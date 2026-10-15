import { messageBatches } from './batches.mjs';
import { clipped, text } from './text.mjs';
const graphPath = link => {
  const u = new URL(link, 'https://graph.microsoft.com');
  if (u.origin !== 'https://graph.microsoft.com' || !u.pathname.startsWith('/v1.0/me/')) throw new Error('Foreign delta URL refused');
  return u.pathname + u.search;
};
const graphBody = m => m.body ? String(m.body.contentType || '').toLowerCase() === 'html' ? text(m.body.content) : String(m.body.content || '') : String(m.bodyPreview || '');
function microsoftMessage(m, folder) {
  return {
    sent: folder === 'sentitems',
    externalId: m.id,
    threadId: m.conversationId || m.id,
    from: m.from?.emailAddress?.address || '',
    fromName: m.from?.emailAddress?.name || '',
    to: (m.toRecipients || []).map(r => r.emailAddress.address),
    cc: (m.ccRecipients || []).map(r => r.emailAddress.address),
    subject: clipped(m.subject || '(No subject)', 2000),
    body: clipped(graphBody(m)),
    truncated: !!m.truncated || Buffer.byteLength(graphBody(m)) > 1_000_000 || Buffer.byteLength(m.subject || '') > 2000,
    receivedAt: m.receivedDateTime || new Date().toISOString(),
    messageId: m.internetMessageId || '',
    unread: !m.isRead,
    archived: folder !== 'inbox',
    attachments: []
  };
}
const graphSelect = 'id,conversationId,from,toRecipients,ccRecipients,subject,bodyPreview,receivedDateTime,isRead,internetMessageId,hasAttachments,parentFolderId';
async function microsoftSync(a, c, io) {
  const {
    read,
    commit
  } = io;
  const folders = ['inbox', 'sentitems', 'archive'];
  c = {
    index: 0,
    links: {},
    complete: {},
    folderIds: {},
    ...c
  };
  // Immutable provider folder IDs distinguish moves from permanent deletion.
  for (const folder of folders) if (!c.folderIds[folder]) c.folderIds[folder] = (await read(a, '/v1.0/me/mailFolders/' + folder + '?$select=id')).id;
  const folder = folders[c.index % folders.length];
  const base = '/v1.0/me/mailFolders/' + folder + '/messages/delta';
  const path = c.links[folder] || base + '?$top=10&$select=' + graphSelect;
  let result;
  try {
    result = await read(a, path);
  } catch (e) {
    if ([404, 410].includes(e.providerStatus)) {
      await commit(a, {}, [], true, "waiting");
      return true;
    }
    throw e;
  }
  const messages = [];
  for (let m of result.value || []) {
    let location = folder;
    if (m['@removed']) {
      try {
        m = await read(a, '/v1.0/me/messages/' + encodeURIComponent(m.id) + '?$select=' + graphSelect);
      } catch (e) {
        if (e.providerStatus === 404) {
          messages.push({
            externalId: m.id,
            deleted: true
          });
          continue;
        }
        throw e;
      }
      location = folders.find(f => c.folderIds[f] === m.parentFolderId);
      if (!location) {
        messages.push({
          externalId: m.id,
          deleted: true
        });
        continue;
      }
    }
    // The initial-history bound applies to every initial page, not only page one.
    if (!c.complete[folder] && Date.parse(m.receivedDateTime) < Date.now() - a.historyDays * 86400000) continue;
    const summary = m;
    try {
      m = await read(a, '/v1.0/me/messages/' + encodeURIComponent(m.id) + '?$select=' + graphSelect + ',body');
    } catch (e) {
      if (e.providerStatus === 404) {
        messages.push({
          externalId: m.id,
          deleted: true
        });
        continue;
      }
      if (e.providerStatus !== 200) throw e;
      m = {
        ...summary,
        truncated: true
      };
    }
    const normalized = microsoftMessage(m, location);
    if (m.hasAttachments) {
      const files = await read(a, '/v1.0/me/messages/' + encodeURIComponent(m.id) + '/attachments?$select=id,name,contentType,size');
      normalized.truncated ||= (files.value || []).length > 50 || !!files['@odata.nextLink'];
      normalized.attachments = (files.value || []).slice(0, 50).map(f => ({
        id: f.id,
        name: f.name,
        mime: f.contentType,
        size: f.size
      }));
    }
    messages.push(normalized);
  }
  const previous = structuredClone(c);
  const next = result['@odata.nextLink'],
    delta = result['@odata.deltaLink'];
  if (next) c.links[folder] = graphPath(next);else {
    if (!delta) throw new Error('Missing provider delta cursor');
    c.links[folder] = graphPath(delta);
    c.complete[folder] = true;
    c.index = (c.index + 1) % folders.length;
  }
  const pending = !!next || !folders.every(f => c.complete[f]);
  const batches = messageBatches(messages);
  if (!batches.length) await commit(a, c, [], false, pending ? 'syncing' : 'ready');
  for (let i = 0; i < batches.length; i++) {
    const last = i === batches.length - 1;
    await commit(a, last ? c : previous, batches[i], false, last && !pending ? 'ready' : 'syncing');
  }
  return pending;
}
export { microsoftMessage, microsoftSync, graphPath };
