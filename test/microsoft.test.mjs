import test from 'node:test';
import assert from 'node:assert/strict';
import { microsoftMessage, microsoftSync, graphPath } from '../src/microsoft.mjs';
const account = {
  id: 'fixture',
  generation: 'one',
  historyDays: 7
};
const message = (id, date = new Date().toISOString()) => ({
  id,
  conversationId: 'thread',
  receivedDateTime: date,
  body: {
    content: '<p>Hello</p>',
    contentType: 'html'
  },
  from: {
    emailAddress: {
      address: 'sender@example.com'
    }
  }
});
const cursor = () => ({
  index: 0,
  links: {
    inbox: '/v1.0/me/mailFolders/inbox/messages/delta?$skiptoken=page'
  },
  complete: {},
  folderIds: {
    inbox: 'in',
    sentitems: 'sent',
    archive: 'arc'
  }
});
test('delta links stay in the reviewed Graph me namespace', () => {
  assert.throws(() => graphPath('https://evil.example/v1.0/me/messages'));
  assert.throws(() => graphPath('https://graph.microsoft.com/v1.0/users/other/messages'));
  assert.equal(graphPath('https://graph.microsoft.com/v1.0/me/messages?$skiptoken=x'), '/v1.0/me/messages?$skiptoken=x');
});
test('initial history bound applies to later pages', async () => {
  let saved;
  await microsoftSync(account, cursor(), {
    read: async (_a, path) => path.includes('/delta') ? {
      value: [message('old', '2000-01-01T00:00:00Z'), message('new')],
      '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=done'
    } : message('new'),
    commit: async (_a, c, m) => {
      saved = {
        c,
        m
      };
    }
  });
  assert.deepEqual(saved.m.map(m => m.externalId), ['new']);
  assert.equal(saved.c.complete.inbox, true);
});
test('removed entries distinguish deletion, archive and trash moves', async () => {
  let saved;
  const c = cursor();
  c.complete.inbox = true;
  await microsoftSync(account, c, {
    read: async (_a, path) => {
      if (path.includes('/delta')) return {
        value: ['deleted', 'moved', 'trash'].map(id => ({
          id,
          '@removed': {
            reason: 'deleted'
          }
        })),
        '@odata.deltaLink': '/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=done'
      };
      if (path.includes('/deleted?')) throw Object.assign(new Error('gone'), {
        providerStatus: 404
      });
      return {
        ...message(path.includes('/moved?') ? 'moved' : 'trash'),
        parentFolderId: path.includes('/moved?') ? 'arc' : 'trash-folder'
      };
    },
    commit: async (_a, _c, m) => {
      saved = m;
    }
  });
  assert.equal(saved.find(m => m.externalId === 'deleted').deleted, true);
  assert.equal(saved.find(m => m.externalId === 'moved').archived, true);
  assert.equal(saved.find(m => m.externalId === 'trash').deleted, true);
});
test('expired Graph cursor explicitly resets membership', async () => {
  let saved;
  await microsoftSync(account, cursor(), {
    read: async () => {
      throw Object.assign(new Error('expired'), {
        providerStatus: 410
      });
    },
    commit: async (_a, c, m, reset, phase) => {
      saved = {
        c,
        m,
        reset,
        phase
      };
    }
  });
  assert.deepEqual(saved, {
    c: {},
    m: [],
    reset: true,
    phase: 'waiting'
  });
});
test('plain comparison text, HTML and UTF-8 bounds preserve meaning', () => {
  assert.equal(microsoftMessage({
    ...message('plain'),
    body: {
      contentType: 'text',
      content: '2 < 3 and 5 > 4'
    }
  }, 'inbox').body, '2 < 3 and 5 > 4');
  assert.equal(microsoftMessage(message('html'), 'inbox').body, 'Hello');
  assert.equal(microsoftMessage(message('sent'), 'sentitems').sent, true);
  assert.equal(microsoftMessage({
    ...message('long'),
    subject: 'x'.repeat(2100)
  }, 'inbox').truncated, true);
});
