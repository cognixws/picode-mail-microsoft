import { createHost } from './host.mjs';
import { microsoftSync } from './microsoft.mjs';
import { run } from './runtime.mjs';
const controller = new AbortController();
process.once('SIGTERM', () => controller.abort());
process.once('SIGINT', () => controller.abort());
await run(createHost(), microsoftSync, {
  signal: controller.signal
});
