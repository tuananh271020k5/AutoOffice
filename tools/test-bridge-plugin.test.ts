// @vitest-environment node
import { describe, expect, it } from 'vitest';
import { Readable } from 'node:stream';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { createTestBroker, localTestBridgePlugin } from './test-bridge-plugin';

const baseUrl = 'https://localhost:3721';
const token = 'a'.repeat(64);
function harness() {
  let now = 1000;
  const handler = createTestBroker(token, baseUrl, () => now);
  const call = async (route: string, body?: unknown, headers: Record<string, string> = {}, peer = '::1') => {
    const req = Readable.from(body === undefined ? [] : [JSON.stringify(body)]) as IncomingMessage;
    Object.assign(req, { url: `/__autooffice_test${route}`, method: body === undefined ? 'GET' : 'POST',
      headers: { host: 'localhost:3721', authorization: `Bearer ${token}`, ...headers },
      socket: { remoteAddress: peer } });
    let status = 0;
    let data: any;
    const res = {
      headersSent: false,
      writeHead(code: number) { status = code; },
      end(text: string) { data = JSON.parse(text); },
    } as unknown as ServerResponse;
    await handler(req, res);
    return { status, data };
  };
  return { call, advance: () => { now += 301_000; } };
}

describe('local Word test broker', () => {
  it('requires token, local peer, exact host and same origin', async () => {
    const { call } = harness();
    expect((await call('/sessions', undefined, { authorization: '' })).status).toBe(401);
    expect((await call('/sessions', undefined, {}, '192.168.1.2')).status).toBe(403);
    expect((await call('/sessions', undefined, { origin: 'https://evil.example' })).status).toBe(403);
    expect((await call('/sessions', undefined, { host: 'evil.example:3721' })).status).toBe(403);
    expect((await call('/sessions')).status).toBe(200);
  });
  it('only bootstraps same-origin browser fetches with the custom header', async () => {
    const { call } = harness();
    expect((await call('/bootstrap')).status).toBe(403);
    expect((await call('/bootstrap', undefined, { 'sec-fetch-site': 'cross-site', 'x-autooffice-client': 'word' })).status).toBe(403);
    expect((await call('/bootstrap', undefined, { 'sec-fetch-site': 'same-origin', 'x-autooffice-client': 'word' })).data.token).toBe(token);
  });
  it('accepts HTTP/2 authority from the real Word WebView', async () => {
    const { call } = harness();
    expect((await call('/sessions', undefined, { host: undefined as unknown as string,
      ':authority': 'localhost:3721' })).status).toBe(200);
    expect((await call('/sessions', undefined, { host: undefined as unknown as string,
      ':authority': 'evil.example' })).status).toBe(403);
  });
  it('isolates windows, dispatches once and accepts an identical result retry', async () => {
    const { call } = harness();
    for (const id of ['one', 'two']) await call('/register', { sessionId: id, documentLabel: `${id}.docx` });
    const { data: command } = await call('/commands', { sessionId: 'one', action: 'run', args: { prompt: 'test' } });
    expect((await call('/commands', { sessionId: 'one', action: 'snapshot' })).status).toBe(409);
    expect((await call('/next?sessionId=two')).data.command).toBeNull();
    expect((await call('/next?sessionId=one')).data.command.id).toBe(command.id);
    expect((await call('/next?sessionId=one')).data.command).toBeNull();
    const result = { id: command.id, sessionId: 'one', ok: true, result: { text: 'done' } };
    expect((await call('/result', { ...result, sessionId: 'two' })).status).toBe(409);
    expect((await call('/result', result)).status).toBe(200);
    expect((await call('/result', result)).status).toBe(200);
    expect((await call('/result', { ...result, result: 'different' })).status).toBe(409);
    expect((await call(`/commands/${command.id}`)).data.state).toBe('done');
  });
  it('rejects arbitrary execution and expires sessions and command records', async () => {
    const { call, advance } = harness();
    await call('/register', { sessionId: 'one', documentLabel: 'test' });
    expect((await call('/commands', { sessionId: 'one', action: 'execute_code' })).status).toBe(400);
    const { data: command } = await call('/commands', { sessionId: 'one', action: 'snapshot' });
    advance();
    expect((await call('/sessions')).data.sessions).toEqual([]);
    expect((await call(`/commands/${command.id}`)).status).toBe(404);
  });
  it('defines the client flag as false in production even when opted in', () => {
    const original = process.env.AUTOOFFICE_TEST_BRIDGE;
    try {
      process.env.AUTOOFFICE_TEST_BRIDGE = '1';
      const config = localTestBridgePlugin().config as Function;
      expect(config({}, { command: 'build' }).define.__AUTOOFFICE_TEST_BRIDGE__).toBe('false');
      expect(config({}, { command: 'serve' }).define.__AUTOOFFICE_TEST_BRIDGE__).toBe('true');
    } finally {
      if (original === undefined) delete process.env.AUTOOFFICE_TEST_BRIDGE;
      else process.env.AUTOOFFICE_TEST_BRIDGE = original;
    }
  });
  it('prevents static access to tokens and reports without replacing Vite defaults', () => {
    const configure = localTestBridgePlugin().config as Function;
    expect(configure({}, { command: 'serve' }).server.fs.deny).toEqual([
      '.env', '.env.*', '*.{crt,pem}', '**/.git/**', '**/.autooffice-test/**',
    ]);
    expect(configure({ server: { fs: { deny: ['custom-secret'] } } }, { command: 'serve' }).server.fs.deny)
      .toContain('custom-secret');
  });
});
