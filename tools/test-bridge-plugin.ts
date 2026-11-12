import { randomBytes, randomUUID, timingSafeEqual } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import type { IncomingMessage, ServerResponse } from 'node:http';
import type { Plugin } from 'vite';

const PREFIX = '/__autooffice_test';
const MAX_BODY = 2 * 1024 * 1024;
const TTL = 5 * 60_000;
type Json = Record<string, any>;

/** In-memory, per-window RPC queue. Never retries a document mutation. */
export function createTestBroker(token: string, baseUrl: string, now = Date.now) {
  const sessions = new Map<string, Json>();
  const commands = new Map<string, Json>();
  let lastDenial: Json | null = null;
  function prune() {
    for (const [id, session] of sessions) if (now() - session.lastSeen > TTL) sessions.delete(id);
    for (const [id, command] of commands) if (now() - command.createdAt > TTL) commands.delete(id);
  }
  return async (req: IncomingMessage, res: ServerResponse) => {
    const reply = (status: number, value: unknown) => {
      res.writeHead(status, {
        'Content-Type': 'application/json', 'Cache-Control': 'no-store',
        // Override the permissive CORS header of the normal dev server.
        'Access-Control-Allow-Origin': baseUrl, 'Vary': 'Origin',
      });
      res.end(JSON.stringify(value));
    };
    try {
      const url = new URL(req.url || '/', baseUrl);
      const peer = req.socket.remoteAddress || '';
      const authority = req.headers.host ?? req.headers[':authority'];
      if (!['127.0.0.1', '::1', '::ffff:127.0.0.1'].includes(peer) ||
          authority !== new URL(baseUrl).host ||
          (req.headers.origin && req.headers.origin !== baseUrl)) {
        lastDenial = { reason: 'local-origin', peer, host: authority, origin: req.headers.origin };
        reply(403, { error: 'Local same-origin requests only' }); return;
      }
      if (url.pathname === `${PREFIX}/bootstrap` && req.method === 'GET') {
        if (req.headers['sec-fetch-site'] !== 'same-origin' || req.headers['x-autooffice-client'] !== 'word') {
          lastDenial = { reason: 'bootstrap', origin: req.headers.origin,
            fetchSite: req.headers['sec-fetch-site'], client: req.headers['x-autooffice-client'] };
          reply(403, { error: 'Bootstrap requires the local Word add-in' }); return;
        }
        reply(200, { token }); return;
      }
      const supplied = Buffer.from(String(req.headers.authorization || ''));
      const expected = Buffer.from(`Bearer ${token}`);
      if (supplied.length !== expected.length || !timingSafeEqual(supplied, expected)) {
        reply(401, { error: 'Invalid test bridge token' }); return;
      }
      prune();
      let body: Json = {};
      if (req.method === 'POST') {
        let size = 0;
        const chunks: Buffer[] = [];
        for await (const chunk of req) {
          size += Buffer.byteLength(chunk);
          if (size > MAX_BODY) { reply(413, { error: 'Payload too large' }); return; }
          chunks.push(Buffer.from(chunk));
        }
        try { body = JSON.parse(Buffer.concat(chunks).toString('utf8')); }
        catch { reply(400, { error: 'Invalid JSON' }); return; }
        if (!body || typeof body !== 'object' || Array.isArray(body)) {
          reply(400, { error: 'Expected object' }); return;
        }
      }
      const route = url.pathname.slice(PREFIX.length);
      if (route === '/diagnostics' && req.method === 'GET') {
        reply(200, { lastDenial }); return;
      }
      if (route === '/sessions' && req.method === 'GET') {
        reply(200, { sessions: [...sessions.values()] }); return;
      }
      if (route === '/register' && req.method === 'POST') {
        if (typeof body.sessionId !== 'string' || typeof body.documentLabel !== 'string') {
          reply(400, { error: 'sessionId and documentLabel required' }); return;
        }
        sessions.set(body.sessionId, {
          sessionId: body.sessionId, documentLabel: body.documentLabel,
          lastSeen: now(),
        });
        reply(200, { ok: true }); return;
      }
      if (route === '/commands' && req.method === 'POST') {
        if (!sessions.has(body.sessionId)) { reply(404, { error: 'Session not connected' }); return; }
        if (!['status', 'snapshot', 'fixture', 'select', 'run'].includes(body.action)) {
          reply(400, { error: 'Unsupported action' }); return;
        }
        if ([...commands.values()].some(c => c.sessionId === body.sessionId && ['queued', 'running'].includes(c.state))) {
          reply(409, { error: 'Session already has a pending command' }); return;
        }
        if (commands.size >= 100) { reply(429, { error: 'Command limit reached; wait for expiry' }); return; }
        const command = { id: randomUUID(), sessionId: body.sessionId, action: body.action,
          args: body.args ?? {}, state: 'queued', createdAt: now() };
        commands.set(command.id, command);
        reply(202, command); return;
      }
      if (route === '/next' && req.method === 'GET') {
        const sessionId = url.searchParams.get('sessionId') || '';
        const session = sessions.get(sessionId);
        if (!session) { reply(404, { error: 'Session expired; reconnect' }); return; }
        session.lastSeen = now();
        const command = [...commands.values()].find(c => c.sessionId === sessionId && c.state === 'queued');
        if (command) command.state = 'running';
        reply(200, { command: command ?? null }); return;
      }
      if (route === '/result' && req.method === 'POST') {
        const command = commands.get(body.id);
        // A response may be lost after committing. Acknowledge identical delivery again.
        if (command && command.sessionId === body.sessionId &&
            command.state === (body.ok === true ? 'done' : 'error') &&
            JSON.stringify(command.result) === JSON.stringify(body.result)) {
          reply(200, { ok: true }); return;
        }
        if (!command || command.sessionId !== body.sessionId || command.state !== 'running') {
          reply(409, { error: 'Command not running in this session' }); return;
        }
        command.state = body.ok === true ? 'done' : 'error';
        command.result = body.result;
        command.completedAt = now();
        reply(200, { ok: true }); return;
      }
      if (route.startsWith('/commands/') && req.method === 'GET') {
        const command = commands.get(route.slice('/commands/'.length));
        reply(command ? 200 : 404, command ?? { error: 'Unknown or expired command; execution may have occurred. Do not replay mutations blindly.' });
        return;
      }
      reply(404, { error: 'Unknown bridge route' });
    } catch (error) {
      if (!res.headersSent) reply(500, { error: error instanceof Error ? error.message : 'Bridge failure' });
    }
  };
}

export function localTestBridgePlugin(): Plugin {
  const enabled = process.env.AUTOOFFICE_TEST_BRIDGE === '1';
  return {
    name: 'autooffice-local-test-bridge',
    config(config, { command }) {
      // Never serve tokens or document reports through Vite's static / @fs routes.
      // Vite computes its deny matcher BEFORE configResolved, so configure it here.
      const defaultDeny = ['.env', '.env.*', '*.{crt,pem}', '**/.git/**'];
      return {
        define: { __AUTOOFFICE_TEST_BRIDGE__: JSON.stringify(enabled && command === 'serve') },
        server: { fs: { deny: [...(config.server?.fs?.deny ?? defaultDeny), '**/.autooffice-test/**'] } },
      };
    },
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        let pathname: string;
        try { pathname = decodeURIComponent((req.url || '').split('?')[0]); }
        catch { res.writeHead(400); res.end(); return; }
        if (/(^|[\\/])\.autooffice-test([\\/]|$)/i.test(pathname)) {
          res.writeHead(403, { 'Cache-Control': 'no-store' }); res.end('Private local test artifacts'); return;
        }
        next();
      });
      if (!enabled) return;
      const token = randomBytes(32).toString('hex');
      // Resolve the actual listening port (Vite can choose another if it is occupied).
      let handler: ReturnType<typeof createTestBroker> | undefined;
      const connectionFile = path.join(server.config.root, '.autooffice-test', 'connection.json');
      server.httpServer?.once('listening', () => {
        const address = server.httpServer!.address();
        if (!address || typeof address === 'string') return;
        const baseUrl = `${server.config.server.https ? 'https' : 'http'}://localhost:${address.port}`;
        handler = createTestBroker(token, baseUrl);
        fs.mkdirSync(path.dirname(connectionFile), { recursive: true });
        fs.writeFileSync(connectionFile, JSON.stringify({ baseUrl, token }, null, 2), { mode: 0o600 });
        server.config.logger.info('AutoOffice local test bridge enabled (test documents only).');
      });
      server.middlewares.use((req, res, next) => {
        const pathname = (req.url || '').split('?')[0];
        if (pathname !== PREFIX && !pathname.startsWith(`${PREFIX}/`)) return next();
        if (!handler) { res.writeHead(503); res.end(); return; }
        void handler(req, res);
      });
      server.httpServer?.once('close', () => {
        try {
          if (JSON.parse(fs.readFileSync(connectionFile, 'utf8')).token === token) fs.unlinkSync(connectionFile);
        } catch { /* Another server owns the connection file, or it was already removed. */ }
      });
    },
  };
}
