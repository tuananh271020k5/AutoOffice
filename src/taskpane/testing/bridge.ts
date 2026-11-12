import { getDocumentIdentity } from '../host/context.ts';
import { setCachedWordSelection } from '../agent/kitchensink-word-tools.ts';

export interface TestCommand {
  id: string;
  action: 'status' | 'snapshot' | 'fixture' | 'select' | 'run';
  args: Record<string, unknown>;
}

export interface TestBridgeAccess {
  busy: () => boolean;
  status: () => Record<string, unknown>;
  send: (prompt: string) => Promise<unknown>;
}

/** Check the actual Office URL, never the name supplied by an HTTP caller. */
export function isTestDocument(url: string): boolean {
  try {
    const name = decodeURIComponent(url).split(/[\\/]/).pop() || '';
    return /^(?:Word add-in [0-9a-f-]{36}(?:\.\d+)?|AutoOffice-Test[\w .-]*)\.docx$/i.test(name);
  } catch { return false; }
}

function assertTestDocument() {
  if (typeof Word === 'undefined' || !isTestDocument(Office.context.document.url || '')) {
    throw new Error('Only the sideload sample or an AutoOffice-Test*.docx document is allowed.');
  }
}

/** Independent read-back via Word APIs; does not trust the agent's success message. */
export async function readWordSnapshot() {
  return Word.run(async context => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    const tables = body.tables;
    const selection = context.document.getSelection();
    body.load('text');
    selection.load('text');
    paragraphs.load('text,style,alignment,font/name,font/size,font/bold,font/italic,font/underline');
    tables.load('values,rowCount');
    await context.sync();
    return {
      text: body.text,
      selectedText: selection.text,
      paragraphs: paragraphs.items.map(p => ({
        text: p.text, style: p.style, alignment: p.alignment,
        font: { name: p.font.name, size: p.font.size, bold: p.font.bold,
          italic: p.font.italic, underline: p.font.underline },
      })),
      tables: tables.items.map(t => ({ values: t.values, rowCount: t.rowCount })),
    };
  });
}

export async function executeTestCommand(command: TestCommand, access: TestBridgeAccess) {
  assertTestDocument();
  if (command.action === 'status') return access.status();
  if (access.busy()) throw new Error('AutoOffice is busy; wait for the current chat turn.');
  if (command.action === 'snapshot') return readWordSnapshot();
  if (command.action === 'fixture') {
    const paragraphs = command.args.paragraphs;
    if (!Array.isArray(paragraphs) || paragraphs.length > 100 ||
        !paragraphs.every(p => typeof p === 'string' && p.length <= 10_000)) {
      throw new Error('fixture requires up to 100 string paragraphs, each up to 10000 characters.');
    }
    await Word.run(async context => {
      const body = context.document.body;
      body.clear();
      for (const [index, text] of paragraphs.entries()) {
        const p = body.insertParagraph(text, Word.InsertLocation.end);
        p.font.name = 'Calibri';
        p.font.size = 11;
        p.font.bold = index === 0 && command.args.boldFirst === true;
        p.font.italic = false;
        p.font.underline = Word.UnderlineType.none;
        p.alignment = Word.Alignment.left;
      }
      body.getRange(Word.RangeLocation.end).select();
      await context.sync();
    });
    setCachedWordSelection('');
    return readWordSnapshot();
  }
  if (command.action === 'select') {
    const text = command.args.text;
    if (typeof text !== 'string' || !text.trim() || text.length > 10_000) {
      throw new Error('select requires non-empty text, up to 10000 characters.');
    }
    await Word.run(async context => {
      const hits = context.document.body.search(text, { matchCase: true, matchWildcards: false });
      hits.load('items');
      await context.sync();
      if (hits.items.length !== 1) throw new Error(`Expected one exact match, found ${hits.items.length}.`);
      hits.items[0].select();
      await context.sync();
    });
    setCachedWordSelection(text);
    return readWordSnapshot();
  }
  if (command.action === 'run') {
    const prompt = command.args.prompt;
    if (typeof prompt !== 'string' || !prompt.trim() || prompt.length > 20_000) {
      throw new Error('run requires a non-empty prompt, up to 20000 characters.');
    }
    const status = access.status();
    if (!status.hasApiKey) throw new Error('Configure an AI provider/model and API key in AutoOffice settings first.');
    // Use the exact same UI send handler. Never override approval or provider settings.
    const before = await readWordSnapshot();
    const turn = await access.send(prompt);
    const after = await readWordSnapshot();
    return { before, turn, after };
  }
  throw new Error('Unsupported test command');
}

export function startTestBridge(access: TestBridgeAccess, onState: (state: string) => void = () => {}): () => void {
  if (!import.meta.env.DEV || typeof Word === 'undefined') {
    onState('Kiểm thử: cần mở trong Word.');
    return () => {};
  }
  if (!isTestDocument(Office.context.document.url || '')) {
    onState('Kiểm thử: hãy lưu tài liệu thành AutoOffice-Test.docx rồi mở lại khung AutoOffice.');
    return () => {};
  }
  const sessionId = getDocumentIdentity().windowSessionId;
  const abort = new AbortController();
  let timer: ReturnType<typeof setTimeout>;
  let token = '';
  let registered = false;
  // Keep a completed result until delivered. Never rerun a command on a network failure.
  let pendingResult: Record<string, unknown> | null = null;
  const request = async (route: string, body?: unknown) => {
    const response = await fetch(`/__autooffice_test${route}`, {
      method: body === undefined ? 'GET' : 'POST',
      headers: {
        'X-AutoOffice-Client': 'word',
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
        ...(body === undefined ? {} : { 'Content-Type': 'application/json' }),
      },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      signal: abort.signal,
    });
    if (!response.ok) {
      if (response.status === 401) { token = ''; registered = false; }
      if (response.status === 404) registered = false;
      throw new Error(`Test bridge HTTP ${response.status}`);
    }
    return response.json();
  };
  const poll = async () => {
    try {
      if (!token) token = (await request('/bootstrap')).token;
      if (!registered) {
        const file = decodeURIComponent(Office.context.document.url).split(/[\\/]/).pop()!;
        await request('/register', { sessionId, documentLabel: file });
        registered = true;
        onState('Kiểm thử local: đã kết nối. Chỉ dùng tài liệu thử.');
      }
      if (pendingResult) {
        await request('/result', pendingResult);
        pendingResult = null;
      }
      const { command } = await request(`/next?sessionId=${encodeURIComponent(sessionId)}`);
      if (command && !abort.signal.aborted) {
        try {
          const result = await executeTestCommand(command, access);
          pendingResult = { id: command.id, sessionId, ok: true, result };
        } catch (error) {
          pendingResult = { id: command.id, sessionId, ok: false,
            result: { error: error instanceof Error ? error.message : String(error) } };
        }
        if (!abort.signal.aborted) {
          await request('/result', pendingResult);
          pendingResult = null;
        }
      }
    } catch (error) {
      if (!abort.signal.aborted) onState(`Kiểm thử: ${error instanceof Error ? error.message : 'mất kết nối'}`);
    }
    finally { if (!abort.signal.aborted) timer = setTimeout(() => void poll(), 700); }
  };
  void poll();
  return () => { clearTimeout(timer); abort.abort(); };
}
