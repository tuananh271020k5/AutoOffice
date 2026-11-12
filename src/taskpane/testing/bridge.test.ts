import { afterEach, describe, expect, it, vi } from 'vitest';
import { executeTestCommand, isTestDocument, type TestBridgeAccess } from './bridge';

afterEach(() => vi.unstubAllGlobals());
describe('Word test bridge document boundary', () => {
  it('only allows explicit test filenames and real sideload sample names', () => {
    expect(isTestDocument('C:\\Temp\\Word add-in ad7b1a94-0e6d-4263-acf1-8272d208af31.2.docx')).toBe(true);
    expect(isTestDocument('file:///D:/AutoOffice-Test-demo.docx')).toBe(true);
    for (const url of ['', 'C:/Contracts.docx', 'C:/AutoOffice-Test/Contract.docx',
      'C:/Word add-in fake.docx', 'C:/AutoOffice-Test.docx.backup']) expect(isTestDocument(url)).toBe(false);
  });
  it('rejects even read commands on personal documents', async () => {
    vi.stubGlobal('Word', {});
    vi.stubGlobal('Office', { context: { document: { url: 'C:/Contract.docx' } } });
    const access = { busy: () => false, status: vi.fn(), send: vi.fn() };
    await expect(executeTestCommand({ id: '1', action: 'snapshot', args: {} }, access)).rejects.toThrow('Only the sideload');
    expect(access.status).not.toHaveBeenCalled();
  });
  it('checks busy state and provider readiness before sending a real turn', async () => {
    vi.stubGlobal('Word', {});
    vi.stubGlobal('Office', { context: { document: { url: 'C:/AutoOffice-Test.docx' } } });
    const access: TestBridgeAccess = { busy: () => true, status: () => ({ hasApiKey: false }), send: vi.fn() };
    const command = { id: '1', action: 'run' as const, args: { prompt: 'test' } };
    await expect(executeTestCommand(command, access)).rejects.toThrow('busy');
    access.busy = () => false;
    await expect(executeTestCommand(command, access)).rejects.toThrow('API key');
    expect(access.send).not.toHaveBeenCalled();
  });
  it('uses the actual send callback between independent before/after document reads', async () => {
    let text = 'Before';
    const load = () => {};
    vi.stubGlobal('Word', { run: async (fn: Function) => fn({
      document: {
        body: { get text() { return text; }, load, paragraphs: { load, items: [] }, tables: { load, items: [] } },
        getSelection: () => ({ text: '', load }),
      }, sync: async () => {},
    }) });
    vi.stubGlobal('Office', { context: { document: { url: 'C:/AutoOffice-Test.docx' } } });
    const send = vi.fn(async () => { text = 'After'; return { errors: [], toolResults: [] }; });
    const result = await executeTestCommand({ id: '1', action: 'run', args: { prompt: 'Change text' } }, {
      busy: () => false, status: () => ({ hasApiKey: true }), send,
    }) as { before: { text: string }; after: { text: string } };
    expect(send).toHaveBeenCalledExactlyOnceWith('Change text');
    expect(result.before.text).toBe('Before');
    expect(result.after.text).toBe('After');
  });
});
