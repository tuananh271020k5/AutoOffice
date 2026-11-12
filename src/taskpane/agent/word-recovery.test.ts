import { afterEach, describe, expect, it, vi } from 'vitest';
import { captureWordRecovery, MAX_RECOVERY_BYTES, openWordRecovery } from './word-recovery';

afterEach(() => { vi.unstubAllGlobals(); vi.useRealTimers(); });
function mockFile(chunks = [[80, 75], [3, 4]]) {
  const file = {
    size: chunks.flat().length, sliceCount: chunks.length,
    getSliceAsync: vi.fn((index, callback) => callback({ status: 'succeeded', value: { data: chunks[index] } })),
    closeAsync: vi.fn(),
  };
  const getFileAsync = vi.fn((_type, _options, callback) => callback({ status: 'succeeded', value: file }));
  vi.stubGlobal('Office', { FileType: { Compressed: 'compressed' }, AsyncResultStatus: { Succeeded: 'succeeded' },
    context: { document: { getFileAsync } } });
  return { file, getFileAsync };
}

describe('full Word recovery copies', () => {
  it('joins compressed slices in order and closes the file handle', async () => {
    const { file, getFileAsync } = mockFile();
    const copy = await captureWordRecovery('Test.docx');
    expect(getFileAsync).toHaveBeenCalledWith('compressed', { sliceSize: 65536 }, expect.any(Function));
    expect(copy.base64).toBe('UEsDBA==');
    expect(copy.bytes).toBe(4);
    expect(copy.documentLabel).toBe('Test.docx');
    expect(file.closeAsync).toHaveBeenCalledOnce();
  });
  it('closes the file and refuses oversized or incomplete backups', async () => {
    const { file } = mockFile();
    file.size = MAX_RECOVERY_BYTES + 1;
    await expect(captureWordRecovery('Test')).rejects.toThrow('20 MB');
    expect(file.closeAsync).toHaveBeenCalledOnce();
    file.size = 9;
    await expect(captureWordRecovery('Test')).rejects.toThrow('Incomplete');
    expect(file.closeAsync).toHaveBeenCalledTimes(2);
  });
  it('releases the file when a slice fails', async () => {
    const { file } = mockFile();
    file.getSliceAsync.mockImplementation((_i, callback) => callback({ status: 'failed' } as any));
    await expect(captureWordRecovery('Test')).rejects.toThrow('Could not read');
    expect(file.closeAsync).toHaveBeenCalledOnce();
  });
  it('closes a file that arrives after the initial timeout', async () => {
    vi.useFakeTimers();
    const { getFileAsync, file } = mockFile();
    let callback!: Function;
    getFileAsync.mockImplementation((_type, _options, cb) => { callback = cb; });
    const assertion = expect(captureWordRecovery('Test')).rejects.toThrow('timed out');
    await vi.advanceTimersByTimeAsync(30_001);
    await assertion;
    callback({ status: 'succeeded', value: file });
    expect(file.closeAsync).toHaveBeenCalledOnce();
  });
  it('opens a separate document without modifying the current document', async () => {
    const open = vi.fn();
    const createDocument = vi.fn(() => ({ open }));
    const sync = vi.fn(async () => {});
    vi.stubGlobal('Word', { run: async (fn: Function) => fn({ application: { createDocument }, sync }) });
    await openWordRecovery({ base64: 'UEsDBA==', bytes: 4, createdAt: 1, documentLabel: 'Test' });
    expect(createDocument).toHaveBeenCalledExactlyOnceWith('UEsDBA==');
    expect(open).toHaveBeenCalledOnce();
    expect(sync).toHaveBeenCalledOnce();
  });
});
