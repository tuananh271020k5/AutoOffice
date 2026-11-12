import { webcrypto } from 'node:crypto';
import { TextEncoder } from 'node:util';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { beginWordDraft, canonicalBodyXml } from './word-draft';

const xml = (text: string) => `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`;
afterEach(() => vi.unstubAllGlobals());
function environment() {
  let value = xml('a');
  const insertOoxml = vi.fn((content: string) => { value = content; });
  const context = { document: { load: vi.fn(), changeTrackingMode: 'Off', getSelection: () => ({ parentBody: { type: 'MainDoc', load() {} } }),
    body: { getOoxml: () => ({ value }), insertOoxml } }, sync: async () => {} };
  vi.stubGlobal('Word', { run: async (fn: Function) => fn(context), ChangeTrackingMode: { off: 'Off' }, InsertLocation: { replace: 'Replace' } });
  vi.stubGlobal('crypto', webcrypto);
  vi.stubGlobal('TextEncoder', TextEncoder);
  return { read: () => value, write: (content: string) => { value = content; }, insertOoxml };
}
describe('inline Word Accept / Reject', () => {
  it('Reject restores the original body in the same document', async () => {
    const word = environment();
    const draft = await beginWordDraft();
    word.write(xml('A'));
    await draft.seal();
    await draft.reject();
    expect(word.read()).toBe(xml('a'));
    expect(word.insertOoxml).toHaveBeenCalledExactlyOnceWith(xml('a'), 'Replace');
  });
  it('Accept keeps the visible change without opening another file', async () => {
    const word = environment();
    const draft = await beginWordDraft();
    word.write(xml('A'));
    await draft.seal();
    draft.accept();
    expect(word.read()).toBe(xml('A'));
    expect(word.insertOoxml).not.toHaveBeenCalled();
  });
  it('does not overwrite user edits made after the draft appeared', async () => {
    const word = environment();
    const draft = await beginWordDraft();
    word.write(xml('A'));
    await draft.seal();
    word.write(xml('New manual edit'));
    await expect(draft.reject()).rejects.toThrow('preserve your newer edits');
    expect(word.read()).toBe(xml('New manual edit'));
    expect(word.insertOoxml).not.toHaveBeenCalled();
  });
  it('ignores generated edit IDs, but never ignores actual text or formatting', () => {
    const first = xml('a').replace('<w:p>', '<w:p w:rsidR="111">');
    const second = xml('a').replace('<w:p>', '<w:p w:rsidR="222">');
    expect(canonicalBodyXml(first)).toBe(canonicalBodyXml(second));
    expect(canonicalBodyXml(xml('a'))).not.toBe(canonicalBodyXml(xml('A')));
    expect(canonicalBodyXml(xml('a').replace('<w:r>', '<w:r><w:rPr><w:b/></w:rPr>'))).not.toBe(canonicalBodyXml(xml('a')));
  });
  it('ignores Word automatic counters but refuses to modify existing revisions', async () => {
    const pkg = (words: number) => `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/docProps/app.xml"><pkg:xmlData><Words>${words}</Words></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>${xml('a')}</pkg:xmlData></pkg:part></pkg:package>`;
    expect(canonicalBodyXml(pkg(1))).toBe(canonicalBodyXml(pkg(2)));
    const word = environment();
    word.write(xml('a').replace('<w:r>', '<w:ins><w:r>').replace('</w:r>', '</w:r></w:ins>'));
    await expect(beginWordDraft()).rejects.toThrow('existing Word revisions');
    expect(word.insertOoxml).not.toHaveBeenCalled();
  });
});
