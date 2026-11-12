import { afterEach, describe, expect, it, vi } from 'vitest';
import { normalizeWordArgs } from './word-formatting';
import { buildKs4wScript, consumeWordSelection, getCachedWordSelection, initWordSelectionTracker, setCachedWordSelection } from './kitchensink-word-tools';
import { Sandbox } from '../executor/sandbox';

afterEach(() => { vi.unstubAllGlobals(); setCachedWordSelection(''); });
function wordEnvironment() {
  const font = { name: 'Calibri', size: 11, bold: true, italic: false, underline: 'None', color: '#000000', highlightColor: null };
  let selected = 'a';
  const paragraph: any = { text: 'a', font, load() {}, insertText(text: string) { this.text = text; selected = ''; return { font }; } };
  const paragraphs = { items: [paragraph], load() {} };
  const range: any = {
    get text() { return selected; }, font, paragraphs, load() {},
    insertText(text: string) { paragraph.text = text; selected = ''; return { font }; },
    getRange: () => ({ select: () => { selected = ''; } }),
  };
  const context = { document: { getSelection: () => range, body: { paragraphs } }, sync: async () => {} };
  vi.stubGlobal('Word', { run: async (fn: Function) => fn(context), InsertLocation: { replace: 'Replace' },
    RangeLocation: { end: 'End' }, Alignment: { left: 'Left' }, UnderlineType: { none: 'None', single: 'Single' } });
  return { font, paragraph };
}
describe('Word font and one-use targets', () => {
  it('normalizes 20pt arguments consistently and rejects invalid sizes', () => {
    expect(normalizeWordArgs({ fontSize: '20', formatting: { bold: true } })).toMatchObject({ font_size: 20, bold: true });
    expect(normalizeWordArgs({ formatting: { size: 20 } })).toMatchObject({ font_size: 20, formatting: { font_size: 20 } });
    expect(() => normalizeWordArgs({ font_size: 'huge' })).toThrow('Font size');
  });
  it('actually applies font_size 20 and bold during text replacement', async () => {
    const { font, paragraph } = wordEnvironment();
    const result = await new Sandbox('word').execute(buildKs4wScript('word_edit', 'edit_selection', { new_text: 'A', font_size: 20, bold: true }));
    expect(result.success).toBe(true);
    expect(paragraph.text).toBe('A');
    expect(font.size).toBe(20);
    expect(font.bold).toBe(true);
    expect(font.name).toBe('Calibri');
  });
  it('consumes bold as a replacement marker unless bold was explicitly requested', async () => {
    const { font, paragraph } = wordEnvironment();
    const result = await new Sandbox('word').execute(buildKs4wScript('word_edit', 'edit_selection', { new_text: 'A' }));
    expect(result.success).toBe(true);
    expect(paragraph.text).toBe('A');
    expect(font.bold).toBe(false);
  });
  it('clears the live and cached selection and updates the selection indicator', async () => {
    wordEnvironment();
    const listener = vi.fn();
    const unsubscribe = initWordSelectionTracker(listener);
    setCachedWordSelection('a');
    await consumeWordSelection();
    expect(getCachedWordSelection()).toBe('');
    expect(listener).toHaveBeenLastCalledWith('');
    unsubscribe();
  });
  it('changes only requested font properties when formatting paragraphs without a live selection', async () => {
    const { font } = wordEnvironment();
    await consumeWordSelection();
    const result = await new Sandbox('word').execute(buildKs4wScript('word_format', 'format_selection', { font_size: 20 }));
    expect(result.success).toBe(true);
    expect(font.size).toBe(20);
    expect(font.name).toBe('Calibri');
    expect(font.bold).toBe(true);
  });
});
