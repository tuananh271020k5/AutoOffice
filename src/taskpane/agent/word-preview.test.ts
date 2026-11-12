import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildWordPreview, sameApprovalState, pinWordSelection, selectPinnedRange, releasePinnedRange } from './word-preview';
import type { WordObservation } from './word-verification';

function state(): WordObservation {
  const font = { name: 'Calibri', size: 11, bold: false, italic: false, underline: 'None', color: '#000000', highlightColor: null };
  return { text: 'Target\rKeep\r', paragraphs: ['Target', 'Keep'].map(text => ({ text, alignment: 'Left', font })),
    selection: { text: 'Target', font }, tables: [] };
}
afterEach(() => vi.unstubAllGlobals());
describe('human-readable Word previews', () => {
  it('shows selected text, proposed replacement and requested formatting', () => {
    const before = state();
    const preview = buildWordPreview('word_edit', 'edit_selection', { new_text: 'New', bold: true }, before);
    expect(preview).toMatchObject({ kind: 'replace', scope: 'selection', before: 'Target', proposed: 'New', properties: [{ name: 'bold', value: 'true' }] });
    expect(before.text).toBe('Target\rKeep\r');
  });
  it('does not label whole-document formatting as cached-selection formatting', () => {
    const before = state();
    before.selection.text = '';
    const preview = buildWordPreview('word_format', 'format_selection', { bold: true }, before, 'stale cached selection');
    expect(preview.scope).toBe('document');
    expect(preview.before).toBe(before.text);
  });
  it('warns when long previews or tables are shortened', () => {
    const before = state();
    before.text = 'x'.repeat(20_000);
    const preview = buildWordPreview('word_structure', 'create_table', { data: Array.from({ length: 25 }, () => ['1']) }, before);
    expect(preview.truncated).toBe(true);
    expect(preview.before).toHaveLength(16_000);
    expect(preview.table).toHaveLength(20);
  });
  it('includes administrative formatting defaults and detects stale approval targets', () => {
    const before = state();
    const preview = buildWordPreview('word_format', 'format_administrative_vn', {}, before);
    expect(preview.properties).toContainEqual({ name: 'font_name', value: 'Times New Roman' });
    expect(sameApprovalState(before, state(), 'Target', 'Target')).toBe(true);
    const moved = state();
    moved.selection.text = 'Keep';
    expect(sameApprovalState(before, moved, 'Target', 'Target')).toBe(false);
    expect(sameApprovalState(before, state(), 'Target', 'Keep')).toBe(false);
  });
  it('retains the actual range across focus changes and releases it after use', async () => {
    const range = { text: 'Repeated text', load: vi.fn(), track: vi.fn(), untrack: vi.fn(), select: vi.fn() };
    const context = { document: { getSelection: () => range }, sync: vi.fn(async () => {}) };
    const run = vi.fn(async (...args: any[]) => args[args.length - 1](context));
    vi.stubGlobal('Word', { run });
    const pinned = await pinWordSelection();
    expect(pinned).toBe(range);
    expect(range.track).toHaveBeenCalledOnce();
    expect(range.select).not.toHaveBeenCalled();
    await selectPinnedRange(pinned!);
    expect(run.mock.calls[run.mock.calls.length - 1]?.[0]).toBe(range);
    expect(range.select).toHaveBeenCalledOnce();
    await releasePinnedRange(pinned!);
    expect(range.untrack).toHaveBeenCalledOnce();
  });
});
