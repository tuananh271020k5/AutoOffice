import type { WordObservation } from './word-verification';

export interface WordChangePreview {
  kind: 'replace' | 'delete' | 'insert' | 'format' | 'table' | 'other';
  scope: 'selection' | 'paragraphs' | 'document';
  before: string;
  proposed?: string;
  properties: Array<{ name: string; value: string }>;
  table?: string[][];
  truncated: boolean;
}

const LIMIT = 16_000;
/** Describes the requested operation; never executes a draft on the live document. */
export function buildWordPreview(tool: string, op: string, args: Record<string, unknown>, before: WordObservation, cached = ''): WordChangePreview {
  const fmt = { ...args, ...((args.formatting as Record<string, unknown>) || {}) };
  if (op === 'format_administrative_vn') {
    fmt.font_name ??= 'Times New Roman';
    fmt.font_size ??= 13;
    fmt.line_spacing_lines ??= 1.3;
  }
  const names = ['font_name', 'font_size', 'bold', 'italic', 'underline', 'color', 'highlight_color',
    'alignment', 'space_before', 'space_after', 'line_spacing_lines'];
  const preview: WordChangePreview = {
    kind: 'other', scope: 'document', before: before.text, properties: [], truncated: false,
  };
  const indices = Array.isArray(args.indices) ? args.indices.map(Number)
    : args.paragraph_index !== undefined || args.index !== undefined ? [Number(args.paragraph_index ?? args.index)] : [];
  const selected = op === 'format_selection' ? before.selection.text : String(args.find || before.selection.text || cached);
  if (op === 'edit_selection' || op === 'replace_selection_or_document' || op === 'format_selection') {
    // The executor can fall back to the document when no selection exists. Make that scope visible.
    preview.scope = selected ? 'selection' : 'document';
    preview.before = selected || before.text;
  } else if (indices.length) {
    preview.scope = 'paragraphs';
    preview.before = indices.map(i => before.paragraphs[i]?.text ?? '').join('\n');
  }
  if (tool === 'word_edit') {
    preview.kind = args.action === 'delete' || op === 'delete_paragraphs' ? 'delete'
      : op === 'insert_paragraphs' ? 'insert' : 'replace';
    if (op === 'insert_paragraphs') {
      const paragraphs = Array.isArray(args.paragraphs) ? args.paragraphs : [args.text ?? ''];
      preview.proposed = paragraphs.map(p => typeof p === 'string' ? p : String(p?.text ?? '')).join('\n');
    } else if (Array.isArray(args.replacements)) {
      preview.proposed = args.replacements.map(r => `${String(r.find ?? '')} → ${String(r.replace ?? '')}`).join('\n');
    } else if (args.action === 'delete' || op === 'delete_paragraphs') {
      preview.proposed = '';
    } else if (args.new_text !== undefined || args.text !== undefined || args.replace !== undefined) {
      preview.proposed = String(args.new_text ?? args.text ?? args.replace ?? '');
    }
  } else if (tool === 'word_format') preview.kind = 'format';
  if (tool === 'word_structure' && op === 'create_table' && Array.isArray(args.data)) {
    preview.kind = 'table';
    preview.table = args.data.slice(0, 20).map(row => Array.isArray(row) ? row.slice(0, 10).map(String) : []);
    if (args.data.length > 20 || args.data.some(row => Array.isArray(row) && row.length > 10)) preview.truncated = true;
  }
  preview.properties = names.filter(name => fmt[name] !== undefined).map(name => ({ name, value: String(fmt[name]) }));
  if (preview.before.length > LIMIT || (preview.proposed?.length ?? 0) > LIMIT) preview.truncated = true;
  preview.before = preview.before.slice(0, LIMIT);
  if (preview.proposed !== undefined) preview.proposed = preview.proposed.slice(0, LIMIT);
  return preview;
}

/** Approval is invalid once the document or the live/cached target changed. */
export function sameApprovalState(a: WordObservation, b: WordObservation, cachedBefore: string, cachedNow: string): boolean {
  return JSON.stringify(a) === JSON.stringify(b) && cachedBefore === cachedNow;
}

/** Retains the actual range, including repeated text, while focus moves into the task pane. */
export async function pinWordSelection(): Promise<Word.Range | undefined> {
  if (typeof Word === 'undefined') return undefined;
  return Word.run(async context => {
    const range = context.document.getSelection();
    range.load('text');
    await context.sync();
    if (!range.text) return undefined;
    range.track();
    await context.sync();
    return range;
  });
}

export async function selectPinnedRange(range: Word.Range): Promise<void> {
  await Word.run(range, async context => { range.select(); await context.sync(); });
}

export async function releasePinnedRange(range: Word.Range): Promise<void> {
  try { await Word.run(range, async context => { range.untrack(); await context.sync(); }); }
  catch { /* A document/pane may have closed; no document edits are performed here. */ }
}
