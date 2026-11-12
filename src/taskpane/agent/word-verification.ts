import type { Sandbox } from '../executor/sandbox.ts';
import { normalizeWordArgs } from './word-formatting';

export interface WordVerification {
  status: 'checking' | 'passed' | 'failed' | 'unverified';
  detail: string;
}

export interface ObservedFont {
  name: string | null;
  size: number | null;
  bold: boolean | null;
  italic: boolean | null;
  underline: string | null;
  color: string | null;
  highlightColor: string | null;
}
export interface ObservedParagraph {
  text: string;
  alignment: string;
  font: ObservedFont;
}
export interface WordObservation {
  text: string;
  paragraphs: ObservedParagraph[];
  selection: { text: string; font: ObservedFont };
  tables: string[][][];
  matches?: Array<{ text: string; font: ObservedFont }>;
  matchCount?: number;
}

/** Serialized into a fresh Word.run: no cached tool output or selection is trusted. */
async function observeWord(context: any, query = ''): Promise<WordObservation> {
  const body = context.document.body;
  const selection = context.document.getSelection();
  const paragraphs = body.paragraphs;
  const tables = body.tables;
  const fontProps = 'font/name,font/size,font/bold,font/italic,font/underline,font/color,font/highlightColor';
  body.load('text');
  selection.load(`text,${fontProps}`);
  paragraphs.load(`text,alignment,${fontProps}`);
  tables.load('values');
  const matches = query && query.length <= 255 && !/[\r\n]/.test(query)
    ? body.search(query, { matchCase: true, matchWildcards: false }) : undefined;
  if (matches) matches.load(`text,${fontProps}`);
  await context.sync();
  const font = (f: any): ObservedFont => ({
    name: f.name, size: f.size, bold: f.bold, italic: f.italic,
    underline: f.underline, color: f.color, highlightColor: f.highlightColor,
  });
  return {
    text: body.text,
    paragraphs: paragraphs.items.map((p: any) => ({ text: p.text, alignment: p.alignment, font: font(p.font) })),
    selection: { text: selection.text, font: font(selection.font) },
    tables: tables.items.map((t: any) => t.values),
    ...(matches ? { matchCount: matches.items.length, matches: matches.items.slice(0, 50).map((r: any) => ({ text: r.text, font: font(r.font) })) } : {}),
  };
}

export async function observeWordDocument(sandbox: Sandbox, timeout: number, query = ''): Promise<WordObservation> {
  const script = `// [AutoOffice] verify.read_document\nreturn await (${observeWord.toString()})(context, ${JSON.stringify(query)});`;
  const result = await sandbox.execute(script, timeout);
  if (!result.success) throw new Error(result.error || 'Word read-back failed');
  const data = typeof result.output === 'string' ? JSON.parse(result.output) : result.output;
  if (!data || typeof data.text !== 'string' || !Array.isArray(data.paragraphs) ||
      !Array.isArray(data.tables) || typeof data.selection?.text !== 'string' || !data.selection.font ||
      !data.paragraphs.every((p: any) => typeof p.text === 'string' && p.font) ||
      !data.tables.every((t: unknown) => Array.isArray(t) && t.every(r => Array.isArray(r) && r.every(c => typeof c === 'string')))) {
    throw new Error('Invalid Word read-back data');
  }
  return data;
}

export function isWordMutation(tool: string, op: string, args: Record<string, unknown>): boolean {
  if (tool === 'word_read' || tool === 'word_help') return false;
  if (tool === 'word_review' && op === 'manage_comment' && args.action === 'list') return false;
  if (tool === 'word_document' && op === 'export_table') return false;
  if (tool === 'word_app' && ['live_scroll_to', 'com_readability_statistics'].includes(op)) return false;
  return true;
}

const text = (value: string) => value.normalize('NFC').replace(/\r\n?/g, '\n');
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
const passed = (detail: string): WordVerification => ({ status: 'passed', detail });
const failed = (detail: string): WordVerification => ({ status: 'failed', detail });
const unverified = (detail: string): WordVerification => ({ status: 'unverified', detail });
const fontKeys = ['font_name', 'font_size', 'bold', 'italic', 'underline', 'color', 'highlight_color'];

function fontMatches(font: ObservedFont, fmt: Record<string, unknown>, requested: string[]): boolean {
  return requested.every(key => {
    if (key === 'font_name') return font.name?.toLowerCase() === String(fmt[key]).toLowerCase();
    if (key === 'font_size') return typeof font.size === 'number' && Math.abs(font.size - Number(fmt[key])) <= 0.05;
    if (key === 'bold' || key === 'italic') return font[key] === Boolean(fmt[key]);
    if (key === 'underline') return font.underline === (fmt[key] ? 'Single' : 'None');
    const observed = key === 'color' ? font.color : font.highlightColor;
    return color(observed) === color(fmt[key]);
  });
}

/** Check only properties we can measure. An observed change alone is not proof of correctness. */
export function verifyWordMutation(
  tool: string, op: string, args: Record<string, unknown>,
  before: WordObservation, after: WordObservation, cachedSelection = '',
): WordVerification {
  args = normalizeWordArgs(args);
  if (tool === 'word_edit' && ['edit_selection', 'replace_selection_or_document'].includes(op)) {
    const target = text(String(args.find || before.selection.text || cachedSelection));
    const action = String(args.action || 'replace');
    if (!target || !['replace', 'delete'].includes(action)) return unverified('Cannot uniquely identify the edited range.');
    const original = text(before.text);
    const index = original.indexOf(target);
    if (index < 0 || original.indexOf(target, index + 1) >= 0) {
      return unverified('The target is missing or appears more than once; the exact edited range cannot be proven.');
    }
    const replacement = action === 'delete' ? '' : text(String(args.new_text ?? args.text ?? args.replace ?? ''));
    const expected = original.slice(0, index) + replacement + original.slice(index + target.length);
    if (text(after.text) !== expected) return failed('The edited text or the surrounding document does not match the expected replacement.');
    const requestedFont = fontKeys.filter(key => args[key] !== undefined);
    if (requestedFont.length && action !== 'delete') {
      const paragraphs = after.paragraphs.filter(p => text(p.text) === replacement);
      const font = paragraphs.length === 1 ? paragraphs[0].font
        : after.matchCount === 1 && text(after.matches?.[0]?.text ?? '') === replacement ? after.matches![0].font : undefined;
      if (!font) return unverified('Text replacement matches, but the exact formatted replacement range is ambiguous.');
      if (!fontMatches(font, args, requestedFont)) return failed('Replacement text matches, but its requested formatting does not.');
      return passed('Replacement text, surrounding text and requested replacement font properties match.');
    }
    return passed('The edited text and surrounding text match the expected replacement.');
  }
  if (tool === 'word_edit' && op === 'set_paragraph_text') {
    const index = Number(args.paragraph_index ?? args.index);
    const desired = String(args.new_text ?? args.text ?? '');
    if (!Number.isInteger(index) || !before.paragraphs[index] || /[\r\n]/.test(desired)) {
      return unverified('Cannot verify this paragraph replacement with paragraph-level read-back.');
    }
    const expected = before.paragraphs.map((p, i) => i === index ? desired : p.text);
    if (!same(after.paragraphs.map(p => p.text), expected)) return failed('The paragraph text or surrounding paragraph text does not match.');
    const requestedFont = fontKeys.filter(key => args[key] !== undefined);
    if (requestedFont.length && !fontMatches(after.paragraphs[index].font, args, requestedFont)) return failed('Paragraph text matches, but its requested formatting does not.');
    return passed('Requested paragraph text matches and surrounding paragraph text is unchanged.');
  }
  if (tool === 'word_edit' && op === 'edit_by_format' && String(args.format || args.target || 'bold') === 'bold') {
    const action = String(args.action || 'replace');
    const replacement = String(args.new_text ?? args.text ?? args.replace ?? '');
    if (action !== 'replace' || !replacement || /[\r\n]/.test(replacement)) return unverified('This format-based operation needs a manual result review.');
    const indices = before.paragraphs.map((p, i) => p.font.bold === true && p.text.trim() ? i : -1).filter(i => i >= 0);
    if (!indices.length) return failed('No bold target paragraph was present before the edit.');
    const expected = before.paragraphs.map((p, i) => indices.includes(i) ? replacement : p.text);
    if (!same(expected, after.paragraphs.map(p => p.text))) return failed('Format-based replacement changed unexpected text.');
    if (indices.some(i => after.paragraphs[i].font.bold !== (args.bold === undefined ? false : Boolean(args.bold)))) return failed('The one-use bold marker was not cleared or the requested bold state does not match.');
    return passed('Bold-marked text was replaced and its one-use marker was cleared unless bold was explicitly requested.');
  }
  if (tool === 'word_format' && ['format_selection', 'set_paragraph_format', 'format_text'].includes(op)) {
    if (text(before.text) !== text(after.text)) return failed('Formatting unexpectedly changed document text.');
    const fmt = { ...args, ...((args.formatting as Record<string, unknown>) || {}) };
    const supported = fontKeys;
    const requested = supported.filter(key => fmt[key] !== undefined);
    // Do not certify properties that were not measured, including paragraph layout and styles.
    const controlKeys = ['indices', 'index', 'paragraph_index', 'formatting', 'find', 'query'];
    const unsupported = Object.keys(fmt).filter(key => !supported.includes(key) && !controlKeys.includes(key));
    let fonts: ObservedFont[];
    if (op === 'format_selection' && before.selection.text && text(after.selection.text) === text(before.selection.text)) {
      fonts = [after.selection.font];
    } else if (['format_selection', 'format_text'].includes(op) && after.matches?.length && after.matchCount === after.matches.length &&
        (op === 'format_text' || after.matchCount === 1)) {
      fonts = after.matches.map(match => match.font);
    } else if (op === 'set_paragraph_format' && Array.isArray(args.indices) && args.indices.length > 0) {
      const indices = args.indices.map(Number);
      if (indices.some(i => !Number.isInteger(i) || !after.paragraphs[i])) return failed('The requested paragraph index is not present.');
      fonts = indices.map(i => after.paragraphs[i].font);
    } else if (op === 'format_selection' && !before.selection.text && !cachedSelection && after.paragraphs.length) {
      fonts = after.paragraphs.map(p => p.font);
    } else return unverified('No stable selected range or explicit paragraph indices were available for formatting verification.');
    if (!requested.length) return unverified('No supported formatting properties were requested.');
    const mismatch = fonts.some(font => !fontMatches(font, fmt, requested));
    if (mismatch) return failed('Read-back formatting does not match the requested formatting.');
    if (unsupported.length) return unverified(`Font properties match, but these properties were not measured: ${unsupported.join(', ')}.`);
    return passed('Requested font properties match and document text is unchanged.');
  }
  if (tool === 'word_structure' && op === 'create_table' && Array.isArray(args.data)) {
    const index = String(args.location || 'end').toLowerCase() === 'start' ? 0 : after.tables.length - 1;
    const inserted = after.tables[index];
    const remaining = after.tables.filter((_, i) => i !== index);
    if (!inserted || !same(remaining, before.tables) || after.tables.length !== before.tables.length + 1) {
      return failed('Expected exactly one new table while preserving existing tables.');
    }
    const normalizeCells = (rows: string[][]) => rows.map(row => row.map(cell => text(cell).replace(/[\n\u0007]+$/, '')));
    const expected = args.data;
    if (!expected.every(row => Array.isArray(row) && row.every(cell => typeof cell === 'string'))) {
      return unverified('Expected table cells must be strings.');
    }
    if (!same(normalizeCells(inserted), normalizeCells(expected as string[][]))) return failed('Table cell values do not match the requested data.');
    if (['font_name', 'font_size', 'header_bg', 'header_row'].some(key => args[key] !== undefined)) {
      return unverified('Table cell values match; table styling was not measured.');
    }
    return passed('Exactly one table was inserted with the requested rows, columns and cell values.');
  }
  return unverified('Document was read back, but this operation has no deterministic postcondition check yet. Review the observed document before claiming completion.');
}

function color(value: unknown): string {
  const key = String(value ?? '').toLowerCase();
  const names: Record<string, string> = { black: '#000000', white: '#ffffff', red: '#ff0000',
    yellow: '#ffff00', blue: '#0000ff', green: '#008000', none: '', null: '' };
  if (/^#[a-f0-9]{3}$/.test(key)) return '#' + [...key.slice(1)].map(c => c + c).join('');
  return names[key] ?? key;
}

export function wordReadBackForModel(observation: WordObservation): string {
  const limit = 12_000;
  const full = JSON.stringify({ text: observation.text, paragraphs: observation.paragraphs, tables: observation.tables,
    matches: observation.matches, matchCount: observation.matchCount });
  return full.length > limit ? `${full.slice(0, limit)}\n[Read-back truncated; use word_read for specific details.]` : full;
}
