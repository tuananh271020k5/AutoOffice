/** Normalize model argument variants once for execution, preview and verification. */
export function normalizeWordArgs(args: Record<string, unknown>): Record<string, unknown> {
  const result = { ...args };
  if (Array.isArray(args.replacements)) result.replacements = args.replacements.map(item => normalizeWordArgs(item as Record<string, unknown>));
  const fmt = { ...args, ...((args.formatting as Record<string, unknown>) || {}) };
  const aliases: Record<string, string[]> = {
    font_name: ['fontName'], font_size: ['fontSize', 'size'], highlight_color: ['highlightColor'],
  };
  for (const [key, names] of Object.entries(aliases)) {
    const value = fmt[key] ?? names.map(name => fmt[name]).find(value => value !== undefined);
    if (value !== undefined) result[key] = value;
    for (const name of names) delete result[name];
  }
  for (const key of ['bold', 'italic', 'underline', 'color']) if (fmt[key] !== undefined) result[key] = fmt[key];
  if (result.font_size !== undefined) {
    const size = Number(result.font_size);
    if (!Number.isFinite(size) || size < 1 || size > 1638) throw new Error('Font size must be between 1 and 1638 pt.');
    result.font_size = size;
  }
  // Keep non-font paragraph properties, but remove aliases that would confuse verification.
  if (args.formatting) {
    const formatting = { ...(args.formatting as Record<string, unknown>) };
    for (const names of Object.values(aliases)) for (const name of names) delete formatting[name];
    for (const key of Object.keys(aliases)) if (result[key] !== undefined) formatting[key] = result[key];
    result.formatting = formatting;
  }
  return result;
}
