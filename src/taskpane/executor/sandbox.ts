// src/taskpane/executor/sandbox.ts
import type { HostKind } from '../host/context.ts';

export interface ExecutionResult {
  success: boolean;
  output?: unknown;
  error?: string;
  stack?: string;
  logs?: string[];
  debugInfo?: {
    code?: string;
    errorLocation?: string;
    statement?: string;
    surroundingStatements?: string[];
    fullStatements?: string[];
    message?: string;
  };
}

const NS: Record<HostKind, 'Word' | 'Excel' | 'PowerPoint'> = {
  word: 'Word',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
};

const formatArg = (a: unknown): string => {
  if (typeof a === 'string') return a;
  if (a instanceof Error) return a.stack || a.message;
  try {
    return JSON.stringify(a);
  } catch {
    return String(a);
  }
};

const makeCapturingConsole = (logs: string[]) => ({
  log:   (...args: unknown[]) => logs.push(args.map(formatArg).join(' ')),
  info:  (...args: unknown[]) => logs.push('[info] '  + args.map(formatArg).join(' ')),
  warn:  (...args: unknown[]) => logs.push('[warn] '  + args.map(formatArg).join(' ')),
  error: (...args: unknown[]) => logs.push('[error] ' + args.map(formatArg).join(' ')),
  debug: (...args: unknown[]) => logs.push('[debug] ' + args.map(formatArg).join(' ')),
});

/**
 * Normalizes common LLM mistakes in generated Word office.js code before execution:
 * - `body.insertParagraph(..., Word.InsertLocation.after)` -> `Word.InsertLocation.end`
 * - `body.insertParagraph(..., Word.InsertLocation.before)` -> `Word.InsertLocation.start`
 * - `insertParagraph("\n", ...)` -> `insertParagraph("", ...)`
 * - lowercase alignment strings (`"centered"`, `"left"`, `"right"`, `"justified"`) -> `Word.Alignment.*`
 * - `.lineSpacing = X` where `0 < X <= 4` (lines instead of points) -> `X * 12` points
 * - `.pagesetup` -> `.pageSetup` and `context.document.pageSetup` -> `context.document.sections.getFirst().pageSetup`
 * - `body.insertField / getBookmarks / insertBookmark / getHyperlinkRanges / insertFootnote / insertEndnote / insertComment` -> forwarded via `body.getRange(...)`
 * - `.temporary = true` -> `.removeWhenEdited = true`
 */
export function normalizeWordCode(code: string): string {
  return code
    // Fix body.insertParagraph / insertText / insertHtml / insertTable / insertBreak with .after / .before
    .replace(
      /\b(body|context\.document\.body)\.(insertParagraph|insertText|insertHtml|insertTable|insertBreak)\(([^)]*?),\s*(?:Word\.InsertLocation\.after|['"]After['"]|['"]after['"])\s*\)/g,
      '$1.$2($3, Word.InsertLocation.end)',
    )
    .replace(
      /\b(body|context\.document\.body)\.(insertParagraph|insertText|insertHtml|insertTable|insertBreak)\(([^)]*?),\s*(?:Word\.InsertLocation\.before|['"]Before['"]|['"]before['"])\s*\)/g,
      '$1.$2($3, Word.InsertLocation.start)',
    )
    // Fix insertParagraph("\n", ...) which throws InvalidArgument in Word
    .replace(/\.insertParagraph\(\s*(?:["']\\n["']|["']\r?\n["'])\s*,/g, '.insertParagraph("",')
    // Fix lowercase alignment string literals
    .replace(/\.alignment\s*=\s*['"](?:centered|center)['"]/gi, '.alignment = Word.Alignment.centered')
    .replace(/\.alignment\s*=\s*['"]left['"]/gi, '.alignment = Word.Alignment.left')
    .replace(/\.alignment\s*=\s*['"]right['"]/gi, '.alignment = Word.Alignment.right')
    .replace(/\.alignment\s*=\s*['"](?:justified|justify)['"]/gi, '.alignment = Word.Alignment.justified')
    // Fix lineSpacing when specified in lines (0 < X <= 4) instead of points (1 line = 12 pt)
    .replace(/\.lineSpacing\s*=\s*([0-9]+(?:\.[0-9]+)?)\b/g, (match, numStr) => {
      const val = parseFloat(numStr);
      if (val > 0 && val <= 4) {
        const pt = Math.round(val * 12 * 100) / 100;
        return `.lineSpacing = ${pt}`;
      }
      return match;
    })
    // Fix .pagesetup (lowercase typo) -> .pageSetup
    .replace(/\.pagesetup\b/gi, '.pageSetup')
    // Fix context.document.pageSetup or body.pageSetup -> context.document.sections.getFirst().pageSetup
    .replace(/\bcontext\.document\.pageSetup\b/g, 'context.document.sections.getFirst().pageSetup')
    .replace(/\b(?:body|context\.document\.body)\.pageSetup\b/g, 'context.document.sections.getFirst().pageSetup')
    // Fix Range-only methods accidentally called on body / context.document.body
    .replace(/\b(body|context\.document\.body)\.insertField\s*\(/g, '$1.getRange(Word.RangeLocation.end).insertField(')
    .replace(/\b(body|context\.document\.body)\.getBookmarks\s*\(/g, '$1.getRange().getBookmarks(')
    .replace(/\b(body|context\.document\.body)\.insertBookmark\s*\(/g, '$1.getRange(Word.RangeLocation.end).insertBookmark(')
    .replace(/\b(body|context\.document\.body)\.getHyperlinkRanges\s*\(/g, '$1.getRange().getHyperlinkRanges(')
    .replace(/\b(body|context\.document\.body)\.insertFootnote\s*\(/g, '$1.getRange(Word.RangeLocation.end).insertFootnote(')
    .replace(/\b(body|context\.document\.body)\.insertEndnote\s*\(/g, '$1.getRange(Word.RangeLocation.end).insertEndnote(')
    .replace(/\b(body|context\.document\.body)\.insertComment\s*\(/g, '$1.getRange().insertComment(')
    // Fix contentControl.temporary -> contentControl.removeWhenEdited
    .replace(/\.temporary\s*=\s*(true|false)\b/g, '.removeWhenEdited = $1');
}

/**
 * Installs defensive runtime shims on global OfficeExtension / Word prototypes
 * so common LLM mistakes (like calling `.load()` on a `ClientResult`, or accessing
 * `section.pageSetup` on older Word builds) degrade gracefully instead of throwing.
 */
export function installWordRuntimeShims(): void {
  const g = globalThis as Record<string, any>;

  // 1. Allow .load() on OfficeExtension.ClientResult (no-op returning this)
  if (g.OfficeExtension?.ClientResult?.prototype && !g.OfficeExtension.ClientResult.prototype.load) {
    g.OfficeExtension.ClientResult.prototype.load = function () {
      return this;
    };
  }

  const WordNS = g.Word;
  if (!WordNS) return;

  // 2. Patch Word.Field.prototype.updateResult (no-op if missing)
  if (WordNS.Field?.prototype && !WordNS.Field.prototype.updateResult) {
    WordNS.Field.prototype.updateResult = function () {
      return this;
    };
  }

  // 3. Patch Word.TrackedChange.prototype.range -> this.getRange()
  if (WordNS.TrackedChange?.prototype && !('range' in WordNS.TrackedChange.prototype)) {
    Object.defineProperty(WordNS.TrackedChange.prototype, 'range', {
      configurable: true,
      get() {
        return typeof this.getRange === 'function' ? this.getRange() : undefined;
      },
    });
  }

  // 4. Patch Word.Body.prototype with Range-only methods if missing
  if (WordNS.Body?.prototype) {
    const bodyProto = WordNS.Body.prototype;
    const endLoc = WordNS.RangeLocation?.end ?? 'End';
    if (!bodyProto.insertField) {
      bodyProto.insertField = function (loc: any, type: any, text?: any, preserve?: any) {
        return this.getRange(endLoc).insertField(loc, type, text, preserve);
      };
    }
    if (!bodyProto.getBookmarks) {
      bodyProto.getBookmarks = function (includeHidden?: boolean, lowercase?: boolean) {
        return this.getRange().getBookmarks(includeHidden, lowercase);
      };
    }
    if (!bodyProto.insertBookmark) {
      bodyProto.insertBookmark = function (name: string) {
        return this.getRange(endLoc).insertBookmark(name);
      };
    }
    if (!bodyProto.getHyperlinkRanges) {
      bodyProto.getHyperlinkRanges = function () {
        return this.getRange().getHyperlinkRanges();
      };
    }
    if (!bodyProto.insertFootnote) {
      bodyProto.insertFootnote = function (text: string) {
        return this.getRange(endLoc).insertFootnote(text);
      };
    }
    if (!bodyProto.insertEndnote) {
      bodyProto.insertEndnote = function (text: string) {
        return this.getRange(endLoc).insertEndnote(text);
      };
    }
    if (!bodyProto.insertComment) {
      bodyProto.insertComment = function (text: string) {
        return this.getRange().insertComment(text);
      };
    }
  }

  // 5. Ensure Word.Section.prototype has a safe pageSetup / pagesetup fallback
  if (WordNS.Section?.prototype) {
    const secProto = WordNS.Section.prototype;
    const origDesc = Object.getOwnPropertyDescriptor(secProto, 'pageSetup');
    if (!origDesc || origDesc.configurable) {
      Object.defineProperty(secProto, 'pageSetup', {
        configurable: true,
        get() {
          const nativeVal = origDesc?.get ? origDesc.get.call(this) : this._pageSetup;
          if (nativeVal) return nativeVal;
          if (!this.__fallbackPageSetup) {
            this.__fallbackPageSetup = {
              orientation: WordNS.PageOrientation?.portrait ?? 'Portrait',
              paperSize: WordNS.PaperType?.a4 ?? 'A4',
              pageWidth: 595.3,
              pageHeight: 841.9,
              topMargin: 56.7,
              bottomMargin: 56.7,
              leftMargin: 85.05,
              rightMargin: 42.52,
              headerDistance: 36,
              footerDistance: 36,
              gutter: 0,
              load() {
                return this;
              },
              set(props: Record<string, unknown>) {
                Object.assign(this, props);
              },
            };
          }
          return this.__fallbackPageSetup;
        },
      });
    }
    if (!('pagesetup' in secProto)) {
      Object.defineProperty(secProto, 'pagesetup', {
        configurable: true,
        get() {
          return this.pageSetup;
        },
      });
    }
  }
}

export class Sandbox {
  constructor(private readonly host: HostKind) {}

  init(): void {}
  destroy(): void {}

  async execute(code: string, timeout: number = 30000): Promise<ExecutionResult> {
    const ns = NS[this.host];
    const otherNamespaces = Object.values(NS).filter((n) => n !== ns);
    const normalized = this.host === 'word' ? normalizeWordCode(code) : code;
    const trimmed = normalized.trim();

    // Reject code targeting the wrong host before running it. Yields a clear
    // error the agent can self-heal on, instead of an opaque "X is not defined".
    for (const other of otherNamespaces) {
      if (trimmed.startsWith(`${other}.run`)) {
        return {
          success: false,
          error: `Code uses ${other}.run but the add-in is running in ${ns}. Rewrite using ${ns}.run.`,
          logs: [],
        };
      }
    }

    if (this.host === 'word') {
      installWordRuntimeShims();
    }

    // Strip any leading comments before checking `await Word.run(...)` at the top level
    const withoutLeadingComments = trimmed.replace(/^(?:\s*(?:\/\/[^\n]*\n?|\/\*[\s\S]*?\*\/))*\s*/, '');
    const strippedAwait = withoutLeadingComments.replace(/^await\s+/, '');
    const isWrapped = strippedAwait.startsWith(`${ns}.run`);
    const execCode = isWrapped
      ? `return (${strippedAwait.replace(/;+\s*$/, '')});`
      : `return ${ns}.run(async function(context) {\n${normalized}\n});`;

    const logs: string[] = [];
    const capturingConsole = makeCapturingConsole(logs);

    const timeoutPromise = new Promise<ExecutionResult>((resolve) =>
      setTimeout(
        () => resolve({ success: false, error: `Execution timed out after ${timeout}ms`, logs }),
        timeout
      )
    );

    const executionPromise = (async (): Promise<ExecutionResult> => {
      try {
        const fn = new Function('console', execCode);
        const result = await fn(capturingConsole);
        return { success: true, output: result, logs };
      } catch (err) {
        const e = err as Error & { code?: string; debugInfo?: ExecutionResult['debugInfo'] };
        const isOfficeError = typeof e.code === 'string' && e.debugInfo !== undefined;
        if (isOfficeError) {
          const dbg = e.debugInfo!;
          return {
            success: false,
            error: `${e.code}: ${e.message || dbg.message || ''}`.trim(),
            stack: e.stack,
            debugInfo: { code: e.code, ...dbg },
            logs,
          };
        }
        return { success: false, error: e.message || String(err), stack: e.stack, logs };
      }
    })();

    return Promise.race([executionPromise, timeoutPromise]);
  }
}
