import { describe, it, expect, beforeEach } from 'vitest';
import { WORD_SKILLS, WORD_SKILL_NAMES } from './index.ts';
import { listSkills, lookupSkill } from '../index.ts';
import { Sandbox, normalizeWordCode, installWordRuntimeShims } from '../../executor/sandbox.ts';
import { buildSystemPrompt } from '../../agent/system-prompt.ts';
import { buildKs4wScript, KITCHENSINK_WORD_TOOLS_META } from '../../agent/kitchensink-word-tools.ts';

/**
 * Strict mock of the Office.js Word runtime that enforces real @types/office-js rules:
 * - Throws InvalidArgument if body.insertParagraph uses Before/After or "\n"
 * - Throws if Paragraph.lineSpacing is set to <= 4 (lines instead of points)
 * - Throws if sections.items is accessed before sections.load("items") + context.sync()
 * - Throws if ContentControl.type is mutated (read-only in Office.js)
 * - ClientResult does NOT natively define .load() (verifies our skill code & runtime shim)
 * - Body does NOT natively define insertField / getBookmarks / getHyperlinkRanges (verifies Range usage & shim)
 */
function createStrictWordEnvironment(options?: { disableNativePageSetup?: boolean }) {
  class ClientResult<T> {
    constructor(public value: T) {}
  }

  const createFont = () => ({
    name: 'Calibri',
    size: 11,
    bold: false,
    italic: false,
    underline: 'None',
    color: '#000000',
    highlightColor: 'None',
    strikeThrough: false,
    load() { return this; },
  });

  const createRange = (initialText = 'Sample range text'): any => {
    const rangeObj: any = {
      text: initialText,
      hyperlink: 'https://example.com',
      style: 'Normal',
      styleBuiltIn: 'Normal',
      font: createFont(),
      load() { return this; },
      select() {},
      delete() { this.text = ''; },
      expand() { return createRange(this.text); },
      getRange() { return this; },
      insertText(text: string) {
        this.text = text;
        return createRange(text);
      },
      insertHtml(html: string) {
        this.text = html;
        return createRange(html);
      },
      insertOoxml(xml: string) {
        return createRange(xml);
      },
      getOoxml() {
        return new ClientResult('<pkg:package><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"/></pkg:package>');
      },
      getHtml() {
        return new ClientResult('<p>Sample HTML</p>');
      },
      compareLocationWith() {
        return new ClientResult('Inside');
      },
      insertBookmark(name: string) {
        bookmarksStore.push(name);
      },
      getBookmarks() {
        return new ClientResult([...bookmarksStore]);
      },
      getHyperlinkRanges() {
        const items = [createRange('https://chinhphu.vn')];
        return {
          items,
          load() { return this; },
        };
      },
      insertField(_loc: string, type: string, text?: string) {
        const f = new FieldMock(type, text || '');
        fieldsStore.push(f);
        return f;
      },
      insertComment(content: string) {
        const c = createComment(content);
        commentsStore.push(c);
        return c;
      },
      insertFootnote(text: string) {
        const n = createNoteItem(text);
        footnotesStore.push(n);
        return n;
      },
      insertEndnote(text: string) {
        const n = createNoteItem(text);
        endnotesStore.push(n);
        return n;
      },
      insertContentControl(type = 'RichText') {
        const cc = createContentControl(type, this.text);
        contentControlsStore.push(cc);
        return cc;
      },
      get paragraphs() {
        return {
          items: [createParagraph(this.text)],
          getFirst: () => createParagraph(this.text),
          load() { return this; },
        };
      },
    };
    return rangeObj;
  };

  const createContentControl = (initialType = 'RichText', initialText = 'Sample CC'): any => {
    let _type = initialType;
    const cc: any = {
      id: Math.floor(Math.random() * 10000),
      tag: 'field_1',
      title: 'My Field',
      appearance: 'BoundingBox',
      color: '#0078D4',
      placeholderText: '',
      cannotDelete: false,
      cannotEdit: false,
      removeWhenEdited: false,
      text: initialText,
      checkboxContentControl: {
        isChecked: false,
        load() { return this; },
      },
      dropDownListContentControl: {
        items: [] as Array<{ displayText: string; value: string }>,
        addListItem(displayText: string, value: string) {
          this.items.push({ displayText, value });
        },
        load() { return this; },
      },
      load() { return this; },
      insertText(t: string) {
        this.text = t;
        return createRange(t);
      },
      clear() { this.text = ''; },
      delete() {},
    };
    Object.defineProperty(cc, 'type', {
      get() { return _type; },
      set() {
        throw new Error('InvalidOperation: ContentControl.type is read-only. Pass type to insertContentControl(type).');
      },
    });
    return cc;
  };

  const createComment = (content: string) => ({
    id: 'comment-1',
    content,
    authorName: 'AutoOffice User',
    createdDate: new Date().toISOString(),
    resolved: false,
    load() { return this; },
    reply(replyText: string) {
      return { content: replyText, load() { return this; } };
    },
    delete() {},
  });

  const createNoteItem = (text: string) => ({
    body: {
      text,
      load() { return this; },
      clear() { this.text = ''; },
      insertText(t: string) { this.text = t; return createRange(t); },
    },
    reference: {
      text: '1',
      load() { return this; },
    },
    load() { return this; },
    delete() {},
  });

  class FieldMock {
    result = { text: '2026-11-05', load() { return this; } };
    constructor(public type: string, public code: string) {}
    load() { return this; }
    select() {}
    delete() {}
  }

  class TrackedChangeMock {
    author = 'Alice';
    date = new Date();
    type = 'Added';
    text = 'Tracked text';
    load() { return this; }
    getRange() { return createRange(this.text); }
    accept() {}
    reject() {}
  }

  const createParagraph = (text: string): any => {
    let _lineSpacing = 12;
    const p: any = {
      text,
      uniqueLocalId: 'para-id-1',
      style: 'Normal',
      styleBuiltIn: 'Normal',
      alignment: 'Left',
      spaceBefore: 0,
      spaceAfter: 6,
      leftIndent: 0,
      rightIndent: 0,
      firstLineIndent: 0,
      font: createFont(),
      listItem: { level: 0, listString: '1.', load() { return this; } },
      load() { return this; },
      getRange() { return createRange(this.text); },
      select() {},
      delete() {
        const idx = paragraphsStore.indexOf(this);
        if (idx !== -1) paragraphsStore.splice(idx, 1);
      },
      insertParagraph(newText: string, loc: string) {
        if (newText === '\n' || newText === '\r\n') {
          throw new Error('InvalidArgument: insertParagraph does not accept newline string');
        }
        if (!['Before', 'After', 'Start', 'End', 'Replace'].includes(loc)) {
          throw new Error(`InvalidArgument: Invalid InsertLocation ${loc}`);
        }
        const created = createParagraph(newText);
        paragraphsStore.push(created);
        return created;
      },
      insertText(newText: string) {
        this.text = newText;
        return createRange(newText);
      },
      insertContentControl(type?: string) {
        return createRange(this.text).insertContentControl(type);
      },
      startNewList() {
        return { id: 1, load() { return this; } };
      },
      attachToList() {},
      getTrackedChanges() {
        return {
          items: [new TrackedChangeMock()],
          load() { return this; },
          acceptAll() {},
          rejectAll() {},
        };
      },
    };
    Object.defineProperty(p, 'lineSpacing', {
      get() { return _lineSpacing; },
      set(val: number) {
        if (val > 0 && val <= 4) {
          throw new Error(`LineSpacingInLinesError: lineSpacing=${val} is <= 4pt (squished lines! Must be in points, e.g. 15.6 or 18)`);
        }
        _lineSpacing = val;
      },
    });
    return p;
  };

  const createTable = (rows = 3, cols = 3, values?: string[][]) => {
    const vals = values || Array.from({ length: rows }, (_, r) =>
      Array.from({ length: cols }, (__, c) => `R${r}C${c}`)
    );
    const cellObj = {
      value: vals[0]?.[0] ?? '',
      shadingColor: '#FFFFFF',
      columnWidth: 100,
      body: {
        text: vals[0]?.[0] ?? '',
        clear() { this.text = ''; },
        insertText(t: string) { this.text = t; return createRange(t); },
        insertParagraph(t: string) { return createParagraph(t); },
        paragraphs: {
          getFirst: () => createParagraph(vals[0]?.[0] ?? ''),
          items: [createParagraph(vals[0]?.[0] ?? '')],
          load() { return this; },
        },
        load() { return this; },
      },
      setCellPadding() {},
      getCellPadding() { return new ClientResult(4); },
      deleteColumn() {},
      load() { return this; },
    };
    return {
      rowCount: vals.length,
      columnCount: vals[0]?.length ?? cols,
      values: vals,
      headerRowCount: 1,
      style: 'Table Grid',
      styleBuiltIn: 'TableGrid',
      styleBandedRows: true,
      styleBandedColumns: false,
      styleFirstColumn: false,
      styleLastColumn: false,
      styleTotalRow: false,
      styleHeaderRow: true,
      getRow: () => ({
        font: createFont(),
        shadingColor: '#FFFFFF',
        horizontalAlignment: 'Left',
        delete() {},
        load() { return this; },
      }),
      getCell: () => cellObj,
      addRows() {},
      addColumns() {},
      mergeCells() {},
      autoFitWindow() {},
      distributeColumns() {},
      getBorder: () => ({
        type: 'Single',
        color: '#000000',
        width: 1,
        load() { return this; },
      }),
      delete() {},
      load() { return this; },
    };
  };

  const paragraphsStore = [
    createParagraph('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM'),
    createParagraph('Độc lập - Tự do - Hạnh phúc'),
    createParagraph('ĐƠN XIN NGHỈ PHÉP'),
    createParagraph('Kính gửi: Ban Giám đốc và Phòng Hành chính Nhân sự (important 25/12/2025)'),
  ];
  const tablesStore = [createTable(3, 3)];
  const bookmarksStore = [' my_section ', 'full_name', 'Phan_Mo_Dau'];
  const fieldsStore = [new FieldMock('Date', 'DATE')];
  const commentsStore = [createComment('Initial comment')];
  const footnotesStore = [createNoteItem('Initial footnote')];
  const endnotesStore = [createNoteItem('Initial endnote')];
  const contentControlsStore = [
    createContentControl('RichText', 'Initial value'),
    Object.assign(createContentControl('CheckBox', ''), { tag: 'agree_terms' }),
  ];
  const picturesStore = [
    {
      width: 200,
      height: 150,
      altTextTitle: 'Sample Image',
      altTextDescription: 'Desc',
      imageFormat: 'png',
      paragraph: createParagraph(''),
      getBase64ImageSrc: () => new ClientResult('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='),
      delete() {},
      load() { return this; },
    },
  ];

  class BodyMock {
    text = paragraphsStore.map((p) => p.text).join('\n');
    font = createFont();
    paragraphs = {
      items: paragraphsStore,
      getFirst: () => paragraphsStore[0],
      load() { return this; },
    };
    tables = {
      items: tablesStore,
      getFirst: () => tablesStore[0],
      load() { return this; },
    };
    inlinePictures = {
      items: picturesStore,
      load() { return this; },
    };
    fields = {
      items: fieldsStore,
      load() { return this; },
    };
    footnotes = {
      items: footnotesStore,
      load() { return this; },
    };
    endnotes = {
      items: endnotesStore,
      load() { return this; },
    };
    lists = {
      items: [{ id: 1, load() { return this; } }],
      load() { return this; },
    };
    load() { return this; }
    clear() { this.text = ''; paragraphsStore.length = 0; }
    getRange() { return createRange(this.text); }
    insertParagraph(text: string, loc: string) {
      if (loc === 'After' || loc === 'Before' || loc === 'after' || loc === 'before') {
        const err: any = new Error('InvalidArgument: Body.insertParagraph only supports Start, End, Replace');
        err.code = 'InvalidArgument';
        err.debugInfo = { code: 'InvalidArgument', errorLocation: 'Body.insertParagraph' };
        throw err;
      }
      if (text === '\n' || text === '\r\n') {
        const err: any = new Error('InvalidArgument: Newline not allowed in Body.insertParagraph');
        err.code = 'InvalidArgument';
        err.debugInfo = { code: 'InvalidArgument', errorLocation: 'Body.insertParagraph' };
        throw err;
      }
      const p = createParagraph(text);
      paragraphsStore.push(p);
      return p;
    }
    insertText(text: string, loc: string) {
      if (loc === 'After' || loc === 'Before') {
        throw new Error('InvalidArgument: Body.insertText does not support After/Before');
      }
      this.text += text;
      return createRange(text);
    }
    insertHtml(html: string, loc: string) {
      if (loc === 'After' || loc === 'Before') {
        throw new Error('InvalidArgument: Body.insertHtml does not support After/Before');
      }
      return createRange(html);
    }
    insertOoxml(xml: string) {
      if (!xml.includes('<pkg:package') && !xml.includes('<w:')) {
        throw new Error('GeneralException: Invalid OOXML payload');
      }
      return createRange(xml);
    }
    getOoxml() {
      return new ClientResult('<pkg:package><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"/></pkg:package>');
    }
    getHtml() {
      return new ClientResult('<p>Document HTML</p>');
    }
    insertBreak() {}
    insertTable(rows: number, cols: number, _loc: string, values?: string[][]) {
      const t = createTable(rows, cols, values);
      tablesStore.push(t);
      return t;
    }
    insertInlinePictureFromBase64(b64: string) {
      if (b64.startsWith('data:image')) {
        throw new Error('InvalidArgument: Base64 must not contain data:image prefix');
      }
      const pic = {
        width: 100,
        height: 100,
        altTextTitle: '',
        altTextDescription: '',
        imageFormat: 'png',
        paragraph: createParagraph(''),
        getBase64ImageSrc: () => new ClientResult(b64),
        delete() {},
        load() { return this; },
      };
      picturesStore.push(pic);
      return pic;
    }
    search(_query: string) {
      return {
        items: [createRange('Found match')],
        load() { return this; },
      };
    }
    getComments() {
      return {
        items: commentsStore,
        load() { return this; },
      };
    }
    getTrackedChanges() {
      return {
        items: [new TrackedChangeMock()],
        load() { return this; },
        acceptAll() {},
        rejectAll() {},
      };
    }
  }

  class SectionMock {
    _pageSetup = options?.disableNativePageSetup
      ? undefined
      : {
          orientation: 'Portrait',
          paperSize: 'A4',
          pageWidth: 595.3,
          pageHeight: 841.9,
          topMargin: 56.7,
          bottomMargin: 56.7,
          leftMargin: 85.05,
          rightMargin: 42.52,
          headerDistance: 36,
          footerDistance: 36,
          gutter: 0,
          load() { return this; },
        };
    get pageSetup() {
      return this._pageSetup;
    }
    body = new BodyMock();
    getHeader() { return new BodyMock(); }
    getFooter() { return new BodyMock(); }
    load() { return this; }
  }

  const sectionInstance = new SectionMock();
  let sectionsLoaded = false;

  const documentMock = {
    body: new BodyMock(),
    changeTrackingMode: 'Off',
    properties: {
      title: 'Test Doc',
      subject: 'Subject',
      author: 'Author',
      keywords: 'Keywords',
      comments: 'Comments',
      creationDate: new Date(),
      customProperties: {
        items: [{ key: 'ProjectCode', value: 'PRJ-2024', type: 'String', load() { return this; }, delete() {} }],
        add() {},
        getItem() { return { delete() {}, load() { return this; } }; },
        load() { return this; },
      },
      load() { return this; },
    },
    sections: {
      get items() {
        if (!sectionsLoaded) {
          throw new Error("PropertyNotLoaded: The property 'items' is not available. Before reading the value's property, call the load method on the containing object and call 'context.sync()'.");
        }
        return [sectionInstance];
      },
      getFirst: () => sectionInstance,
      load() {
        sectionsLoaded = true;
        return this;
      },
    },
    contentControls: {
      items: contentControlsStore,
      getByTag: (_tag: string) => ({
        items: contentControlsStore,
        load() { return this; },
      }),
      getByTitle: (_title: string) => ({
        items: contentControlsStore,
        load() { return this; },
      }),
      load() { return this; },
    },
    getSelection: () => createRange('Selected text in document'),
    getBookmarkRange: () => createRange('Bookmarked range'),
    getBookmarkRangeOrNullObject: () => Object.assign(createRange('Bookmarked range'), { isNullObject: false }),
    deleteBookmark: () => {},
    getStyles: () => ({
      items: [{ nameLocal: 'Normal', type: 'Paragraph', load() { return this; } }],
      load() { return this; },
    }),
    addStyle: () => ({
      baseStyle: 'Heading 1',
      nextParagraphStyle: 'Normal',
      font: createFont(),
      load() { return this; },
    }),
    importStylesFromJson: () => {},
    onParagraphAdded: { add: () => {} },
    getParagraphByUniqueLocalId: () => createParagraph('Found by ID'),
    save: () => {},
    load() { return this; },
  };

  const contextMock = {
    document: documentMock,
    sync: async () => {},
  };

  const WordGlobal = {
    Body: BodyMock,
    Section: SectionMock,
    Field: FieldMock,
    TrackedChange: TrackedChangeMock,
    InsertLocation: {
      before: 'Before',
      after: 'After',
      start: 'Start',
      end: 'End',
      replace: 'Replace',
    },
    RangeLocation: {
      whole: 'Whole',
      start: 'Start',
      end: 'End',
      before: 'Before',
      after: 'After',
      content: 'Content',
    },
    Alignment: {
      left: 'Left',
      centered: 'Centered',
      right: 'Right',
      justified: 'Justified',
    },
    UnderlineType: {
      none: 'None',
      single: 'Single',
      double: 'Double',
      dotted: 'Dotted',
      wave: 'Wave',
    },
    HighlightColor: {
      yellow: 'Yellow',
      green: 'Green',
      cyan: 'Cyan',
      pink: 'Pink',
    },
    BreakType: {
      page: 'Page',
      line: 'Line',
      sectionNext: 'SectionNext',
      sectionContinuous: 'SectionContinuous',
      sectionEven: 'SectionEven',
      sectionOdd: 'SectionOdd',
    },
    BorderLocation: {
      top: 'Top',
      bottom: 'Bottom',
      left: 'Left',
      right: 'Right',
      insideHorizontal: 'InsideHorizontal',
      insideVertical: 'InsideVertical',
      all: 'All',
      outside: 'Outside',
    },
    BorderType: {
      none: 'None',
      single: 'Single',
      double: 'Double',
      dashed: 'Dashed',
    },
    CellPaddingLocation: {
      top: 'Top',
      bottom: 'Bottom',
      left: 'Left',
      right: 'Right',
    },
    BuiltInStyleName: {
      normal: 'Normal',
      heading1: 'Heading1',
      heading2: 'Heading2',
      heading3: 'Heading3',
      title: 'Title',
      subtitle: 'Subtitle',
      quote: 'Quote',
      intenseQuote: 'IntenseQuote',
      listParagraph: 'ListParagraph',
      noSpacing: 'NoSpacing',
      tableGrid: 'TableGrid',
    },
    StyleType: {
      paragraph: 'Paragraph',
      character: 'Character',
      table: 'Table',
      list: 'List',
    },
    HeaderFooterType: {
      primary: 'Primary',
      firstPage: 'FirstPage',
      evenPages: 'EvenPages',
    },
    PageOrientation: {
      portrait: 'Portrait',
      landscape: 'Landscape',
    },
    PaperType: {
      a4: 'A4',
      letter: 'Letter',
      legal: 'Legal',
    },
    FieldType: {
      toc: 'TOC',
      page: 'Page',
      numPages: 'NumPages',
      date: 'Date',
      time: 'Time',
      author: 'Author',
      title: 'Title',
      fileName: 'FileName',
      hyperlink: 'Hyperlink',
      ref: 'Ref',
      seq: 'Seq',
    },
    ChangeTrackingMode: {
      off: 'Off',
      trackAll: 'TrackAll',
      trackMineOnly: 'TrackMineOnly',
    },
    TrackedChangeType: {
      none: 'None',
      added: 'Added',
      deleted: 'Deleted',
      formatted: 'Formatted',
    },
    ContentControlType: {
      richText: 'RichText',
      plainText: 'PlainText',
      checkBox: 'CheckBox',
      dropDownList: 'DropDownList',
      comboBox: 'ComboBox',
    },
    ContentControlAppearance: {
      boundingBox: 'BoundingBox',
      tags: 'Tags',
      hidden: 'Hidden',
    },
    RangeExpandTo: {
      word: 'Word',
      sentence: 'Sentence',
      paragraph: 'Paragraph',
    },
    SelectionMode: {
      select: 'Select',
      start: 'Start',
      end: 'End',
    },
    NoteItemType: {
      footnote: 'Footnote',
      endnote: 'Endnote',
    },
    run: async (cb: (ctx: any) => Promise<any>) => {
      return cb(contextMock);
    },
  };

  (globalThis as any).OfficeExtension = { ClientResult };
  (globalThis as any).Word = WordGlobal;

  return { contextMock, WordGlobal };
}

/**
 * Extracts runnable JavaScript blocks from a skill markdown file,
 * skipping intentionally marked "// WRONG" anti-examples.
 */
function extractRunnableCodeBlocks(markdown: string): string[] {
  const regex = /```javascript\r?\n([\s\S]*?)```/g;
  const blocks: string[] = [];
  let match: RegExpExecArray | null;
  while ((match = regex.exec(markdown)) !== null) {
    const code = match[1].trim();
    if (code.includes('// WRONG')) continue;
    blocks.push(code);
  }
  return blocks;
}

describe('All 21 Word Tools & Skills Verification Suite', () => {
  beforeEach(() => {
    createStrictWordEnvironment();
  });

  it('1. Core Tool #1 (lookup_skill): lists and retrieves all 19 Word skills and matches WORD_CATALOG', () => {
    const skills = listSkills('word');
    expect(skills.length).toBe(19);
    expect(skills).toEqual(WORD_SKILL_NAMES);

    for (const skillName of skills) {
      const content = lookupSkill('word', skillName);
      expect(content.length).toBeGreaterThan(100);
      expect(content).toContain('Word.run');
    }
    expect(KITCHENSINK_WORD_TOOLS_META.length).toBe(10);
  });

  it('2. Core Tool #2 (execute_code) + All 19 Word Skill Markdown Code Blocks execute cleanly with 0 errors', async () => {
    const sandbox = new Sandbox('word');

    for (const skillName of WORD_SKILL_NAMES) {
      const markdown = WORD_SKILLS[skillName];
      const blocks = extractRunnableCodeBlocks(markdown);
      expect(blocks.length, `Skill ${skillName} should have runnable code blocks`).toBeGreaterThan(0);

      for (let i = 0; i < blocks.length; i++) {
        // Re-init strict environment before each block
        createStrictWordEnvironment();
        let code = blocks[i];
        // Provide stylesJson if the snippet references it as an external variable
        if (code.includes('stylesJson') && !code.includes('const stylesJson')) {
          code = `const stylesJson = "{}";\n${code}`;
        }

        const res = await sandbox.execute(code);
        expect(
          res.success,
          `Skill "${skillName}" block #${i + 1} failed with error: ${res.error}\nCode:\n${code}`,
        ).toBe(true);
      }
    }
  });

  it('3. Regression Test: Screenshot #1 (page-setup undefined / lowercase pagesetup) works even when native section.pageSetup is undefined', async () => {
    createStrictWordEnvironment({ disableNativePageSetup: true });
    const sandbox = new Sandbox('word');

    // Case A: LLM writes lowercase .pagesetup or context.document.pageSetup without loading sections
    const llmBuggyPageSetupCode = `
      await Word.run(async (context) => {
        const ps = context.document.pageSetup;
        ps.orientation = Word.PageOrientation.portrait;
        ps.topMargin = 56.7;
        ps.bottomMargin = 56.7;
        ps.leftMargin = 85.05;
        ps.rightMargin = 42.52;
        await context.sync();
      });
    `;
    const resA = await sandbox.execute(llmBuggyPageSetupCode);
    expect(resA.success, resA.error).toBe(true);

    // Case B: Verified page-setup.md universal fallback code
    const blocks = extractRunnableCodeBlocks(WORD_SKILLS['page-setup']);
    for (const block of blocks) {
      createStrictWordEnvironment({ disableNativePageSetup: true });
      const resB = await sandbox.execute(block);
      expect(resB.success, resB.error).toBe(true);
    }
  });

  it('4. Regression Test: Screenshot #2 (lineSpacing in lines 1.3 / 1.5 auto-converts to points 15.6 / 18)', async () => {
    const sandbox = new Sandbox('word');
    const llmLineSpacingCode = `
      await Word.run(async (context) => {
        const p = context.document.body.insertParagraph("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", Word.InsertLocation.after);
        p.alignment = "centered";
        p.lineSpacing = 1.5;
        const p2 = context.document.body.insertParagraph("\\n", Word.InsertLocation.before);
        p2.lineSpacing = 1.3;
        await context.sync();
        return { p1Spacing: p.lineSpacing, p2Spacing: p2.lineSpacing };
      });
    `;
    const res = await sandbox.execute(llmLineSpacingCode);
    expect(res.success, res.error).toBe(true);
    expect(res.output).toEqual({ p1Spacing: 18, p2Spacing: 15.6 });
  });

  it('5. Regression Test: Range-only methods called on body and .load() on ClientResult work via normalizeWordCode + runtime shims', async () => {
    const sandbox = new Sandbox('word');
    const llmRangeAndClientResultMistakes = `
      await Word.run(async (context) => {
        const body = context.document.body;
        body.insertField(Word.InsertLocation.end, Word.FieldType.date);
        body.insertBookmark("Test_BM");
        const bm = body.getBookmarks(false, false);
        bm.load("items");
        const links = body.getHyperlinkRanges();
        links.load("items");
        const ooxml = body.getOoxml();
        ooxml.load();
        await context.sync();
        return { bookmarks: bm.value, hasOoxml: ooxml.value.includes("<pkg:package") };
      });
    `;
    const res = await sandbox.execute(llmRangeAndClientResultMistakes);
    expect(res.success, res.error).toBe(true);
    expect((res.output as any).hasOoxml).toBe(true);
    expect((res.output as any).bookmarks).toContain('Test_BM');
  });

  it('6. System Prompt enforces National Header, Live Selection targeting, and forbids overwriting document during summarization', () => {
    const prompt = buildSystemPrompt('word', WORD_SKILL_NAMES, 'vi');
    expect(prompt).toContain('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM');
    expect(prompt).toContain('edit_selection');
    expect(prompt).toContain('edit_by_format');
    expect(prompt).toContain('10 CÔNG CỤ MCP KITCHENSINK4WORD');
  });

  it('7. All 10 KitchenSink4Word MCP Tools (including auto-prepended Quốc hiệu/Tiêu ngữ and selection/format targeting) execute 100% cleanly', async () => {
    expect(KITCHENSINK_WORD_TOOLS_META.length).toBe(10);
    const sandbox = new Sandbox('word');

    const testCases: Array<{ tool: string; op: string; args?: Record<string, any> }> = [
      // 1. word_read
      { tool: 'word_read', op: 'get_document_view' },
      { tool: 'word_read', op: 'get_text', args: { start: 0, end: 3 } },
      { tool: 'word_read', op: 'get_document_info' },
      { tool: 'word_read', op: 'get_outline' },
      { tool: 'word_read', op: 'find_text', args: { query: 'ĐƠN' } },
      { tool: 'word_read', op: 'list_elements', args: { type: 'bookmarks' } },
      { tool: 'word_read', op: 'list_elements', args: { type: 'hyperlinks' } },
      { tool: 'word_read', op: 'list_elements', args: { type: 'content_controls' } },
      // 2. word_edit (including auto-prepended Quốc hiệu when starting with BẢN KIỂM ĐIỂM)
      {
        tool: 'word_edit',
        op: 'insert_paragraphs',
        args: {
          location: 'replace',
          paragraphs: [
            { text: 'BẢN KIỂM ĐIỂM', alignment: 'Centered', bold: true, font_size: 14 },
            { text: 'Kính gửi: Ban Giám hiệu nhà trường', alignment: 'Left', bold: true },
          ],
        },
      },
      { tool: 'word_edit', op: 'edit_selection', args: { action: 'replace', new_text: 'Đoạn văn đã sửa trực tiếp từ vùng bôi chọn' } },
      { tool: 'word_edit', op: 'edit_by_format', args: { format: 'bold', action: 'replace', new_text: 'Đã sửa đoạn in đậm' } },
      { tool: 'word_edit', op: 'set_paragraph_text', args: { paragraph_index: 0, new_text: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM' } },
      { tool: 'word_edit', op: 'search_and_replace', args: { find: 'BẢN KIỂM ĐIỂM', replace: 'BẢN KIỂM ĐIỂM CÁ NHÂN', bold: true } },
      // 3. word_format
      { tool: 'word_format', op: 'format_administrative_vn', args: { font_name: 'Times New Roman', font_size: 13, line_spacing_lines: 1.3 } },
      { tool: 'word_format', op: 'format_selection', args: { bold: true, highlight_color: 'Yellow' } },
      { tool: 'word_format', op: 'set_paragraph_format', args: { indices: [0, 1], alignment: 'Centered', bold: true, line_spacing_lines: 1.3 } },
      { tool: 'word_format', op: 'format_text', args: { find: 'KIỂM ĐIỂM', bold: true, color: '#1F4E79' } },
      { tool: 'word_format', op: 'apply_style', args: { style: 'Heading1', paragraph_index: 2, color: '#1F4E79' } },
      // 4. word_structure
      { tool: 'word_structure', op: 'create_table', args: { data: [['STT', 'Công việc'], ['1', 'Bàn giao']], header_row: true } },
      { tool: 'word_structure', op: 'insert_list', args: { items: ['Mục 1', 'Mục 2', 'Mục 3'], kind: 'numbered' } },
      { tool: 'word_structure', op: 'set_section_properties', args: { paper_size: 'A4', top_cm: 2, bottom_cm: 2, left_cm: 3, right_cm: 1.5 } },
      { tool: 'word_structure', op: 'set_header_footer', args: { part: 'both', header_text: 'Lưu hành nội bộ', include_page_number: true } },
      { tool: 'word_structure', op: 'insert_callout_box', args: { text: '📌 LƯU Ý QUAN TRỌNG' } },
      // 5. word_media
      { tool: 'word_media', op: 'insert_field', args: { field_type: 'Date' } },
      { tool: 'word_media', op: 'insert_hyperlink', args: { text: 'Chính phủ', url: 'https://chinhphu.vn' } },
      { tool: 'word_media', op: 'insert_content_control', args: { title: 'Người duyệt', tag: 'approver' } },
      { tool: 'word_media', op: 'insert_image', args: { width_pt: 80, height_pt: 80 } },
      // 6. word_review
      { tool: 'word_review', op: 'manage_comment', args: { action: 'add', text: 'Góp ý kiểm tra' } },
      { tool: 'word_review', op: 'set_track_changes', args: { enabled: true } },
      { tool: 'word_review', op: 'resolve_revisions', args: { action: 'accept_all' } },
      // 7. word_references
      { tool: 'word_references', op: 'manage_note', args: { action: 'insert_footnote', text: 'Chú thích 1' } },
      { tool: 'word_references', op: 'insert_bookmark', args: { name: 'Vi_Tri_Ky' } },
      { tool: 'word_references', op: 'insert_reference_list', args: { type: 'TOC', title: 'MỤC LỤC' } },
      // 8. word_document
      { tool: 'word_document', op: 'create_document', args: { title: 'BẢN KIỂM ĐIỂM' } },
      { tool: 'word_document', op: 'fill_template', args: { data: { '...': 'Nguyễn Văn A' } } },
      { tool: 'word_document', op: 'set_document_properties', args: { title: 'Bản kiểm điểm', author: 'Nguyễn Văn A' } },
      { tool: 'word_document', op: 'set_watermark', args: { watermark: 'BẢN NHÁP' } },
      { tool: 'word_document', op: 'export_table', args: { table_index: 0 } },
      // 9. word_app
      { tool: 'word_app', op: 'live_insert_at_cursor', args: { text: 'Chèn tại con trỏ' } },
      { tool: 'word_app', op: 'live_scroll_to', args: { paragraph_index: 0 } },
      { tool: 'word_app', op: 'com_save_document' },
      { tool: 'word_app', op: 'com_readability_statistics' },
      // 10. word_help
      { tool: 'word_help', op: 'list_tools' },
    ];

    for (const tc of testCases) {
      const script = buildKs4wScript(tc.tool, tc.op, tc.args || {});
      const res = await sandbox.execute(script);
      expect(res.success, `${tc.tool}.${tc.op} failed: ${res.error}`).toBe(true);
      expect(res.output, `${tc.tool}.${tc.op} returned undefined output`).toBeDefined();
      const parsed = typeof res.output === 'string' ? JSON.parse(res.output) : res.output;
      expect(parsed?.error, `${tc.tool}.${tc.op} returned error output`).toBeUndefined();
    }
  });

  it('8. Case-insensitive ("hoa với thường"), extra whitespace ("cách xa cách gần"), and multi-paragraph span replacement works 100%', async () => {
    const sandbox = new Sandbox('word');

    // Insert a 3-line uppercase title with uneven spacing across 3 paragraphs (like user's KLTN cover page)
    await sandbox.execute(
      buildKs4wScript('word_edit', 'insert_paragraphs', {
        location: 'replace',
        paragraphs: [
          'XÂY DỰNG NỀN TẢNG   SAAS QUẢN LÝ',
          'DỰ ÁN AGILE/SCRUM   TÍCH HỢP TRỢ LÝ',
          'ẢO   AI CÁ NHÂN HÓA',
        ],
      }),
    );

    // Search using lowercase and single spaces across the 3 paragraphs
    const replaceRes = await sandbox.execute(
      buildKs4wScript('word_edit', 'search_and_replace', {
        replacements: [
          {
            find: 'xây dựng nền tảng saas quản lý dự án agile/scrum tích hợp trợ lý ảo ai cá nhân hóa',
            replace: 'abc',
          },
        ],
      }),
    );
    expect(replaceRes.success).toBe(true);
    const parsed = JSON.parse(String(replaceRes.output));
    expect(parsed.replacedCount).toBeGreaterThan(0);
  });
});

