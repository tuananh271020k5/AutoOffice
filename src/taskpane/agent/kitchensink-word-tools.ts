import { tool, jsonSchema, type Tool } from 'ai';
import type { Sandbox } from '../executor/sandbox.ts';
import type { AppSettings } from '../store/settings.ts';
import type { OrchestratorCallbacks } from './orchestrator.ts';
import { formatError } from './errors.ts';
import {
  isWordMutation, observeWordDocument, verifyWordMutation, wordReadBackForModel,
  type WordObservation, type WordVerification,
} from './word-verification.ts';

export interface Ks4wToolMeta {
  name: string;
  category: 'read' | 'edit' | 'format' | 'structure' | 'media' | 'review' | 'references' | 'document' | 'app' | 'help';
  titleVi: string;
  titleEn: string;
  descVi: string;
  descEn: string;
  ops: string[];
  samplePromptVi: string;
  samplePromptEn: string;
}

// Cache the latest non-empty selection in Word so that when the user clicks
// into the Task Pane chat input (which can blur the Word canvas), we still
// know the exact text they had selected ("bôi đen / bôi xanh").
let cachedSelectionText = '';
let selectionListenersAtached = false;

export function getCachedWordSelection(): string {
  return cachedSelectionText;
}

export function setCachedWordSelection(text: string): void {
  cachedSelectionText = text.trim();
}

export function initWordSelectionTracker(onSelectionChange?: (text: string) => void): void {
  if (typeof Office === 'undefined' || !Office.context?.document?.addHandlerAsync) return;
  if (selectionListenersAtached) return;
  selectionListenersAtached = true;

  const readCurrentSelection = () => {
    try {
      if (typeof Word === 'undefined' || !Word.run) return;
      void Word.run(async (context) => {
        const sel = context.document.getSelection();
        sel.load('text');
        await context.sync();
        const raw = (sel.text || '').replace(/\r/g, '\n').trim();
        if (raw) {
          cachedSelectionText = raw;
          onSelectionChange?.(raw);
        } else {
          onSelectionChange?.('');
        }
      }).catch(() => {});
    } catch {
      // Ignore outside Word runtime
    }
  };

  try {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => {
        readCurrentSelection();
      },
      () => {
        readCurrentSelection();
      },
    );
  } catch {
    // Ignore if host doesn't support DocumentSelectionChanged
  }
}

export const KITCHENSINK_WORD_TOOLS_META: Ks4wToolMeta[] = [
  {
    name: 'word_read',
    category: 'read',
    titleVi: '1. Đọc nội dung, vùng bôi chọn & thống kê tài liệu',
    titleEn: '1. Read Document, Selection & Statistics',
    descVi: 'Đọc toàn bộ văn bản, nhận diện vùng đang bôi đen/bôi xanh, chữ in đậm/tô màu, bảng biểu, đếm số từ và dàn ý tiêu đề.',
    descEn: 'Read paragraphs, active selection, bold/highlighted text, tables, word count, and outline.',
    ops: ['get_document_view', 'get_text', 'get_document_info', 'word_count', 'get_outline', 'find_text', 'list_elements'],
    samplePromptVi: 'Hãy đọc tài liệu hiện tại và thống kê giúp tôi số đoạn văn, số từ, bảng biểu và nội dung đang bôi chọn.',
    samplePromptEn: 'Inspect the current document and report paragraph count, word count, tables, and selection.',
  },
  {
    name: 'word_edit',
    category: 'edit',
    titleVi: '2. Soạn thảo văn bản, sửa vùng bôi chọn & tìm/thay thế',
    titleEn: '2. Draft Documents, Edit Selection & Search/Replace',
    descVi: 'Soạn văn bản hành chính đầy đủ Quốc hiệu Tiêu ngữ, sửa trực tiếp phần đang bôi đen/bôi đậm/tô màu, tìm & thay thế hoặc xóa đoạn văn.',
    descEn: 'Draft full administrative documents with national header, edit selected/bold/highlighted text, search & replace, or delete paragraphs.',
    ops: ['insert_paragraphs', 'edit_selection', 'edit_by_format', 'set_paragraph_text', 'search_and_replace', 'delete_paragraphs', 'replace_selection_or_document'],
    samplePromptVi: 'Hãy soạn cho tôi một mẫu Đơn xin nghỉ phép chuẩn hành chính Việt Nam có đầy đủ Quốc hiệu, Tiêu ngữ, Kính gửi, thông tin người làm đơn, lý do và phần chữ ký.',
    samplePromptEn: 'Draft a complete administrative leave application form.',
  },
  {
    name: 'word_format',
    category: 'format',
    titleVi: '3. Chuẩn hóa thể thức hành chính & định dạng phông chữ',
    titleEn: '3. Administrative Formatting, Styles & Selection Format',
    descVi: 'Tự động căn chuẩn văn bản hành chính (Quốc hiệu, Tiêu ngữ, Tiêu đề căn giữa; nội dung căn đều 2 bên), hoặc in đậm/đổi màu/tô sáng phần đang bôi chọn.',
    descEn: 'Format paragraph alignment, line spacing, fonts, highlight colors, Heading styles, and Vietnamese administrative layout.',
    ops: ['format_administrative_vn', 'format_selection', 'set_paragraph_format', 'format_text', 'apply_style', 'clear_formatting'],
    samplePromptVi: 'Hãy chuẩn hóa toàn bộ văn bản hiện tại sang phông Times New Roman cỡ 13pt, giãn dòng 1.3 và căn chuẩn thể thức hành chính Việt Nam.',
    samplePromptEn: 'Standardize the document to Times New Roman 13pt, 1.3 line spacing, and proper alignment.',
  },
  {
    name: 'word_structure',
    category: 'structure',
    titleVi: '4. Bảng biểu, Danh sách, Khổ giấy A4 & Tiêu đề đầu/chân trang',
    titleEn: '4. Tables, Lists, Page Setup, Headers/Footers & Callouts',
    descVi: 'Tạo bảng biểu chuyên nghiệp, danh sách đánh số/gạch đầu dòng, căn lề trang A4 chuẩn Nghị định 30, chèn Header/Footer và khung ghi chú.',
    descEn: 'Create styled tables, bulleted/numbered lists, A4 page margins, headers/footers, and callout boxes.',
    ops: ['create_table', 'format_table', 'insert_list', 'set_section_properties', 'set_header_footer', 'insert_callout_box', 'insert_break'],
    samplePromptVi: 'Hãy chèn một bảng danh sách nơi nhận gồm 3 dòng (Ban Giám đốc, Phòng HCNS, Lưu VT) có đóng khung rõ ràng ở cuối tài liệu.',
    samplePromptEn: 'Insert a styled 3-row recipient table at the end of the document.',
  },
  {
    name: 'word_media',
    category: 'media',
    titleVi: '5. Trường thông tin tự động, Liên kết, Biểu mẫu & Hình ảnh',
    titleEn: '5. Fields, Hyperlinks, Content Controls & Images',
    descVi: 'Chèn ngày tháng/số trang tự động, đường dẫn liên kết (Hyperlink), ô điền biểu mẫu và hình ảnh minh họa.',
    descEn: 'Insert dynamic Word fields (Date/Page/TOC), hyperlinks, content controls, and inline images.',
    ops: ['insert_field', 'insert_hyperlink', 'insert_content_control', 'insert_image'],
    samplePromptVi: 'Hãy chèn một ô biểu mẫu có tiêu đề "Ý kiến phê duyệt" ở cuối văn bản.',
    samplePromptEn: 'Insert a form field titled "Approval Note" at the end of the document.',
  },
  {
    name: 'word_review',
    category: 'review',
    titleVi: '6. Nhận xét góp ý & Theo dõi chỉnh sửa tài liệu',
    titleEn: '6. Comments & Track Changes',
    descVi: 'Thêm/xóa/trả lời nhận xét bên lề trang, bật/tắt chế độ theo dõi sửa đổi và chấp nhận hoặc từ chối các thay đổi.',
    descEn: 'Add/reply/delete comments, toggle Track Changes mode, and accept or reject revisions.',
    ops: ['manage_comment', 'set_track_changes', 'list_revisions', 'resolve_revisions'],
    samplePromptVi: 'Hãy chèn một ghi chú nhận xét vào đoạn đầu tiên với nội dung: "Vui lòng kiểm tra lại thông tin trước khi ký duyệt."',
    samplePromptEn: 'Add a review comment on the first paragraph.',
  },
  {
    name: 'word_references',
    category: 'references',
    titleVi: '7. Mục lục tự động, Chú thích chân trang & Đánh dấu trang',
    titleEn: '7. Table of Contents, Footnotes & Bookmarks',
    descVi: 'Chèn chú thích chân trang (Footnote/Endnote), đánh dấu vị trí quan trọng và tạo bảng Mục lục tự động.',
    descEn: 'Insert footnotes/endnotes, manage bookmarks, and generate automatic Tables of Contents (TOC).',
    ops: ['manage_note', 'insert_bookmark', 'insert_reference_list', 'insert_caption'],
    samplePromptVi: 'Hãy chèn một chú thích chân trang vào đoạn đầu tiên với nội dung: "Căn cứ Bộ luật Lao động 2019."',
    samplePromptEn: 'Add a footnote to the first paragraph.',
  },
  {
    name: 'word_document',
    category: 'document',
    titleVi: '8. Điền thông tin biểu mẫu tự động, Đóng dấu mờ & Xuất bảng',
    titleEn: '8. Create Document, Fill Template, Properties & Export Table',
    descVi: 'Tự động điền hàng loạt chỗ trống "..." hoặc {{...}} trong mẫu đơn, gán thông tin tác giả/tiêu đề file, chèn chữ chìm và xuất bảng biểu.',
    descEn: 'Create clean documents, fill {{placeholder}} or [...] templates in place, set document properties, watermark, and export tables.',
    ops: ['create_document', 'fill_template', 'set_document_properties', 'set_watermark', 'export_table'],
    samplePromptVi: 'Hãy điền toàn bộ các chỗ trống "..." trong mẫu đơn hiện tại bằng thông tin mẫu hợp lý (Họ tên: Nguyễn Văn An, Bộ phận: Kỹ thuật, Ngày: 05/11/2026).',
    samplePromptEn: 'Fill all "..." placeholders in the current form with realistic sample data.',
  },
  {
    name: 'word_app',
    category: 'app',
    titleVi: '9. Chèn tại con trỏ, Di chuyển tới đoạn văn, Lưu file & Phân tích câu',
    titleEn: '9. Insert at Cursor, Scroll to Spot, Save Document & Readability Stats',
    descVi: 'Chèn văn bản ngay tại vị trí con trỏ đang đặt, cuộn màn hình Word tới đoạn văn bất kỳ, lưu tài liệu hiện tại và phân tích độ dễ đọc.',
    descEn: 'Insert text at live cursor, scroll Word window to a paragraph/phrase, save the open document, and compute readability statistics.',
    ops: ['live_insert_at_cursor', 'live_scroll_to', 'com_save_document', 'live_set_track_changes', 'com_readability_statistics'],
    samplePromptVi: 'Hãy phân tích thống kê chi tiết độ dài câu, số từ và mức độ dễ đọc của văn bản hiện tại.',
    samplePromptEn: 'Compute detailed readability statistics for the current document.',
  },
  {
    name: 'word_help',
    category: 'help',
    titleVi: '10. Tra cứu danh sách tính năng & hướng dẫn sử dụng nhanh',
    titleEn: '10. Inspect AutoOffice Features & Quick Guide',
    descVi: 'Xem tóm tắt toàn bộ 10 nhóm tính năng hỗ trợ soạn thảo, chỉnh sửa và định dạng văn bản trong AutoOffice.',
    descEn: 'List all 10 AutoOffice tool groups and their main capabilities.',
    ops: ['list_tools', 'inspect_op', 'location_guide'],
    samplePromptVi: 'Hãy giới thiệu tóm tắt các tính năng chính mà AutoOffice có thể hỗ trợ tôi trên tài liệu Word này.',
    samplePromptEn: 'Summarize the main features AutoOffice can help me with on this Word document.',
  },
];

export interface Ks4wCallInput {
  op: string;
  args?: Record<string, unknown>;
}

/**
 * Self-contained function serialized via `.toString()` and executed inside `Word.run` in the sandbox.
 * Keeping this as a real TypeScript function (instead of a template literal string) prevents
 * backslash escaping bugs in regular expressions and gives deterministic behavior.
 */
async function executeKs4wInWord(
  context: any,
  payload: { toolName: string; op: string; args: Record<string, unknown>; cachedSelection?: string },
): Promise<string> {
  const { toolName, op, cachedSelection = '' } = payload;
  const args: Record<string, unknown> = payload.args || {};
  const body = context.document.body;

  const toAlign = (val: unknown, fallback = Word.Alignment.left) => {
    const s = String(val || '').toLowerCase().trim();
    if (s === 'center' || s === 'centered' || s === 'giữa') return Word.Alignment.centered;
    if (s === 'right' || s === 'phải') return Word.Alignment.right;
    if (s === 'justify' || s === 'justified' || s === 'đều') return Word.Alignment.justified;
    if (s === 'left' || s === 'trái') return Word.Alignment.left;
    return fallback;
  };

  const isDocTitleText = (upper: string): boolean => {
    const prefixes = [
      'ĐƠN ', 'BẢN KIỂM ĐIỂM', 'BIÊN BẢN', 'BÁO CÁO', 'TỜ TRÌNH',
      'QUYẾT ĐỊNH', 'THÔNG BÁO', 'HỢP ĐỒNG', 'GIẤY ĐỀ NGHỊ',
      'GIẤY XÁC NHẬN', 'GIẤY ỦY QUYỀN', 'GIẤY MỜI', 'CÔNG VĂN',
      'CHƯƠNG TRÌNH', 'KẾ HOẠCH', 'QUY CHẾ', 'THƯ ', 'BẢN CAM KẾT',
      'PHIẾU ', 'ĐỀ ÁN',
    ];
    return prefixes.some(p => upper.startsWith(p));
  };

  const isSigOrDateText = (upper: string): boolean => {
    const prefixes = [
      'NGÀY ', 'HÀ NỘI', 'TP.', 'THÀNH PHỐ', 'NGƯỜI LÀM ĐƠN',
      'NGƯỜI VIẾT', 'NGƯỜI LẬP', 'TRÂN TRỌNG', '(KÝ', 'KÝ TÊN',
      'XÁC NHẬN CỦA', 'GIÁM ĐỐC', 'THỦ TRƯỞNG', 'HIỆU TRƯỞNG',
    ];
    return prefixes.some(p => upper.startsWith(p)) || upper.includes(', NGÀY ');
  };

  // Normalizes Unicode (NFC vs NFD tổ hợp/dựng sẵn) and builds a case-insensitive,
  // whitespace-tolerant regex so differences in UPPER/lowercase ("hoa với thường")
  // or multiple spaces / soft line breaks ("cách xa cách gần") always match 100%!
  const toNfc = (s: unknown): string => String(s ?? '').normalize('NFC');

  const escapeRegexChar = (s: string): string => {
    const specials = '\\^$.*+?()[]{}|';
    let out = '';
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (specials.indexOf(ch) !== -1) out += '\\' + ch;
      else out += ch;
    }
    return out;
  };

  const buildFlexibleRegex = (query: string, globalFlag = true): RegExp | null => {
    const cleaned = toNfc(query).replace(/[\s\u00a0\u000b\r\n]+/g, ' ').trim();
    if (!cleaned) return null;
    const tokens = cleaned.split(' ').filter(Boolean).map(escapeRegexChar);
    if (tokens.length === 0) return null;
    // Allow any number of spaces, non-breaking spaces, tabs, or line breaks between words
    const pattern = tokens.join('[\\s\\u00a0\\u000b\\r\\n]+');
    return new RegExp(pattern, globalFlag ? 'gi' : 'i');
  };

  const flexibleIncludes = (haystack: string, needle: string): boolean => {
    const rx = buildFlexibleRegex(needle, false);
    if (!rx) return false;
    return rx.test(toNfc(haystack));
  };

  const flexibleReplace = (haystack: string, needle: string, replacement: string): { changed: boolean; text: string } => {
    const rx = buildFlexibleRegex(needle, true);
    if (!rx) return { changed: false, text: haystack };
    const normHay = toNfc(haystack);
    if (!rx.test(normHay)) return { changed: false, text: haystack };
    rx.lastIndex = 0;
    return {
      changed: true,
      text: normHay.replace(rx, toNfc(replacement)),
    };
  };

  // Collects paragraphs from both the active selection's parent container (e.g., floating Text Box
  // on a thesis cover page via sel.parentBody.paragraphs / sel.paragraphs) and the main document body
  // without duplicating paragraphs when the cursor is already in the main body.
  const collectSearchableParagraphs = (
    bodyParas: any,
    parentBodyParas?: any,
    selParas?: any,
  ): any[] => {
    const mainItems: any[] = Array.isArray(bodyParas?.items) ? bodyParas.items : [];
    const parentItems: any[] = Array.isArray(parentBodyParas?.items) ? parentBodyParas.items : [];
    const selItems: any[] = Array.isArray(selParas?.items) ? selParas.items : [];

    const isSameAsMain =
      parentItems.length === mainItems.length &&
      parentItems.every((p, idx) => p === mainItems[idx] || (p?.text || '') === (mainItems[idx]?.text || ''));

    const result: any[] = [];
    if (parentItems.length > 0 && !isSameAsMain) {
      result.push(...parentItems);
    } else if (selItems.length > 0) {
      const mainTexts = new Set(mainItems.map(p => toNfc(p?.text || '').trim()));
      for (const sp of selItems) {
        const t = toNfc(sp?.text || '').trim();
        if (t && !mainTexts.has(t)) {
          result.push(sp);
        }
      }
    }
    result.push(...mainItems);
    return result;
  };

  const applyParagraphFormatting = (
    p: any,
    fmt: Record<string, unknown>,
    defaultFontName = 'Times New Roman',
    defaultFontSize = 13,
  ) => {
    p.font.name = String(fmt.font_name || fmt.fontName || defaultFontName);
    p.font.size = Number(fmt.font_size || fmt.fontSize || defaultFontSize);
    if (fmt.bold !== undefined) p.font.bold = Boolean(fmt.bold);
    if (fmt.italic !== undefined) p.font.italic = Boolean(fmt.italic);
    // Always set underline explicitly so Word never leaks underline from "Độc lập - Tự do - Hạnh phúc"
    p.font.underline = fmt.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
    if (fmt.color) p.font.color = String(fmt.color);
    if (fmt.highlight_color !== undefined || fmt.highlightColor !== undefined) {
      const hc = fmt.highlight_color ?? fmt.highlightColor;
      p.font.highlightColor = hc ? String(hc) : null;
    }
    if (fmt.alignment) p.alignment = toAlign(fmt.alignment);
    const spAfter = fmt.space_after ?? fmt.spaceAfter;
    const spBefore = fmt.space_before ?? fmt.spaceBefore;
    p.spaceAfter = spAfter !== undefined ? Number(spAfter) : 6;
    p.spaceBefore = spBefore !== undefined ? Number(spBefore) : 0;
    const lineSpLines = fmt.line_spacing_lines ?? fmt.lineSpacingLines;
    const lineSpPt = fmt.line_spacing_pt ?? fmt.lineSpacing;
    if (lineSpLines !== undefined) {
      p.lineSpacing = Number(lineSpLines) * 12;
    } else if (lineSpPt !== undefined) {
      const n = Number(lineSpPt);
      p.lineSpacing = n <= 4 ? n * 12 : n;
    } else {
      p.lineSpacing = 15.6;
    }
    if (fmt.first_line_indent_pt !== undefined) {
      p.firstLineIndent = Number(fmt.first_line_indent_pt);
    }
    if (fmt.left_indent_pt !== undefined) {
      p.leftIndent = Number(fmt.left_indent_pt);
    }
  };

  // ==================== 1. WORD_READ ====================
  if (toolName === 'word_read') {
    if (
      op === 'get_document_view' ||
      op === 'get_text' ||
      op === 'get_document_info' ||
      op === 'word_count' ||
      op === 'get_outline'
    ) {
      const sel = context.document.getSelection();
      const selParas = sel.paragraphs;
      const parentBodyParas = sel.parentBody && sel.parentBody.paragraphs ? sel.parentBody.paragraphs : null;
      const paragraphs = body.paragraphs;
      const tables = body.tables;
      const footnotes = body.footnotes;
      const comments = body.getComments();
      const ccs = body.contentControls || context.document.contentControls;
      const pics = body.inlinePictures;

      const paraProps = 'text,style,alignment,font/name,font/size,font/bold,font/italic,font/underline,font/highlightColor,font/color';
      sel.load('text');
      if (selParas && typeof selParas.load === 'function') selParas.load(paraProps);
      if (parentBodyParas && typeof parentBodyParas.load === 'function') parentBodyParas.load(paraProps);
      paragraphs.load(paraProps);
      tables.load('rowCount');
      if (footnotes && typeof footnotes.load === 'function') footnotes.load('items');
      if (comments && typeof comments.load === 'function') comments.load('items');
      if (ccs && typeof ccs.load === 'function') ccs.load('items');
      if (pics && typeof pics.load === 'function') pics.load('items');
      await context.sync();

      const liveSelText = (sel.text || '').replace(/\r/g, '\n').trim();
      const effectiveSelectedText = liveSelText || cachedSelection || '';
      const selectedParaTexts = new Set(
        (selParas?.items || []).map((sp: any) => toNfc(sp.text || '').trim()).filter(Boolean),
      );

      const allParaItems = collectSearchableParagraphs(paragraphs, parentBodyParas, selParas);

      const paraList = allParaItems.map((p: any, idx: number) => {
        const text = toNfc(p.text || '').trim();
        const isSelected = Boolean(
          (liveSelText && text && (selectedParaTexts.has(text) || flexibleIncludes(text, liveSelText) || flexibleIncludes(liveSelText, text))) ||
          (!liveSelText && cachedSelection && text && (flexibleIncludes(text, cachedSelection) || flexibleIncludes(cachedSelection, text)))
        );
        return {
          index: idx,
          text,
          style: p.style,
          alignment: p.alignment,
          fontName: p.font?.name,
          fontSize: p.font?.size,
          bold: Boolean(p.font?.bold),
          italic: Boolean(p.font?.italic),
          underline: Boolean(p.font?.underline && p.font.underline !== 'None' && p.font.underline !== Word.UnderlineType.none),
          highlightColor: p.font?.highlightColor || null,
          color: p.font?.color || null,
          isSelected,
        };
      });

      const fullText = paraList.map((p: any) => p.text).filter(Boolean).join('\n');
      const words = fullText ? fullText.split(/\s+/).filter(Boolean).length : 0;

      if (op === 'word_count' || op === 'get_document_info') {
        return JSON.stringify({
          paragraphCount: paraList.length,
          nonEmptyParagraphCount: paraList.filter((p: any) => p.text.length > 0).length,
          wordCount: words,
          characterCount: fullText.length,
          tableCount: tables.items.length,
          commentCount: comments.items.length,
          footnoteCount: footnotes.items.length,
          contentControlCount: ccs.items.length,
          imageCount: pics.items.length,
          selectedText: effectiveSelectedText,
          boldParagraphs: paraList.filter((p: any) => p.bold && p.text).map((p: any) => ({ index: p.index, text: p.text })),
          highlightedParagraphs: paraList.filter((p: any) => p.highlightColor && p.text).map((p: any) => ({ index: p.index, text: p.text, highlightColor: p.highlightColor })),
        });
      }

      if (op === 'get_outline') {
        const headings = paraList.filter((p: any) =>
          p.text && (/heading|tiêu đề|title/i.test(String(p.style)) || (p.bold && (p.fontSize || 0) >= 14)),
        );
        return JSON.stringify({ headings, totalParagraphs: paraList.length });
      }

      if (op === 'get_text') {
        return JSON.stringify({
          selectedText: effectiveSelectedText,
          wordCount: words,
          paragraphs: paraList,
        });
      }

      const mdLines = paraList.map((p: any) => {
        const tags: string[] = [String(p.alignment || 'Left')];
        if (p.bold) tags.push('Bold/InĐậm');
        if (p.italic) tags.push('Italic');
        if (p.underline) tags.push('Underline/GạchChân');
        if (p.highlightColor) tags.push(`Highlight/BôiMàu:${p.highlightColor}`);
        if (p.color && p.color !== '#000000' && p.color !== 'black') tags.push(`Color:${p.color}`);
        if (p.isSelected) tags.push('🎯ĐANG_BÔI_CHỌN');
        return `[P${p.index}] (${tags.join(', ')}) ${p.text}`;
      });
      return JSON.stringify({
        selectedText: effectiveSelectedText,
        stats: { paragraphs: paraList.length, words, tables: tables.items.length, comments: comments.items.length },
        documentView: mdLines.join('\n'),
      });
    }

    if (op === 'find_text') {
      const q = String(args.query || args.find || '').trim();
      if (!q) return JSON.stringify({ matches: [] });
      const results = body.search(q, { matchCase: Boolean(args.match_case), matchWholeWord: Boolean(args.whole_word) });
      results.load('text');
      await context.sync();
      return JSON.stringify({
        query: q,
        matchCount: results.items.length,
        matches: results.items.map((r: any, i: number) => ({ index: i, text: r.text })),
      });
    }

    if (op === 'list_elements') {
      const elType = String(args.type || 'tables').toLowerCase();
      if (elType === 'bookmarks') {
        const bms = context.document.body.getRange().getBookmarks(true);
        await context.sync();
        return JSON.stringify({ type: 'bookmarks', items: bms.value || [] });
      }
      if (elType === 'content_controls') {
        const ccs = body.contentControls || context.document.contentControls;
        ccs.load('title,tag,text,type');
        await context.sync();
        return JSON.stringify({
          type: 'content_controls',
          items: ccs.items.map((c: any) => ({ title: c.title, tag: c.tag, text: c.text })),
        });
      }
      if (elType === 'hyperlinks') {
        const paras = body.paragraphs;
        paras.load('text,hyperlink');
        await context.sync();
        return JSON.stringify({
          type: 'hyperlinks',
          items: paras.items
            .filter((p: any) => p.hyperlink)
            .map((p: any, i: number) => ({ index: i, text: p.text, url: p.hyperlink })),
        });
      }
      const tables = body.tables;
      tables.load('rowCount,values');
      await context.sync();
      return JSON.stringify({
        type: 'tables',
        count: tables.items.length,
        items: tables.items.map((t: any, i: number) => ({ index: i, rowCount: t.rowCount, values: t.values })),
      });
    }
  }

  // ==================== 2. WORD_EDIT ====================
  if (toolName === 'word_edit') {
    if (op === 'insert_paragraphs') {
      const loc = String(args.location || 'end').toLowerCase();
      if (loc === 'replace') {
        body.clear();
      }
      const rawItems = Array.isArray(args.paragraphs)
        ? (args.paragraphs as Array<Record<string, unknown> | string>)
        : args.text
          ? [String(args.text)]
          : [];

      const normalizedItems: Array<Record<string, unknown>> = rawItems.map(item =>
        typeof item === 'string' ? { text: item } : { ...item },
      );

      // Deterministic guarantee: If creating/replacing a Vietnamese administrative document
      // (BẢN KIỂM ĐIỂM, ĐƠN XIN..., BIÊN BẢN..., BÁO CÁO..., TỜ TRÌNH..., HỢP ĐỒNG..., etc.)
      // and the AI omitted Quốc hiệu & Tiêu ngữ at the top, automatically prepend them!
      if (normalizedItems.length > 0 && (loc === 'replace' || loc === 'start' || loc === 'end')) {
        const firstThreeUpper = normalizedItems
          .slice(0, 3)
          .map(it => String(it.text || '').trim().toUpperCase());
        const hasQuocHieu = firstThreeUpper.some(
          t => t.includes('CỘNG HÒA XÃ HỘI') || t.includes('ĐỘC LẬP - TỰ DO'),
        );
        const startsWithAdminTitle = firstThreeUpper.some(t => isDocTitleText(t));

        if (!hasQuocHieu && startsWithAdminTitle) {
          normalizedItems.unshift(
            {
              text: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',
              alignment: 'Centered',
              bold: true,
              underline: false,
              font_name: 'Times New Roman',
              font_size: 12,
              space_before: 0,
              space_after: 2,
            },
            {
              text: 'Độc lập - Tự do - Hạnh phúc',
              alignment: 'Centered',
              bold: true,
              underline: true,
              font_name: 'Times New Roman',
              font_size: 13,
              space_before: 0,
              space_after: 12,
            },
          );
        }
      }

      const insertLoc = loc === 'start' ? Word.InsertLocation.start : Word.InsertLocation.end;
      const ordered = loc === 'start' ? [...normalizedItems].reverse() : normalizedItems;

      for (const item of ordered) {
        const text = String(item.text ?? '');
        const upper = text.trim().toUpperCase();
        const styleObj = { ...item };

        // Auto-apply standard Vietnamese administrative formatting on inserted lines
        if (upper.includes('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM')) {
          styleObj.alignment = 'Centered';
          styleObj.bold = true;
          styleObj.underline = false;
        } else if (upper.includes('ĐỘC LẬP - TỰ DO - HẠNH PHÚC')) {
          styleObj.alignment = 'Centered';
          styleObj.bold = true;
          styleObj.underline = true;
          if (styleObj.space_after === undefined) styleObj.space_after = 12;
        } else if (isDocTitleText(upper) && text.trim().length < 90) {
          styleObj.alignment = 'Centered';
          styleObj.bold = true;
          styleObj.underline = false;
          if (!styleObj.font_size) styleObj.font_size = 14;
          if (styleObj.space_before === undefined) styleObj.space_before = 6;
          if (styleObj.space_after === undefined) styleObj.space_after = 10;
        }

        const p = body.insertParagraph(text, insertLoc);
        if (styleObj.style) {
          try {
            p.styleBuiltIn = styleObj.style;
          } catch {
            // Ignore custom style error
          }
        }
        applyParagraphFormatting(p, styleObj);
      }
      await context.sync();
      return JSON.stringify({
        status: 'ok',
        insertedParagraphs: ordered.length,
        hasNationalHeader: true,
      });
    }

    if (op === 'edit_selection' || op === 'replace_selection_or_document') {
      const newText = String(args.new_text ?? args.text ?? args.replace ?? '');
      const action = String(args.action || 'replace').toLowerCase(); // 'replace' | 'delete' | 'append'
      const explicitFind = toNfc(args.find || '').trim();
      const sel = context.document.getSelection();
      const selParas = sel.paragraphs;
      const parentBodyParas = sel.parentBody && sel.parentBody.paragraphs ? sel.parentBody.paragraphs : null;
      const paras = body.paragraphs;

      sel.load('text,font/name,font/size');
      if (selParas && typeof selParas.load === 'function') selParas.load('text');
      if (parentBodyParas && typeof parentBodyParas.load === 'function') parentBodyParas.load('text');
      paras.load('text');
      await context.sync();

      const liveSel = toNfc(sel.text || '').replace(/\r/g, '\n').trim();
      const targetText = explicitFind || liveSel || toNfc(cachedSelection || '').trim();
      const allParas = collectSearchableParagraphs(paras, parentBodyParas, selParas);

      // Use liveSel directly unless explicitFind is provided and is wider than liveSel (e.g. user selected 3 words of a 16-word multi-line title)
      const shouldUseLiveSelDirectly =
        Boolean(liveSel) && (!explicitFind || flexibleIncludes(liveSel, explicitFind));

      if (shouldUseLiveSelDirectly) {
        if (action === 'delete') {
          sel.delete();
        } else {
          const r = sel.insertText(newText, Word.InsertLocation.replace);
          if (args.bold !== undefined) r.font.bold = Boolean(args.bold);
          if (args.italic !== undefined) r.font.italic = Boolean(args.italic);
          if (args.underline !== undefined) {
            r.font.underline = args.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
          }
          if (args.color) r.font.color = String(args.color);
        }
        await context.sync();
        return JSON.stringify({
          status: 'ok',
          target: 'live_selection',
          originalText: liveSel,
          newText: action === 'delete' ? '' : newText,
        });
      }

      if (targetText) {
        let updatedCount = 0;
        for (const p of allParas) {
          const pText = toNfc(p.text || '').trim();
          if (!pText) continue;
          const repRes = flexibleReplace(pText, targetText, action === 'delete' ? '' : newText);
          if (repRes.changed) {
            if (action === 'delete' && !repRes.text.trim()) {
              p.delete();
            } else {
              p.insertText(repRes.text, Word.InsertLocation.replace);
              if (args.bold !== undefined) p.font.bold = Boolean(args.bold);
              if (args.underline !== undefined) {
                p.font.underline = args.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
              }
            }
            updatedCount++;
            break;
          }
        }
        // Multi-paragraph span fallback if selection/title crossed 2..5 lines
        if (updatedCount === 0 && allParas.length > 1) {
          for (let win = 2; win <= Math.min(5, allParas.length) && updatedCount === 0; win++) {
            for (let i = 0; i <= allParas.length - win; i++) {
              const joined = allParas
                .slice(i, i + win)
                .map((p: any) => toNfc(p.text || '').trim())
                .join(' ');
              const repRes = flexibleReplace(joined, targetText, action === 'delete' ? '' : newText);
              if (repRes.changed) {
                allParas[i].insertText(repRes.text, Word.InsertLocation.replace);
                for (let k = i + win - 1; k >= i + 1; k--) {
                  allParas[k].delete();
                }
                updatedCount++;
                break;
              }
            }
          }
        }
        // If paragraph search still found 0 matches and liveSel is non-empty, replace liveSel!
        if (updatedCount === 0 && liveSel) {
          if (action === 'delete') {
            sel.delete();
          } else {
            sel.insertText(newText, Word.InsertLocation.replace);
          }
          updatedCount++;
        }
        await context.sync();
        return JSON.stringify({
          status: updatedCount > 0 ? 'ok' : 'no_match',
          target: 'cached_selection',
          originalText: targetText,
          newText: action === 'delete' ? '' : newText,
          updatedCount,
        });
      }

      return JSON.stringify({
        status: 'no_selection',
        message: 'Không tìm thấy văn bản đang bôi chọn. Hãy truyền paragraph_index hoặc find để chỉnh sửa đoạn văn cụ thể.',
      });
    }

    if (op === 'edit_by_format') {
      // Edit or delete paragraphs/runs that are currently Bold ("bôi đậm"),
      // Highlighted ("bôi xanh / bôi vàng / tô sáng"), or Underlined ("gạch chân")!
      const matchFormat = String(args.format || args.target || 'bold').toLowerCase(); // 'bold' | 'highlight' | 'underline'
      const action = String(args.action || 'replace').toLowerCase(); // 'replace' | 'delete' | 'clear_format'
      const newText = String(args.new_text ?? args.text ?? args.replace ?? '');

      const sel = context.document.getSelection();
      const selParas = sel.paragraphs;
      const parentBodyParas = sel.parentBody && sel.parentBody.paragraphs ? sel.parentBody.paragraphs : null;
      const paras = body.paragraphs;
      const fmtProps = 'text,font/bold,font/underline,font/highlightColor,font/color';
      if (selParas && typeof selParas.load === 'function') selParas.load(fmtProps);
      if (parentBodyParas && typeof parentBodyParas.load === 'function') parentBodyParas.load(fmtProps);
      paras.load(fmtProps);
      await context.sync();

      const allParas = collectSearchableParagraphs(paras, parentBodyParas, selParas);
      const matchedIndices: number[] = [];
      for (let i = allParas.length - 1; i >= 0; i--) {
        const p = allParas[i];
        const t = (p.text || '').trim();
        if (!t) continue;
        const isBold = Boolean(p.font?.bold);
        const isUnderline = Boolean(p.font?.underline && p.font.underline !== 'None' && p.font.underline !== Word.UnderlineType.none);
        const isHighlight = Boolean(p.font?.highlightColor);

        const matched =
          (matchFormat.includes('bold') || matchFormat.includes('đậm')) ? isBold :
          (matchFormat.includes('highlight') || matchFormat.includes('xanh') || matchFormat.includes('vàng') || matchFormat.includes('màu')) ? isHighlight :
          (matchFormat.includes('underline') || matchFormat.includes('gạch')) ? isUnderline :
          (isBold || isHighlight);

        if (matched) {
          matchedIndices.unshift(i);
          if (action === 'delete') {
            p.delete();
          } else if (action === 'clear_format') {
            p.font.bold = false;
            p.font.underline = Word.UnderlineType.none;
            p.font.highlightColor = null;
          } else if (newText) {
            p.insertText(newText, Word.InsertLocation.replace);
          }
          if (args.bold !== undefined) p.font.bold = Boolean(args.bold);
          if (args.underline !== undefined) {
            p.font.underline = args.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
          }
          if (args.highlight_color !== undefined) {
            p.font.highlightColor = args.highlight_color ? String(args.highlight_color) : null;
          }
        }
      }
      await context.sync();
      return JSON.stringify({
        status: 'ok',
        matchFormat,
        action,
        matchedCount: matchedIndices.length,
        matchedParagraphIndices: matchedIndices,
      });
    }

    if (op === 'set_paragraph_text') {
      const idx = Number(args.paragraph_index ?? args.index ?? 0);
      const newText = String(args.new_text ?? args.text ?? args.replace ?? '');
      const sel = context.document.getSelection();
      const selParas = sel.paragraphs;
      const parentBodyParas = sel.parentBody && sel.parentBody.paragraphs ? sel.parentBody.paragraphs : null;
      const paras = body.paragraphs;
      if (selParas && typeof selParas.load === 'function') selParas.load('text');
      if (parentBodyParas && typeof parentBodyParas.load === 'function') parentBodyParas.load('text');
      paras.load('text');
      await context.sync();
      const allParas = collectSearchableParagraphs(paras, parentBodyParas, selParas);
      if (idx < 0 || idx >= allParas.length) {
        return JSON.stringify({ error: `Paragraph index ${idx} out of range (0..${allParas.length - 1})` });
      }
      const target = allParas[idx];
      target.insertText(newText, Word.InsertLocation.replace);
      if (args.bold !== undefined || args.alignment !== undefined || args.font_size !== undefined || args.underline !== undefined) {
        applyParagraphFormatting(target, args);
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', updatedParagraph: idx, newText });
    }

    if (op === 'search_and_replace') {
      // Support BOTH args.replacements array AND direct top-level args.find / args.replace!
      const rawReplacements = Array.isArray(args.replacements)
        ? (args.replacements as Array<Record<string, unknown>>)
        : (args.find !== undefined || args.query !== undefined)
          ? [args]
          : [];

      let totalReplaced = 0;
      const sel = context.document.getSelection();
      const selParas = sel.paragraphs;
      const parentBodyParas = sel.parentBody && sel.parentBody.paragraphs ? sel.parentBody.paragraphs : null;
      const paras = body.paragraphs;

      sel.load('text');
      if (selParas && typeof selParas.load === 'function') selParas.load('text');
      if (parentBodyParas && typeof parentBodyParas.load === 'function') parentBodyParas.load('text');
      paras.load('text');
      await context.sync();

      const liveSel = toNfc(sel.text || '').replace(/\r/g, '\n').trim();
      const allParas = collectSearchableParagraphs(paras, parentBodyParas, selParas);

      for (const rep of rawReplacements) {
        const findText = toNfc(rep.find ?? rep.query ?? rep.target ?? '').trim();
        if (!findText) continue;
        const replaceText = rep.replace !== undefined
          ? String(rep.replace)
          : rep.new_text !== undefined
            ? String(rep.new_text)
            : undefined;

        let matchedCurrent = 0;

        // 1. If the user currently has text selected on screen AND `liveSel` already contains the entire `findText`,
        // apply directly to the live selection first!
        // (If `findText` is LONGER than `liveSel` — e.g., user only highlighted 3 words of a 16-word multi-line title —
        // we search `allParas` first so the full 16-word title is replaced rather than only the 3 words!)
        if (liveSel && flexibleIncludes(liveSel, findText)) {
          if (replaceText !== undefined) {
            const repSel = flexibleReplace(liveSel, findText, replaceText);
            const finalSelText = repSel.changed ? repSel.text : replaceText;
            const r = sel.insertText(finalSelText, Word.InsertLocation.replace);
            if (rep.bold !== undefined) r.font.bold = Boolean(rep.bold);
            if (rep.italic !== undefined) r.font.italic = Boolean(rep.italic);
            if (rep.underline !== undefined) {
              r.font.underline = rep.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
            }
            if (rep.color) r.font.color = String(rep.color);
            if (rep.highlight_color !== undefined) {
              r.font.highlightColor = rep.highlight_color ? String(rep.highlight_color) : null;
            }
          }
          matchedCurrent++;
          totalReplaced++;
          continue;
        }

        // 2. Single-paragraph case-insensitive + whitespace-tolerant + NFC-normalized search
        // across both floating Text Box / parentBody paragraphs and main document paragraphs!
        for (const p of allParas) {
          const pText = toNfc(p.text || '').trim();
          if (!pText) continue;
          if (flexibleIncludes(pText, findText)) {
            if (replaceText !== undefined) {
              const repRes = flexibleReplace(pText, findText, replaceText);
              if (repRes.changed) {
                p.insertText(repRes.text, Word.InsertLocation.replace);
              }
            }
            if (rep.bold !== undefined) p.font.bold = Boolean(rep.bold);
            if (rep.italic !== undefined) p.font.italic = Boolean(rep.italic);
            if (rep.underline !== undefined) {
              p.font.underline = rep.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
            }
            if (rep.color) p.font.color = String(rep.color);
            if (rep.highlight_color !== undefined) {
              p.font.highlightColor = rep.highlight_color ? String(rep.highlight_color) : null;
            }
            matchedCurrent++;
            totalReplaced++;
          }
        }

        // 3. Multi-paragraph sliding window search (when a title or sentence is split across 2..5 lines!)
        if (matchedCurrent === 0 && allParas.length > 1) {
          for (let win = 2; win <= Math.min(5, allParas.length) && matchedCurrent === 0; win++) {
            for (let i = 0; i <= allParas.length - win; i++) {
              const joined = allParas
                .slice(i, i + win)
                .map((p: any) => toNfc(p.text || '').trim())
                .join(' ');
              if (flexibleIncludes(joined, findText)) {
                if (replaceText !== undefined) {
                  const repRes = flexibleReplace(joined, findText, replaceText);
                  allParas[i].insertText(repRes.changed ? repRes.text : replaceText, Word.InsertLocation.replace);
                  for (let k = i + win - 1; k >= i + 1; k--) {
                    allParas[k].delete();
                  }
                }
                if (rep.bold !== undefined) allParas[i].font.bold = Boolean(rep.bold);
                if (rep.color) allParas[i].font.color = String(rep.color);
                matchedCurrent++;
                totalReplaced++;
                break;
              }
            }
          }
        }

        // 4. Final fallback: if still 0 matches and the user had an active/cached selection, replace that selection!
        if (matchedCurrent === 0 && replaceText !== undefined) {
          if (liveSel) {
            sel.insertText(replaceText, Word.InsertLocation.replace);
            matchedCurrent++;
            totalReplaced++;
          } else if (cachedSelection) {
            for (const p of allParas) {
              const pText = toNfc(p.text || '').trim();
              if (pText && flexibleIncludes(pText, cachedSelection)) {
                const repRes = flexibleReplace(pText, cachedSelection, replaceText);
                if (repRes.changed) {
                  p.insertText(repRes.text, Word.InsertLocation.replace);
                  matchedCurrent++;
                  totalReplaced++;
                  break;
                }
              }
            }
          }
        }
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', replacedCount: totalReplaced });
    }

    if (op === 'delete_paragraphs') {
      const sel = context.document.getSelection();
      const selParas = sel.paragraphs;
      const parentBodyParas = sel.parentBody && sel.parentBody.paragraphs ? sel.parentBody.paragraphs : null;
      const paras = body.paragraphs;
      const delProps = 'text,font/bold,font/highlightColor,font/underline';
      if (selParas && typeof selParas.load === 'function') selParas.load(delProps);
      if (parentBodyParas && typeof parentBodyParas.load === 'function') parentBodyParas.load(delProps);
      paras.load(delProps);
      await context.sync();

      const allParas = collectSearchableParagraphs(paras, parentBodyParas, selParas);

      // Support deleting by indices array, single index, find text, or start..end range
      const indicesArg = Array.isArray(args.indices)
        ? (args.indices as number[]).map(n => Number(n))
        : (args.paragraph_index !== undefined || args.index !== undefined)
          ? [Number(args.paragraph_index ?? args.index)]
          : null;

      let deleted = 0;
      if (indicesArg && indicesArg.length > 0) {
        const sortedDesc = [...new Set(indicesArg)].sort((a, b) => b - a);
        for (const i of sortedDesc) {
          if (i >= 0 && i < allParas.length) {
            allParas[i].delete();
            deleted++;
          }
        }
      } else if (args.find) {
        const q = String(args.find).trim();
        for (let i = allParas.length - 1; i >= 0; i--) {
          const t = (allParas[i].text || '').trim();
          if (t && flexibleIncludes(t, q)) {
            allParas[i].delete();
            deleted++;
          }
        }
      } else {
        const start = Number(args.start ?? 0);
        const end = Number(args.end ?? start);
        for (let i = Math.min(end, allParas.length - 1); i >= Math.max(0, start); i--) {
          allParas[i].delete();
          deleted++;
        }
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', deletedCount: deleted });
    }
  }

  // ==================== 3. WORD_FORMAT ====================
  if (toolName === 'word_format') {
    if (op === 'format_administrative_vn') {
      const fontName = String(args.font_name || 'Times New Roman');
      const fontSize = Number(args.font_size || 13);
      const lineSpacingPt = Number(args.line_spacing_lines || 1.3) * 12;

      const paras = body.paragraphs;
      paras.load('text');
      await context.sync();

      // Ensure Quốc hiệu & Tiêu ngữ exist at the top if this is an administrative document
      const nonEmptyTexts = paras.items.map((p: any) => (p.text || '').trim().toUpperCase()).filter(Boolean);
      const hasQuocHieu = nonEmptyTexts.slice(0, 3).some(
        (t: string) => t.includes('CỘNG HÒA XÃ HỘI') || t.includes('ĐỘC LẬP - TỰ DO'),
      );
      if (!hasQuocHieu && nonEmptyTexts.length > 0 && isDocTitleText(nonEmptyTexts[0])) {
        const pSub = body.insertParagraph('Độc lập - Tự do - Hạnh phúc', Word.InsertLocation.start);
        pSub.font.name = fontName;
        pSub.font.size = fontSize;
        pSub.font.bold = true;
        pSub.font.underline = Word.UnderlineType.single;
        pSub.alignment = Word.Alignment.centered;
        pSub.spaceAfter = 12;

        const pHead = body.insertParagraph('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', Word.InsertLocation.start);
        pHead.font.name = fontName;
        pHead.font.size = fontSize - 1;
        pHead.font.bold = true;
        pHead.font.underline = Word.UnderlineType.none;
        pHead.alignment = Word.Alignment.centered;
        pHead.spaceAfter = 2;
        await context.sync();
        paras.load('text');
        await context.sync();
      }

      for (let i = 0; i < paras.items.length; i++) {
        const p = paras.items[i];
        const raw = (p.text || '').trim();
        const upper = raw.toUpperCase();
        p.font.name = fontName;
        p.font.size = fontSize;
        p.lineSpacing = lineSpacingPt;
        p.spaceBefore = 0;
        p.spaceAfter = 6;
        p.font.underline = Word.UnderlineType.none;

        if (!raw) continue;

        const isQuocHieu = upper.includes('CỘNG HÒA XÃ HỘI') || upper.includes('ĐỘC LẬP - TỰ DO');
        const isTitle = isDocTitleText(upper) && raw.length < 90;
        const isSigOrDate = isSigOrDateText(upper) && raw.length < 70;

        if (isQuocHieu || isTitle) {
          p.alignment = Word.Alignment.centered;
          p.font.bold = true;
          if (upper.includes('ĐỘC LẬP - TỰ DO')) {
            p.font.underline = Word.UnderlineType.single;
            p.spaceAfter = 12;
          }
          if (isTitle) {
            p.font.size = fontSize + 1;
            p.spaceBefore = 6;
            p.spaceAfter = 12;
          }
        } else if (isSigOrDate) {
          p.alignment = Word.Alignment.right;
          if (upper.startsWith('NGƯỜI ') || upper.includes('GIÁM ĐỐC') || upper.includes('HIỆU TRƯỞNG')) {
            p.font.bold = true;
          }
        } else {
          p.alignment = Word.Alignment.justified;
        }
      }
      await context.sync();
      return JSON.stringify({
        status: 'ok',
        formattedParagraphs: paras.items.length,
        fontName,
        fontSize,
      });
    }

    if (op === 'format_selection' || op === 'set_paragraph_format') {
      // Merge top-level args and args.formatting so either parameter shape works!
      const fmt: Record<string, unknown> = {
        ...args,
        ...((args.formatting as Record<string, unknown>) || {}),
      };
      const indices = Array.isArray(args.indices)
        ? (args.indices as number[]).map(n => Number(n))
        : args.paragraph_index !== undefined || args.index !== undefined
          ? [Number(args.paragraph_index ?? args.index)]
          : [];

      const paras = body.paragraphs;
      paras.load('items,text');
      await context.sync();

      if (op === 'format_selection' || indices.length === 0) {
        const sel = context.document.getSelection();
        sel.load('text');
        await context.sync();
        const liveSel = (sel.text || '').trim();
        if (liveSel) {
          if (fmt.font_name) sel.font.name = String(fmt.font_name);
          if (fmt.font_size) sel.font.size = Number(fmt.font_size);
          if (fmt.bold !== undefined) sel.font.bold = Boolean(fmt.bold);
          if (fmt.italic !== undefined) sel.font.italic = Boolean(fmt.italic);
          if (fmt.underline !== undefined) {
            sel.font.underline = fmt.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
          }
          if (fmt.color) sel.font.color = String(fmt.color);
          if (fmt.highlight_color !== undefined) {
            sel.font.highlightColor = fmt.highlight_color ? String(fmt.highlight_color) : null;
          }
          await context.sync();
          return JSON.stringify({ status: 'ok', formattedTarget: 'live_selection', selectedText: liveSel });
        }
      }

      const targets = indices.length > 0
        ? indices.filter(i => i >= 0 && i < paras.items.length).map(i => paras.items[i])
        : paras.items;

      for (const p of targets) {
        applyParagraphFormatting(p, fmt);
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', updatedCount: targets.length });
    }

    if (op === 'format_text') {
      const findText = String(args.find || args.query || cachedSelection || '').trim();
      const fmt: Record<string, unknown> = {
        ...args,
        ...((args.formatting as Record<string, unknown>) || {}),
      };
      if (!findText) {
        const sel = context.document.getSelection();
        if (fmt.bold !== undefined) sel.font.bold = Boolean(fmt.bold);
        if (fmt.italic !== undefined) sel.font.italic = Boolean(fmt.italic);
        if (fmt.underline !== undefined) {
          sel.font.underline = fmt.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
        }
        if (fmt.color) sel.font.color = String(fmt.color);
        if (fmt.highlight_color !== undefined) {
          sel.font.highlightColor = fmt.highlight_color ? String(fmt.highlight_color) : null;
        }
        await context.sync();
        return JSON.stringify({ status: 'ok', formattedTarget: 'selection' });
      }
      const matches = body.search(findText, { matchCase: false });
      matches.load('items');
      await context.sync();
      let formattedCount = matches.items.length;
      for (const m of matches.items) {
        if (fmt.font_name) m.font.name = String(fmt.font_name);
        if (fmt.font_size) m.font.size = Number(fmt.font_size);
        if (fmt.bold !== undefined) m.font.bold = Boolean(fmt.bold);
        if (fmt.italic !== undefined) m.font.italic = Boolean(fmt.italic);
        if (fmt.underline !== undefined) {
          m.font.underline = fmt.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
        }
        if (fmt.color) m.font.color = String(fmt.color);
        if (fmt.highlight_color !== undefined) {
          m.font.highlightColor = fmt.highlight_color ? String(fmt.highlight_color) : null;
        }
      }
      if (formattedCount === 0) {
        const paras = body.paragraphs;
        paras.load('text');
        await context.sync();
        for (const p of paras.items) {
          if (flexibleIncludes(p.text || '', findText)) {
            applyParagraphFormatting(p, fmt);
            formattedCount++;
          }
        }
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', matchedCount: formattedCount });
    }

    if (op === 'apply_style') {
      const styleName = String(args.style || 'Heading1');
      const findText = String(args.find || '');
      const paras = body.paragraphs;
      paras.load('text');
      await context.sync();
      let count = 0;
      for (let i = 0; i < paras.items.length; i++) {
        const p = paras.items[i];
        if ((findText && p.text.includes(findText)) || (!findText && i === Number(args.paragraph_index ?? 0))) {
          try {
            p.styleBuiltIn = styleName;
          } catch {
            p.style = styleName;
          }
          if (args.color) p.font.color = String(args.color);
          if (args.font_size) p.font.size = Number(args.font_size);
          if (args.bold !== undefined) p.font.bold = Boolean(args.bold);
          count++;
        }
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', appliedStyle: styleName, count });
    }

    if (op === 'clear_formatting') {
      const paras = body.paragraphs;
      paras.load('items');
      await context.sync();
      for (const p of paras.items) {
        p.font.bold = false;
        p.font.italic = false;
        p.font.underline = Word.UnderlineType.none;
        p.font.highlightColor = null;
        p.font.color = '#000000';
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', clearedParagraphs: paras.items.length });
    }
  }

  // ==================== 4. WORD_STRUCTURE ====================
  if (toolName === 'word_structure') {
    if (op === 'create_table') {
      const data = (args.data as string[][]) || [
        ['STT', 'Hạng mục', 'Ghi chú'],
        ['1', 'Nội dung 1', 'Hoàn thành'],
      ];
      const rows = data.length;
      const cols = data[0]?.length || 1;
      const loc = String(args.location || 'end').toLowerCase() === 'start'
        ? Word.InsertLocation.start
        : Word.InsertLocation.end;
      const table = body.insertTable(rows, cols, loc, data);
      table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light || Word.BuiltInStyleName.tableGrid;
      if (table.font) {
        table.font.name = String(args.font_name || 'Times New Roman');
        table.font.size = Number(args.font_size || 12);
      }
      if (args.header_row !== false && rows > 0) {
        const headerRow = table.rows ? table.rows.getFirst() : table.getRow(0);
        headerRow.font.bold = true;
        headerRow.shadingColor = String(args.header_bg || '#D9E1F2');
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', rows, cols });
    }

    if (op === 'insert_list') {
      const items = (args.items as string[]) || [];
      const kind = String(args.kind || 'bulleted').toLowerCase();
      for (let i = 0; i < items.length; i++) {
        const prefix = kind === 'numbered' ? `${i + 1}. ` : '• ';
        const p = body.insertParagraph(`${prefix}${items[i]}`, Word.InsertLocation.end);
        p.font.name = 'Times New Roman';
        p.font.size = 13;
        p.font.underline = Word.UnderlineType.none;
        p.leftIndent = 18;
        p.spaceAfter = 4;
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', listItemsInserted: items.length });
    }

    if (op === 'set_section_properties') {
      const section = context.document.sections.getFirst();
      const ps = section.pageSetup;
      const paperEnum = (Word as any).PaperSize || (Word as any).PaperType;
      if (paperEnum && paperEnum.a4) ps.paperSize = paperEnum.a4;
      ps.orientation = String(args.orientation || 'portrait').toLowerCase() === 'landscape'
        ? Word.PageOrientation.landscape
        : Word.PageOrientation.portrait;
      const cmToPt = 28.3465;
      ps.topMargin = Number(args.top_cm ?? 2) * cmToPt;
      ps.bottomMargin = Number(args.bottom_cm ?? 2) * cmToPt;
      ps.leftMargin = Number(args.left_cm ?? 3) * cmToPt;
      ps.rightMargin = Number(args.right_cm ?? 1.5) * cmToPt;
      await context.sync();
      return JSON.stringify({ status: 'ok', paperSize: 'A4', orientation: ps.orientation });
    }

    if (op === 'set_header_footer') {
      const section = context.document.sections.getFirst();
      const part = String(args.part || 'header').toLowerCase();
      if (part === 'header' || part === 'both') {
        const header = section.getHeader(Word.HeaderFooterType.primary);
        header.clear();
        const hp = header.insertParagraph(String(args.header_text || args.text || ''), Word.InsertLocation.start);
        hp.alignment = toAlign(args.alignment || 'Right');
        hp.font.name = 'Times New Roman';
        hp.font.size = 10;
        hp.font.italic = true;
      }
      if (part === 'footer' || part === 'both') {
        const footer = section.getFooter(Word.HeaderFooterType.primary);
        footer.clear();
        const fp = footer.insertParagraph(String(args.footer_text || 'Trang '), Word.InsertLocation.start);
        fp.alignment = Word.Alignment.centered;
        fp.font.name = 'Times New Roman';
        fp.font.size = 10;
        if (args.include_page_number !== false) {
          fp.getRange(Word.RangeLocation.end).insertField(Word.InsertLocation.end, Word.FieldType.page, '', true);
        }
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', part });
    }

    if (op === 'insert_callout_box') {
      const text = String(args.text || '');
      const tbl = body.insertTable(1, 1, Word.InsertLocation.end, [[text]]);
      const cell = tbl.getCell(0, 0);
      cell.shadingColor = String(args.bg_color || '#E8F4FD');
      if (cell.body.font) {
        cell.body.font.name = 'Times New Roman';
        cell.body.font.size = 12;
        cell.body.font.bold = Boolean(args.bold ?? true);
      }
      const border = tbl.getBorder(Word.BorderLocation.outside);
      border.type = Word.BorderType.single;
      border.color = String(args.border_color || '#0078D4');
      border.width = 1.5;
      await context.sync();
      return JSON.stringify({ status: 'ok', calloutInserted: true });
    }

    if (op === 'insert_break') {
      const breakType = String(args.break_type || 'Page').toLowerCase() === 'line'
        ? Word.BreakType.line
        : Word.BreakType.page;
      body.insertBreak(breakType, Word.InsertLocation.end);
      await context.sync();
      return JSON.stringify({ status: 'ok', breakType });
    }
  }

  // ==================== 5. WORD_MEDIA ====================
  if (toolName === 'word_media') {
    if (op === 'insert_field') {
      const fTypeRaw = String(args.field_type || 'Date').toLowerCase();
      const p = body.insertParagraph(String(args.prefix_text || ''), Word.InsertLocation.end);
      p.font.name = 'Times New Roman';
      p.font.size = 13;
      p.font.underline = Word.UnderlineType.none;
      const fType = fTypeRaw === 'page'
        ? Word.FieldType.page
        : fTypeRaw === 'toc'
          ? Word.FieldType.toc
          : Word.FieldType.date;
      p.getRange(Word.RangeLocation.end).insertField(Word.InsertLocation.end, fType, String(args.switches || ''), true);
      await context.sync();
      return JSON.stringify({ status: 'ok', fieldType: fType });
    }

    if (op === 'insert_hyperlink') {
      const text = String(args.text || args.url || 'Liên kết');
      const url = String(args.url || 'https://chinhphu.vn');
      const p = body.insertParagraph(text, Word.InsertLocation.end);
      p.font.name = 'Times New Roman';
      p.font.size = 13;
      p.hyperlink = url;
      await context.sync();
      return JSON.stringify({ status: 'ok', text, url });
    }

    if (op === 'insert_content_control') {
      const title = String(args.title || 'Mục thông tin');
      const tag = String(args.tag || 'cc_field');
      const text = String(args.text || `[Nhập ${title}]`);
      const p = body.insertParagraph(text, Word.InsertLocation.end);
      p.font.name = 'Times New Roman';
      p.font.size = 13;
      p.font.underline = Word.UnderlineType.none;
      const cc = p.insertContentControl();
      cc.title = title;
      cc.tag = tag;
      cc.appearance = Word.ContentControlAppearance.boundingBox;
      await context.sync();
      return JSON.stringify({ status: 'ok', title, tag });
    }

    if (op === 'insert_image') {
      const defaultPng = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAALUlEQVR42u3OMQEAAAgDoJvc6FEOhB40gK5d27Zt27Zt27Zt27Zt27Zt2/Y8A7h8A4G89pL1AAAAAElFTkSuQmCC';
      const base64 = String(args.base64 || defaultPng);
      const img = body.insertInlinePictureFromBase64(base64, Word.InsertLocation.end);
      img.width = Number(args.width_pt || 80);
      img.height = Number(args.height_pt || 80);
      img.altTextTitle = String(args.alt_title || 'Hình ảnh');
      await context.sync();
      return JSON.stringify({ status: 'ok', width: img.width, height: img.height });
    }
  }

  // ==================== 6. WORD_REVIEW ====================
  if (toolName === 'word_review') {
    if (op === 'manage_comment') {
      const action = String(args.action || 'add').toLowerCase();
      if (action === 'list') {
        const comments = body.getComments();
        comments.load('content,authorName,resolved');
        await context.sync();
        return JSON.stringify({
          count: comments.items.length,
          comments: comments.items.map((c: any) => ({
            content: c.content,
            author: c.authorName,
            resolved: c.resolved,
          })),
        });
      }
      const text = String(args.text || 'Cần rà soát lại nội dung này.');
      const firstPara = body.paragraphs.getFirst();
      firstPara.getRange(Word.RangeLocation.content).insertComment(text);
      await context.sync();
      return JSON.stringify({ status: 'ok', commentAdded: text });
    }

    if (op === 'set_track_changes') {
      const enabled = args.enabled !== false;
      context.document.changeTrackingMode = enabled
        ? Word.ChangeTrackingMode.trackAll
        : Word.ChangeTrackingMode.off;
      await context.sync();
      return JSON.stringify({ status: 'ok', trackChangesEnabled: enabled });
    }

    if (op === 'resolve_revisions') {
      const action = String(args.action || 'accept_all').toLowerCase();
      const tracked = body.getTrackedChanges();
      tracked.load('items');
      await context.sync();
      if (action === 'reject_all') {
        tracked.rejectAll();
      } else {
        tracked.acceptAll();
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', action, count: tracked.items.length });
    }
  }

  // ==================== 7. WORD_REFERENCES ====================
  if (toolName === 'word_references') {
    if (op === 'manage_note') {
      const text = String(args.text || 'Chú thích tài liệu.');
      const firstPara = body.paragraphs.getFirst();
      firstPara.getRange(Word.RangeLocation.end).insertFootnote(text);
      await context.sync();
      return JSON.stringify({ status: 'ok', footnoteInserted: text });
    }

    if (op === 'insert_bookmark') {
      const rawName = String(args.name || 'Moc_Tai_Lieu').replace(/[^A-Za-z0-9_]/g, '_');
      const name = /^[A-Za-z]/.test(rawName) ? rawName : `BM_${rawName}`;
      const p = body.paragraphs.getLast ? body.paragraphs.getLast() : body.paragraphs.getFirst();
      p.getRange(Word.RangeLocation.whole).insertBookmark(name);
      await context.sync();
      return JSON.stringify({ status: 'ok', bookmarkName: name });
    }

    if (op === 'insert_reference_list') {
      const title = String(args.title || 'MỤC LỤC');
      const tocPara = body.insertParagraph('', Word.InsertLocation.start);
      tocPara.getRange(Word.RangeLocation.start).insertField(Word.InsertLocation.start, Word.FieldType.toc, '\\o "1-3" \\h \\z \\u', true);
      const headingPara = body.insertParagraph(title, Word.InsertLocation.start);
      headingPara.alignment = Word.Alignment.centered;
      headingPara.font.name = 'Times New Roman';
      headingPara.font.size = 14;
      headingPara.font.bold = true;
      headingPara.font.underline = Word.UnderlineType.none;
      await context.sync();
      return JSON.stringify({ status: 'ok', referenceList: 'TOC', title });
    }
  }

  // ==================== 8. WORD_DOCUMENT ====================
  if (toolName === 'word_document') {
    if (op === 'create_document') {
      body.clear();
      const title = String(args.title || '').trim();
      if (title) {
        const pHead = body.insertParagraph('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', Word.InsertLocation.end);
        pHead.alignment = Word.Alignment.centered;
        pHead.font.name = 'Times New Roman';
        pHead.font.size = 12;
        pHead.font.bold = true;
        pHead.font.underline = Word.UnderlineType.none;
        pHead.spaceAfter = 2;

        const pSub = body.insertParagraph('Độc lập - Tự do - Hạnh phúc', Word.InsertLocation.end);
        pSub.alignment = Word.Alignment.centered;
        pSub.font.name = 'Times New Roman';
        pSub.font.size = 13;
        pSub.font.bold = true;
        pSub.font.underline = Word.UnderlineType.single;
        pSub.spaceAfter = 12;

        const pTitle = body.insertParagraph(title.toUpperCase(), Word.InsertLocation.end);
        pTitle.alignment = Word.Alignment.centered;
        pTitle.font.name = 'Times New Roman';
        pTitle.font.size = 14;
        pTitle.font.bold = true;
        pTitle.font.underline = Word.UnderlineType.none;
        pTitle.spaceBefore = 6;
        pTitle.spaceAfter = 12;
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', created: true, title });
    }

    if (op === 'fill_template') {
      const data = (args.data as Record<string, string>) || {};
      const paras = body.paragraphs;
      paras.load('text');
      await context.sync();

      let updatedCount = 0;
      for (const p of paras.items) {
        let text = p.text || '';
        let changed = false;
        for (const [k, v] of Object.entries(data)) {
          const val = String(v ?? '');
          const patterns = [`{{${k}}}`, `[${k}]`, k];
          for (const pat of patterns) {
            if (pat && text.includes(pat)) {
              text = text.split(pat).join(val);
              changed = true;
            }
          }
        }
        if (changed) {
          p.insertText(text, Word.InsertLocation.replace);
          updatedCount++;
        }
      }
      await context.sync();
      return JSON.stringify({ status: 'ok', filledParagraphs: updatedCount });
    }

    if (op === 'set_document_properties') {
      const props = context.document.properties;
      if (args.title !== undefined) props.title = String(args.title);
      if (args.author !== undefined) props.author = String(args.author);
      if (args.subject !== undefined) props.subject = String(args.subject);
      if (args.keywords !== undefined) props.keywords = String(args.keywords);
      if (args.category !== undefined) props.category = String(args.category);
      if (args.comments !== undefined) props.comments = String(args.comments);
      await context.sync();
      return JSON.stringify({ status: 'ok', updatedProperties: Object.keys(args) });
    }

    if (op === 'set_watermark') {
      const wmText = String(args.watermark || args.text || 'BẢN DỰ THẢO');
      const section = context.document.sections.getFirst();
      const header = section.getHeader(Word.HeaderFooterType.primary);
      const hp = header.insertParagraph(wmText, Word.InsertLocation.end);
      hp.alignment = Word.Alignment.centered;
      hp.font.name = 'Times New Roman';
      hp.font.size = 11;
      hp.font.color = '#A6A6A6';
      hp.font.italic = true;
      await context.sync();
      return JSON.stringify({ status: 'ok', watermark: wmText });
    }

    if (op === 'export_table') {
      const tblIdx = Number(args.table_index ?? 0);
      const tables = body.tables;
      tables.load('values,rowCount');
      await context.sync();
      if (tblIdx < 0 || tblIdx >= tables.items.length) {
        return JSON.stringify({ error: `Table index ${tblIdx} out of range (total: ${tables.items.length})` });
      }
      const tbl = tables.items[tblIdx];
      return JSON.stringify({
        status: 'ok',
        tableIndex: tblIdx,
        rowCount: tbl.rowCount,
        values: tbl.values,
      });
    }
  }

  // ==================== 9. WORD_APP ====================
  if (toolName === 'word_app') {
    if (op === 'live_insert_at_cursor') {
      const text = String(args.text || '');
      const sel = context.document.getSelection();
      const r = sel.insertText(text, Word.InsertLocation.replace);
      r.font.name = 'Times New Roman';
      r.font.size = 13;
      await context.sync();
      return JSON.stringify({ status: 'ok', insertedAtCursor: text });
    }

    if (op === 'live_scroll_to') {
      const findText = String(args.find || '').trim();
      const pIdx = args.paragraph_index !== undefined ? Number(args.paragraph_index) : -1;
      if (findText) {
        const matches = body.search(findText, { matchCase: false });
        matches.load('items');
        await context.sync();
        if (matches.items.length > 0) {
          matches.items[0].select();
          await context.sync();
          return JSON.stringify({ status: 'ok', scrolledToText: findText });
        }
      }
      const paras = body.paragraphs;
      paras.load('items');
      await context.sync();
      const targetIdx = pIdx >= 0 && pIdx < paras.items.length ? pIdx : 0;
      if (paras.items[targetIdx]) {
        paras.items[targetIdx].select();
        await context.sync();
      }
      return JSON.stringify({ status: 'ok', scrolledToParagraph: targetIdx });
    }

    if (op === 'com_save_document') {
      if (typeof context.document.save === 'function') {
        context.document.save();
        await context.sync();
      } else {
        throw new Error('Saving is not supported by this Word runtime. The document has not been confirmed saved.');
      }
      return JSON.stringify({ status: 'ok', saved: true });
    }

    if (op === 'live_set_track_changes') {
      const enabled = args.enabled !== false;
      context.document.changeTrackingMode = enabled
        ? Word.ChangeTrackingMode.trackAll
        : Word.ChangeTrackingMode.off;
      await context.sync();
      return JSON.stringify({ status: 'ok', trackChangesEnabled: enabled });
    }

    if (op === 'com_readability_statistics') {
      const paras = body.paragraphs;
      paras.load('text');
      await context.sync();
      const nonEmpty = paras.items.map((p: any) => (p.text || '').trim()).filter(Boolean);
      const fullText = nonEmpty.join(' ');
      const words = fullText ? fullText.split(/\s+/).filter(Boolean) : [];
      const sentences = fullText ? fullText.split(/[.!?]+/).map((s: string) => s.trim()).filter(Boolean) : [];
      const avgWordsPerSentence = sentences.length > 0 ? Number((words.length / sentences.length).toFixed(1)) : words.length;
      return JSON.stringify({
        status: 'ok',
        paragraphs: nonEmpty.length,
        words: words.length,
        characters: fullText.length,
        sentences: sentences.length,
        avgWordsPerSentence,
      });
    }
  }

  // ==================== 10. WORD_HELP ====================
  if (toolName === 'word_help') {
    return JSON.stringify({
      status: 'ok',
      tools: [
        'word_read(get_document_view | get_text | get_document_info | word_count | get_outline | find_text | list_elements)',
        'word_edit(insert_paragraphs | edit_selection | edit_by_format | set_paragraph_text | search_and_replace | delete_paragraphs | replace_selection_or_document)',
        'word_format(format_administrative_vn | format_selection | set_paragraph_format | format_text | apply_style | clear_formatting)',
        'word_structure(create_table | insert_list | set_section_properties | set_header_footer | insert_callout_box | insert_break)',
        'word_media(insert_field | insert_hyperlink | insert_content_control | insert_image)',
        'word_review(manage_comment | set_track_changes | resolve_revisions)',
        'word_references(manage_note | insert_bookmark | insert_reference_list)',
        'word_document(create_document | fill_template | set_document_properties | set_watermark | export_table)',
        'word_app(live_insert_at_cursor | live_scroll_to | com_save_document | live_set_track_changes | com_readability_statistics)',
        'word_help(list_tools | inspect_op)',
      ],
    });
  }

  return JSON.stringify({ status: 'ok', toolName, op });
}

/**
 * Generates clean, deterministic Office.js code for each KitchenSink4Word tool operation.
 */
export function buildKs4wScript(toolName: string, op: string, args: Record<string, unknown> = {}): string {
  const payload = JSON.stringify({
    toolName,
    op,
    args,
    cachedSelection: cachedSelectionText,
  });
  return `// [KitchenSink4Word] ${toolName}.${op}(${JSON.stringify(args)})
await Word.run(async (context) => {
  const fn = ${executeKs4wInWord.toString()};
  return await fn(context, ${payload});
});`;
}

/**
 * Silently reads a live snapshot of the current Word document & user selection
 * before the LLM generates its turn response so the AI ALWAYS sees what the user
 * has selected ("bôi đen / bôi xanh"), bolded ("bôi đậm"), or highlighted ("tô màu")!
 */
export async function captureLiveWordSnapshot(sandbox: Sandbox): Promise<string> {
  if (typeof Word === 'undefined') return '';
  try {
    const code = buildKs4wScript('word_read', 'get_document_view', {});
    const res = await sandbox.execute(code, 5000);
    if (!res.success || !res.output) return '';
    const parsed = typeof res.output === 'string' ? JSON.parse(res.output) : res.output;
    if (!parsed || typeof parsed !== 'object') return '';
    const selectedText = String(parsed.selectedText || cachedSelectionText || '').trim();
    const docView = String(parsed.documentView || '').trim();
    if (!docView && !selectedText) return '';

    const selectionBanner = selectedText
      ? `🎯 VÙNG VĂN BẢN NGƯỜI DÙNG ĐANG BÔI CHỌN (BÔI ĐEN / BÔI XANH): "${selectedText}"\n(Nếu người dùng yêu cầu sửa, xóa, in đậm hoặc viết lại phần đang bôi chọn/bôi đen/bôi xanh, hãy tác động trực tiếp vào đoạn này bằng word_edit op="edit_selection" hoặc op="search_and_replace" hoặc set_paragraph_text!)`
      : `Hiện người dùng không bôi chọn bằng chuột (nếu người dùng nói "phần bôi đậm / tô màu / gạch chân", hãy nhìn các nhãn Bold/InĐậm, Highlight/BôiMàu, Underline/GạchChân ở danh sách đoạn văn bên dưới và dùng word_edit op="edit_by_format" hoặc set_paragraph_text).`;

    return `\n\n=== TRẠNG THÁI TÀI LIỆU WORD HIỆN TẠI (LIVE SNAPSHOT) ===\n${selectionBanner}\n\nDanh sách đoạn văn hiện tại trong tài liệu:\n${docView || '(Tài liệu đang trống)'}\n========================================================`;
  } catch {
    return '';
  }
}

export function makeKitchenSinkWordTools(
  sandbox: Sandbox,
  settings: AppSettings,
  callbacks: OrchestratorCallbacks,
): Record<string, Tool> {
  const tools: Record<string, Tool> = {};
  // The SDK may request several tools concurrently. Keep each edit + read-back atomic
  // relative to other built-in calls so a second edit cannot contaminate its evidence.
  let queue: Promise<unknown> = Promise.resolve();
  const sequential = <T>(work: () => Promise<T>): Promise<T> => {
    const next = queue.then(work);
    queue = next.catch(() => {});
    return next;
  };

  for (const meta of KITCHENSINK_WORD_TOOLS_META) {
    tools[meta.name] = tool({
      description: `${meta.titleEn}: ${meta.descEn} Supported ops: ${meta.ops.join(', ')}.`,
      inputSchema: jsonSchema<Ks4wCallInput>({
        type: 'object',
        properties: {
          op: {
            type: 'string',
            enum: meta.ops,
            description: `Operation to execute in ${meta.name}: ${meta.ops.join(' | ')}`,
          },
          args: {
            type: 'object',
            description: 'Arguments object for the selected op.',
            additionalProperties: true,
          },
        },
        required: ['op'],
        additionalProperties: false,
      }),
      execute: async ({ op, args = {} }, { toolCallId }) => sequential(async () => {
        const code = buildKs4wScript(meta.name, op, args);
        try {
          if (!settings.autoApprove) {
            callbacks.onUpsertCodeBlock(toolCallId, { code, status: 'pending' });
          }
          const approved = settings.autoApprove || await callbacks.requestApproval(code);
          if (!approved) {
            callbacks.onUpsertCodeBlock(toolCallId, { code, status: 'rejected' });
            return 'User rejected the operation.';
          }

          callbacks.onUpsertCodeBlock(toolCallId, { code, status: 'running' });
          const mutation = isWordMutation(meta.name, op, args);
          const selectionBefore = getCachedWordSelection();
          let before: WordObservation | undefined;
          if (mutation) {
            try { before = await observeWordDocument(sandbox, settings.executionTimeout); }
            catch { /* Still read after the action, but do not certify it without a baseline. */ }
          }
          const result = await sandbox.execute(code, settings.executionTimeout);

          if (result.success) {
            const outputText = typeof result.output === 'string'
              ? result.output
              : JSON.stringify(result.output, null, 2);
            if (mutation) {
              callbacks.onUpsertCodeBlock(toolCallId, {
                code, status: 'running', verification: { status: 'checking', detail: 'Reading the document back after the operation.' },
              });
              let verification: WordVerification;
              let readBack = '';
              try {
                const after = await observeWordDocument(sandbox, settings.executionTimeout);
                readBack = wordReadBackForModel(after);
                verification = before
                  ? verifyWordMutation(meta.name, op, args, before, after, selectionBefore)
                  : { status: 'unverified', detail: 'Read-back succeeded, but the pre-edit document could not be observed.' };
              } catch (error) {
                verification = { status: 'unverified', detail: `The operation ran, but document read-back failed: ${error instanceof Error ? error.message : String(error)}` };
              }
              const evidence = JSON.stringify({ operationOutput: outputText, verification, readBack,
                instruction: verification.status === 'passed'
                  ? 'These measured postconditions passed. Compare the observed result against the user request before reporting completion; this is not proof of every aspect of the request.'
                  : 'Do not claim verified completion. The operation may already have modified the document. Explain the uncertainty or mismatch, inspect the current document, and do not blindly repeat the operation.' });
              callbacks.onUpsertCodeBlock(toolCallId, {
                code, status: verification.status === 'failed' ? 'error' : 'success',
                result: evidence, verification,
              });
              return evidence;
            }
            callbacks.onUpsertCodeBlock(toolCallId, {
              code,
              status: 'success',
              result: `Output:\n${outputText}`,
            });
            return outputText || JSON.stringify({ status: 'ok', tool: meta.name, op });
          }

          const uiResult = `Error: ${result.error}\n${result.stack || ''}`;
          callbacks.onUpsertCodeBlock(toolCallId, { code, status: 'error', result: uiResult });
          return `Operation ${meta.name}.${op} failed: ${result.error}`;
        } catch (err) {
          const formatted = formatError(err, { phase: 'tool-execute', tool: meta.name });
          callbacks.onMessage({ role: 'assistant', content: '', error: formatted });
          return `Tool failed: ${formatted.title}. ${formatted.detail}`;
        }
      }),
    });
  }

  return tools;
}
