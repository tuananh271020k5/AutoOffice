import React, { useMemo } from 'react';
import { makeStyles, tokens, Button, Text, Spinner } from '@fluentui/react-components';
import {
  CheckmarkCircle16Filled,
  DismissCircle16Filled,
  DismissCircle24Regular,
  Play24Regular,
} from '@fluentui/react-icons';
import { useTranslation } from '../i18n/index.ts';

const useStyles = makeStyles({
  compactRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    flexWrap: 'wrap',
    gap: '6px',
    padding: '6px 10px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: '8px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    maxWidth: '100%',
    minWidth: 0,
    boxSizing: 'border-box',
  },
  leftGroup: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    flex: '1 1 0%',
    minWidth: 0,
  },
  label: {
    flex: '1 1 0%',
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    color: tokens.colorNeutralForeground2,
    fontSize: '12px',
  },
  iconSuccess: {
    color: tokens.colorPaletteGreenForeground1,
    flexShrink: 0,
  },
  iconError: {
    color: tokens.colorPaletteRedForeground1,
    flexShrink: 0,
  },
  statusPill: {
    display: 'inline-block',
    padding: '1px 7px',
    borderRadius: '999px',
    fontSize: '11px',
    fontWeight: 600,
    lineHeight: '16px',
    whiteSpace: 'nowrap',
    flexShrink: 0,
  },
  statusInformative: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
  },
  statusSuccess: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground1,
  },
  statusDanger: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorPaletteRedForeground1,
  },
  statusWarning: {
    backgroundColor: tokens.colorPaletteYellowBackground2,
    color: tokens.colorPaletteYellowForeground2,
  },
  actions: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '6px',
    width: '100%',
    paddingTop: '4px',
  },
  errorText: {
    width: '100%',
    fontSize: '11.5px',
    color: tokens.colorPaletteRedForeground1,
    wordBreak: 'break-word',
    paddingTop: '2px',
  },
});

type CodeStatus = 'streaming' | 'pending' | 'rejected' | 'running' | 'success' | 'error';

interface CodeBlockProps {
  code: string;
  status: CodeStatus;
  result?: string;
  onApprove?: () => void;
  onReject?: () => void;
}

const FRIENDLY_ACTION_LABELS_VI: Record<string, string> = {
  'word_read.get_document_view': 'Đọc nội dung tài liệu',
  'word_read.get_text': 'Đọc văn bản tài liệu',
  'word_read.get_document_info': 'Kiểm tra thống kê tài liệu',
  'word_read.word_count': 'Đếm số từ & đoạn văn',
  'word_read.get_outline': 'Đọc cấu trúc tiêu đề',
  'word_read.find_text': 'Tìm kiếm văn bản',
  'word_read.list_elements': 'Kiểm tra thành phần tài liệu',
  'word_edit.insert_paragraphs': 'Soạn thảo nội dung vào tài liệu',
  'word_edit.edit_selection': 'Chỉnh sửa vùng văn bản đang chọn',
  'word_edit.replace_selection_or_document': 'Cập nhật nội dung văn bản',
  'word_edit.edit_by_format': 'Chỉnh sửa phần văn bản theo định dạng',
  'word_edit.set_paragraph_text': 'Cập nhật đoạn văn',
  'word_edit.search_and_replace': 'Tìm và thay thế nội dung',
  'word_edit.delete_paragraphs': 'Xóa đoạn văn bản',
  'word_format.format_administrative_vn': 'Chuẩn hóa thể thức văn bản hành chính',
  'word_format.set_font': 'Định dạng phông chữ',
  'word_format.set_paragraph': 'Căn chỉnh đoạn văn & giãn dòng',
  'word_format.apply_style': 'Áp dụng kiểu tiêu đề',
  'word_format.highlight_and_color': 'Tô màu & làm nổi bật văn bản',
  'word_format.create_list': 'Tạo danh sách liệt kê',
  'word_structure.insert_table': 'Tạo bảng biểu trong tài liệu',
  'word_structure.format_table': 'Định dạng bảng biểu',
  'word_structure.table_row_col': 'Cập nhật hàng/cột bảng biểu',
  'word_structure.insert_break': 'Chèn ngắt trang',
  'word_structure.set_page_setup': 'Thiết lập khổ giấy & căn lề trang',
  'word_structure.set_header_footer': 'Thiết lập tiêu đề đầu/chân trang',
  'word_media.insert_image': 'Chèn hình ảnh vào tài liệu',
  'word_media.resize_image': 'Điều chỉnh kích thước ảnh',
  'word_media.set_hyperlink': 'Chèn liên kết vào văn bản',
  'word_media.manage_bookmark': 'Cập nhật đánh dấu trang',
  'word_review.add_comment': 'Thêm nhận xét vào tài liệu',
  'word_review.list_comments': 'Đọc danh sách nhận xét',
  'word_review.delete_comments': 'Xóa nhận xét',
  'word_review.set_track_changes': 'Cấu hình theo dõi sửa đổi',
  'word_references.insert_toc': 'Tạo mục lục tự động',
  'word_references.update_fields': 'Cập nhật trường dữ liệu & mục lục',
  'word_references.insert_footnote': 'Chèn chú thích chân trang',
  'word_references.insert_endnote': 'Chèn chú thích cuối tài liệu',
  'word_references.insert_field': 'Chèn trường thông tin tự động',
  'word_document.insert_content_control': 'Tạo vùng điều khiển nội dung',
  'word_document.fill_content_control': 'Điền thông tin biểu mẫu',
  'word_document.get_set_properties': 'Cập nhật thông tin tài liệu',
  'word_document.ooxml_transfer': 'Xử lý cấu trúc tài liệu',
};

function getFriendlyActionLabel(rawCode: string, isVi: boolean): string {
  const trimmed = rawCode.trim();
  const firstLine = trimmed.split('\n')[0] || '';
  const match = /^\/\/\s*\[(?:KitchenSink4Word|AutoOffice)\]\s*([a-z0-9_]+)\.([a-z0-9_]+)/i.exec(firstLine);
  if (match) {
    const key = `${match[1]}.${match[2]}`;
    if (isVi && FRIENDLY_ACTION_LABELS_VI[key]) {
      return FRIENDLY_ACTION_LABELS_VI[key];
    }
    if (!isVi) {
      return `Document action (${match[2].replace(/_/g, ' ')})`;
    }
  }
  return isVi ? 'Thực hiện thao tác trên tài liệu' : 'Executing document action';
}

export function CodeBlock({ code, status, result, onApprove, onReject }: CodeBlockProps) {
  const styles = useStyles();
  const { t, locale } = useTranslation();
  const isVi = locale === 'vi';

  const STATUS_LABELS = {
    streaming: t('code.statusStreaming'),
    pending:   t('code.statusPending'),
    rejected:  t('code.statusRejected'),
    running:   t('code.statusRunning'),
    success:   t('code.statusSuccess'),
    error:     t('code.statusError'),
  } as const;

  const friendlyLabel = useMemo(() => getFriendlyActionLabel(code, isVi), [code, isVi]);

  const statusPillClass =
    status === 'success' ? styles.statusSuccess :
    status === 'error' ? styles.statusDanger :
    status === 'rejected' ? styles.statusWarning :
    styles.statusInformative;

  const isRunning = status === 'streaming' || status === 'running';

  return (
    <div className={styles.compactRow}>
      <div className={styles.leftGroup}>
        {isRunning ? (
          <Spinner size="extra-tiny" />
        ) : status === 'error' ? (
          <DismissCircle16Filled className={styles.iconError} />
        ) : (
          <CheckmarkCircle16Filled className={styles.iconSuccess} />
        )}
        <Text className={styles.label}>{friendlyLabel}</Text>
      </div>

      <span className={`${styles.statusPill} ${statusPillClass}`}>
        {STATUS_LABELS[status]}
      </span>

      {status === 'pending' && onApprove && onReject && (
        <div className={styles.actions}>
          <Button appearance="primary" icon={<Play24Regular />} size="small" onClick={onApprove}>
            {t('code.approveButton')}
          </Button>
          <Button appearance="subtle" icon={<DismissCircle24Regular />} size="small" onClick={onReject}>
            {t('code.rejectButton')}
          </Button>
        </div>
      )}

      {status === 'error' && result && (
        <div className={styles.errorText}>
          {isVi
            ? 'Không thể thực hiện thao tác này trên tài liệu. Vui lòng thử lại.'
            : 'Could not complete this action on the document. Please try again.'}
        </div>
      )}
    </div>
  );
}
