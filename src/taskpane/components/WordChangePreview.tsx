import { Button, Text } from '@fluentui/react-components';
import { useTranslation, type TranslationKey } from '../i18n';
import type { WordChangePreview as Preview } from '../agent/word-preview';

const kinds: Record<Preview['kind'], TranslationKey> = {
  replace: 'preview.replace', delete: 'preview.delete', insert: 'preview.insert',
  format: 'preview.format', table: 'preview.table', other: 'preview.other',
};
const scopes: Record<Preview['scope'], TranslationKey> = {
  selection: 'preview.scopeSelection', paragraphs: 'preview.scopeParagraphs', document: 'preview.scopeDocument',
};
const properties: Record<string, TranslationKey> = {
  font_name: 'preview.font', font_size: 'preview.fontSize', bold: 'preview.bold', italic: 'preview.italic',
  underline: 'preview.underline', color: 'preview.color', highlight_color: 'preview.highlight',
  alignment: 'preview.alignment', space_before: 'preview.spaceBefore', space_after: 'preview.spaceAfter',
  line_spacing_lines: 'preview.lineSpacing',
};

export function WordChangePreview({ preview, onApprove, onReject }: {
  preview: Preview; onApprove: () => void; onReject: () => void;
}) {
  const { t } = useTranslation();
  const box = { margin: '4px 0 10px', padding: 8, border: '1px solid #c8c8c8',
    borderRadius: 4, whiteSpace: 'pre-wrap' as const, overflowWrap: 'anywhere' as const };
  return (
    <section aria-label={t('preview.title')} style={{ fontSize: 12 }}>
      <Text weight="semibold">{t('preview.title')}: {t(kinds[preview.kind])}</Text>
      <p>{t(preview.phase === 'applied' ? 'preview.appliedNotice' : 'preview.planNotice')}</p>
      <details>
      <summary>{t('preview.details')}</summary>
      <p>{t(scopes[preview.scope])}</p>
      <strong>{t('preview.before')}</strong>
      <div style={box}>{preview.before || t('preview.empty')}</div>
      {preview.proposed !== undefined && <>
        <strong>{t('preview.proposed')}</strong>
        <div style={box}>{preview.proposed || t('preview.empty')}</div>
      </>}
      {preview.properties.length > 0 && <dl>
        {preview.properties.map(property => <div key={property.name}>
          <dt style={{ display: 'inline' }}>{t(properties[property.name])}: </dt>
          <dd style={{ display: 'inline', margin: 0 }}>
            {property.value === 'true' ? t('preview.enabled') : property.value === 'false' ? t('preview.disabled') : property.value}
          </dd>
        </div>)}
      </dl>}
      {preview.table && <div style={{ overflowX: 'auto', margin: '8px 0' }}>
        <table style={{ borderCollapse: 'collapse' }}><tbody>
          {preview.table.map((row, i) => <tr key={i}>{row.map((cell, j) =>
            <td key={j} style={{ border: '1px solid #aaa', padding: 4 }}>{cell}</td>)}</tr>)}
        </tbody></table>
      </div>}
      {preview.kind === 'other' && <p>{t('preview.otherNotice')}</p>}
      {preview.truncated && <p role="note">{t('preview.truncated')}</p>}
      </details>
      <div style={{ display: 'flex', gap: 8, paddingTop: 6 }}>
        <Button size="small" appearance="primary" onClick={onApprove}>{preview.phase === 'applied' ? 'Accept' : t('preview.apply')}</Button>
        <Button size="small" onClick={onReject}>{preview.phase === 'applied' ? 'Reject' : t('preview.reject')}</Button>
      </div>
    </section>
  );
}
