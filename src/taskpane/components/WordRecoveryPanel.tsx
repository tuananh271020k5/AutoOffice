import { Button, Spinner } from '@fluentui/react-components';
import { useTranslation } from '../i18n';

export function WordRecoveryPanel({ documentLabel, createdAt, preparing, opening, busy, error, onOpen }: {
  documentLabel?: string; createdAt?: number; preparing: boolean; opening: boolean;
  busy: boolean; error: boolean; onOpen: () => void;
}) {
  const { t, locale } = useTranslation();
  if (!documentLabel && !preparing && !error) return null;
  return <section aria-label={t('recovery.title')} style={{ flexShrink: 0, padding: '8px 10px',
    borderTop: '1px solid #ccc', fontSize: 12 }}>
    {preparing || opening ? <div role="status" style={{ display: 'flex', gap: 6 }}>
      <Spinner size="extra-tiny" />{t(preparing ? 'recovery.preparing' : 'recovery.opening')}
    </div> : <>
      {documentLabel && <>
        <strong>{t('recovery.title')}</strong>
        <div>{documentLabel} · {createdAt ? new Date(createdAt).toLocaleTimeString(locale) : ''}</div>
        <p style={{ margin: '4px 0' }}>{t('recovery.notice')}</p>
        <Button size="small" disabled={busy} onClick={onOpen}>{t('recovery.open')}</Button>
      </>}
      {error && <p role="alert">{t('recovery.failure')}</p>}
    </>}
  </section>;
}
