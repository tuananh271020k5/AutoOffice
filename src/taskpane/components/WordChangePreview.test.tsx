import { describe, expect, it, vi } from 'vitest';
import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { LanguageProvider } from '../i18n';
import { WordChangePreview } from './WordChangePreview';
import { WordRecoveryPanel } from './WordRecoveryPanel';

describe('preview and recovery controls', () => {
  it('shows content as plain text and applies only when the user clicks apply', async () => {
    const onApprove = vi.fn();
    const onReject = vi.fn();
    render(<LanguageProvider initialLocale="en"><WordChangePreview preview={{
      kind: 'replace', scope: 'selection', before: 'Old paragraph', proposed: '<script>danger()</script>',
      properties: [{ name: 'bold', value: 'true' }], truncated: false,
    }} onApprove={onApprove} onReject={onReject} /></LanguageProvider>);
    expect(await screen.findByRole('region', { name: 'Review changes' })).toBeInTheDocument();
    expect(screen.getByText('Old paragraph')).toBeInTheDocument();
    expect(screen.getByText('<script>danger()</script>')).toBeInTheDocument();
    expect(document.querySelector('script')).toBeNull();
    expect(onApprove).not.toHaveBeenCalled();
    await userEvent.click(screen.getByRole('button', { name: 'Apply changes' }));
    expect(onApprove).toHaveBeenCalledOnce();
    expect(onReject).not.toHaveBeenCalled();
  });
  it('rejects without applying and warns when the preview is shortened', async () => {
    const onApprove = vi.fn();
    const onReject = vi.fn();
    render(<LanguageProvider initialLocale="en"><WordChangePreview preview={{
      kind: 'delete', scope: 'document', before: 'Existing content', proposed: '', properties: [], truncated: true,
    }} onApprove={onApprove} onReject={onReject} /></LanguageProvider>);
    await userEvent.click(await screen.findByRole('button', { name: 'Keep document unchanged' }));
    expect(onReject).toHaveBeenCalledOnce();
    expect(onApprove).not.toHaveBeenCalled();
    expect(screen.getByText(/preview is shortened/)).toBeInTheDocument();
  });
  it('disables recovery while edits are running and explains that it opens a separate document', async () => {
    render(<LanguageProvider initialLocale="en"><WordRecoveryPanel documentLabel="Test.docx" createdAt={1}
      preparing={false} opening={false} busy error={false} onOpen={vi.fn()} /></LanguageProvider>);
    const button = await screen.findByRole('button', { name: 'Open recovery copy' });
    expect(button).toBeDisabled();
    expect(screen.getByText(/Opens a separate Word document/)).toBeInTheDocument();
  });
});
