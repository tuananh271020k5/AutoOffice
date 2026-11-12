import { describe, expect, it } from 'vitest';
import { render, screen } from '@testing-library/react';
import { CodeBlock } from './CodeBlock';
import { LanguageProvider } from '../i18n';

describe('Word verification status shown to users', () => {
  it('does not show success when the Word API ran but verification was unavailable', async () => {
    render(<LanguageProvider initialLocale="vi"><CodeBlock
      code="// [AutoOffice] word_edit.edit_selection" status="success"
      verification={{ status: 'unverified', detail: 'Read timed out' }}
    /></LanguageProvider>);
    expect(await screen.findByText('Chưa xác minh')).toBeInTheDocument();
    expect(screen.queryByText('Hoàn tất')).not.toBeInTheDocument();
    expect(screen.getByText(/Tài liệu có thể đã thay đổi/)).toBeInTheDocument();
    expect(screen.queryByText('Read timed out')).not.toBeInTheDocument();
  });
  it('warns about a mismatch without telling users to blindly retry a completed mutation', async () => {
    render(<LanguageProvider initialLocale="en"><CodeBlock
      code="" status="error" result="Mismatch"
      verification={{ status: 'failed', detail: 'Mismatch' }}
    /></LanguageProvider>);
    expect(await screen.findByText('Result does not match')).toBeInTheDocument();
    expect(screen.queryByText(/Please try again/)).not.toBeInTheDocument();
  });
  it('shows checking and then checked as the evidence becomes available', async () => {
    const { rerender } = render(<LanguageProvider initialLocale="en"><CodeBlock code="" status="running"
      verification={{ status: 'checking', detail: '' }} /></LanguageProvider>);
    expect(await screen.findByText('Checking document…')).toBeInTheDocument();
    rerender(<LanguageProvider initialLocale="en"><CodeBlock code="" status="success"
      verification={{ status: 'passed', detail: '' }} /></LanguageProvider>);
    expect(screen.getByText('Checked')).toBeInTheDocument();
  });
});
