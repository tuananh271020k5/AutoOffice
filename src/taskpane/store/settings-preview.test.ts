import { afterEach, describe, expect, it } from 'vitest';
import { loadSettings } from './settings';

afterEach(() => localStorage.removeItem('autooffice_settings'));
describe('Word protection settings migration', () => {
  it('enables preview/recovery for existing settings without changing their auto-approval choice', () => {
    localStorage.setItem('autooffice_settings', JSON.stringify({ autoApprove: true }));
    const settings = loadSettings();
    expect(settings.previewWordChanges).toBe(true);
    expect(settings.autoApprove).toBe(true);
  });
  it('preserves an explicit choice to turn Word protection off', () => {
    localStorage.setItem('autooffice_settings', JSON.stringify({ previewWordChanges: false }));
    expect(loadSettings().previewWordChanges).toBe(false);
  });
});
