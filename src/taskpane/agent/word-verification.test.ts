import { describe, expect, it, vi } from 'vitest';
import { isWordMutation, observeWordDocument, verifyWordMutation, type WordObservation } from './word-verification';
import { makeKitchenSinkWordTools } from './kitchensink-word-tools';
import type { Sandbox } from '../executor/sandbox';
import type { AppSettings } from '../store/settings';
import type { OrchestratorCallbacks } from './orchestrator';

function observation(texts = ['Before', 'Keep']): WordObservation {
  const font = { name: 'Calibri', size: 11, bold: false, italic: false, underline: 'None', color: '#000000', highlightColor: null };
  return { text: texts.join('\r') + '\r', paragraphs: texts.map(text => ({ text, alignment: 'Left', font: { ...font } })),
    selection: { text: texts[0], font: { ...font } }, tables: [] };
}
const settings = { autoApprove: true, executionTimeout: 30000 } as AppSettings;
function setup() {
  const execute = vi.fn();
  const callbacks: OrchestratorCallbacks = {
    onMessage: vi.fn(), onStreamToken: vi.fn(), onUpsertCodeBlock: vi.fn(), requestApproval: vi.fn(async () => true),
  };
  const sandbox = { execute } as unknown as Sandbox;
  const tools = makeKitchenSinkWordTools(sandbox, settings, callbacks);
  const call = (name: string, op: string, args = {}) => tools[name].execute!(
    { op, args }, { toolCallId: name + op, messages: [] },
  );
  return { execute, callbacks, call };
}
const result = (output: unknown) => ({ success: true, output });

describe('Word postcondition checks', () => {
  it('detects a no-op replacement even when the tool reported success', () => {
    const before = observation();
    expect(verifyWordMutation('word_edit', 'edit_selection', { new_text: 'After' }, before, before).status).toBe('failed');
    const after = observation(['After', 'Keep']);
    expect(verifyWordMutation('word_edit', 'edit_selection', { new_text: 'After' }, before, after).status).toBe('passed');
    after.text = 'After\rAccidentally changed\r';
    expect(verifyWordMutation('word_edit', 'edit_selection', { new_text: 'After' }, before, after).status).toBe('failed');
  });
  it('does not guess between repeated text ranges', () => {
    const before = observation(['Before', 'Before']);
    expect(verifyWordMutation('word_edit', 'edit_selection', { new_text: 'After' }, before, observation(['After', 'Before'])).status).toBe('unverified');
  });
  it('also checks formatting when replacement is used to bold a full paragraph', () => {
    const before = observation();
    const after = observation();
    const args = { new_text: 'Before', bold: true };
    expect(verifyWordMutation('word_edit', 'edit_selection', args, before, after).status).toBe('failed');
    after.paragraphs[0].font.bold = true;
    expect(verifyWordMutation('word_edit', 'edit_selection', args, before, after).status).toBe('passed');
    const partialBefore = observation(['Some Before text', 'Keep']);
    partialBefore.selection.text = 'Before';
    expect(verifyWordMutation('word_edit', 'edit_selection', args, partialBefore, partialBefore).status).toBe('unverified');
  });
  it('detects wrong paragraph and content changed outside the requested paragraph', () => {
    const before = observation();
    expect(verifyWordMutation('word_edit', 'set_paragraph_text', { paragraph_index: 0, new_text: 'After' }, before, observation(['After', 'Keep'])).status).toBe('passed');
    expect(verifyWordMutation('word_edit', 'set_paragraph_text', { paragraph_index: 0, new_text: 'After' }, before, observation(['Before', 'After'])).status).toBe('failed');
  });
  it('accepts already-correct formatting without requiring a change', () => {
    const before = observation();
    before.selection.font.bold = true;
    expect(verifyWordMutation('word_format', 'format_selection', { bold: true }, before, before).status).toBe('passed');
  });
  it('rejects missing formatting and unexpected text changes', () => {
    const before = observation();
    expect(verifyWordMutation('word_format', 'format_selection', { bold: true }, before, before).status).toBe('failed');
    const after = observation(['Changed', 'Keep']);
    after.selection.font.bold = true;
    expect(verifyWordMutation('word_format', 'format_selection', { bold: true }, before, after).status).toBe('failed');
  });
  it('does not certify paragraph layout or mixed formatting it did not measure', () => {
    const before = observation();
    expect(verifyWordMutation('word_format', 'format_selection', { alignment: 'Centered' }, before, before).status).toBe('unverified');
    const after = observation();
    after.selection.font.bold = null;
    expect(verifyWordMutation('word_format', 'format_selection', { bold: true }, before, after).status).toBe('failed');
  });
  it('checks table dimensions and values, not the mutation output', () => {
    const before = observation();
    const after = observation();
    after.tables = [[['Name', 'Count'], ['Pen', '2']]];
    expect(verifyWordMutation('word_structure', 'create_table', { data: after.tables[0] }, before, after).status).toBe('passed');
    expect(verifyWordMutation('word_structure', 'create_table', { data: [['Wrong']] }, before, after).status).toBe('failed');
    expect(verifyWordMutation('word_structure', 'create_table', { data: after.tables[0], header_bg: '#ff0000' }, before, after).status).toBe('unverified');
  });
  it('does not claim saved state or page margins from body content', () => {
    const state = observation();
    expect(verifyWordMutation('word_app', 'com_save_document', {}, state, state).status).toBe('unverified');
    expect(verifyWordMutation('word_structure', 'set_section_properties', {}, state, state).status).toBe('unverified');
  });
  it('skips extra reads for read-only operations', () => {
    expect(isWordMutation('word_read', 'get_text', {})).toBe(false);
    expect(isWordMutation('word_review', 'manage_comment', { action: 'list' })).toBe(false);
    expect(isWordMutation('word_document', 'export_table', {})).toBe(false);
    expect(isWordMutation('word_review', 'manage_comment', { action: 'add' })).toBe(true);
  });
});

describe('Word tool verification integration', () => {
  it('reads before/action/after and sends independent evidence back to the model', async () => {
    const { execute, callbacks, call } = setup();
    execute.mockResolvedValueOnce(result(observation()))
      .mockResolvedValueOnce(result('{"status":"ok"}'))
      .mockResolvedValueOnce(result(observation(['After', 'Keep'])));
    const output = JSON.parse(await call('word_edit', 'edit_selection', { new_text: 'After' }) as string);
    expect(execute).toHaveBeenCalledTimes(3);
    expect(execute.mock.calls[0][0]).toContain('verify.read_document');
    expect(execute.mock.calls[1][0]).toContain('word_edit.edit_selection');
    expect(execute.mock.calls[2][0]).toContain('verify.read_document');
    expect(output.verification.status).toBe('passed');
    expect(output.readBack).toContain('After');
    expect(callbacks.onUpsertCodeBlock).toHaveBeenLastCalledWith('word_editedit_selection', expect.objectContaining({
      status: 'success', verification: expect.objectContaining({ status: 'passed' }),
    }));
  });
  it('marks false success as a mismatch without replaying the action', async () => {
    const { execute, callbacks, call } = setup();
    execute.mockResolvedValueOnce(result(observation())).mockResolvedValueOnce(result('{"status":"ok"}'))
      .mockResolvedValueOnce(result(observation()));
    const output = JSON.parse(await call('word_edit', 'edit_selection', { new_text: 'After' }) as string);
    expect(output.verification.status).toBe('failed');
    expect(execute).toHaveBeenCalledTimes(3);
    expect(callbacks.onUpsertCodeBlock).toHaveBeenLastCalledWith(expect.any(String), expect.objectContaining({ status: 'error' }));
  });
  it('warns when post-read fails rather than converting execution success into verified success', async () => {
    const { execute, call } = setup();
    execute.mockResolvedValueOnce(result(observation())).mockResolvedValueOnce(result('ok'))
      .mockResolvedValueOnce({ success: false, error: 'Read timeout' });
    const output = JSON.parse(await call('word_edit', 'edit_selection', { new_text: 'After' }) as string);
    expect(output.verification.status).toBe('unverified');
    expect(output.verification.detail).toContain('Read timeout');
    expect(execute).toHaveBeenCalledTimes(3);
  });
  it('does not certify without a pre-edit baseline and does not verify rejected actions', async () => {
    const { execute, call } = setup();
    execute.mockResolvedValueOnce({ success: false, error: 'Read failed' }).mockResolvedValueOnce(result('ok'))
      .mockResolvedValueOnce(result(observation(['After', 'Keep'])));
    expect(JSON.parse(await call('word_edit', 'edit_selection', { new_text: 'After' }) as string).verification.status).toBe('unverified');
    const callbacks = { onMessage: vi.fn(), onStreamToken: vi.fn(), onUpsertCodeBlock: vi.fn(), requestApproval: vi.fn(async () => false) };
    const sandbox = { execute: vi.fn() } as unknown as Sandbox;
    const tool = makeKitchenSinkWordTools(sandbox, { ...settings, autoApprove: false }, callbacks).word_edit;
    await tool.execute!({ op: 'edit_selection', args: { new_text: 'After' } }, { toolCallId: 'rejected', messages: [] });
    expect(sandbox.execute).not.toHaveBeenCalled();
  });
  it('serializes concurrent edits so each read-back belongs to its own mutation', async () => {
    const { execute, call } = setup();
    execute.mockResolvedValueOnce(result(observation())).mockResolvedValueOnce(result('ok'))
      .mockResolvedValueOnce(result(observation(['After', 'Keep'])))
      .mockResolvedValueOnce(result(observation(['After', 'Keep']))).mockResolvedValueOnce(result('ok'))
      .mockResolvedValueOnce(result(observation(['Final', 'Keep'])));
    const outputs = await Promise.all([
      call('word_edit', 'set_paragraph_text', { paragraph_index: 0, new_text: 'After' }),
      call('word_edit', 'set_paragraph_text', { paragraph_index: 0, new_text: 'Final' }),
    ]);
    expect(outputs.map(o => JSON.parse(o as string).verification.status)).toEqual(['passed', 'passed']);
    expect(execute.mock.calls[2][0]).toContain('verify.read_document');
    expect(execute.mock.calls[3][0]).toContain('verify.read_document');
  });
  it('performs one execution for reads, and rejects malformed read-back', async () => {
    const { execute, call } = setup();
    execute.mockResolvedValueOnce(result('read result'));
    expect(await call('word_read', 'get_text')).toBe('read result');
    expect(execute).toHaveBeenCalledTimes(1);
    execute.mockResolvedValueOnce(result({ text: 'not a complete snapshot' }));
    await expect(observeWordDocument({ execute } as unknown as Sandbox, 5000)).rejects.toThrow('Invalid Word');
  });
});
