import React, { useState, useRef, useEffect, useLayoutEffect, useMemo } from 'react';
import {
  makeStyles,
  tokens,
  Badge,
  Textarea,
  Button,
  Text,
  Tooltip,
  Input,
  Select,
} from '@fluentui/react-components';
import {
  Send24Regular,
  Settings24Regular,
  History24Regular,
  Add24Regular,
  Sparkle16Regular,
  Key20Regular,
  Dismiss16Regular,
} from '@fluentui/react-icons';
import type { ChatMessage } from '../agent/orchestrator.ts';
import type { HostContext, HostKind } from '../host/context.ts';
import { CrossHostBanner } from './CrossHostBanner.tsx';
import { useTranslation, type TranslationKey } from '../i18n/index.ts';
import { MessageBubble } from './MessageBubble.tsx';
import { CodeBlock } from './CodeBlock.tsx';
import { WordChangePreview } from './WordChangePreview';
import type { WordChangePreview as Preview } from '../agent/word-preview';
import { CostBadge } from './CostBadge.tsx';
import type { CallCost } from '../agent/pricing.ts';
import { listSkills } from '../skills/index.ts';
import {
  KITCHENSINK_WORD_TOOLS_META,
  initWordSelectionTracker,
  getCachedWordSelection,
  setCachedWordSelection,
} from '../agent/kitchensink-word-tools.ts';

interface QuickActionDef {
  id: string;
  labelKey: TranslationKey;
  promptKey: TranslationKey;
}

const QUICK_ACTIONS_BY_HOST: Record<HostKind, QuickActionDef[]> = {
  word: [
    {
      id: 'word-summarize',
      labelKey: 'quickActions.wordSummarizeLabel',
      promptKey: 'quickActions.wordSummarizePrompt',
    },
    {
      id: 'word-formalize',
      labelKey: 'quickActions.wordFormalizeLabel',
      promptKey: 'quickActions.wordFormalizePrompt',
    },
    {
      id: 'word-proofread',
      labelKey: 'quickActions.wordProofreadLabel',
      promptKey: 'quickActions.wordProofreadPrompt',
    },
    {
      id: 'word-format-standard',
      labelKey: 'quickActions.wordFormatStandardLabel',
      promptKey: 'quickActions.wordFormatStandardPrompt',
    },
    {
      id: 'word-table-toc',
      labelKey: 'quickActions.wordTableTocLabel',
      promptKey: 'quickActions.wordTableTocPrompt',
    },
    {
      id: 'word-translate',
      labelKey: 'quickActions.wordTranslateLabel',
      promptKey: 'quickActions.wordTranslatePrompt',
    },
  ],
  excel: [
    {
      id: 'excel-analyze',
      labelKey: 'quickActions.excelAnalyzeLabel',
      promptKey: 'quickActions.excelAnalyzePrompt',
    },
    {
      id: 'excel-format-table',
      labelKey: 'quickActions.excelFormatTableLabel',
      promptKey: 'quickActions.excelFormatTablePrompt',
    },
    {
      id: 'excel-chart',
      labelKey: 'quickActions.excelChartLabel',
      promptKey: 'quickActions.excelChartPrompt',
    },
  ],
  powerpoint: [
    {
      id: 'ppt-add-slide',
      labelKey: 'quickActions.pptAddSlideLabel',
      promptKey: 'quickActions.pptAddSlidePrompt',
    },
    {
      id: 'ppt-format-titles',
      labelKey: 'quickActions.pptFormatTitlesLabel',
      promptKey: 'quickActions.pptFormatTitlesPrompt',
    },
  ],
};

const QUICK_DEFAULT_MODELS: Record<string, string> = {
  openai: 'gpt-4o',
  google: 'gemini-2.5-pro',
  anthropic: 'claude-sonnet-4-6',
  deepseek: 'deepseek-chat',
};

const useStyles = makeStyles({
  container: {
    flex: '1 1 0%',
    minHeight: 0,
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    width: '100%',
    minWidth: 0,
    overflow: 'hidden',
    boxSizing: 'border-box',
  },
  header: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    padding: '6px 10px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
    minWidth: 0,
  },
  headerTopRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '6px',
    minWidth: 0,
  },
  brand: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    minWidth: 0,
    overflow: 'hidden',
  },
  headerActions: {
    display: 'flex',
    alignItems: 'center',
    gap: '2px',
    flexShrink: 0,
  },
  headerSubRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '6px',
    minWidth: 0,
  },
  docSessionPill: {
    display: 'inline-block',
    flex: '1 1 0%',
    minWidth: 0,
    maxWidth: '100%',
    padding: '2px 8px',
    borderRadius: '999px',
    fontSize: '11px',
    lineHeight: '16px',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  headerMetaRight: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    flexShrink: 0,
  },
  logo: {
    width: '22px',
    height: '22px',
    flexShrink: 0,
  },
  title: {
    fontWeight: 600,
    fontSize: '15px',
    whiteSpace: 'nowrap',
  },
  hostPill: {
    display: 'inline-block',
    padding: '1px 7px',
    borderRadius: '999px',
    fontSize: '11px',
    fontWeight: 600,
    lineHeight: '16px',
    whiteSpace: 'nowrap',
    flexShrink: 0,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    border: `1px solid ${tokens.colorBrandStroke1}`,
  },
  messageList: {
    flex: 1,
    overflowY: 'auto',
    overflowX: 'hidden',
    padding: '8px 0',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    minWidth: 0,
  },
  empty: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    flex: 1,
    gap: '10px',
    color: tokens.colorNeutralForeground3,
    padding: '16px 12px',
    textAlign: 'center',
    minWidth: 0,
  },
  apiKeyCard: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    width: '100%',
    maxWidth: '400px',
    padding: '10px',
    borderRadius: '8px',
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorBrandStroke1}`,
    textAlign: 'left',
    boxSizing: 'border-box',
    minWidth: 0,
  },
  quickActionsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(125px, 1fr))',
    gap: '6px',
    width: '100%',
    maxWidth: '420px',
    marginTop: '6px',
  },
  quickActionCard: {
    justifyContent: 'flex-start',
    textAlign: 'left',
    fontWeight: 500,
    fontSize: '12px',
    padding: '6px 8px',
    height: 'auto',
    whiteSpace: 'normal',
    wordBreak: 'break-word',
  },
  quickActionsBar: {
    display: 'flex',
    flexWrap: 'wrap',
    alignItems: 'center',
    gap: '4px',
    padding: '6px 10px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    overflowX: 'hidden',
    flexShrink: 0,
    minWidth: 0,
  },
  quickActionPill: {
    fontSize: '11px',
    padding: '2px 7px',
    height: 'auto',
    minHeight: '24px',
    maxWidth: '100%',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  selectionBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '6px',
    padding: '4px 10px',
    backgroundColor: tokens.colorBrandBackground2,
    borderTop: `1px solid ${tokens.colorBrandStroke2}`,
    fontSize: '12px',
    flexShrink: 0,
    minWidth: 0,
  },
  selectionText: {
    flex: 1,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    color: tokens.colorNeutralForeground1,
  },
  approvalArea: {
    maxHeight: '45vh',
    overflowY: 'auto',
    padding: '8px 10px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
    minWidth: 0,
  },
  inputArea: {
    display: 'flex',
    alignItems: 'flex-end',
    gap: '6px',
    padding: '8px 10px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
    minWidth: 0,
  },
  input: {
    flex: 1,
    minWidth: 0,
  },
  textarea: {
    width: '100%',
    minHeight: '32px',
    maxHeight: '160px',
    resize: 'none',
    overflowY: 'auto',
  },
});

interface ChatPanelProps {
  host: HostContext;
  messages: ChatMessage[];
  isLoading: boolean;
  pendingApproval: string | null;
  pendingPreview?: Preview | null;
  /** Host of the currently-loaded conversation; null = no active conversation. */
  activeChatHost: HostKind | null;
  /** Running total cost for the active conversation. */
  cost: CallCost | undefined;
  /** Currently-selected provider id, used to label local-model costs as free. */
  providerId?: string;
  /** Label for the current isolated window/document project */
  documentLabel?: string;
  /** True if the active provider has a personal API key configured on this machine */
  hasApiKey?: boolean;
  onSaveQuickApiKey?: (providerId: string, apiKey: string, model: string) => void;
  onSend: (text: string) => void;
  onApprove: (approved: boolean) => void;
  onOpenSettings: () => void;
  onOpenHistory: () => void;
  onOpenToolsCatalog?: () => void;
  onNewChat: () => void;
  draftPrompt?: string;
  onClearDraftPrompt?: () => void;
}

export function ChatPanel({
  host, messages, isLoading, pendingApproval, pendingPreview, activeChatHost, cost, providerId,
  documentLabel, hasApiKey = true, onSaveQuickApiKey,
  onSend, onApprove, onOpenSettings, onOpenHistory, onOpenToolsCatalog, onNewChat,
  draftPrompt, onClearDraftPrompt,
}: ChatPanelProps) {
  const styles = useStyles();
  const { t, locale } = useTranslation();
  const totalBuiltInTools = host.kind === 'word'
    ? KITCHENSINK_WORD_TOOLS_META.length
    : 2 + listSkills(host.kind).length;
  const hostDisplay = t(
    host.kind === 'word' ? 'chat.hostWord' :
    host.kind === 'excel' ? 'chat.hostExcel' :
    'chat.hostPowerpoint',
  );
  const hostNoun = t(
    host.kind === 'word' ? 'chat.hostNounWord' :
    host.kind === 'excel' ? 'chat.hostNounExcel' :
    'chat.hostNounPowerpoint',
  );
  const [inputText, setInputText] = useState('');
  const [quickProvider, setQuickProvider] = useState<string>(providerId || 'openai');
  const [quickKeyInput, setQuickKeyInput] = useState('');
  const [selectedWordText, setSelectedWordText] = useState<string>(() => getCachedWordSelection());
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const quickActions = useMemo(
    () => QUICK_ACTIONS_BY_HOST[host.kind] ?? QUICK_ACTIONS_BY_HOST.word,
    [host.kind],
  );

  // Track Word selection in real time so when the user selects text ("bôi đen / bôi xanh")
  // and clicks into the chat input box, AutoOffice keeps & displays their selection!
  useEffect(() => {
    if (host.kind !== 'word') return;
    return initWordSelectionTracker(setSelectedWordText);
  }, [host.kind]);

  useEffect(() => {
    if (draftPrompt) {
      setInputText(draftPrompt);
      onClearDraftPrompt?.();
      setTimeout(() => {
        textareaRef.current?.focus();
      }, 50);
    }
  }, [draftPrompt, onClearDraftPrompt]);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, pendingApproval]);

  useLayoutEffect(() => {
    const el = textareaRef.current;
    if (!el) return;
    el.style.height = 'auto';
    el.style.height = `${Math.min(el.scrollHeight, 200)}px`;
  }, [inputText]);

  const handleSubmit = () => {
    if (!inputText.trim() || isLoading) return;
    onSend(inputText);
    setInputText('');
  };

  const handleQuickAction = (promptKey: TranslationKey) => {
    if (isLoading) return;
    onSend(t(promptKey));
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  };

  const handleSavePersonalKey = () => {
    if (!quickKeyInput.trim() || !onSaveQuickApiKey) return;
    const defaultModel = QUICK_DEFAULT_MODELS[quickProvider] || 'gpt-4o';
    onSaveQuickApiKey(quickProvider, quickKeyInput.trim(), defaultModel);
    setQuickKeyInput('');
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <div className={styles.headerTopRow}>
          <div className={styles.brand}>
            <img
              src={`${import.meta.env.BASE_URL}assets/icon-64.png`}
              alt=""
              className={styles.logo}
            />
            <Text className={styles.title}>AutoOffice</Text>
            <span className={styles.hostPill}>
              {host.displayName}
            </span>
          </div>
          <div className={styles.headerActions}>
            <Tooltip content={t('chat.historyTooltip')} relationship="label">
              <Button size="small" appearance="subtle" icon={<History24Regular />} onClick={onOpenHistory} disabled={isLoading} />
            </Tooltip>
            <Tooltip content={t('chat.newChatTooltip')} relationship="label">
              <Button size="small" appearance="subtle" icon={<Add24Regular />} onClick={onNewChat} disabled={isLoading} />
            </Tooltip>
            <Tooltip content={t('chat.settingsTooltip')} relationship="label">
              <Button size="small" appearance="subtle" icon={<Settings24Regular />} onClick={onOpenSettings} />
            </Tooltip>
          </div>
        </div>

        {(documentLabel || onOpenToolsCatalog || cost) && (
          <div className={styles.headerSubRow}>
            {documentLabel ? (
              <Tooltip
                content={
                  locale === 'vi'
                    ? `Phiên làm việc độc lập cho cửa sổ này (${documentLabel}). Mở nhiều cửa sổ Word sẽ tự động chạy đa luồng trên các phiên riêng biệt.`
                    : `Isolated session for this window (${documentLabel}).`
                }
                relationship="description"
              >
                <span className={styles.docSessionPill}>
                  📄 {documentLabel}
                </span>
              </Tooltip>
            ) : <span style={{ flex: 1 }} />}

            <div className={styles.headerMetaRight}>
              {onOpenToolsCatalog && (
                <Tooltip
                  content={locale === 'vi' ? 'Xem bộ 10 Công cụ MCP AutoOffice' : 'View 10 AutoOffice MCP Tools'}
                  relationship="label"
                >
                  <Button
                    appearance="subtle"
                    size="small"
                    style={{ padding: '2px 6px', minWidth: 'auto', fontSize: '11px', whiteSpace: 'nowrap' }}
                    onClick={onOpenToolsCatalog}
                  >
                    🛠️ {totalBuiltInTools} {locale === 'vi' ? 'công cụ' : 'tools'}
                  </Button>
                </Tooltip>
              )}
              <CostBadge cost={cost} providerId={providerId} />
            </div>
          </div>
        )}
      </div>

      {activeChatHost && activeChatHost !== host.kind && (
        <CrossHostBanner chatHost={activeChatHost} currentHost={host.kind} />
      )}

      <div className={styles.messageList}>
        {messages.length === 0 && !pendingApproval ? (
          <div className={styles.empty}>
            <Text size={400} weight="semibold">{t('chat.welcomeTitle')}</Text>
            <Text size={200}>{t('chat.welcomeMessage', { host: hostDisplay, noun: hostNoun })}</Text>

            {!hasApiKey && onSaveQuickApiKey && (
              <div className={styles.apiKeyCard} data-testid="personal-api-key-card">
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <Key20Regular style={{ flexShrink: 0 }} />
                  <Text weight="semibold" size={300}>
                    {locale === 'vi'
                      ? 'Nhập Khóa API Cá Nhân (Lưu riêng trên máy này)'
                      : 'Enter Personal API Key (Stored only on this device)'}
                  </Text>
                </div>
                <Text size={200}>
                  {locale === 'vi'
                    ? 'Mỗi người dùng sử dụng khóa API và lịch sử riêng biệt trên máy cá nhân. Khóa của bạn không bao giờ bị chia sẻ cho người khác.'
                    : 'Each user uses their own private API key and local history on their device.'}
                </Text>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                  <Select
                    size="small"
                    value={quickProvider}
                    onChange={(_, d) => setQuickProvider(d.value)}
                    style={{ minWidth: '110px', flex: '0 0 auto' }}
                  >
                    <option value="openai">OpenAI</option>
                    <option value="google">Google Gemini</option>
                    <option value="anthropic">Anthropic</option>
                    <option value="deepseek">DeepSeek</option>
                  </Select>
                  <Input
                    size="small"
                    type="password"
                    style={{ flex: '1 1 130px', minWidth: 0 }}
                    placeholder="sk-..."
                    value={quickKeyInput}
                    onChange={(_, d) => setQuickKeyInput(d.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') handleSavePersonalKey();
                    }}
                  />
                </div>
                <Button
                  appearance="primary"
                  size="small"
                  disabled={!quickKeyInput.trim()}
                  onClick={handleSavePersonalKey}
                >
                  {locale === 'vi' ? 'Lưu Khóa API Cá Nhân' : 'Save Personal API Key'}
                </Button>
              </div>
            )}

            <Text size={200}>
              {host.kind === 'word'
                ? t('chat.exampleWord')
                : host.kind === 'excel'
                  ? t('chat.exampleExcel')
                  : t('chat.examplePowerpoint')}
            </Text>

            <div className={styles.quickActionsGrid} data-testid="quick-actions-grid">
              {quickActions.map(action => (
                <Tooltip key={action.id} content={t(action.promptKey)} relationship="description">
                  <Button
                    appearance="outline"
                    size="small"
                    icon={<Sparkle16Regular />}
                    className={styles.quickActionCard}
                    disabled={isLoading}
                    onClick={() => handleQuickAction(action.promptKey)}
                  >
                    {t(action.labelKey)}
                  </Button>
                </Tooltip>
              ))}
            </div>
          </div>
        ) : (
          messages.map((msg, i) => (
            <MessageBubble key={i} message={msg} />
          ))
        )}
        <div ref={messagesEndRef} />
      </div>

      {pendingApproval && (
        <div className={styles.approvalArea}>
          {pendingPreview ? <WordChangePreview
            preview={pendingPreview} onApprove={() => onApprove(true)} onReject={() => onApprove(false)}
          /> : <CodeBlock
            code={pendingApproval}
            status="pending"
            onApprove={() => onApprove(true)}
            onReject={() => onApprove(false)}
          />}
        </div>
      )}

      {messages.length > 0 && (
        <div className={styles.quickActionsBar} data-testid="quick-actions-bar">
          {quickActions.map(action => (
            <Tooltip key={action.id} content={t(action.promptKey)} relationship="description">
              <Button
                appearance="subtle"
                size="small"
                icon={<Sparkle16Regular />}
                className={styles.quickActionPill}
                disabled={isLoading}
                onClick={() => handleQuickAction(action.promptKey)}
              >
                {t(action.labelKey)}
              </Button>
            </Tooltip>
          ))}
        </div>
      )}

      {host.kind === 'word' && selectedWordText && (
        <div className={styles.selectionBar} data-testid="word-selection-bar">
          <div className={styles.selectionText}>
            <strong>🎯 {locale === 'vi' ? 'Đang bôi chọn:' : 'Selected:'}</strong> &ldquo;{selectedWordText}&rdquo;
          </div>
          <Tooltip content={locale === 'vi' ? 'Bỏ ghim vùng chọn' : 'Clear selection'} relationship="label">
            <Button
              appearance="subtle"
              size="small"
              icon={<Dismiss16Regular />}
              style={{ minWidth: 'auto', padding: '2px' }}
              onClick={() => {
                setCachedWordSelection('');
                setSelectedWordText('');
              }}
            />
          </Tooltip>
        </div>
      )}

      <div className={styles.inputArea}>
        <Textarea
          className={styles.input}
          textarea={{ ref: textareaRef, className: styles.textarea }}
          placeholder={t('chat.inputPlaceholder', { noun: hostNoun })}
          value={inputText}
          onChange={(_, data) => setInputText(data.value)}
          onKeyDown={handleKeyDown}
          disabled={isLoading}
          rows={1}
        />
        <Button
          appearance="primary"
          icon={<Send24Regular />}
          aria-label={t('chat.sendButton')}
          onClick={handleSubmit}
          disabled={!inputText.trim() || isLoading}
        />
      </div>
    </div>
  );
}
