import React, { useState, useCallback, useRef, useEffect } from 'react';
import { makeStyles, tokens } from '@fluentui/react-components';
import type { ModelMessage } from 'ai';
import {
  getDocumentIdentity,
  getWindowActiveConversationId,
  setWindowActiveConversationId,
  type HostContext,
  type HostKind,
} from './host/context.ts';
import { ChatPanel } from './components/ChatPanel.tsx';
import { SettingsPanel } from './components/SettingsPanel.tsx';
import { HistoryPanel } from './components/HistoryPanel.tsx';
import { ToolsCatalogPanel } from './components/ToolsCatalogPanel.tsx';
import { runAgent, type ChatMessage, type OrchestratorCallbacks } from './agent/orchestrator.ts';
import { generateTitle } from './agent/title.ts';
import { Sandbox } from './executor/sandbox.ts';
import { formatError } from './agent/errors.ts';
import {
  loadSettings,
  saveSettings,
  hasActiveProviderApiKey,
  type AppSettings,
} from './store/settings.ts';
import {
  saveConversation,
  getConversation,
  listConversations,
  renameConversation,
  deleteConversation,
  mostRecentForDocument,
  CURRENT_VERSION,
  type Conversation,
  type ConversationSummary,
} from './store/history.ts';
import { translationService } from './i18n/index.ts';
import { sumCallCosts, emptyCallCost, isCallCostEmpty, type CallCost } from './agent/pricing.ts';
import type { WordChangePreview } from './agent/word-preview';
import { captureWordRecovery, openWordRecovery, type WordRecoveryCopy } from './agent/word-recovery';
import { WordRecoveryPanel } from './components/WordRecoveryPanel';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    backgroundColor: tokens.colorNeutralBackground1,
    overflow: 'hidden',
  },
});

const SAVE_DEBOUNCE_MS = 300;
const PLACEHOLDER_LEN = 40;

function placeholderTitle(firstUserMessage: string): string {
  const oneLine = firstUserMessage.replace(/\s+/g, ' ').trim();
  if (!oneLine) return translationService.t('history.newChatPlaceholder');
  return oneLine.length <= PLACEHOLDER_LEN ? oneLine : oneLine.slice(0, PLACEHOLDER_LEN);
}

interface AppProps {
  host: HostContext;
}

export function App({ host }: AppProps) {
  const styles = useStyles();
  const [docIdentity, setDocIdentity] = useState(() => getDocumentIdentity());
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [testBridgeState, setTestBridgeState] = useState('');
  const [showSettings, setShowSettings] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [showToolsCatalog, setShowToolsCatalog] = useState(false);
  const [draftPrompt, setDraftPrompt] = useState<string>('');
  const [settings, setSettings] = useState<AppSettings>(loadSettings);
  const [pendingApproval, setPendingApproval] = useState<string | null>(null);
  const [pendingPreview, setPendingPreview] = useState<WordChangePreview | null>(null);
  const [recoveryInfo, setRecoveryInfo] = useState<Pick<WordRecoveryCopy, 'documentLabel' | 'createdAt'> | null>(null);
  const [preparingRecovery, setPreparingRecovery] = useState(false);
  const [openingRecovery, setOpeningRecovery] = useState(false);
  const [recoveryError, setRecoveryError] = useState(false);
  const [activeConversationId, setActiveConversationId] = useState<string | null>(null);
  const [activeChatHost, setActiveChatHost] = useState<HostKind | null>(null);
  const [activeCost, setActiveCost] = useState<CallCost | undefined>(undefined);
  const [historySummaries, setHistorySummaries] = useState<ConversationSummary[]>(() => listConversations());

  const conversationHistory = useRef<ModelMessage[]>([]);
  const sandboxRef = useRef<Sandbox | null>(null);
  const approvalResolveRef = useRef<((approved: boolean) => void) | null>(null);
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const turnRunningRef = useRef(false);
  const recoveryRef = useRef<WordRecoveryCopy | null>(null);

  useEffect(() => () => {
    recoveryRef.current = null;
    approvalResolveRef.current?.(false);
    approvalResolveRef.current = null;
  }, []);

  // Hydrate ONLY the conversation belonging to THIS specific window/document session.
  // Opening a new Word window or a different document starts a fresh, isolated chat
  // so multiple Word windows never collide or inherit a closed file's context.
  useEffect(() => {
    const identity = getDocumentIdentity();
    setDocIdentity(identity);

    const windowConvId = getWindowActiveConversationId();
    const convFromWindow = windowConvId ? getConversation(windowConvId) : null;
    const recentDoc = !convFromWindow
      ? mostRecentForDocument(host.kind, identity.documentKey)
      : null;
    const conv = convFromWindow ?? (recentDoc ? getConversation(recentDoc.id) : null);

    if (!conv) {
      setMessages([]);
      conversationHistory.current = [];
      setActiveConversationId(null);
      setActiveChatHost(null);
      setActiveCost(undefined);
      return;
    }

    setMessages(conv.uiMessages);
    conversationHistory.current = conv.modelMessages;
    setActiveConversationId(conv.id);
    setWindowActiveConversationId(conv.id);
    setActiveChatHost(conv.host);
    setActiveCost(conv.cost);
  }, [host.kind]);

  // Sync history list & settings when another Word window updates localStorage concurrently
  useEffect(() => {
    const onStorageChange = (e: StorageEvent) => {
      if (!e.key) return;
      if (e.key.startsWith('autooffice_history_')) {
        setHistorySummaries(listConversations());
      } else if (e.key === 'autooffice_settings') {
        setSettings(loadSettings());
      }
    };
    window.addEventListener('storage', onStorageChange);
    return () => window.removeEventListener('storage', onStorageChange);
  }, []);

  useEffect(() => {
    const sandbox = new Sandbox(host.kind);
    sandbox.init();
    sandboxRef.current = sandbox;
    return () => sandbox.destroy();
  }, [host.kind]);

  const refreshSummaries = useCallback(() => {
    setHistorySummaries(listConversations());
  }, []);

  const persistDebounced = useCallback((conv: Conversation) => {
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      saveConversation(conv);
      refreshSummaries();
    }, SAVE_DEBOUNCE_MS);
  }, [refreshSummaries]);

  const persistImmediate = useCallback((conv: Conversation) => {
    if (saveTimerRef.current) {
      clearTimeout(saveTimerRef.current);
      saveTimerRef.current = null;
    }
    saveConversation(conv);
    refreshSummaries();
  }, [refreshSummaries]);

  const handleSettingsChange = useCallback((newSettings: AppSettings) => {
    setSettings(newSettings);
    saveSettings(newSettings);
  }, []);

  const handleSaveQuickApiKey = useCallback((providerId: string, apiKey: string, model: string) => {
    const updatedProviders = settings.providers.map(p =>
      p.id === providerId ? { ...p, apiKey: apiKey.trim() } : p
    );
    const nextSettings: AppSettings = {
      ...settings,
      selectedProviderId: providerId,
      selectedModel: model || settings.selectedModel || 'gpt-4o',
      providers: updatedProviders,
    };
    setSettings(nextSettings);
    saveSettings(nextSettings);
  }, [settings]);

  const handleApprove = useCallback((approved: boolean) => {
    if (approvalResolveRef.current) {
      approvalResolveRef.current(approved);
      approvalResolveRef.current = null;
      setPendingApproval(null);
      setPendingPreview(null);
    }
  }, []);

  const handleOpenRecovery = useCallback(async () => {
    if (turnRunningRef.current || !recoveryRef.current) return;
    turnRunningRef.current = true;
    setOpeningRecovery(true);
    setRecoveryError(false);
    try { await openWordRecovery(recoveryRef.current); }
    catch { setRecoveryError(true); }
    finally { turnRunningRef.current = false; setOpeningRecovery(false); }
  }, []);

  const cancelPendingSave = useCallback(() => {
    if (saveTimerRef.current) {
      clearTimeout(saveTimerRef.current);
      saveTimerRef.current = null;
    }
  }, []);

  const handleNewChat = useCallback(() => {
    if (isLoading) return;
    cancelPendingSave();
    setMessages([]);
    conversationHistory.current = [];
    setActiveConversationId(null);
    setWindowActiveConversationId(null);
    setActiveChatHost(null);
    setActiveCost(undefined);
  }, [isLoading, cancelPendingSave]);

  const handleLoadConversation = useCallback((id: string) => {
    if (isLoading) return;
    const conv = getConversation(id);
    if (!conv) return;
    cancelPendingSave();
    setMessages(conv.uiMessages);
    conversationHistory.current = conv.modelMessages;
    setActiveConversationId(conv.id);
    setWindowActiveConversationId(conv.id);
    setActiveChatHost(conv.host);
    setActiveCost(conv.cost);
    setShowHistory(false);
  }, [isLoading, cancelPendingSave]);

  const handleRename = useCallback((id: string, title: string) => {
    renameConversation(id, title);
    refreshSummaries();
  }, [refreshSummaries]);

  const handleDelete = useCallback((id: string) => {
    deleteConversation(id);
    if (id === activeConversationId) {
      cancelPendingSave();
      setMessages([]);
      conversationHistory.current = [];
      setActiveConversationId(null);
      setWindowActiveConversationId(null);
      setActiveChatHost(null);
      setActiveCost(undefined);
    }
    refreshSummaries();
  }, [activeConversationId, cancelPendingSave, refreshSummaries]);

  const handleSend = useCallback(async (text: string) => {
    if (!text.trim() || isLoading || turnRunningRef.current) return;
    turnRunningRef.current = true;

    const currentIdentity = getDocumentIdentity();
    setDocIdentity(currentIdentity);

    let convId = activeConversationId;
    let convHost: HostKind = activeChatHost ?? host.kind;
    let isFirstTurn = false;
    if (convId === null) {
      convId = crypto.randomUUID();
      convHost = host.kind;
      isFirstTurn = true;
      setActiveConversationId(convId);
      setWindowActiveConversationId(convId);
      setActiveChatHost(convHost);
    }

    setMessages(prev => [...prev, { role: 'user', content: text }]);
    setIsLoading(true);

    let turnCost: CallCost | null = null;
    const historyStartLength = conversationHistory.current.length;
    const turnErrors: unknown[] = [];
    const toolResults: unknown[] = [];
    let recoveryCaptured = false;

    const callbacks: OrchestratorCallbacks = {
      onMessage: (msg) => {
        if (msg.error) turnErrors.push(msg.error);
        setMessages(prev => [...prev, msg]);
      },
      onTurnCost: (cost) => { turnCost = cost; },
      onStreamToken: (token) => {
        setMessages(prev => {
          const copy = [...prev];
          const last = copy[copy.length - 1];
          if (last?.role === 'assistant' && !last.codeBlock && !last.toolActivity) {
            copy[copy.length - 1] = { ...last, content: last.content + token };
          }
          return copy;
        });
      },
      onUpsertCodeBlock: (toolCallId, patch) => {
        if (patch.status === 'success' || patch.status === 'error' || patch.status === 'rejected') {
          toolResults.push({ toolCallId, status: patch.status, result: patch.result });
        }
        setMessages(prev => {
          const idx = prev.findIndex(m => m.codeBlock?.toolCallId === toolCallId);
          if (idx === -1) {
            return [...prev, {
              role: 'assistant',
              content: '',
              codeBlock: {
                toolCallId,
                code: patch.code ?? '',
                status: patch.status ?? 'streaming',
                result: patch.result,
              },
            }];
          }
          const copy = [...prev];
          const existing = copy[idx];
          copy[idx] = {
            ...existing,
            codeBlock: { ...existing.codeBlock!, ...patch },
          };
          return copy;
        });
      },
      requestApproval: (code, preview) => {
        setPendingApproval(code);
        setPendingPreview(preview ?? null);
        return new Promise<boolean>((resolve) => {
          approvalResolveRef.current = resolve;
        });
      },
      beforeWordMutation: async () => {
        if (recoveryCaptured) return;
        setPreparingRecovery(true);
        setRecoveryError(false);
        try {
          const copy = await captureWordRecovery(currentIdentity.documentLabel);
          recoveryRef.current = copy;
          setRecoveryInfo({ documentLabel: copy.documentLabel, createdAt: copy.createdAt });
          recoveryCaptured = true;
        } catch {
          setRecoveryError(true);
          throw new Error(translationService.t('recovery.captureFailed'));
        } finally { setPreparingRecovery(false); }
      },
    };

    const placeholder = isFirstTurn ? (placeholderTitle(text) || translationService.t('history.newChatPlaceholder')) : '';

    try {
      const history = await runAgent(
        text,
        conversationHistory.current,
        settings,
        sandboxRef.current!,
        host.kind,
        callbacks,
      );
      conversationHistory.current = history;
    } catch (e) {
      const formatted = formatError(e, { phase: 'agent' });
      turnErrors.push(formatted);
      setMessages(prev => [...prev, { role: 'assistant', content: '', error: formatted }]);
    } finally {
      turnRunningRef.current = false;
      setIsLoading(false);
      setPendingApproval(null);
      setPendingPreview(null);
    }

    setMessages(currentMessages => {
      const now = Date.now();
      const existing = isFirstTurn ? null : getConversation(convId!);
      const accumulatedCost = turnCost && !isCallCostEmpty(turnCost)
        ? sumCallCosts([activeCost ?? emptyCallCost('estimated'), turnCost])
        : activeCost;
      const conv: Conversation = {
        id: convId!,
        v: CURRENT_VERSION,
        title: isFirstTurn ? placeholder : (existing?.title ?? translationService.t('history.newChatPlaceholder')),
        host: convHost,
        createdAt: existing?.createdAt ?? now,
        updatedAt: now,
        messageCount: currentMessages.length,
        uiMessages: currentMessages,
        modelMessages: conversationHistory.current,
        cost: accumulatedCost,
        totalUsd: accumulatedCost?.totalUsd,
        costSource: accumulatedCost?.source,
        documentKey: existing?.documentKey ?? currentIdentity.documentKey,
        documentLabel: existing?.documentLabel ?? currentIdentity.documentLabel,
      };
      setActiveCost(accumulatedCost);
      if (isFirstTurn) persistImmediate(conv);
      else persistDebounced(conv);
      return currentMessages;
    });

    if (isFirstTurn) {
      void generateTitle(conversationHistory.current, settings).then((newTitle) => {
        if (!newTitle) return;
        const current = getConversation(convId!);
        if (!current) return;
        if (current.title !== placeholder) return;
        renameConversation(convId!, newTitle);
        refreshSummaries();
      });
    }
    return { modelMessages: conversationHistory.current.slice(historyStartLength), errors: turnErrors, toolResults };
  }, [isLoading, settings, host, activeConversationId, activeChatHost, activeCost, persistDebounced, persistImmediate, refreshSummaries]);

  // One connection per task pane. Access current React callbacks without reconnecting mid-turn.
  const testAccessRef = useRef({
    busy: () => turnRunningRef.current,
    status: () => ({
      provider: settings.selectedProviderId, model: settings.selectedModel,
      hasApiKey: hasActiveProviderApiKey(settings), autoApprove: settings.autoApprove,
      busy: turnRunningRef.current, pendingApproval: Boolean(pendingApproval),
      previewWordChanges: settings.previewWordChanges !== false,
      recoveryAvailable: Boolean(recoveryRef.current),
    }),
    send: handleSend,
  });
  testAccessRef.current = {
    busy: () => turnRunningRef.current,
    status: () => ({
      provider: settings.selectedProviderId, model: settings.selectedModel,
      hasApiKey: hasActiveProviderApiKey(settings), autoApprove: settings.autoApprove,
      busy: turnRunningRef.current, pendingApproval: Boolean(pendingApproval),
      previewWordChanges: settings.previewWordChanges !== false,
      recoveryAvailable: Boolean(recoveryRef.current),
    }),
    send: handleSend,
  };
  useEffect(() => {
    if (!import.meta.env.DEV || !__AUTOOFFICE_TEST_BRIDGE__ || host.kind !== 'word') return;
    let disposed = false;
    let stop: (() => void) | undefined;
    void import('./testing/bridge.ts').then(({ startTestBridge }) => {
      if (!disposed) stop = startTestBridge({
        busy: () => testAccessRef.current.busy(),
        status: () => testAccessRef.current.status(),
        send: prompt => testAccessRef.current.send(prompt),
      }, setTestBridgeState);
    });
    return () => { disposed = true; stop?.(); };
  }, [host.kind]);

  if (showSettings) {
    return (
      <div className={styles.root}>
        <SettingsPanel
          settings={settings}
          onChange={handleSettingsChange}
          onClose={() => setShowSettings(false)}
          onHistoryCleared={() => {
            refreshSummaries();
            handleNewChat();
          }}
        />
      </div>
    );
  }

  if (showHistory) {
    return (
      <div className={styles.root}>
        <HistoryPanel
          conversations={historySummaries}
          currentHost={host.kind}
          activeId={activeConversationId}
          currentDocumentKey={docIdentity.documentKey}
          onSelect={handleLoadConversation}
          onRename={handleRename}
          onDelete={handleDelete}
          onClose={() => setShowHistory(false)}
        />
      </div>
    );
  }

  if (showToolsCatalog) {
    return (
      <div className={styles.root}>
        <ToolsCatalogPanel
          host={host.kind}
          mcpServers={settings.mcpServers}
          onClose={() => setShowToolsCatalog(false)}
          onRunSample={(prompt) => { void handleSend(prompt); }}
          onFillSample={(prompt) => { setDraftPrompt(prompt); }}
        />
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <ChatPanel
        host={host}
        messages={messages}
        isLoading={isLoading || openingRecovery}
        pendingApproval={pendingApproval}
        pendingPreview={pendingPreview}
        activeChatHost={activeChatHost}
        cost={activeCost}
        providerId={settings.selectedProviderId}
        documentLabel={docIdentity.documentLabel}
        hasApiKey={hasActiveProviderApiKey(settings)}
        onSaveQuickApiKey={handleSaveQuickApiKey}
        onSend={handleSend}
        onApprove={handleApprove}
        onOpenSettings={() => setShowSettings(true)}
        onOpenHistory={() => setShowHistory(true)}
        onOpenToolsCatalog={() => setShowToolsCatalog(true)}
        onNewChat={handleNewChat}
        draftPrompt={draftPrompt}
        onClearDraftPrompt={() => setDraftPrompt('')}
      />
      {host.kind === 'word' && <WordRecoveryPanel
        documentLabel={recoveryInfo?.documentLabel} createdAt={recoveryInfo?.createdAt}
        preparing={preparingRecovery} opening={openingRecovery} busy={isLoading || openingRecovery}
        error={recoveryError} onOpen={() => void handleOpenRecovery()}
      />}
      {import.meta.env.DEV && __AUTOOFFICE_TEST_BRIDGE__ && testBridgeState && (
        <div role="status" style={{ padding: '4px 10px', fontSize: 11, background: '#fff4ce' }}>
          {testBridgeState}
        </div>
      )}
    </div>
  );
}
