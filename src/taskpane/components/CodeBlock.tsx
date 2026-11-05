import React, { useMemo } from 'react';
import { makeStyles, tokens, Button, Text } from '@fluentui/react-components';
import {
  DismissCircle24Regular,
  Play24Regular,
} from '@fluentui/react-icons';
import { useTranslation } from '../i18n/index.ts';

const useStyles = makeStyles({
  container: {
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: '8px',
    overflow: 'hidden',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    maxWidth: '100%',
    minWidth: 0,
    boxSizing: 'border-box',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    flexWrap: 'wrap',
    gap: '6px',
    padding: '6px 10px',
    backgroundColor: tokens.colorNeutralBackground4,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    minWidth: 0,
  },
  headerTitle: {
    flex: '1 1 0%',
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  statusPill: {
    display: 'inline-block',
    padding: '2px 8px',
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
  codeArea: {
    padding: '10px',
    overflow: 'auto',
    maxHeight: '180px',
    minWidth: 0,
  },
  code: {
    fontFamily: 'Consolas, "Courier New", monospace',
    fontSize: '11.5px',
    lineHeight: '1.45',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    margin: 0,
    color: tokens.colorNeutralForeground1,
  },
  caret: {
    display: 'inline-block',
    width: '7px',
    height: '1em',
    marginLeft: '1px',
    verticalAlign: 'text-bottom',
    backgroundColor: tokens.colorNeutralForeground2,
    animationName: {
      '0%, 49%': { opacity: 1 },
      '50%, 100%': { opacity: 0 },
    },
    animationDuration: '1s',
    animationIterationCount: 'infinite',
  },
  emptyHint: {
    color: tokens.colorNeutralForeground3,
    fontStyle: 'italic',
  },
  actions: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '8px',
    padding: '8px 10px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  details: {
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  detailsError: {
    backgroundColor: tokens.colorPaletteRedBackground1,
  },
  summary: {
    padding: '6px 10px',
    cursor: 'pointer',
    fontSize: '12px',
    fontWeight: 600,
    color: tokens.colorNeutralForeground2,
    userSelect: 'none',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground4Hover,
    },
  },
  summaryError: {
    color: tokens.colorPaletteRedForeground1,
    '&:hover': {
      backgroundColor: tokens.colorPaletteRedBackground2,
    },
  },
  resultBody: {
    padding: '8px 10px 10px 10px',
    fontSize: '11.5px',
    fontFamily: 'Consolas, "Courier New", monospace',
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    maxHeight: '200px',
    overflow: 'auto',
    color: tokens.colorNeutralForeground1,
  },
  resultBodyError: {
    color: tokens.colorPaletteRedForeground1,
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

/**
 * Formats verbose KitchenSink4Word runtime scripts into a clean, readable MCP call preview
 * so the 1,500-line helper body doesn't flood narrow taskpanes.
 */
function formatDisplayCode(rawCode: string): { headerLabel: string; formattedCode: string } {
  const trimmed = rawCode.trim();
  if (trimmed.startsWith('// [KitchenSink4Word]') || trimmed.startsWith('// [AutoOffice]')) {
    const firstNewline = trimmed.indexOf('\n');
    const firstLine = firstNewline !== -1 ? trimmed.slice(0, firstNewline).trim() : trimmed;
    const match = /^\/\/\s*\[(?:KitchenSink4Word|AutoOffice)\]\s*([a-z0-9_]+)\.([a-z0-9_]+)\((.*)\)$/i.exec(firstLine);
    if (match) {
      const [, toolName, op, rawJson] = match;
      try {
        const parsed = JSON.parse(rawJson);
        return {
          headerLabel: `🛠️ ${toolName}.${op}`,
          formattedCode: `// AutoOffice MCP Tool\nawait ${toolName}("${op}", ${JSON.stringify(parsed, null, 2)});`,
        };
      } catch {
        return {
          headerLabel: `🛠️ ${toolName}.${op}`,
          formattedCode: firstLine,
        };
      }
    }
  }
  return { headerLabel: 'office.js', formattedCode: rawCode };
}

export function CodeBlock({ code, status, result, onApprove, onReject }: CodeBlockProps) {
  const styles = useStyles();
  const { t } = useTranslation();
  const STATUS_LABELS = {
    streaming: t('code.statusStreaming'),
    pending:   t('code.statusPending'),
    rejected:  t('code.statusRejected'),
    running:   t('code.statusRunning'),
    success:   t('code.statusSuccess'),
    error:     t('code.statusError'),
  } as const;
  const isError = status === 'error';
  const showResult = (status === 'success' || status === 'error') && !!result;

  const { headerLabel, formattedCode } = useMemo(() => formatDisplayCode(code), [code]);

  const statusPillClass =
    status === 'success' ? styles.statusSuccess :
    status === 'error' ? styles.statusDanger :
    status === 'rejected' ? styles.statusWarning :
    styles.statusInformative;

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Text size={200} weight="semibold" className={styles.headerTitle}>
          {headerLabel}
        </Text>
        <span className={`${styles.statusPill} ${statusPillClass}`}>
          {STATUS_LABELS[status]}
        </span>
      </div>

      <div className={styles.codeArea}>
        <pre className={styles.code}>
          {code === '' && status === 'streaming'
            ? <span className={styles.emptyHint}>{t('code.generatingCode')}</span>
            : formattedCode}
          {status === 'streaming' && <span className={styles.caret} aria-hidden />}
        </pre>
      </div>

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

      {showResult && (
        <details className={`${styles.details} ${isError ? styles.detailsError : ''}`} open={isError}>
          <summary className={`${styles.summary} ${isError ? styles.summaryError : ''}`}>
            {isError ? t('code.errorDetails') : t('code.result')}
          </summary>
          <div className={`${styles.resultBody} ${isError ? styles.resultBodyError : ''}`}>
            {result}
          </div>
        </details>
      )}
    </div>
  );
}
