import React, { useEffect, useState } from "react";
import {
  makeStyles,
  tokens,
  Button,
  Badge,
  Text,
  Divider,
  Spinner,
} from "@fluentui/react-components";
import {
  ArrowLeft24Regular,
  Sparkle16Regular,
  Edit16Regular,
  ArrowClockwise24Regular,
} from "@fluentui/react-icons";
import type { HostKind } from "../host/context.ts";
import { listSkills } from "../skills/index.ts";
import type { McpServerConfig } from "../store/settings.ts";
import { inspectMcpServer, type McpToolInfo } from "../mcp/client.ts";
import { useTranslation } from "../i18n/index.ts";
import { KITCHENSINK_WORD_TOOLS_META } from "../agent/kitchensink-word-tools.ts";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    width: "100%",
    minWidth: 0,
    overflow: "hidden",
    boxSizing: "border-box",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "8px 10px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
    minWidth: 0,
  },
  headerTitle: {
    flex: 1,
    minWidth: 0,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  content: {
    flex: 1,
    overflowY: "auto",
    overflowX: "hidden",
    padding: "10px",
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    minWidth: 0,
    boxSizing: "border-box",
  },
  summaryBox: {
    padding: "10px 12px",
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: "8px",
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    minWidth: 0,
    wordBreak: "break-word",
  },
  card: {
    padding: "10px 12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: "8px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    minWidth: 0,
    boxSizing: "border-box",
  },
  cardTopBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    flexWrap: "wrap",
    gap: "6px",
    minWidth: 0,
  },
  toolCodePill: {
    display: "inline-flex",
    alignItems: "center",
    padding: "2px 8px",
    borderRadius: "999px",
    fontSize: "11px",
    fontWeight: 600,
    lineHeight: "16px",
    whiteSpace: "nowrap",
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    border: `1px solid ${tokens.colorBrandStroke2}`,
    maxWidth: "100%",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  cardTitle: {
    display: "block",
    lineHeight: "1.35",
    wordBreak: "break-word",
  },
  actionRow: {
    display: "flex",
    alignItems: "center",
    flexWrap: "wrap",
    gap: "4px",
    flexShrink: 0,
  },
  opsRow: {
    display: "flex",
    flexWrap: "wrap",
    alignItems: "center",
    gap: "4px",
    marginTop: "2px",
    minWidth: 0,
  },
  opBadge: {
    display: "inline-block",
    padding: "1px 6px",
    borderRadius: "4px",
    fontSize: "11px",
    lineHeight: "16px",
    whiteSpace: "nowrap",
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground2,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  samplePreview: {
    padding: "6px 8px",
    borderRadius: "6px",
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px dashed ${tokens.colorNeutralStroke2}`,
    fontSize: "11px",
    lineHeight: "1.4",
    color: tokens.colorNeutralForeground2,
    wordBreak: "break-word",
  },
  sectionHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    flexWrap: "wrap",
    gap: "6px",
    marginTop: "2px",
    minWidth: 0,
  },
  brandPill: {
    display: "inline-block",
    padding: "2px 8px",
    borderRadius: "999px",
    fontSize: "11px",
    fontWeight: 600,
    lineHeight: "16px",
    whiteSpace: "nowrap",
    flexShrink: 0,
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
});

interface ToolsCatalogPanelProps {
  host: HostKind;
  mcpServers: McpServerConfig[];
  onClose: () => void;
  onRunSample: (prompt: string) => void;
  onFillSample?: (prompt: string) => void;
}

export function ToolsCatalogPanel({
  host,
  mcpServers,
  onClose,
  onRunSample,
  onFillSample,
}: ToolsCatalogPanelProps) {
  const styles = useStyles();
  const { locale } = useTranslation();
  const isVi = locale === "vi";
  const skills = listSkills(host);

  const [mcpTools, setMcpTools] = useState<McpToolInfo[]>([]);
  const [mcpErrors, setMcpErrors] = useState<
    Array<{ server: string; message: string }>
  >([]);
  const [loadingMcp, setLoadingMcp] = useState(false);

  const enabledMcpServers = mcpServers.filter(
    (s) => s.enabled && s.url && s.url.trim().length > 0,
  );

  const loadMcpTools = async () => {
    if (enabledMcpServers.length === 0) {
      setMcpTools([]);
      setMcpErrors([]);
      return;
    }
    setLoadingMcp(true);
    const found: McpToolInfo[] = [];
    const errs: Array<{ server: string; message: string }> = [];
    for (const s of enabledMcpServers) {
      try {
        const list = await inspectMcpServer(s);
        found.push(...list);
      } catch (e) {
        errs.push({
          server: s.name,
          message: e instanceof Error ? e.message : String(e),
        });
      }
    }
    setMcpTools(found);
    setMcpErrors(errs);
    setLoadingMcp(false);
  };

  useEffect(() => {
    void loadMcpTools();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mcpServers]);

  const totalActionCount = KITCHENSINK_WORD_TOOLS_META.reduce(
    (sum, item) => sum + item.ops.length,
    0,
  );
  const totalCount =
    host === "word"
      ? KITCHENSINK_WORD_TOOLS_META.length + mcpTools.length
      : 2 + skills.length + mcpTools.length;

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<ArrowLeft24Regular />}
          onClick={onClose}
          style={{ flexShrink: 0 }}
        />
        <Text weight="semibold" size={300} className={styles.headerTitle}>
          {isVi
            ? `Bộ 10 Công cụ MCP AutoOffice (${totalCount})`
            : `AutoOffice MCP Tools (${totalCount})`}
        </Text>
      </div>

      <div className={styles.content}>
        <div className={styles.summaryBox}>
          <Text weight="semibold" size={300}>
            {host === "word"
              ? isVi
                ? `Đang kích hoạt: 10 Công cụ MCP AutoOffice (${totalActionCount} Actions/Ops chuyên sâu)`
                : `Active: 10 AutoOffice MCP Tools (${totalActionCount} Operations)`
              : isVi
                ? `Đang kích hoạt: 2 Công cụ lõi + ${skills.length} Module ${host.toUpperCase()}`
                : `Active: 2 Core Tools + ${skills.length} ${host.toUpperCase()} Modules`}
          </Text>
          <Text size={200}>
            {isVi
              ? "Đã chuẩn hóa về bộ 10 công cụ MCP chuyên sâu của AutoOffice. Hỗ trợ nhận diện trực tiếp phần văn bản đang bôi đen / bôi xanh / bôi đậm và tự động đính kèm Quốc hiệu - Tiêu ngữ cho mọi đơn từ hành chính."
              : "Streamlined to the 10 AutoOffice MCP tools with live selection, bold/highlight targeting, and automatic Vietnamese administrative header formatting."}
          </Text>
        </div>

        {host === "word" ? (
          <>
            <div className={styles.sectionHeader}>
              <Text weight="semibold" size={300}>
                {isVi
                  ? `10 Công cụ MCP AutoOffice (${KITCHENSINK_WORD_TOOLS_META.length})`
                  : `10 AutoOffice MCP Tools (${KITCHENSINK_WORD_TOOLS_META.length})`}
              </Text>
              <span className={styles.brandPill}>
                AutoOffice MCP
              </span>
            </div>

            {KITCHENSINK_WORD_TOOLS_META.map((item, idx) => {
              const title = isVi ? item.titleVi : item.titleEn;
              const desc = isVi ? item.descVi : item.descEn;
              const sample = isVi ? item.samplePromptVi : item.samplePromptEn;
              return (
                <div key={item.name} className={styles.card}>
                  <div className={styles.cardTopBar}>
                    <span className={styles.toolCodePill}>
                      🛠️ {isVi ? `Nhóm ${idx + 1} (${item.ops.length} tác vụ)` : `Group ${idx + 1} (${item.ops.length} actions)`}
                    </span>
                    <div className={styles.actionRow}>
                      {onFillSample && (
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<Edit16Regular />}
                          onClick={() => {
                            onClose();
                            onFillSample(sample);
                          }}
                        >
                          {isVi ? "Sửa lệnh" : "Edit"}
                        </Button>
                      )}
                      <Button
                        size="small"
                        appearance="primary"
                        icon={<Sparkle16Regular />}
                        onClick={() => {
                          onClose();
                          onRunSample(sample);
                        }}
                      >
                        {isVi ? "Chạy thử" : "Run"}
                      </Button>
                    </div>
                  </div>
                  <Text weight="semibold" size={300} className={styles.cardTitle}>
                    {title}
                  </Text>
                  <Text size={200}>{desc}</Text>
                  <div className={styles.samplePreview}>
                    💬 {isVi ? "Mẫu thử:" : "Sample:"} &ldquo;{sample}&rdquo;
                  </div>
                </div>
              );
            })}
          </>
        ) : (
          skills.map((skillName) => (
            <div key={skillName} className={styles.card}>
              <div className={styles.cardTopBar}>
                <Text weight="semibold" size={300}>
                  {skillName}
                </Text>
                <Button
                  size="small"
                  appearance="primary"
                  icon={<Sparkle16Regular />}
                  onClick={() => {
                    onClose();
                    onRunSample(`Use ${skillName} on the current document.`);
                  }}
                >
                  {isVi ? "Chạy thử" : "Run"}
                </Button>
              </div>
            </div>
          ))
        )}

        {enabledMcpServers.length > 0 && (
          <>
            <Divider />
            <div className={styles.sectionHeader}>
              <Text weight="semibold" size={300}>
                {isVi
                  ? `Máy chủ MCP bên ngoài (${mcpTools.length})`
                  : `External MCP Server Tools (${mcpTools.length})`}
              </Text>
              <Button
                size="small"
                appearance="subtle"
                icon={<ArrowClockwise24Regular />}
                onClick={() => {
                  void loadMcpTools();
                }}
                disabled={loadingMcp}
              >
                {isVi ? "Làm mới" : "Refresh"}
              </Button>
            </div>
            {loadingMcp && <Spinner size="extra-tiny" />}
            {mcpErrors.map((err, i) => (
              <Text
                key={i}
                size={200}
                style={{ color: tokens.colorPaletteRedForeground1 }}
              >
                ✗ {err.server}: {err.message}
              </Text>
            ))}
          </>
        )}
      </div>
    </div>
  );
}
