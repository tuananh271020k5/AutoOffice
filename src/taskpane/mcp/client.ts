import { experimental_createMCPClient as createMCPClient } from '@ai-sdk/mcp';
import type { McpServerConfig } from '../store/settings.ts';
import type { ToolSet } from 'ai';

function resolveUrl(url: string): string {
  if (import.meta.env.DEV && /^https?:\/\//.test(url)) {
    return `${window.location.origin}/api/mcp-proxy?target=${encodeURIComponent(url)}`;
  }
  return url;
}

export interface McpConnectFailure {
  serverName: string;
  url: string;
  error: unknown;
}

export interface McpToolsResult {
  tools: ToolSet;
  failures: McpConnectFailure[];
}

export interface McpToolInfo {
  name: string;
  description: string;
  serverName: string;
}

export async function inspectMcpServer(server: McpServerConfig): Promise<McpToolInfo[]> {
  const trimmedUrl = server.url.trim();
  if (!trimmedUrl) return [];
  const client = await createMCPClient({
    transport: {
      type: server.transport === 'sse' ? 'sse' : 'http',
      url: resolveUrl(trimmedUrl),
      fetch: (url: RequestInfo | URL, init?: RequestInit) => fetch(url, init),
    },
  });
  const tools = await client.tools();
  return Object.entries(tools).map(([name, def]) => ({
    name,
    description: (def as { description?: string })?.description || '',
    serverName: server.name,
  }));
}

export async function getMcpTools(servers: McpServerConfig[]): Promise<McpToolsResult> {
  const allTools: ToolSet = {};
  const failures: McpConnectFailure[] = [];

  const enabledServers = servers.filter(s => s.enabled && s.url && s.url.trim().length > 0);

  for (const server of enabledServers) {
    try {
      const client = await createMCPClient({
        transport: {
          type: server.transport === 'sse' ? 'sse' : 'http',
          url: resolveUrl(server.url.trim()),
          fetch: (url: RequestInfo | URL, init?: RequestInit) => fetch(url, init),
        },
      });
      const tools = await client.tools();
      Object.assign(allTools, tools);
    } catch (e) {
      failures.push({ serverName: server.name, url: server.url, error: e });
    }
  }

  return { tools: allTools, failures };
}
