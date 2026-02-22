/** Transport type for an MCP server */
export type McpTransportType = 'http' | 'sse' | 'stdio';

/** A configured MCP server imported from a mcp.json file */
export interface McpServerConfig {
  /** Display name (used as identifier) */
  name: string;
  /** Optional description shown in the UI */
  description?: string;
  /** MCP server endpoint URL (required for http/sse) */
  url?: string;
  /** Transport protocol */
  transport: McpTransportType;
  /** Optional HTTP headers (e.g. Authorization) — http/sse only */
  headers?: Record<string, string>;
  /** Command to spawn (stdio only) */
  command?: string;
  /** Arguments for the command (stdio only) */
  args?: string[];
}
