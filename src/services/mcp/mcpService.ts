import type { McpServerConfig, McpTransportType } from '@/types';

/** Shape accepted from both Claude Desktop and VS Code mcp.json formats */
interface RawMcpEntry {
  url?: string;
  command?: string;
  args?: string[];
  type?: string;
  transport?: string;
  headers?: Record<string, string>;
  description?: string;
}

/**
 * Parse a `mcp.json` File into a normalized McpServerConfig array.
 *
 * Supports:
 *   - Claude Desktop format: `{ mcpServers: { name: { url, type, headers } } }`
 *   - VS Code format:        `{ servers:    { name: { url, type, headers } } }`
 *
 * Supports HTTP, SSE, and stdio transports. Stdio entries use `command` + `args`
 * and are executed server-side by the proxy.
 */
export async function parseMcpJsonFile(file: File): Promise<McpServerConfig[]> {
  if (!file.name.endsWith('.json')) {
    throw new Error('File must be a .json file.');
  }

  const text = await file.text();
  let parsed: unknown;
  try {
    parsed = JSON.parse(text) as unknown;
  } catch {
    throw new Error('Invalid JSON: could not parse mcp.json.');
  }

  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('mcp.json must be a JSON object.');
  }

  // Support both { mcpServers: {...} } and { servers: {...} }
  const raw = parsed as Record<string, unknown>;
  const serversMap =
    (raw.mcpServers as Record<string, RawMcpEntry> | undefined) ??
    (raw.servers as Record<string, RawMcpEntry> | undefined);

  if (!serversMap || typeof serversMap !== 'object' || Array.isArray(serversMap)) {
    throw new Error('mcp.json must contain a "mcpServers" or "servers" object.');
  }

  const configs: McpServerConfig[] = [];

  for (const [name, entry] of Object.entries(serversMap)) {
    if (typeof entry !== 'object' || entry === null) continue;

    const rawTransport = (entry.type ?? entry.transport ?? '').toLowerCase();
    const url = entry.url;
    const command = entry.command;

    // stdio entry: has command, no url
    if (typeof command === 'string' && command) {
      configs.push({
        name,
        description: typeof entry.description === 'string' ? entry.description : undefined,
        transport: 'stdio',
        command,
        args: Array.isArray(entry.args) ? entry.args.map(String) : [],
      });
      continue;
    }

    // http/sse entry: has url and valid transport
    if (typeof url !== 'string' || !url) continue;
    if (rawTransport && rawTransport !== 'http' && rawTransport !== 'sse') continue; // skip unknown (e.g. grpc)
    const transport = rawTransport === 'sse' ? 'sse' : 'http';

    configs.push({
      name,
      description: typeof entry.description === 'string' ? entry.description : undefined,
      url,
      transport: transport as McpTransportType,
      headers: entry.headers,
    });
  }

  if (configs.length === 0) {
    throw new Error(
      'No valid MCP servers found in mcp.json. ' +
        'Each entry needs either a "url" (http/sse) or a "command" (stdio).'
    );
  }

  return configs;
}

/** Return only the servers that are currently active (null = all active). */
export function resolveActiveMcpServers(
  servers: McpServerConfig[],
  activeNames: string[] | null
): McpServerConfig[] {
  if (activeNames === null) return servers;
  return servers.filter(s => activeNames.includes(s.name));
}
