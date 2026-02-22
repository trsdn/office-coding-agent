/**
 * mcpClient.mjs — connects to MCP servers and returns their tools as Copilot SDK tools.
 *
 * Each MCP tool is mapped to a Copilot SDK tool with a handler that calls
 * the MCP server directly on the proxy server (no browser round-trip needed).
 *
 * Supports three transport types:
 *   - http:  StreamableHTTPClientTransport (default for url-based configs)
 *   - sse:   SSEClientTransport
 *   - stdio: StdioClientTransport (spawns a subprocess, e.g. WorkIQ)
 */

import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { SSEClientTransport } from '@modelcontextprotocol/sdk/client/sse.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

/**
 * Connect to a list of MCP servers and return Copilot SDK-compatible tools.
 *
 * @param {Array<{ name: string, url?: string, transport: 'http'|'sse'|'stdio', headers?: Record<string, string>, command?: string, args?: string[] }>} mcpServers
 * @returns {Promise<{ tools: Array<{ name: string, description: string, parameters: object, handler: Function }>, clients: Array<{ name: string, client: Client }> }>}
 */
export async function loadMcpTools(mcpServers) {
  const tools = [];
  const clients = [];

  for (const server of mcpServers) {
    try {
      console.log(`[mcp] Connecting to MCP server '${server.name}' (${server.transport})...`);

      const client = new Client({ name: 'office-coding-agent', version: '1.0.0' });

      let transport;
      if (server.transport === 'stdio') {
        if (!server.command) {
          console.warn(`[mcp] Server '${server.name}': stdio transport requires a command, skipping`);
          continue;
        }
        transport = new StdioClientTransport({
          command: server.command,
          args: server.args || [],
        });
      } else if (server.transport === 'sse') {
        transport = new SSEClientTransport(new URL(server.url), {
          requestInit: server.headers ? { headers: server.headers } : undefined,
        });
      } else {
        transport = new StreamableHTTPClientTransport(new URL(server.url), {
          requestInit: server.headers ? { headers: server.headers } : undefined,
        });
      }

      await client.connect(transport);
      clients.push({ name: server.name, client });

      const { tools: mcpTools } = await client.listTools();
      console.log(
        `[mcp] Server '${server.name}': ${mcpTools.length} tool(s) — ${mcpTools.map(t => t.name).join(', ')}`
      );

      for (const tool of mcpTools) {
        tools.push({
          name: tool.name,
          description: tool.description || '',
          parameters: tool.inputSchema || {},
          handler: async (args) => {
            try {
              const result = await client.callTool({ name: tool.name, arguments: args });
              const text =
                result.content
                  ?.map(c => (c.type === 'text' ? c.text : JSON.stringify(c)))
                  .join('\n') ?? JSON.stringify(result);
              return { textResultForLlm: text, resultType: 'success', toolTelemetry: {} };
            } catch (err) {
              const message = err instanceof Error ? err.message : String(err);
              console.error(`[mcp] Tool '${tool.name}' error:`, message);
              return {
                textResultForLlm: message,
                resultType: 'failure',
                error: message,
                toolTelemetry: {},
              };
            }
          },
        });
      }
    } catch (err) {
      console.warn(
        `[mcp] Failed to connect to '${server.name}' (${server.url}):`,
        err.message || err
      );
    }
  }

  return { tools, clients };
}

/**
 * Disconnect all MCP clients.
 * @param {Array<{ name: string, client: Client }>} clients
 */
export async function closeMcpClients(clients) {
  for (const { name, client } of clients) {
    try {
      await client.close();
    } catch {
      console.warn(`[mcp] Failed to close client '${name}'`);
    }
  }
}
