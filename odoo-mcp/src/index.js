#!/usr/bin/env node
// Odoo MCP connector — stdio MCP server exposing an Odoo instance to Claude.
//
// Configuration comes from environment variables (see .env.example):
//   ODOO_URL       Base URL of the Odoo instance
//   ODOO_DB        Database name
//   ODOO_USERNAME  Login (email)
//   ODOO_API_KEY   API key (Settings ▸ Account Security ▸ New API Key) or password
//   ODOO_TIMEOUT   Optional request timeout in ms (default 30000)
//
// Secrets are never hard-coded — they must be supplied via the environment,
// typically through the Claude Desktop / Claude Code MCP server config.

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { OdooClient } from "./odoo-client.js";
import { buildTools } from "./tools.js";

function loadConfig() {
  const {
    ODOO_URL,
    ODOO_DB,
    ODOO_USERNAME,
    ODOO_API_KEY,
    ODOO_PASSWORD,
    ODOO_TIMEOUT,
  } = process.env;

  const missing = [];
  if (!ODOO_URL) missing.push("ODOO_URL");
  if (!ODOO_DB) missing.push("ODOO_DB");
  if (!ODOO_USERNAME) missing.push("ODOO_USERNAME");
  if (!ODOO_API_KEY && !ODOO_PASSWORD) missing.push("ODOO_API_KEY");

  if (missing.length) {
    process.stderr.write(
      `[odoo-mcp] Missing required environment variables: ${missing.join(", ")}\n` +
        `See .env.example for the full list.\n`
    );
    process.exit(1);
  }

  return {
    url: ODOO_URL,
    db: ODOO_DB,
    username: ODOO_USERNAME,
    apiKey: ODOO_API_KEY || ODOO_PASSWORD,
    timeout: ODOO_TIMEOUT ? Number(ODOO_TIMEOUT) : 30000,
  };
}

async function main() {
  const config = loadConfig();
  const client = new OdooClient(config);
  const { tools, call } = buildTools(client);

  const server = new Server(
    { name: "odoo-mcp", version: "1.0.0" },
    { capabilities: { tools: {} } }
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;
    return call(name, args);
  });

  const transport = new StdioServerTransport();
  await server.connect(transport);

  // Logs go to stderr so they never corrupt the stdio JSON-RPC stream.
  process.stderr.write(
    `[odoo-mcp] connected — ${tools.length} tools available for ${config.url} (db: ${config.db})\n`
  );
}

main().catch((err) => {
  process.stderr.write(`[odoo-mcp] fatal: ${err.stack || err}\n`);
  process.exit(1);
});
