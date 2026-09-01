#!/usr/bin/env node
/*
 * Odoo MCP connector — single-file, zero-dependency edition.
 *
 * Connects Claude to an Odoo instance via Odoo's external JSON-RPC API.
 * Needs only Node.js 18+ (uses the built-in fetch). No npm install required.
 *
 * Two ways to run it:
 *   1) As a Claude connector (default): Claude launches it and talks over stdio.
 *   2) Discovery / health check:  node odoo-mcp.js --discover
 *      (prints server version, database list, and confirms your login)
 *
 * Configuration comes from environment variables (never hard-code secrets):
 *   ODOO_URL       e.g. https://qacoatwork.com
 *   ODOO_DB        database name (e.g. bakertilly)
 *   ODOO_USERNAME  login (email)
 *   ODOO_API_KEY   API key (or password)
 *   ODOO_TIMEOUT   optional request timeout in ms (default 30000)
 */

"use strict";

// ---------------------------------------------------------------------------
// Odoo JSON-RPC client
// ---------------------------------------------------------------------------

function cfg() {
  const url = (process.env.ODOO_URL || "").replace(/\/+$/, "");
  const db = process.env.ODOO_DB;
  const username = process.env.ODOO_USERNAME;
  const apiKey = process.env.ODOO_API_KEY || process.env.ODOO_PASSWORD;
  const timeout = process.env.ODOO_TIMEOUT ? Number(process.env.ODOO_TIMEOUT) : 30000;
  return { url, db, username, apiKey, timeout };
}

let _uid = null;

async function jsonrpc(url, timeout, service, method, args) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeout);
  let resp;
  try {
    resp = await fetch(`${url}/jsonrpc`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        jsonrpc: "2.0",
        method: "call",
        params: { service, method, args },
        id: Math.floor(Math.random() * 1e9),
      }),
      signal: controller.signal,
    });
  } catch (err) {
    if (err.name === "AbortError") throw new Error(`Request to Odoo timed out after ${timeout}ms`);
    throw new Error(`Network error contacting Odoo at ${url}: ${err.message}`);
  } finally {
    clearTimeout(t);
  }
  if (!resp.ok) {
    const text = await resp.text().catch(() => "");
    throw new Error(`HTTP ${resp.status} from Odoo: ${text.slice(0, 500)}`);
  }
  const json = await resp.json();
  if (json.error) {
    const data = json.error.data || {};
    throw new Error(data.message || json.error.message || "Unknown Odoo error");
  }
  return json.result;
}

async function version() {
  const { url, timeout } = cfg();
  return jsonrpc(url, timeout, "common", "version", []);
}

async function authenticate() {
  if (_uid) return _uid;
  const { url, db, username, apiKey, timeout } = cfg();
  if (!url) throw new Error("Missing ODOO_URL");
  if (!db) throw new Error("Missing ODOO_DB");
  if (!username) throw new Error("Missing ODOO_USERNAME");
  if (!apiKey) throw new Error("Missing ODOO_API_KEY");
  const uid = await jsonrpc(url, timeout, "common", "authenticate", [db, username, apiKey, {}]);
  if (!uid) throw new Error("Authentication failed — check ODOO_DB, ODOO_USERNAME, and ODOO_API_KEY.");
  _uid = uid;
  return uid;
}

async function execute(model, method, args = [], kwargs = {}) {
  const { url, db, apiKey, timeout } = cfg();
  const uid = await authenticate();
  return jsonrpc(url, timeout, "object", "execute_kw", [db, uid, apiKey, model, method, args, kwargs]);
}

// ---------------------------------------------------------------------------
// MCP tools
// ---------------------------------------------------------------------------

const TOOLS = [
  {
    name: "odoo_search_read",
    description:
      "Search Odoo records and return chosen fields. Use an Odoo domain (list of triples, " +
      'e.g. [["customer_rank",">",0]]). Empty domain [] returns all.',
    inputSchema: {
      type: "object",
      properties: {
        model: { type: "string", description: "Model, e.g. res.partner, sale.order, account.move" },
        domain: { type: "array", default: [] },
        fields: { type: "array", items: { type: "string" } },
        limit: { type: "integer", default: 50 },
        offset: { type: "integer", default: 0 },
        order: { type: "string" },
      },
      required: ["model"],
    },
    run: (a) => {
      const kwargs = {};
      if (a.fields) kwargs.fields = a.fields;
      kwargs.limit = a.limit ?? 50;
      kwargs.offset = a.offset ?? 0;
      if (a.order) kwargs.order = a.order;
      return execute(a.model, "search_read", [a.domain ?? []], kwargs);
    },
  },
  {
    name: "odoo_search_count",
    description: "Count Odoo records matching a domain.",
    inputSchema: {
      type: "object",
      properties: { model: { type: "string" }, domain: { type: "array", default: [] } },
      required: ["model"],
    },
    run: (a) => execute(a.model, "search_count", [a.domain ?? []]),
  },
  {
    name: "odoo_read",
    description: "Read specific Odoo records by their IDs.",
    inputSchema: {
      type: "object",
      properties: {
        model: { type: "string" },
        ids: { type: "array", items: { type: "integer" } },
        fields: { type: "array", items: { type: "string" } },
      },
      required: ["model", "ids"],
    },
    run: (a) => execute(a.model, "read", [a.ids], a.fields ? { fields: a.fields } : {}),
  },
  {
    name: "odoo_create",
    description: "Create a new Odoo record. Returns the new record ID.",
    inputSchema: {
      type: "object",
      properties: { model: { type: "string" }, values: { type: "object" } },
      required: ["model", "values"],
    },
    run: (a) => execute(a.model, "create", [a.values]),
  },
  {
    name: "odoo_write",
    description: "Update existing Odoo records. Returns true on success.",
    inputSchema: {
      type: "object",
      properties: {
        model: { type: "string" },
        ids: { type: "array", items: { type: "integer" } },
        values: { type: "object" },
      },
      required: ["model", "ids", "values"],
    },
    run: (a) => execute(a.model, "write", [a.ids, a.values]),
  },
  {
    name: "odoo_unlink",
    description: "Delete Odoo records by ID. Returns true on success. Use with care.",
    inputSchema: {
      type: "object",
      properties: { model: { type: "string" }, ids: { type: "array", items: { type: "integer" } } },
      required: ["model", "ids"],
    },
    run: (a) => execute(a.model, "unlink", [a.ids]),
  },
  {
    name: "odoo_name_search",
    description: "Fuzzy lookup by display name. Returns [id, display_name] pairs.",
    inputSchema: {
      type: "object",
      properties: {
        model: { type: "string" },
        name: { type: "string" },
        limit: { type: "integer", default: 20 },
      },
      required: ["model", "name"],
    },
    run: (a) => execute(a.model, "name_search", [a.name], { limit: a.limit ?? 20 }),
  },
  {
    name: "odoo_fields_get",
    description: "Describe a model's fields (name, type, help, required, relation, selection).",
    inputSchema: {
      type: "object",
      properties: { model: { type: "string" }, attributes: { type: "array", items: { type: "string" } } },
      required: ["model"],
    },
    run: (a) =>
      execute(a.model, "fields_get", [], {
        attributes: a.attributes || ["string", "type", "help", "required", "relation", "selection"],
      }),
  },
  {
    name: "odoo_list_models",
    description:
      "List available Odoo models via menu actions (label + technical res_model). Optional " +
      "substring filter. Uses ir.actions.act_window because some accounts cannot read ir.model.",
    inputSchema: {
      type: "object",
      properties: { filter: { type: "string" }, limit: { type: "integer", default: 300 } },
    },
    run: (a) => {
      const domain = a.filter
        ? ["|", ["res_model", "ilike", a.filter], ["name", "ilike", a.filter]]
        : [];
      return execute("ir.actions.act_window", "search_read", [domain], {
        fields: ["name", "res_model"],
        limit: a.limit ?? 300,
        order: "res_model",
      });
    },
  },
  {
    name: "odoo_post_note",
    description:
      "Post an HTML note to a record's chatter. Handles the Odoo 17 quirk where message_post " +
      "escapes HTML: posts plain text first, then writes the real markup to mail.message.body. " +
      "Chatter notes can notify real staff — confirm before posting.",
    inputSchema: {
      type: "object",
      properties: {
        model: { type: "string" },
        id: { type: "integer", description: "Record ID to post the note on." },
        html: { type: "string", description: "Note body (HTML allowed)." },
      },
      required: ["model", "id", "html"],
    },
    run: async (a) => {
      const plain = a.html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
      const msgId = await execute(a.model, "message_post", [[a.id]], {
        body: plain,
        message_type: "comment",
        subtype_xmlid: "mail.mt_comment",
      });
      await execute("mail.message", "write", [[msgId], { body: a.html }]);
      return { message_id: msgId };
    },
  },
  {
    name: "odoo_execute",
    description: "Escape hatch: call any model method via execute_kw (model, method, args, kwargs).",
    inputSchema: {
      type: "object",
      properties: {
        model: { type: "string" },
        method: { type: "string" },
        args: { type: "array", default: [] },
        kwargs: { type: "object", default: {} },
      },
      required: ["model", "method"],
    },
    run: (a) => execute(a.model, a.method, a.args ?? [], a.kwargs ?? {}),
  },
  {
    name: "odoo_version",
    description: "Return Odoo server version and verify connectivity (no auth required).",
    inputSchema: { type: "object", properties: {} },
    run: () => version(),
  },
  {
    name: "odoo_whoami",
    description: "Authenticate and return the current user's id, name, login, and company.",
    inputSchema: { type: "object", properties: {} },
    run: async () => {
      const uid = await authenticate();
      const [user] = await execute("res.users", "read", [[uid]], {
        fields: ["name", "login", "company_id", "email"],
      });
      return { uid, ...user };
    },
  },
];

const TOOL_MAP = new Map(TOOLS.map((t) => [t.name, t]));
const TOOL_LISTING = TOOLS.map(({ run, ...rest }) => rest);

async function callTool(name, args) {
  const tool = TOOL_MAP.get(name);
  if (!tool) {
    return { isError: true, content: [{ type: "text", text: `Unknown tool: ${name}` }] };
  }
  try {
    const result = await tool.run(args || {});
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  } catch (err) {
    return { isError: true, content: [{ type: "text", text: `Odoo error: ${err.message}` }] };
  }
}

// ---------------------------------------------------------------------------
// Minimal MCP stdio server (newline-delimited JSON-RPC 2.0)
// ---------------------------------------------------------------------------

function send(msg) {
  process.stdout.write(JSON.stringify(msg) + "\n");
}

async function handle(msg) {
  const { id, method, params } = msg;

  // Notifications (no id) require no response.
  if (id === undefined || id === null) return;

  try {
    if (method === "initialize") {
      return send({
        jsonrpc: "2.0",
        id,
        result: {
          protocolVersion: (params && params.protocolVersion) || "2024-11-05",
          capabilities: { tools: {} },
          serverInfo: { name: "odoo-mcp", version: "1.0.0" },
        },
      });
    }
    if (method === "ping") {
      return send({ jsonrpc: "2.0", id, result: {} });
    }
    if (method === "tools/list") {
      return send({ jsonrpc: "2.0", id, result: { tools: TOOL_LISTING } });
    }
    if (method === "tools/call") {
      const result = await callTool(params && params.name, params && params.arguments);
      return send({ jsonrpc: "2.0", id, result });
    }
    // Unknown method.
    send({ jsonrpc: "2.0", id, error: { code: -32601, message: `Method not found: ${method}` } });
  } catch (err) {
    send({ jsonrpc: "2.0", id, error: { code: -32603, message: String(err && err.message || err) } });
  }
}

function runServer() {
  let buffer = "";
  process.stdin.setEncoding("utf8");
  process.stdin.on("data", (chunk) => {
    buffer += chunk;
    let idx;
    while ((idx = buffer.indexOf("\n")) >= 0) {
      const line = buffer.slice(0, idx).trim();
      buffer = buffer.slice(idx + 1);
      if (!line) continue;
      let msg;
      try {
        msg = JSON.parse(line);
      } catch {
        continue; // ignore malformed lines
      }
      handle(msg);
    }
  });
  process.stdin.on("end", () => process.exit(0));

  const { url, db } = cfg();
  process.stderr.write(
    `[odoo-mcp] ready — ${TOOLS.length} tools for ${url || "(ODOO_URL unset)"} (db: ${db || "unset"})\n`
  );
}

// ---------------------------------------------------------------------------
// Discovery / health check mode
// ---------------------------------------------------------------------------

async function runDiscover() {
  const { url, db, username, apiKey } = cfg();
  if (!url) {
    console.error("Set at least ODOO_URL (ideally ODOO_DB, ODOO_USERNAME, ODOO_API_KEY too).");
    process.exit(1);
  }
  console.log(`Odoo URL: ${url}`);

  try {
    const v = await version();
    console.log("Server version:", JSON.stringify(v));
  } catch (err) {
    console.error("Could not reach Odoo:", err.message);
    process.exit(2);
  }

  // Try to list databases (often disabled on hosted Odoo).
  try {
    const resp = await fetch(`${url}/web/database/list`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ jsonrpc: "2.0", method: "call", params: {} }),
    });
    const j = await resp.json();
    if (j.result) {
      console.log("Databases on this server:", j.result);
      if (!db) console.log("-> Pick one above and use it as ODOO_DB.");
    } else throw new Error("no list");
  } catch {
    console.log("Database listing is disabled on this server (normal for hosted Odoo).");
    if (!db) console.log('-> Set ODOO_DB explicitly (for this server it is "bakertilly").');
  }

  if (!db || !username || !apiKey) {
    console.log("\nSet ODOO_DB, ODOO_USERNAME and ODOO_API_KEY to test the login.");
    return;
  }
  try {
    const uid = await authenticate();
    const [me] = await execute("res.users", "read", [[uid]], { fields: ["name", "login", "company_id"] });
    console.log(`\nSUCCESS: authenticated as uid=${uid}: ${me.name} <${me.login}> (company: ${JSON.stringify(me.company_id)})`);
    console.log("Your connector is correctly configured.");
  } catch (err) {
    console.error("\nFAILED to authenticate:", err.message);
    console.error("Check ODOO_DB, ODOO_USERNAME (login/email), and ODOO_API_KEY.");
    process.exit(3);
  }
}

// ---------------------------------------------------------------------------
// Relay CLI mode:  node odoo-mcp.js call <model> <method> [jsonArgs] [jsonKwargs]
// Lets you run a single Odoo call from the command line and print the result.
// Example: node odoo-mcp.js call res.partner search_count "[[]]"
// ---------------------------------------------------------------------------

async function runCall(a) {
  const [model, method, argsJson = "[]", kwargsJson = "{}"] = a;
  if (!model || !method) {
    console.error('Usage: node odoo-mcp.js call <model> <method> [jsonArgs] [jsonKwargs]');
    process.exit(2);
  }
  let args, kwargs;
  try {
    args = JSON.parse(argsJson);
    kwargs = JSON.parse(kwargsJson);
  } catch (err) {
    console.error("Could not parse the JSON arguments:", err.message);
    process.exit(2);
  }
  try {
    const result = await execute(model, method, args, kwargs);
    console.log(JSON.stringify(result, null, 2));
  } catch (err) {
    console.error("ERROR:", err.message);
    process.exit(1);
  }
}

// ---------------------------------------------------------------------------
// Base64 relay mode:  node odoo-mcp.js callb64 <base64>
// The base64 decodes to JSON {model, method, args?, kwargs?}. Base64 contains
// only safe characters, so the command survives Windows Command Prompt intact
// even when the query uses >, <, |, or quotes. This is the mode Claude hands you.
// ---------------------------------------------------------------------------

async function runCallB64(a) {
  const b64 = a[0];
  if (!b64) {
    console.error("Usage: node odoo-mcp.js callb64 <base64-encoded JSON>");
    process.exit(2);
  }
  let req;
  try {
    req = JSON.parse(Buffer.from(b64, "base64").toString("utf8"));
  } catch (err) {
    console.error("Could not decode/parse the request:", err.message);
    process.exit(2);
  }
  const { model, method, args = [], kwargs = {} } = req || {};
  if (!model || !method) {
    console.error("Decoded request must include at least { model, method }.");
    process.exit(2);
  }
  try {
    const result = await execute(model, method, args, kwargs);
    console.log(JSON.stringify(result, null, 2));
  } catch (err) {
    console.error("ERROR:", err.message);
    process.exit(1);
  }
}

// ---------------------------------------------------------------------------

const _argv = process.argv.slice(2);
if (_argv.includes("--discover")) {
  runDiscover().catch((err) => {
    console.error("Unexpected error:", err);
    process.exit(1);
  });
} else if (_argv[0] === "call") {
  runCall(_argv.slice(1)).catch((err) => {
    console.error("Unexpected error:", err);
    process.exit(1);
  });
} else if (_argv[0] === "callb64") {
  runCallB64(_argv.slice(1)).catch((err) => {
    console.error("Unexpected error:", err);
    process.exit(1);
  });
} else {
  runServer();
}
