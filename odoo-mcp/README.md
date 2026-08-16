# Odoo MCP Connector

An [MCP](https://modelcontextprotocol.io) server that connects Claude to an
**Odoo** instance through Odoo's external **JSON-RPC** API. Once configured,
Claude can search, read, create, update, and delete Odoo records — partners,
sales orders, invoices, products, inventory, CRM leads, anything the logged-in
user is allowed to touch.

It runs as a local **stdio** MCP server, so it works with **Claude Desktop**,
**Claude Code**, and any other MCP client.

> ⚠️ **Where this runs:** the connector talks to Odoo from the machine it runs
> on (your laptop / server). That machine must be able to reach your Odoo URL.

---

## 1. Prerequisites

- **Node.js ≥ 18** (uses the built-in `fetch`).
- An Odoo **API key** — create one at:
  Odoo → top-right avatar → **My Profile** → **Account Security** → **New API Key**.
- Your Odoo **database name** (see [Finding your database name](#finding-your-database-name)).

## 2. Install

```bash
cd odoo-mcp
npm install
```

## 3. Configure

The connector reads four settings from the environment:

| Variable        | Required | Description                                            |
|-----------------|----------|--------------------------------------------------------|
| `ODOO_URL`      | ✅       | Base URL, e.g. `https://qacoatwork.com`                |
| `ODOO_DB`       | ✅       | Database name                                          |
| `ODOO_USERNAME` | ✅       | Login (email)                                          |
| `ODOO_API_KEY`  | ✅       | API key (or account password if keys aren't enforced) |
| `ODOO_TIMEOUT`  | ❌       | Request timeout in ms (default `30000`)               |

For local testing, copy `.env.example` to `.env` and fill it in. **Never commit
`.env`** — it is gitignored. When wiring into Claude, you'll pass these in the
MCP config `env` block instead (shown below).

## 4. Verify before wiring into Claude

Run the discovery/health check from the machine that can reach Odoo:

```bash
ODOO_URL="https://qacoatwork.com" \
ODOO_DB="<your-db>" \
ODOO_USERNAME="idrees.ca@gmail.com" \
ODOO_API_KEY="<your-api-key>" \
npm run discover
```

It reports the server version, lists databases (if the server allows it),
authenticates, and prints who you are. Fix any error here **before** adding it
to Claude — it isolates config problems from Claude problems.

## 5. Add to Claude

### Claude Code (CLI)

```bash
claude mcp add odoo \
  --env ODOO_URL="https://qacoatwork.com" \
  --env ODOO_DB="<your-db>" \
  --env ODOO_USERNAME="idrees.ca@gmail.com" \
  --env ODOO_API_KEY="<your-api-key>" \
  -- node /absolute/path/to/odoo-mcp/src/index.js
```

### Claude Desktop

Edit the config file:

- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

Add an `odoo` entry under `mcpServers`:

```json
{
  "mcpServers": {
    "odoo": {
      "command": "node",
      "args": ["/absolute/path/to/odoo-mcp/src/index.js"],
      "env": {
        "ODOO_URL": "https://qacoatwork.com",
        "ODOO_DB": "<your-db>",
        "ODOO_USERNAME": "idrees.ca@gmail.com",
        "ODOO_API_KEY": "<your-api-key>"
      }
    }
  }
}
```

Restart Claude Desktop. You should see the `odoo` tools appear (the 🔌 / tools menu).

## 6. Try it

Ask Claude things like:

- *"Use odoo_whoami to confirm the connection."*
- *"List the 10 most recent sales orders with their totals."*
- *"How many customers do we have?"*
- *"Find the partner named 'Acme' and show their contact details."*
- *"Create a CRM lead for 'Website enquiry' from john@acme.com."*

---

## Available tools

| Tool                | What it does                                                        |
|---------------------|--------------------------------------------------------------------|
| `odoo_version`      | Server version / connectivity check (no auth).                     |
| `odoo_whoami`       | Authenticate and return the current user + company.                |
| `odoo_list_models`  | List models (optionally filtered), e.g. everything matching "sale".|
| `odoo_fields_get`   | Describe a model's fields (type, help, required, relation).        |
| `odoo_search_read`  | Search with an Odoo domain and return chosen fields.               |
| `odoo_search_count` | Count records matching a domain.                                   |
| `odoo_read`         | Read specific records by ID.                                       |
| `odoo_name_search`  | Fuzzy lookup by display name → `[id, name]` pairs.                 |
| `odoo_create`       | Create a record.                                                   |
| `odoo_write`        | Update records.                                                    |
| `odoo_unlink`       | Delete records.                                                    |
| `odoo_execute`      | Escape hatch: call any model method via `execute_kw`.              |

**Odoo domains** are lists of triples, e.g.
`[["customer_rank", ">", 0], ["country_id.code", "=", "PK"]]`.
Start with `odoo_list_models` and `odoo_fields_get` to discover what's available.

## Finding your database name

The DB name is **not** in the web URL (the `cids=1` there is a *company* id, not
the database). To find it:

- Run `npm run discover` — if database listing is enabled, it prints the list.
- Or in Odoo, go to **Settings → Activate developer mode**, then the database
  name appears in **Settings → Technical**, or in the page's About dialog.
- On **Odoo Online (SaaS)** it's usually your subdomain, e.g. `mycompany` for
  `mycompany.odoo.com`.

## Security notes

- **The API key is a credential.** Keep it in the MCP `env` block or a local
  `.env`; never commit it. It grants Claude the same access your user has —
  scope the Odoo user's permissions accordingly.
- `odoo_write` / `odoo_unlink` **modify and delete data.** In Claude, keep tool
  approvals on for these until you trust a given workflow. To make the connector
  strictly read-only, remove the `odoo_create`, `odoo_write`, `odoo_unlink`, and
  `odoo_execute` tools from `src/tools.js`.
- Rotate the key (Account Security → API Keys) if it's ever exposed.

## Troubleshooting

| Symptom                                   | Fix                                                                 |
|-------------------------------------------|---------------------------------------------------------------------|
| `Authentication failed`                   | Check `ODOO_DB`, `ODOO_USERNAME` (the login/email), and `ODOO_API_KEY`. |
| `Missing required environment variables`  | One of the required env vars isn't set in the MCP config.           |
| `Network error contacting Odoo`           | The machine can't reach `ODOO_URL` (firewall/VPN/DNS).             |
| Tools don't appear in Claude              | Use an **absolute** path to `src/index.js`; restart the client.     |
| Slow/timeout                              | Raise `ODOO_TIMEOUT`, or narrow domains and set `limit`.            |

## How it works

- `src/odoo-client.js` — tiny JSON-RPC client (`authenticate` + `execute_kw`), no XML-RPC dependency.
- `src/tools.js` — MCP tool definitions and handlers mapping to Odoo ORM methods.
- `src/index.js` — stdio MCP server wiring the two together.
- `src/discover.js` — standalone connectivity/credentials check.
