// Minimal Odoo JSON-RPC client.
//
// Uses Odoo's external API over JSON-RPC (POST <url>/jsonrpc), so it needs no
// XML-RPC dependency — just the global fetch built into Node 18+.
//
// The client authenticates once (lazily) and caches the resulting uid. All
// model access goes through `execute_kw`, mirroring Odoo's ORM RPC surface.

export class OdooError extends Error {
  constructor(message, data) {
    super(message);
    this.name = "OdooError";
    this.data = data;
  }
}

export class OdooClient {
  /**
   * @param {object} cfg
   * @param {string} cfg.url       Base URL, e.g. https://example.odoo.com
   * @param {string} cfg.db        Database name
   * @param {string} cfg.username  Login (email)
   * @param {string} cfg.apiKey    API key (or password)
   * @param {number} [cfg.timeout] Request timeout in ms (default 30000)
   */
  constructor({ url, db, username, apiKey, timeout = 30000 }) {
    if (!url) throw new Error("Missing ODOO_URL");
    if (!db) throw new Error("Missing ODOO_DB");
    if (!username) throw new Error("Missing ODOO_USERNAME");
    if (!apiKey) throw new Error("Missing ODOO_API_KEY");

    this.url = url.replace(/\/+$/, "");
    this.db = db;
    this.username = username;
    this.apiKey = apiKey;
    this.timeout = timeout;
    this._uid = null;
  }

  async _jsonrpc(service, method, args) {
    const body = {
      jsonrpc: "2.0",
      method: "call",
      params: { service, method, args },
      id: Math.floor(Math.random() * 1e9),
    };

    const controller = new AbortController();
    const t = setTimeout(() => controller.abort(), this.timeout);
    let resp;
    try {
      resp = await fetch(`${this.url}/jsonrpc`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
        signal: controller.signal,
      });
    } catch (err) {
      if (err.name === "AbortError") {
        throw new OdooError(`Request to Odoo timed out after ${this.timeout}ms`);
      }
      throw new OdooError(`Network error contacting Odoo at ${this.url}: ${err.message}`);
    } finally {
      clearTimeout(t);
    }

    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      throw new OdooError(`HTTP ${resp.status} from Odoo: ${text.slice(0, 500)}`);
    }

    const json = await resp.json();
    if (json.error) {
      const data = json.error.data || {};
      const msg = data.message || json.error.message || "Unknown Odoo error";
      throw new OdooError(msg, data);
    }
    return json.result;
  }

  /** Authenticate and cache the uid. */
  async authenticate() {
    if (this._uid) return this._uid;
    const uid = await this._jsonrpc("common", "authenticate", [
      this.db,
      this.username,
      this.apiKey,
      {},
    ]);
    if (!uid) {
      throw new OdooError(
        "Authentication failed — check ODOO_DB, ODOO_USERNAME, and ODOO_API_KEY."
      );
    }
    this._uid = uid;
    return uid;
  }

  /** Return Odoo server version info (no auth required). */
  async version() {
    return this._jsonrpc("common", "version", []);
  }

  /**
   * Call any model method via execute_kw.
   * @param {string} model    e.g. "res.partner"
   * @param {string} method   e.g. "search_read"
   * @param {any[]}  args      positional args
   * @param {object} [kwargs] keyword args
   */
  async execute(model, method, args = [], kwargs = {}) {
    const uid = await this.authenticate();
    return this._jsonrpc("object", "execute_kw", [
      this.db,
      uid,
      this.apiKey,
      model,
      method,
      args,
      kwargs,
    ]);
  }

  // ---- Convenience wrappers over the common ORM methods ----

  searchRead(model, domain = [], { fields, limit, offset, order } = {}) {
    const kwargs = {};
    if (fields) kwargs.fields = fields;
    if (limit != null) kwargs.limit = limit;
    if (offset != null) kwargs.offset = offset;
    if (order) kwargs.order = order;
    return this.execute(model, "search_read", [domain], kwargs);
  }

  searchCount(model, domain = []) {
    return this.execute(model, "search_count", [domain]);
  }

  read(model, ids, fields) {
    const kwargs = fields ? { fields } : {};
    return this.execute(model, "read", [ids], kwargs);
  }

  create(model, values) {
    return this.execute(model, "create", [values]);
  }

  write(model, ids, values) {
    return this.execute(model, "write", [ids, values]);
  }

  unlink(model, ids) {
    return this.execute(model, "unlink", [ids]);
  }

  fieldsGet(model, attributes = ["string", "type", "help", "required", "relation", "selection"]) {
    return this.execute(model, "fields_get", [], { attributes });
  }

  nameSearch(model, name, { limit = 20 } = {}) {
    return this.execute(model, "name_search", [name], { limit });
  }
}
