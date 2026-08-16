// MCP tool definitions and handlers for the Odoo connector.
//
// Each tool maps to an Odoo ORM operation. Results are returned as compact
// JSON text so Claude can read them directly.

/** JSON-serialise a result for return as MCP text content. */
function asText(value) {
  return { content: [{ type: "text", text: JSON.stringify(value, null, 2) }] };
}

function asError(err) {
  const detail = err && err.message ? err.message : String(err);
  return {
    isError: true,
    content: [{ type: "text", text: `Odoo error: ${detail}` }],
  };
}

/**
 * Build the list of tools and a dispatch map bound to a given OdooClient.
 * @param {import("./odoo-client.js").OdooClient} client
 */
export function buildTools(client) {
  const tools = [
    {
      name: "odoo_search_read",
      description:
        "Search Odoo records and return chosen fields in one call. Use an Odoo domain " +
        "(list of triples, e.g. [[\"customer_rank\",\">\",0]]). Empty domain [] returns all.",
      inputSchema: {
        type: "object",
        properties: {
          model: { type: "string", description: "Model name, e.g. res.partner, sale.order, account.move" },
          domain: {
            type: "array",
            description: "Odoo search domain, e.g. [[\"name\",\"ilike\",\"acme\"]]. Use [] for no filter.",
            default: [],
          },
          fields: {
            type: "array",
            items: { type: "string" },
            description: "Field names to return. Omit for a default set (can be large).",
          },
          limit: { type: "integer", description: "Max records (default 50).", default: 50 },
          offset: { type: "integer", description: "Records to skip (paging).", default: 0 },
          order: { type: "string", description: "Sort clause, e.g. \"create_date desc\"." },
        },
        required: ["model"],
      },
      handler: (a) =>
        client.searchRead(a.model, a.domain ?? [], {
          fields: a.fields,
          limit: a.limit ?? 50,
          offset: a.offset ?? 0,
          order: a.order,
        }),
    },
    {
      name: "odoo_search_count",
      description: "Count Odoo records matching a domain (fast, returns just a number).",
      inputSchema: {
        type: "object",
        properties: {
          model: { type: "string" },
          domain: { type: "array", default: [] },
        },
        required: ["model"],
      },
      handler: (a) => client.searchCount(a.model, a.domain ?? []),
    },
    {
      name: "odoo_read",
      description: "Read specific Odoo records by their IDs.",
      inputSchema: {
        type: "object",
        properties: {
          model: { type: "string" },
          ids: { type: "array", items: { type: "integer" }, description: "Record IDs to read." },
          fields: { type: "array", items: { type: "string" }, description: "Fields to return (optional)." },
        },
        required: ["model", "ids"],
      },
      handler: (a) => client.read(a.model, a.ids, a.fields),
    },
    {
      name: "odoo_create",
      description: "Create a new Odoo record. Returns the new record ID.",
      inputSchema: {
        type: "object",
        properties: {
          model: { type: "string" },
          values: { type: "object", description: "Field/value map for the new record." },
        },
        required: ["model", "values"],
      },
      handler: (a) => client.create(a.model, a.values),
    },
    {
      name: "odoo_write",
      description: "Update existing Odoo records. Returns true on success.",
      inputSchema: {
        type: "object",
        properties: {
          model: { type: "string" },
          ids: { type: "array", items: { type: "integer" }, description: "Record IDs to update." },
          values: { type: "object", description: "Field/value map to write." },
        },
        required: ["model", "ids", "values"],
      },
      handler: (a) => client.write(a.model, a.ids, a.values),
    },
    {
      name: "odoo_unlink",
      description: "Delete Odoo records by ID. Returns true on success. Use with care.",
      inputSchema: {
        type: "object",
        properties: {
          model: { type: "string" },
          ids: { type: "array", items: { type: "integer" } },
        },
        required: ["model", "ids"],
      },
      handler: (a) => client.unlink(a.model, a.ids),
    },
    {
      name: "odoo_name_search",
      description:
        "Quick fuzzy lookup by display name. Returns [id, display_name] pairs — handy for " +
        "resolving a partner/product/etc. before referencing it.",
      inputSchema: {
        type: "object",
        properties: {
          model: { type: "string" },
          name: { type: "string", description: "Text to match against the record display name." },
          limit: { type: "integer", default: 20 },
        },
        required: ["model", "name"],
      },
      handler: (a) => client.nameSearch(a.model, a.name, { limit: a.limit ?? 20 }),
    },
    {
      name: "odoo_fields_get",
      description:
        "Describe the fields of an Odoo model (name, type, help, required, relation). " +
        "Use this to discover what fields exist before querying or writing.",
      inputSchema: {
        type: "object",
        properties: {
          model: { type: "string" },
          attributes: {
            type: "array",
            items: { type: "string" },
            description: "Field attributes to return (default: string,type,help,required,relation,selection).",
          },
        },
        required: ["model"],
      },
      handler: (a) =>
        a.attributes ? client.fieldsGet(a.model, a.attributes) : client.fieldsGet(a.model),
    },
    {
      name: "odoo_list_models",
      description:
        "List available Odoo models (technical name + label). Optionally filter by a " +
        "substring matched against the model name or label.",
      inputSchema: {
        type: "object",
        properties: {
          filter: { type: "string", description: "Optional substring, e.g. \"sale\" or \"partner\"." },
          limit: { type: "integer", default: 200 },
        },
      },
      handler: (a) => {
        const domain = a.filter
          ? ["|", ["model", "ilike", a.filter], ["name", "ilike", a.filter]]
          : [];
        return client.searchRead("ir.model", domain, {
          fields: ["model", "name"],
          limit: a.limit ?? 200,
          order: "model",
        });
      },
    },
    {
      name: "odoo_execute",
      description:
        "Escape hatch: call an arbitrary model method via execute_kw " +
        "(model, method, args, kwargs). Use when no dedicated tool fits, e.g. action_confirm.",
      inputSchema: {
        type: "object",
        properties: {
          model: { type: "string" },
          method: { type: "string", description: "Method name, e.g. search_read, action_post." },
          args: { type: "array", description: "Positional args array.", default: [] },
          kwargs: { type: "object", description: "Keyword args object.", default: {} },
        },
        required: ["model", "method"],
      },
      handler: (a) => client.execute(a.model, a.method, a.args ?? [], a.kwargs ?? {}),
    },
    {
      name: "odoo_version",
      description: "Return Odoo server version info and verify connectivity (no auth required).",
      inputSchema: { type: "object", properties: {} },
      handler: () => client.version(),
    },
    {
      name: "odoo_whoami",
      description: "Authenticate and return the current user's id, name, login, and company.",
      inputSchema: { type: "object", properties: {} },
      handler: async () => {
        const uid = await client.authenticate();
        const [user] = await client.read("res.users", [uid], [
          "name",
          "login",
          "company_id",
          "email",
        ]);
        return { uid, ...user };
      },
    },
  ];

  const dispatch = new Map(tools.map((t) => [t.name, t.handler]));

  async function call(name, args) {
    const handler = dispatch.get(name);
    if (!handler) {
      return asError(new Error(`Unknown tool: ${name}`));
    }
    try {
      const result = await handler(args || {});
      return asText(result);
    } catch (err) {
      return asError(err);
    }
  }

  // Strip the internal handler when advertising tools over MCP.
  const listing = tools.map(({ handler, ...rest }) => rest);

  return { tools: listing, call };
}
