#!/usr/bin/env node
// Connectivity & discovery helper. Run this on the machine that can reach your
// Odoo instance BEFORE wiring the connector into Claude:
//
//   ODOO_URL=... ODOO_DB=... ODOO_USERNAME=... ODOO_API_KEY=... node src/discover.js
//
// It will: report the server version, list databases (if the server exposes
// them), authenticate, and print who you are. Any failure here tells you which
// setting is wrong before Claude ever touches it.

import { OdooClient } from "./odoo-client.js";

const url = (process.env.ODOO_URL || "").replace(/\/+$/, "");
const db = process.env.ODOO_DB;
const username = process.env.ODOO_USERNAME;
const apiKey = process.env.ODOO_API_KEY || process.env.ODOO_PASSWORD;

if (!url) {
  console.error("Set at least ODOO_URL. Ideally also ODOO_DB, ODOO_USERNAME, ODOO_API_KEY.");
  process.exit(1);
}

async function listDatabases() {
  // Odoo exposes db listing at /web/database/list when db-listing is enabled.
  try {
    const resp = await fetch(`${url}/web/database/list`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ jsonrpc: "2.0", method: "call", params: {} }),
    });
    const json = await resp.json();
    return json.result || null;
  } catch {
    return null;
  }
}

async function main() {
  console.log(`Odoo URL: ${url}`);

  // Version (no auth) — proves basic reachability.
  try {
    const probe = new OdooClient({ url, db: db || "x", username: username || "x", apiKey: apiKey || "x" });
    const version = await probe.version();
    console.log("Server version:", JSON.stringify(version));
  } catch (err) {
    console.error("Could not reach Odoo (version call failed):", err.message);
    process.exit(2);
  }

  const dbs = await listDatabases();
  if (dbs) {
    console.log("Databases visible on this server:", dbs);
    if (!db) console.log("→ Pick one of the above and set it as ODOO_DB.");
  } else {
    console.log("Database listing is disabled on this server (normal for SaaS/Online).");
    if (!db) console.log("→ You must set ODOO_DB explicitly. For Odoo Online it is usually your subdomain, e.g. \"mycompany\".");
  }

  if (!db || !username || !apiKey) {
    console.log("\nSet ODOO_DB, ODOO_USERNAME and ODOO_API_KEY to test authentication.");
    return;
  }

  const client = new OdooClient({ url, db, username, apiKey });
  try {
    const uid = await client.authenticate();
    const [me] = await client.read("res.users", [uid], ["name", "login", "company_id"]);
    console.log(`\n✅ Authenticated as uid=${uid}: ${me.name} <${me.login}> (company: ${JSON.stringify(me.company_id)})`);
    console.log("Your connector is correctly configured.");
  } catch (err) {
    console.error("\n❌ Authentication failed:", err.message);
    console.error("Check ODOO_DB, ODOO_USERNAME (login/email), and ODOO_API_KEY.");
    process.exit(3);
  }
}

main().catch((err) => {
  console.error("Unexpected error:", err);
  process.exit(1);
});
