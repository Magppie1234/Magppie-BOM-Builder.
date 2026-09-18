import fs from "node:fs/promises";
import path from "node:path";

const REQUIRED_ENV = ["ZOHO_ORGANIZATION_ID", "ZOHO_REFRESH_TOKEN", "ZOHO_CLIENT_ID", "ZOHO_CLIENT_SECRET"];
for (const name of REQUIRED_ENV) {
  if (!process.env[name]) throw new Error(`Missing ${name}. Load the production credentials before syncing.`);
}

const catalogPath = path.join(process.cwd(), "src/data/item_catalog.json");
const finishPath = path.join(process.cwd(), "src/data/stone_finishes.json");
const current = JSON.parse(await fs.readFile(catalogPath, "utf8"));
const finishGroups = JSON.parse(await fs.readFile(finishPath, "utf8"));
const knownFinishes = [...new Set(Object.values(finishGroups).flat())]
  .map((value) => String(value).toUpperCase())
  .sort((a, b) => b.length - a.length);

const tokenParams = new URLSearchParams({
  refresh_token: process.env.ZOHO_REFRESH_TOKEN,
  client_id: process.env.ZOHO_CLIENT_ID,
  client_secret: process.env.ZOHO_CLIENT_SECRET,
  grant_type: "refresh_token",
});
const tokenResponse = await fetch(`https://accounts.zoho.in/oauth/v2/token?${tokenParams}`, { method: "POST" });
const tokenBody = await tokenResponse.json();
if (!tokenResponse.ok || !tokenBody.access_token) throw new Error("Zoho authentication failed while syncing the item catalogue.");

const zohoItems = [];
for (let page = 1; ; page += 1) {
  const url = new URL("https://www.zohoapis.in/inventory/v1/items");
  url.searchParams.set("organization_id", process.env.ZOHO_ORGANIZATION_ID);
  url.searchParams.set("page", String(page));
  url.searchParams.set("per_page", "200");
  const response = await fetch(url, { headers: { Authorization: `Zoho-oauthtoken ${tokenBody.access_token}` } });
  const body = await response.json();
  if (!response.ok || (typeof body.code === "number" && body.code > 0)) {
    throw new Error(`Zoho item page ${page} failed during catalogue sync.`);
  }
  zohoItems.push(...(body.items || []));
  process.stdout.write(`Fetched ${zohoItems.length} Zoho items\r`);
  if (!body.page_context?.has_more_page) break;
}
process.stdout.write("\n");

function dimensionsFromName(name) {
  const normalized = String(name || "").toUpperCase();
  const match = normalized.match(/\b(\d{2,4})\s*[X×]\s*(\d{2,4})\s*[X×]\s*(\d{1,3})\s*(?:MM)?\b/);
  if (!match) return { height: "", width: "", thickness: "" };
  return { height: match[1], width: match[2], thickness: match[3] };
}

function finishFromName(name) {
  const normalized = String(name || "").toUpperCase();
  return knownFinishes.find((finish) => normalized.includes(finish)) || "";
}

function profileFromName(name) {
  return String(name || "").toUpperCase().match(/\b(HM\d{3}|MD[123]|CL[12]|CM[12]|ELEN|STP|SLF|SINK|DBS|DBC)\b/)?.[1] || "";
}

function normalizeZohoItem(item) {
  const name = item.name || item.item_name || "";
  const dims = dimensionsFromName(name);
  const stone = /\bSTONE\b/i.test(name);
  const profileCode = item.cf_profile_code || profileFromName(name);
  const profile = /\bPROFILE\b/i.test(name) || !!profileCode;
  const custom = {
    cf_group: item.cf_group || (stone || profile ? "Raw Material" : ""),
    cf_sub_group: item.cf_sub_group || (stone ? "Stone" : profile ? "Profile" : ""),
    cf_finish: item.cf_finish || finishFromName(name),
    cf_profile_code: profileCode,
    cf_height: item.cf_height || dims.height,
    cf_width: item.cf_width || dims.width,
    cf_thickness: item.cf_thickness || dims.thickness,
    cf_msl: item.cf_msl ?? null,
    cf_rol: item.cf_rol ?? item.reorder_level ?? null,
  };
  return {
    item_id: String(item.item_id),
    name,
    item_name: item.item_name || name,
    sku: item.sku || "",
    ...custom,
    unit: item.unit || "",
    status: item.status || "active",
    item_type: item.item_type || "inventory",
    product_type: item.product_type || "goods",
    description: item.description || "",
    group_name: item.group_name || "",
    source: "Zoho catalogue snapshot",
    sync_status: "Synced",
    created_time: item.created_time || "",
    last_modified_time: item.last_modified_time || "",
    stock_on_hand: item.stock_on_hand ?? null,
    available_stock: item.available_stock ?? null,
    actual_available_stock: item.actual_available_stock ?? null,
    is_combo_product: item.is_combo_product === true || item.is_combo_product === "true",
    custom_field_hash: custom,
  };
}

const mergedById = new Map(Object.values(current).map((item) => [String(item.item_id), item]));
for (const item of zohoItems) mergedById.set(String(item.item_id), normalizeZohoItem(item));

// Prefer the current Zoho record when the old spreadsheet and Zoho contain the same SKU/name.
const zohoIds = new Set(zohoItems.map((item) => String(item.item_id)));
const bestByIdentity = new Map();
for (const item of mergedById.values()) {
  const identity = String(item.sku || item.name || item.item_id).trim().toLowerCase();
  const existing = bestByIdentity.get(identity);
  if (!existing || zohoIds.has(String(item.item_id))) bestByIdentity.set(identity, item);
}
const sorted = [...bestByIdentity.values()].sort((a, b) =>
  String(a.name || "").localeCompare(String(b.name || ""), undefined, { sensitivity: "base" }),
);
const output = Object.fromEntries(sorted.map((item) => [String(item.item_id), item]));
await fs.writeFile(catalogPath, `${JSON.stringify(output)}\n`, "utf8");
console.log(`Catalogue synced: ${Object.keys(current).length} existing + ${zohoItems.length} Zoho → ${sorted.length} unique items.`);
