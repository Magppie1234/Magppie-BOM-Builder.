# AI HANDOFF — Carcass BOM Builder (Next.js + Zoho Inventory)

> **Read this first.** This file is a self-contained brief so any AI can resume work
> on this project without asking the user anything. It contains: what the app does,
> the architecture, the full history of recent changes, the exact code conventions,
> known issues, and the next pending tasks. The complete source is in this zip
> (everything under `src/`). `node_modules`, `.next`, `.cache`, `.git` were excluded —
> run `npm install` to restore dependencies.

---

## 0. How to run

```bash
npm install
npm run dev        # dev server at http://localhost:3000  (main page: /builder)
npm run build      # production build (used to verify type-safety)
npx tsc --noEmit   # type-check only
```

Environment: copy `.env.example` → `.env.local` and fill in the Zoho credentials
(`ZOHO_ORGANIZATION_ID`, `ZOHO_REFRESH_TOKEN`, `ZOHO_CLIENT_ID`, `ZOHO_CLIENT_SECRET`,
optional `ZOHO_ACCOUNTS_BASE_URL`, `ZOHO_INVENTORY_BASE_URL`). The app targets the
**Zoho India** data center by default (`accounts.zoho.in` / `zohoapis.in`).

---

## 1. What the app is

A modular kitchen-cabinet **Bill of Materials (BOM) builder**. The user configures a
cabinet (zone, family, configuration, handle/profile, finishes, dimensions) and the app:

1. Generates a deterministic **cabinet code** (e.g. `BC-DW-XXX-2HB-XCJ-450-720-560-15 ARMANI GRIS/BLACK STATUARIO NUVOLATO`).
2. Computes the full **packet/panel/profile/hardware/consumable** breakdown.
3. Lets the user build a multi-unit **project** and export a styled multi-sheet **Excel** workbook.
4. Pushes the entire nested BOM into **Zoho Inventory** as composite items attached to a Sales Order.

### Main screens / routes
- `/builder` — the primary tool. Renders `<CarcassBomBuilder/>`.
- `/reorder` — reorder report (`<ReorderReport/>`).
- `/qr/bom/[orderId]` — QR-linked BOM view (`<QrBomPage/>`).
- `/` — landing.

---

## 2. Architecture & key files

| File | Purpose |
|---|---|
| `src/components/CarcassBomBuilder.tsx` | **The heart of the app (~5000+ lines).** All BOM logic: model building, packet/panel math, Excel export data, and the full Zoho creation flow. **Almost all recent work happens here.** |
| `src/lib/zoho.ts` | Server-side Zoho OAuth token cache + `zohoInventoryGet/Put/Post` helpers + `toApiError`. Token cached at `.cache/zoho-access-token.json` with refresh skew + rate-limit backoff. |
| `src/lib/types.ts` | Shared TS types: `SalesOrderSummary/Detail`, `SalesOrderLineItem`, `ItemDetail`, `CompositeItemDetail`, `CompositeMappedItem`, `BomReportRow`, `StockStatus`, etc. |
| `src/lib/export.ts` | Excel export helpers (`xlsx-js-style`). |
| `src/lib/report.ts`, `src/lib/stock.ts`, `src/lib/rawmaterial.ts` | Report building, stock status, raw-material resolution. |
| `src/app/api/zoho/*/route.ts` | Next.js API routes proxying Zoho Inventory (items, compositeitems, salesorders, contacts, purchaseorders, reorder-report, status). All `dynamic = "force-dynamic"`. |
| `src/data/*.json` | `stone_finishes.json`, `profile_finishes.json` — finish option catalogs keyed by thickness. |

### Zoho API route conventions
- `GET /api/zoho/items?search=<text>` → uses Zoho `search_text` (per_page 200).
- `GET /api/zoho/items?name=<text>` → uses Zoho `name_contains` (fallback search).
- `POST /api/zoho/items` → create item (simple or service).
- `compositeitems` routes → create/read composite (bundle) items.
- Errors funnel through `toApiError` → `{ error: string }` with 500.

---

## 3. BOM hierarchy model (critical mental model)

The BOM is a nested tree. Levels (as created in Zoho and mirrored in the Excel "Full BOM" tree):

```
Level 0  Cabinet (Carcass)                         e.g. "SO-00029 BC-DW-..."
  Level 1  Set of Parts (packet) / Elenor composite / Hardware Pack / Shutter BOM / Consumable
    Level 2  Part (drilled wrapper) / Set of Profiles / Elenor components
      Level 3  Panel (raw stone + cutting)         (drilled panels are wrapped by a Part)
```

- **Panels under packs** get a **Part** wrapper (Level 2) and the **Panel** sits at Level 3.
  - Part name: `Part ${cleanName} ${drill || "No Drill"} ${finish}` — note: **"Panels" prefix is stripped** via `p.name.replace(/^Panels/i, "")` and the word "Panels" is NOT in Part names.
  - Panel name: `Panel ${cleanName} ${finish}`.
- **Profiles** carry side info in their display name, e.g.
  `Panel Profile ${cleanedProfName} ${len}mm ${finish}(LH+RH)` or `(TP+BT)`.
- **Shutter profiles**: `Panel Profile Shutter Profile (${design}) ${profV}mm ${finish} (LH+RH)` and `... ${profH}mm ... (TP+BT)`.

### Core data interfaces (in CarcassBomBuilder.tsx)
```ts
interface Hardware   { name: string; qty: number; uom?: string; pack?: string; }
interface Consumable { name: string; qty: number; uom: string; pack: string; }
interface Panel      { name: string; w: number; h: number; qty: number; drill: string | null; pack: string; }
interface Profile    { name: string; len: number; qty: number; type: string; pack: string; }
// pkRows entries are tuples: [pkType, name, dimOrNote, qty, uom]
// pkType ∈ "panel" | "prof" | "hard" | "cons" | "shut" | "elen_bom"
```

---

## 4. Light Profiles (Elenor, Skirting, Duplay, L-Profile Dado)

All light profiles have a defined BOM using **exact Zoho item names** (sourced from the
"Light Profile Boms" sheet in `ZOHO BOM Shutter and Cabinet Codes_ New (4).xlsx`).
Each adds **10% waste** to the profile length for length-based components. Drivers are
**selected manually** by the user via dropdown (`DRIVER_OPTIONS`), never auto-added.

```ts
const ELENOR_WASTE = 0.10;

const LIGHT_PROFILE_BOMS: Record<string, {
  profile: string; led: string; ledMultiplier: number;
  diffuser: string; diffuserMultiplier: number;
  wire: string; wireQty: number;
  tape: string | null; tapeMultiplier: number; waste: number;
}> = {
  skirting:  { profile: "ALU PROF FOR SKIRTING WITH LIGHT 3000X100X9.8 ANODISED CHAMPAGNE HM-507 MINA",
               led: "LIGHT FLEXIBLE LED LIGHT 3000K, 180 LED/MTR, 72W, W-5MM XX LED", ledMultiplier: 1,
               diffuser: "PVC DIFFUSER FOR GRAND PROF 3000X6X WHITE 9099 VAI", diffuserMultiplier: 1,
               wire: "LIGHT EXTENSION WIRE TWP SP XX 14/38 SNO", wireQty: 1,
               tape: null, tapeMultiplier: 0, waste: 0.10 },
  duplay:    { profile: "ALU PROF DUPLY LIGHT 3000X45X14 ANODISED CHAMPAGNE HM-534 MINA",
               led: "...", ledMultiplier: 2, diffuser: "...", diffuserMultiplier: 2,
               wire: "...", wireQty: 2,
               tape: "TAPE THERMAL 210 MIC 50 MTR X50X0.3 391990 EURO", tapeMultiplier: 1, waste: 0.10 },
  lprofile:  { profile: "ALU PROF DADO TOOL 3000X35..5X28 ANODISED CHAMPAGNE HM-514 MINA",
               led: "...", ledMultiplier: 1, diffuser: "...", diffuserMultiplier: 1,
               wire: "...", wireQty: 1, tape: null, tapeMultiplier: 0, waste: 0.10 },
  elenor:    { profile: "ALU PROF FOR ELENOR 3000X15X15 ANODISED CHAMPAGNE HM-519 MINA",
               led: "...", ledMultiplier: 1, diffuser: "...", diffuserMultiplier: 1,
               wire: "...", wireQty: 2,
               tape: "TAPE FOR COVER CAP ADH 3M X20X 91031 3M", tapeMultiplier: 1, waste: 0.10 },
};

const DRIVER_OPTIONS = [
  "",
  "LIGHT DRIVER LED STRIP 12V, 2AMP , 24W XX LED",
  "LIGHT DRIVER LED STRIP 12V, 5AMP , 60W XX LED",
];
```

### Elenor = composite item
"**Elenor with Light**" is a single **composite** item in the cabinet BOM. It contains all
its components (profile, LED, diffuser, tape, wire) nested inside. It is injected via:

```ts
function addElenor(H, profiles, cons, hardware, pkRows) {
  const lenWithWaste = Math.ceil(H * (1 + ELENOR_WASTE));
  pkRows.push(["elen_bom", "Elenor with Light", `L=${H}`, 2, "Set"]);
  profiles.push({ name: "ALU PROF FOR ELENOR 3000X15X15 ANODISED CHAMPAGNE HM-519 MINA", len: H, qty: 2, type: "ELEN", pack: "Elenor with Light" });
  cons.push({ name: "LIGHT FLEXIBLE LED LIGHT 3000K, 180 LED/MTR, 72W, W-5MM XX LED", qty: 2*lenWithWaste, uom: "MM", pack: "Elenor with Light" });
  cons.push({ name: "PVC DIFFUSER FOR GRAND PROF 3000X6X WHITE 9099 VAI", qty: 2*lenWithWaste, uom: "MM", pack: "Elenor with Light" });
  cons.push({ name: "TAPE FOR COVER CAP ADH 3M X20X 91031 3M", qty: 2*lenWithWaste, uom: "MM", pack: "Elenor with Light" });
  cons.push({ name: "LIGHT EXTENSION WIRE TWP SP XX 14/38 SNO", qty: 2, uom: "Mtr", pack: "Elenor with Light" });
  // Driver selected manually in Accessories
}
```

`classifyRowType("elen_bom")` returns `"sub_bom"`; `subGroupFromPkType("elen_bom", …)` returns `"Elenor Light"`.

### Accessory type (extended for connectors + driver)
```ts
type AccessoryEntry = {
  enabled: boolean; size: string; elevation: string; actualQty: number | null;
  straightConnectors: number;  // Skirting only
  lConnectors: number;         // Skirting only
  driver: string;              // manual dropdown for all light profiles
};
```
- Skirting UI exposes number inputs for straight + L connectors.
- All `PROFILE_ACC_KEYS` expose a driver dropdown.
- `buildAccessoryExportRows` expands light-profile accessories into individual component
  rows (profile, LED, diffuser, wire, tape, driver, connectors). Non-light accessories
  (J handle, C handle) stay as single rows.

---

## 5. Hardware-pack replacement (done)

The item **"PVC PACKING FOR WALL HANGING 100X100X12 K018 LOC"** was removed from all
**Wall, Loft, and MD** units and replaced with
**"HARDWARE PACK WALL HANGER BRACKET SET/1"** (qty 1, "Set").

---

## 6. Zoho creation flow (in CarcassBomBuilder.tsx)

Helpers (search-then-create pattern everywhere):

- `searchItemByName(name)` — **normalized** match (see §7). Returns the existing Zoho item or `null`.
- `resolveSimpleItem(name, isService=false, customFields?, uom?)` — search; if missing, create.
  Unit precedence: explicit `uom` > inferred from name keywords (set→pcs, kg/glue→Kg,
  ml/silicone→ML, mtr/wire→Mtr, mm/led/diffuser/tape/prof→MM) > default.
- `resolveCompositeItem(rawName, rawSku, mappedItems, customFields?)` — names truncated to 100 chars via `truncZoho`.
- `resolveHardwarePack(hardwareName, qty, packingHardwareId, cf?)` — wraps a physical hardware item + packing into a "Pack Hardware …" composite.
- `resolveProfileBOM(...)` — builds "Set of Profiles" composites.
- `findExistingCompositeByRawMaterial(...)` — dedup panels by raw stone so SO-prefix is dropped when an identical panel already exists.

### Recently implemented (the last work session)
In the carcass packet build loop:
1. **Elenor profiles excluded from regular pack loop** — `carcassProfilesByPack` skips `p.pack === "Elenor with Light"`.
2. **Hardware loop** — items whose name starts with `"HARDWARE PACK"` are searched directly
   and used as-is (created as a simple tracked item with `uom || "Set"` if missing), instead of being
   wrapped again by `resolveHardwarePack`. Other hardware still uses `resolveHardwarePack`.
   Elenor-pack hardware is skipped here.
3. **Consumable loop** — passes `c.uom` into `resolveSimpleItem(c.name, false, undefined, c.uom)`.
   Elenor-pack consumables are skipped here.
4. **Elenor composite** — after the loops, all `pack === "Elenor with Light"` profiles/cons/hardware
   are resolved individually and assembled into a single `resolveCompositeItem("Elenor with Light", …)`
   with cf `{ group: "Elenor Light", subGroup: zoneName }`, then pushed to `carcassMappedItems`
   with the qty from the `elen_bom` pkRow.

---

## 7. MOST RECENT FIX — duplicate items in Zoho

**Symptom:** "GLUE SILICONE XX TRANSPERENT DOWSIL 789 AGG" already exists in Zoho, but
the creation flow created a *new* duplicate instead of reusing it.

**Root cause:** `searchItemByName` compared names with strict equality
(`it.name === name`). Any invisible difference (trailing/double spaces, case) failed the
match → fell through to create.

**Fix applied** (in `searchItemByName`):
```ts
function normalizeItemName(s) {
  return String(s ?? "").replace(/\s+/g, " ").trim().toLowerCase();
}
async function searchItemByName(name) {
  const target = normalizeItemName(name);
  const findMatch = (items) => items.find(it =>
    normalizeItemName(it.name) === target || normalizeItemName(it.sku) === target);
  // search=…  then fallback name=…  both use findMatch
}
```

**Still open / caveat:** if the app's hardcoded name differs in *spelling/wording* from
Zoho (e.g. "TRANSPERENT" vs "TRANSPARENT"), normalization won't help — the app's literal
string must be corrected to the exact Zoho spelling. Fuzzy matching was deliberately
avoided (risk of selecting the wrong item). A suggested optional improvement: log
"close-but-not-exact" candidates during a creation run to surface spelling mismatches.

---

## 8. Excel export
4 sheets via `xlsx-js-style`: **Full BOM** (the nested tree), **Out of Stock**, **Opti**
(optimization), **Accessories**. The Full BOM tree mirrors the Zoho hierarchy in §3
(Part wrapper for drilled panels; Elenor nested as a composite; profiles carry side labels).

---

## 9. External reference files (NOT in this zip — they live in the user's Downloads)
- `Hardware Packs.xlsx` — master list of hardware packs. (PVC PACKING row was given qty 2, UOM "PCS".)
- `ZOHO BOM Shutter and Cabinet Codes_ New (4).xlsx` — read-only reference; "Light Profile Boms" sheet provides exact Zoho item names.

If you need these, ask the user to re-share them.

---

## 10. Conventions & gotchas
- **Always run `npx tsc --noEmit` (or `npm run build`) after edits** — the user verifies type-safety this way. Both currently pass clean.
- Item/composite names are **UPPERCASE** matching Zoho catalog spelling. Treat catalog strings as exact.
- Zoho names are truncated to **100 chars** on create (`truncZoho`).
- The app targets **Zoho India** endpoints.
- `pkRows` tuple order is `[pkType, name, dimOrNote, qty, uom]`.
- When grouping by pack, "Elenor with Light" is a **special pack** handled as a composite — never as a "Set of Parts".

---

## 11. Pending / suggested next tasks
1. **Verify the Elenor composite end-to-end** against a real Sales Order push (the
   creation code is written but should be confirmed against live Zoho data).
2. **Skirting connector Zoho names** — confirm exact item names for straight/L connectors
   (currently may use placeholders "SKIRTING STRAIGHT CONNECTOR" / "SKIRTING L CONNECTOR").
3. **Spelling-mismatch surfacing** — optionally log near-miss candidates from
   `searchItemByName` so the user can fix any hardcoded names that differ from Zoho.
4. **Confirm driver-by-pack wiring** — ensure manually-selected drivers flow into both the
   Excel accessory rows and (if intended) the Zoho composite.

---

## 12. Full change log of the recent sessions (chronological intent)
8. **Fix duplicate creation** — normalize name comparison in `searchItemByName` (case + whitespace).
9. **AI Copilot Integration**:
   - Added `OPENAI_API_KEY` configuration.
   - Built server API chat handler at `src/app/api/ai/chat/route.ts`.
   - Built dynamic context hooks in `BomDashboard.tsx` and `CarcassBomBuilder.tsx`.
   - Wired floating frontend widget `AiCopilot.tsx` globally into root `layout.tsx`.

10. **Wall Stone Shutter & Shutter Glass Restructuring** (June 13, 2026):
    - Added the `"WST"` (Wall Stone Shutter) cabinet family with standard variants and support for 3 wood shelves.
    - Added size and zone mappings for `WST` under `ZOHO_BOM_SIZES`, `famSetOf` for WB zone, and `shutSpec`.
    - Added wood shelf and HM511 shelf profile generation for wall/loft cabinets in `buildCarcassInnerRaw` when `f.shelf` is defined.
    - Restructured the `"WGL"` (Wall Glass Shutter) BOM: removed glass from carcass cabinet level and moved it directly to the shutter BOM as a glass panel replacement (`GLASS TOUGH EP ...`).
    - Updated Zoho sync for `WGL`: bypassed `shutterStoneId` select validation and pushed resolved glass simple items directly into `shutterParts` instead of generating a Level 4 panel composite wrapper.
    - Added glass shutter panels to `runStockCheck`.

---

## 13. AI Copilot Integration & Dynamic Context

The application has been transformed into an AI-driven system by adding a floating **AI Copilot** widget on all pages:
- **Server API Route**: `src/app/api/ai/chat/route.ts` proxies requests to OpenAI's Chat Completions (`gpt-4o-mini`). It is configured with a rich system prompt that includes cabinetry rules, waste percentages, shutter designs, and design exclusions.
- **Floating Chat Widget**: `src/components/AiCopilot.tsx` implements a premium chat widget with linear gradient buttons, smooth animations, glassmorphism panel styles, and robust error handling.
- **Global Integration**: Wired into the root layout at `src/app/layout.tsx`.
- **Active Page Context**:
  - **Dashboard**: Exposes critical low-stock/out-of-stock items, active order list, and dashboard stats.
  - **Builder**: Exposes the currently active cabinet code, dimensions, materials, styles, active family, and valid designs.
- **Known Issue**: The OpenAI API Key provided (`sk-proj-...MgYA`) was deactivated/revoked by OpenAI's automated public key scanner. You need to replace this key in `.env.local` and Vercel environment variables with a new, valid key to re-enable AI functionality.

> The previous handoff notes `CHAT_SUMMARY.md`, `HANDOFF-SUMMARY.md`, and `HANDOFF.md`
> are older snapshots also included in the zip for additional context.
