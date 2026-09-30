# APP ARCHITECTURE BRIEF — Magppie Carcass BOM Builder

A complete, AI-portable description of how this app is built: data model, construction
pipeline (how cabinets/BOMs are generated), naming convention, BOM explosion, and the
Zoho Inventory sync. Read this top-to-bottom and you can continue development on any AI
without the prior chat history.

---

## 0. TL;DR — what the app does

It is a **kitchen-carcass BOM (Bill of Materials) builder**. A user configures a cabinet
(zone → family → variant → size → finishes), and the app:

1. **Constructs a cabinet model** — derives every panel, profile, hardware pack, and
   shutter with exact cut dimensions from a set of parametric formulas.
2. **Explodes it into a multi-level BOM tree** (master → pack → part → panel → raw),
   applying material-waste factors and a cutting/drilling service model.
3. **Pushes it to Zoho Inventory (India DC)** as composite items built bottom-up, then
   attaches them to a Sales Order.

The entire domain logic lives in **ONE component**: `src/components/CarcassBomBuilder.tsx`
(~9,062 lines). Everything else (libs, API routes, other components) is plumbing around it.

---

## 1. Environment & project layout

- **Stack:** Next.js 16 (App Router, webpack + wasm SWC), React 19, TypeScript 5.9, Node 24.
- **Active codebase (work ONLY here):** `~/Documents/bom-nextjs-app/Inventory - Magppie`
  - NOT `~/Downloads` copies; NOT the parent `~/Documents/bom-nextjs-app` (a *different*
    app, "BOM Backorder Report", only `/qr`).
- **Main route:** `/builder` → `src/app/builder/page.tsx` → renders `<CarcassBomBuilder/>`.
- **Env (`.env.local`):** `ZOHO_*` creds (India DC), `OPENAI_API_KEY`.
- **Not a git repo.**

### Verify after EVERY edit (required)
```bash
cd "/Users/Apple/Documents/bom-nextjs-app/Inventory - Magppie"
node scratch/test-drawers.js      # drawer unit/architecture tests → "ALL TESTS PASSED SUCCESSFULLY!"
npx tsc --noEmit                  # type check → exit 0
```
**Dev-server gotcha:** the Claude preview MCP launches the WRONG parent app, so `/builder`
404s. To run the real server use `npm run dev` (port 3000) **inside the `Inventory - Magppie`
folder**. Logic is covered by `scratch/test-drawers.js`; browser preview is not part of the loop.

### Source tree
```
src/
  app/
    page.tsx, layout.tsx, globals.css
    builder/page.tsx               → <CarcassBomBuilder/>  (the app)
    reorder/page.tsx               → <ReorderReport/>
    qr/bom/[orderId]/page.tsx      → <QrBomPage/>  (read-only mobile BOM by SO)
    api/
      ai/chat/route.ts             → OpenAI gpt-4o-mini copilot
      zoho/
        items/route.ts, items/[id]/vendors/route.ts
        compositeitems/route.ts
        salesorders/route.ts
        contacts/route.ts
        purchaseorders/route.ts
        po-vendor-map/route.ts
        reorder-report/route.ts
        status/route.ts
  components/
    CarcassBomBuilder.tsx          ← HEART OF APP (~9,062 lines)
    BomDashboard.tsx (~2,300)      ← SO browser / BOM viewer
    BomOrderCard.tsx, DraftPoModal.tsx, ReorderReport.tsx
    AuthGate.tsx                   ← password gate (Factory@1234, localStorage)
    AiCopilot.tsx                  ← chat UI to /api/ai/chat
    QrBomPage.tsx
  lib/
    types.ts        ← shared TS interfaces
    naming.ts       ← BOM name parsing/normalization helpers
    zoho.ts         ← Zoho OAuth + REST wrappers (India DC)
    stock.ts        ← stock-check helpers
    rawmaterial.ts  ← raw-material resolution
    report.ts       ← reorder/report aggregation
    export.ts       ← Excel/CSV export
  data/
    profile_finishes.json
    stone_finishes.json
scratch/
  test-drawers.js   ← self-contained Node test mirroring component logic
```

---

## 2. Data model (core interfaces)

These are the shapes flowing through construction → BOM → Zoho. (Defined inline in the
component and/or `src/lib/types.ts`.)

```ts
Panel    { name; w; h; qty; drill; pack; t? }          // a flat sheet part with cut size w×h
Profile  { name; len; qty; type; pack }                // an extruded aluminium profile, cut to len
Hardware { name; qty; uom?; pack? }                    // a purchased hardware pack
Consumable { name; qty; uom; pack }                    // glue, edgeband, etc.
Shutter  { code; kind; design; w; h; pw; ph; profV; profH; hinge; hq? }

CarcassModel {
  code; mat; finishes;                                  // cabinet code string + material/finish
  f; v;                                                 // family record, variant record
  sn; lh; rh;                                           // side flags / left-right handed
  W; H; D; t;                                           // outer Width, Height, Depth, panel thickness
  isCJ;                                                 // corner-joint flag
  pkRows;                                               // pack rows [type,name,dim,qty,uom]
  panels; profiles; hardware; cons;                     // component arrays
  netSqft; nPanels; ops;                                // rollups
  shutters;                                             // shutter array
  zk; fk;                                               // zone key, family key
}

ProjectLine { qty; m; rate? }                           // a model placed in the project (m = model)
StockItem   { ... }                                     // Zoho stock snapshot row
FullBomRow  { ... }                                     // a flattened BOM tree row for display/export
```

### Constants
```
STEP        = 2          // generic step/clearance
CJ_CUT      = 23         // corner-joint cut allowance
SHELF_OFF   = 40         // shelf setback
SQDIV       = 92903.04   // mm² per ft² (area conversion)
KG_SQFT     = 3.45       // stone weight per ft²
STONE_WASTE   = 0.15     // +15% stone
PROFILE_WASTE = 0.20     // +20% profile
ELENOR_WASTE  = 0.10     // +10% elenor profile
```

### Profile codes (aluminium extrusions)
```
STP  = HM509   SLF = HM511   SINK = HM510   ELEN = HM519
DBS  = HM513   (drawer side profile, LH/RH split when qty===2)
DBC  = HM535   (drawer center profile, flat cut)
NO_DRILL_PROFILES = { STP, ELEN, SLF, SINK }   // DBS/DBC are drilled-class (NOT in this set)
```

---

## 3. Configuration taxonomy (ZONES → FAMILIES → VARIANTS → SIZES)

The UI funnels the user through a 4-level pick. Each level is a data table inside the component.

### ZONES (13)
`BC, BCL, BB, BBL, WC, WB, TC, TB, TCL, LO, LB, LBF, MD`
Each zone record carries: `p1` (code prefix), `construct`, `mount`, feature flags, and a
`fams` list of family keys it allows. `famSetOf(zone)` maps a zone's `fams` → the family
record table to use.

### FAMILIES (grouped by construction archetype)
- **BASE:** `DW (Drawer), HO, SK (Sink), SH (Shelf), GD, AP, BPO, WBP`
- **BLIND:** `LMC, BSH, PLB`
- **WALL:** `WGL, WST, WOP, WDR`
- **TALL:** `SHF, APP, PAN, DPN, PPN, REF`
- **TALL_BLIND:** `BLND`
- **TALL_LOW:** `LOWS`
- **LOFT:** `LFT`
- **MD:** `MDR`

Each family record has: `p2` (code segment), feature flags (`top/shelf/glass/bracket/holes`),
and a `variants` list. Each **variant** = `{ id, p3, p4, side, handed }` (e.g. DW `3dr` = a
3-drawer variant). `SIZES`/`defSizes` give default W/H/D per `zone.fam`.

---

## 4. Construction pipeline (how a cabinet is built)

Entry: **`buildModel(...)`** orchestrates per placed line. Core chain:

```
buildModel
 └─ buildCarcassInner            (→ mergeCarcassPanels to combine duplicate panels)
     └─ buildCarcassInnerRaw     (5 branches by zone/family archetype)
         ├─ Tall      branch
         ├─ Loft      branch
         ├─ MD        branch
         ├─ Wall      branch
         └─ Base-default branch  ← cabinet code override + drawer entry point
     ├─ addElenor(...)           (adds ELEN profile where applicable)
     └─ addDrawerBoxes(...)      (drawer packs — see §6)
 └─ buildShutters(...)           (doors/drawer fronts appended; uses shutSpec + validDesigns)
```

### Cabinet code
Built as:
```
code = [p1, p2, p3, p4, handle, W, H, D, t].join("-")  + (mat && mat!=="(none)" ? "-"+mat : "")
```
**Override (base-default zone, in `buildCarcassInnerRaw`):** for the inbuilt 2HB+1BL case,
P3·P4 become `2HB-1BL`:
```ts
const isInbuilt2hb1bl = z.p1 === "BC" && fk === "DW" && v.id === "3dr" && inbuiltDrawers === "2hb1bl";
const codeP3 = isInbuilt2hb1bl ? "2HB" : v.p3;
const codeP4 = isInbuilt2hb1bl ? "1BL" : p4;
const code = [z.p1, f.p2, codeP3, codeP4, handle, W, H, D, t].join("-") + (mat && mat!=="(none)" ? "-"+mat : "");
```

### Pack rows
`pkRows` entries are `[type, name, dim, qty, uom]`. A pack groups its child panels/profiles/
hardware. Drawer packs use `["panel", drawerPackName, "Set", n, "Set"]` where `n` = drawer count.

### shutSpec / buildShutters
`shutSpec(zoneKey, familyKey, variant, inbuiltDrawers)` plans the door/drawer fronts (kind,
count, hb codes). `buildShutters` produces `Shutter` records with profile V/H, hinge, panel
sizes; `validDesigns` constrains allowed front designs.

---

## 5. Naming convention & BOM tree levels

`src/lib/naming.ts` parses/normalizes BOM item names. The tree nesting is:

```
L0  Master composite (the cabinet)
L1  Pack            — "Set of Parts- ..." OR "Drawer Pack- ..."
L2  Part / Panel / Profile / Hardware
L3  Panel / Raw
L4  Raw material
```
Helpers: `explodePanelForTree`, `getPartBaseName`, `getPanelBaseName`,
`normalizePartOrPanelName`, `splitDrillAndFinish`.

- **Drilled parts** nest `Part (JD) → Panel → Raw`.
- **Flat panels** skip the Part wrapper: `Panel → Raw`.

---

## 6. Drawer packs (the most recently restructured subsystem)

> Drawer BOM was collapsed to a SINGLE Level-1 pack per variant. See `DRAWER_PACK_HANDOFF.md`
> for the deep dive; summary here.

**`Drawer Pack- Cab Drawer Box <Low|High|Semi High>`** — old `Set of Parts- Cab Drawer Box`
wrapper eliminated. Inside each pack:
- **Back panel:** drilled → `Part (JD) → Panel → Raw`.
- **Bottom panel:** flat (no Part wrapper, no drill).
- **Fascia** (built-in drawers only): drilled JD, carcass colour → `Part (JD) → Panel → Raw`.
- **Side profile HM513 (DBS):** KEEPS `Set of Profile Parts (LH/RH)` wrapper; LH/RH split
  fires only when per-entry `qty === 2`.
- **Center profile HM535 (DBC):** FLAT cut profile (raw + cutting), like the bottom panel.
- **Runner hardware:** nested FLAT under the pack at **Level 2**.

**Composite-multiplier quantity model:** pack L1 qty = drawer count **N**. Child arrays store
**TOTAL** quantities (per-drawer × N). At render/Zoho, children are divided by N to show
per-drawer consumption; profile explosion passes `q × N` so exploded total = per-drawer × N ×
cabinetQty. `getDrawerCount` derives N from the back-panel qty within the pack.

Applies consistently in: **BOM Output, Excel/Full-BOM tree (`treeGroups`), Stock Check, and
Zoho Sales Order sync.**

### Inbuilt dropdown (DW/HO "Base Drawer", 3-drawer variant only)
Options are exactly **`None`** and **`2HB + 1BL`** (`"2hb1bl"`).
`hasInbuiltDrawerOption = (DW||HO) && variant.id === "3dr"`.
(Forced internal branches remain for GD/WBP→`1lb`, DPN→`fixed_dpn`.)

**2HB+1BL behavior:** replaces the default 2LB+1HB layout with **3 built-in (fascia) drawers**:
2× High Back + 1× Low Back, all with fascia. Standard (no-fascia) box generation is
**suppressed**.
- High Back built-in: back 212, fascia 210×562, hardware H239 → `Drawer Pack ... High` (qty 2).
- Low  Back built-in: back  63, fascia 110×562, hardware H90  → `Drawer Pack ... Low`  (qty 1).
- Fascia width: `faciaW = W − 2·t − 8`, which at locked `t=15` equals `W − 38` (600→562).

### Dimension formulas (emit loop)
```
backW   = W - 72   (t=15, drill JD)
bottomW = W - 53   bottomD = D - 77   (t=6, flat)
faciaW  = W - 2*t - 8
DBS profile: len 483, qty 2*n
DBC profile: len W - 78, qty 4*n
```

---

## 7. BOM explosion — `buildFullBomData`

Flattens a `CarcassModel` into display/export rows (`FullBomRow`). Row builders:
- `makeRow`, `addPanelRow`, `addProfilePanelAndRawRow` (waste 20% ELEN / 25% other),
  `addHardwareRow`, `buildProfileBomRows`.
- `actualQty = qty × q` (q = cabinet quantity in the project line).
- **Drawer branch** (guard `pkType==="panel" && name.startsWith("Drawer Pack- Cab Drawer Box")`
  then `return`): `dCount=getDrawerCount`, `qN=q*dCount`; per-parent qty `p.qty/dCount`,
  actualQty `p.qty*q`; DBC via `addProfilePanelAndRawRow(...qN...)` flat; DBS via
  `buildProfileBomRows(...[{len,qty:perDrawerQty,name}]...qN...)` (keeps LH/RH set); hardware
  via `addHardwareRow` at L2 (`h.qty/dCount`).
- **`isMatchedHardware`** returns `true` for any `h.pack` starting `"Drawer Pack- Cab Drawer
  Box"` so drawer hardware is nested-only, never duplicated as flat L1.
- **Stock:** `runStockCheck` + `stockMap` annotate rows with on-hand availability.

### Tree grouping — `treeGroups` useMemo
`nonDrawerPanels` excludes packs starting `Drawer Pack- Cab Drawer Box`/`... Pullout Shelf`
plus legacy `Set of Parts-` prefixes; a loop adds one group per `pkRow` starting `Drawer Pack-`
/`Set of Parts-`.

---

## 8. Zoho Inventory sync — `handleAddToZohoSO`

**Bottom-up composite creation**, idempotent by item name. New raw-bearing composites get an
SO-number prefix (dedup mechanism). Order: leaf simple items → composite parts → composite
packs → master, then PUT to the Sales Order `line_items`.

Resolvers:
- `resolveSimpleItem`, `resolveCompositeItem`, `findExistingCompositeByRawMaterial`,
  `resolveProfileBOM`, `resolveHardwarePackComposite`, `resolveHardwarePack`.
- **Service items:** `Cutting-1`, `Drilling-1`, `Packing Item-Carcass`, `Packing Item-Shutter`,
  `Hardware`.
- **Stone qty** = area × 1.15 / slabArea (4 decimals).

**Drawer branch:** `drawerCount=getDrawerCount`; panels resolve stone+cutting composite;
back/fascia wrapped in drilled `Part`, pushed `quantity: p.qty/drawerCount`; DBC → flat
cut-profile composite (`group:"Profile"`, name `${p.name} ${p.len}mm ${pFinish}`, raw+cutting)
pushed `quantity: p.qty/drawerCount`; DBS → `resolveProfileBOM(...qty: p.qty/drawerCount...)`;
hardware via `resolveHardwarePackComposite`/`resolveHardwarePack`. Pack composite pushed at end
with `quantity: drawerCount`.

### `src/lib/zoho.ts`
OAuth refresh-token flow. Env: `ZOHO_ORGANIZATION_ID`, `ZOHO_REFRESH_TOKEN`, `ZOHO_CLIENT_ID`,
`ZOHO_CLIENT_SECRET`. Defaults to **India DC**: `accounts.zoho.in` + `zohoapis.in/inventory/v1`.
In-memory token cache. Exports `zohoInventoryGet/Put/Post`.

---

## 9. API routes (`src/app/api/...`)
- `zoho/items`, `zoho/items/[id]/vendors` — item + vendor lookup.
- `zoho/compositeitems` — create/find composite items.
- `zoho/salesorders` — list/get/update SO line items.
- `zoho/contacts` — customer lookup.
- `zoho/purchaseorders`, `zoho/po-vendor-map` — PO creation + vendor mapping.
- `zoho/reorder-report` — low-stock/reorder aggregation.
- `zoho/status` — connectivity/health.
- `ai/chat` — OpenAI `gpt-4o-mini` copilot backend.

---

## 10. Other components & libs
- **BomDashboard** (~2,300) — browse Sales Orders, view their BOM trees, stock check, export.
- **BomOrderCard** — per-SO summary card.
- **AuthGate** — password gate (`Factory@1234`, localStorage).
- **AiCopilot** — chat UI hitting `/api/ai/chat`.
- **ReorderReport** — reorder/low-stock view.
- **DraftPoModal** — draft purchase-order creation UI.
- **QrBomPage** — read-only mobile BOM by SO (route `/qr/bom/[orderId]`).
- **libs:** `export.ts` (Excel/CSV), `stock.ts` (stock check), `naming.ts` (name parsing),
  `report.ts` (reorder aggregation), `rawmaterial.ts` (raw resolution), `types.ts` (interfaces).

---

## 10b. Elevation & "Other Accessories" (added feature)

### Elevation (mandatory at cabinet add)
- Constant `ELEVATION_OPTIONS = [AA, BB, CC, DD, EE, FF, GG, HH, II, JJ, KK]` (near the
  accessory types, ~line 131).
- State `elevation` (default `""`). `ProjectLine` carries `elevation?: string`.
- The **Add to Project** row renders an Elevation `<select>` (red border until chosen); the
  `+ Add to Project` button is **disabled** until an elevation is picked. `addLine` also
  guards (`if (!elevation) { alert(...); return; }`) and pushes `elevation` onto the line.
- Shown per line in the project table (under the code) as `Elevation: XX`.

### Other Accessories (panels) — new section before ⑨ Accessories
- Types: `OtherAccItemKey = "chimney" | "dishwasher"`; `OTHER_ACC_ITEMS = [Chimney Panel,
  Dishwasher Panel]`; `OtherAccRow = { id, item, width, height, profile, color, qty }`.
- State `otherAccRows` / `otherAccCounter`; helpers `addOtherAccRow` (defaults
  `profile = shutterProfileColor`, `color = shutterMat`, `width/height = customW/customH`),
  `updateOtherAccRow`, `removeOtherAccRow`.
- **Color** dropdown = stone finishes (`shutterOptions`, i.e. `finishesMap["6"]`); default =
  shutter stone (`shutterMat`). Changing it selects a different stone for that panel.
- **Profile** dropdown = union of all `profile_finishes.json` values (`otherAccProfileOptions`);
  default = shutter profile colour (`shutterProfileColor`).
- UI: a "Other Accessories" card (with `+ add row`) rendered immediately **before** the
  `⑨ Accessories` card; each row = Item / Width / Height / Profile / Color / Qty / remove.
- Export: `buildAccessoryExportRows` appends one row per Other Accessory
  (`Accessory: "Other Accessory"`, `Item Name: "<label> - <color> (<profile>)"`,
  `Size: "<w>x<h>"`, `Actual Qty: qty`).
- **Raw-material requirement:** `rawAggregation` adds a `shutter|<color>|6` stone key per row
  (netSqft = `sqft(w,h)·qty`), so the chosen colour appears in the raw-material selection UI and
  the user picks a Zoho raw stone (populates `selectedRawIds`).
- **Zoho SO push (`handleAddToZohoSO`, "Step 2b"):** each Other Accessory becomes its own SO
  line. Stone resolved via `selectedRawIds["shutter|<color>|6"]` (throws a clear error if not
  selected); stone qty = `sqft(w,h)·(1+STONE_WASTE)/slabArea` (4dp). A composite
  `"<label> <w>x<h>x6 <color> (<profile>)"` is created with `[stone, Cutting-1]` (custom-field
  group `"Other Accessory"`), pushed at `quantity: row.qty`, `rate: 0`, with line
  `description: "Profile: <profile>"`. (Profile colour is recorded in the composite name +
  line description; it is NOT pushed as a separate raw aluminium-profile line since these
  panels have no profile *design* code.)

### Elevation on Zoho
- Each cabinet SO line now carries `description: "Elevation: <XX>"` (from the per-line
  `elevation`). Wired through `itemsToAdd → newCabinetLines.description → linesToAdd`.

## 11. Key facts / gotchas
- **Pullout Shelf packs are vestigial** — never generated; legacy two-pack defensive branches
  left intact (`getDrawerCount` handles via name-rewrite).
- Line numbers in handoff docs are **approximate** (file shifts on edit) — search by anchor
  strings.
- Drawer subsystem is fully covered by `scratch/test-drawers.js` (scenarios A–D + Architecture
  + PPN). Always run it after touching drawer logic.
- Fascia width formula left as `W − 2·t − 8` (== `W − 38` at t=15) for code consistency.

---

## 12. Where to start for common tasks
- **New family/variant:** add to the FAMILIES table + `SIZES`/`defSizes`, then handle in the
  matching `buildCarcassInnerRaw` branch.
- **Change a cut formula:** find it in `buildCarcassInnerRaw` (panels) or `addDrawerBoxes`
  (drawers); re-run tsc + test-drawers.
- **Change BOM display/explosion:** `buildFullBomData` + `treeGroups`.
- **Change Zoho push:** `handleAddToZohoSO` + resolvers; `src/lib/zoho.ts` for transport.
- **Drawer changes:** read `DRAWER_PACK_HANDOFF.md` first, then §6 above.
