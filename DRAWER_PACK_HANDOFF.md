# HANDOFF — Magppie Carcass BOM Builder: Drawer-Pack Restructure

This document plus the bundled source lets another AI continue from the current state.

## 1. Project & environment

- **App:** Next.js 16 (React 19, TypeScript, webpack + wasm SWC, Node 24) — "Carcass BOM Builder" integrating with **Zoho Inventory (India DC)**.
- **Active codebase (work ONLY here):** `~/Documents/bom-nextjs-app/Inventory - Magppie`
  (NOT the `~/Downloads` copies, NOT the parent `~/Documents/bom-nextjs-app` — the parent is a *different* app titled "BOM Backorder Report" with only `/qr`).
- **Heart of the app:** `src/components/CarcassBomBuilder.tsx` (~6,900 lines) — all BOM logic + Zoho creation flow. Main route `/builder` (`src/app/builder/page.tsx` just renders `<CarcassBomBuilder/>`).
- **Verify after every edit (REQUIRED):**
  ```bash
  cd "/Users/Apple/Documents/bom-nextjs-app/Inventory - Magppie"
  node scratch/test-drawers.js      # drawer unit/architecture tests
  npx tsc --noEmit                  # type check (expect exit 0)
  ```
- **Dev-server gotcha:** the Claude preview MCP (`preview_start`) launches the WRONG parent app, so `/builder` 404s. To run the real server: `npm run dev` (port 3000) **inside the `Inventory - Magppie` folder**. Logic is covered by `scratch/test-drawers.js`; browser preview was NOT used for this work.
- Env in `.env.local`: `ZOHO_*` creds (India), `OPENAI_API_KEY`. Not a git repo.

## 2. Spec implemented (current state — ALL DONE & VERIFIED)

### A. Drawer BOM collapsed to a SINGLE Level-1 pack per variant
`Drawer Pack- Cab Drawer Box <Low|High|Semi High>` — the old `Set of Parts- Cab Drawer Box` wrapper was **eliminated**. Inside each pack:
- **Back panel:** drilled → `Part (JD) → Panel → Raw`.
- **Bottom panel:** flat (no Part wrapper, no drill).
- **Fascia panel** (built-in drawers only): drilled `JD`, carcass colour → `Part (JD) → Panel → Raw`.
- **Side profile HM513 (`DBS`):** KEEPS its `Set of Profile Parts (LH/RH)` wrapper (LH/RH drilling). LH/RH split fires only when the per-entry `qty === 2`.
- **Center profile HM535 (`DBC`):** FLAT cut profile (raw + cutting), like the bottom panel — no set wrapper.
- **Runner hardware:** nested FLAT under the pack at **Level 2** (NOT a flat Level-1 carcass-hardware row).
- **Quantity convention (composite-multiplier model):** the pack's Level-1 qty = **drawer count N**. Child arrays store **TOTAL** quantities (per-drawer × N). At render/Zoho, children are divided by N to show per-drawer consumption; profile explosion passes `q × N` so exploded total = per-drawer × N × cabinetQty. `getDrawerCount` derives N from the back-panel qty within the pack.
- Applies consistently in: **BOM Output, Excel/Full-BOM tree (`treeGroups`), Stock Check, and Zoho Sales Order sync.**

### B. Inbuilt dropdown (DW/HO "Base Drawer", 3-drawer variant only)
Options are now exactly **`None`** and **`2HB + 1BL`** (value `"2hb1bl"`). Old `1lb/1sb/2lb/2sb` removed from the UI (their internal `effectiveInbuiltDrawers` branches remain for GD/WBP→`1lb`, DPN→`fixed_dpn`, which are forced, not user-selectable). `hasInbuiltDrawerOption = (DW||HO) && variant.id === "3dr"`.

### C. "2HB + 1BL" behavior (was corrected mid-task — read carefully)
- Originally said to change ONLY the cabinet code. **Corrected:** it now ALSO changes the drawer pack.
- Selecting `2hb1bl` **replaces** the default 2LB+1HB layout with **3 built-in (fascia) drawers**: **2 × High Back + 1 × Low Back**, all with fascia.
  - High Back built-in: back 212, **fascia 210 × 562**, hardware `H239` → `Drawer Pack- Cab Drawer Box High` (qty 2).
  - Low Back built-in: back 63, **fascia 110 × 562**, hardware `H90` → `Drawer Pack- Cab Drawer Box Low` (qty 1).
  - Standard (no-fascia) box generation is **suppressed** for this option.
- **Fascia width:** user formula is `Outer − 38`. Code uses `faciaW = W − 2·t − 8`, which at stone-locked `t = 15` equals `W − 38` (600 → 562). Identical — left as-is for consistency.
- **Cabinet code:** when `z.p1 === "BC" && fk === "DW" && v.id === "3dr" && inbuiltDrawers === "2hb1bl"`, the code's P3·P4 becomes `2HB-1BL` (default `2LB-1HB`). ONLY this case.

### D. UI placement
**Drawer Model** + **Inbuilt Drawers** dropdowns now render immediately **after the Configuration (P3·P4)** `SearchableSelect`.

### Key facts
- `NO_DRILL_PROFILES = {STP, ELEN, SLF, SINK}`. `DBS`/`DBC` are NOT in it (drilled-class).
- **Pullout Shelf packs are vestigial / never generated** — defensive branches left untouched (legacy two-pack shape; `getDrawerCount` handles via a name-rewrite).
- Constants: `STONE_WASTE = 0.15`, `PROFILE_WASTE = 0.20`.

## 3. Exact change locations in `src/components/CarcassBomBuilder.tsx`

Line numbers approximate (file shifts as you edit). Search by anchor strings.

1. **`addDrawerBoxes`** (≈550–707) — single-pack aggregation; the `bump()` aggregator; standard section guarded by `effectiveInbuiltDrawers !== "2hb1bl"`; built-in section has `numHigh/numLow/numSemi` incl. the `2hb1bl → numHigh=2, numLow=1` branch; emit loop pushes back/bottom/fascia panels, DBS+DBC profiles, hardware, and the pkRow `["panel", drawerPackName, "Set", n, "Set"]`.
2. **`getDrawerCount`** (≈2481) — returns back-panel qty (N); rewrites name for Pullout.
3. **`buildFullBomData` drawer branch** (≈2584) — guard `if (pkType === "panel" && name.startsWith("Drawer Pack- Cab Drawer Box"))`, then `return`s (skips generic handling). `dCount = getDrawerCount(...)`, `qN = q * dCount`. Back/fascia → drilled `Part` (per-parent qty `p.qty/dCount`, actualQty `p.qty*q`); bottom → flat panel; `DBC` → `addProfilePanelAndRawRow(... perDrawerQty ... qN ...)` flat; `DBS` → `buildProfileBomRows(... [{len,qty:perDrawerQty,name}] ... qN ...)` (keeps LH/RH set); hardware via `addHardwareRow` at Level 2 (`h.qty/dCount`).
4. **`isMatchedHardware`** (≈2877) — returns `true` for any `h.pack` starting `"Drawer Pack- Cab Drawer Box"` so drawer hardware is nested-only, never duplicated as flat Level-1.
5. **Zoho `handleAddToZohoSO` drawer-pack branch** (≈5099) — `if (packName.startsWith("Drawer Pack- Cab Drawer Box") || ... "Drawer Pack- Cab Pullout Shelf")`. `drawerCount = getDrawerCount(...)`. Panels resolve stone+cutting composite; back/fascia wrapped in drilled `Part`; pushed `quantity: p.qty / drawerCount`. Profiles grouped by type: `DBC` → flat cut-profile composite (`group: "Profile"`, name `${p.name} ${p.len}mm ${pFinish}`, raw+cutting) pushed `quantity: p.qty/drawerCount`; others (`DBS`) → `resolveProfileBOM(... qty: p.qty/drawerCount ...)`. Hardware via `resolveHardwarePackComposite`/`resolveHardwarePack`. Pack composite pushed at end with `quantity: drawerCount`.
6. **`treeGroups` useMemo** (≈6772) — `nonDrawerPanels` now excludes packs starting `"Drawer Pack- Cab Drawer Box"`/`"Drawer Pack- Cab Pullout Shelf"` (plus legacy `Set of Parts-` prefixes); a loop adds one group per `pkRow` starting `"Drawer Pack-"`/`"Set of Parts-"`.
7. **Cabinet code override** (≈1625, in `buildCarcassInnerRaw`, base-default zone): right before `const code = [...].join("-")`:
   ```ts
   const isInbuilt2hb1bl = z.p1 === "BC" && fk === "DW" && v.id === "3dr" && inbuiltDrawers === "2hb1bl";
   const codeP3 = isInbuilt2hb1bl ? "2HB" : v.p3;
   const codeP4 = isInbuilt2hb1bl ? "1BL" : p4;
   const code = [z.p1, f.p2, codeP3, codeP4, handle, W, H, D, t].join("-") + (mat && mat !== "(none)" ? "-" + mat : "");
   ```
8. **Inbuilt dropdown options** (UI, in the Configure card) — `[{ value: "none", label: "None" }, { value: "2hb1bl", label: "2HB + 1BL" }]`.
9. **UI dropdown move** — `{hasDrawers && (...Drawer Model...)}` and `{hasInbuiltDrawerOption && (...Inbuilt Drawers...)}` now sit right after the Configuration `SearchableSelect` (they were after the thickness input).

## 4. Test coverage (`scratch/test-drawers.js`)

Self-contained Node script mirroring the component logic. Scenarios:
- **A** DW-3dr `none` → Drawer Pack Low (n=2) + High (n=1); totals DBS/DBC = per-drawer × n; H90 qty2, H239 qty1; no Set of Parts.
- **B** GD → standard High + 1 LB inbuilt (fascia 110×562 under Low pack; High pack has no fascia).
- **C** DPN `fixed_dpn` → Low (n=2) + Semi High (n=3) with fascias 110/210; H90 qty2, H175 qty3.
- **D** DW-3dr `2hb1bl` → High (n=2, fascia 210×562) + Low (n=1, fascia 110×562); H239 qty2, H90 qty1; no standard boxes.
- **Architecture** BOM tree: one Drawer Pack per variant at L1 with qty = drawer count; back as drilled Part, bottom flat, DBS set, DBC flat, hardware nested L2 (no flat L1 drawer hardware). Zoho: pack composites pushed with qty = drawer count; no Set-of-Parts composite.
- **PPN** Pullout Shelf legacy two-pack path still passes (vestigial).

Run: `node scratch/test-drawers.js` → ends with `ALL TESTS PASSED SUCCESSFULLY!`

## 5. Current status

All 10 tasks complete and verified: `node scratch/test-drawers.js` passes; `npx tsc --noEmit` exits 0. No known open issues. Browser preview not exercised (dev-server gotcha) — start `npm run dev` inside the folder for a live `/builder` check if needed.

## 6. Bundled files in this zip
- `DRAWER_PACK_HANDOFF.md` — this document.
- `src/components/CarcassBomBuilder.tsx` — full component (current state).
- `src/app/builder/page.tsx` — route.
- `scratch/test-drawers.js` — full test suite (current state).
- `CONTINUATION_README.md` — original project brief.
