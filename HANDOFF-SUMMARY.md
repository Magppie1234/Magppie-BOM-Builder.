# BOM Backorder Report App — Handoff Summary

## Project Overview
A Next.js 16 App Router application connected to Zoho Inventory via REST API. It generates BOM (Bill of Materials) Backorder Reports, Packing Lists, Labels (PDF with barcodes/QR), Excel exports, and an Accessories Subform for Sales Orders.

**Stack**: Next.js 16, TypeScript, Zoho Inventory API, jsPDF, JsBarcode, qrcode, xlsx

**Working Directory**: `/Users/apple/Downloads/magppie_handoff_render/Inventory/bom-nextjs-app`
**Dev Server**: `npm run dev` on port 3002

---

## Key Files

| File | Purpose |
|------|---------|
| `src/components/BomDashboard.tsx` | Main UI component. Contains all packing list, label PDF, accessory, QR modal, and dashboard logic |
| `src/components/BomOrderCard.tsx` | Renders a single BOM order card with tree-structured rows |
| `src/components/QrBomPage.tsx` | QR code page for viewing a single order's BOM |
| `src/lib/stock.ts` | BOM expansion engine: expands composites into rows, calculates stock/deficit, assigns groups |
| `src/lib/types.ts` | All TypeScript types (BomReportRow, SalesOrderDetail, ItemDetail, etc.) |
| `src/lib/report.ts` | Fetches orders/items from API, resolves composite tree, builds report |
| `src/lib/export.ts` | Excel/CSV export: Full BOM sheet, Out of Stock sheet (now with Pcs logic), Opti sheet, Accessories sheet |
| `src/lib/zoho.ts` | Zoho API token management and helpers |
| `src/app/globals.css` | All styles including packing modal, dashboard, BOM cards |

---

## BOM Data Model

### Hierarchy
```
Sales Order
  └─ Line Item (level 0, rowType: "master" or "plain")
       └─ Child (level 1, rowType: "component" or "sub_bom")
            └─ Grandchild (level 2+, rowType: "component" or "sub_bom")
```

### Key Fields on BomReportRow
- `masterGroup`: parent's cfGroup (used for group filtering in packing list)
- `cfGroup`: item's own group from Zoho `cf_group` field
- `cfSubGroup`: item's own sub-group from Zoho `cf_sub_group` field
- `cfFinish`: item color/finish from Zoho
- `parentItemId`: points to the parent composite's itemId
- `quantityNeeded`: `perUnit × parentQty` (BOM quantity, includes SO qty multiplier)
- `actualQuantity`: `quantityNeeded + waste%` (for stock calculations)
- `rowType`: "master" | "plain" | "component" | "sub_bom"
- `level`: depth in BOM tree (0 = SO line item, 1 = direct child, etc.)

---

## Packing List Logic (in BomDashboard.tsx)

### Pipeline
1. **`isPackingListRow(row)`** — Includes `plain` items OR `component`/`sub_bom` at level 1 only
2. **`filterPackingRowsByGroup(orderRows, group)`** — Filters by masterGroup matching the selected group (e.g., "Carcass")
3. **`consolidatePackingRows(rows, allOrderRows)`** — Three-way consolidation:
   - **Hardware `sub_bom`** items: merged by item name, PCS = merged qty (GLUE: qty÷270)
   - **Non-hardware `sub_bom`** items: NOT merged (each stays as own row), boxes = parentSoQty, PCS = set multiplier
   - **`plain`/`component`** items: merged by itemId, boxes = parentSoQty, PCS = perUnit × set multiplier
4. **`groupByParentMaster(consolidated, allOrderRows)`** — Groups items by parentItemId, resolves master name
5. **`buildPackingListHtml(order, orderRows, group)`** — Generates HTML packing list

### Set Item Logic (`isSetItem`)
- **Carcass Pack** (cfSubGroup OR cfGroup contains "carcass pack") AND SKU does NOT contain "prof" → PCS = 2
- **Profile Pack** (cfSubGroup OR cfGroup contains "profile pack") → PCS = 2 always
- Everything else → PCS = 1

### Hardware Section
- Hardware items identified by: name starts with "hardware pack" OR name contains "glue silicone"
- Same-name hardware packs are merged, quantities summed
- GLUE SILICONE: ML ÷ 270 → pcs
- All hardware items share 1 PACK/BOX (using `rowspan` for merge+center in HTML)
- Each hardware item listed as separate row with its own PCS

### Box Dimension Calculation (`calcBoxDimension`)
- Extracts WxHxT from item name (regex: `(\d+)x(\d+)x(\d+)`)
- Adds 20mm padding to each dimension
- For set items: thickness = T×2 + 20 (double thickness for 2 pieces)
- For non-set: thickness = T + 20

### Parent SO Qty Resolution
- `masterQtyMap` built from level-0 rows (master/plain) mapping itemId → quantityNeeded
- Each child's `parentSoQty` = parent's quantityNeeded from masterQtyMap
- boxes = parentSoQty (not child's quantityNeeded which includes perUnit multiplier)

---

## Label PDF Logic (`buildLabelsPdf`)
- Uses jsPDF, landscape 150×100mm per label
- One label per box (iterates `entry.boxes` times per consolidated row)
- Label contains: MAGPPIE branding, item name, From/To addresses, barcode (SKU via CODE128), QR code (SO number + customer name), box number, PCS per box, finish, dimension

---

## Excel Export (export.ts)

### Sheets
1. **Full BOM** — All rows with all fields (now contains a **Pcs** column for stone raw materials, positioned next to Actual Qty).
2. **Out of Stock** — Consolidated deficit items (now contains a consolidated **Pcs** column).
3. **Opti** — Cutting optimization sheet with **Profile Code** column (extracts 3 chars before last `-` in item name)
4. **Accessories** — From accessory subform (only enabled items exported)

### Stone Pieces Logic (Pcs)
- Matches raw materials where name, cfGroup, or masterGroup contains `"stone"` (case-insensitive).
- Uses a regex parser `/(\d{3,4})\s*[xX]\s*(\d{3,4})/` to extract height and width from the item name.
- Calculates single slab area: `Slab Area (sqft) = (height * width) / (304.8 * 304.8)`.
- Calculates number of slab pieces required: `Pcs = Actual Qty / Slab Area`, rounded to 2 decimal places. Non-stone rows display `"-"`.
- Recalculates pieces count dynamically during OOS consolidation using `Consolidated Actual Qty / Slab Area`.

### Profile Code (`extractProfileCode`)
```typescript
function extractProfileCode(itemName: string): string {
  const lastDash = itemName.lastIndexOf("-");
  if (lastDash < 3) return "";
  return itemName.substring(lastDash - 3, lastDash);
}
```

---

## Accessories Subform
- 5 accessory types: skirting, duplay, lprofile, jhandle, chandle
- Profile mutual exclusion: skirting/duplay/lprofile — only 1 at a time
- Handles: NO mutual exclusion (both jhandle + chandle allowed)
- No duplicate accessories per item
- Only enabled items exported to Excel

---

## Recent Changes (this session)
1. **Raw Material Selection Popup Fixes**:
   - Reverted panel sqft calculation to `panel.quantity` (net area in sqft) since Zoho already stores it in sqft, fixing the squaring error.
   - Automatically searches for default stone from Zoho using `finish + height(2800) + width(1200) + group(stone) + thickness`.
   - Loosened group name filters to check if `group_name` contains "stone" (`.includes("stone")`), preventing matching candidates from being filtered out.
   - Excluded the default stone from the candidates selection dropdown.
   - Fixed Zoho duplicate component replacement error (Zoho code `2057`) by immediately setting `stoneReplaced = true` in tree traversal on first matched stone.
2. **Excel Export Pieces Column**:
   - Implemented a `Pcs` column in the `Full BOM` and `Out of Stock` sheets in `export.ts` based on regex parsing of name dimensions.
   - Configured custom column widths for `Full BOM` and `Out of Stock` sheets to align headers properly.

---

## Known Behaviors
- `quantityNeeded` at level 1 = `perUnit × parentQty` (already includes SO qty)
- `actualQuantity` = `quantityNeeded + waste%` (used for stock, NOT for packing)
- GLUE SILICONE is classified as hardware/other but measured in ML, converted to pcs (÷270)
- Profile mutual exclusion only applies to skirting/duplay/lprofile, NOT handles
