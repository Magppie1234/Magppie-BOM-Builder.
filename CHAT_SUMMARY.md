# BOM Backorder Report — Raw Material Selection Feature — Full Chat Summary

## Project Overview
Next.js 16 App Router app connected to **Zoho Inventory** REST API. Generates BOM (Bill of Materials) backorder reports for sales orders. The major feature being built is a **Raw Material Selection Popup** that appears before BOM loads.

**Tech Stack:** Next.js 16, TypeScript, Zoho Inventory API, jsPDF, QRCode, JsBarcode

**Dev Server:** Runs on `localhost:3002`

---

## BOM Hierarchy (Critical to understand)
```
Sales Order (SO)
  └── Carcass/Shutter (SO line item, composite, group: "carcass"/"shutter")
        ├── Set of Parts 1 (composite)
        │   ├── Part 1 (composite)
        │   │   └── Panel 1 (composite, group: "Panels- SH Strip" / "Panels- SH" / "Panels- Cab")
        │   │       ├── Shade/Stone item (raw material, group: "shade"/"shades" — THIS gets replaced)
        │   │       └── Cutting-1 (service item — DO NOT replace)
        │   ├── Part 2
        │   │   └── Panel 2 → Shade item + Cutting-1
        │   └── ...
        ├── Set of Parts 2
        │   └── ... more Parts → Panels
        └── Hardware Pack (SKIP entirely)
```

### Key Data Points
- **Carcass thickness** (e.g., 15mm, 31mm) ≠ **Stone thickness** (e.g., 6mm)
- For **Carcass**: use the carcass item's own `cf_thickness` for stone matching
- For **Shutter**: use the Panel's thickness from BOM traversal (`chain.stoneThickness`)
- Default stone found by **Zoho search**: `group=Stone + finish + thickness + height=2800 + width=1200`
- Stone items in dropdown: `group=Stone + finish + thickness`
- The raw material INSIDE Panels has group **"shade"/"shades"**, NOT "stone"
- The "stone" group items are what you SEARCH for as replacements
- Only replace in composites whose group is: **"Panels- SH Strip"**, **"Panels- SH"**, **"Panels- Cab"**
- Items like "Cutting-1" (service) must NOT be replaced
- Custom fields accessed via `getCfValue(item, "cf_group")`, `getCfValue(item, "cf_finish")`, etc.
- Numeric custom field values come as strings like "2800.00" — use `parseFloat()` for comparison

---

## Two Scenarios for Stone Selection

### Scenario 2 (Default — checkbox unchecked)
- Default stone is auto-found by searching Zoho: `group=Stone + finish + thickness + height=2800 + width=1200`
- On confirm: `updateStonesInTree()` recursively walks the ENTIRE BOM tree from each carcass → all Sets → all Parts → all Panels
- At each Panel composite (group matches PANEL_GROUPS), replaces the first non-composite, non-hardware item with the default stone
- Uses **PUT** on existing composites — no new items created
- If default stone not found: checkbox is auto-checked (mandatory), dropdown auto-loads

### Scenario 1 (User changes stone — checkbox checked)
- User selects from dropdown filtered by `group=Stone + finish + thickness`
- On confirm: `cloneCompositeTree()` recursively clones the ENTIRE BOM tree top-down
- Each composite gets `-{SO number}` appended to name and SKU
- At Panel level, replaces the shade item with the user-selected stone
- Checks for existing clones (by name) to avoid duplicates
- Updates SO to point to new cloned carcass items
- All tax fields copied from original items (Zoho requires `tax_id` or `tax_exemption_id`)

---

## Files Modified/Created

### 1. `/src/lib/rawmaterial.ts` (NEW — core raw material logic)
**Full current code is in the file.** Key functions:

- **Types**: `StoneGroup`, `ProfileGroup`, `BomChainItem`, `RawMaterialSelection`, `TraversalResult`
- **`getCarcassCategory(item)`** — returns "carcass", "shutter", or null
- **`isPanelGroupMatch(item)`** — checks if composite's group is "Panels- SH Strip", "Panels- SH", or "Panels- Cab"
- **`isHardwareItem()`**, **`isHardwareComposite()`** — filters to skip hardware
- **`findStoneInBomChain()`** — traverses BOM tree to find stone at Panel level (for initial popup data)
- **`findDefaultStone(finish, thickness)`** — searches Zoho for default stone: group=Stone + finish + thickness + 2800x1200
- **`resolveRawMaterials(orders, ...)`** — main function that groups carcass/shutter items, finds defaults, returns `RawMaterialSelection`
- **`searchStoneItems(finish, thickness)`** — dropdown search for Scenario 1
- **`updateStonesInTree(compositeId, defaultStoneId, ...)`** — Scenario 2: recursively PUT-updates all Panels
- **`cloneCompositeTree(compositeId, soNumber, newStoneId, ...)`** — Scenario 1: recursively clones entire tree
- **`applyStoneChange(group, soNumber, orderId, order)`** — dispatches to Scenario 1 or 2
- **`applyProfileChange(group, ...)`** — same for profiles
- **`createComposite(name, sku, original, mappedItems)`** — creates composite in Zoho with tax fields
- **`findExistingComposite(name)`** — duplicity check by name
- **`buildSoLineItem(li)`** — builds SO line item payload preserving all tax fields
- **`copyTaxFields(source, target)`** — copies tax fields between objects

### 2. `/src/components/BomDashboard.tsx` (~2100+ lines)
Key additions for raw material feature:

**State variables** (line ~581):
```typescript
const [rawMatModal, setRawMatModal] = useState(false);
const [rawMatSelection, setRawMatSelection] = useState<RawMaterialSelection | null>(null);
const [rawMatLoading, setRawMatLoading] = useState(false);
const [rawMatDropdownItems, setRawMatDropdownItems] = useState<Record<string, Array<{id,name,sku}>>>({});
const [rawMatDropdownLoading, setRawMatDropdownLoading] = useState<Record<string, boolean>>({});
const [rawMatOrders, setRawMatOrders] = useState<SalesOrderDetail[]>([]);
```

**`loadSelectedOrders()`** (line ~775):
- Fetches sales order details via `/api/zoho/salesorders/${id}`
- Calls `resolveRawMaterials()` to find stone/profile groups
- Shows modal if groups found, stores orders in `rawMatOrders` (avoids refetch)
- Auto-loads dropdown for groups where default not found (`changed=true`)
- Falls back to direct BOM load if order fetch fails

**`handleRawMatConfirm()`** (line ~840):
- Uses stored `rawMatOrders` (no refetch)
- Calls `applyStoneChange()` for ALL groups (both Scenario 1 and 2)
- Calls `applyProfileChange()` for changed profile groups
- Closes modal and loads BOM

**`handleStoneGroupToggle(groupKey)`** (line ~900):
- Toggles checkbox; if checked, fetches dropdown items via `searchStoneItems(finish, thickness)`
- Cannot uncheck if default stone not found (mandatory)

**`handleStoneSelect(groupKey, itemId, itemName)`** (line ~926):
- Sets `newStoneId` and `newStoneName` on the group

**Modal JSX** (line ~1960):
- Separate sections for Carcass and Shutter stones
- Each group shows: finish, thickness, total sqft
- Checkbox "Change stone" / "Change stone (required)"
- Default display or dropdown select
- Profile section below
- Confirm button disabled if any changed group has no selection

### 3. `/src/lib/zoho.ts`
- Added `zohoInventoryPost<T>()` for POST requests
- Existing: `zohoInventoryGet`, `zohoInventoryPut`, token management

### 4. `/src/app/api/zoho/items/route.ts` (NEW)
- GET: Search items (`?search=` or `?name=`) with `per_page: 200`
- POST: Create new items

### 5. `/src/app/api/zoho/items/[id]/route.ts`
- GET: Fetch single item
- PUT: Update item

### 6. `/src/app/api/zoho/compositeitems/route.ts` (NEW)
- POST: Create composite items

### 7. `/src/app/api/zoho/compositeitems/[id]/route.ts`
- GET: Fetch composite item with mapped items
- PUT: Update composite item (mapped items, etc.)

### 8. `/src/app/api/zoho/salesorders/[id]/route.ts`
- GET: Fetch sales order detail
- PUT: Update sales order (line items)

### 9. `/src/lib/export.ts`
- Component row color changed from `EEEBE5` to `D6EAF8` (distinct from master `FFF1DD`)

### 10. `/src/lib/stock.ts`
- Has `getCfValue(item, fieldName)` — resolves custom fields from `cf_*` direct props, `custom_field_hash`, or `custom_fields` array
- Has `isComposite(item)` — checks if item is composite type

---

## Other Fixes Applied

### Barcode Fix
- Increased resolution: width:3, height:80, fontSize:16, margin:4
- Guard for empty/"-" SKU: falls back to item name

### Labels Per Box (Hardware)
- Hardware: `boxes: 1, pcsPerBox: pcs` (1 label per box, not per PCS)
- Packing HTML: `entry.pcsPerBox` for PCS column, `entry.boxes` for BOX column

### Excel Export
- Component row color: `D6EAF8` (blue tint, distinct from master `FFF1DD` yellow)

---

## Known Issues / Pending Work

1. **Sqft calculation** — Currently uses Panel dimensions from BOM traversal, but only follows first path. May not capture all panels for accurate total sqft across multiple Sets/Parts.

2. **Profile raw material** — Logic exists but hasn't been fully tested with real data.

3. **Console.log debug statements** — Many `[RawMat]` prefixed logs throughout; should be removed after everything works.

4. **Zoho API rate limiting** — The app makes many sequential API calls during BOM traversal and cloning. Orders are cached in `rawMatOrders` state to avoid refetching, but the clone/update operations still make many calls.

5. **Error "duplicate components"** — Was caused by replacing ALL non-composite items in a Panel (including service items like "Cutting-1") with the stone. Fixed by:
   - Using `isPanelGroupMatch()` to only replace in Panel composites
   - Only replacing the FIRST non-composite, non-hardware item (using `stoneReplaced` flag)

6. **Tax fields** — Zoho requires `tax_id` or `tax_exemption_id` when creating composite items and updating SO line items. All tax fields are now copied from originals.

---

## Zoho Inventory API Patterns Used
```
GET  /items?search=...&per_page=200          → { items: ItemDetail[] }
GET  /items/{id}                              → { item: ItemDetail }
POST /items                                   → { item: ItemDetail }
GET  /compositeitems/{id}                     → { composite_item: CompositeItemDetail }
POST /compositeitems                          → { composite_item: CompositeItemDetail }
PUT  /compositeitems/{id}                     → { composite_item: CompositeItemDetail }
GET  /salesorders/{id}                        → { salesorder: SalesOrderDetail }
PUT  /salesorders/{id}                        → { salesorder: SalesOrderDetail }
```

All proxied through Next.js API routes at `/api/zoho/...`

Composite item creation requires: `item_type: "inventory"`, tax info, mapped_items array.
SO line item update requires: tax fields from original line item.

---

## Custom Fields Reference
| Field | Description | Example Values |
|-------|-------------|----------------|
| `cf_group` | Item group | "Stone", "Carcass", "Shutter", "Panels- SH", "Hardware" |
| `cf_sub_group` | Sub group | "carcass", "shutter", "hardware" |
| `cf_finish` | Finish/color | "ONYX MYSTIC", "ONYX OMAN" |
| `cf_thickness` | Thickness in mm | "6", "15", "31" |
| `cf_height` | Height in mm | "2800.00", "560" |
| `cf_width` | Width in mm | "1200.00", "690" |
| `cf_waste_percentage` | Waste % | "5", "10" |
| `group_name` | Alternative group field | "Stone", "Shade" |
