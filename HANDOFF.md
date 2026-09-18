# Kitchen Carcass + Shutter BoM Builder — Handoff Document

This document provides full context for continuing work on the **BOM Backorder Report & Carcass BOM Builder** Next.js application.

---

## 1. Project Overview & Architecture
This is a **Next.js 16 (App Router)** React & TypeScript application connected to Zoho Inventory. It contains two main features:
1. **BOM Backorder Report & Packing Lists** (mounted at the root `/`):
   * Searches sales orders, expands multi-level BOM trees, checks stock, handles group-based packing lists, prints barcode/QR labels (PDF), and exports data.
2. **Carcass & Shutter BOM Builder** (mounted at `/builder`):
   * An interactive stone-kitchen builder. You configure a cabinet (zone, family, dimensions, shutter design, finishes) and it calculates carcass stone panels, profiles, accessories, and shutter dimensions.
   * Consolidates configurations into a local session-only project list.
   * Creates corresponding Level 4 Panels, Level 3 Parts, Level 2 Packets, and Level 1 Cabinet composites inside Zoho Inventory and attaches them directly as line items on selected Sales Orders.

---

## 2. Key Technology Stack
* **Framework**: Next.js 16.2.6 (App Router)
* **Language**: TypeScript
* **Libraries**: jsPDF, JsBarcode, qrcode, xlsx, xlsx-js-style
* **Zoho Proxy Routes**: Server-side API endpoints under `src/app/api/zoho/` proxy all Zoho Inventory requests to bypass CORS restrictions.
* **Dev Server**: Runs on `http://localhost:3000` (PID `911`).

---

## 3. Directory Structure
```
bom-nextjs-app/
├── src/
│   ├── app/
│   │   ├── api/zoho/          # Zoho API proxy endpoints (items, salesorders, status, etc.)
│   │   ├── builder/page.tsx   # Renders the Carcass BOM Builder page
│   │   ├── page.tsx           # Renders the main BOM Backorder Report Dashboard
│   │   └── globals.css        # Core design system stylesheet
│   ├── components/
│   │   ├── CarcassBomBuilder.tsx  # Interactive builder, project lines, and Zoho integration
│   │   ├── BomDashboard.tsx       # Packing lists, accessories, report filters, and exports
│   │   └── BomOrderCard.tsx       # Individual Sales Order card renderer
│   ├── lib/
│   │   ├── zoho.ts            # OAuth token refresh, base request helpers (GET, PUT, POST)
│   │   ├── stock.ts           # BOM expansion engine (deficit, stock check)
│   │   └── export.ts          # Excel spreadsheet generation (Full BOM, Out of Stock, etc.)
```

---

## 4. Key Rules and Implementations

### A. "+ Add to Project" vs. "Add to Zoho Sales Order"
The builder UI separates local planning from Zoho integration:
* **`+ Add to Project`**: Found in the main configuration Output Panel. Adds the current cabinet to the session-only project array.
* **`Projects & Totals` Tab (Tab 3)**:
  * Houses the **`Zoho Sales Order`** search select and the **`Add to Zoho Sales Order`** button.
  * Displays the **`Project Lines`** table. Rates are editable inline. Line totals and the Grand Total recalculate dynamically.
  * Clicking the Zoho button resolves and creates all composite items for *all* cabinets in the local project and appends them to the selected Zoho Sales Order (or falls back to the currently active configured cabinet if the project is empty).

### B. Level 3 Part Splitting (LH/RH & Top/Bottom)
To accommodate different drilling locations:
* If a carcass/shutter profile set has `qty === 2`, `resolveProfileBOM` automatically splits it into two Level 3 composite parts of quantity 1:
  * **Vertical profiles** (names containing `V`, `vertical`, `LH+RH`, `elen`): Split into `Part Profile LH ...` and `Part Profile RH ...`.
  * **Horizontal profiles** (names containing `H`, `horizontal`, `TP+BT`, `edge`): Split into `Part Profile Top ...` and `Part Profile Bottom ...`.
* Both split Level 3 parts refer to the **same** Level 4 uncut panel item to prevent duplicate raw material listings.

### C. Unique Composite Names (Finish Inclusion)
To prevent duplicate composite items in Zoho (error `2057`), finish names are appended to Level 2 sets and Level 3 parts:
* E.g., `Part Panels - Left Wall 15x336x1070 GREY STONE` and `Part Profile LH MD1 327mm CHAMPAGNE`.
* If a composite item with that clean name/SKU is already resolved, the app reuses it.

### D. Single-Item Composite Safeguard (Drilling-1)
Zoho Inventory yields error `2056` if a composite item contains only one line item. To prevent this, **every Level 3 Part** has the `"Drilling-1"` service item automatically appended as a second component.

---

## 5. Verification Commands
To test type safety and build correctness, execute:
```bash
npm run build
```
To run the local server:
```bash
npm run dev
```
*(Runs on port 3000 by default. If port 3000 is occupied, it uses port 3001).*

---

## 6. How to Continue
1. Extract the handoff zip.
2. Ensure you have your environment variables defined in `.env.local` (Client ID, Secret, Refresh Token).
3. Run `npm install` and `npm run dev`.
4. Proceed with further modifications.
