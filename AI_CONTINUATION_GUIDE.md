# Cabinet Carcass BOM Builder — AI Continuation Guide

This file provides a complete handoff and status guide for any AI coding assistant to resume development of the **Cabinet Carcass BOM Builder** project without needing context questions.

---

## 1. Project Overview & Architecture
This is a **Next.js** application built with TypeScript, React, TailwindCSS, and custom design patterns. It generates complete Bill of Materials (BOM) for customized kitchen cabinet carcasses, panels, profile raw materials, hardware, and assembly consumables.

### Core Files
* **Carcass BOM Builder Component**: [`src/components/CarcassBomBuilder.tsx`](file:///Users/Apple/Documents/bom-nextjs-app/Inventory%20-%20Magppie/src/components/CarcassBomBuilder.tsx)
  * Contains the entire pricing, sizing, profile cutting, hardware count, and assembly estimation logic for all 5 cabinet categories.
* **Raw Materials Database**: [`src/lib/rawmaterial.ts`](file:///Users/Apple/Documents/bom-nextjs-app/Inventory%20-%20Magppie/src/lib/rawmaterial.ts)
  * Translates generic items (e.g. side panels, profiles, drawer sets) to exact SKU inventory items in Magppie's catalog.
* **Rules Excel Compiler**: [`scratch/generate_rules_excel.js`](file:///Users/Apple/Documents/bom-nextjs-app/Inventory%20-%20Magppie/scratch/generate_rules_excel.js)
  * Compiles the entire cabinet filters, sizes, formulas, and lookups into an interactive Excel sheet (`cabinet_filter_rules_latest.xlsx`).

---

## 2. Carcass Sizing & Construction Rules
The carcass builder handles 5 main categories with different structural caps:
1. **Base (BC/BCL/BB/BBL)** & **Tall (TC/TB/TCL)**:
   * **Construction**: `fulltb` (Top/Bottom panels Cap/Overlay Left/Right side panels).
   * **Sides Sizing**: Side Height = `Height - 2 * Thickness` (`H - 2*t`). Side Depth = `Depth` (`D`) for Base, or `Depth - 15` (`D - 15`) for Tall.
   * **Top/Bottom Sizing**: Width = `W`. Depth = `D`.
2. **Wall (WC/WB)** & **Loft (LO/LB/LBF)** & **Mid (MD)**:
   * **Construction**: `fullsides` (Left/Right side panels Cap/Overlay Top/Bottom panels).
   * **Sides Sizing**: Side Height = `Height - 15` (`H - 15`) for Wall, or `Height` (`H`) for Loft/Mid. Side Depth = `Depth - 15` (`D - 15`) for Wall/Loft, or `Depth` (`D`) for Mid.
   * **Top/Bottom Sizing**: Width = `W - 2 * t`. Depth = `D` (or `D - 23` if XCJ profile top handle is active).
3. **Fillers & Visible Panels**:
   * **Filler Panels**: Width = `W`, Height = `H` (no size reduction).
   * **Visible Panels**: DepthRecess sizing matches side panel rules.

### Hardware & Consumables
* **PVC Legs**: 2 if W <= 150mm, 6 if W >= 1050mm, 4 otherwise (0 for wall-mounted units).
* **Wall Hangers**: 1 set for wall-mounted units (WC, WB, LO, LB, LBF, MD).
* **Fixing HPL Brackets**: 1 set for all units (large for floor-mounted, small for wall-hung).
* **Elenor with Light**: Integrated on Wall (non-dishrack), Loft, and Tall side panels. Adds ALU profiles (HM519), LED strip lights (MM), PVC diffusers (MM), and 3M tape (MM) with a `10%` waste factor.
* **Assembly Glue (DowSil 789)**: Sized in ML using the perimeter formula: `ROUND(((4 * D + BackPerimeter) / 1000) * 15)`.
* **Stepper Silicone (DowSil 789)**: Calculated using the run lookup matrix for back wall stepper profiles.

---

## 3. Included Metadata (Inside Project Zip)
To allow you to resume work seamlessly, this project zip contains a **`project_metadata`** directory:
* **`logs/transcript.jsonl`**: The complete chronological chat history and tool trace from the previous AI session. You can read/parse this to see exactly what decisions were made.
* **`cabinet_filter_rules_latest.xlsx`**: The dynamic Excel spreadsheet summarizing all cascade filters, sizing formulas, shutter spec tables, and silicone weights.
* **`cabinet_filter_rules.md`**: Markdown specification listing all cell formulas.
* **`task.md`** & **`walkthrough.md`**: Checklist of completed and verified features.

---

## 4. How to Run & Validate
* **Start Dev Server**:
  ```bash
  npm run dev
  ```
* **Compile Interactive Excel Rules**:
  ```bash
  node scratch/generate_rules_excel.js
  ```
* **Run BOM Tests**:
  ```bash
  node scratch/test_bom.js
  ```

---

## 5. Next Steps
1. Parse `project_metadata/logs/transcript.jsonl` using a JSONL reader if you need the full chat details.
2. Verify with the user if they want to make any structural modifications to the TSX file or raw materials mapping based on the compiled Excel sheet reviews.
