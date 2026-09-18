# AI Handover & Continuation Guide — Cabinet BOM Builder

If you are an AI assistant resuming work on this codebase, this document contains everything you need to know to understand the current state, run the test suites, and continue development seamlessly.

---

## 1. Project Context & Purpose
This is a **Next.js React application** for designing cabinet carcasses and stone/glass shutters and generating their production **BOM (Bill of Materials)**.
- **Carcass & Shutter Builder Component**: The main building logic resides in [CarcassBomBuilder.tsx](src/components/CarcassBomBuilder.tsx). It calculates panels, parts, profiles, hardware, and consumables.
- **Zoho Inventory Integration**: Generates composite items bottom-up and pushes them to Zoho Sales Orders via `handleAddToZohoSO`.
- **UI Explosion & Excel View**: Displays a hierarchical tree view of the cabinet parts, profiles, and hardware.

---

## 2. Current State (Reverted to 11:30 AM Today)
The codebase has been successfully reverted to the morning session's stable state as of **11:30 AM today (June 15, 2026)**. The features active in the codebase are:

### A. Automatic Drawer Box Sizing (Lian Calculations)
- **Carcass Drawer Boxes** are automatically calculated when a cabinet variant has drawers.
- **Low Back Drawer Box**: Back panel (Height: 63mm), Bottom panel (6mm thick). Side profiles (HM513, len 483), center profiles (HM535, len $W-78$), and low hardware pack (H90).
- **High Back Drawer Box**: Back panel (Height: 212mm), Bottom panel (6mm thick). Side profiles (HM513, len 483), center profiles (HM535, len $W-78$), and high hardware pack (H239).

### B. Built-in (Inbuilt) Drawers
- **Trigger**: Selectable in the sidebar selector (Inbuilt Drawers dropdown: `none`, `1lb` - 1 Low Back, `1sb` - 1 Semi-High, `2lb` - 2 Low Back, `2sb` - 2 Semi-High).
- **Automatic Sizing & Fascia**:
  - Adds built-in drawer boxes using Lian sizing rules.
  - **Fascia Panel**: 15mm stone fascia panel (Width: $W - 2t - 8$, Height: 110mm for Low, 210mm for High/Semi). Drill: `JD`.
  - **Semi-High Drawer Box**: Pushes H175 runner hardware pack and uses a 148mm height back panel.

### C. Separate Level 1 composites (BOM Tree & Zoho Sync)
- **Separate composites**: Rather than nesting `Set of Parts-...` under `Drawer Pack-...` at Level 2, they are processed as **separate Level 1 composite items** (both in the BOM tree, Zoho sync, and Excel tree view groups).
- **Flat runner hardware**: Drawer runner hardware (e.g. H90, H175, H239) is processed as flat Level 1 carcass hardware in Zoho SO and in the BOM tree.

### D. Appliance Unit Shelves
- **Restored Custom Shelves**: Custom shelves for the Appliance cabinet (`APP` family) are sized at `depth = D - 13` for additional shelves.

---

## 3. Repository Structure & Active Files
- **Logic & UI Builder**: [src/components/CarcassBomBuilder.tsx](src/components/CarcassBomBuilder.tsx) (Contains all UI state, building logic, BOM generation, and Zoho sync).
- **Unit Test Script**: [scratch/test-drawers.js](scratch/test-drawers.js) (Validates drawer sizing, inbuilt parameters, PPN tall pantry shelves, and BOM tree/Zoho sync structures).

---

## 4. How to Verify & Build

### A. Run Automated Unit Tests
To verify the drawer box dimensions, built-in fascias, and architecture mappings:
```bash
node scratch/test-drawers.js
```
*Expected Output:*
```
=== RUNNING REFINED DRAWER BOX TESTS ===
Scenario A Passed!
Scenario B Passed!
Scenario C Passed!
Scenario D Passed!
=== RUNNING ARCHITECTURE TESTS (BOM TREE & ZOHO SYNC) ===
BOM Tree Architecture Verification Passed!
Zoho SO Architecture Verification Passed!
--- Testing Tall PO Shelf Pantry (PPN) Cabinet ---
PPN BOM Tree Architecture Verification Passed!
PPN Zoho SO Architecture Verification Passed!

ALL TESTS PASSED SUCCESSFULLY!
```

### B. Run TypeScript Compiler Checks
To ensure there are no compilation or type check errors:
```bash
npx tsc --noEmit
```

### C. Build the Application
To run the Next.js production build:
```bash
npm run build
```

---

## 5. Next Steps
You can continue implementing feature requests from this stable morning baseline. Feel free to reference `artifacts/walkthrough.md` for historical modifications and check the unit test assertions in `scratch/test-drawers.js` to see exact expected outputs.
