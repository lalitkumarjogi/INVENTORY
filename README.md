

┌────────────────────────────────────────────────────────────────────────┐
│                        InventoryPro Web App                            │
│           (Dashboard.html / stock.html / Client UI)                    │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │ google.script.run (RPC)
┌───────────────────────────────────▼────────────────────────────────────┐
│                       Google Apps Script (Code.js)                     │
│  ┌──────────────────────┬──────────────────────┬────────────────────┐  │
│  │ Multi-Level Caching  │ Concurrency Locks    │ Anti-Duplication   │  │
│  │ (CacheService Chunk) │ (LockService 15s)    │ (30s Time Hash)    │  │
│  └──────────────────────┴──────────────────────┴────────────────────┘  │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │ SpreadsheetApp API
┌───────────────────────────────────▼────────────────────────────────────┐
│                        Google Sheets Database                          │
│  • Item_Master           • Purchase_Orders     • Stock_Adjustments     │
│  • Vendor_List           • Inventory_Balance   • Item_Daily_Summary    │
│  • Stock_Transactions    • Vendor and Material List                    │
└────────────────────────────────────────────────────────────────────────┘
 The system coordinates the complete manufacturing and warehouse inventory lifecycle:

1. Item & Vendor Master Setup

Item Registration: New raw materials (RM), semi-finished goods (SFG), and finished goods (FG) are cataloged in Item_Master with assigned Category, Sub-Category, and physical warehouse coordinates (Line, Rack, Box).

Vendor Onboarding: Vendors are added to Vendor_List with auto-generated alphanumeric identification codes (VEN-0X-MM-YYYY).

2. Purchase Order (PO) Lifecycle & Backordering

Creation: Procurement issues purchase orders (BHT/PO/XXX) with atomic line items, unit rates, expected delivery dates, and terms.

Receiving (Inward): Store staff receive physical shipments against the PO number.

Shortfall & Backorder Engine:

If a delivery is partial and marked for Backorder, the original PO line closes with status CLOSED-SHORT, and a new linked backorder PO (BHT/PO/XXX-BO) is automatically drafted for the remaining units.

If fully received, the line updates to FULFILLED.

3. Physical Stock Movements (Inward / Outward)

Inward (IN): Logs material receipts into inventory, updates vendor and invoice references, and immediately increments closing stock.

Outward (OUT): Validates live stock availability before issuing materials to production lines or dispatches. Throws a concurrency error if requested quantity exceeds available stock.

Rapid-Fire Guard: Blocks accidental double-submissions within 30 seconds by hashing recent user transactions.
