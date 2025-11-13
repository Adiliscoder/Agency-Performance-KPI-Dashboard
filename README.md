✅ How to Use the Dashboard

This dashboard is designed for end-users, meaning you can interact with the analytics, but you cannot modify the structure, formulas, VBA, or data model.

All sheets are protected, and the VBA project is locked to ensure integrity and prevent accidental changes.

Follow these steps:

🔐 1. Open the Dashboard

Open the file:
Agency-Performance-KPI-Dashboard.xlsm

Make sure you are using Excel for Windows
(PowerPivot + VBA required — not supported on Mac or Excel Online)

🔄 2. Enable Macros & External Data

When Excel opens:

Click “Enable Content” (yellow warning bar)

Click “Enable Macros”

If prompted, allow external connections for Power Query

This allows:

Automated refresh

Button actions

KPI recalculations

Slicer interactions

🖱️ 3. Use Only the Interactive Controls

The workbook is protected.
You can use the dashboard but not modify:

✔️ Allowed actions (for users):

Click on buttons to:

Refresh KPIs

Navigate to dashboard sections

Trigger automated scripts (if VBA button actions exist)

Use slicers to filter:

Zone

Bureau

Mutuelle

Usage

Lignes de produits (Auto, RD, RC, etc.)

Interact with PivotTables (filters only)

Explore charts and visual decompositions

❌ Not allowed (sheet is protected):

Modifying formulas

Changing PivotTables structure

Editing Power Query steps

Accessing VBA code (password protected)

Adding new data directly into tables

🔄 4. Refresh the Dashboard

To update KPIs with the latest data imports:

Option 1 — Refresh using the button (recommended)

If the dashboard includes a refresh button:

Click “Refresh Data”

Wait for:

Power Query updates

PowerPivot recalculations

DAX measures refresh

Option 2 — Manual

Go to:
Data → Refresh All

⚠️ The refresh process may take several seconds depending on your data size.

📌 5. Protected Workbook Behavior

The dashboard is intentionally protected so that:

The layout remains stable

DAX measures and data model remain consistent

Power Query

VBA automation cannot be overwritten

Business rules stay intact

No password is required for normal use.
Only the developer can unlock advanced layers.

🧪 6. Troubleshooting
If slicers don’t respond

➡️ Make sure macros and data connections are enabled.

If PivotTables show “Data Model not available”

➡️ You must use Excel for Windows (PowerPivot required).

If buttons do nothing

➡️ Macros are blocked — reopen the file and click Enable Macros.

🎯 Summary for Users

✔️ You can:
Filter, click, explore, refresh

❌ You cannot:
Modify formulas, VBA, Power Query, or structure
