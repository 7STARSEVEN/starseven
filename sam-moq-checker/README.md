# SAM Order Checker + Tracker – Piston Agency, Nagpur

Two buttons:

| Button | What it does |
|---|---|
| **Check SAM Order** | Checks a SAM purchase order **before you send it**: MOQ scheme slabs, pack sizes, order value vs weekly cap, what is still open with SAM, and **stock cover** (are we short or over-stocked). Makes `<order name>_CHECK.xlsx` next to the order. |
| **Update SAM Tracker** | Shows everything still **to come from SAM**: SAM's pending list + orders sent after it − SAM bills received after it. Makes `SAM_Order_Tracker.xlsx`. |

The tool only **reads** your files. It never changes them.

---

## 1. One-time setup (each computer)

1. Install **Python** from python.org. On the first screen **tick "Add python.exe to PATH"**, then *Install Now*.
2. Unzip `SAM_MOQ_Checker.zip` onto the **Desktop**, so you have `Desktop\SAM_MOQ_Checker\`.
3. Open the `Desktop shortcuts` folder inside it and copy **`Check SAM Order.bat`** and
   **`Update SAM Tracker.bat`** onto the Desktop (next to the `SAM_MOQ_Checker` folder).
4. Double-click **Check SAM Order** once **with internet on** – the first time it installs two free
   free readers (`openpyxl`, `xlrd` for Excel, `pypdf` for SAM's PDF bills). After that no internet is needed.
   If Windows says "Windows protected your PC": *More info* → *Run anyway*.

## 2. The folders – keep them filled

| Folder | Put here | How often |
|---|---|---|
| `Pending` | SAM's pending-orders list (Excel) | Every time SAM sends a new one |
| `Orders Sent` | A copy of **every PO you send** to SAM. **Date in the file name**, e.g. `SAM_PO_23.11.2026.xls` | Every order |
| `SAM Bills` | Every SAM bill / tax invoice – the **PDF as SAM sends it** (Excel also works) | Every bill |
| `Stock` | *(only if `stock_check = yes`)* Tally **Stock Summary** exported to Excel (with Opening / Inwards / Outwards / Closing columns and the period at the top) | Every week |

The newest file is used in `Pending` and `Stock`. In `Orders Sent` and `SAM Bills` every file is used.

## 3. Checking an order

1. Fill SAM's PO in Excel and **save** it.
2. Double-click **Check SAM Order** and pick the PO (or drag the PO onto the icon).
3. Read the screen, then the report that opens. Fix the PO, save, check again.
4. When it's right: send it to SAM **and save a copy in `Orders Sent`** (date in the name).

**Report sheets**

| Sheet | Shows |
|---|---|
| **Summary** | Key numbers, a numbered **"What to fix before sending"** list, the MOQ table (GREEN = Slab-2 hit, RED = Slab-2 not hit), models not in the scheme, models still open with SAM, warnings. |
| **Detail** | One row per model, sizes across, rate, value, open with SAM. |
| **Stock cover** | Per model + size: stock in hand, open with SAM, this order, monthly sales, months of cover. **SHORT** (red) = under 1 month → "ADD x pcs". **OVER** (orange) = more than 3 months → "CUT x pcs from this order". **NO SALE** (yellow) = ordered but nothing sold. |
| **Open with SAM** | Every open GDA / PO line with ordered, received, balance. |
| **SAM bills read** | Every bill line and which order it was set against. |

Stock cover = (stock in hand + still open with SAM + this order) ÷ monthly sales from Tally.
Quantities are rounded to the pack size. If a CUT would drop an MOQ group below Slab-2, the note says so –
then decide which matters more.

## 4. The tracker

Double-click **Update SAM Tracker**. `SAM_Order_Tracker.xlsx` opens:

- **Open with SAM** – every line still to come: GDA lines from SAM's pending list, plus every PO in
  `Orders Sent` dated **after** the latest GDA date in that list. Status OPEN / PART (yellow) / DONE (green).
- **SAM bills read** – each bill line is set against the **oldest** open line of the same Gr Code + size.
  Bills dated on/before the pending list's date are SKIPPED (SAM's list already shows them).
  If the bill names its GDAs ("Order No: GDA NO.1534,1545,1424"), those GDAs are used first.
  A bill quantity that matches no open order is flagged **EXTRA** – check it with SAM.
- **MOQ check on bills** – the "MOQ Scheme (Rs.)" column of every bill is compared with `moq.csv`:
  the tool adds up each MOQ group per GDA and works out the benefit due. **SHORT-PAID** lines are shown in red
  with the amount to claim from SAM. One PDF may hold many invoices – each page is read separately.

## 5. Updating the MOQ scheme – `moq.csv`

When SAM sends a new scheme: open `moq.csv` in Excel, one line per group, columns in this order:
`Group, Gr Codes (space separated), Slab-1 MOQ, Slab-1 Rs/pc, Slab-2 MOQ, Slab-2 Rs/pc`. Save as CSV.
A code may be in only one group. Then update `scheme_name` / `scheme_end_date` in `settings.ini`.

## 6. Rates – `rates.csv`

Lists every code in SAM's PO. Type the **rate incl GST** in the third column and save as CSV.
Until rates are filled, order value and the weekly-cap check stay incomplete.

## 7. Settings – `settings.ini` (open with Notepad)

| Setting | Default | Meaning |
|---|---|---|
| `count_mode` | group | `group` = all codes of a group added (SAM's rule); `model` = each code on its own |
| `weekly_cap` | 6,00,000 | Warn when the order value is above this |
| `scheme_end_date` | blank | Reminder to update `moq.csv` after this date |
| `scheme_check_ignore` | Bravo codes | Differences with SAM's PO summary you already confirmed |
| `pending_customer` | blank | Only use pending lines whose customer name contains this |
| `pending_max_age_days` | 7 | Warn if the pending list is older |
| `stock_check` | no | Stock cover check from Tally – switched **off** until Tally stock is reliable. Set `yes` to use it. |
| `stock_months` | blank | Months in the stock file (blank = read from its period) |
| `min_cover_months` / `target_cover_months` / `max_cover_months` | 1 / 1.5 / 3 | Stock cover limits |
| `stock_check_all_codes` | yes | Also list selling sizes NOT in the order that are short |
| `open_report` | yes | Open the report automatically |

Blank = default. Unknown lines are ignored.

## 8. Common problems

| Message | What to do |
|---|---|
| "Python is not installed" | Setup step 1 – tick "Add python.exe to PATH". |
| "openpyxl / xlrd / pypdf is not installed" | Run the .bat once with internet on. |
| "No quantities found" | Empty PO picked, or not saved. |
| "No stock file found" | Export Tally Stock Summary into the `Stock` folder. |
| "No SAM pending list found" | Put SAM's pending Excel into the `Pending` folder. |
| "Could not read bill …" | That bill's format is unusual – send it to whoever set up the tool. Other bills still count. |
| "Model Code is 'GRXXXX' – NOT checked" | A quantity on a line with no real code. Ask SAM for the code. |
| Report saved as `..._153012.xlsx` | The old report was open in Excel. Close old reports. |
| Anything else | The window shows "send this message to whoever set up the tool" – send a photo of it. |
