# Canara Bank e-passbook: monthly cash report

This tool reads the Canara Bank e-passbook PDFs for the Piston Agency current account. It produces one Excel report and prints a summary on screen. It runs entirely on your PC and does not use the internet.

## How to use

1. Install Python 3 from python.org. During setup, tick **"Add Python to PATH"**.
2. Copy the e-passbook PDFs into the `pdfs` folder. You can add several months at once. Statements whose dates overlap are fine, because repeated rows are dropped.
3. Double-click **`run_report.bat`**.
4. The report opens automatically. It is saved in `output\Cash_Report_<date>.xlsx`.

The first run installs two Python packages, `pdfplumber` and `openpyxl`. This needs internet once. If the PC never has internet, do this on another PC that does:
`pip download -r requirements.txt -d wheels`
Then copy the `wheels` folder next to `run_report.bat`.

**Password-protected PDFs:** open `run_report.bat` in Notepad and fill in `set PDF_PASSWORD=`.

## What's in the Excel file

| Sheet | Contents |
|---|---|
| **Summary** | Balance check result and a **3-month comparison** with averages: collections by mode, payments by payee group, SAM count, total and average, net cash flow, flags |
| **One sheet per month** | Collections by mode (UPI, NEFT/RTGS, IMPS, cheque clearing, cash deposit, Canara transfer) **by week**, payments by payee group, SAM statistics, that month's flags |
| **Flags** | All flags in one list |
| **Transactions** | Every transaction with deposit/withdrawal, balance, mode, payee group and balance check. Filters are switched on |
| **Unmapped payees** | Payments that went to "Others", with a suggested `payees.csv` line for each |
| **Checks** | Per PDF: opening balance, totals, last balance compared with the printed closing balance, and any rows that don't tally |

Weeks are counted by day of the month: Week 1 is days 1–7, Week 2 is days 8–14, and so on. Week 5 is days 29–31.

## How it works out the numbers

- **Deposit or withdrawal** comes from the change in balance compared with the previous row, not from which column the amount sits in. Every row is checked: previous balance ± amount must equal the new balance. Any row that doesn't tally is marked **MISMATCH**. The tool also compares the last balance with the PDF's printed closing balance. "ALL BALANCES TALLY" on the Summary sheet means the PDFs were read completely.
- **Collections** are all credits except owner money, loan/drawdown credits, charge reversals and returned payments (for example, "NEFT RETURN"). Those are shown separately as "Other credits". "Canara transfer" means a customer paid from their own Canara account ("FUNDS TRANSFER …").
- **Which month a row counts in:** Canara cuts each statement by posting date. Sometimes a row shows an earlier date than the row above it, for example 30-09 printed after 01-10. That row is counted on the latest date already reached, so month balances follow on exactly. The **Transactions** sheet shows these rows in the "Counted on" column.
- **Flags**
  - Two *different* parties received the **same amount within 7 days**. This could be a wrong payment. Bank charges, loan debits and amounts under Rs 1,000 are ignored.
  - The **same party paid the same amount twice within 7 days**. This could be a double payment. SAM, owner transfers and cash withdrawals are left out, because repeats are normal for them.
  - **Cheque returns** (CHQ RETURN, CHQ RTN and similar) and **returned payments** (NEFT RETURN and similar).
  - **Payments over Rs 2,00,000 to anyone other than SAM.**
  - Balance mismatches.

You can change these limits at the top of `passbook_report.py` (`LARGE_PAYMENT_LIMIT`, `DUPLICATE_WINDOW_DAYS`, and so on).

## payees.csv: mapping names to groups

Open it in Excel or Notepad. Each line is `keyword,group`:

```
SAMKRG*,SAM
MACO,MACO
RAJESH KUMAR,Owner transfer      <- add the owner's name(s) here
SUNDAR ELECTRICALS,Electrical    <- any new group name becomes its own row in the report
```

- The keyword is looked for in the Particulars. Case doesn't matter and whole words are matched. Add `*` at the end to match the start of a word.
- Rules are checked from the top down, and the **first match wins**.
- Special groups are `Owner transfer`, `Loan/Drawdown`, `Bank charges` and `Cash withdrawal`.
- Add the owner's name(s) **and the owner's own account number(s)**. Some owner transfers only show the account number. Otherwise owner transfers appear under "Others". Keep your filled-in copy on your PC; don't commit it.
- After editing, run the `.bat` again. Check the **Unmapped payees** sheet for names to add.

## If something looks wrong

Run this from a command prompt in this folder. It prints every row the tool read, so you can compare it with the PDF:

```
python passbook_report.py --dump
```

Other options: `--months 6` gives a 6-month comparison, and `python passbook_report.py D:\statements` reads PDFs from another folder.

Your bank PDFs and reports are kept out of git (see `.gitignore`).
