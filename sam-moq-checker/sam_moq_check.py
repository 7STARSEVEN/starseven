#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SAM ORDER CHECKER + TRACKER  -  Piston Agency, Nagpur
======================================================
1. ORDER CHECK  (Check_SAM_Order.bat)
   Checks a SAM purchase order before it is sent: MOQ slabs per group, pack sizes,
   order value vs weekly cap, what is still open with SAM, and stock cover
   (stock in hand + open with SAM + this order, against monthly sales from Tally).
   Writes <order name>_CHECK.xlsx next to the order.

2. TRACKER  (Update_SAM_Tracker.bat  =  sam_moq_check.py --tracker)
   Open with SAM = SAM pending list + orders sent after it - SAM bills after it.
   Writes SAM_Order_Tracker.xlsx in this folder.

Input files are only READ, never changed. Libraries: standard library + openpyxl + xlrd.
Settings: settings.ini   MOQ table: moq.csv   Rates: rates.csv
"""

import argparse
import configparser
import csv
import datetime as dt
import hashlib
import math
import os
import re
import subprocess
import sys
import traceback
import warnings
from collections import OrderedDict
from types import SimpleNamespace

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except Exception:
        pass

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("\nPROBLEM: the 'openpyxl' package is not installed on this computer.\n"
          "Double-click Check_SAM_Order.bat once while connected to the internet - it installs it,\n"
          "or open Command Prompt and type:   py -m pip install openpyxl xlrd\n")
    sys.exit(2)

warnings.filterwarnings("ignore", module="openpyxl")

TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
MOQ_FILE = os.path.join(TOOL_DIR, "moq.csv")
RATES_FILE = os.path.join(TOOL_DIR, "rates.csv")
SETTINGS_FILE = os.path.join(TOOL_DIR, "settings.ini")
PENDING_DIR = os.path.join(TOOL_DIR, "Pending")
ORDERS_DIR = os.path.join(TOOL_DIR, "Orders Sent")
BILLS_DIR = os.path.join(TOOL_DIR, "SAM Bills")
STOCK_DIR = os.path.join(TOOL_DIR, "Stock")
TRACKER_FILE = os.path.join(TOOL_DIR, "SAM_Order_Tracker.xlsx")
LAST_FOLDER_FILE = os.path.join(TOOL_DIR, ".last_folder")

EXCEL_EXTS = (".xls", ".xlsx", ".xlsm")
BILL_EXTS = EXCEL_EXTS + (".pdf",)
STANDARD_SIZES = ["STD", "0.2", "0.25", "0.4", "0.5", "0.6", "0.75", "0.8", "1", "1.25", "1.5"]
PLACEHOLDER_CODES = {"GR0000"}

CODE_CELL_RE = re.compile(r"^\s*GR\s*[-_]?\s*[0-9O]{3,6}\s*$", re.I)
CODE_IN_TEXT_RE = re.compile(r"(?i)(?<![A-Z])[GC]R\s*-?\s*([0-9O]{3,6})(?![0-9])")
SCHEME_AREA_RE = re.compile(r"\bslab\b|summary\s+of\s+moq|scheme\s+model|eligible", re.I)
SIZE_TOKEN_RE = re.compile(r"^\+?(STD|STANDARD|\d?\.\d{1,3})$", re.I)
PERIOD_RE = re.compile(r"(\d{1,2}-[A-Za-z]{3}-\d{2,4})\s*to\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})", re.I)


class CheckError(Exception):
    """A problem we can explain to the user in plain words."""


# --------------------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------------------

def key(value):
    """Header text -> lower-case letters and digits only ('Gr. Code' -> 'grcode')."""
    return re.sub(r"[^a-z0-9]", "", str(value).lower()) if value is not None else ""


def get(row, j):
    """Cell value from a row list; blank strings become None."""
    if j is None or j < 0 or j >= len(row):
        return None
    v = row[j]
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def to_number(value):
    """Cell -> float.  Blank / '-' -> 0.  Text that is not a number -> None."""
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = re.sub(r"(?i)rs\.?|inr|₹|/-|,", "", str(value)).strip()
    if s in ("", "-", "--"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return None


def qty_number(value):
    """Quantity cells from Tally/bills: 120, '120 Nos', '(-)5 Pcs', '-5'.  Unreadable -> 0."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    m = re.search(r"(\(-\)|-)?\s*([\d,]*\.?\d+)", str(value or ""))
    if not m:
        return 0.0
    n = float(m.group(2).replace(",", ""))
    return -n if m.group(1) else n


def norm_code(text):
    """'GR 238' / 'gr-0238' / 'CR0238' / 'GRO238' -> 'GR0238'."""
    s = re.sub(r"[\s\-_./]", "", str(text).upper())
    m = re.match(r"^[GC]R([0-9O]+)$", s)
    if m:
        return "GR%04d" % int(m.group(1).replace("O", "0"))
    return s


def valid_code(value):
    """Normalised code if the cell really holds a SAM model code, else ''."""
    if not isinstance(value, str) or not CODE_CELL_RE.match(value):
        return ""
    code = norm_code(value)
    return "" if code in PLACEHOLDER_CODES else code


def codes_in_text(text):
    return [norm_code("GR" + m) for m in CODE_IN_TEXT_RE.findall(str(text or ""))]


def parse_size(value):
    """Size heading/cell -> 'STD', '0.25', '1', ...  or None."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        f = float(value)
    else:
        s = re.sub(r"\s+", "", str(value).upper())
        if s in ("STD", "STD.", "STANDARD", "S.T.D"):
            return "STD"
        s = re.sub(r"^(O/S|OS|\+)", "", s)
        s = re.sub(r"MM$", "", s)
        if s.startswith("."):
            s = "0" + s
        try:
            f = float(s)
        except ValueError:
            return None
    return ("%g" % round(f, 3)) if 0 < f <= 3 else None


def size_in_name(text):
    """Last size token in an item name: 'GR0316 H.S.SPLENDOR 0.25 WOSP' -> '0.25'.  None if no size."""
    found = None
    for tok in re.split(r"[\s,;()\[\]]+", str(text or "")):
        if SIZE_TOKEN_RE.match(tok):
            s = parse_size(tok.lstrip("+"))
            if s in STANDARD_SIZES:
                found = s
    return found


def size_sort_key(s):
    return -1.0 if s == "STD" else (float(s) if re.match(r"^[\d.]+$", s) else 99.0)


def clean_text(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def ref_text(v):
    if isinstance(v, float):
        return fmt_qty(v)
    return clean_text(v)


def inr(n):
    """Indian digit grouping: 600000 -> '6,00,000'."""
    if n is None:
        return "-"
    n = int(round(n))
    sign, s = ("-" if n < 0 else ""), str(abs(n))
    if len(s) > 3:
        head, tail, parts = s[:-3], s[-3:], []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        s = ",".join(parts + [tail])
    return sign + s


def fmt_qty(q):
    if q is None:
        return ""
    q = float(q)
    return ("%d" % q) if q.is_integer() else ("%.2f" % q)


def parse_date(value):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and 20000 < value < 80000:  # Excel serial number
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(value))
    s = str(value or "").strip()
    for f in ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%d-%b-%Y", "%d-%b-%y",
              "%d-%m-%y", "%d/%m/%y", "%d.%m.%y", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def date_in_filename(path):
    """'SAM_PO_23.11.2026.xls' / 'bill 26-11-26' / 'Stock_24Nov26' -> date, else None."""
    name = os.path.splitext(os.path.basename(path))[0]
    for m in re.finditer(r"(?<!\d)(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{4}|\d{2})(?!\d)", name):
        d, mo, y = (int(x) for x in m.groups())
        try:
            return dt.date(y + 2000 if y < 100 else y, mo, d)
        except ValueError:
            continue
    m = re.search(r"(?<!\d)(\d{1,2})([A-Za-z]{3})(\d{4}|\d{2})(?!\d)", name)
    if m:
        return parse_date("%s-%s-%s" % m.groups())
    return None


def file_date(path):
    """Date in the file name, else the file's saved date."""
    return date_in_filename(path) or dt.date.fromtimestamp(os.path.getmtime(path))


def ddmmyyyy(d):
    return d.strftime("%d-%m-%Y") if d else ""


def file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def excel_files_in(folder, exts=EXCEL_EXTS):
    """Excel files in a folder, newest first.  Skips Excel lock files and this tool's own reports."""
    if not os.path.isdir(folder):
        return []
    out = []
    for name in os.listdir(folder):
        low = name.lower()
        if (low.endswith(exts) and not low.startswith("~$") and "_check" not in low
                and not low.startswith("sam_order_tracker")):
            out.append(os.path.join(folder, name))
    return sorted(out, key=os.path.getmtime, reverse=True)


def ceil_to(x, pack):
    pack = pack if pack and pack > 0 else 1
    return math.ceil(round(x / pack, 6)) * pack


def floor_to(x, pack):
    pack = pack if pack and pack > 0 else 1
    return math.floor(round(x / pack, 6)) * pack


# --------------------------------------------------------------------------------------
# Settings, moq.csv, rates.csv
# --------------------------------------------------------------------------------------

DEFAULT_SETTINGS = OrderedDict([
    ("company_name", "Piston Agency, Nagpur"), ("count_mode", "group"), ("weekly_cap", "600000"),
    ("show_groups_not_ordered", "yes"), ("scheme_name", ""), ("scheme_end_date", ""),
    ("check_scheme_in_order_file", "yes"), ("scheme_check_ignore", ""), ("pending_file", ""),
    ("pending_customer", ""), ("pending_max_age_days", "7"), ("order_sheets", ""), ("open_report", "yes"),
    ("stock_check", "no"), ("stock_file", ""), ("stock_months", ""), ("min_cover_months", "1"), ("target_cover_months", "1.5"),
    ("max_cover_months", "3"), ("stock_check_all_codes", "yes"),
])


def yes(value):
    return str(value).strip().lower() in ("yes", "y", "true", "1", "on")


def read_settings_file(path):
    values = dict(DEFAULT_SETTINGS)
    if not os.path.exists(path):
        return values
    cp = configparser.ConfigParser(inline_comment_prefixes=(";", "#"), interpolation=None)
    try:
        try:
            cp.read(path, encoding="utf-8-sig")
        except UnicodeDecodeError:
            cp.read(path, encoding="cp1252")
    except configparser.Error as e:
        raise CheckError("settings.ini has a mistake in it:\n  %s" % e)
    for section in cp.sections():
        for k, v in cp.items(section):
            if (v or "").strip():          # blank = keep the default
                values[k.lower()] = v.strip()
    return values


def number_setting(values, name):
    n = to_number(values[name])
    if n is None or n < 0:
        raise CheckError("settings.ini: %s must be a number (found '%s')" % (name, values[name]))
    return n


class Settings:
    def __init__(self, path=SETTINGS_FILE):
        v = read_settings_file(path)
        self.company = v["company_name"]
        self.mode = "model" if "model" in v["count_mode"].lower() else "group"
        self.weekly_cap = number_setting(v, "weekly_cap")
        self.show_not_ordered = yes(v["show_groups_not_ordered"])
        self.scheme_name = v["scheme_name"]
        self.scheme_end = parse_date(v["scheme_end_date"]) if v["scheme_end_date"] else None
        self.check_po_scheme = yes(v["check_scheme_in_order_file"])
        self.scheme_ignore = {norm_code(c) for c in re.split(r"[\s,;/]+", v["scheme_check_ignore"]) if c}
        self.pending_file = v["pending_file"].strip('"')
        self.pending_customer = v["pending_customer"]
        self.pending_max_age = number_setting(v, "pending_max_age_days")
        self.order_sheets = [s.strip().lower() for s in v["order_sheets"].split(",") if s.strip()]
        self.open_report = yes(v["open_report"])
        self.stock_check = yes(v["stock_check"])
        self.stock_file = v["stock_file"].strip('"')
        self.stock_months = number_setting(v, "stock_months") if v["stock_months"] else None
        self.min_cover = number_setting(v, "min_cover_months")
        self.target_cover = number_setting(v, "target_cover_months")
        self.max_cover = number_setting(v, "max_cover_months")
        self.stock_all = yes(v["stock_check_all_codes"])


def read_csv_rows(path):
    raw = None
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, newline="", encoding=enc) as fh:
                raw = fh.read()
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        raise CheckError("Cannot read %s" % path)
    lines = raw.splitlines()
    first = lines[0] if lines else ""
    delim = ";" if first.count(";") > first.count(",") else ("\t" if first.count("\t") > first.count(",") else ",")
    return [[c.strip() for c in row] for row in csv.reader(lines, delimiter=delim)]


def parse_moq_row(n, row, problems):
    if len(row) < 6:
        problems.append("line %d: needs 6 columns (Group, Gr Codes, Slab-1 MOQ, Slab-1 Rs/pc, Slab-2 MOQ, Slab-2 Rs/pc)" % n)
        return None
    name, codes = row[0], []
    for c in re.split(r"[\s,;/+]+", row[1]):
        if c and norm_code(c) not in codes:
            codes.append(norm_code(c))
    if not name or not codes:
        problems.append("line %d: needs a group name and at least one Gr Code" % n)
        return None
    nums = [to_number(x) for x in row[2:6]]
    bad = ["%s = '%s'" % (t, row[2 + k]) for k, t in
           enumerate(("Slab-1 MOQ", "Slab-1 Rs/pc", "Slab-2 MOQ", "Slab-2 Rs/pc")) if nums[k] is None or nums[k] < 0]
    if bad:
        problems.append("line %d (%s): not a number: %s" % (n, name, ", ".join(bad)))
        return None
    if nums[2] < nums[0]:
        problems.append("line %d (%s): Slab-2 MOQ is smaller than Slab-1 MOQ" % (n, name))
    return SimpleNamespace(name=name, codes=codes, s1=nums[0], r1=nums[1], s2=nums[2], r2=nums[3])


def load_moq(path=MOQ_FILE):
    """moq.csv columns (in this order): Group, Gr Codes, Slab-1 MOQ, Slab-1 Rs/pc, Slab-2 MOQ, Slab-2 Rs/pc"""
    if not os.path.exists(path):
        raise CheckError("moq.csv is missing. It must be in the same folder as this tool:\n  %s" % path)
    groups, seen, problems = [], {}, []
    for n, row in enumerate(read_csv_rows(path), start=1):
        if not any(row) or row[0].startswith("#"):
            continue
        if n == 1 and to_number(row[2] if len(row) > 2 else "x") is None:
            continue  # header line
        g = parse_moq_row(n, row, problems)
        if not g:
            continue
        for code in g.codes:
            if code in seen:
                problems.append("%s is in two groups: '%s' and '%s'" % (code, seen[code], g.name))
            seen[code] = g.name
        groups.append(g)
    if problems:
        raise CheckError("moq.csv has mistakes - please fix them and run again:\n  - " + "\n  - ".join(problems))
    if not groups:
        raise CheckError("moq.csv has no MOQ groups in it.")
    return groups


def load_rates(path=RATES_FILE):
    """rates.csv: Gr Code, Model (optional, for reference), Rate incl GST.  Blank rate = unknown."""
    rates, labels, problems = {}, {}, []
    if not os.path.exists(path):
        return rates, labels, ["rates.csv not found - order value cannot be worked out."]
    rows = read_csv_rows(path)
    if not rows:
        return rates, labels, problems
    has_header = valid_code(rows[0][0] if rows[0] else "") == ""
    header = [key(h) for h in rows[0]] if has_header else []
    rate_col = next((j for j, k in enumerate(header) if "rate" in k or "price" in k), len(rows[0]) - 1)
    for n, row in enumerate(rows[1:] if has_header else rows, start=2 if has_header else 1):
        if not row or not row[0]:
            continue
        code = norm_code(row[0])
        if len(row) > 1 and rate_col != 1:
            labels[code] = row[1]
        val = row[rate_col] if rate_col < len(row) else ""
        if val == "":
            continue
        r = to_number(val)
        if r is None or r < 0:
            problems.append("rates.csv line %d: rate '%s' for %s is not a number - ignored" % (n, val, code))
        elif r > 0:
            rates[code] = r
    return rates, labels, problems


# --------------------------------------------------------------------------------------
# Reading Excel files (.xlsx with openpyxl, SAM's old .xls with xlrd) - READ ONLY
# --------------------------------------------------------------------------------------

def xls_cell(xlrd, book, cell):
    if cell.ctype == xlrd.XL_CELL_DATE:
        try:
            return xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)
        except Exception:
            return cell.value
    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
        return None
    return cell.value


def load_xls(path):
    try:
        import xlrd
    except ImportError:
        raise CheckError("This is an old-style .xls file. Reading it needs the free 'xlrd' package.\n"
                         "  Double-click Check_SAM_Order.bat once while connected to the internet, or type in\n"
                         "  Command Prompt:   py -m pip install xlrd")
    try:
        book = xlrd.open_workbook(path, on_demand=True)
    except Exception as e:
        raise CheckError("Cannot open %s\n  (%s)" % (os.path.basename(path), e))
    try:
        out = []
        for i in range(book.nsheets):
            sh = book.sheet_by_index(i)
            rows = [[xls_cell(xlrd, book, sh.cell(r, c)) for c in range(sh.ncols)] for r in range(sh.nrows)]
            out.append((sh.name, rows))
            book.unload_sheet(i)
        return out
    finally:
        book.release_resources()


def load_xlsx(path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        raise CheckError("Cannot open %s\n  (%s)" % (os.path.basename(path), e))
    try:
        out = []
        for ws in wb.worksheets:
            try:
                ws.reset_dimensions()
            except Exception:
                pass
            out.append((ws.title, [list(r) for r in ws.iter_rows(values_only=True)]))
        return out
    finally:
        wb.close()


def load_sheets(path):
    """[(sheet name, [row values...]), ...]  The file is opened read-only and never saved."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return load_xlsx(path)
    if ext == ".xls":
        return load_xls(path)
    raise CheckError("'%s' is not an Excel file (.xls / .xlsx)." % os.path.basename(path))


# --------------------------------------------------------------------------------------
# SAM order (PO) reader
#
# Per sheet: every column holding GR codes starts a block that runs until the next such
# column (the cylinder-kit sheet has two tables side by side). Walking down a block, the
# latest heading row tells which column is which size, TOTAL, MODEL NAME, Unit Packing,
# DLP... Rows holding a valid code are order lines. SAM's "Summary of MOQ Scheme" area at
# the bottom of a sheet is read separately.
# --------------------------------------------------------------------------------------

HEADER_WORDS = {
    "std", "size", "sizes", "sizesrequired", "oem", "seg", "dlp", "rate", "mrp", "modelname",
    "modelcode", "grcode", "groupcode", "model", "description", "concernplants", "kitcolour",
    "kitcolor", "totamt", "amount", "partname", "partno", "sno", "srno", "slno",
}


def is_header_word(k):
    return (k in HEADER_WORDS or "packing" in k or k.startswith("moqeach")
            or k.startswith("rate") or k.startswith("modelname"))


def classify_header(k):
    if "packing" in k or k.startswith("moqeach"):
        return "pack"
    if k in ("totamt", "amount", "amt", "totalamount", "totalvalue", "value", "netamount"):
        return "amount"
    if k in ("total", "totalqty", "qty", "quantity", "orderqty", "totalquantity", "totalpcs", "ordqty"):
        return "total"
    if k in ("dlp", "nlc", "dp", "unitprice", "price") or (k.startswith("rate") and k[4:7] != "gst"):
        return "rate"
    if k.startswith("modelname") or k in ("model", "description", "partname"):
        return "model"
    return None


def scheme_cutoff(rows):
    for i, row in enumerate(rows):
        if any(isinstance(v, str) and SCHEME_AREA_RE.search(v) for v in row):
            return i
    return len(rows)


def update_block_headings(block, cells, texts, code_col):
    sizes = OrderedDict((j, parse_size(v)) for j, v in cells if j != code_col and parse_size(v))
    if sizes:
        block.sizes = sizes
    for j, v in texts:
        kind = classify_header(key(v))
        if kind:
            block.cols[kind] = j
            if kind == "rate":
                block.rate_label = clean_text(v)


def stray_qty_warning(sheet, i, row, c, block, texts, keys):
    """A quantity typed on a line without a valid Gr Code (GRXXXX, U/D, blank...)."""
    q = sum(x for x in (to_number(get(row, j)) for j in block.sizes) if x and x > 0)
    if q <= 0 or any("total" in k for k in keys):
        return None
    code_text = get(row, c)
    other = [clean_text(v) for j, v in texts if j != c]
    if code_text is None and not other:
        return None
    return ("Sheet '%s' row %d: %s pcs typed for '%s' but its Model Code is '%s' - this line is NOT checked "
            "(no valid Gr Code)." % (sheet.strip(), i + 1, fmt_qty(q), other[0] if other else "",
                                    clean_text(code_text) or "blank"))


def note_catalog(catalog, code, row, block):
    pack = to_number(get(row, block.cols.get("pack"))) if "pack" in block.cols else None
    if pack and pack > 0:
        catalog["packs"].setdefault(code, pack)
    if "model" in block.cols:
        name = clean_text(get(row, block.cols["model"]))
        if name:
            catalog["names"].setdefault(code, name)


def read_block(sheet, rows, c, end, cutoff, catalog):
    lines, warns = [], []
    block = SimpleNamespace(sizes=OrderedDict(), cols={}, rate_label="")
    for i in range(cutoff):
        row = rows[i]
        code = valid_code(get(row, c))
        if code:
            note_catalog(catalog, code, row, block)
            line = read_order_line(sheet, i, row, code, block)
            if line:
                lines.append(line)
            continue
        cells = [(j, get(row, j)) for j in range(c, min(end, len(row)))]
        texts = [(j, v) for j, v in cells if isinstance(v, str)]
        keys = [key(v) for _, v in texts]
        if any(parse_size(v) == "STD" for _, v in texts) or any(is_header_word(k) for k in keys if k != "total"):
            update_block_headings(block, cells, texts, c)
            continue
        w = stray_qty_warning(sheet, i, row, c, block, texts, keys)
        if w:
            warns.append(w)
    return lines, warns


def read_order_sheet(sheet, rows, catalog):
    """Returns (order lines, warnings, cutoff row index)."""
    cutoff = scheme_cutoff(rows)
    code_cols = sorted({j for i in range(cutoff) for j, v in enumerate(rows[i]) if valid_code(v)})
    lines, warns = [], []
    if not code_cols:
        return lines, warns, cutoff
    width = max(len(r) for r in rows)
    for n, c in enumerate(code_cols):
        end = code_cols[n + 1] if n + 1 < len(code_cols) else width
        bl, bw = read_block(sheet, rows, c, end, cutoff, catalog)
        lines.extend(bl)
        warns.extend(bw)
    return lines, warns, cutoff


def line_sizes(row, block, notes):
    sizes = OrderedDict()
    for j, s in block.sizes.items():
        v = get(row, j)
        q = to_number(v)
        if q is None:
            if str(v).strip().upper() not in ("X", "NA", "N/A", "-"):
                notes.append("'%s' in size %s is not a number - ignored" % (v, s))
            continue
        if q:
            sizes[s] = sizes.get(s, 0) + q
    return sizes


def read_order_line(sheet, i, row, code, block):
    notes = []
    sizes = line_sizes(row, block, notes)
    qty = sum(sizes.values())
    total = to_number(get(row, block.cols.get("total"))) if "total" in block.cols else None
    if total and qty == 0:
        qty, sizes["STD" if not block.sizes else "Other"] = total, total
        if block.sizes:
            notes.append("size columns blank - TOTAL column (%s) used" % fmt_qty(total))
    elif total and abs(total - qty) > 0.001:
        notes.append("TOTAL column says %s but sizes add up to %s - sizes used" % (fmt_qty(total), fmt_qty(qty)))
    if qty == 0:
        return None
    if qty < 0 or any(q < 0 for q in sizes.values()):
        notes.append("NEGATIVE quantity - please check")
    pack = to_number(get(row, block.cols.get("pack"))) if "pack" in block.cols else None
    if pack and pack > 0:
        odd = ["%s: %s" % (s, fmt_qty(q)) for s, q in sizes.items()
               if s != "Other" and abs(q / pack - round(q / pack)) > 1e-9]
        if odd:
            notes.append("not in multiples of pack size %s (%s)" % (fmt_qty(pack), ", ".join(odd)))
    rate = to_number(get(row, block.cols.get("rate"))) if "rate" in block.cols else None
    amount = to_number(get(row, block.cols.get("amount"))) if "amount" in block.cols else None
    return {"sheet": sheet.strip(), "row": i + 1, "code": code,
            "model": clean_text(get(row, block.cols.get("model"))) if "model" in block.cols else "",
            "sizes": sizes, "qty": qty, "pack": pack,
            "rate": rate if rate and rate > 0 else None, "rate_label": block.rate_label,
            "amount": amount if amount and amount > 0 else None, "notes": notes}


def scheme_line(sheet, r, k, s1c, r1c, s2c, r2c):
    codes, label = [], ""
    for j in range(0, s1c):
        v = get(r, j)
        if isinstance(v, str):
            found = codes_in_text(v)
            if found:
                codes.extend(found)
            elif not label:
                label = clean_text(v)
    nums = [to_number(get(r, x)) for x in (s1c, r1c, s2c, r2c)]
    if codes and all(x is not None and x > 0 for x in nums):
        return {"sheet": sheet.strip(), "row": k + 1, "codes": codes, "label": label,
                "s1": nums[0], "r1": nums[1], "s2": nums[2], "r2": nums[3]}
    return None


def read_scheme_area(sheet, rows, cutoff):
    """SAM's own 'Summary of MOQ Scheme' rows at the bottom of a PO sheet."""
    out, i = [], cutoff
    while i < len(rows):
        row = rows[i]
        elig = [j for j, v in enumerate(row) if isinstance(v, str) and "eligible" in v.lower()]
        ben = [j for j, v in enumerate(row) if isinstance(v, str) and re.search(r"(?i)bene?fit", v)]
        if len(elig) >= 2 and ben:
            s1c, s2c = elig[0], elig[1]
            r1c = next((j for j in ben if s1c < j < s2c), None)
            r2c = next((j for j in ben if j > s2c), None)
            if r1c is not None and r2c is not None:
                k = i + 1
                while k < len(rows) and any(get(rows[k], j) is not None for j in range(len(rows[k]))):
                    line = scheme_line(sheet, rows[k], k, s1c, r1c, s2c, r2c)
                    if line:
                        out.append(line)
                    k += 1
                i = k
        i += 1
    return out


def read_order(path, only_sheets):
    """Returns dict: lines, warns, scheme_rows, sheet_info, packs, names."""
    sheets = load_sheets(path)
    res = {"lines": [], "warns": [], "scheme_rows": [], "sheet_info": [], "packs": {}, "names": {}}
    catalog = {"packs": res["packs"], "names": res["names"]}
    for name, rows in sheets:
        if only_sheets and name.strip().lower() not in only_sheets:
            continue
        sl, sw, cutoff = read_order_sheet(name, rows, catalog)
        res["lines"].extend(sl)
        res["warns"].extend(sw)
        res["scheme_rows"].extend(read_scheme_area(name, rows, cutoff))
        if sl:
            res["sheet_info"].append((name.strip(), len(sl), sum(x["qty"] for x in sl)))
    if only_sheets and not any(n.strip().lower() in only_sheets for n, _ in sheets):
        raise CheckError("None of the sheets named in settings.ini 'order_sheets' exist in this file.\n"
                         "  Sheets in the file: %s" % ", ".join(n for n, _ in sheets))
    return res


# --------------------------------------------------------------------------------------
# SAM pending list
# --------------------------------------------------------------------------------------

def looks_like_pending(path):
    try:
        for _, rows in load_sheets(path):
            for row in rows[:15]:
                ks = {key(v) for v in row if isinstance(v, str)}
                if ("grcode" in ks or "groupcode" in ks) and any(k.startswith("gda") for k in ks):
                    return True
    except Exception:
        return False
    return False


def find_pending_file(cfg, order_path=None):
    if cfg.pending_file:
        p = cfg.pending_file if os.path.isabs(cfg.pending_file) else os.path.join(TOOL_DIR, cfg.pending_file)
        if not os.path.exists(p):
            raise CheckError("settings.ini pending_file points to a file that does not exist:\n  %s" % p)
        return p, "settings.ini"
    for p in excel_files_in(PENDING_DIR):
        if looks_like_pending(p):
            return p, "newest file in the Pending folder"
    if order_path:
        for p in excel_files_in(os.path.dirname(os.path.abspath(order_path)))[:25]:
            if os.path.abspath(p) != os.path.abspath(order_path) and looks_like_pending(p):
                return p, "found next to the order"
    return None, ""


def pending_columns(header):
    cols, sizes = {}, OrderedDict()
    for j, v in enumerate(header):
        s = parse_size(v)
        if s:
            sizes[j] = s
            continue
        k = key(v)
        if k in ("grcode", "groupcode"):
            cols["code"] = j
        elif k.startswith("cust") or "party" in k or "dealer" in k:
            cols.setdefault("cust", j)
        elif "date" in k:
            cols.setdefault("date", j)
        elif k.startswith("gda") or k in ("orderno", "pono", "sono", "docno"):
            cols.setdefault("gda", j)
        elif "model" in k or "desc" in k:
            cols.setdefault("model", j)
        elif k in ("total", "totalqty", "qty", "pendingqty", "balqty", "balance"):
            cols.setdefault("total", j)
    return cols, sizes


def pending_row_lines(row, cols, sizes):
    sz = OrderedDict()
    for j, s in sizes.items():
        q = to_number(get(row, j)) or 0
        if q > 0:
            sz[s] = sz.get(s, 0) + q
    total = to_number(get(row, cols.get("total"))) if "total" in cols else None
    if not sz and total and total > 0:
        sz["STD"] = total
    return sz


def read_pending_lines(path, customer_filter):
    """One open line per GDA x Gr Code x size."""
    lines = []
    for _, rows in load_sheets(path):
        hdr_i = next((i for i, row in enumerate(rows[:30])
                      if any(key(v) in ("grcode", "groupcode") for v in row if isinstance(v, str))), None)
        if hdr_i is None:
            continue
        cols, sizes = pending_columns(rows[hdr_i])
        for row in rows[hdr_i + 1:]:
            code = valid_code(get(row, cols["code"]))
            if not code:
                continue
            if customer_filter and customer_filter.lower() not in clean_text(get(row, cols.get("cust"))).lower():
                continue
            d = parse_date(get(row, cols.get("date"))) if "date" in cols else None
            for s, q in pending_row_lines(row, cols, sizes).items():
                lines.append({"kind": "GDA", "ref": ref_text(get(row, cols.get("gda"))) or "?", "date": d,
                              "code": code, "model": clean_text(get(row, cols.get("model"))), "size": s,
                              "ordered": q, "received": 0.0, "bills": []})
    return lines


# --------------------------------------------------------------------------------------
# SAM bills (format not fixed - found by headings)
# --------------------------------------------------------------------------------------

BILL_QTY_WORDS = {"qty", "quantity", "billedqty", "billedquantity", "nos", "pcs", "qtynos", "qtypcs",
                  "billqty", "invqty", "dispatchqty", "dispqty", "qtyinpcs", "qtyinnos"}


def is_bill_qty(k):
    return k in BILL_QTY_WORDS or (k.startswith("qty") and "rate" not in k)


def bill_columns(header):
    cols, sizes = {}, OrderedDict()
    for j, v in enumerate(header):
        k = key(v)
        s = parse_size(v) if isinstance(v, str) else None
        if s:
            sizes[j] = s
        elif is_bill_qty(k):
            cols.setdefault("qty", j)
        elif k in ("grcode", "code", "itemcode", "partno", "partcode", "modelcode", "groupcode"):
            cols.setdefault("code", j)
        elif k in ("size", "sizes", "os", "oversize"):
            cols.setdefault("size", j)
        elif "invoice" in k or "inv" == k[:3] or k in ("billno", "vchno", "docno", "voucherno"):
            if "date" in k:
                cols.setdefault("date", j)
            else:
                cols.setdefault("inv", j)
        elif "date" in k:
            cols.setdefault("date", j)
        elif any(w in k for w in ("desc", "item", "particular", "product", "model", "name", "material")):
            cols.setdefault("desc", j)
    return cols, sizes


def bill_header_row(rows):
    for i, row in enumerate(rows[:40]):
        ks = [key(v) for v in row if isinstance(v, str)]
        if any(is_bill_qty(k) for k in ks) or any(parse_size(v) == "STD" for v in row if isinstance(v, str)):
            return i
    return None


def bill_row_code(row, cols):
    """Gr Code from a code column, else from the item text ('GR0745 SHINE PISTON 0.25')."""
    texts = [get(row, cols["desc"])] if "desc" in cols else [v for v in row if isinstance(v, str)]
    desc = next((clean_text(t) for t in texts if t), "")
    code = valid_code(get(row, cols.get("code"))) if "code" in cols else ""
    if not code:
        for t in texts:
            found = codes_in_text(t)
            if found:
                return found[0], clean_text(t)
    return code, desc


def bill_row_lines(row, cols, sizes, desc):
    if sizes:
        return [(s, to_number(get(row, j)) or 0) for j, s in sizes.items()]
    size = parse_size(get(row, cols["size"])) if "size" in cols else None
    return [(size or size_in_name(desc) or "STD", qty_number(get(row, cols.get("qty"))))]


PDF_NUM_RE = re.compile(r"\d[\d,]*\.?\d*")


def pdf_pages(path):
    try:
        import pypdf
    except ImportError:
        raise CheckError("Reading PDF bills needs the free 'pypdf' package. Run Check_SAM_Order.bat once with "
                         "internet on, or type:  py -m pip install pypdf")
    try:
        return [(pg.extract_text() or "") for pg in pypdf.PdfReader(path).pages]
    except Exception as e:
        raise CheckError("cannot read the PDF (%s)" % e)


def pdf_qty_rate(numbers):
    """Find qty and rate among the numbers of an item line: qty x rate = one of the other numbers."""
    for i, q in enumerate(numbers):
        if q <= 0 or not float(q).is_integer():
            continue
        for j, r in enumerate(numbers):
            if j != i and r > 0 and any(abs(q * r - t) <= 0.6 for k, t in enumerate(numbers) if k not in (i, j)):
                return q, r
    return None, None


def pdf_item_line(line):
    """'Set 30 12404.70 248.09 11676.61 2 413.49 480.00GR0745 H.SHINE 0.75 WOSP' -> dict or None."""
    m = re.search(r"(?i)(?<![A-Z])GR\s*-?\s*([0-9O]{3,6})\s+(.*)$", line)
    if not m:
        return None
    before = line[:m.start()]
    nums = [float(x.replace(",", "")) for x in PDF_NUM_RE.findall(before)]
    qty, rate = pdf_qty_rate(nums)
    if not qty:
        return None
    desc = clean_text(m.group(2))
    moq = nums[-1] if before.rstrip() and before.rstrip()[-1].isdigit() and nums[-1] != rate else None
    return {"code": norm_code("GR" + m.group(1)), "desc": desc, "size": size_in_name(desc) or "STD",
            "qty": qty, "rate": rate, "moq_rs": moq}


def pdf_header(text, path):
    inv = re.search(r"(?<!\d)(\d{4}-\d{2}/\d{3,})(?!\d)", text)
    d = re.search(r"INVOICE\s*DATE\s*:?\s*(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})", text, re.I)
    gda = re.search(r"GDA\s*NO\.?\s*([\d,\s]+)", text, re.I)
    return (inv.group(1) if inv else "", (parse_date(d.group(1).replace(".", "/")) if d else None) or file_date(path),
            [g for g in re.split(r"[,\s]+", gda.group(1)) if g] if gda else [])


def read_bill_pdf(path):
    """A SAM PDF can hold several invoices (one or more pages each). Header is read per page;
    a page without its own invoice number continues the previous invoice."""
    out, inv, date, gdas = [], "", file_date(path), []
    for text in pdf_pages(path):
        p_inv, p_date, p_gdas = pdf_header(text, path)
        if p_inv:
            inv, date, gdas = p_inv, p_date, p_gdas
        for line in text.splitlines():
            it = pdf_item_line(line)
            if it:
                it.update(file=os.path.basename(path), inv=inv, date=date, gdas=list(gdas),
                          status="", against="", extra=0.0)
                out.append(it)
    return out


def read_bill_file(path):
    if path.lower().endswith(".pdf"):
        return read_bill_pdf(path)
    out = []
    fallback = file_date(path)
    for _, rows in load_sheets(path):
        hdr = bill_header_row(rows)
        if hdr is None:
            continue
        cols, sizes = bill_columns(rows[hdr])
        for row in rows[hdr + 1:]:
            code, desc = bill_row_code(row, cols)
            if not code:
                continue
            d = parse_date(get(row, cols.get("date"))) if "date" in cols else None
            for size, q in bill_row_lines(row, cols, sizes, desc):
                if q and q > 0:
                    out.append({"file": os.path.basename(path), "inv": ref_text(get(row, cols.get("inv"))),
                                "date": d or fallback, "code": code, "size": size, "qty": q, "desc": desc,
                                "rate": None, "moq_rs": None, "gdas": [], "status": "", "against": "", "extra": 0.0})
    return out


# --------------------------------------------------------------------------------------
# Tally Stock Summary
# --------------------------------------------------------------------------------------

def find_stock_file(cfg):
    if cfg.stock_file:
        p = cfg.stock_file if os.path.isabs(cfg.stock_file) else os.path.join(TOOL_DIR, cfg.stock_file)
        return (p, "settings.ini") if os.path.exists(p) else (None, "")
    files = excel_files_in(STOCK_DIR)
    return (files[0], "newest file in the Stock folder") if files else (None, "")


def stock_period_months(rows):
    for row in rows[:15]:
        for v in row:
            m = PERIOD_RE.search(str(v or ""))
            if m:
                a, b = parse_date(m.group(1)), parse_date(m.group(2))
                if a and b and b >= a:
                    return ((b - a).days + 1) / 30.44, "%s to %s" % (m.group(1), m.group(2))
    return None, ""


def stock_qty_columns(rows):
    """Column of 'Quantity' under 'Outwards' and under 'Closing Balance'."""
    for i, row in enumerate(rows[:30]):
        ks = [key(v) for v in row]
        out_c = next((j for j, k in enumerate(ks) if k.startswith("outward")), None)
        close_c = next((j for j, k in enumerate(ks) if k.startswith("closing")), None)
        if out_c is None or close_c is None or i + 1 >= len(rows):
            continue
        sub = [key(v) for v in rows[i + 1]]
        def qty_col(start):
            return next((j for j in range(start, len(sub)) if sub[j].startswith("quant")), start)
        return qty_col(out_c), qty_col(close_c), i + 2
    return None, None, None


def read_stock_file(path, cfg):
    """{(code, size): {'name','closing','outwards'}}, months, period text"""
    stock, months, period = {}, None, ""
    for _, rows in load_sheets(path):
        out_c, close_c, start = stock_qty_columns(rows)
        if out_c is None:
            continue
        m, p = stock_period_months(rows)
        months, period = months or m, period or p
        for row in rows[start:]:
            name = next((clean_text(v) for v in row if isinstance(v, str) and v.strip()), "")
            code = valid_code(name.split(" ")[0]) if name else ""
            if not code:
                continue
            k = (code, size_in_name(name) or "STD")
            s = stock.setdefault(k, {"name": name, "closing": 0.0, "outwards": 0.0})
            s["closing"] += qty_number(get(row, close_c))
            s["outwards"] += qty_number(get(row, out_c))
    if cfg.stock_months:
        months = cfg.stock_months
    return stock, months, period


# --------------------------------------------------------------------------------------
# Tracker: open with SAM = pending list + orders sent after it - SAM bills after it
# --------------------------------------------------------------------------------------

def line_balance(l):
    return l["ordered"] - l["received"]


def line_label(l):
    return ("GDA %s" % l["ref"]) if l["kind"] == "GDA" else ("PO %s" % l["ref"])


def line_status(l):
    if line_balance(l) <= 0.0001:
        return "DONE"
    return "PART" if l["received"] > 0 else "OPEN"


def po_lines(path):
    res = read_order(path, [])
    ref = os.path.splitext(os.path.basename(path))[0]
    d = file_date(path)
    out = []
    for ln in res["lines"]:
        for s, q in ln["sizes"].items():
            if q > 0:
                out.append({"kind": "PO", "ref": ref, "date": d, "code": ln["code"], "model": ln["model"],
                            "size": s, "ordered": q, "received": 0.0, "bills": []})
    return out


def add_orders_sent(t, skip_path, skip_hash):
    for p in sorted(excel_files_in(ORDERS_DIR), key=file_date):
        name = os.path.basename(p)
        if skip_path and (os.path.abspath(p) == os.path.abspath(skip_path) or file_sha256(p) == skip_hash):
            t["notes"].append("%s: this is the order being checked - not counted as already sent." % name)
            continue
        d = file_date(p)
        if t["orders_asof"] and d <= t["orders_asof"]:
            t["notes"].append("%s (dated %s): on/before the pending list's latest GDA date %s - assumed already "
                              "in SAM's list, skipped." % (name, ddmmyyyy(d), ddmmyyyy(t["orders_asof"])))
            continue
        try:
            new = po_lines(p)
        except Exception as e:
            t["warns"].append("Could not read order '%s' in 'Orders Sent' - skipped (%s)." % (name, str(e).split("\n")[0]))
            continue
        t["lines"].extend(new)
        t["orders_used"].append((name, d, sum(l["ordered"] for l in new)))


def allocate_bill(b, lines):
    named = set(b.get("gdas") or [])
    cands = sorted((l for l in lines if l["code"] == b["code"] and l["size"] == b["size"] and line_balance(l) > 0),
                   key=lambda l: (0 if (l["kind"] == "GDA" and l["ref"] in named) else 1,
                                  l["date"] or dt.date.min, 0 if l["kind"] == "GDA" else 1, l["ref"]))
    left, used = b["qty"], []
    for l in cands:
        if left <= 0:
            break
        take = min(left, line_balance(l))
        l["received"] += take
        l["bills"].append("%s %s: %s" % (b["inv"] or b["file"], ddmmyyyy(b["date"]), fmt_qty(take)))
        used.append("%s (%s)" % (line_label(l), fmt_qty(take)))
        left -= take
    b["against"] = ", ".join(used)
    b["extra"] = left
    b["status"] = "SET OFF" if left <= 0 else ("PART EXTRA" if used else "EXTRA")


def add_bills(t):
    bills = []
    for p in excel_files_in(BILLS_DIR, BILL_EXTS):
        try:
            got = read_bill_file(p)
        except Exception as e:
            t["warns"].append("Could not read bill '%s' in 'SAM Bills' - skipped (%s)."
                              % (os.path.basename(p), str(e).split("\n")[0]))
            continue
        if not got:
            t["warns"].append("Bill '%s': no lines with a Gr Code and quantity found - check its format."
                              % os.path.basename(p))
        bills.extend(got)
    for b in sorted(bills, key=lambda b: (b["date"], b["file"])):
        if t["bills_asof"] and b["date"] <= t["bills_asof"]:
            b["status"] = "SKIPPED"
            b["against"] = "dated on/before the pending list (%s) - already reflected" % ddmmyyyy(t["bills_asof"])
        else:
            allocate_bill(b, t["lines"])
    t["bills"] = bills


def bill_order_key(b):
    """Which SAM order a bill line belongs to: its GDA no(s), else its invoice."""
    return "GDA " + "/".join(b["gdas"]) if b.get("gdas") else "invoice " + (b["inv"] or b["file"])


def moq_bill_check(b, g, group_qty):
    """MOQ benefit SAM gave on a bill line vs. the scheme. The billed group total we can see is a MINIMUM
    of the GDA's real total, so benefit due on it is the least SAM owes: less than that = surely short-paid."""
    if b.get("moq_rs") is None:
        return "", 0.0
    per = b["moq_rs"] / b["qty"] if b["qty"] else 0
    if not g:
        return ("bonus Rs %s/pc (not a scheme model)" % fmt_qty(per)) if per > 0 else "", 0.0
    due = evaluate(group_qty, g)
    basis = "%s: '%s' billed %s pcs" % (bill_order_key(b), g.name, fmt_qty(group_qty))
    if abs(per - due["rate"]) < 0.01:
        what = ("Slab-%d Rs %s/pc" % (due["slab"], fmt_qty(per))) if due["slab"] else \
               "no benefit due (Slab-1 needs %s)" % fmt_qty(g.s1)
        return "OK - %s (%s)" % (what, basis), 0.0
    if per < due["rate"]:
        short = (due["rate"] - per) * b["qty"]
        return ("SHORT-PAID: Rs %s/pc given, Rs %s/pc due (%s) - claim Rs %s from SAM"
                % (fmt_qty(per), fmt_qty(due["rate"]), basis, inr(short)), short)
    # SAM counts the slab on the whole GDA; bills we have may show only part of it - so paying more is fine
    slab = "Slab-2" if abs(per - g.r2) < 0.01 else "Slab-1" if abs(per - g.r1) < 0.01 else "Rs %s/pc" % fmt_qty(per)
    return "OK - %s Rs %s/pc given (whole GDA counts; %s)" % (slab, fmt_qty(per), basis), 0.0


def check_bill_moq(t):
    try:
        groups = load_moq()
    except CheckError:
        return
    by_code = {c: g for g in groups for c in g.codes}
    totals = {}
    for b in t["bills"]:
        g = by_code.get(b["code"])
        if g:
            k = (bill_order_key(b), g.name)
            totals[k] = totals.get(k, 0) + b["qty"]
    t["moq_short"] = 0.0
    for b in t["bills"]:
        g = by_code.get(b["code"])
        b["moq_note"], short = moq_bill_check(b, g, totals.get((bill_order_key(b), g.name), 0) if g else 0)
        t["moq_short"] += short
        if short:
            t["warns"].append("Bill %s %s %s: %s" % (b["inv"] or b["file"], b["code"], b["size"], b["moq_note"]))


def add_pending(t, cfg, order_path):
    ppath, how = find_pending_file(cfg, order_path)
    if not ppath:
        t["warns"].append("No SAM pending list found - put SAM's latest pending-orders Excel in the folder:  %s"
                          % PENDING_DIR)
        return
    lines = read_pending_lines(ppath, cfg.pending_customer)
    t["lines"].extend(lines)
    dates = [l["date"] for l in lines if l["date"]]
    t["orders_asof"] = max(dates) if dates else None
    t["bills_asof"] = file_date(ppath)
    t["pending_path"] = ppath
    n_rows = len({(l["ref"], l["code"], l["date"]) for l in lines})
    t["pending_desc"] = "%s  (%d lines, %s pcs, latest GDA %s, %s)" % (
        os.path.basename(ppath), n_rows, inr(sum(l["ordered"] for l in lines)), ddmmyyyy(t["orders_asof"]), how)
    age = (dt.date.today() - dt.date.fromtimestamp(os.path.getmtime(ppath))).days
    if cfg.pending_max_age and age > cfg.pending_max_age:
        t["warns"].append("Pending list '%s' is %d days old - get SAM's latest pending list." % (os.path.basename(ppath), age))
    if not lines:
        t["warns"].append("Pending list has no usable lines%s."
                          % (" for customer '%s'" % cfg.pending_customer if cfg.pending_customer else ""))


def build_tracker(cfg, order_path=None, order_hash=None):
    t = {"lines": [], "bills": [], "notes": [], "warns": [], "orders_used": [], "pending_path": None,
         "pending_desc": "NOT FOUND", "orders_asof": None, "bills_asof": None}
    add_pending(t, cfg, order_path)
    add_orders_sent(t, order_path, order_hash)
    add_bills(t)
    check_bill_moq(t)
    t["lines"].sort(key=lambda l: (l["date"] or dt.date.min, l["kind"], l["ref"], l["code"], size_sort_key(l["size"])))
    return t


def tracker_open_by_code(t):
    out = OrderedDict()
    for l in t["lines"]:
        bal = line_balance(l)
        if bal <= 0:
            continue
        p = out.setdefault(l["code"], {"qty": 0.0, "sizes": OrderedDict(), "gdas": OrderedDict(), "model": ""})
        p["qty"] += bal
        p["sizes"][l["size"]] = p["sizes"].get(l["size"], 0) + bal
        k = (line_label(l), l["date"])
        p["gdas"][k] = p["gdas"].get(k, 0) + bal
        p["model"] = p["model"] or l["model"]
    return out


def tracker_open_by_key(t):
    out = {}
    for l in t["lines"]:
        if line_balance(l) > 0:
            k = (l["code"], l["size"])
            out[k] = out.get(k, 0) + line_balance(l)
    return out


def tracker_totals(t):
    return (sum(l["ordered"] for l in t["lines"]), sum(l["received"] for l in t["lines"]),
            sum(max(0, line_balance(l)) for l in t["lines"]))


# --------------------------------------------------------------------------------------
# MOQ check
# --------------------------------------------------------------------------------------

def evaluate(qty, g):
    slab, rate = (2, g.r2) if qty >= g.s2 else ((1, g.r1) if qty >= g.s1 else (0, 0))
    earned = qty * rate
    short2 = 0 if slab == 2 else g.s2 - qty
    short1 = g.s1 - qty if slab == 0 else 0
    lost = 0 if slab == 2 else g.s2 * g.r2 - earned
    if qty <= 0:
        colour, status = "RED", "NOT IN THIS ORDER (needs %s for Slab-2)" % fmt_qty(g.s2)
    elif slab == 2:
        colour, status = "GREEN", "OK - Slab-2 hit"
    elif slab == 1:
        colour, status = "RED", "Slab-1 only - add %s pcs for Slab-2" % fmt_qty(short2)
    else:
        colour, status = "RED", "Below Slab-1 - add %s pcs for Slab-1, %s for Slab-2" % (fmt_qty(short1), fmt_qty(short2))
    return {"qty": qty, "s1": g.s1, "r1": g.r1, "s2": g.s2, "r2": g.r2, "slab": slab, "rate": rate,
            "earned": earned, "short1": short1, "short2": short2, "lost": lost,
            "colour": colour, "status": status, "ordered": qty > 0}


def scheme_row_notes(s, by_code, where):
    notes = []
    known = [c for c in s["codes"] if c in by_code]
    unknown = [c for c in s["codes"] if c not in by_code]
    gs = []
    for c in known:
        if by_code[c] not in gs:
            gs.append(by_code[c])
    for g in gs:
        if (g.s1, g.r1, g.s2, g.r2) != (s["s1"], s["r1"], s["s2"], s["r2"]):
            notes.append("%s says %s: Slab-1 %s @ Rs %s, Slab-2 %s @ Rs %s - but moq.csv '%s' has %s @ Rs %s, %s @ Rs %s."
                         % (where, "/".join(s["codes"]), fmt_qty(s["s1"]), fmt_qty(s["r1"]), fmt_qty(s["s2"]),
                            fmt_qty(s["r2"]), g.name, fmt_qty(g.s1), fmt_qty(g.r1), fmt_qty(g.s2), fmt_qty(g.r2)))
    if len(gs) > 1:
        notes.append("%s counts %s together, but moq.csv splits them: %s." % (where, "/".join(known), "; ".join(g.name for g in gs)))
    elif gs and not unknown and set(gs[0].codes) - set(known):
        notes.append("%s counts %s on its own, but moq.csv group '%s' also adds %s."
                     % (where, "/".join(known), gs[0].name, "/".join(sorted(set(gs[0].codes) - set(known)))))
    return notes


def check_scheme_vs_po(scheme_rows, groups, ignore=()):
    """Compare moq.csv with the MOQ summary SAM prints inside its own PO."""
    by_code = {c: g for g in groups for c in g.codes}
    notes = []
    for s in scheme_rows:
        if set(s["codes"]) & set(ignore):
            continue
        where = "SAM's MOQ summary in the order (sheet '%s' row %d)" % (s["sheet"], s["row"])
        if not any(c in by_code for c in s["codes"]):
            notes.append("%s lists %s - none of these codes are in moq.csv." % (where, "/".join(s["codes"])))
            continue
        notes.extend(scheme_row_notes(s, by_code, where))
    return notes


def collect_models(lines):
    models = OrderedDict()
    for ln in lines:
        m = models.setdefault(ln["code"], {"code": ln["code"], "model": "", "sizes": OrderedDict(), "qty": 0.0,
                                           "lines": [], "notes": [], "value": 0.0, "priced_qty": 0.0, "rate_src": []})
        m["lines"].append(ln)
        m["qty"] += ln["qty"]
        if ln["model"] and ln["model"] not in m["model"]:
            m["model"] = (m["model"] + " / " + ln["model"]) if m["model"] else ln["model"]
        for s, q in ln["sizes"].items():
            m["sizes"][s] = m["sizes"].get(s, 0) + q
        m["notes"].extend("%s row %d: %s" % (ln["sheet"], ln["row"], n) for n in ln["notes"])
        if ln["rate"] or ln["amount"]:
            m["value"] += ln["qty"] * ln["rate"] if ln["rate"] else ln["amount"]
            m["priced_qty"] += ln["qty"]
            src = "order file (%s)" % (ln["rate_label"] or "rate") if ln["rate"] else "order file (amount)"
            if src not in m["rate_src"]:
                m["rate_src"].append(src)
    return models


def price_models(models, rates, rate_labels, by_code, pending):
    for code, m in models.items():
        g = by_code.get(code)
        m["group"] = g.name if g else ""
        m["model"] = m["model"] or rate_labels.get(code) or pending.get(code, {}).get("model", "")
        rest, m["unpriced_qty"] = m["qty"] - m["priced_qty"], 0.0
        if rest > 0.0001:
            if code in rates:
                m["value"] += rest * rates[code]
                m["rate_src"].append("rates.csv")
            else:
                m["unpriced_qty"] = rest
        priced = m["qty"] - m["unpriced_qty"]
        m["rate"] = m["value"] / priced if priced > 0 else None
        m["pending"] = pending.get(code)


def group_codes_text(g, models):
    ordered = [c for c in g.codes if c in models]
    if len(g.codes) <= 4 or not ordered:
        return " | ".join("%s: %s" % (c, fmt_qty(models[c]["qty"]) if c in models else "0") for c in g.codes)
    return (" | ".join("%s: %s" % (c, fmt_qty(models[c]["qty"])) for c in ordered)
            + "  (+%d other codes: 0)" % (len(g.codes) - len(ordered)))


def moq_units(groups, models, rate_labels, cfg):
    units = []
    for g in groups:
        if cfg.mode == "group":
            qty = sum(models[c]["qty"] for c in g.codes if c in models)
            if qty > 0 or cfg.show_not_ordered:
                u = evaluate(qty, g)
                u.update(name=g.name, group=g.name, model="", codes=group_codes_text(g, models))
                units.append(u)
            continue
        for c in g.codes:
            qty = models[c]["qty"] if c in models else 0
            if qty > 0 or cfg.show_not_ordered:
                u = evaluate(qty, g)
                u.update(name=c, group=g.name, codes=c,
                         model=models[c]["model"] if c in models else rate_labels.get(c, ""))
                units.append(u)
    return units


def analyse(lines, groups, rates, rate_labels, pending, cfg):
    by_code = {c: g for g in groups for c in g.codes}
    models = collect_models(lines)
    price_models(models, rates, rate_labels, by_code, pending)
    units = moq_units(groups, models, rate_labels, cfg)
    ordered = [u for u in units if u["ordered"]]
    total_value = sum(m["value"] for m in models.values())
    return {"models": models, "units": units, "by_code": by_code,
            "total_qty": sum(m["qty"] for m in models.values()), "total_value": total_value,
            "unpriced": [m for m in models.values() if m["unpriced_qty"] > 0],
            "earned": sum(u["earned"] for u in units), "lost": sum(u["lost"] for u in ordered),
            "n_green": sum(1 for u in ordered if u["slab"] == 2), "n_orange": sum(1 for u in ordered if u["slab"] == 1),
            "n_red": sum(1 for u in ordered if u["slab"] == 0),
            "n_not_ordered": sum(1 for u in units if not u["ordered"]),
            "not_in_scheme": [m for m in models.values() if not m["group"]],
            "pending_hits": [m for m in models.values() if m["pending"]],
            "over_cap": total_value > cfg.weekly_cap}


# --------------------------------------------------------------------------------------
# Stock cover:  have = stock in hand + open with SAM + this order;  cover = have / monthly sales
# --------------------------------------------------------------------------------------

def cover_row(key_, ordered, st, open_qty, months, pack, cfg, name):
    closing = max(0.0, st["closing"]) if st else 0.0
    monthly = (st["outwards"] / months) if st and months else 0.0
    have = closing + open_qty + ordered
    r = {"code": key_[0], "size": key_[1], "name": name, "pack": pack, "stock": closing, "open": open_qty,
         "ordered": ordered, "have": have, "monthly": monthly, "cover": None, "status": "OK", "action": "", "qty": 0,
         "note": "" if st else "not in stock file"}
    if monthly <= 0:
        r["status"] = "NO SALE" if ordered > 0 else ""
        return r
    r["cover"] = have / monthly
    if r["cover"] < cfg.min_cover:
        r["qty"] = ceil_to(cfg.target_cover * monthly - have, pack)
        r["status"], r["action"] = "SHORT", "ADD %s pcs" % fmt_qty(r["qty"])
    elif r["cover"] > cfg.max_cover and ordered > 0:
        cut = min(ordered, floor_to(have - cfg.target_cover * monthly, pack))
        if cut > 0:
            r["qty"] = cut
            r["status"], r["action"] = "OVER", "CUT %s pcs from this order" % fmt_qty(cut)
    return r


def cut_moq_notes(rows, res):
    for r in rows:
        g = res["by_code"].get(r["code"])
        if r["status"] != "OVER" or not g:
            continue
        gq = sum(res["models"][c]["qty"] for c in g.codes if c in res["models"])
        if gq >= g.s2 > gq - r["qty"]:
            r["note"] = ("careful: this cut alone drops '%s' below Slab-2 (%s -> %s, needs %s)"
                         % (g.name, fmt_qty(gq), fmt_qty(gq - r["qty"]), fmt_qty(g.s2)))


def stock_cover(res, catalog, stock, months, open_key, cfg):
    order_keys = {(c, s): q for c, m in res["models"].items() for s, q in m["sizes"].items() if q > 0}
    keys = set(order_keys)
    if cfg.stock_all:
        keys |= {k for k, st in stock.items() if st["outwards"] > 0}
    rows = []
    for k in sorted(keys, key=lambda k: (k[0], size_sort_key(k[1]))):
        name = (res["models"].get(k[0], {}).get("model") or catalog["names"].get(k[0])
                or (stock.get(k) or {}).get("name", ""))
        r = cover_row(k, order_keys.get(k, 0.0), stock.get(k), open_key.get(k, 0.0), months,
                      catalog["packs"].get(k[0], 1.0), cfg, name)
        if r["ordered"] > 0 or r["status"] == "SHORT":
            rows.append(r)
    cut_moq_notes(rows, res)
    return rows


# --------------------------------------------------------------------------------------
# "What to fix before sending"
# --------------------------------------------------------------------------------------

def moq_actions(res, cfg):
    acts = []
    bad = [u for u in res["units"] if u["ordered"] and u["slab"] < 2]
    for u in sorted(bad, key=lambda u: -u["lost"]):
        what = u["name"] if cfg.mode == "group" else "%s (%s)" % (u["name"], u["group"])
        if u["slab"] == 0:
            acts.append(("RED", "%s: ordered %s - BELOW Slab-1, earns Rs 0. Add %s pcs for Slab-1 (Rs %s) or %s pcs "
                                "for Slab-2 (Rs %s)." % (what, fmt_qty(u["qty"]), fmt_qty(u["short1"]),
                                                         inr(u["s1"] * u["r1"]), fmt_qty(u["short2"]), inr(u["s2"] * u["r2"]))))
        else:
            acts.append(("RED", "%s: ordered %s - Slab-1 only. Add %s pcs to reach Slab-2 (+Rs %s benefit)."
                         % (what, fmt_qty(u["qty"]), fmt_qty(u["short2"]), inr(u["lost"]))))
    return acts


def short_list(rows, n=6):
    return ", ".join("%s %s %s%s" % (r["code"], r["size"], "+" if r["status"] == "SHORT" else "-", fmt_qty(r["qty"]))
                     for r in sorted(rows, key=lambda r: -r["qty"])[:n]) + (" ..." if len(rows) > n else "")


def stock_actions(cover, cfg):
    acts = []
    short_in = [r for r in cover if r["status"] == "SHORT" and r["ordered"] > 0]
    short_out = [r for r in cover if r["status"] == "SHORT" and r["ordered"] == 0]
    over = [r for r in cover if r["status"] == "OVER"]
    nosale = [r for r in cover if r["status"] == "NO SALE"]
    if short_in:
        acts.append(("STOCK", "STOCK SHORT: %d size(s) in this order still under %s month cover - ADD: %s"
                     % (len(short_in), fmt_qty(cfg.min_cover), short_list(short_in))))
    if short_out:
        acts.append(("STOCK", "NOT ORDERED but SHORT: %d selling size(s) - consider adding: %s"
                     % (len(short_out), short_list(short_out))))
    if over:
        acts.append(("OVER", "STOCK OVER: %d size(s) give more than %s months cover - CUT: %s"
                     % (len(over), fmt_qty(cfg.max_cover), short_list(over))))
    if nosale:
        acts.append(("NOSALE", "NO SALE: %d size(s) ordered with no sales in the stock period: %s"
                     % (len(nosale), ", ".join("%s %s (%s)" % (r["code"], r["size"], fmt_qty(r["ordered"])) for r in nosale[:8])
                        + (" ..." if len(nosale) > 8 else ""))))
    return acts


def build_actions(res, cover, cfg):
    acts = moq_actions(res, cfg)
    hits = res["pending_hits"]
    if hits:
        top = sorted(hits, key=lambda m: -m["pending"]["qty"])[:4]
        acts.append(("PENDING", "%d model(s) in this order are still OPEN with SAM - see 'Open with SAM'. Biggest: %s."
                     % (len(hits), ", ".join("%s %s pcs" % (m["code"], fmt_qty(m["pending"]["qty"])) for m in top))))
    acts.extend(stock_actions(cover, cfg))
    packs = [m for m in res["models"].values() if any("pack size" in n for n in m["notes"])]
    if packs:
        acts.append(("PACK", "%d model(s) have sizes not in multiples of the pack size: %s"
                     % (len(packs), ", ".join(m["code"] for m in packs))))
    if res["over_cap"]:
        acts.append(("CAP", "Order value Rs %s is ABOVE the weekly cap of Rs %s by Rs %s."
                     % (inr(res["total_value"]), inr(cfg.weekly_cap), inr(res["total_value"] - cfg.weekly_cap))))
    if res["unpriced"]:
        codes = [m["code"] for m in res["unpriced"]]
        acts.append(("RATE", "%d model(s) have no rate - order value is INCOMPLETE. Add to rates.csv: %s"
                     % (len(codes), ", ".join(codes[:10]) + (" ..." if len(codes) > 10 else ""))))
    return acts


# --------------------------------------------------------------------------------------
# Excel output
# --------------------------------------------------------------------------------------

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=16, color="1F3864")
SECTION_FONT = Font(bold=True, size=12, color="1F3864")
BOLD = Font(bold=True)
ROW_FILL = {"GREEN": "E2F0D9", "ORANGE": "FCE4D6", "RED": "FFE1E1", "GREY": "F2F2F2", "YELLOW": "FFF2CC"}
STATUS_FILL = {"GREEN": "00B050", "ORANGE": "ED7D31", "RED": "E00000"}
ACTION_FILL = {"RED": "FFE1E1", "PENDING": "FFF2CC", "CAP": "FFE1E1", "RATE": "FCE4D6", "STOCK": "FFE1E1",
               "OVER": "FCE4D6", "NOSALE": "FFF2CC", "PACK": "FFF2CC"}
STOCK_FILL = {"SHORT": "FFE1E1", "OVER": "FCE4D6", "NO SALE": "FFF2CC", "OK": "E2F0D9"}
INR_FMT = r'[>=10000000]##\,##\,##\,##0;[>=100000]##\,##\,##0;##,##0'
QTY_FMT = "#,##0"
RATE_FMT = "#,##0.00"
COVER_FMT = "0.0"


def fill(hex_colour):
    return PatternFill("solid", fgColor=hex_colour)


def put(ws, r, c, value, fmt=None, font=None, bg=None, border=False, wrap=False, align=None):
    cell = ws.cell(row=r, column=c, value=value)
    if fmt:
        cell.number_format = fmt
    if font:
        cell.font = font
    if bg:
        cell.fill = fill(bg)
    if border:
        cell.border = BORDER
    if wrap or align or border:
        cell.alignment = Alignment(wrap_text=wrap, vertical="center", horizontal=align)
    return cell


def header_row(ws, r, titles, height=32):
    for c, t in enumerate(titles, start=1):
        cell = put(ws, r, c, t, font=HEAD_FONT, border=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[r].height = height


def fill_row(ws, r, first, last, colour):
    for c in range(first, last + 1):
        ws.cell(row=r, column=c).fill = fill(colour)


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def landscape(ws):
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def table(ws, r, titles, rows, fmts, colours=None, wrap_cols=(), height=30):
    """Header + rows; returns next free row."""
    header_row(ws, r, titles, height)
    for n, vals in enumerate(rows):
        r += 1
        bg = colours[n] if colours else None
        for c, (v, f) in enumerate(zip(vals, fmts), start=1):
            put(ws, r, c, v, fmt=f, border=True, bg=bg, wrap=c in wrap_cols)
    return r + 1


def save_workbook(wb, path):
    try:
        wb.save(path)
        return path
    except PermissionError:
        alt = path[:-5] + "_" + dt.datetime.now().strftime("%H%M%S") + ".xlsx"
        print("  (The old file is open in Excel - saving as %s instead.)" % os.path.basename(alt))
        wb.save(alt)
        return alt


# ---- Summary sheet -------------------------------------------------------------------

def summary_info(ws, ctx, r):
    cfg = ctx["cfg"]
    info = [("Order file", ctx["order_name"]),
            ("Sheets with quantities", ", ".join("%s (%d lines, %s pcs)" % (n, k, fmt_qty(q)) for n, k, q in ctx["sheet_info"])),
            ("MOQ scheme", "%s%d groups from moq.csv" % ((cfg.scheme_name + " - ") if cfg.scheme_name else "", len(ctx["groups"]))),
            ("MOQ counted", "Per GROUP - all sizes of all models in a group added together" if cfg.mode == "group"
             else "Per MODEL - each Gr Code must reach the MOQ on its own"),
            ("SAM pending list", ctx["tracker"]["pending_desc"]),
            ("Orders sent added", ", ".join("%s (%s pcs)" % (n, fmt_qty(q)) for n, _, q in ctx["tracker"]["orders_used"]) or "none"),
            ("SAM bills set off", "%d bill line(s)" % sum(1 for b in ctx["tracker"]["bills"] if b["status"] not in ("SKIPPED",))),
            ("Stock file (Tally)", ctx["stock_desc"])]
    for label, value in info:
        put(ws, r, 1, label, font=BOLD)
        put(ws, r, 2, value)
        r += 1
    return r + 1


def value_comment(res, cfg):
    if res["over_cap"]:
        txt, bg = "ABOVE weekly cap of Rs %s by Rs %s" % (inr(cfg.weekly_cap), inr(res["total_value"] - cfg.weekly_cap)), ROW_FILL["RED"]
    else:
        txt, bg = "Within weekly cap of Rs %s" % inr(cfg.weekly_cap), ROW_FILL["GREEN"]
    if res["unpriced"]:
        txt = ((txt + "  |  ") if res["over_cap"] else "") + "INCOMPLETE: %d model(s) have no rate%s" % (
            len(res["unpriced"]), " (real value is even higher)" if res["over_cap"] else " - cap cannot be confirmed")
        bg = ROW_FILL["RED"] if res["over_cap"] else ROW_FILL["ORANGE"]
    return txt, bg


def summary_key_numbers(ws, ctx, r):
    cfg, res, cover = ctx["cfg"], ctx["res"], ctx["cover"]
    put(ws, r, 1, "KEY NUMBERS", font=SECTION_FONT)
    r += 1
    vtxt, vbg = value_comment(res, cfg)
    n = {s: sum(1 for x in cover if x["status"] == s) for s in ("SHORT", "OVER", "NO SALE")}
    rows = [
        ("Total pieces ordered", res["total_qty"], QTY_FMT, "%d models" % len(res["models"]), None),
        ("Order value Rs (incl GST)", res["total_value"], INR_FMT, vtxt, vbg),
        ("MOQ benefit earned Rs", res["earned"], INR_FMT, "", ROW_FILL["GREEN"]),
        ("MOQ benefit LOST Rs", res["lost"], INR_FMT, "fix the RED groups below" if res["lost"] else "nothing lost",
         ROW_FILL["RED"] if res["lost"] else ROW_FILL["GREEN"]),
        ("Groups: Slab-2 / Slab-1 only / below Slab-1", "%d / %d / %d" % (res["n_green"], res["n_orange"], res["n_red"]),
         None, ("%d group(s) not in this order" % res["n_not_ordered"]) if res["n_not_ordered"] else "", None),
        ("Models NOT in MOQ scheme", len(res["not_in_scheme"]), None,
         "%s pcs - no MOQ benefit" % fmt_qty(sum(m["qty"] for m in res["not_in_scheme"])), None),
        ("Models still OPEN with SAM", len(res["pending_hits"]), None, "see 'Open with SAM' list",
         ROW_FILL["YELLOW"] if res["pending_hits"] else None),
        ("Stock cover: short / over / no sale", "%d / %d / %d" % (n["SHORT"], n["OVER"], n["NO SALE"]), None,
         "see 'Stock cover' sheet" if ctx["stock_used"] else ctx["stock_desc"],
         None if not ctx["cfg"].stock_check else ROW_FILL["RED"] if n["SHORT"] or not ctx["stock_used"] else ROW_FILL["GREEN"]),
    ]
    for label, value, fmt, comment, bg in rows:
        put(ws, r, 1, label, font=BOLD, border=True)
        put(ws, r, 2, value, fmt=fmt, font=Font(bold=True, size=12), border=True, align="left")
        put(ws, r, 3, comment)
        if bg:
            fill_row(ws, r, 3, 11, bg)
        r += 1
    return r + 1


def summary_actions(ws, ctx, r):
    put(ws, r, 1, "WHAT TO FIX BEFORE SENDING", font=SECTION_FONT)
    r += 1
    if not ctx["actions"]:
        put(ws, r, 1, "Nothing - all good.", font=Font(bold=True, color="006100"))
        return r + 2
    for n, (tag, text) in enumerate(ctx["actions"], start=1):
        put(ws, r, 1, "%d. %s" % (n, text))
        fill_row(ws, r, 1, 11, ACTION_FILL.get(tag, "FFFFFF"))
        r += 1
    return r + 1


def summary_moq(ws, ctx, r):
    cfg = ctx["cfg"]
    put(ws, r, 1, "MOQ CHECK - per %s" % ("GROUP" if cfg.mode == "group" else "MODEL"), font=SECTION_FONT)
    titles = (["MOQ Group", "Gr Codes : qty ordered"] if cfg.mode == "group" else ["Gr Code - Model", "MOQ Group"]) + [
        "Qty ordered", "Slab-1 MOQ", "Slab-2 MOQ", "Slab hit", "Rs / pc", "Benefit earned Rs", "Pcs short of Slab-2",
        "Benefit lost Rs", "Status"]
    rows, colours = [], []
    for u in ctx["res"]["units"]:
        first = u["name"] if cfg.mode == "group" else ("%s  %s" % (u["name"], u["model"])).strip()
        second = u["codes"] if cfg.mode == "group" else u["group"]
        rows.append([first, second, u["qty"], u["s1"], u["s2"], ("Slab-%d" % u["slab"]) if u["slab"] else "None",
                     u["rate"], u["earned"], u["short2"] or None, u["lost"] or None, u["status"]])
        colours.append(ROW_FILL[u["colour"]] if u["ordered"] else ROW_FILL["GREY"])
    fmts = [None, None, QTY_FMT, QTY_FMT, QTY_FMT, None, QTY_FMT, INR_FMT, QTY_FMT, INR_FMT, None]
    start = r + 1
    end = table(ws, start, titles, rows, fmts, colours, wrap_cols=(1, 2))
    for n, u in enumerate(ctx["res"]["units"]):
        st = ws.cell(row=start + 1 + n, column=11)
        if u["ordered"]:
            st.fill = fill(STATUS_FILL[u["colour"]])
            st.font = Font(bold=True, color="FFFFFF")
        else:
            st.font = Font(bold=True, color="C00000")
    put(ws, end, 1, "GREEN = Slab-2 hit.  RED = Slab-2 NOT hit.  Grey = not in this order.  Benefit lost = "
                    "(Slab-2 MOQ x Slab-2 Rs/pc) - benefit earned now.", font=Font(italic=True, size=9))
    return end + 2


def summary_not_in_scheme(ws, ctx, r):
    put(ws, r, 1, "MODELS NOT IN THE MOQ SCHEME (no MOQ benefit)", font=SECTION_FONT)
    ms = ctx["res"]["not_in_scheme"]
    if not ms:
        put(ws, r + 1, 1, "None.")
        return r + 3
    rows = [[m["code"], m["model"], m["qty"], m["value"] or None] for m in ms]
    return table(ws, r + 1, ["Gr Code", "Model", "Qty ordered", "Value Rs"], rows, [None, None, QTY_FMT, INR_FMT],
                 [ROW_FILL["GREY"]] * len(rows), height=20) + 1


def pending_row(m, today):
    p = m["pending"]
    dates = [d for (_, d) in p["gdas"] if d]
    oldest = min(dates) if dates else None
    refs = ", ".join("%s (%s): %s" % (g, d.strftime("%d-%m") if d else "?", fmt_qty(q))
                     for (g, d), q in sorted(p["gdas"].items(), key=lambda kv: (kv[0][1] or dt.date.min)))
    sizes = ", ".join("%s: %s" % (s, fmt_qty(q)) for s, q in sorted(p["sizes"].items(), key=lambda kv: size_sort_key(kv[0])))
    return ["%s  %s" % (m["code"], m["model"]), m["group"] or "not in scheme", m["qty"], p["qty"], sizes,
            ddmmyyyy(oldest), (today - oldest).days if oldest else None, refs]


def summary_pending(ws, ctx, r):
    put(ws, r, 1, "STILL OPEN WITH SAM (pending list + orders sent - bills)", font=SECTION_FONT)
    hits = ctx["res"]["pending_hits"]
    if not hits:
        put(ws, r + 1, 1, "None of the models in this order are open with SAM.")
        return r + 3
    rows = [pending_row(m, ctx["now"].date()) for m in hits]
    return table(ws, r + 1, ["Gr Code - Model", "MOQ Group", "Ordering now", "Open with SAM", "Open sizes",
                             "Oldest", "Days old", "GDA / PO (date): pcs"], rows,
                 [None, None, QTY_FMT, QTY_FMT, None, None, QTY_FMT, None],
                 [ROW_FILL["YELLOW"]] * len(rows), wrap_cols=(1, 2, 5, 8)) + 1


def summary_warnings(ws, ctx, r):
    put(ws, r, 1, "WARNINGS & NOTES", font=SECTION_FONT)
    r += 1
    for w in ctx["warnings"] or ["None."]:
        put(ws, r, 1, ("! " + w) if ctx["warnings"] else w, font=Font(color="9C0006"))
        r += 1
    put(ws, r + 1, 1, "Order file was only read, never changed (checked: %s)." % ctx["untouched"],
        font=Font(italic=True, size=9, color="666666"))


def write_summary(ws, ctx):
    set_widths(ws, [40, 34, 11, 10, 10, 10, 9, 13, 11, 13, 48])
    put(ws, 1, 1, "SAM ORDER CHECK", font=TITLE_FONT)
    put(ws, 1, 4, ctx["cfg"].company, font=Font(bold=True, size=12))
    put(ws, 2, 1, "Checked on %s" % ctx["now"].strftime("%d-%m-%Y %H:%M"), font=Font(italic=True, color="666666"))
    r = summary_info(ws, ctx, 4)
    r = summary_key_numbers(ws, ctx, r)
    r = summary_actions(ws, ctx, r)
    r = summary_moq(ws, ctx, r)
    r = summary_not_in_scheme(ws, ctx, r)
    r = summary_pending(ws, ctx, r)
    summary_warnings(ws, ctx, r)
    ws.sheet_view.zoomScale = 90
    landscape(ws)


# ---- Detail sheet --------------------------------------------------------------------

def detail_values(m, size_cols):
    p = m["pending"]
    notes = list(m["notes"])
    if m["unpriced_qty"]:
        notes.insert(0, "NO RATE for %s pcs - add to rates.csv" % fmt_qty(m["unpriced_qty"]))
    if not m["group"]:
        notes.insert(0, "NOT in MOQ scheme")
    oldest = min((d for (_, d) in p["gdas"] if d), default=None) if p else None
    return ([m["code"], m["model"], m["group"] or "NOT IN SCHEME"] + [m["sizes"].get(s) for s in size_cols] +
            [m["qty"], m["rate"], ", ".join(m["rate_src"]) or "-", m["value"] or None, p["qty"] if p else None,
             ", ".join(g for (g, _) in p["gdas"]) if p else None, ddmmyyyy(oldest) or None,
             "; ".join("%s r%d" % (ln["sheet"], ln["row"]) for ln in m["lines"]), "; ".join(notes)])


def write_detail(ws, ctx):
    models = ctx["res"]["models"]
    extra = sorted({s for m in models.values() for s in m["sizes"] if s not in STANDARD_SIZES and s != "Other"},
                   key=size_sort_key)
    size_cols = STANDARD_SIZES + extra + (["Other"] if any("Other" in m["sizes"] for m in models.values()) else [])
    titles = (["Gr Code", "Model (from order)", "MOQ Group"] + size_cols +
              ["Total Qty", "Rate Rs", "Rate from", "Value Rs", "Open with SAM", "Open GDA / PO",
               "Oldest open", "Read from (sheet row)", "Notes"])
    fmts = [None, None, None] + [QTY_FMT] * len(size_cols) + [QTY_FMT, RATE_FMT, None, INR_FMT, QTY_FMT, None, None, None, None]
    rows = [detail_values(m, size_cols) for m in models.values()]
    colours = [ROW_FILL["GREY"] if not m["group"] else None for m in models.values()]
    end = table(ws, 1, titles, rows, fmts, colours, height=34)
    put(ws, end, 1, "TOTAL", font=BOLD, border=True)
    for c, name in enumerate(titles, start=1):
        if name in size_cols or name in ("Total Qty", "Value Rs"):
            col = get_column_letter(c)
            put(ws, end, c, "=SUM(%s2:%s%d)" % (col, col, end - 1), fmt=INR_FMT if name == "Value Rs" else QTY_FMT,
                font=BOLD, border=True, bg="D9E1F2")
    widths = {"Gr Code": 10, "Model (from order)": 38, "MOQ Group": 30, "Total Qty": 9, "Rate Rs": 10, "Rate from": 16,
              "Value Rs": 12, "Open with SAM": 10, "Open GDA / PO": 22, "Oldest open": 11,
              "Read from (sheet row)": 28, "Notes": 60}
    set_widths(ws, [widths.get(t, 7) for t in titles])
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(titles)), max(end - 1, 1))
    landscape(ws)


# ---- Stock cover sheet ---------------------------------------------------------------

def write_stock_sheet(ws, ctx):
    cfg = ctx["cfg"]
    put(ws, 1, 1, "STOCK COVER - have = stock in hand + open with SAM + this order;  cover = have / monthly sales",
        font=SECTION_FONT)
    put(ws, 2, 1, "%s.  SHORT below %s month -> add to reach %s months (rounded up to pack).  OVER above %s months -> "
                  "cut to %s months (rounded down to pack)." % (ctx["stock_desc"], fmt_qty(cfg.min_cover),
                                                               fmt_qty(cfg.target_cover), fmt_qty(cfg.max_cover),
                                                               fmt_qty(cfg.target_cover)), font=Font(italic=True, size=9))
    titles = ["Gr Code", "Item", "Size", "Pack", "Stock in hand", "Open with SAM", "This order", "Have",
              "Monthly sales", "Cover (months)", "Status", "Action", "MOQ group", "Note"]
    fmts = [None, None, None, QTY_FMT, QTY_FMT, QTY_FMT, QTY_FMT, QTY_FMT, COVER_FMT, COVER_FMT, None, None, None, None]
    order = {"SHORT": 0, "OVER": 1, "NO SALE": 2, "OK": 3}
    rows_ = sorted(ctx["cover"], key=lambda r: (order.get(r["status"], 4), r["code"], size_sort_key(r["size"])))
    rows = [[r["code"], r["name"], r["size"], r["pack"], r["stock"], r["open"], r["ordered"] or None, r["have"],
             r["monthly"], r["cover"], r["status"], r["action"],
             getattr(ctx["res"]["by_code"].get(r["code"]), "name", ""), r["note"]] for r in rows_]
    colours = [STOCK_FILL.get(r["status"]) for r in rows_]
    end = table(ws, 4, titles, rows, fmts, colours, wrap_cols=(2, 14))
    if not rows:
        put(ws, end, 1, "No stock rows - no stock file, or nothing to show.")
    set_widths(ws, [10, 40, 7, 7, 10, 10, 10, 9, 10, 10, 10, 26, 28, 50])
    ws.freeze_panes = "C5"
    ws.auto_filter.ref = "A4:N%d" % max(end - 1, 4)
    landscape(ws)


# ---- Open with SAM + SAM bills sheets (also used by the tracker file) ----------------

def write_open_sheet(ws, t, highlight=()):
    titles = ["Order date", "GDA / PO", "Gr Code", "Model", "Size", "Ordered", "Received", "Balance", "Status",
              "Bills set off"]
    fmts = [None, None, None, None, None, QTY_FMT, QTY_FMT, QTY_FMT, None, None]
    rows, colours = [], []
    for l in t["lines"]:
        st = line_status(l)
        rows.append([ddmmyyyy(l["date"]), line_label(l), l["code"], l["model"], l["size"], l["ordered"],
                     l["received"] or None, line_balance(l), st, "; ".join(l["bills"])])
        colours.append(ROW_FILL["GREEN"] if st == "DONE" else ROW_FILL["YELLOW"] if st == "PART" else None)
    end = table(ws, 1, titles, rows, fmts, colours, height=24)
    for n, l in enumerate(t["lines"]):
        if l["code"] in highlight:
            ws.cell(row=2 + n, column=3).font = BOLD
    o, rc, bal = tracker_totals(t)
    put(ws, end, 1, "TOTAL", font=BOLD, border=True)
    for c, v in ((6, o), (7, rc), (8, bal)):
        put(ws, end, c, v, fmt=QTY_FMT, font=BOLD, border=True, bg="D9E1F2")
    set_widths(ws, [11, 26, 9, 34, 6, 9, 9, 9, 8, 40])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:J%d" % max(end - 1, 1)


def write_bills_sheet(ws, t):
    titles = ["File", "Invoice", "Date", "Gr Code", "Size", "Qty", "Status", "Set against", "Extra", "Item text",
              "Rate Rs", "MOQ benefit Rs", "MOQ check"]
    fmts = [None, None, None, None, None, QTY_FMT, None, None, QTY_FMT, None, RATE_FMT, RATE_FMT, None]
    rows = [[b["file"], b["inv"], ddmmyyyy(b["date"]), b["code"], b["size"], b["qty"], b["status"], b["against"],
             b["extra"] or None, b["desc"], b.get("rate"), b.get("moq_rs"), b.get("moq_note", "")] for b in t["bills"]]
    colours = [ROW_FILL["RED"] if "EXTRA" in b["status"] else ROW_FILL["GREY"] if b["status"] == "SKIPPED"
               else ROW_FILL["GREEN"] for b in t["bills"]]
    end = table(ws, 1, titles, rows, fmts, colours, height=24)
    if not rows:
        put(ws, end, 1, "No SAM bills found in: %s" % BILLS_DIR)
    for n, b in enumerate(t["bills"]):
        note = b.get("moq_note", "")
        if note.startswith("SHORT"):
            ws.cell(row=2 + n, column=13).font = Font(bold=True, color="C00000")
    set_widths(ws, [26, 14, 11, 9, 6, 8, 11, 44, 8, 40, 9, 10, 50])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:M%d" % max(end - 1, 1)


def write_check_report(out_path, ctx):
    wb = openpyxl.Workbook()
    write_summary(wb.active, ctx)
    wb.active.title = "Summary"
    write_detail(wb.create_sheet("Detail"), ctx)
    if ctx["cfg"].stock_check:
        write_stock_sheet(wb.create_sheet("Stock cover"), ctx)
    write_open_sheet(wb.create_sheet("Open with SAM"), ctx["tracker"], highlight=set(ctx["res"]["models"]))
    write_bills_sheet(wb.create_sheet("SAM bills read"), ctx["tracker"])
    return save_workbook(wb, out_path)


def write_tracker_report(path, t):
    wb = openpyxl.Workbook()
    wb.active.title = "Open with SAM"
    write_open_sheet(wb.active, t)
    write_bills_sheet(wb.create_sheet("SAM bills read"), t)
    return save_workbook(wb, path)


# --------------------------------------------------------------------------------------
# Screen summary
# --------------------------------------------------------------------------------------

def colour_on():
    if not sys.stdout.isatty():
        return False
    if os.name == "nt":
        os.system("")  # switches on colour codes in the Windows console
    return True


ANSI = {"GREEN": "\033[92m", "ORANGE": "\033[93m", "RED": "\033[91m", "END": "\033[0m"}


def col(text, c):
    return (ANSI[c] + text + ANSI["END"]) if colour_on() and c in ANSI else text


def print_value(res, cfg):
    v = "Rs %s" % inr(res["total_value"])
    if res["over_cap"]:
        v += "   " + col("** ABOVE weekly cap Rs %s by Rs %s **" % (inr(cfg.weekly_cap), inr(res["total_value"] - cfg.weekly_cap)), "RED")
    elif not res["unpriced"]:
        v += "   (weekly cap Rs %s - OK)" % inr(cfg.weekly_cap)
    if res["unpriced"]:
        v += col("  INCOMPLETE: %d model(s) have no rate - %s" % (
            len(res["unpriced"]), "real value is even higher" if res["over_cap"] else "cap cannot be confirmed"), "ORANGE")
    print("  ORDER VALUE  : " + v)
    print("  MOQ BENEFIT  : earned Rs %s   |   LOST Rs %s" % (inr(res["earned"]), col(inr(res["lost"]), "RED") if res["lost"] else "0"))


def print_moq(res, cfg):
    print("  %-8s %-36s %6s %6s  %s" % ("", "MOQ GROUP" if cfg.mode == "group" else "MODEL (GROUP)", "QTY", "SLAB", "RESULT"))
    for u in res["units"]:
        name = u["name"] if cfg.mode == "group" else "%s (%s)" % (u["name"], u["group"])
        if not u["ordered"]:
            result = "not in this order"
        elif u["slab"] == 2:
            result = "OK"
        elif u["slab"] == 1:
            result = "add %s pcs for Slab-2  (+Rs %s)" % (fmt_qty(u["short2"]), inr(u["lost"]))
        else:
            result = "add %s for Slab-1 / %s for Slab-2  (+Rs %s)" % (fmt_qty(u["short1"]), fmt_qty(u["short2"]), inr(u["lost"]))
        tag = "[%-6s]" % (u["colour"] if u["ordered"] else "-")
        print("  %s %-36s %6s %6s  %s" % (col(tag, u["colour"]) if u["ordered"] else tag, name[:36], fmt_qty(u["qty"]),
                                         str(u["slab"]) if u["slab"] else "-", result))


def print_pending(res):
    hits = sorted(res["pending_hits"], key=lambda m: -m["pending"]["qty"])
    if not hits:
        return
    print("  STILL OPEN WITH SAM (%d models, biggest first):" % len(hits))
    for m in hits[:8]:
        p = m["pending"]
        oldest = min((d for (_, d) in p["gdas"] if d), default=None)
        print("     %-7s ordering %5s, open %5s%s   %s" % (m["code"], fmt_qty(m["qty"]), fmt_qty(p["qty"]),
              (" since " + ddmmyyyy(oldest)) if oldest else "", m["model"][:30]))
    if len(hits) > 8:
        print("     ... %d more in the report" % (len(hits) - 8))


def print_stock(cover, stock_used, on=True):
    if not on:
        return
    if not stock_used:
        print("  STOCK COVER  : " + col("not checked - no stock file", "ORANGE"))
        return
    n = {s: [r for r in cover if r["status"] == s] for s in ("SHORT", "OVER", "NO SALE")}
    print("  STOCK COVER  : %s short, %s over, %s no sale  (details: 'Stock cover' sheet)"
          % (col(str(len(n["SHORT"])), "RED") if n["SHORT"] else "0", len(n["OVER"]), len(n["NO SALE"])))
    for label, rows in (("ADD", n["SHORT"]), ("CUT", n["OVER"])):
        if rows:
            print("     %s: %s" % (label, short_list(rows, 8)))


def print_warnings(warns):
    if not warns:
        return
    print("-" * 78)
    print("  WARNINGS (%d):" % len(warns))
    for w in warns[:12]:
        print("   ! " + w)
    if len(warns) > 12:
        print("   ... %d more in the report" % (len(warns) - 12))


def print_summary(ctx):
    cfg, res = ctx["cfg"], ctx["res"]
    line = "-" * 78
    print("=" * 78)
    print("  SAM ORDER CHECK" + cfg.company.rjust(60))
    print("=" * 78)
    print("  Order      : %s" % ctx["order_name"])
    print("  Read       : %s models, %s pcs from %d sheet(s)" % (len(res["models"]), inr(res["total_qty"]), len(ctx["sheet_info"])))
    print("  Pending    : %s" % ctx["tracker"]["pending_desc"])
    print("  Stock      : %s" % ctx["stock_desc"])
    print(line)
    print_value(res, cfg)
    print_stock(ctx["cover"], ctx["stock_used"], ctx["cfg"].stock_check)
    print(line)
    print_moq(res, cfg)
    print(line)
    if res["not_in_scheme"]:
        print("  NOT IN MOQ SCHEME: %d models, %s pcs (no MOQ benefit) - list in the report"
              % (len(res["not_in_scheme"]), fmt_qty(sum(m["qty"] for m in res["not_in_scheme"]))))
    print_pending(res)
    print_warnings(ctx["warnings"])
    print(line)
    print("  Report saved : %s" % ctx["out_path"])
    print("  Order file was only read - not changed (%s)." % ctx["untouched"])
    print("=" * 78)


def print_tracker(t, path):
    o, rc, bal = tracker_totals(t)
    print("=" * 78)
    print("  SAM ORDER TRACKER - open with SAM = pending list + orders sent - SAM bills")
    print("=" * 78)
    print("  Pending list : %s" % t["pending_desc"])
    print("  Orders sent  : %s" % (", ".join("%s (%s pcs)" % (n, fmt_qty(q)) for n, _, q in t["orders_used"]) or "none added"))
    applied = [b for b in t["bills"] if b["status"] != "SKIPPED"]
    print("  SAM bills    : %d line(s) read, %d set off, %d skipped (already in pending list)"
          % (len(t["bills"]), len(applied), len(t["bills"]) - len(applied)))
    got = sum(b.get("moq_rs") or 0 for b in t["bills"])
    if got:
        short = t.get("moq_short", 0)
        print("  MOQ benefit on bills: Rs %s received%s" % (inr(got), col("  |  SHORT-PAID Rs %s - claim from SAM" % inr(short), "RED")
                                                          if short else "  |  all correct as per moq.csv"))
    print("-" * 78)
    print("  Ordered %s   |   Received %s   |   STILL TO COME %s pcs" % (inr(o), inr(rc), inr(bal)))
    extra = [b for b in t["bills"] if b["extra"] > 0]
    if extra:
        print("  " + col("EXTRA (billed but not on any open order):", "RED"))
        for b in extra:
            print("     %s %s %s: %s pcs  (%s %s)" % (b["code"], b["size"], b["desc"][:30], fmt_qty(b["extra"]),
                                                   b["inv"] or b["file"], ddmmyyyy(b["date"])))
    for n in t["notes"]:
        print("  note: " + n)
    print_warnings(t["warns"])
    print("-" * 78)
    print("  Tracker saved: %s" % path)
    print("=" * 78)


# --------------------------------------------------------------------------------------
# File picking and main
# --------------------------------------------------------------------------------------

def gui_pick(title, initial_dir):
    """File-open dialog. Returns path, '' if cancelled, or None if no dialog is possible."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
    except Exception:
        return None
    try:
        root.withdraw()
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass
        path = filedialog.askopenfilename(parent=root, title=title, initialdir=initial_dir or TOOL_DIR,
                                          filetypes=[("Excel files", "*.xls *.xlsx *.xlsm"), ("All files", "*.*")])
        return path or ""
    finally:
        root.destroy()


def ask_order_file():
    initial = ""
    try:
        with open(LAST_FOLDER_FILE, encoding="utf-8") as fh:
            initial = fh.read().strip()
    except Exception:
        pass
    path = gui_pick("Select the SAM ORDER file to check", initial if os.path.isdir(initial) else "")
    if path is None:
        try:
            path = input("Type or paste the full path of the SAM order file: ").strip().strip('"')
        except EOFError:
            path = ""
    if not path:
        raise CheckError("No order file was selected.")
    try:
        with open(LAST_FOLDER_FILE, "w", encoding="utf-8") as fh:
            fh.write(os.path.dirname(os.path.abspath(path)))
    except Exception:
        pass
    return path


def open_file(path):
    try:
        if os.name == "nt":
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        print("  (Could not open the file automatically: %s)" % e)


def check_order_path(order_path):
    order_path = os.path.abspath(order_path.strip('"'))
    base = os.path.basename(order_path)
    if not os.path.exists(order_path):
        raise CheckError("File not found:\n  %s" % order_path)
    if base.startswith("~$"):
        raise CheckError("'%s' is Excel's temporary lock file, not the order. Pick the real order file." % base)
    if "_CHECK" in os.path.splitext(base)[0].upper():
        raise CheckError("'%s' is a CHECK report made by this tool. Please pick the original SAM order file." % base)
    return order_path


def load_stock(cfg, warns):
    if not cfg.stock_check:
        return {}, None, "switched off (settings.ini: stock_check = no)", False
    spath, how = find_stock_file(cfg)
    if not spath:
        warns.append("No stock file found - put Tally's Stock Summary Excel in the folder:  %s" % STOCK_DIR)
        return {}, None, "NOT CHECKED - no stock file", False
    try:
        stock, months, period = read_stock_file(spath, cfg)
    except Exception as e:
        warns.append("Could not read stock file '%s' (%s) - stock cover not checked." % (os.path.basename(spath), str(e).split("\n")[0]))
        return {}, None, "NOT CHECKED - stock file unreadable", False
    if not stock:
        warns.append("Stock file '%s': no GR items with Outwards / Closing Balance columns found." % os.path.basename(spath))
    if not months:
        months = 1.0
        warns.append("Stock file has no period (e.g. '1-Oct-26 to 31-Oct-26') - sales taken as 1 month. "
                     "Set stock_months in settings.ini if wrong.")
    age = (dt.date.today() - dt.date.fromtimestamp(os.path.getmtime(spath))).days
    if age > 7:
        warns.append("Stock file '%s' is %d days old - export a fresh Stock Summary from Tally." % (os.path.basename(spath), age))
    desc = "%s  (%d items, period %s = %.2f months, %s)" % (os.path.basename(spath), len(stock), period or "?", months, how)
    return stock, months, desc, bool(stock)


def order_warnings(cfg, order, groups, rates_problems, res):
    warns = []
    if cfg.scheme_end and dt.date.today() > cfg.scheme_end:
        warns.append("The MOQ scheme%s ended on %s. Check moq.csv is the CURRENT scheme, then update scheme_name / "
                     "scheme_end_date in settings.ini." % ((" '%s'" % cfg.scheme_name) if cfg.scheme_name else "",
                                                           ddmmyyyy(cfg.scheme_end)))
    if cfg.check_po_scheme:
        warns.extend(check_scheme_vs_po(order["scheme_rows"], groups, cfg.scheme_ignore))
    warns.extend(order["warns"])
    warns.extend(rates_problems)
    for m in res["models"].values():
        warns.extend("%s: %s" % (m["code"], n) for n in m["notes"])
    return warns


def run_check(args, cfg):
    order_path = check_order_path(args.order or ask_order_file())
    base = os.path.basename(order_path)
    print("\n  Checking %s ..." % base)
    groups = load_moq()
    rates, rate_labels, rate_problems = load_rates()
    hash_before = file_sha256(order_path)
    order = read_order(order_path, cfg.order_sheets)
    if not order["lines"]:
        raise CheckError("No quantities found in '%s'.\n  Every Model Code line is empty - is this the filled-in order?" % base)
    tracker = build_tracker(cfg, order_path, hash_before)
    res = analyse(order["lines"], groups, rates, rate_labels, tracker_open_by_code(tracker), cfg)
    warns = tracker["warns"] + order_warnings(cfg, order, groups, rate_problems, res)
    stock, months, stock_desc, stock_used = load_stock(cfg, warns)
    cover = stock_cover(res, order, stock, months, tracker_open_by_key(tracker), cfg) if stock_used else []
    ctx = {"cfg": cfg, "res": res, "groups": groups, "order_name": base, "sheet_info": order["sheet_info"],
           "tracker": tracker, "stock_desc": stock_desc, "stock_used": stock_used, "cover": cover,
           "warnings": warns, "now": dt.datetime.now(), "actions": build_actions(res, cover, cfg),
           "untouched": "verified" if file_sha256(order_path) == hash_before else "CHANGED?!"}
    stem = os.path.splitext(base)[0]
    ctx["out_path"] = write_check_report(os.path.join(os.path.dirname(order_path), stem + "_CHECK.xlsx"), ctx)
    if file_sha256(order_path) != hash_before:
        ctx["untouched"] = "WARNING: order file changed while checking"
    print_summary(ctx)
    if cfg.open_report and not args.no_open:
        open_file(ctx["out_path"])
    return 0


def run_tracker(args, cfg):
    print("\n  Updating SAM order tracker ...")
    t = build_tracker(cfg)
    path = write_tracker_report(TRACKER_FILE, t)
    print_tracker(t, path)
    if cfg.open_report and not args.no_open:
        open_file(path)
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Check a SAM purchase order / update the SAM order tracker.")
    ap.add_argument("order", nargs="?", help="SAM order Excel file (.xls / .xlsx). If left out, a file picker opens.")
    ap.add_argument("--tracker", action="store_true", help="update SAM_Order_Tracker.xlsx instead of checking an order")
    ap.add_argument("--mode", choices=["group", "model"], help="override count_mode from settings.ini")
    ap.add_argument("--no-open", action="store_true", help="do not open the report afterwards")
    args = ap.parse_args(argv)
    try:
        cfg = Settings()
        if args.mode:
            cfg.mode = args.mode
        return run_tracker(args, cfg) if args.tracker else run_check(args, cfg)
    except CheckError as e:
        print("\n  PROBLEM: %s\n" % e)
        return 1
    except Exception:
        print("\n  Something unexpected went wrong. Please send this message to whoever set up the tool:\n")
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
