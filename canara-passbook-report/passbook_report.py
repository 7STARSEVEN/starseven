#!/usr/bin/env python3
"""
Canara Bank e-passbook  ->  monthly cash report (Excel + on-screen summary).

Reads every PDF in the input folder (default: ./pdfs), rebuilds the transaction
list, works out deposit vs withdrawal from the change in balance, checks that the
balances add up, groups collections and payments, raises flags and writes one
Excel file to ./output.

Usage:
    python passbook_report.py                      (reads ./pdfs)
    python passbook_report.py C:\\path\\to\\pdfs
    python passbook_report.py --password XXXX      (if the PDFs are password protected)
    python passbook_report.py --dump               (also print every parsed row, to check parsing)

Needs: pdfplumber, openpyxl  (no internet needed once installed).
"""

import argparse
import csv
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime

# ----------------------------------------------------------------------------
# Settings you may want to change
# ----------------------------------------------------------------------------
LARGE_PAYMENT_LIMIT = 200000.00   # flag payments above this to anyone except SAM (Rs)
DUPLICATE_WINDOW_DAYS = 7          # same amount to two parties within this many days
DUPLICATE_MIN_AMOUNT = 1000.00     # ignore small identical amounts (Rs)
NO_REPEAT_FLAG = {"SAM", "Owner transfer", "Cash withdrawal"}  # regular repeat payees: no double-payment flag
COMPARE_MONTHS = 3                 # months in the comparison table
OWN_NAME_WORDS = {"PISTON", "AGENCY", "AGENCIES"}  # our own name, ignored when guessing the other party

# Fixed group names used by the report (payees.csv may add more groups)
SAM = "SAM"
OWNER = "Owner transfer"
LOAN = "Loan/Drawdown"
CHARGES = "Bank charges"
CASH_WDL = "Cash withdrawal"
CHQ_RETURN = "Cheque return"
OTHERS = "Others"
SUPPLIER_ORDER = [SAM, "MACO", "Rane", "Mitsuboshi", "Jai Auto", "H M Industries"]
SPECIAL_ORDER = [OWNER, LOAN, CHARGES, CASH_WDL, CHQ_RETURN]

MODES = ["UPI", "NEFT/RTGS", "IMPS", "Cheque clearing", "Cash deposit", "Canara transfer", "Other"]
WEEKS = ["Week 1 (1-7)", "Week 2 (8-14)", "Week 3 (15-21)", "Week 4 (22-28)", "Week 5 (29-31)"]

# ----------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------
MONTHS = {m: i for i, m in enumerate(
    ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"], 1)}

RE_DATE_NUM = re.compile(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$")
RE_DATE_MON = re.compile(r"^(\d{1,2})[-/ ]?([A-Za-z]{3})[A-Za-z]*[-/ ,]*(\d{2}|\d{4})$")
RE_TIME = re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$")
RE_MONEY = re.compile(r"^(?:Rs\.?|INR)?(-)?((?:\d{1,3}(?:,\d{2,3})+|\d+)\.\d{2})(Cr|Dr|CR|DR|cr|dr)?\.?$")
RE_CRDR = re.compile(r"^\(?(Cr|Dr|CR|DR|cr|dr)\)?\.?$")
RE_FOOTER = re.compile(
    r"PAGE\s*\d+\s*(OF|/)\s*\d+|COMPUTER GENERATED|DOES NOT REQUIRE|SIGNATURE|"
    r"UNLESS THE CONSTITUENT|TOLL FREE|WWW\.|CUSTOMER CARE|STATEMENT GENERATED|"
    r"END OF STATEMENT|\*\*\*|^\s*PAGE\s*\d+\s*$", re.I)
RE_OPENING = re.compile(r"OPENING\s*BAL|BALANCE\s*B\s*/?\s*F|BROUGHT\s*FORWARD|^\s*B\s*/\s*F\b", re.I)
RE_CLOSING = re.compile(r"CLOSING\s*BAL|BALANCE\s*C\s*/?\s*F|CARRIED\s*FORWARD|^\s*C\s*/\s*F\b", re.I)
RE_RETURN = re.compile(r"\b(RTN|RETURN|RETURNED|RET|DISHONOU?R(ED)?|BOUNCED?|UNPAID|INWARD RET|OUTWARD RET)\b")
RE_EPAY = re.compile(r"\b(UPI|IMPS|NEFT|RTGS)\b")
RE_CHQ_CTX = re.compile(r"\b(CHQ|CHEQUE|CLG|CTS|GEFU|INWARD|OUTWARD|OW|IW|ECS|NACH|MICR)\b")
RE_EPAY_RETURN = re.compile(r"\b(NEFT|RTGS|IMPS|UPI)\s*(RETURN|RTN|RET|REVERSAL|REV|REFUND)\b")
# Canara narration formats: pull out the other party's name
PARTY_PATTERNS = [re.compile(p) for p in [
    r"NEFT DR\s+CNRBH\d+\s+(.+?)\s+[A-Z]{4}0[A-Z0-9]{6}\b",
    r"NEFT DR-\s*CNRBH\d+-\s*[A-Z]{4}0[A-Z0-9]{6}-\s*([^-]+)",
    r"NEFT CR-\s*\S+?-\s*[A-Z]{4}0[A-Z0-9]{6}-\s*([^-]+)",
    r"NEFT RETURN-\S+-\s*([^-]+)",
    r"RTGS (?:DR|CR)-\s*\S+?-\s*(?:[A-Z]{4}0[A-Z0-9]{6}-\s*)?([^-]+)",
    r"IMPS-(?:DR|CR)/([^/]+)/",
    r"IMPS\s*-?\s*CR/([^/]+)/",
    r"UPI/(?:DR|CR)/\d+/([^/]+)/",
    r"MB/\d+/PISTON AGENCY/\d+/([^/]+)/",
    r"CASH DEPOSIT(?: BY)?\s+(.+)",
    r"FUNDS TRANSFER (?:DEBIT|CREDIT)\s*[\d ]*-\s*(.+)",
    r"BY CLG:\S+ CLG SEC-?\s*[^,]+,\s*(.+)",
]]
RE_IFSC = re.compile(r"^[A-Z]{4}0[A-Z0-9]{6}$")


def parse_date(text):
    t = text.strip().rstrip(",")
    m = RE_DATE_NUM.match(t)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = RE_DATE_MON.match(t)
        if not m or m.group(2).upper()[:3] not in MONTHS:
            return None
        d, mo, y = int(m.group(1)), MONTHS[m.group(2).upper()[:3]], int(m.group(3))
    if y < 100:
        y += 2000
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def parse_money(text):
    """'1,23,456.78' -> (12345678 paise, None|'CR'|'DR')"""
    m = RE_MONEY.match(text.strip().replace("\u20b9", ""))
    if not m:
        return None
    paise = int(round(float(m.group(2).replace(",", "")) * 100))
    if m.group(1):
        paise = -paise
    return paise, (m.group(3) or "").upper() or None


def norm(text):
    return " " + re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9]+", " ", text.upper())).strip() + " "


def inr(paise, blank_zero=False):
    """Indian-style number: 12,34,567.89"""
    if paise is None:
        return ""
    if blank_zero and paise == 0:
        return "-"
    neg = paise < 0
    s = f"{abs(paise) / 100:.2f}"
    whole, frac = s.split(".")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        head = ",".join(re.findall(r"\d{1,2}", head[::-1]))[::-1]
        whole = head + "," + tail
    return ("-" if neg else "") + whole + "." + frac


# ----------------------------------------------------------------------------
# PDF reading
# ----------------------------------------------------------------------------
@dataclass
class Row:
    date: date
    text: str
    money: list            # [(paise, crdr, x_center)]
    page: int
    file: str


@dataclass
class Txn:
    date: date
    particulars: str
    amount: int            # paise, always positive
    kind: str              # 'CR' deposit or 'DR' withdrawal
    balance: int
    file: str
    page: int
    check: str = "OK"
    mode: str = ""
    group: str = ""
    category: str = ""     # collection mode / non-collection type / payment group
    is_collection: bool = False
    is_return: bool = False
    ret_type: str = ""
    party: str = ""
    book_date: date = None   # date used for month/week; differs only for back-dated rows

    @property
    def bdate(self):
        return self.book_date or self.date

    @property
    def month(self):
        return (self.bdate.year, self.bdate.month)


@dataclass
class FileResult:
    file: str
    pages: int = 0
    txns: list = field(default_factory=list)
    opening: int = None
    opening_source: str = ""
    closing_printed: int = None
    notes: list = field(default_factory=list)


def cluster_lines(words, tol=3.0):
    words = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    lines = []
    for w in words:
        if lines and abs(w["top"] - lines[-1]["top"]) <= tol:
            lines[-1]["words"].append(w)
            lines[-1]["bottom"] = max(lines[-1]["bottom"], w["bottom"])
        else:
            lines.append({"top": w["top"], "bottom": w["bottom"], "words": [w]})
    for ln in lines:
        ln["words"].sort(key=lambda w: w["x0"])
        ln["words"] = merge_date_words(ln["words"])
        ln["text"] = " ".join(w["text"] for w in ln["words"])
        ln["mid"] = (ln["top"] + ln["bottom"]) / 2
    return lines


def merge_date_words(ws):
    """Join '01' 'Oct' '2025' (three words) into one date word."""
    out, i = [], 0
    while i < len(ws):
        if i + 2 < len(ws) and re.fullmatch(r"\d{1,2}", ws[i]["text"]) and \
                ws[i + 1]["text"][:3].upper() in MONTHS and re.fullmatch(r"\d{2,4},?", ws[i + 2]["text"]) and \
                ws[i + 2]["x0"] - ws[i]["x1"] < 40:
            w = dict(ws[i])
            w["text"] = " ".join(x["text"] for x in ws[i:i + 3])
            w["x1"] = ws[i + 2]["x1"]
            out.append(w)
            i += 3
        else:
            out.append(ws[i])
            i += 1
    return out


def find_header(lines):
    """Return (header_bottom, columns) or (None, None)."""
    for idx, ln in enumerate(lines):
        up = ln["text"].upper()
        if "BALANCE" in up and "DATE" in up and re.search(r"PARTICULARS|DESCRIPTION|NARRATION|DETAILS", up):
            cols = {}
            for w in ln["words"]:
                t = w["text"].upper()
                c = (w["x0"] + w["x1"]) / 2
                if t.startswith("DATE") and "date" not in cols:
                    cols["date"] = w
                    cols["date_x1"] = w["x1"]
                elif re.match(r"PARTICULARS|DESCRIPTION|NARRATION|DETAILS", t):
                    cols["part_x0"] = w["x0"]
                elif t.startswith("DEPOSIT") or t.startswith("CREDIT"):
                    cols["dep"] = c
                elif t.startswith("WITHDRAW") or t.startswith("DEBIT"):
                    cols["wd"] = c
                elif t.startswith("BALANCE"):
                    cols["bal"] = c
            bottom = ln["bottom"]
            # a second header line like "(Rs.)  (Rs.)" directly below
            if idx + 1 < len(lines) and lines[idx + 1]["top"] - bottom < 12 and \
                    not any(parse_date(w["text"]) for w in lines[idx + 1]["words"]) and \
                    re.fullmatch(r"[\s()A-Za-z./]*", lines[idx + 1]["text"]) and len(lines[idx + 1]["text"]) < 40:
                bottom = lines[idx + 1]["bottom"]
            return bottom, cols
    return None, None


def horizontal_rules(page):
    ys = []
    pw = float(page.width)
    for ln in page.lines:
        if abs(ln["top"] - ln["bottom"]) < 1.5 and abs(ln["x1"] - ln["x0"]) > pw * 0.4:
            ys.append(ln["top"])
    for r in page.rects:
        if abs(r["x1"] - r["x0"]) > pw * 0.4:
            if r["bottom"] - r["top"] < 2:
                ys.append(r["top"])
            else:
                ys.extend([r["top"], r["bottom"]])
    ys.sort()
    out = []
    for y in ys:
        if not out or y - out[-1] > 1.5:
            out.append(y)
    return out


def read_pdf(path, password=None, dump=False):
    import pdfplumber
    res = FileResult(file=os.path.basename(path))
    pages = []          # per page: list of classified line dicts
    cols = None
    with pdfplumber.open(path, password=password or None) as pdf:
        res.pages = len(pdf.pages)
        for pno, page in enumerate(pdf.pages, 1):
            words = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)
            lines = cluster_lines(words)
            hb, hcols = find_header(lines)
            if hcols:
                cols = hcols
            pw = float(page.width)
            part_x0 = (cols or {}).get("part_x0")
            money_x = pw * 0.45
            if cols and any(k in cols for k in ("dep", "wd", "bal")):
                money_x = min(cols.get(k, pw) for k in ("dep", "wd", "bal")) - 60
            if part_x0:
                money_x = max(money_x, part_x0 + 40)
            if res.opening is None:   # "Opening Balance : 1,23,456.78" printed above the table
                for ln in lines:
                    if hb is not None and ln["top"] <= hb and RE_OPENING.search(ln["text"]):
                        ms = [parse_money(w["text"]) for w in ln["words"]]
                        ms = [m for m in ms if m]
                        if ms:
                            dr = ms[-1][1] == "DR" or re.search(r"\bDr\b", ln["text"])
                            res.opening = -abs(ms[-1][0]) if dr else ms[-1][0]
                            res.opening_source = "printed"
            content = [ln for ln in lines if hb is None or ln["top"] > hb + 0.5]
            # dates must start left of this x (particulars may start left of their heading)
            date_zone = (part_x0 - 3) if part_x0 else pw * 0.22
            if cols and cols.get("date_x1"):
                date_zone = min(date_zone, cols["date_x1"] + 30)
            plines = []
            for ln in content:
                ws = ln["words"]
                first = ws[0]
                d = parse_date(first["text"])
                is_anchor = d is not None and first["x0"] < date_zone
                money, text = [], []
                for i, w in enumerate(ws):
                    if is_anchor and i == 0:
                        continue
                    pm = parse_money(w["text"])
                    if pm and w["x0"] >= money_x:
                        money.append([pm[0], pm[1], (w["x0"] + w["x1"]) / 2])
                        continue
                    if RE_CRDR.match(w["text"]) and money and w["x0"] >= money_x:
                        money[-1][1] = RE_CRDR.match(w["text"]).group(1).upper()
                        continue
                    if is_anchor and (RE_TIME.match(w["text"]) or
                                      (parse_date(w["text"]) and w["x0"] < date_zone + 25)):
                        continue   # value date / time column
                    text.append(w["text"])
                kind = "text"
                if is_anchor:
                    kind = "anchor"
                elif RE_OPENING.search(ln["text"]) and money:
                    kind = "opening"
                elif RE_CLOSING.search(ln["text"]) and money:
                    kind = "closing"
                elif RE_FOOTER.search(ln["text"]):
                    kind = "stop"
                elif re.match(r"CHQ\b", ln["text"], re.I):
                    kind = "chq"      # Canara e-passbook: every row ends with a "Chq: <no>" line
                elif first["x0"] < (date_zone if part_x0 else -1) and not money:
                    kind = "stop"     # text starting in the date column = not part of the table
                plines.append({"kind": kind, "date": d, "text": " ".join(text), "money": money,
                               "mid": ln["mid"], "top": ln["top"], "page": pno})
            pages.append({"lines": plines, "rules": horizontal_rules(page), "cols": cols})

    rows = assemble_rows(pages, res)
    if not rows:
        res.notes.append("No transaction rows found - is this a Canara e-passbook PDF?")
    build_txns(rows, res, cols)
    if dump:
        print(f"\n--- {res.file}: parsed rows ---")
        for t in res.txns:
            print(f"{t.date:%d-%m-%Y} {t.kind} {inr(t.amount):>15} bal {inr(t.balance):>15}  {t.check:<6} {t.particulars[:70]}")
    return res


def assemble_rows(pages, res):
    """Attach particulars lines to their date row (handles top-aligned, centred and ruled tables)."""
    n_chq = sum(l["kind"] == "chq" for pg in pages for l in pg["lines"])
    n_anchor = sum(l["kind"] == "anchor" for pg in pages for l in pg["lines"])
    if n_anchor and n_chq >= 0.5 * n_anchor:
        return assemble_chq_rows(pages, res)
    # Decide layout: centred cells put lines above the date.
    centred = False
    for pg in pages:
        ls = pg["lines"]
        anchors = [i for i, l in enumerate(ls) if l["kind"] == "anchor"]
        if not anchors:
            continue
        if any(l["kind"] == "text" and l["text"] for l in ls[:anchors[0]]) and pg is pages[0]:
            centred = True
        for i in anchors:
            if not ls[i]["text"]:
                near = [l for l in ls[max(0, i - 1):i + 2] if l["kind"] == "text" and l["text"]]
                if len(near) >= 2:
                    centred = True
    rows = []
    for pg in pages:
        ls = pg["lines"]
        for l in ls:
            if l["kind"] == "opening" and res.opening is None:
                res.opening = l["money"][-1][0] if l["money"][-1][1] != "DR" else -abs(l["money"][-1][0])
                res.opening_source = "printed"
            if l["kind"] == "closing":
                m = l["money"][-1]
                res.closing_printed = -abs(m[0]) if m[1] == "DR" else m[0]
        anchors = [i for i, l in enumerate(ls) if l["kind"] == "anchor"]
        page_rows = []
        for i in anchors:
            l = ls[i]
            page_rows.append({"row": Row(l["date"], "", list(l["money"]), l["page"], res.file),
                              "parts": [(l["mid"], l["text"])] if l["text"] else [], "mid": l["mid"]})
        prev_row = rows[-1] if rows else None
        rules = pg["rules"]
        ruled = False
        if len(anchors) >= 2 and rules:
            gaps = sum(1 for a, b in zip(page_rows, page_rows[1:])
                       if any(a["mid"] < y < b["mid"] for y in rules))
            ruled = gaps >= 0.8 * (len(page_rows) - 1)

        def attach(target, l):
            if target is None:
                return
            if l["text"]:
                target["parts"].append((l["mid"], l["text"]))
            if l["money"] and not target["row"].money:
                target["row"].money = list(l["money"])

        if ruled:
            def band(y):
                above = [r for r in rules if r <= y]
                below = [r for r in rules if r > y]
                return (above[-1] if above else -1e9, below[0] if below else 1e9)
            bands = [band(pr["mid"]) for pr in page_rows]
            stopped = False
            for idx, l in enumerate(ls):
                if l["kind"] == "stop":
                    stopped = True
                if l["kind"] != "text":
                    continue
                b = band(l["mid"])
                tgt = next((pr for pr, bb in zip(page_rows, bands) if bb == b), None)
                if tgt is None and page_rows and l["mid"] < page_rows[0]["mid"] and not stopped:
                    tgt = prev_row          # row continued from previous page
                attach(tgt, l)
        elif not centred:
            current = prev_row
            for l in ls:
                if l["kind"] == "anchor":
                    current = next(pr for pr in page_rows if pr["mid"] == l["mid"])
                elif l["kind"] in ("stop", "opening", "closing"):
                    current = None
                elif l["kind"] == "text":
                    attach(current, l)
        else:
            # centred: split text lines between consecutive dates symmetrically
            segs, cur = [[]], 0
            for l in ls:
                if l["kind"] == "anchor":
                    segs.append([])
                elif l["kind"] == "text":
                    segs[-1].append(l)
                elif l["kind"] in ("stop", "closing"):
                    segs[-1].append(None)    # marker: nothing after this belongs to a row
            above = 0
            for k, seg in enumerate(segs):
                if None in seg:
                    seg = seg[:seg.index(None)]
                if k == 0:
                    if page_rows:
                        for l in seg:
                            attach(page_rows[0], l)
                        above = len(seg)
                    continue
                me = page_rows[k - 1]
                nxt = page_rows[k] if k < len(page_rows) else None
                if nxt is None:
                    take = above
                else:
                    take = split_point(seg, above)
                for l in seg[:take]:
                    attach(me, l)
                for l in seg[take:]:
                    attach(nxt, l)
                above = len(seg) - take
        for pr in page_rows:
            pr["parts"].sort(key=lambda p: p[0])
            pr["row"].text = " ".join(t for _, t in pr["parts"]).strip()
            rows.append(pr)
    return [pr["row"] for pr in rows]


def assemble_chq_rows(pages, res):
    """Canara e-passbook layout: a row is every line from one 'Chq:' line to the next,
    with the date and amounts on one of the middle lines. Works across page breaks."""
    rows, pending, cur = [], [], None
    for pg in pages:
        for l in pg["lines"]:
            k = l["kind"]
            if k == "opening" and res.opening is None:
                m = l["money"][-1]
                res.opening = -abs(m[0]) if m[1] == "DR" else m[0]
                res.opening_source = "printed"
            elif k == "closing":
                m = l["money"][-1]
                res.closing_printed = -abs(m[0]) if m[1] == "DR" else m[0]
            elif k == "anchor":
                if cur is not None:          # two dates without a Chq line between them
                    rows.append(cur)
                cur = {"row": Row(l["date"], "", list(l["money"]), l["page"], res.file),
                       "parts": [p["text"] for p in pending if p["text"]] + ([l["text"]] if l["text"] else [])}
                pending = []
            elif k == "text":
                if cur is not None:
                    cur["parts"].append(l["text"])
                    if l["money"] and not cur["row"].money:
                        cur["row"].money = list(l["money"])
                else:
                    pending.append(l)
            elif k == "chq":
                if cur is not None:
                    no = re.sub(r"(?i)^chq\s*:?\s*", "", l["text"]).strip()
                    if no and no.strip("0") and len(no) <= 6:
                        cur["parts"].append(f"Chq {no}")
                    rows.append(cur)
                elif pending and rows:       # tail of a row split by a page break
                    rows[-1]["parts"] += [p["text"] for p in pending if p["text"]]
                cur, pending = None, []
            elif k in ("stop",):
                pass                          # page footer/header: ignore, row may continue
    if cur is not None:
        rows.append(cur)
    out = []
    for r in rows:
        r["row"].text = " ".join(p for p in r["parts"] if p).strip()
        out.append(r["row"])
    return out


def split_point(seg, above_prev):
    """For centred rows: how many of the lines between two dates belong to the upper row."""
    n = len(seg)
    if n == 0:
        return 0
    gaps = [seg[i + 1]["top"] - seg[i]["top"] for i in range(n - 1)]
    if len(gaps) >= 2:
        g = sorted(gaps)
        med = g[len(g) // 2]
        big = max(gaps)
        if big > med * 1.4:
            return gaps.index(big) + 1
    return min(n, above_prev)


def guess_kind(text):
    u = norm(text)
    if re.search(r" (UPI CR|CR|BY|CREDIT|DEP|DEPOSIT|RECEIVED) ", u) or "/CR/" in text.upper():
        return "CR"
    if re.search(r" (UPI DR|DR|TO|DEBIT|PAID|WDL) ", u) or "/DR/" in text.upper():
        return "DR"
    return None


def build_txns(rows, res, cols):
    prev = res.opening
    for r in rows:
        if not r.money:
            res.notes.append(f"Row {r.date:%d-%m-%Y} '{r.text[:40]}' has no amount/balance - skipped")
            continue
        # balance = token in balance column, else right-most
        if cols and "bal" in cols:
            bal_tok = min(r.money, key=lambda m: abs(m[2] - cols["bal"]))
        else:
            bal_tok = max(r.money, key=lambda m: m[2])
        bal = -abs(bal_tok[0]) if bal_tok[1] == "DR" else bal_tok[0]
        others = [m for m in r.money if m is not bal_tok]
        if RE_OPENING.search(r.text) and not others:
            if prev is None:
                res.opening, res.opening_source = bal, "printed"
            prev = bal
            continue
        if RE_CLOSING.search(r.text) and not others:
            res.closing_printed = bal
            continue
        amount, hint = None, None
        if others:
            a = others[0]
            amount = abs(a[0])
            if a[1] in ("CR", "DR"):
                hint = a[1]
            elif cols and "dep" in cols and "wd" in cols:
                hint = "CR" if abs(a[2] - cols["dep"]) < abs(a[2] - cols["wd"]) else "DR"
        check = "OK"
        if prev is not None:
            delta = bal - prev
            if amount is None:
                amount = abs(delta)
                kind = "CR" if delta >= 0 else "DR"
            elif abs(delta) == amount and delta != 0:
                kind = "CR" if delta > 0 else "DR"
            else:
                kind = hint or guess_kind(r.text) or ("CR" if delta > 0 else "DR")
                check = f"MISMATCH (prev bal {inr(prev)} -> {inr(bal)}, amount {inr(amount)})"
        else:
            if amount is None:
                res.notes.append(f"First row {r.date:%d-%m-%Y} has no amount and no opening balance - skipped")
                prev = bal
                continue
            kind = hint or guess_kind(r.text) or "CR"
            res.opening = bal - amount if kind == "CR" else bal + amount
            res.opening_source = "worked out from first row"
            check = "OK" if hint else "OK (direction guessed)"
        res.txns.append(Txn(r.date, re.sub(r"\s+", " ", r.text), amount, kind, bal, r.file, r.page, check))
        prev = bal


# ----------------------------------------------------------------------------
# Classification
# ----------------------------------------------------------------------------
def load_payees(path):
    rules = []
    if not os.path.exists(path):
        print(f"WARNING: {path} not found - all payments will be 'Others'.")
        return rules
    try:
        with open(path, newline="", encoding="utf-8-sig") as f:
            lines = list(csv.reader(f))
    except UnicodeDecodeError:          # saved from Excel in the Windows code page
        with open(path, newline="", encoding="cp1252") as f:
            lines = list(csv.reader(f))
    for row in lines:
        if not row or not row[0].strip() or row[0].strip().startswith("#"):
            continue
        kw = row[0].strip()
        grp = row[1].strip() if len(row) > 1 else ""
        if kw.lower() == "keyword" or not grp:
            continue
        prefix = kw.endswith("*")
        core = norm(kw.rstrip("*")).strip()
        if not core:
            continue
        pat = r"(?<![A-Z0-9])" + re.escape(core) + ("" if prefix else r"(?![A-Z0-9])")
        rules.append((re.compile(pat), grp, kw))
    return rules


def match_group(text, rules):
    u = norm(text)
    for pat, grp, _ in rules:
        if pat.search(u):
            return grp
    return ""


def pay_mode(text):
    u = norm(text)
    if " UPI " in u:
        return "UPI"
    if re.search(r" (NEFT|RTGS) ", u):
        return "NEFT/RTGS"
    if " IMPS " in u:
        return "IMPS"
    if re.search(r" (CLG|CTS|CHQ|CHEQUE|CLEARING|MICR|INWARD|OUTWARD) ", u) or re.search(r" CHQ ?NO", u):
        return "Cheque clearing"
    if re.search(r" (CASH|CDM|BNA|BY CASH|CSH) ", u):
        return "Cash deposit"
    if re.search(r" (FUNDS TRANSFER|FUND TRANSFER|TRANSFER FROM|BY TRANSFER|IFT) ", u):
        return "Canara transfer"
    return "Other"


def counterparty(text):
    """Best guess at the other party's name from the narration."""
    up = text.upper()
    for pat in PARTY_PATTERNS:
        m = pat.search(up)
        if m:
            name = re.sub(r"\s+", " ", m.group(1)).strip(" -/.,")
            name = re.sub(r"\s*\b(NGP CEN AVEN|CHQ \d+)$", "", name)
            if re.search(r"[A-Z]{3}", name):
                return name[:40]
    parts = re.split(r"[/\-:|]+|\s{2,}", up)
    skip = {"UPI", "CR", "DR", "NEFT", "RTGS", "IMPS", "P2A", "P2P", "P2M", "MOB", "CHQ", "CLG", "CTS", "BY",
            "TO", "TRANSFER", "TRF", "PAYMENT", "PAY", "PAID", "FROM", "INB", "MB", "NA", "IFT", "CHEQUE", "MICR",
            "INWARD", "OUTWARD", "CLEARING", "DEPOSIT", "DEP", "NO", "REF", "TXN", "BANK"}
    for p in parts:
        p = p.strip(" .,")
        words = p.split()
        if not words or "@" in p or RE_IFSC.match(p.replace(" ", "")):
            continue
        if all(w in skip or w in OWN_NAME_WORDS or not re.search(r"[A-Z]", w) for w in words):
            continue
        letters = sum(c.isalpha() for c in p)
        if letters < 3 or (len(p) <= 4 and p.isalpha() and p.isupper() and len(parts) > 3):
            continue   # bank short codes like HDFC/SBIN
        if re.fullmatch(r"[A-Z0-9]{10,}", p.replace(" ", "")) and re.search(r"\d", p):
            continue   # reference numbers
        name = " ".join(w for w in words if w not in skip and not w.isdigit()).strip()[:40]
        if name:
            return name
    return ""


def classify(t, rules):
    t.mode = pay_mode(t.particulars)
    t.group = match_group(t.particulars, rules)
    u = norm(t.particulars)
    up = t.particulars.upper()
    if RE_RETURN.search(u) and RE_CHQ_CTX.search(u):
        t.ret_type = "Cheque return"
    elif t.kind == "CR" and RE_EPAY_RETURN.search(up):
        t.ret_type = "Payment returned"          # e.g. NEFT RETURN - our payment bounced back
    t.is_return = bool(t.ret_type)
    t.party = counterparty(t.particulars)
    if t.kind == "CR":
        if t.ret_type == "Cheque return":
            t.category = "Cheque return (credited back)"
        elif t.ret_type:
            t.category = "Payment returned (credited back)"
        elif t.group == OWNER:
            t.category = "Owner funds in"
        elif t.group == LOAN:
            t.category = "Loan/Drawdown credit"
        elif t.group == CHARGES:
            t.category = "Charges reversed"
        else:
            t.category = t.mode
            t.is_collection = True
    else:
        if t.is_return:
            t.category = CHARGES if t.group == CHARGES else CHQ_RETURN
        elif t.group:
            t.category = t.group
        else:
            t.category = OTHERS


def week_of(d):
    return WEEKS[min((d.day - 1) // 7, 4)]


# ----------------------------------------------------------------------------
# Analysis
# ----------------------------------------------------------------------------
def party_key(t):
    if t.category not in (OTHERS, ""):
        return t.category
    return "?" + re.sub(r"[^A-Z0-9]", "", (t.party or t.particulars[:25]).upper())


def find_flags(txns):
    flags = []   # dict(month, type, date, amount, detail, txns)
    pays = [t for t in txns if t.kind == "DR" and t.category not in (CHARGES, CHQ_RETURN, LOAN)
            and t.amount >= DUPLICATE_MIN_AMOUNT * 100]
    for i, a in enumerate(pays):
        for b in pays[i + 1:]:
            if b.amount == a.amount and abs((b.date - a.date).days) <= DUPLICATE_WINDOW_DAYS \
                    and party_key(a) != party_key(b):
                flags.append({"month": b.month, "type": "Same amount, 2 parties", "date": b.date,
                              "amount": b.amount,
                              "detail": f"{a.date:%d-%m} {short_party(a)}  AND  {b.date:%d-%m} {short_party(b)}"
                                        " - possible wrong payment, please check"})
    for i, a in enumerate(pays):
        for b in pays[i + 1:]:
            if b.amount == a.amount and abs((b.date - a.date).days) <= DUPLICATE_WINDOW_DAYS \
                    and party_key(a) == party_key(b) and a.category not in NO_REPEAT_FLAG:
                flags.append({"month": b.month, "type": "Same amount, same party", "date": b.date,
                              "amount": b.amount,
                              "detail": f"{a.date:%d-%m} and {b.date:%d-%m} to {short_party(b)}"
                                        " - possible double payment, please check"})
    for t in txns:
        if t.is_return:
            flags.append({"month": t.month, "type": t.ret_type, "date": t.date, "amount": t.amount,
                          "detail": f"{'Credit' if t.kind == 'CR' else 'Debit'}: {t.particulars[:90]}"})
    for t in txns:
        if t.kind == "DR" and t.category != SAM and t.amount > LARGE_PAYMENT_LIMIT * 100:
            flags.append({"month": t.month, "type": "Large non-SAM payment", "date": t.date, "amount": t.amount,
                          "detail": f"{t.category}: {t.particulars[:90]}"})
    for t in txns:
        if t.check.startswith("MISMATCH"):
            flags.append({"month": t.month, "type": "Balance mismatch", "date": t.date, "amount": t.amount,
                          "detail": f"{t.check} - {t.particulars[:60]}"})
    flags.sort(key=lambda f: (f["date"], f["type"]))
    return flags


def short_party(t):
    name = t.category if t.category != OTHERS else (t.party or t.particulars[:30])
    return f"{name} ({t.mode})"


def payment_groups(txns):
    seen = []
    for t in txns:
        if t.kind == "DR" and t.category not in seen:
            seen.append(t.category)
    custom = sorted(g for g in seen if g not in SUPPLIER_ORDER + SPECIAL_ORDER + [OTHERS])
    return SUPPLIER_ORDER + custom + [OTHERS] + SPECIAL_ORDER


def month_stats(txns, mkey, flags, groups):
    mt = [t for t in txns if t.month == mkey]
    s = {"key": mkey, "txns": mt}
    s["first"] = min(t.bdate for t in mt)
    s["last"] = max(t.bdate for t in mt)
    f0 = mt[0]
    s["opening"] = f0.balance - f0.amount if f0.kind == "CR" else f0.balance + f0.amount
    s["closing"] = mt[-1].balance
    cols = [t for t in mt if t.is_collection]
    s["collections"] = sum(t.amount for t in cols)
    s["coll_count"] = len(cols)
    s["by_mode"] = {m: sum(t.amount for t in cols if t.category == m) for m in MODES}
    s["mode_count"] = {m: sum(1 for t in cols if t.category == m) for m in MODES}
    s["by_week"] = {w: sum(t.amount for t in cols if week_of(t.bdate) == w) for w in WEEKS}
    s["week_mode"] = {w: {m: sum(t.amount for t in cols if week_of(t.bdate) == w and t.category == m)
                          for m in MODES} for w in WEEKS}
    other_cr = defaultdict(int)
    for t in mt:
        if t.kind == "CR" and not t.is_collection:
            other_cr[t.category] += t.amount
    s["other_credits"] = dict(other_cr)
    s["total_credits"] = sum(t.amount for t in mt if t.kind == "CR")
    s["total_debits"] = sum(t.amount for t in mt if t.kind == "DR")
    s["pay"] = {g: sum(t.amount for t in mt if t.kind == "DR" and t.category == g) for g in groups}
    s["pay_count"] = {g: sum(1 for t in mt if t.kind == "DR" and t.category == g) for g in groups}
    sam = [t for t in mt if t.kind == "DR" and t.category == SAM]
    s["sam_count"] = len(sam)
    s["sam_total"] = sum(t.amount for t in sam)
    s["sam_avg"] = round(s["sam_total"] / len(sam)) if sam else 0
    s["sam_max"] = max((t.amount for t in sam), default=0)
    s["flags"] = [f for f in flags if f["month"] == mkey]
    return s


def month_name(mkey):
    return date(mkey[0], mkey[1], 1).strftime("%b %Y")


# ----------------------------------------------------------------------------
# Excel output
# ----------------------------------------------------------------------------
INR_FMT = '#,##0.00;[Red]-#,##0.00'


def write_excel(path, months, cmp_months, txns, flags, files, groups, rules_path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    bold = Font(bold=True)
    title = Font(bold=True, size=14, color="1F3864")
    head_fill = PatternFill("solid", fgColor="1F3864")
    head_font = Font(bold=True, color="FFFFFF")
    sub_fill = PatternFill("solid", fgColor="D9E1F2")
    flag_fill = PatternFill("solid", fgColor="FCE4D6")
    ok_fill = PatternFill("solid", fgColor="E2EFDA")
    thin = Side(style="thin", color="BFBFBF")
    box = Border(top=thin, bottom=thin, left=thin, right=thin)

    def rs(p):
        return None if p is None else p / 100

    def header(ws, r, values, c0=1):
        for i, v in enumerate(values):
            c = ws.cell(r, c0 + i, v)
            c.fill, c.font, c.border = head_fill, head_font, box
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def put(ws, r, c, v, money=False, b=False, fill=None):
        cell = ws.cell(r, c, rs(v) if money else v)
        cell.border = box
        if money:
            cell.number_format = INR_FMT
        if b:
            cell.font = bold
        if fill:
            cell.fill = fill
        return cell

    def widths(ws, w):
        for i, x in enumerate(w, 1):
            ws.column_dimensions[get_column_letter(i)].width = x

    # ---------------- Summary ----------------
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Piston Agency - Canara Bank current account - Monthly cash report"
    ws["A1"].font = title
    ws["A2"] = f"Generated {datetime.now():%d-%m-%Y %H:%M}  |  {len(txns)} transactions  |  " \
               f"{txns[0].date:%d-%m-%Y} to {txns[-1].date:%d-%m-%Y}  |  payee rules: {os.path.basename(rules_path)}"
    bad = [t for t in txns if t.check.startswith("MISMATCH")]
    closing_bad = [f for f in files if f.closing_printed is not None and f.txns and
                   f.closing_printed != f.txns[-1].balance]
    ok = not bad and not closing_bad
    ws["A3"] = "Balance check: ALL BALANCES TALLY" if ok else \
        f"Balance check: PROBLEMS FOUND - see 'Checks' sheet ({len(bad)} row mismatches, " \
        f"{len(closing_bad)} closing balance differences)"
    ws["A3"].font = Font(bold=True, color="375623" if ok else "C00000")

    r = 5
    ws.cell(r, 1, f"Last {len(cmp_months)} months comparison (Rs)").font = Font(bold=True, size=12)
    r += 1
    header(ws, r, ["Item"] + [month_name(m["key"]) for m in cmp_months] + ["Average"])
    r += 1

    def row(label, getter, money=True, b=False, fill=None, avg=True):
        nonlocal r
        put(ws, r, 1, label, b=b, fill=fill)
        vals = []
        for i, m in enumerate(cmp_months):
            v = getter(m)
            vals.append(v)
            put(ws, r, 2 + i, v, money=money, b=b, fill=fill)
        if avg and vals:
            a = sum(vals) / len(vals)
            c = put(ws, r, 2 + len(cmp_months), round(a) if money else round(a, 1), money=money, b=b, fill=fill)
            if not money:
                c.number_format = "0.0"
        else:
            put(ws, r, 2 + len(cmp_months), "", fill=fill)
        r += 1

    row("Days covered", lambda m: f"{m['first']:%d %b} - {m['last']:%d %b}", money=False, avg=False)
    row("Opening balance", lambda m: m["opening"], avg=False)
    row("TOTAL COLLECTIONS", lambda m: m["collections"], b=True, fill=sub_fill)
    for md in MODES:
        row(f"   {md}", lambda m, md=md: m["by_mode"][md])
    row("Number of collections", lambda m: m["coll_count"], money=False)
    other_keys = sorted({k for m in cmp_months for k in m["other_credits"]})
    for k in other_keys:
        row(f"Other credits: {k}", lambda m, k=k: m["other_credits"].get(k, 0))
    row("TOTAL PAYMENTS (all debits)", lambda m: m["total_debits"], b=True, fill=sub_fill)
    for g in groups:
        if any(m["pay"].get(g) for m in cmp_months):
            row(f"   {g}", lambda m, g=g: m["pay"].get(g, 0))
    row("SAM - number of payments", lambda m: m["sam_count"], money=False, fill=sub_fill)
    row("SAM - total paid", lambda m: m["sam_total"], fill=sub_fill)
    row("SAM - average payment", lambda m: m["sam_avg"], fill=sub_fill)
    row("Net cash flow (credits - debits)", lambda m: m["total_credits"] - m["total_debits"], b=True)
    row("Closing balance", lambda m: m["closing"], b=True, avg=False)
    row("Flags raised", lambda m: len(m["flags"]), money=False)
    ws.cell(r + 1, 1, "Average = simple average of the months shown. A month with only part of its days "
                      "in the PDFs ('Days covered') will pull the average down.").font = Font(italic=True, size=9)
    widths(ws, [38] + [16] * (len(cmp_months) + 1))
    ws.freeze_panes = "B7"

    # ---------------- Month sheets ----------------
    for m in months:
        ws = wb.create_sheet(month_name(m["key"]))
        ws["A1"] = f"{month_name(m['key'])} - cash report"
        ws["A1"].font = title
        ws["A2"] = f"Transactions {m['first']:%d-%m-%Y} to {m['last']:%d-%m-%Y}  |  " \
                   f"Opening {inr(m['opening'])}  |  Closing {inr(m['closing'])}"
        r = 4
        ws.cell(r, 1, "COLLECTIONS by mode and week (Rs)").font = Font(bold=True, size=12)
        r += 1
        header(ws, r, ["Week"] + MODES + ["Total"])
        r += 1
        for w in WEEKS:
            put(ws, r, 1, w)
            for i, md in enumerate(MODES):
                put(ws, r, 2 + i, m["week_mode"][w][md], money=True)
            put(ws, r, 2 + len(MODES), m["by_week"][w], money=True, b=True)
            r += 1
        put(ws, r, 1, "Total", b=True, fill=sub_fill)
        for i, md in enumerate(MODES):
            put(ws, r, 2 + i, m["by_mode"][md], money=True, b=True, fill=sub_fill)
        put(ws, r, 2 + len(MODES), m["collections"], money=True, b=True, fill=sub_fill)
        r += 1
        put(ws, r, 1, "No. of receipts")
        for i, md in enumerate(MODES):
            put(ws, r, 2 + i, m["mode_count"][md])
        put(ws, r, 2 + len(MODES), m["coll_count"])
        r += 2
        if m["other_credits"]:
            ws.cell(r, 1, "Other credits (not counted as collections)").font = bold
            r += 1
            for k, v in m["other_credits"].items():
                put(ws, r, 1, k)
                put(ws, r, 2, v, money=True)
                r += 1
            r += 1

        ws.cell(r, 1, "PAYMENTS by payee group (Rs)").font = Font(bold=True, size=12)
        r += 1
        header(ws, r, ["Payee group", "No.", "Total", "Average", "% of payments"])
        r += 1
        tot = m["total_debits"] or 1
        for g in groups:
            n, v = m["pay_count"][g], m["pay"][g]
            if not n and g not in SUPPLIER_ORDER:
                continue
            fill = sub_fill if g == SAM else None
            put(ws, r, 1, g, b=(g == SAM), fill=fill)
            put(ws, r, 2, n, fill=fill)
            put(ws, r, 3, v, money=True, fill=fill)
            put(ws, r, 4, round(v / n) if n else 0, money=True, fill=fill)
            c = put(ws, r, 5, v / tot, fill=fill)
            c.number_format = "0.0%"
            r += 1
        put(ws, r, 1, "Total payments", b=True, fill=sub_fill)
        put(ws, r, 2, sum(m["pay_count"].values()), b=True, fill=sub_fill)
        put(ws, r, 3, m["total_debits"], money=True, b=True, fill=sub_fill)
        put(ws, r, 4, "", fill=sub_fill)
        put(ws, r, 5, "", fill=sub_fill)
        r += 2

        ws.cell(r, 1, "SAM (SAMKRG) payments").font = Font(bold=True, size=12)
        r += 1
        for lab, v, mon in [("Number of payments", m["sam_count"], False), ("Total paid", m["sam_total"], True),
                            ("Average payment", m["sam_avg"], True), ("Largest payment", m["sam_max"], True)]:
            put(ws, r, 1, lab)
            put(ws, r, 2, v, money=mon)
            r += 1
        r += 1
        ws.cell(r, 1, f"FLAGS ({len(m['flags'])})").font = Font(bold=True, size=12, color="C00000")
        r += 1
        if m["flags"]:
            header(ws, r, ["Flag", "Date", "Amount", "Details"])
            r += 1
            for f in m["flags"]:
                put(ws, r, 1, f["type"], fill=flag_fill)
                c = put(ws, r, 2, f["date"], fill=flag_fill)
                c.number_format = "DD-MM-YYYY"
                put(ws, r, 3, f["amount"], money=True, fill=flag_fill)
                put(ws, r, 4, f["detail"], fill=flag_fill)
                ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
                r += 1
        else:
            put(ws, r, 1, "No flags this month", fill=ok_fill)
        widths(ws, [30, 14, 16, 16, 16, 16, 16, 16])

    # ---------------- Flags ----------------
    ws = wb.create_sheet("Flags")
    header(ws, 1, ["Month", "Flag", "Date", "Amount", "Details"])
    for i, f in enumerate(flags, 2):
        put(ws, i, 1, month_name(f["month"]))
        put(ws, i, 2, f["type"])
        put(ws, i, 3, f["date"]).number_format = "DD-MM-YYYY"
        put(ws, i, 4, f["amount"], money=True)
        put(ws, i, 5, f["detail"])
    if not flags:
        ws.cell(2, 1, "No flags")
    widths(ws, [10, 24, 12, 15, 110])
    ws.freeze_panes = "A2"

    # ---------------- Transactions ----------------
    ws = wb.create_sheet("Transactions")
    header(ws, 1, ["Date", "Particulars", "Deposit", "Withdrawal", "Balance", "Mode", "Category / payee group",
                   "Collection?", "Other party (guess)", "Week", "Balance check", "File", "Page",
                   "Counted on (back-dated rows)"])
    for i, t in enumerate(txns, 2):
        put(ws, i, 1, t.date).number_format = "DD-MM-YYYY"
        put(ws, i, 2, t.particulars)
        put(ws, i, 3, t.amount if t.kind == "CR" else None, money=True)
        put(ws, i, 4, t.amount if t.kind == "DR" else None, money=True)
        put(ws, i, 5, t.balance, money=True)
        put(ws, i, 6, t.mode)
        put(ws, i, 7, t.category)
        put(ws, i, 8, "Yes" if t.is_collection else "")
        put(ws, i, 9, t.party)
        put(ws, i, 10, week_of(t.bdate))
        put(ws, i, 11, t.check, fill=None if t.check.startswith("OK") else flag_fill)
        put(ws, i, 12, t.file)
        put(ws, i, 13, t.page)
        c = put(ws, i, 14, t.book_date)
        if t.book_date:
            c.number_format = "DD-MM-YYYY"
    widths(ws, [11, 70, 14, 14, 15, 15, 24, 11, 28, 15, 14, 26, 6, 14])
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:N{len(txns) + 1}"

    # ---------------- Unmapped payees ----------------
    ws = wb.create_sheet("Unmapped payees")
    ws["A1"] = "Payments that fell into 'Others'. Add a keyword for any of these to payees.csv and run again."
    ws["A1"].font = bold
    agg = defaultdict(lambda: [0, 0, None, None])
    for t in txns:
        if t.kind == "DR" and t.category == OTHERS:
            name = t.party or t.particulars[:40]
            a_ = agg[re.sub(r"[^A-Z0-9]", "", name.upper())]   # 'NARAN G MOTORS' = 'NARANG MOTORS'
            a_[0] += 1
            a_[1] += t.amount
            a_[2] = a_[2] or t.particulars
            if a_[3] is None or len(name.split()) < len(a_[3].split()):
                a_[3] = name
    header(ws, 3, ["Other party (guess)", "No.", "Total", "Example particulars", "Suggested payees.csv line"])
    for i, (n, v, ex, k) in enumerate(sorted(agg.values(), key=lambda x: -x[1]), 4):
        put(ws, i, 1, k)
        put(ws, i, 2, n)
        put(ws, i, 3, v, money=True)
        put(ws, i, 4, ex)
        put(ws, i, 5, f"{k},<group name>")
    widths(ws, [32, 6, 15, 80, 30])

    # ---------------- Checks ----------------
    ws = wb.create_sheet("Checks")
    header(ws, 1, ["File", "Pages", "Transactions", "Opening balance", "Opening from", "Deposits", "Withdrawals",
                   "Last balance", "Printed closing", "Result", "Notes"])
    for i, f in enumerate(files, 2):
        dep = sum(t.amount for t in f.txns if t.kind == "CR")
        wd = sum(t.amount for t in f.txns if t.kind == "DR")
        last = f.txns[-1].balance if f.txns else None
        problems = [t for t in f.txns if t.check.startswith("MISMATCH")]
        arith = f.opening is not None and last is not None and f.opening + dep - wd == last
        res = []
        if problems:
            res.append(f"{len(problems)} row mismatch(es)")
        if f.closing_printed is not None and last is not None and f.closing_printed != last:
            res.append(f"closing differs by {inr(f.closing_printed - last)}")
        if not arith and f.txns:
            res.append("opening + deposits - withdrawals does not equal last balance")
        put(ws, i, 1, f.file)
        put(ws, i, 2, f.pages)
        put(ws, i, 3, len(f.txns))
        put(ws, i, 4, f.opening, money=True)
        put(ws, i, 5, f.opening_source)
        put(ws, i, 6, dep, money=True)
        put(ws, i, 7, wd, money=True)
        put(ws, i, 8, last, money=True)
        put(ws, i, 9, f.closing_printed if f.closing_printed is not None else "not printed",
            money=f.closing_printed is not None)
        put(ws, i, 10, "OK - tallies" if not res else "; ".join(res), fill=ok_fill if not res else flag_fill)
        put(ws, i, 11, " | ".join(f.notes))
    r = len(files) + 3
    ws.cell(r, 1, "Rows where balance does not tally (these usually mean a row was misread or is missing):").font = bold
    r += 1
    bad_rows = [t for t in txns if not t.check.startswith("OK")]
    if not bad_rows:
        ws.cell(r, 1, "None - every row's balance = previous balance +/- amount.")
    for t in bad_rows:
        ws.cell(r, 1, f"{t.date:%d-%m-%Y}  {t.file} p{t.page}  {t.check}  {t.particulars[:80]}")
        r += 1
    widths(ws, [32, 7, 12, 16, 22, 16, 16, 16, 16, 40, 60])

    wb.save(path)


# ----------------------------------------------------------------------------
# Console summary
# ----------------------------------------------------------------------------
def print_summary(months, cmp_months, txns, flags, files, groups, out_path):
    line = "=" * 78
    print(line)
    print(" PISTON AGENCY - CANARA BANK CURRENT ACCOUNT - CASH REPORT")
    print(line)
    for f in files:
        last = f.txns[-1].balance if f.txns else None
        bad = sum(1 for t in f.txns if t.check.startswith("MISMATCH"))
        cl = ""
        if f.closing_printed is not None:
            cl = "closing matches" if f.closing_printed == last else \
                f"CLOSING DIFFERS (printed {inr(f.closing_printed)}, computed {inr(last)})"
        status = "OK" if not bad else f"{bad} ROW MISMATCHES"
        print(f" {f.file[:38]:<38} {len(f.txns):>4} txns  balances {status}  {cl}")
        for n in f.notes:
            print(f"     note: {n}")
    print(f" Period: {txns[0].date:%d-%m-%Y} to {txns[-1].date:%d-%m-%Y}   Closing balance: Rs {inr(txns[-1].balance)}")

    for m in cmp_months:
        print("\n" + "-" * 78)
        print(f" {month_name(m['key']).upper()}   ({m['first']:%d %b} - {m['last']:%d %b})   "
              f"Opening {inr(m['opening'])}  Closing {inr(m['closing'])}")
        print("-" * 78)
        print(f" Collections: Rs {inr(m['collections'])}  ({m['coll_count']} receipts)")
        print("   " + "   ".join(f"{md}: {inr(m['by_mode'][md])}" for md in MODES if m["by_mode"][md]))
        print("   " + "   ".join(f"W{i + 1}: {inr(m['by_week'][w])}" for i, w in enumerate(WEEKS)
                                 if m["by_week"][w] or i < 4))
        print(f" Payments:    Rs {inr(m['total_debits'])}")
        for g in groups:
            if m["pay_count"][g]:
                print(f"   {g:<22} {m['pay_count'][g]:>3} x   Rs {inr(m['pay'][g]):>15}")
        print(f" SAM: {m['sam_count']} payments, total Rs {inr(m['sam_total'])}, average Rs {inr(m['sam_avg'])}")
        if m["flags"]:
            print(f" FLAGS ({len(m['flags'])}):")
            for f in m["flags"]:
                print(f"   ! {f['date']:%d-%m} {f['type']:<23} Rs {inr(f['amount']):>13}  {f['detail'][:60]}")
        else:
            print(" Flags: none")

    print("\n" + "-" * 78)
    print(f" {len(cmp_months)}-MONTH COMPARISON (Rs)")
    print("-" * 78)
    hdr = f" {'':<24}" + "".join(f"{month_name(m['key']):>14}" for m in cmp_months) + f"{'Average':>14}"
    print(hdr)

    def pr(label, fn, money=True):
        vals = [fn(m) for m in cmp_months]
        avg = sum(vals) / len(vals)
        fmt = (lambda v: inr(round(v))) if money else (lambda v: f"{v:,.1f}" if isinstance(v, float) else str(v))
        print(f" {label:<24}" + "".join(f"{fmt(v):>14}" for v in vals) + f"{fmt(avg):>14}")

    pr("Collections", lambda m: m["collections"])
    pr("Payments (all)", lambda m: m["total_debits"])
    pr("SAM total", lambda m: m["sam_total"])
    pr("SAM no. of payments", lambda m: m["sam_count"], money=False)
    pr("SAM average", lambda m: m["sam_avg"])
    pr("Net cash flow", lambda m: m["total_credits"] - m["total_debits"])
    pr("Flags", lambda m: len(m["flags"]), money=False)
    print(line)
    print(f" Excel report saved: {out_path}")
    print(line)


# ----------------------------------------------------------------------------
def main():
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description="Canara Bank e-passbook PDFs -> monthly cash report")
    ap.add_argument("inputs", nargs="*", help="PDF files or folders (default: pdfs folder next to this script)")
    ap.add_argument("--out", default=os.path.join(here, "output"), help="output folder")
    ap.add_argument("--payees", default=os.path.join(here, "payees.csv"), help="payee mapping CSV")
    ap.add_argument("--password", default=os.environ.get("PDF_PASSWORD", ""), help="PDF password, if any")
    ap.add_argument("--months", type=int, default=COMPARE_MONTHS, help="months in comparison table")
    ap.add_argument("--dump", action="store_true", help="print every parsed row")
    a = ap.parse_args()

    inputs = a.inputs or [os.path.join(here, "pdfs")]
    pdfs = []
    for p in inputs:
        if os.path.isdir(p):
            pdfs += [os.path.join(p, f) for f in sorted(os.listdir(p)) if f.lower().endswith(".pdf")]
        elif p.lower().endswith(".pdf") and os.path.exists(p):
            pdfs.append(p)
        else:
            print(f"Not found: {p}")
    if not pdfs:
        print("No PDF files found. Put the Canara e-passbook PDFs in the 'pdfs' folder and run again.")
        return 1

    files = []
    for p in pdfs:
        try:
            files.append(read_pdf(p, a.password, a.dump))
        except Exception as e:  # noqa: BLE001
            msg = str(e) or repr(e)
            if "password" in repr(e).lower() or "encrypt" in repr(e).lower():
                msg = "PDF is password protected - set PDF_PASSWORD in run_report.bat (or use --password)"
            print(f"ERROR reading {os.path.basename(p)}: {msg}")
    files = [f for f in files if f.txns]
    if not files:
        print("No transactions could be read.")
        return 1

    # combine files in date order, dropping rows repeated in overlapping statements
    files.sort(key=lambda f: (f.txns[0].date, f.file))
    txns, seen = [], set()
    for f in files:
        for t in f.txns:
            key = (t.date, t.amount, t.kind, t.balance)
            if key in seen:
                continue
            seen.add(key)
            txns.append(t)
    for prev, t in zip(txns, txns[1:]):
        if t.file != prev.file and t.check.startswith("OK"):
            exp = prev.balance + (t.amount if t.kind == "CR" else -t.amount)
            if exp != t.balance:
                t.check = f"MISMATCH (gap between files: {prev.file} ends {inr(prev.balance)})"

    # Canara cuts statements by posting date; a row printed with an earlier date than the row
    # before it (e.g. 30-09 after 01-10) is counted on the latest date seen so far.
    latest = None
    for t in txns:
        latest = t.date if latest is None or t.date > latest else latest
        if latest != t.date:
            t.book_date = latest

    rules = load_payees(a.payees)
    for t in txns:
        classify(t, rules)
    groups = payment_groups(txns)
    flags = find_flags(txns)
    mkeys = sorted({t.month for t in txns})
    months = [month_stats(txns, k, flags, groups) for k in mkeys]
    cmp_months = months[-a.months:]

    os.makedirs(a.out, exist_ok=True)
    out = os.path.join(a.out, f"Cash_Report_{datetime.now():%Y%m%d_%H%M}.xlsx")
    write_excel(out, months, cmp_months, txns, flags, files, groups, a.payees)
    print_summary(months, cmp_months, txns, flags, files, groups, out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
