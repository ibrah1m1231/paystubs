#!/usr/bin/env python3
"""Cash workers: paystubs + time cards in one PDF.

Input is EITHER
  * a cash paystubs PDF (the "Cash Paystubs" export, one employee per page), or
  * the Weekly Pay .xlsx workbook (one tab per week; the cash columns are
    Hours / Per Hour / Cash, e.g. G / H / I).

Only the employees named in --names are kept (default "Marvin, Santiago";
pass an empty string to keep everyone). For each kept employee this builds a
time card whose daily hours add up to the paystub total, then puts the paystub
on top and the time card below, one employee per page.

When the input is the workbook, the paystubs themselves are generated in the
same layout as the cash paystub PDF. Pay is turned into hours the same way the
cash paystubs do it: amount / rate, first 40 hours regular, the rest overtime
at 1.5x, rounded up to the next minute.

Usage:
    python3 cash.py --input "Weekly Pay.xlsx" --out combined.pdf [--sheet "11.02-11.08"]
    python3 cash.py --input "Cash Paystubs.pdf" --out combined.pdf [--names "Marvin, Santiago"]
    python3 cash.py --input "Weekly Pay.xlsx" --list-sheets
"""
import argparse
import datetime
import json
import math
import os
import re
import subprocess
import sys
import tempfile

import openpyxl
import pdfplumber
from pypdf import PdfReader, PdfWriter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import generate_timecards as gt  # noqa: E402
import combine as cb  # noqa: E402

DEFAULT_NAMES = "Marvin, Santiago"
DEFAULT_RATE = 17.00
OT_AFTER_MIN = 40 * 60
OT_MULT = 1.5
COMPANY = ["KABABJEES LLC", "495-18 S BROADWAY", "HICKSVILLE", "NY 11801"]
COMPANY_FOOTER = "KABABJEES LLC, 495-18 S BROADWAY, HICKSVILLE, NY 11801"
PAY_DATE_AFTER_END = 4  # period ends Sunday, paid the following Thursday


def fail(msg):
    print(msg, file=sys.stderr)
    sys.exit(1)


def norm(s):
    return re.sub(r"[^a-z]", "", (s or "").lower())


def clean_name(s):
    s = re.sub(r"\*{3}-\*{2}-\d+", "", s or "")
    s = s.split(",")[0]
    s = re.sub(r"[\s.]+$", "", s.strip())
    return re.sub(r"\s+", " ", s).strip()


def name_filter(raw):
    toks = [norm(t) for t in re.split(r"[,;\n]+", raw or "")]
    return [t for t in toks if t]


def keep(name, toks):
    if not toks:
        return True
    n = norm(name)
    return any(t in n for t in toks)


def min_to_hm(m):
    return f"{m // 60}:{m % 60:02d}"


def hm_to_min(s):
    h, m = s.split(":")
    return int(h) * 60 + int(m)


def fmt_date(d):
    return f"{d.month:02d}/{d.day:02d}/{d.year}"


def money(x):
    return f"{x:,.2f}"


# --------------------------------------------------------------------------
# Paystub PDF input
# --------------------------------------------------------------------------
def parse_cash_pdf(path):
    """Return list of {page, name, minutes, start, end} for every paystub page."""
    out = []
    with pdfplumber.open(path) as pdf:
        for idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            lines = [l.strip() for l in text.splitlines()]
            name = None
            for i, l in enumerate(lines):
                if re.match(r"^Employee\s+SSN\b", l) and i + 1 < len(lines):
                    name = clean_name(lines[i + 1])
                    break
            if not name:
                m = re.search(r"^([A-Za-z][A-Za-z .'-]+?),.*\*\*\*-\*\*-", text, re.M)
                name = clean_name(m.group(1)) if m else None
            mins = 0
            for m in re.finditer(r"^(?:Hourly|Over\s*Time|Overtime|Double\s*Time)\s+(\d{1,3}:\d{2})\s",
                                 text, re.M | re.I):
                mins += hm_to_min(m.group(1))
            pm = re.search(r"Pay Period:\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})", text)
            start = end = None
            if pm:
                start = datetime.datetime.strptime(pm.group(1), "%m/%d/%Y").date()
                end = datetime.datetime.strptime(pm.group(2), "%m/%d/%Y").date()
            if name and mins:
                out.append({"page": idx, "name": name.title() if name.isupper() else name,
                            "minutes": mins, "start": start, "end": end})
    return out


def from_pdf(path, toks, workdir):
    stubs = parse_cash_pdf(path)
    if not stubs:
        fail("Couldn't find any paystubs with hours in that PDF.")
    picked = [s for s in stubs if keep(s["name"], toks)]
    if not picked:
        fail("None of the names you entered were found. Names in this PDF: "
             + ", ".join(s["name"] for s in stubs))
    reader = PdfReader(path)
    writer = PdfWriter()
    for s in picked:
        writer.add_page(reader.pages[s["page"]])
    ps_pdf = os.path.join(workdir, "cash-paystubs.pdf")
    with open(ps_pdf, "wb") as fh:
        writer.write(fh)
    emps = [(s["name"], min_to_hm(s["minutes"])) for s in picked]
    start = next((s["start"] for s in picked if s["start"]), None)
    return ps_pdf, emps, start


# --------------------------------------------------------------------------
# Weekly Pay workbook input
# --------------------------------------------------------------------------
def list_sheets(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


def sheet_dates(sheet, today=None):
    """'11.02-11.08' -> (date(…,11,2), date(…,11,8)). The year is the latest
    one that doesn't put the week more than a month in the future."""
    m = re.match(r"\s*(\d{1,2})[./](\d{1,2})\s*-\s*(\d{1,2})[./](\d{1,2})", sheet)
    if not m:
        return None, None
    m1, d1, m2, d2 = map(int, m.groups())
    today = today or datetime.date.today()
    limit = today + datetime.timedelta(days=31)
    for y in range(today.year + 1, today.year - 10, -1):
        try:
            start = datetime.date(y, m1, d1)
        except ValueError:
            continue
        if start <= limit:
            ey = y + 1 if (m2, d2) < (m1, d1) else y
            try:
                return start, datetime.date(ey, m2, d2)
            except ValueError:
                return start, start + datetime.timedelta(days=6)
    return None, None


_SAFE = re.compile(r"^[\d\s.+\-*/()]+$")


def num(v):
    """Cell value -> float, evaluating plain-arithmetic formulas like '=208.33*6'."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().lstrip("=").replace(",", "")
    if s and _SAFE.match(s):
        try:
            return float(eval(s, {"__builtins__": {}}, {}))  # digits/operators only
        except Exception:
            return None
    return None


def read_cash_rows(path, sheet):
    wb_v = openpyxl.load_workbook(path, data_only=True, read_only=True)
    wb_f = openpyxl.load_workbook(path, data_only=False, read_only=True)
    if sheet not in wb_v.sheetnames:
        fail(f'Tab "{sheet}" not found in the workbook.')
    rows_v = list(wb_v[sheet].iter_rows(values_only=True))
    rows_f = list(wb_f[sheet].iter_rows(values_only=True))
    wb_v.close(); wb_f.close()
    if not rows_v:
        fail(f'Tab "{sheet}" is empty.')

    # Find the "Cash" header; hours and rate sit just left of it (G / H / I).
    cash_col = None
    for r in rows_v[:5]:
        for c, v in enumerate(r or ()):
            if isinstance(v, str) and v.strip().lower() == "cash":
                cash_col = c
                break
        if cash_col is not None:
            break
    if cash_col is None or cash_col < 2:
        fail(f'Couldn\'t find a "Cash" column on tab "{sheet}".')
    hrs_col, rate_col = cash_col - 2, cash_col - 1

    def cell(rows, r, c):
        row = rows[r] if r < len(rows) else ()
        return row[c] if row and c < len(row) else None

    out = []
    for r in range(1, len(rows_v)):
        name = cell(rows_v, r, 0)
        if not isinstance(name, str) or not name.strip():
            continue
        name = name.strip()
        if "," in name:  # "Chauhan, Pritpal" -> "Pritpal Chauhan"
            last, first = [p.strip() for p in name.split(",", 1)]
            name = f"{first} {last}".strip()
        if name.isupper() or any(w[:1].islower() for w in name.split()):
            name = name.title()  # "JASHANDEEP SINGH" / "Marvin morales"
        amount = num(cell(rows_v, r, cash_col))
        if amount is None:
            amount = num(cell(rows_f, r, cash_col))
        if amount is None:
            h, rt = num(cell(rows_v, r, hrs_col)), num(cell(rows_v, r, rate_col))
            if h is not None and rt is not None:
                amount = h * rt
        if amount is None or amount <= 0.004:
            continue
        out.append({"name": name, "amount": round(amount, 2),
                    "hours": num(cell(rows_v, r, hrs_col)),
                    "cash_rate": num(cell(rows_v, r, rate_col))})
    return out


def split_pay(amount, rate):
    """Dollar amount -> (reg_minutes, ot_minutes), rounded up to the minute."""
    if amount <= rate * OT_AFTER_MIN / 60 + 1e-9:
        reg = math.ceil(amount / rate * 60 - 1e-6)
        return min(reg, OT_AFTER_MIN), 0
    ot = math.ceil((amount - rate * OT_AFTER_MIN / 60) / (rate * OT_MULT) * 60 - 1e-6)
    return OT_AFTER_MIN, ot


def render_paystubs(stubs, start, end, pay_date, out_path):
    from reportlab.pdfgen import canvas

    W, H = 612.0, 792.0
    c = canvas.Canvas(out_path, pagesize=(W, H))
    c.setTitle("Cash Paystubs")

    def y(top):  # pdfplumber "top" (from page top) -> reportlab baseline
        return H - top - 6.0

    def txt(x, top, s, bold=False, size=6.9, right=False, font=None):
        c.setFont(font or ("Helvetica-Bold" if bold else "Helvetica"), size)
        (c.drawRightString if right else c.drawString)(x, y(top), s)

    def hline(x0, x1, top, w=0.96):
        c.setLineWidth(w)
        c.line(x0, H - top, x1, H - top)

    period = f"Pay Period: {fmt_date(start)} - {fmt_date(end)}" if start and end else "Pay Period:"
    pdate = f"Pay Date: {fmt_date(pay_date)}" if pay_date else "Pay Date:"
    for s in stubs:
        for i, line in enumerate(COMPANY):
            txt(40.6, 33.4 + 11.75 * i, line, size=9.9, font="Times-Roman")
        txt(72.0, 155.0, s["name"], size=9.9, font="Times-Roman")

        c.setDash(1.5, 1.5); hline(14, 585, 252.0, 0.375); c.setDash()
        txt(22.6, 254.9, "Employee Pay Stub", bold=True)
        txt(153.0, 254.9, "Check number:")
        txt(310.5, 254.9, period)
        txt(476.9, 254.9, pdate)

        txt(22.6, 274.4, "Employee", bold=True)
        txt(310.6, 274.4, "SSN", bold=True)
        hline(23, 297, 283.6); hline(311, 585, 283.6)
        txt(22.6, 286.5, s["name"])

        QX, RX, CX = 153.0, 189.0, 238.5
        top = 314.8
        txt(22.6, top, "Earnings and Hours", bold=True)
        txt(QX, top, "Qty", bold=True, right=True)
        txt(RX, top, "Rate", bold=True, right=True)
        txt(CX, top, "Current", bold=True, right=True)
        hline(23, 239, top + 9.2)
        top += 12.2
        reg_pay = round(s["reg"] / 60 * s["rate"], 2)
        ot_rate = round(s["rate"] * OT_MULT, 2)
        ot_pay = round(s["ot"] / 60 * ot_rate, 2)
        txt(22.6, top, "Hourly"); txt(QX, top, min_to_hm(s["reg"]), right=True)
        txt(RX, top, money(s["rate"]), right=True); txt(CX, top, money(reg_pay), right=True)
        total = reg_pay
        if s["ot"]:
            top += 8.2
            txt(22.6, top, "Over Time"); txt(QX, top, min_to_hm(s["ot"]), right=True)
            txt(RX, top, money(ot_rate), right=True); txt(CX, top, money(ot_pay), right=True)
            total = round(reg_pay + ot_pay, 2)
            hline(122, 239, top + 7.2)
            top += 10.8
            txt(QX, top, min_to_hm(s["reg"] + s["ot"]), right=True)
            txt(CX, top, money(total), right=True)
        top += 11.6
        txt(22.6, top, "Taxes", bold=True)
        txt(CX, top, "Current", bold=True, right=True)
        hline(23, 239, top + 9.2)
        top += 12.2
        txt(22.6, top, "Federal Withholding"); txt(CX, top, "0.00", right=True)
        top += 18.9
        txt(22.6, top, "Net Pay", bold=True); txt(CX, top, money(total), bold=True, right=True)

        txt(22.6, 736.5, COMPANY_FOOTER)
        c.showPage()
    c.save()


def from_xlsx(path, sheet, toks, rate, workdir):
    sheets = list_sheets(path)
    if not sheets:
        fail("The workbook has no tabs.")
    sheet = sheet or sheets[0]
    rows = read_cash_rows(path, sheet)
    if not rows:
        fail(f'No cash pay found on tab "{sheet}".')
    picked = [r for r in rows if keep(r["name"], toks)]
    if not picked:
        fail(f'None of the names you entered have cash pay on tab "{sheet}". '
             "Names with cash pay: " + ", ".join(r["name"] for r in rows))
    start, end = sheet_dates(sheet)
    pay_date = end + datetime.timedelta(days=PAY_DATE_AFTER_END) if end else None
    stubs = []
    for r in picked:
        reg, ot = split_pay(r["amount"], rate)
        stubs.append({"name": r["name"], "reg": reg, "ot": ot, "rate": rate})
        print(f"  {r['name']:<26} cash ${r['amount']:>9,.2f} -> "
              f"{min_to_hm(reg)} reg + {min_to_hm(ot)} OT")
    ps_pdf = os.path.join(workdir, "cash-paystubs.pdf")
    render_paystubs(stubs, start, end, pay_date, ps_pdf)
    emps = [(s["name"], min_to_hm(s["reg"] + s["ot"])) for s in stubs]
    return ps_pdf, emps, start, sheet


# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--input", required=True, help="cash paystubs .pdf or Weekly Pay .xlsx")
    ap.add_argument("--out", help="combined PDF to write")
    ap.add_argument("--names", default=DEFAULT_NAMES,
                    help='comma-separated names to keep (default "%(default)s"; "" = everyone)')
    ap.add_argument("--sheet", help="workbook tab (default: the first tab = latest week)")
    ap.add_argument("--rate", type=float, default=DEFAULT_RATE,
                    help="hourly rate printed on generated paystubs (default %(default).2f)")
    ap.add_argument("--daysoff", help="JSON file: {name: [weekday abbrevs]}")
    ap.add_argument("--workdir", help="folder for intermediate files (paystubs, time cards)")
    ap.add_argument("--list-sheets", action="store_true", help="print the workbook tabs as JSON")
    a = ap.parse_args()

    ext = os.path.splitext(a.input)[1].lower()
    if a.list_sheets:
        print(json.dumps(list_sheets(a.input)))
        return
    if not a.out:
        ap.error("--out is required")
    if a.rate <= 0:
        fail("Rate must be more than 0.")

    workdir = a.workdir or tempfile.mkdtemp(prefix="cash_")
    os.makedirs(workdir, exist_ok=True)
    toks = name_filter(a.names)

    if ext == ".pdf":
        ps_pdf, emps, start = from_pdf(a.input, toks, workdir)
        label = os.path.splitext(os.path.basename(a.input))[0]
    elif ext in (".xlsx", ".xlsm"):
        ps_pdf, emps, start, label = from_xlsx(a.input, a.sheet, toks, a.rate, workdir)
    else:
        fail("Please upload a paystubs .pdf or a Weekly Pay .xlsx.")
    print(f"Employees: {', '.join(n + ' ' + h for n, h in emps)}; week starting {start}")

    # 1) time cards (.xlsx). Cash workers have no saved schedule -> auto days.
    daysoff = {}
    if a.daysoff:
        daysoff, _ = gt.resolve_daysoff(emps, json.load(open(a.daysoff)))
    seed = int(start.strftime("%Y%m%d")) if start else 42
    tc_xlsx = os.path.join(workdir, "cash-timecards.xlsx")
    gt.build(emps, start, tc_xlsx, daysoff, seed, schedule={})
    if not gt.verify(tc_xlsx, daysoff):
        fail("Time card hours didn't add up to the paystub hours.")

    # 2) time cards -> PDF (LibreOffice)
    tc_pdf = os.path.join(workdir, "cash-timecards.pdf")
    proc = subprocess.run([sys.executable, os.path.join(HERE, "timecard_to_pdf.py"),
                           tc_xlsx, tc_pdf, "--skip-sheet", ""],
                          capture_output=True, text=True)
    if proc.returncode != 0 or not os.path.isfile(tc_pdf):
        fail("Time cards to PDF failed: " + (proc.stderr or proc.stdout or "").strip())

    # 3) paystub on top, time card below
    out, report = cb.combine(ps_pdf, tc_pdf, a.out)
    for pn, tn, r, ph, th in report:
        print(f"  {pn:<26} <- {tn:<26} paystub {ph} min / time card {th} min")
    print(f"Saved {out} ({len(report)} pages); label={label}")


if __name__ == "__main__":
    main()
