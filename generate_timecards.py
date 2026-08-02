#!/usr/bin/env python3
"""Generate a multi-sheet Excel time-card workbook from a W2 paystubs PDF.

One employee per sheet. Daily hours are synthesized to sum exactly to each
employee's paystub total (a model/mock, not a real attendance record).

Days off can be specified per employee via --daysoff (a JSON file mapping a
name/alias to a list of weekday abbreviations, e.g. {"Qadir": ["Tue","Fri"]}).
Aliases are fuzzy-matched to the paystub names; any names left unmatched on both
sides are paired by order. Employees with no entry fall back to auto scheduling.

Usage:
    python3 generate_timecards.py --pdf paystubs.pdf --out timecards.xlsx \
        [--daysoff daysoff.json] [--seed 42]
"""
import argparse, json, random, re, sys, datetime, difflib
import pdfplumber
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
ABBR = {"mon":0,"tue":1,"wed":2,"thu":3,"fri":4,"sat":5,"sun":6}
CLOSED_DAYS = {1}  # weekdays the restaurant is closed (Tue) -> always OFF for everyone

def hm_to_min(s):
    h, m = s.split(":"); return int(h)*60 + int(m)
def min_to_hm(t):
    return f"{t//60}:{t%60:02d}"

def parse_days(vals):
    out = []
    for v in vals:
        k = str(v).strip().lower()[:3]
        if k in ABBR: out.append(ABBR[k])
    return sorted(set(out))

def parse_pdf(path):
    """Return (employees, week_start_date). employees: list of (name, 'HH:MM')."""
    emps, period_start = [], None
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            m = re.search(r"Pay Period:\s*(\d{2}/\d{2}/\d{4})", text)
            if m and not period_start:
                period_start = datetime.datetime.strptime(m.group(1), "%m/%d/%Y").date()
            name = None
            ssn = re.search(r"^([A-Z][A-Z .'-]+?),.*\*\*\*-\*\*-", text, re.M)
            if ssn:
                name = ssn.group(1).strip()
            hrs = re.search(r"Hourly\s+(\d{1,3}:\d{2})", text)
            if name and hrs:
                emps.append((name.title(), hrs.group(1)))
    return emps, period_start

def _norm(s):
    return re.sub(r"[^a-z]", "", s.lower())

def resolve_daysoff(emps, spec):
    """Map spec {alias: [day-abbrevs]} to {full_employee_name: [day-indices]}.

    Fuzzy-match each alias to the closest paystub name; leftovers on both sides
    are paired by order (handles aliases that don't resemble the legal name).
    Returns (mapping, report) where report lists how each alias resolved.
    """
    if not spec: return {}, []
    names = [n for n, _ in emps]
    remaining = set(names)
    resolved, report = {}, []
    # score all alias/name pairs, assign greedily by best similarity
    scored = []
    for alias in spec:
        na = _norm(alias)
        for name in names:
            nn = _norm(name)
            r = difflib.SequenceMatcher(None, na, nn).ratio()
            if na and (na in nn or nn.startswith(na)):
                r = max(r, 0.9)
            scored.append((r, alias, name))
    scored.sort(reverse=True)
    used_alias = set()
    for r, alias, name in scored:
        if alias in used_alias or name not in remaining or r < 0.5:
            continue
        resolved[name] = parse_days(spec[alias]); remaining.discard(name)
        used_alias.add(alias); report.append((alias, name, r))
    # pair leftovers by order
    left_alias = [a for a in spec if a not in used_alias]
    left_names = [n for n in names if n in remaining]
    for alias, name in zip(left_alias, left_names):
        resolved[name] = parse_days(spec[alias])
        report.append((alias, name, 0.0))
    return resolved, report

def distribute(total_min, ndays, rng):
    base = total_min // ndays
    vals = [base]*ndays
    for i in range(total_min - base*ndays):
        vals[i] += 1
    for _ in range(ndays):
        i, j = rng.sample(range(ndays), 2)
        d = rng.randint(1, 4)
        vals[i] += d; vals[j] -= d
    return vals

DEFAULT_SCHEDULE = {
    "Mariya": {
        "Tue": [
            "10:57",
            352
        ],
        "Wed": [
            "10:56",
            358
        ],
        "Thu": [
            "15:59",
            320
        ],
        "Fri": [
            "10:58",
            354
        ],
        "Sat": [
            "10:58",
            320
        ],
        "Sun": [
            "10:58",
            364
        ]
    },
    "Reyna": {
        "Mon": [
            "15:57",
            347
        ],
        "Wed": [
            "15:58",
            350
        ],
        "Thu": [
            "11:58",
            318
        ],
        "Fri": [
            "16:02",
            342
        ],
        "Sat": [
            "16:26",
            326
        ],
        "Sun": [
            "16:32",
            379
        ]
    },
    "Taran": {
        "Mon": [
            "15:56",
            404
        ],
        "Tue": [
            "15:58",
            403
        ],
        "Thu": [
            "16:00",
            369
        ],
        "Fri": [
            "15:59",
            400
        ],
        "Sat": [
            "11:59",
            377
        ],
        "Sun": [
            "11:58",
            397
        ]
    },
    "Qadir": {
        "Mon": [
            "12:59",
            482
        ],
        "Wed": [
            "12:59",
            477
        ],
        "Thu": [
            "11:59",
            450
        ],
        "Sat": [
            "11:56",
            450
        ],
        "Sun": [
            "13:01",
            487
        ]
    },
    "Daniyal": {
        "Thu": [
            "12:58",
            468
        ],
        "Fri": [
            "12:57",
            470
        ],
        "Sun": [
            "12:56",
            390
        ]
    },
    "Gaurab": {
        "Wed": [
            "12:00",
            497
        ],
        "Thu": [
            "11:58",
            497
        ],
        "Fri": [
            "12:31",
            461
        ],
        "Sat": [
            "12:26",
            468
        ],
        "Sun": [
            "12:57",
            382
        ]
    },
    "Amarjit": {
        "Wed": [
            "14:56",
            485
        ],
        "Thu": [
            "14:59",
            479
        ],
        "Fri": [
            "15:26",
            453
        ],
        "Sat": [
            "15:27",
            451
        ],
        "Sun": [
            "14:50",
            501
        ]
    },
    "Darling": {
        "Mon": [
            "18:00",
            359
        ],
        "Tue": [
            "18:00",
            359
        ],
        "Wed": [
            "18:00",
            359
        ],
        "Thu": [
            "18:00",
            320
        ],
        "Sat": [
            "18:00",
            359
        ],
        "Sun": [
            "18:00",
            319
        ]
    },
    "Martha": {
        "Mon": [
            "9:58",
            397
        ],
        "Wed": [
            "10:02",
            393
        ],
        "Thu": [
            "9:56",
            402
        ],
        "Fri": [
            "10:03",
            389
        ],
        "Sat": [
            "9:59",
            398
        ],
        "Sun": [
            "10:01",
            317
        ]
    },
    "Ganesh Bahadur": {
        "Mon": [
            "11:58",
            297
        ],
        "Wed": [
            "12:02",
            292
        ],
        "Thu": [
            "11:57",
            295
        ],
        "Fri": [
            "12:03",
            289
        ],
        "Sat": [
            "11:59",
            294
        ],
        "Sun": [
            "12:01",
            307
        ]
    },
    "Joghishwar": {
        "Mon": [
            "12:01",
            348
        ],
        "Wed": [
            "12:05",
            343
        ],
        "Thu": [
            "12:00",
            345
        ],
        "Fri": [
            "12:06",
            341
        ],
        "Sat": [
            "12:02",
            344
        ],
        "Sun": [
            "12:04",
            341
        ]
    }
}


IDX2ABBR = {0:"Mon",1:"Tue",2:"Wed",3:"Thu",4:"Fri",5:"Sat",6:"Sun"}

def distribute_weighted(total_min, weights, rng):
    """Split total_min across days ~proportionally to weights, exact sum,
    with a little minute-level jitter so no two days look identical."""
    sw = sum(weights) or 1
    vals = [int(total_min * w / sw) for w in weights]
    rem = total_min - sum(vals)
    order = sorted(range(len(vals)), key=lambda k: -weights[k])
    i = 0
    while rem > 0 and order:
        vals[order[i % len(order)]] += 1; rem -= 1; i += 1
    if len(vals) > 1:
        for _ in range(len(vals)):
            a, b = rng.sample(range(len(vals)), 2)
            if vals[b] > 6:
                d = rng.randint(1, 4); vals[a] += d; vals[b] -= d
    return vals

def resolve_schedule(emps, spec):
    """Map schedule aliases (Mariya, Qadir, ...) to paystub employee names,
    fuzzy-matched like the days-off resolver. Returns {emp_name: profile}."""
    if not spec: return {}
    names = [n for n, _ in emps]
    remaining = set(names)
    resolved = {}
    scored = []
    for alias in spec:
        na = _norm(alias)
        for name in names:
            nn = _norm(name)
            r = difflib.SequenceMatcher(None, na, nn).ratio()
            if na and (na in nn or nn.startswith(na)):
                r = max(r, 0.9)
            scored.append((r, alias, name))
    scored.sort(reverse=True)
    used = set()
    for r, alias, name in scored:
        if alias in used or name not in remaining or r < 0.5:
            continue
        resolved[name] = spec[alias]; remaining.discard(name); used.add(alias)
    left_alias = [a for a in spec if a not in used]
    left_names = [n for n in names if n in remaining]
    for alias, name in zip(left_alias, left_names):
        resolved[name] = spec[alias]
    return resolved


def build(emps, week_start, out_path, daysoff=None, seed=42, schedule=None):
    daysoff = daysoff or {}
    schedule = schedule or {}
    rng = random.Random(seed)
    dates = [(week_start + datetime.timedelta(days=i)) for i in range(7)] if week_start else [None]*7
    date_strs = [f"{d.month}/{d.day}/{d.year}" if d else "" for d in dates]
    week_of = f"{week_start.month:02d}/{week_start.day:02d}/{week_start.year}" if week_start else ""

    navy, blue, green = "1F3864", "2E75B6", "375623"
    lightblue, lightgreen, gray = "DDEBF7", "E2EFDA", "F2F2F2"
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    wb = Workbook(); wb.remove(wb.active)
    for name, tot in emps:
        total_min = hm_to_min(tot)
        day_start = {}
        weights = None
        prof = schedule.get(name)
        if prof:
            items = [(ABBR[k.lower()], v) for k, v in prof.items()
                     if k.lower() in ABBR and ABBR[k.lower()] not in CLOSED_DAYS]
            items.sort()
            work_idx = [i for i, _ in items]
            weights = [v[1] for _, v in items]
            for i, v in items:
                day_start[i] = hm_to_min(v[0])
        else:
            if name in daysoff:
                off0 = set(daysoff[name])
            else:
                ndays = max(4, min(7, round(total_min/60/5.7)))
                off = 7 - ndays
                off0 = set([1,3,5,6,4,2,0][:off]) if off > 0 else set()
            work_idx = [i for i in range(7) if i not in off0 and i not in CLOSED_DAYS]
        # restaurant-closed days are always OFF; keep at least one working day
        work_idx = [i for i in work_idx if i not in CLOSED_DAYS]
        if not work_idx:
            work_idx = [next(d for d in range(7) if d not in CLOSED_DAYS)]
            weights, day_start = None, {}
        off_idx = set(range(7)) - set(work_idx)
        if weights is not None:
            shifts = distribute_weighted(total_min, weights, rng)
        else:
            shifts = distribute(total_min, len(work_idx), rng)
        day_hours = {work_idx[k]: shifts[k] for k in range(len(work_idx))}

        ws = wb.create_sheet(title=name[:31])
        for c, w in enumerate([12,13,11,12,11,16], start=1):
            ws.column_dimensions[get_column_letter(c)].width = w

        ws.merge_cells("A1:F1")
        t = ws["A1"]; t.value = "EMPLOYEE TIME CARD"
        t.font = Font(bold=True, color="FFFFFF", size=20)
        t.fill = PatternFill("solid", fgColor=navy); t.alignment = center
        ws.row_dimensions[1].height = 34

        ws.merge_cells("A2:C2"); ws.merge_cells("D2:F2")
        e = ws["A2"]; e.value = f"Employee:  {name}"
        e.font = Font(bold=True, size=12); e.fill = PatternFill("solid", fgColor=lightblue); e.alignment = center
        wk = ws["D2"]; wk.value = f"Week of:  {week_of}"
        wk.font = Font(bold=True, size=12); wk.fill = PatternFill("solid", fgColor=lightblue); wk.alignment = center
        ws.row_dimensions[2].height = 24
        for col in "ABCDEF": ws[f"{col}2"].border = border

        for c, h in enumerate(["Day","Date","Clock In","Clock Out","Hours","Notes"], start=1):
            cell = ws.cell(row=3, column=c, value=h)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill("solid", fgColor=blue); cell.alignment = center; cell.border = border
        ws.row_dimensions[3].height = 20

        r = 4
        for i in range(7):
            fill = gray if i % 2 == 1 else "FFFFFF"
            for c in range(1, 7):
                cell = ws.cell(row=r, column=c); cell.border = border
                cell.fill = PatternFill("solid", fgColor=fill); cell.alignment = center
            ws.cell(row=r, column=1, value=DAYS[i]).alignment = Alignment(horizontal="left", vertical="center")
            ws.cell(row=r, column=2, value=date_strs[i])
            if i in off_idx:
                oc = ws.cell(row=r, column=3, value="OFF"); oc.font = Font(italic=True, color="808080")
                ws.cell(row=r, column=5, value="-")
            else:
                if i in day_start:
                    cin = max(0, day_start[i] + rng.randint(-3, 4))
                else:
                    cin = 12*60 + rng.randint(0, 7)
                mins = day_hours[i]
                ws.cell(row=r, column=3, value=min_to_hm(cin))
                ws.cell(row=r, column=4, value=min_to_hm(cin+mins))
                ws.cell(row=r, column=5, value=min_to_hm(mins))
            ws.row_dimensions[r].height = 20
            r += 1

        ws.merge_cells(f"A{r}:D{r}")
        tc = ws.cell(row=r, column=1, value="TOTAL HOURS")
        tc.font = Font(bold=True, color="FFFFFF", size=12); tc.alignment = center
        for c in range(1, 5):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=green)
            ws.cell(row=r, column=c).border = border
        hc = ws.cell(row=r, column=5, value=tot)
        hc.font = Font(bold=True, size=13); hc.alignment = center
        hc.fill = PatternFill("solid", fgColor=lightgreen); hc.border = border
        ws.cell(row=r, column=6).fill = PatternFill("solid", fgColor=lightgreen)
        ws.cell(row=r, column=6).border = border
        ws.row_dimensions[r].height = 26
        r += 2

        ws.merge_cells(f"A{r}:C{r}"); ws.merge_cells(f"D{r}:F{r}")
        ws.cell(row=r, column=1, value="Employee Signature: ______________________").font = Font(size=10)
        ws.cell(row=r, column=4, value="Manager Signature: ______________________").font = Font(size=10)

        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_area = f"A1:F{r}"

    wb.save(out_path)
    return out_path

def verify(path, daysoff=None):
    daysoff = daysoff or {}
    wb = load_workbook(path); ok = True
    print(f"{'Sheet':<26} {'sum':>6} {'total':>6} {'off':>10}  status")
    for ws in wb.worksheets:
        s = 0; offs = []
        for i, row in enumerate(ws.iter_rows(min_row=4, max_row=10, min_col=3, max_col=5)):
            cin, _, hrs = row[0].value, row[1].value, row[2].value
            if cin == "OFF": offs.append(DAYS[i][:3])
            if hrs and hrs != "-": s += hm_to_min(hrs)
        total = None
        for row in ws.iter_rows():
            for c in row:
                if c.value == "TOTAL HOURS":
                    total = ws.cell(row=c.row, column=5).value
        got = min_to_hm(s)
        exp_off = sorted(daysoff.get(ws.title, []))
        off_ok = (exp_off == sorted(ABBR[o.lower()] for o in offs)) if ws.title in daysoff else True
        status = "OK" if (got == total and off_ok) else "MISMATCH"
        if status != "OK": ok = False
        print(f"{ws.title:<26} {got:>6} {str(total):>6} {','.join(offs):>10}  {status}")
    return ok

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--daysoff", help="JSON file: {alias: [weekday abbrevs]}")
    ap.add_argument("--schedule", help="JSON file of real schedules; overrides the "
                    "built-in one. Pass 'none' to disable schedules entirely.")
    ap.add_argument("--seed", type=int, default=None,
                    help="Fix the randomization. Default: derived from the pay-period "
                    "week so each week differs but a given week stays consistent.")
    a = ap.parse_args()
    emps, week_start = parse_pdf(a.pdf)
    if not emps:
        print("No employees found in PDF.", file=sys.stderr); sys.exit(1)
    print(f"Found {len(emps)} employees; week starting {week_start}")
    daysoff = {}
    if a.daysoff:
        spec = json.load(open(a.daysoff))
        daysoff, report = resolve_daysoff(emps, spec)
        print("Days-off mapping:")
        for alias, name, r in report:
            tag = "by-order" if r == 0.0 else f"{r:.2f}"
            print(f"  {alias:<18} -> {name:<26} ({tag}) off={[DAYS[i][:3] for i in daysoff[name]]}")
    spec = DEFAULT_SCHEDULE
    if a.schedule:
        spec = {} if a.schedule.lower() == "none" else json.load(open(a.schedule))
    schedule = resolve_schedule(emps, spec)
    if schedule:
        print("Using real schedules for:", ", ".join(sorted(schedule)))
    seed = a.seed
    if seed is None:
        seed = int(week_start.strftime("%Y%m%d")) if week_start else random.randrange(1 << 30)
    build(emps, week_start, a.out, daysoff, seed, schedule)
    if not verify(a.out, daysoff):
        print("Verification failed.", file=sys.stderr); sys.exit(2)
    print(f"Saved {a.out}")

if __name__ == "__main__":
    main()
