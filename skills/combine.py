#!/usr/bin/env python3
"""Combine a paystubs PDF and a time-cards PDF into one PDF.

Each employee's paystub is placed on the top half of a US-Letter page and that
same employee's time card on the bottom half, one employee per page. Employees
are matched by name (fuzzy, tolerating spelling differences) and confirmed with
total hours. Each source page is cropped to its content bounding box; the
isolated footer line at the bottom of each paystub is removed.

Usage:
    python3 combine.py --paystubs paystubs.pdf --timecards timecards.pdf --out combined.pdf
"""
import argparse
import difflib
import re

import pdfplumber
from pypdf import PdfReader, PdfWriter
from pypdf.generic import RectangleObject

PAGE_W, PAGE_H = 612.0, 792.0
MARGIN = 18.0
GAP = 12.0


def norm(s):
    return re.sub(r"[^a-z]", "", (s or "").lower())


def hm_to_min(s):
    try:
        h, m = s.split(":")
        return int(h) * 60 + int(m)
    except Exception:
        return None


def content_bbox(page, keys=("chars", "lines", "rects", "curves")):
    xs0, xs1, tops, bottoms = [], [], [], []
    for key in keys:
        for o in getattr(page, key, []) or []:
            xs0.append(o["x0"]); xs1.append(o["x1"])
            tops.append(o["top"]); bottoms.append(o["bottom"])
    if not xs0:
        return (0, 0, float(page.width), float(page.height))
    return (min(xs0), min(tops), max(xs1), max(bottoms))


def text_lines(page):
    rows = {}
    for c in page.chars or []:
        key = round(c["top"] / 2.0)
        t, b = rows.get(key, (c["top"], c["bottom"]))
        rows[key] = (min(t, c["top"]), max(b, c["bottom"]))
    return sorted(rows.values())


def strip_footer(page, bbox):
    lines = text_lines(page)
    if len(lines) < 3:
        return bbox
    gaps = [(lines[i + 1][0] - lines[i][1], i) for i in range(len(lines) - 1)]
    med = sorted(g for g, _ in gaps)[len(gaps) // 2]
    # a footer is a low line preceded by a large gap: judge by where the NEXT line sits
    lower_gaps = [(g, i) for g, i in gaps if lines[i + 1][0] > page.height * 0.55]
    if not lower_gaps:
        return bbox
    g, i = max(lower_gaps, key=lambda t: t[0])
    if g > max(22.0, 3.0 * max(med, 1.0)):
        new_bottom = lines[i][1] + 2
        x0, top, x1, _ = bbox
        return (x0, top, x1, min(bbox[3], new_bottom))
    return bbox


def parse_pdf(path, is_paystub):
    out = []
    with pdfplumber.open(path) as pdf:
        for idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            name = None
            hours = None
            if is_paystub:
                m = re.search(r"^([A-Z][A-Z .'-]+?),.*\*\*\*-\*\*-", text, re.M)
                if m:
                    name = m.group(1).strip().title()
                else:
                    # Cash stubs: no SSN on the name line; the name is the line
                    # right after the "Employee  SSN" header.
                    lines = text.splitlines()
                    for li, line in enumerate(lines[:-1]):
                        if re.match(r"^\s*Employee\s+SSN\b", line):
                            nm = re.sub(r"\*{3}-\*{2}-\d+", "", lines[li + 1]).split(",")[0]
                            nm = re.sub(r"[\s.]+$", "", nm).strip()
                            if nm:
                                name = nm.title() if nm.isupper() else nm
                            break
                h = re.search(r"Hourly\s+(\d{1,3}:\d{2})", text)
                if h:
                    hours = hm_to_min(h.group(1))
            else:
                m = (re.search(r"Employee:?\s*([A-Za-z][A-Za-z .'-]+?)(?:\s+Week\b|\s{2,}|$)", text, re.M)
                     or re.search(r"^([A-Z][A-Za-z .'-]{2,})$", text, re.M))
                if m:
                    name = re.sub(r"\s+", " ", m.group(1)).strip().title()
                tm = re.search(r"TOTAL[^\n]*?(\d{1,3}:\d{2})", text, re.I)
                if tm:
                    hours = hm_to_min(tm.group(1))
                else:
                    cands = [hm_to_min(x) for x in re.findall(r"\d{1,3}:\d{2}", text)]
                    cands = [c for c in cands if c is not None]
                    if cands:
                        hours = max(cands)
            bbox = content_bbox(page)
            if is_paystub:
                bbox = strip_footer(page, bbox)
            x0, top, x1, bottom = bbox
            ph = float(page.height)
            yb, yt = ph - bottom, ph - top
            out.append({
                "page": idx,
                "name": name or f"(page {idx + 1})",
                "hours_min": hours,
                "bbox_native": (x0, yb, x1, yt),
                "w": max(1.0, x1 - x0),
                "h": max(1.0, yt - yb),
            })
    return out


def match(paystubs, timecards):
    scored = []
    for pi, ps in enumerate(paystubs):
        for ti, tc in enumerate(timecards):
            r = difflib.SequenceMatcher(None, norm(ps["name"]), norm(tc["name"])).ratio()
            na, nb = norm(ps["name"]), norm(tc["name"])
            if na and nb and (na in nb or nb in na or nb.startswith(na) or na.startswith(nb)):
                r = max(r, 0.9)
            hb = 0.0
            if ps["hours_min"] and tc["hours_min"]:
                diff = abs(ps["hours_min"] - tc["hours_min"])
                hb = 0.25 if diff == 0 else (0.12 if diff <= 60 else 0.0)
            scored.append((r + hb, r, pi, ti))
    scored.sort(reverse=True)
    used_p, used_t = set(), set()
    pairs, report = [], []
    for tot, r, pi, ti in scored:
        if pi in used_p or ti in used_t or r < 0.4:
            continue
        used_p.add(pi); used_t.add(ti)
        pairs.append((paystubs[pi], timecards[ti]))
        report.append((paystubs[pi]["name"], timecards[ti]["name"], r,
                       paystubs[pi]["hours_min"], timecards[ti]["hours_min"]))
    left_p = [p for i, p in enumerate(paystubs) if i not in used_p]
    left_t = [t for i, t in enumerate(timecards) if i not in used_t]
    for ps, tc in zip(left_p, left_t):
        pairs.append((ps, tc))
        report.append((ps["name"], tc["name"], 0.0, ps["hours_min"], tc["hours_min"]))
    return pairs, report


def place(dest, reader, item, y_bottom, scale):
    src = reader.pages[item["page"]]
    x0, yb, x1, yt = item["bbox_native"]
    # pypdf clips the merged page to its crop box (older versions: trim box),
    # so shrink both to the content box -- otherwise anything outside it
    # (e.g. the paystub footer line) bleeds onto the time card.
    box = RectangleObject((x0 - 1, yb - 1, x1 + 1, yt + 1))
    src.cropbox = box
    src.trimbox = box
    tx = MARGIN
    ctm = (scale, 0, 0, scale, tx - x0 * scale, y_bottom - yb * scale)
    dest.merge_transformed_page(src, ctm, expand=False)


def combine(paystubs_pdf, timecards_pdf, out_pdf):
    ps_items = parse_pdf(paystubs_pdf, True)
    tc_items = parse_pdf(timecards_pdf, False)
    if not ps_items:
        raise SystemExit("No pages found in paystubs PDF.")
    if not tc_items:
        raise SystemExit("No pages found in timecards PDF.")
    pairs, report = match(ps_items, tc_items)
    ps_reader = PdfReader(paystubs_pdf)
    tc_reader = PdfReader(timecards_pdf)
    writer = PdfWriter()
    content_w = PAGE_W - 2 * MARGIN
    avail_h = PAGE_H - 2 * MARGIN - GAP
    for ps, tc in pairs:
        page = writer.add_blank_page(width=PAGE_W, height=PAGE_H)
        s_ps = content_w / ps["w"]
        s_tc = content_w / tc["w"]
        nh_ps = ps["h"] * s_ps
        nh_tc = tc["h"] * s_tc
        total = nh_ps + nh_tc
        if total > avail_h:
            f = avail_h / total
            s_ps *= f; s_tc *= f
            nh_ps *= f; nh_tc *= f
        ps_bottom = (PAGE_H - MARGIN) - nh_ps
        place(page, ps_reader, ps, ps_bottom, s_ps)
        tc_bottom = ps_bottom - GAP - nh_tc
        place(page, tc_reader, tc, tc_bottom, s_tc)
    with open(out_pdf, "wb") as fh:
        writer.write(fh)
    return out_pdf, report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--paystubs", required=True)
    ap.add_argument("--timecards", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    out, report = combine(a.paystubs, a.timecards, a.out)
    print(f"{'Paystub':<24} {'Timecard':<24} {'score':>6} {'ps_hrs':>7} {'tc_hrs':>7}")
    for pn, tn, r, ph, th in report:
        tag = "by-order" if r == 0.0 else f"{r:.2f}"
        print(f"{pn:<24} {tn:<24} {tag:>6} {str(ph):>7} {str(th):>7}")
    print(f"Saved {out} ({len(report)} pages)")


if __name__ == "__main__":
    main()
