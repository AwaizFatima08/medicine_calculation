"""
Build a fresh Comparative Statement (CST) from the vendor RFQ quote files.

Usage:   python3 build_cst.py
Inputs:  quotes/*.xls                       vendor-filled RFQ sheets (sheet "RFQ MEDICINES")
         reference/Previous_CST_*.xls       the earlier hand-made CST; used only for
                                            (a) MCOLSON and SAMI quotes, which have no file,
                                            (b) pack sizes it recorded, (c) its awards, for review
Output:  CST_RFQ_Medicines.xlsx

Vendor file layout (all identical):
  A SR# | B Item Code | C Description | D UOM | E QTY | F Price | G Total | H Quoted Brand | I Remarks
"""
import glob
import os
import re
import shutil
import subprocess
import tempfile
from collections import OrderedDict

import xlrd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "CST_RFQ_Medicines.xlsx")

# Vendor groups in CST column order (same order as the previous CST).
GROUPS = ["ABBOTT (Jamil Traders)", "GSK (Jamil Traders)", "HIGHNOON", "MCOLSON",
          "OBS GROUP (M&P)", "HORIZON (Aslam Traders)", "SANOFI (Aslam Traders)", "SAMI"]
ABBOTT, GSK, HIGHNOON, MCOLSON, OBS, HORIZON, SANOFI, SAMI = GROUPS
# Column index of each group in the previous CST (price column, 0-based)
PREV_COLS = OrderedDict(zip(GROUPS, range(5, 37, 4)))

# Jamil Traders quotes Abbott and GSK products in one file; GSK brands are:
GSK_BRANDS = ("CAL POL", "CALPOL", "PANADOL", "CUTIVATE", "FIXVAL", "BETNOVATE", "POLYFAX",
              "AUGM", "BETNESOL", "CLOBEVATE", "REVITALE", "FEFOLVIT")

# Quotes kept visible but excluded from price comparison (spec mismatch).
EXCLUDE = {
    (HORIZON, 59): "WRONG ITEM: Estadol 500 mg quoted against Propranolol 10 mg",
}

# Review flags shown in the CST row (details in 'Review Notes').
FLAGS = {
    1: "Horizon 15 mg (4,250) is below its own 2.5 mg (7,650); likely typo. Previous CST awarded on this figure. Confirm in writing.",
    2: "Horizon tirzepatide prices run opposite to strength (see SR 1).",
    3: "Horizon tirzepatide prices run opposite to strength (see SR 1).",
    4: "Horizon tirzepatide prices run opposite to strength (see SR 1).",
    5: "SAMI price has 3 decimals (discounted figure); confirm rounding.",
    16: "OBS brand reads 'CINKLARE CAP' for a syrup; verify.",
    22: "OBS and SAMI prices carry 3-4 decimals; confirm final unit price.",
    23: "Pack sizes unknown for Abbott, OBS, Horizon, SAMI; Highnoon is 14's. Compare per tablet before awarding.",
    30: "OBS quoted Canesten (clotrimazole only); item needs hydrocortisone + clotrimazole. Previous CST awarded it. Check spec.",
    37: "SAMI price 168.24135 is a computed discount; confirm.",
    38: "SAMI price 212.89185 is a computed discount; confirm.",
    44: "Highnoon note: Kestine pack changed from 14 to 20 tablets. Pack sizes of OBS/SAMI unknown.",
    46: "Pack sizes differ (20/14/14/30). Per-tablet L1 is Highnoon, not the lowest pack price (MCOLSON).",
    59: "Horizon's Estadol is the wrong item; excluded. No valid quote.",
    65: "Only GSK quoted (Calpol / Panadol, both 617.96). Previous CST left it unawarded.",
    76: "Pack sizes differ (30/45/30/30). Abbott Surbex Z is L1 per pack and per tablet.",
    77: "Pack sizes differ (20/10/30/10). Per-tablet L1 is Highnoon; SAMI has the lowest pack price.",
    84: "OBS price has 3 decimals; confirm.",
    85: "Pack sizes differ (20/30/20). Per-tablet L1 is Horizon; SAMI has the lowest pack price.",
    103: "Horizon quoted FERRIUM 120 ML against a 60 ML item.",
    111: "Sanofi Lantus is a pack of 5 pens; OBS Basagine is single. Per pen: Sanofi 1,467.35 vs OBS 1,134.75.",
    118: "OBS price has 3 decimals; confirm.",
    120: "Pack sizes differ (30 vs 56). Per-capsule L1 is GSK Fefolvit; Abbott Iberet has the lower pack price.",
    121: "Pack sizes differ (30 vs 20). Per-capsule L1 is Horizon Mucogin; Sanofi has the lower pack price.",
    123: "OBS quoted a chewable tablet against a capsule item. SAMI is L1 per pack and per capsule.",
    124: "Abbott quoted a tablet against a capsule item. Abbott is L1 per pack and per capsule.",
}


def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def read_vendor_file(path):
    """items: SR -> dict(code, desc, uom, qty). quotes: list of dict(sr, price, brand, remarks, row)."""
    sh = xlrd.open_workbook(path).sheet_by_name("RFQ MEDICINES")
    items, quotes = {}, []
    for r in range(2, sh.nrows):
        sr = sh.cell_value(r, 0)
        if not isinstance(sr, float) or sr == 0:
            continue
        sr = int(sr)
        items.setdefault(sr, dict(code=clean(sh.cell_value(r, 1)), desc=clean(sh.cell_value(r, 2)),
                                  uom=clean(sh.cell_value(r, 3)), qty=sh.cell_value(r, 4)))
        price = sh.cell_value(r, 5)
        if isinstance(price, float) and price > 0:
            quotes.append(dict(sr=sr, price=price, brand=clean(sh.cell_value(r, 7)),
                               remarks=clean(sh.cell_value(r, 8)), row=r + 1))
    return items, quotes


def read_previous_cst(path):
    """Returns (quotes, packs, awards).
    quotes: {(group, sr): (price, brand, remarks)} for every priced cell
    packs:  {(group, sr): n} from the 'PACK OF n' / 'SINGLE' sub-rows
    awards: {sr: group} where the price cell is filled yellow (colour index 13)"""
    book = xlrd.open_workbook(path, formatting_info=True)
    sh = book.sheet_by_index(0)
    quotes, packs, awards = {}, {}, {}
    sr = None
    for r in range(3, sh.nrows):
        a = sh.cell_value(r, 0)
        if isinstance(a, float):
            sr = int(a)
            for g, c in PREV_COLS.items():
                p = sh.cell_value(r, c)
                if isinstance(p, float) and p > 0:
                    quotes[(g, sr)] = (p, clean(sh.cell_value(r, c + 2)), clean(sh.cell_value(r, c + 3)))
                    xf = book.xf_list[sh.cell_xf_index(r, c)]
                    if xf.background.fill_pattern and xf.background.pattern_colour_index == 13:
                        awards[sr] = g
        elif sr is not None:
            for g, c in PREV_COLS.items():
                txt = clean(sh.cell_value(r, c + 2)).upper()
                m = re.search(r"PACK OF\s*,?\s*(\d+)", txt)
                if m:
                    packs[(g, sr)] = int(m.group(1))
                elif txt == "SINGLE":
                    packs[(g, sr)] = 1
    return quotes, packs, awards


def split_jamil(q):
    """Jamil Traders sometimes lists two brands: first brand priced in Price, second in Remarks
    as 'BRAND 123.45'. Return a list of (group, price, brand) quotes."""
    brands = [b.strip() for b in q["brand"].split(",") if b.strip()]
    grp = lambda b: GSK if any(k in b.upper() for k in GSK_BRANDS) else ABBOTT
    m = re.match(r"(.+?)\s+([\d.]+)$", q["remarks"])
    if not m:
        return [(grp(q["brand"]), q["price"], q["brand"])]
    first, alt_brand, alt_price = brands[0], m.group(1).strip(), float(m.group(2))
    out = [(grp(first), q["price"], first), (grp(alt_brand), alt_price, alt_brand)]
    if out[0][0] == out[1][0] and abs(out[0][1] - out[1][1]) < 1e-9:   # SR 65: Calpol / Panadol, same price
        return [(out[0][0], q["price"], f"{first} / {alt_brand} (both {q['price']:,.2f})")]
    return out


def pack_from_brand(brand):
    m = re.search(r"(\d+)\s*'S\b", brand.upper())
    return int(m.group(1)) if m else None


def main():
    quotes_dir = os.path.join(HERE, "quotes")
    prev_path = (glob.glob(os.path.join(HERE, "reference", "Previous_CST_*.xls")) or [None])[0]

    master = None
    grid = {g: {} for g in GROUPS}      # group -> sr -> dict(price, brand, pack, src)
    raw = []                            # audit trail rows
    prev_quotes, prev_packs, prev_awards = ({}, {}, {})
    if prev_path:
        prev_quotes, prev_packs, prev_awards = read_previous_cst(prev_path)

    def put(group, sr, price, brand, src, row):
        if sr in grid[group]:
            raise SystemExit(f"Duplicate quote for {group} SR {sr}: {grid[group][sr]} and {price} {brand}")
        grid[group][sr] = dict(price=price, brand=brand, src=src)
        raw.append((src, row, group, sr, price, brand))

    file_map = [("Highnoon.xls", lambda q: [(HIGHNOON, q["price"], q["brand"])]),
                ("Hoechst.xls", lambda q: [(SANOFI, q["price"], q["brand"])]),
                ("Horizon.xls", lambda q: [((SANOFI if q["remarks"].lower() == "sanofi" else HORIZON), q["price"], q["brand"])]),
                ("Jamil_Traders.xls", split_jamil),
                ("MP.xls", lambda q: [(OBS, q["price"], q["brand"])])]
    for fname, splitter in file_map:
        items, quotes = read_vendor_file(os.path.join(quotes_dir, fname))
        if master is None:
            master = items
        elif items != master:
            bad = [sr for sr in items if items[sr] != master.get(sr)]
            raise SystemExit(f"{fname}: item list differs from first file at SR {bad}")
        for q in quotes:
            for group, price, brand in splitter(q):
                if group == SANOFI and q["remarks"].lower() == "sanofi" and q["sr"] in grid[SANOFI]:
                    # Horizon file repeats the Sanofi (Hoechst) quote; must be identical
                    if abs(grid[SANOFI][q["sr"]]["price"] - price) > 1e-9:
                        raise SystemExit(f"Sanofi price conflict at SR {q['sr']}")
                    raw.append((fname, q["row"], SANOFI, q["sr"], price, brand + " (repeat of Hoechst quote)"))
                    continue
                put(group, q["sr"], price, brand, fname, q["row"])

    # MCOLSON and SAMI exist only in the previous CST
    for (g, sr), (price, brand, rem) in prev_quotes.items():
        if g in (MCOLSON, SAMI):
            put(g, sr, price, brand, "Previous CST", "")

    # Cross-check: every previous-CST figure for file-backed groups must match the files
    mismatches = []
    for (g, sr), (price, brand, rem) in prev_quotes.items():
        if g in (MCOLSON, SAMI):
            continue
        mine = grid[g].get(sr)
        if mine is None or abs(mine["price"] - price) > 1e-6:
            mismatches.append((g, sr, price, mine and mine["price"]))
    for g in GROUPS:
        if g in (MCOLSON, SAMI):
            continue
        for sr in grid[g]:
            if (g, sr) not in prev_quotes:
                mismatches.append((g, sr, None, grid[g][sr]["price"]))
    print("Previous CST vs vendor files:", "all figures match" if not mismatches else mismatches)

    # pack sizes: previous CST sub-rows first, then Highnoon's "14'S" style brand text
    for g in GROUPS:
        for sr, q in grid[g].items():
            q["pack"] = prev_packs.get((g, sr)) or pack_from_brand(q["brand"])

    srs = sorted(master)
    ng = len(GROUPS)

    # ---------- styles ----------
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_font = Font(bold=True, color="FFFFFF")
    vend_fills = ["CCFFCC", "CCCCFF", "99CCFF", "D9D9D9", "FFFFCC", "F8CBAD", "C6E0B4", "FFE699"]
    sum_fill = PatternFill("solid", fgColor="D9D9D9")
    green = PatternFill("solid", fgColor="92D050")
    blue = PatternFill("solid", fgColor="9BC2E6")
    red = PatternFill("solid", fgColor="F4B6B6")
    grey_font = Font(color="808080", italic=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap = Alignment(vertical="top", wrap_text=True)
    money = '#,##0.00'

    wb = Workbook()

    # ================= Sheet 1: CST =================
    ws = wb.active
    ws.title = "CST"
    FIRST, LAST = 6, 5 + len(srs)
    TOT = LAST + 1
    PER = ["Price", "Pack", "Unit Price", "Total", "Quoted Brand / Remarks"]
    fixed = ["SR#", "Item Code", "Description", "UOM", "Qty"]

    ws.cell(1, 1, "COMPARATIVE STATEMENT (CST) - REQUEST FOR QUOTATION: MEDICINES").font = Font(bold=True, size=14)
    ws.cell(2, 1, "Price = vendor's quoted price per UOM. Pack = units per pack where stated (previous CST sub-rows or vendor text); "
                  "for UOM EACH with no pack stated, 1 is assumed. Unit Price = Price / Pack.").font = Font(italic=True, color="555555")
    ws.cell(3, 1, "Green = lowest Price (L1 per pack).  Blue = lowest Unit Price (L1 per tablet/pen/bottle).  "
                  "Pink rows = no valid quote.  Grey italics = quote excluded (see brand cell).  Read 'Review Notes' before awarding.").font = Font(italic=True, color="555555")

    for i, h in enumerate(fixed, 1):
        ws.merge_cells(start_row=4, start_column=i, end_row=5, end_column=i)
        c = ws.cell(4, i, h)
        c.font, c.fill, c.alignment, c.border = hdr_font, hdr_fill, center, border
        ws.cell(5, i).border = border

    c0 = len(fixed) + 1
    gcol = {}   # group -> dict of column indexes
    for gi, g in enumerate(GROUPS):
        base = c0 + len(PER) * gi
        gcol[g] = dict(price=base, pack=base + 1, unit=base + 2, total=base + 3, brand=base + 4)
        ws.merge_cells(start_row=4, start_column=base, end_row=4, end_column=base + len(PER) - 1)
        c = ws.cell(4, base, g)
        c.font, c.fill, c.alignment, c.border = Font(bold=True), PatternFill("solid", fgColor=vend_fills[gi]), center, border
        for k, h in enumerate(PER):
            hc = ws.cell(5, base + k, h)
            hc.font, hc.fill, hc.alignment, hc.border = hdr_font, hdr_fill, center, border
            ws.cell(4, base + k).border = border

    s0 = c0 + len(PER) * ng
    summary = ["No. of Quotes", "L1 Price (per pack)", "L1 Vendor (per pack)", "L1 Total (Qty x L1 Price)",
               "L1 Unit Price", "L1 Vendor (per unit)", "Quotes with Pack Unknown", "Review Flag"]
    for i, h in enumerate(summary):
        ws.merge_cells(start_row=4, start_column=s0 + i, end_row=5, end_column=s0 + i)
        c = ws.cell(4, s0 + i, h)
        c.font, c.fill, c.alignment, c.border = hdr_font, hdr_fill, center, border
        ws.cell(5, s0 + i).border = border
    C_N, C_MIN, C_VEND, C_L1TOT, C_UMIN, C_UVEND, C_NOPACK, C_FLAG = (col(s0 + i) for i in range(8))
    P = [col(gcol[g]["price"]) for g in GROUPS]
    U = [col(gcol[g]["unit"]) for g in GROUPS]
    LASTCOL = s0 + len(summary) - 1

    for i, sr in enumerate(srs):
        r = FIRST + i
        it = master[sr]
        ws.cell(r, 1, sr)
        ws.cell(r, 2, it["code"])
        ws.cell(r, 3, it["desc"]).alignment = wrap
        ws.cell(r, 4, it["uom"])
        ws.cell(r, 5, it["qty"]).number_format = '#,##0'
        for cc in (1, 2, 4, 5):
            ws.cell(r, cc).alignment = Alignment(vertical="top")
        for g in GROUPS:
            cc = gcol[g]
            q = grid[g].get(sr)
            pc, kc, uc, tc, bc = (col(cc[k]) for k in ("price", "pack", "unit", "total", "brand"))
            if q:
                excl = EXCLUDE.get((g, sr))
                if excl:
                    ws.cell(r, cc["brand"], f"{q['brand']} @ {q['price']:,.2f} - {excl} - EXCLUDED").font = grey_font
                else:
                    ws.cell(r, cc["price"], q["price"]).number_format = money
                    if q["pack"]:
                        ws.cell(r, cc["pack"], q["pack"])
                    ws.cell(r, cc["brand"], q["brand"])
                ws.cell(r, cc["brand"]).alignment = wrap
            ws.cell(r, cc["unit"], f'=IF({pc}{r}="","",IF({kc}{r}<>"",{pc}{r}/{kc}{r},IF($D{r}="EACH",{pc}{r},"")))').number_format = money
            ws.cell(r, cc["total"], f'=IF({pc}{r}="","",$E{r}*{pc}{r})').number_format = money
        plist = ",".join(f"{p}{r}" for p in P)
        ulist = ",".join(f"{u}{r}" for u in U)
        ws.cell(r, s0, f"=COUNT({plist})")
        ws.cell(r, s0 + 1, f'=IF({C_N}{r}=0,"",MIN({plist}))').number_format = money
        expr = '""'
        for p, name in reversed(list(zip(P, GROUPS))):
            expr = f'IF({p}{r}={C_MIN}{r},"{name}",{expr})'
        ws.cell(r, s0 + 2, f'=IF({C_N}{r}=0,"",{expr})')
        ws.cell(r, s0 + 3, f'=IF({C_N}{r}=0,"",$E{r}*{C_MIN}{r})').number_format = money
        ws.cell(r, s0 + 4, f'=IF(COUNT({ulist})=0,"",MIN({ulist}))').number_format = money
        uexpr = '""'
        for u, name in reversed(list(zip(U, GROUPS))):
            uexpr = f'IF({u}{r}={C_UMIN}{r},"{name}",{uexpr})'
        ws.cell(r, s0 + 5, f'=IF(COUNT({ulist})=0,"",{uexpr})')
        ws.cell(r, s0 + 6, f'=IF({C_N}{r}=0,"",{C_N}{r}-COUNT({ulist}))')
        ws.cell(r, s0 + 7, FLAGS.get(sr, "")).alignment = wrap
        for cc in range(1, LASTCOL + 1):
            ws.cell(r, cc).border = border

    ws.cell(TOT, 3, "TOTAL (quoted items only)").font = Font(bold=True)
    ws.cell(TOT, 5, f"=SUM(E{FIRST}:E{LAST})").number_format = '#,##0'
    for g in GROUPS:
        pc, tc = gcol[g]["price"], gcol[g]["total"]
        ws.cell(TOT, pc, f'=COUNT({col(pc)}{FIRST}:{col(pc)}{LAST})&" items"').alignment = center
        ws.cell(TOT, tc, f"=SUM({col(tc)}{FIRST}:{col(tc)}{LAST})").number_format = money
    ws.cell(TOT, s0, f'=COUNTIF({C_N}{FIRST}:{C_N}{LAST},">0")&" of {len(srs)}"')
    ws.cell(TOT, s0 + 3, f"=SUM({C_L1TOT}{FIRST}:{C_L1TOT}{LAST})").number_format = money
    for cc in range(1, LASTCOL + 1):
        c = ws.cell(TOT, cc)
        c.fill, c.border, c.font = sum_fill, border, Font(bold=True)

    for p in P:
        ws.conditional_formatting.add(f"{p}{FIRST}:{p}{LAST}",
            FormulaRule(formula=[f'AND({p}{FIRST}<>"",{p}{FIRST}=${C_MIN}{FIRST})'], fill=green))
    for u in U:
        ws.conditional_formatting.add(f"{u}{FIRST}:{u}{LAST}",
            FormulaRule(formula=[f'AND({u}{FIRST}<>"",{u}{FIRST}=${C_UMIN}{FIRST})'], fill=blue))
    ws.conditional_formatting.add(f"A{FIRST}:E{LAST}", FormulaRule(formula=[f"${C_N}{FIRST}=0"], fill=red))

    for cc, w in {1: 5, 2: 23, 3: 44, 4: 8, 5: 7}.items():
        ws.column_dimensions[col(cc)].width = w
    for g in GROUPS:
        cc = gcol[g]
        for k, w in (("price", 10), ("pack", 6), ("unit", 9), ("total", 12), ("brand", 28)):
            ws.column_dimensions[col(cc[k])].width = w
    for i, w in enumerate((7, 11, 18, 14, 10, 18, 9, 50)):
        ws.column_dimensions[col(s0 + i)].width = w
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 32
    ws.freeze_panes = ws.cell(FIRST, c0)
    ws.auto_filter.ref = f"A5:{col(LASTCOL)}{LAST}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "4:5"

    # ================= Sheet 2: Vendor Summary =================
    vs = wb.create_sheet("Vendor Summary")
    vs.cell(1, 1, "VENDOR SUMMARY").font = Font(bold=True, size=14)
    heads = ["Vendor", "Items Quoted", f"Items Not Quoted (of {len(srs)})", "Value of Quoted Items (Qty x Price)",
             "Items L1 per pack", "Value where L1 per pack", "Share", "Items L1 per unit", "Awards in previous CST"]
    for i, h in enumerate(heads, 1):
        c = vs.cell(3, i, h)
        c.font, c.fill, c.alignment, c.border = hdr_font, hdr_fill, center, border
    tot = 4 + ng
    for gi, g in enumerate(GROUPS):
        r = 4 + gi
        pc, tc = col(gcol[g]["price"]), col(gcol[g]["total"])
        vs.cell(r, 1, g)
        vs.cell(r, 2, f"=COUNT(CST!{pc}{FIRST}:{pc}{LAST})")
        vs.cell(r, 3, f"={len(srs)}-B{r}")
        vs.cell(r, 4, f"=SUM(CST!{tc}{FIRST}:{tc}{LAST})").number_format = money
        vs.cell(r, 5, f'=COUNTIF(CST!{C_VEND}{FIRST}:{C_VEND}{LAST},A{r})')
        vs.cell(r, 6, f'=SUMIF(CST!{C_VEND}{FIRST}:{C_VEND}{LAST},A{r},CST!{C_L1TOT}{FIRST}:{C_L1TOT}{LAST})').number_format = money
        vs.cell(r, 7, f"=IF($F${tot}=0,0,F{r}/$F${tot})").number_format = '0.0%'
        vs.cell(r, 8, f'=COUNTIF(CST!{C_UVEND}{FIRST}:{C_UVEND}{LAST},A{r})')
        vs.cell(r, 9, sum(1 for v in prev_awards.values() if v == g))
        for cc in range(1, 10):
            vs.cell(r, cc).border = border
    vs.cell(tot, 1, "TOTAL").font = Font(bold=True)
    for cc, f in ((2, None), (5, f"=SUM(E4:E{tot-1})"), (6, f"=SUM(F4:F{tot-1})"), (7, f"=SUM(G4:G{tot-1})"),
                  (8, f"=SUM(H4:H{tot-1})"), (9, f"=SUM(I4:I{tot-1})")):
        if f:
            vs.cell(tot, cc, f)
    vs.cell(tot, 6).number_format = money
    vs.cell(tot, 7).number_format = '0.0%'
    for cc in range(1, 10):
        vs.cell(tot, cc).fill, vs.cell(tot, cc).border = sum_fill, border
    vs.cell(tot + 2, 1, "Items with at least one valid quote:")
    vs.cell(tot + 2, 4, f'=COUNTIF(CST!{C_N}{FIRST}:{C_N}{LAST},">0")')
    vs.cell(tot + 3, 1, "Items with no valid quote:")
    vs.cell(tot + 3, 4, f'=COUNTIF(CST!{C_N}{FIRST}:{C_N}{LAST},0)')
    vs.cell(tot + 4, 1, "Items with a single quote (no competition):")
    vs.cell(tot + 4, 4, f'=COUNTIF(CST!{C_N}{FIRST}:{C_N}{LAST},1)')
    vs.cell(tot + 5, 1, "Sum of L1 totals, per-pack basis:")
    vs.cell(tot + 5, 4, f"=SUM(CST!{C_L1TOT}{FIRST}:{C_L1TOT}{LAST})").number_format = money
    vs.cell(tot + 7, 1, "Ties go to the left-most vendor in the CST. 'Value where L1' counts each item once at the lowest pack price. "
                        "Per-unit counts only cover quotes with a known pack size.").font = Font(italic=True, color="555555")
    for cc, w in zip(range(1, 10), (26, 12, 20, 24, 14, 22, 9, 14, 20)):
        vs.column_dimensions[col(cc)].width = w

    # ================= Sheet 3: Review vs Previous CST (static snapshot) =================
    rv = wb.create_sheet("Review vs Previous CST")
    rv.cell(1, 1, "REVIEW: previous CST awards (yellow cells) compared with this CST's L1").font = Font(bold=True, size=14)
    rv.cell(2, 1, "Static snapshot computed when this file was built. Live figures are on the CST sheet.").font = Font(italic=True, color="555555")
    heads = ["SR#", "Description", "Qty", "Previous award: vendor", "Previous award: price",
             "New L1 per pack: vendor", "New L1 per pack: price", "New L1 per unit: vendor", "New L1 per unit: price", "Pack sizes known?", "Result"]
    for i, h in enumerate(heads, 1):
        c = rv.cell(4, i, h)
        c.font, c.fill, c.alignment, c.border = hdr_font, hdr_fill, center, border
    diff_count = 0
    r = 5
    for sr in srs:
        valid = {g: grid[g][sr] for g in GROUPS if sr in grid[g] and (g, sr) not in EXCLUDE}
        prev_g = prev_awards.get(sr)
        if not valid and not prev_g:
            continue
        it = master[sr]
        l1g = min(valid, key=lambda g: (valid[g]["price"], GROUPS.index(g))) if valid else None
        units = {g: q["price"] / (q["pack"] or (1 if it["uom"] == "EACH" else 0)) for g, q in valid.items() if q["pack"] or it["uom"] == "EACH"}
        u1g = min(units, key=lambda g: (units[g], GROUPS.index(g))) if units else None
        packs_known = "all" if len(units) == len(valid) and valid else ("none" if not units else f"{len(units)} of {len(valid)}")
        prev_price = prev_quotes.get((prev_g, sr), (None,))[0] if prev_g else None
        if not valid:
            result = "No valid quote (previous CST also unawarded)" if not prev_g else "Previous award now excluded"
        elif prev_g is None:
            result = "Previous CST left this unawarded"
        elif prev_g == l1g and (u1g in (None, prev_g)):
            result = "Same"
        elif prev_g == u1g:
            result = "Same (previous CST chose by unit price)"
        elif prev_g == l1g and packs_known != "all":
            result = "Same per pack (unit comparison incomplete: pack sizes missing)"
        elif prev_g == l1g:
            result = "Same per pack; differs per unit"
        else:
            result = "DIFFERS"
        if result != "Same":
            diff_count += 1
        vals = [sr, it["desc"], it["qty"], prev_g or "", prev_price, l1g or "", valid[l1g]["price"] if l1g else None,
                u1g or "", units[u1g] if u1g else None, packs_known, result]
        for cc, v in enumerate(vals, 1):
            c = rv.cell(r, cc, v)
            c.border = border
            c.alignment = wrap
        for cc in (5, 7, 9):
            rv.cell(r, cc).number_format = money
        if result != "Same":
            rv.cell(r, 11).font = Font(bold=True, color="C00000" if ("DIFFERS" in result or "unawarded" in result) else "7F6000")
        r += 1
    rv.auto_filter.ref = f"A4:K{r-1}"
    rv.freeze_panes = "C5"
    for cc, w in zip(range(1, 12), (5, 44, 7, 22, 12, 22, 12, 22, 12, 12, 40)):
        rv.column_dimensions[col(cc)].width = w

    # ================= Sheet 4: Review Notes =================
    quoted_counts = [sum(1 for g in GROUPS if sr in grid[g] and (g, sr) not in EXCLUDE) for sr in srs]
    n_any = sum(1 for n in quoted_counts if n > 0)
    n_single = sum(1 for n in quoted_counts if n == 1)
    n_pack_known = sum(1 for g in GROUPS for sr, q in grid[g].items() if q["pack"])
    n_quotes = sum(len(grid[g]) for g in GROUPS)
    rn = wb.create_sheet("Review Notes")
    rn.cell(1, 1, "REVIEW NOTES - read before awarding").font = Font(bold=True, size=14)
    notes = [
        ("Sources", f"Five vendor files were provided (Highnoon, Hoechst, Horizon, Jamil Traders, MP). MCOLSON and SAMI have no file; their {len(grid[MCOLSON])} + {len(grid[SAMI])} "
                    "quotes were copied from the previous CST and could not be verified against an original quotation. Obtain their original quotes for the file."),
        ("Sources", "Every price in the previous CST for the five file-backed vendors matches the vendor files exactly. Transcription in the previous CST was accurate."),
        ("Structure", "Vendor columns follow the previous CST: Jamil Traders is split into ABBOTT and GSK principals; Aslam Traders is split into HORIZON and SANOFI. "
                      "The 'Hoechst' file is the Sanofi line, and the Horizon file repeats the same five Sanofi rows at identical prices (shown once, under SANOFI)."),
        ("Coverage", f"{len(srs)} items requested. {n_any} have at least one valid quote, {len(srs) - n_any} have none (pink rows), and {n_single} have a single quote, so there is no competition on them."),
        ("Pack size", f"Most prices are per PACKET but only {n_pack_known} of {n_quotes} quotes state a pack size (from the previous CST's 'PACK OF' sub-rows and Highnoon's brand text). "
                      "The previous CST compared per unit only on about a dozen items, by hand. This CST adds Pack / Unit Price columns and an L1-per-unit column so the comparison is explicit, "
                      "but it is only as good as the pack data. Ask every vendor to confirm units per pack for all quoted items; then fill the Pack cells and the sheet recalculates."),
        ("Awards", "On the per-pack basis this CST's L1 agrees with the previous award on almost every item. The sheet 'Review vs Previous CST' lists each item; "
                   f"{diff_count} rows are marked other than 'Same' (unit-price differences, unawarded items, and excluded quotes)."),
        ("SR 1-4", "Horizon tirzepatide: 15 mg at 4,250 but 2.5 mg at 7,650, 5 mg at 10,200, 7.5 mg at 11,900. Price falls as strength rises, which is implausible, "
                   "and the previous CST awarded SR 1 (200 pens, 850,000) to Horizon on the 4,250 figure. Get written confirmation before award; the next price is OBS at 5,525."),
        ("SR 30", "OBS quoted Canesten, which is clotrimazole alone; the item is hydrocortisone + clotrimazole. The previous CST awarded it. Check specification compliance."),
        ("SR 59", "Horizon's 'ESTADOL 500MG' is not propranolol. The previous CST marked it WRONG ITEM; it is excluded here, so the item has no valid quote."),
        ("SR 65", "GSK quoted Calpol and Panadol at the same 617.96. Previous CST left it unawarded although it is the only quote."),
        ("SR 111", "Sanofi Lantus 7,336.77 is a pack of 5 pens (1,467.35 per pen); OBS Basagine 1,134.75 is single. The UOM is EACH, so the per-pen figure is the right comparison."),
        ("Form / volume", "SR 103 Horizon 120 ML vs 60 ML item; SR 123 OBS chewable tablet vs capsule; SR 124 Abbott tablet vs capsule; SR 16 OBS brand text says 'CAP' for a syrup. "
                          "Technical acceptance is outside this price comparison and should be recorded separately."),
        ("Decimals", "OBS (SR 22, 84, 118) and SAMI (SR 5, 17, 21-23, 37, 38, 44, 54, 66, 99, 100, 104, 105, 118, 123, 124) prices carry 3 to 6 decimals, i.e. a discount was applied to a list price. "
                     "Confirm final rounded unit prices with both vendors before issuing orders."),
        ("Method", "L1 Vendor = first vendor from the left whose price equals the row minimum; exact ties go to the left-most column. "
                   "All totals, L1 and the Vendor Summary are live formulas: correcting a price or filling a Pack cell updates everything."),
    ]
    for i, h in enumerate(("#", "Area", "Finding / Action"), 1):
        c = rn.cell(3, i, h)
        c.font, c.fill, c.alignment, c.border = hdr_font, hdr_fill, center, border
    for i, (area, text) in enumerate(notes, 1):
        r = 3 + i
        rn.cell(r, 1, i)
        rn.cell(r, 2, area)
        rn.cell(r, 3, text)
        for cc in range(1, 4):
            rn.cell(r, cc).border = border
            rn.cell(r, cc).alignment = wrap
    rn.column_dimensions["A"].width = 4
    rn.column_dimensions["B"].width = 14
    rn.column_dimensions["C"].width = 130

    # ================= Sheet 5: Source Quotes (audit trail) =================
    sq = wb.create_sheet("Source Quotes")
    sq.cell(1, 1, "SOURCE QUOTES - every priced row as found in the source, before any grouping").font = Font(bold=True, size=14)
    heads = ["Source File", "Source Row", "Vendor Group", "SR#", "Item Code", "Description", "UOM", "Qty", "Price", "Total (Qty x Price)", "Quoted Brand"]
    for i, h in enumerate(heads, 1):
        c = sq.cell(3, i, h)
        c.font, c.fill, c.alignment, c.border = hdr_font, hdr_fill, center, border
    raw.sort(key=lambda x: (x[3], GROUPS.index(x[2]), x[0]))
    for i, (src, row, group, sr, price, brand) in enumerate(raw):
        r = 4 + i
        it = master[sr]
        vals = [src, row, group, sr, it["code"], it["desc"], it["uom"], it["qty"], price, f"=H{r}*I{r}", brand]
        for cc, v in enumerate(vals, 1):
            sq.cell(r, cc, v).border = border
        sq.cell(r, 9).number_format = money
        sq.cell(r, 10).number_format = money
    for cc, w in zip(range(1, 12), (16, 9, 24, 6, 23, 44, 8, 7, 11, 15, 44)):
        sq.column_dimensions[col(cc)].width = w
    sq.freeze_panes = "A4"
    sq.auto_filter.ref = f"A3:K{3 + len(raw)}"

    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    print(f"Wrote {OUT}: {len(srs)} items, {n_quotes} quotes in grid, {len(raw)} source rows, {ng} vendor groups")
    recalc(OUT)


def recalc(path):
    """Re-save through LibreOffice (if installed) so formula results are stored in the file.
    Excel recalculates on open anyway; this helps viewers that do not."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice not found; formulas will calculate when opened in Excel.")
        return
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", tmp, path],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        shutil.copyfile(os.path.join(tmp, os.path.basename(path)), path)
    print("Recalculated with LibreOffice.")


if __name__ == "__main__":
    main()
