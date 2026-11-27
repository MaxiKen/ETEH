# -*- coding: utf-8 -*-
"""
ETEH Bootcamp - Workbook generator.

Reads the curriculum spec (curriculum.py) and the standardised clean datasets,
and produces:

  Section_Examples/WeekN_Section_Examples.xlsx   (one worksheet per module,
                                                  each with dataset-derived
                                                  worked examples + live formulas)
  Assignments/WeekN_Assignments.xlsx             (one worksheet per module with
                                                  practice tasks + an Answer Key)

All example tables are sliced from the real curated datasets so the content is
grounded, simple and relatable.
"""
import os, datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from curriculum import WEEKS, DATASET_FILES

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLEAN = os.path.join(ROOT, "Datasets", "Clean")
SEC   = os.path.join(ROOT, "Section_Examples")
ASG   = os.path.join(ROOT, "Assignments")
os.makedirs(SEC, exist_ok=True); os.makedirs(ASG, exist_ok=True)

C_HEAD = "11487A"; C_SUB = "1F6FB2"; C_NOTE = "FFF4CE"; C_TASK = "E8F0FE"
THIN = Side(style="thin", color="BBBBBB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

_cache = {}
def load_ds(key):
    if key not in _cache:
        _cache[key] = pd.read_csv(os.path.join(CLEAN, DATASET_FILES[key]))
    return _cache[key].copy()

def get_slice(spec):
    kind, payload = spec
    if kind == "custom":
        return pd.DataFrame(payload)
    df = load_ds(kind)
    p = payload
    cols = p.get("cols", list(df.columns))
    df = df[[c for c in cols if c in df.columns]]
    if "filter" in p:
        col, val = p["filter"]
        df = df[df[col].astype(str) == str(val)]
    if "group_by" in p:
        gb = p["group_by"]
        if gb == "month" and "order_date" in df.columns:
            df = df.copy()
            df["month"] = df["order_date"].astype(str).str[:7]
            gb = "month"
        num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c]) and c != gb]
        if not num:
            num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        g = df.groupby(gb).agg({c: "sum" for c in num}).reset_index()
        if num:
            g = g.sort_values(num[0], ascending=False)
        if gb == "month":
            g = g.rename(columns={"month": "order_date"})
        df = g.head(p.get("n", 10))
    else:
        if "sort_by" in p:
            df = df.sort_values(p["sort_by"], ascending=p.get("ascending", True))
        df = df.head(p.get("n", 10))
    return df.reset_index(drop=True)

def col_widths(df, show):
    for c in show:
        try:
            w = max([len(str(c))] + [len(str(v)) for v in df[c].head(20)])
        except Exception:
            w = 12
        yield min(max(w + 2, 10), 40)

def render_block(ws, r, block):
    """Render one example block starting at row r. Returns next free row."""
    # block title
    ws.cell(r, 1, "✔ " + block["title"]).font = Font(bold=True, color=C_SUB, size=11)
    r += 1
    ws.cell(r, 1, block["intro"]).font = Font(italic=True, color="555555", size=9)
    ws.cell(r, 1).alignment = Alignment(wrap_text=True)
    r += 1
    df = get_slice(block["src"])
    show = block.get("show") or list(df.columns)
    derive = block.get("derive", {}) or {}
    # ensure show columns exist
    show = [c for c in show if c in df.columns]
    # table header
    hdr = show + list(derive.keys())
    for j, h in enumerate(hdr, 1):
        c = ws.cell(r, j, h.replace("_", " ").title())
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = PatternFill("solid", fgColor=C_HEAD)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    header_row = r
    r += 1
    first_data = r
    # data rows
    for _, row in df.iterrows():
        for j, col in enumerate(show, 1):
            v = row[col]
            if pd.isna(v):
                v = ""
            elif isinstance(v, (pd.Timestamp,)):
                v = str(v)
            cell = ws.cell(r, j, v)
            cell.border = BORDER
            cell.font = Font(size=9)
            if isinstance(v, (int, float)):
                cell.alignment = Alignment(horizontal="right")
        # derived formula columns
        for k, (dcol, formula) in enumerate(derive.items(), start=len(show) + 1):
            f = formula.replace("{r}", str(r))
            cell = ws.cell(r, k, f)
            cell.border = BORDER
            cell.font = Font(size=9, color="1F6FB2")
        r += 1
    last_data = r - 1
    # totals row
    num_cols = [c for c in show if pd.api.types.is_numeric_dtype(df[c])] + list(derive.keys())
    if block.get("totals") and num_cols:
        ws.cell(r, 1, "TOTAL").font = Font(bold=True, size=9)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor="E8EEF5")
        for j, col in enumerate(hdr, 1):
            if col in num_cols:
                L = get_column_letter(j)
                ws.cell(r, j, f"=SUM({L}{first_data}:{L}{last_data})").font = Font(bold=True, size=9)
                ws.cell(r, j).fill = PatternFill("solid", fgColor="E8EEF5")
            ws.cell(r, j).border = BORDER
        r += 1
    # worked note
    r += 1
    wcell = ws.cell(r, 1, "Worked example: " + block["worked"])
    wcell.fill = PatternFill("solid", fgColor=C_NOTE)
    wcell.alignment = Alignment(wrap_text=True, vertical="top")
    wcell.font = Font(size=9)
    ws.merge_cells(start_row=r, start_column=1, end_row=r+2, end_column=max(4, len(hdr)))
    r += 3
    # task note
    tcell = ws.cell(r, 1, "Your turn: " + block["task"])
    tcell.fill = PatternFill("solid", fgColor=C_TASK)
    tcell.alignment = Alignment(wrap_text=True, vertical="top")
    tcell.font = Font(size=9, italic=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r+2, end_column=max(4, len(hdr)))
    r += 4
    return r

def style_sheet(ws, week_title, module):
    ws.cell(1, 1, week_title).font = Font(bold=True, size=14, color=C_HEAD)
    ws.cell(2, 1, f"Module {module['id']}  ·  {module['name']}").font = Font(bold=True, size=11, color=C_SUB)
    ws.cell(3, 1, "Objective: " + module["objective"]).font = Font(italic=True, size=9, color="555555")
    ws.cell(3, 1).alignment = Alignment(wrap_text=True)
    ws.column_dimensions["A"].width = 26
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 18
    ws.sheet_view.showGridLines = False
    return 5  # first content row

def sheet_name(module):
    nm = f"{module['id']} {module['name']}"
    for ch in [":", "\\", "/", "?", "*", "[", "]"]:
        nm = nm.replace(ch, " ")
    return nm[:31]

def build_section():
    for week in WEEKS:
        wb = Workbook(); wb.remove(wb.active)
        # overview sheet
        ov = wb.create_sheet("Overview")
        ov.cell(1,1, week["title"]).font = Font(bold=True, size=13, color=C_HEAD)
        ov.cell(2,1, "Section Examples — one worksheet per module, each with relatable worked examples drawn from the curated datasets.").font = Font(italic=True, size=9)
        ov.cell(4,1,"Module"); ov.cell(4,2,"Topic"); ov.cell(4,3,"Examples")
        for c in range(1,4):
            ov.cell(4,c).font = Font(bold=True, color="FFFFFF"); ov.cell(4,c).fill = PatternFill("solid", fgColor=C_HEAD)
        rr=5
        for m in week["modules"]:
            ov.cell(rr,1,m["id"]); ov.cell(rr,2,m["name"]); ov.cell(rr,3,len(m["blocks"]))
            rr+=1
        ov.column_dimensions["A"].width=8; ov.column_dimensions["B"].width=46; ov.column_dimensions["C"].width=10
        for m in week["modules"]:
            ws = wb.create_sheet(sheet_name(m))
            r = style_sheet(ws, week["title"], m)
            for blk in m["blocks"]:
                r = render_block(ws, r, blk)
        out = os.path.join(SEC, week["file"])
        wb.save(out)
        print(f"[Section] {week['file']}: {len(week['modules'])} module sheets")

def build_assignments():
    for week in WEEKS:
        wb = Workbook(); wb.remove(wb.active)
        answers = []  # (module_id, name, q, expected)
        for m in week["modules"]:
            ws = wb.create_sheet(sheet_name(m))
            ws.cell(1,1, week["title"]).font = Font(bold=True, size=13, color=C_HEAD)
            ws.cell(2,1, f"Assignment — Module {m['id']}: {m['name']}").font = Font(bold=True, size=11, color=C_SUB)
            ws.cell(3,1,"Complete the tasks below using the dataset slice provided, then check the Answer Key sheet.").font = Font(italic=True, size=9, color="555555")
            r = 5
            for i, blk in enumerate(m["blocks"], 1):
                ws.cell(r,1, f"Task {i}: {blk['title']}").font = Font(bold=True, color=C_SUB, size=11)
                r += 1
                ws.cell(r,1, blk["intro"]).font = Font(italic=True, size=9, color="555555")
                ws.cell(r,1).alignment = Alignment(wrap_text=True)
                r += 1
                df = get_slice(blk["src"])
                show = blk.get("show") or list(df.columns)
                show = [c for c in show if c in df.columns]
                hdr = show + list((blk.get("derive") or {}).keys())
                for j,h in enumerate(hdr,1):
                    c = ws.cell(r,j,h.replace("_"," ").title())
                    c.font=Font(bold=True,color="FFFFFF",size=9); c.fill=PatternFill("solid",fgColor=C_HEAD); c.border=BORDER
                r += 1
                for _, row in df.iterrows():
                    for j,col in enumerate(show,1):
                        v=row[col]
                        if pd.isna(v): v=""
                        cell=ws.cell(r,j,v); cell.border=BORDER; cell.font=Font(size=9)
                        if isinstance(v,(int,float)): cell.alignment=Alignment(horizontal="right")
                    for k,dcol in enumerate((blk.get("derive") or {}).keys(), start=len(show)+1):
                        ws.cell(r,k,"").border=BORDER
                    r += 1
                r += 1
                q = ws.cell(r,1, "Submit: " + blk["task"])
                q.fill = PatternFill("solid", fgColor=C_TASK); q.font=Font(size=9, italic=True)
                q.alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=r, start_column=1, end_row=r+2, end_column=max(4,len(hdr)))
                r += 4
                answers.append((m["id"], m["name"], f"{blk['title']} — {blk['task'][:90]}",
                               blk["worked"][:160]))
            ws.column_dimensions["A"].width=26
            for col in "BCDEFG": ws.column_dimensions[col].width=18
            ws.sheet_view.showGridLines=False
        # answer key sheet
        ak = wb.create_sheet("Answer Key")
        ak.cell(1,1,"Answer Key & Mark Scheme (sample solutions / techniques)").font=Font(bold=True,size=13,color=C_HEAD)
        ak.cell(3,1,"Module"); ak.cell(3,2,"Topic"); ak.cell(3,3,"Task"); ak.cell(3,4,"Expected approach / result")
        for c in range(1,5):
            ak.cell(3,c).font=Font(bold=True,color="FFFFFF"); ak.cell(3,c).fill=PatternFill("solid",fgColor=C_HEAD)
        rr=4
        for mid,name,q,exp in answers:
            ak.cell(rr,1,mid); ak.cell(rr,2,name); ak.cell(rr,3,q); ak.cell(rr,4,exp)
            for c in range(1,5): ak.cell(rr,c).alignment=Alignment(wrap_text=True, vertical="top"); ak.cell(rr,c).font=Font(size=9)
            rr+=1
        ak.column_dimensions["A"].width=8; ak.column_dimensions["B"].width=30
        ak.column_dimensions["C"].width=55; ak.column_dimensions["D"].width=60
        # order: overview first
        out = os.path.join(ASG, week["file"].replace("Section_Examples","Assignments"))
        wb.save(out)
        print(f"[Assign]  {os.path.basename(out)}: {len(week['modules'])} module sheets + Answer Key")

if __name__ == "__main__":
    build_section()
    build_assignments()
    print("Workbooks complete.")
