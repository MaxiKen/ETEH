"""Shared openpyxl helpers for ETEH bootcamp workbooks."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation

NAVY="1F3864"; TEAL="0E6B7A"; LIGHT="D9E2F3"; GOLD="FFF2CC"; GREY="F2F2F2"; GREEN="E2EFDA"; RED="FCE4EC"; WHITE="FFFFFF"
TITLE_FONT=Font(name="Calibri",size=14,bold=True,color=NAVY)
SUB_FONT=Font(name="Calibri",size=10,italic=True,color="595959")
HDR_FONT=Font(name="Calibri",size=10,bold=True,color="FFFFFF")
HDR_FILL=PatternFill("solid",fgColor=NAVY)
SEC_FONT=Font(name="Calibri",size=11,bold=True,color=NAVY)
BODY_FONT=Font(name="Calibri",size=10)
SMALL_FONT=Font(name="Calibri",size=9)
SMALL_GREY=Font(name="Calibri",size=9,italic=True,color="595959")
BOLD10=Font(name="Calibri",size=10,bold=True)
THIN=Side(style="thin",color="B4C6E7")
BORDER=Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
WRAP=Alignment(wrap_text=True,vertical="top")
CENTER=Alignment(horizontal="center",vertical="center",wrap_text=True)
LEFT_TOP=Alignment(horizontal="left",vertical="top",wrap_text=True)

def setup_sheet(ws, title, subtitle, tab_color=NAVY, landscape=True):
    ws.sheet_properties.tabColor=tab_color
    ws["A1"]=title; ws["A1"].font=TITLE_FONT
    ws["A2"]=subtitle; ws["A2"].font=SUB_FONT
    ws["A2"].alignment=Alignment(wrap_text=True,vertical="top")
    ws.row_dimensions[2].height=28
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.orientation="landscape" if landscape else "portrait"
    ws.page_setup.paperSize=ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.page_margins.left=0.4; ws.page_margins.right=0.4; ws.page_margins.top=0.5; ws.page_margins.bottom=0.5
    ws.oddHeader.center.text="ETEH Ltd — Excel Data Analytics Bootcamp"
    ws.oddHeader.center.font="Calibri,Regular"; ws.oddHeader.center.size=9
    ws.oddFooter.center.text="Page &P of &N"
    ws.oddFooter.center.size=9
    return ws

def section(ws, row, text, ncols=8):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncols)
    c=ws.cell(row=row,column=1,value=text)
    c.font=SEC_FONT; c.fill=PatternFill("solid",fgColor=LIGHT)
    c.alignment=Alignment(vertical="center")
    ws.row_dimensions[row].height=20
    return row+1

def note(ws, row, text, ncols=8, height=30, bold_prefix=None):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncols)
    c=ws.cell(row=row,column=1,value=text)
    c.font=BODY_FONT; c.alignment=WRAP
    ws.row_dimensions[row].height=height
    return row+1

def header_row(ws, row, headers, widths=None):
    for j,h in enumerate(headers,1):
        c=ws.cell(row=row,column=j,value=h)
        c.font=HDR_FONT; c.fill=HDR_FILL; c.alignment=CENTER; c.border=BORDER
    ws.row_dimensions[row].height=32
    if widths:
        for j,w in enumerate(widths,1):
            ws.column_dimensions[get_column_letter(j)].width=w
    return row+1

def write_table(ws, start_row, headers, rows, widths=None, money_cols=(), pct_cols=(), num_fmt=None, table_name=None, banded=True):
    """Write header + rows. rows: list of lists (values or formula strings starting with =)."""
    r = header_row(ws, start_row, headers, widths)
    first_data = r
    for i,row in enumerate(rows):
        for j,v in enumerate(row,1):
            c=ws.cell(row=r,column=j,value=v)
            c.font=Font(name="Calibri",size=9); c.border=BORDER
            c.alignment=Alignment(vertical="center",wrap_text=True)
            if banded and i%2==1:
                c.fill=PatternFill("solid",fgColor=GREY)
            if j in money_cols:
                c.number_format='#,##0.00'
            if j in pct_cols:
                c.number_format='0.0%'
        ws.row_dimensions[r].height=15
        r+=1
    last_data = r-1
    # filter + freeze
    ws.auto_filter.ref=f"A{start_row}:{get_column_letter(len(headers))}{last_data}"
    ws.freeze_panes=ws.cell(row=first_data,column=1).coordinate
    # Excel Table object (optional; only if name given and no merged cells conflict)
    if table_name:
        try:
            tab=Table(displayName=table_name, ref=f"A{start_row}:{get_column_letter(len(headers))}{last_data}")
            tab.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            ws.add_table(tab)
        except Exception as e:
            print(f"  table warn {table_name}: {e}")
    return r  # next free row

def add_check_box(ws, row, text, ncols=8):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncols)
    c=ws.cell(row=row,column=1,value=f"☐  {text}")
    c.font=BODY_FONT; c.alignment=WRAP
    ws.row_dimensions[row].height=22
    return row+1

def add_answer_row(ws, row, label, formula_or_value, expected=None, ncols=8):
    ws.cell(row=row,column=1,value=label).font=BOLD10
    c=ws.cell(row=row,column=2,value=formula_or_value)
    c.font=Font(name="Calibri",size=10); c.border=BORDER
    if isinstance(formula_or_value,str) and formula_or_value.startswith("="):
        c.fill=PatternFill("solid",fgColor=GREEN)
    ws.merge_cells(start_row=row,start_column=3,end_row=row,end_column=ncols)
    exp = f"Expected: {expected}" if expected is not None else "Check with =FORMULAEVALUATION or manual calc. See ANSWERS sheet."
    ws.cell(row=row,column=3,value=exp).font=SMALL_GREY
    ws.cell(row=row,column=3).alignment=WRAP
    ws.row_dimensions[row].height=28
    return row+1

def readme_sheet(wb, title, subtitle, contents, dataset_map, howto):
    ws=wb.active; ws.title="00_README"
    setup_sheet(ws, title, subtitle, tab_color=NAVY, landscape=False)
    r=4
    r=section(ws,r,"What this workbook is",ncols=6)
    for line in contents:
        r=note(ws,r,line,ncols=6,height=30)
    r+=1
    r=section(ws,r,"Dataset map — which file, which sheet, which module",ncols=6)
    r=header_row(ws,r,["Dataset (CLEAN file)","Sheet in this workbook","Used in modules","Rows in sheet","Grain"],widths=[38,26,30,16,30])
    for row in dataset_map:
        for j,v in enumerate(row,1):
            c=ws.cell(row=r,column=j,value=v); c.font=Font(name="Calibri",size=9); c.border=BORDER; c.alignment=WRAP
        ws.row_dimensions[r].height=22; r+=1
    r+=1
    r=section(ws,r,"How to work (read first)",ncols=6)
    for line in howto:
        r=note(ws,r,line,ncols=6,height=32)
    ws.column_dimensions["A"].width=38; ws.column_dimensions["B"].width=26
    ws.column_dimensions["C"].width=30; ws.column_dimensions["D"].width=16; ws.column_dimensions["E"].width=30
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.print_title_rows="1:2"
    return ws

def data_sheet(wb, name, df, title, subtitle, tab_color=TEAL, max_rows=500, table_name=None):
    ws=wb.create_sheet(title=name[:31])
    setup_sheet(ws, title, subtitle, tab_color=tab_color)
    # header at row 4 (rows 1-2 title, row 3 blank/note)
    ws["A3"]=f"Source: see 00_README. One header row, one record per row. Do not insert blank rows/columns. Showing first {min(len(df),max_rows):,} of {len(df):,} rows."
    ws["A3"].font=SMALL_GREY
    headers=list(df.columns)
    r=4
    r=header_row(ws,r,headers)
    # widths
    for j,h in enumerate(headers,1):
        ws.column_dimensions[get_column_letter(j)].width=min(max(11,len(str(h))+2),26)
    n=min(len(df),max_rows)
    for i in range(n):
        row=df.iloc[i]
        for j,v in enumerate(row,1):
            import pandas as pd, numpy as np
            if pd.isna(v): val=None
            elif isinstance(v,(np.integer,)): val=int(v)
            elif isinstance(v,(np.floating,)): val=float(v)
            else: val=v
            c=ws.cell(row=r,column=j,value=val)
            c.font=Font(name="Calibri",size=9); c.border=BORDER
            if i%2==1: c.fill=PatternFill("solid",fgColor=GREY)
        ws.row_dimensions[r].height=13
        r+=1
    ws.auto_filter.ref=f"A4:{get_column_letter(len(headers))}{r-1}"
    ws.freeze_panes="A5"
    ws.print_title_rows="1:4"
    if table_name:
        try:
            tab=Table(displayName=table_name, ref=f"A4:{get_column_letter(len(headers))}{r-1}")
            tab.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(tab)
        except Exception as e:
            print(f"  table warn {table_name}: {e}")
    return ws
