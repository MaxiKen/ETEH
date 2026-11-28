"""Shared data and helpers for the Week 2 edition-3 workbook (02_sec_w2.py).

Every table name matches the slides (w2_data.py): tSales25, tStores, tDirty25, tSM, tAdm23 sit on Data_ sheets starting in A1,
so a cell address quoted on a slide (e.g. tDirty25 cell R23 holds the text '₦239,240') is the same cell in the workbook.
Small tables (tFiles, tMonthFiles, tWebStores, tBulletin, tCatMap, tZoneCat) sit on the module sheet that uses them.
"""
from __future__ import annotations

import math
import re
from types import SimpleNamespace

import numpy as np
import pandas as pd

import w2_data as W
from wk_common import L
from xlhelp import LAST_COL

N = "naira"
XL0 = pd.Timestamp("1899-12-30")
MONTHS = W.MONTHS


def serial(t):
    """Excel date serial for a Timestamp (cached values of date formulas)."""
    return (pd.Timestamp(t) - XL0).days


def iso_or_text(v):
    if isinstance(v, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        return pd.Timestamp(v)
    return v


def data():
    d = SimpleNamespace()
    d.S = W.table("tSales25")
    d.ST = W.table("tStores")
    d.SM = W.sales().copy()
    dr = L("dirty_sm").copy()
    dr["Order_Date"] = dr.Order_Date.map(iso_or_text)
    dr["Ship_Date"] = dr.Ship_Date.map(iso_or_text)
    d.DR = dr.astype(object).where(dr.notna(), None)
    d.RAW = W.dirty_raw()                       # the same file as plain text (what Power Query reads)
    d.A = W.table("tAdm23")
    d.FILES = W.table("tFiles")
    d.MF = W.table("tMonthFiles")
    d.WEB = W.table("tWebStores")
    d.BUL = W.table("tBulletin")
    d.ZC = W.table("tZoneCat")
    d.DEPT = W.table("tDept23")
    return d


def cell_of(df, col, i):
    """A1 address of data row i (0-based) of column `col` in a table that starts in A1."""
    from xlsxwriter.utility import xl_rowcol_to_cell
    return xl_rowcol_to_cell(i + 1, list(df.columns).index(col))


def lit_grid(m, rows, header=None, widths_chars=None, merge_last=False):
    """Bordered grid whose cells are always literal (text that starts with '=' is shown, never calculated)."""
    s = m.s
    ws, F = s.ws, s.F
    if header:
        for j, h in enumerate(header):
            if merge_last and j == len(header) - 1 and 1 + j < LAST_COL:
                ws.merge_range(s.r, 1 + j, s.r, LAST_COL, h, F["hdr"])
            else:
                ws.write_string(s.r, 1 + j, str(h), F["hdr"])
        s.r += 1
    for row in rows:
        lines = 1
        for j, v in enumerate(row):
            last = merge_last and j == len(row) - 1
            fmt = F["cell"]
            if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
                fmt_n = m.b.nf("int" if float(v).is_integer() else "dec2")
                if last and 1 + j < LAST_COL:
                    ws.merge_range(s.r, 1 + j, s.r, LAST_COL, "", fmt_n)
                ws.write_number(s.r, 1 + j, float(v), fmt_n)
                continue
            v = "" if v is None else str(v)
            if last and 1 + j < LAST_COL:
                ws.merge_range(s.r, 1 + j, s.r, LAST_COL, "", fmt)
                w = (LAST_COL - j) * 12
            else:
                w = 12
            ws.write_string(s.r, 1 + j, v, fmt)
            lines = max(lines, math.ceil(len(v) / (w * 1.1)) if v else 1)
        if lines > 1:
            ws.set_row(s.r, 14 * lines + 2)
        s.r += 1
    s.r += 1


def code(m, text, label=None):
    if label:
        m.h(label)
    m.s.code(text)
    m.s.r += 1


# ---------------------------------------------------------------- M scripts used on several sheets
M_DIRTY = '''let
    Source     = Csv.Document(File.Contents(DataFolder & "SuperMart_Sales_2025_DIRTY.csv"),
                     [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted   = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    NoColumn1  = Table.RemoveColumns(Promoted, {"Column1"}),
    NoBlanks   = Table.SelectRows(NoColumn1, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    Dedup      = Table.Distinct(NoBlanks),
    Stores     = {"Lagos - Ikeja", "Lagos - Lekki", "Ibadan - Ring Road", "Port Harcourt - GRA", "Benin City - Ring Road",
                  "Uyo - Ewet", "Abuja - Wuse", "Abuja - Garki", "Jos - Rayfield", "Enugu - Independence Layout",
                  "Onitsha - Main Market", "Owerri - New Owerri"},            // official spellings (Text.Proper would give "Gra")
    Trimmed    = Table.TransformColumns(Dedup, {{"Zone", Text.Trim, type text},
                     {"Store_Name", (s) => List.First(List.Select(Stores, (n) => Text.Lower(n) = Text.Lower(Text.Trim(s))), "CHECK"), type text},
                     {"Customer_Name", each Text.Trim(Text.Clean(Text.Replace(_, "#(00A0)", " "))), type text}}),
    Category   = Table.TransformColumns(Trimmed, {{"Category", each
                     let c = Text.Lower(Text.Trim(_)) in
                     if Text.StartsWith(c, "fur") then "Furniture" else if Text.StartsWith(c, "off") then "Office Supplies"
                     else if Text.StartsWith(c, "tec") then "Technology" else "CHECK", type text}}),
    Money      = Table.TransformColumns(Category, {
                     {"Sales", each Number.From(Text.Remove(_, {"₦", ","}), "en-US"), type number},
                     {"Discount", each if Text.EndsWith(_, "%") then Number.From(Text.TrimEnd(_, "%"), "en-US") / 100
                                       else Number.From(_, "en-US"), type number}}),
    Typed      = Table.TransformColumnTypes(Money, {{"Profit", type number}, {"Quantity", Int64.Type}, {"Row_ID", Int64.Type}}, "en-US"),
    Dates      = Table.TransformColumns(Typed, {{"Order_Date", each
                     if Text.Contains(_, "/") then Date.FromText(_, [Format = "dd/MM/yyyy"])
                     else if Text.Contains(_, ".") then Date.FromText(Text.Replace(_, ".", "-"))
                     else Date.FromText(_, [Culture = "en-GB"]), type date}}),
    NoNA       = Table.ReplaceValue(Dates, "N/A", null, Replacer.ReplaceValue, {"Customer_Name"})
in
    NoNA'''

M_MERCYLAND = '''let
    Source    = Csv.Document(File.Contents(DataFolder & "Mercyland_Episodes_2023_DIRTY.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted  = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Dedup     = Table.Distinct(Promoted),
    Names     = Table.TransformColumns(Dedup, {{"Patient_Name", each Text.Proper(Text.Trim(_)), type text}}),
    WardMap   = Table.TransformColumns(Names, {{"Ward", each
                    let w = Text.Lower(Text.Trim(Text.Replace(Text.Replace(_, "  ", " "), "Saphire", "Mahogany"))) in
                    if Text.StartsWith(w, "iroko") then "Iroko Ward" else if Text.StartsWith(w, "mahogany") then "Mahogany Ward"
                    else if Text.StartsWith(w, "obeche") then "Obeche Ward" else if Text.StartsWith(w, "baobab") then "Baobab Ward" else "CHECK", type text}}),
    Gender    = Table.TransformColumns(WardMap, {{"Gender", each if List.Contains({"m", "male"}, Text.Lower(_)) then "Male" else "Female", type text}}),
    AdmType   = Table.TransformColumns(Gender, {{"Admission_Route", each
                    let t = Text.Lower(Text.Trim(_)) in if Text.StartsWith(t, "emerg") then "Emergency" else Text.Proper(t), type text}}),
    Bill      = Table.TransformColumnTypes(AdmType, {{"Bill_Amount_NGN", type number}}, "en-US"),   // '2,507,101.35' parses with en-US
    Dates     = Table.TransformColumnTypes(Bill, {{"Admit_Date", type date}, {"Discharge_Date", type date}, {"Age", Int64.Type}}),
    LOS       = Table.AddColumn(Dates, "LOS_Days", each if [Discharge_Date] = null then null else Duration.Days([Discharge_Date] - [Admit_Date]), Int64.Type),
    Exception = Table.AddColumn(LOS, "Exception", each
                    if [Bill_Amount_NGN] < 0 then "Negative bill" else if [Discharge_Date] = null then "No discharge date"
                    else if [LOS_Days] < 0 then "Discharge before admission" else if [Age] = 0 or [Age] > 120 then "Impossible age" else null, type text),
    Clean     = Table.SelectRows(Exception, each [Exception] = null)
in
    Clean'''

M_KINGSWAY = '''let
    Source  = Csv.Document(File.Contents(CoreFolder & "Kingsway_Student_Results.csv"), [Delimiter = ",", Encoding = 65001]),
    Promote = Table.PromoteHeaders(Source),
    Typed   = Table.TransformColumnTypes(Promote, {{"Lecture_Attendance_Pct", type number}, {"Final_Score", Int64.Type}, {"Study_Hours_Wk", Int64.Type}}),
    Band    = Table.AddColumn(Typed, "Attendance_Band2", each
                if [Lecture_Attendance_Pct] >= 90 then "90%+" else if [Lecture_Attendance_Pct] >= 75 then "75-89%" else "Below 75%", type text),
    Risk    = Table.AddColumn(Band, "Needs_Mentor", each [Final_Score] < 60 and [Lecture_Attendance_Pct] < 75, type logical)
in
    Risk'''

M_STORE_MONTH = '''let
    Source  = qSales_Clean,                                     // reference the cleaned query (2.2.9)
    Month   = Table.AddColumn(Source, "Month", each Date.ToText([Order_Date], "yyyy-MM"), type text),
    Grouped = Table.Group(Month, {"Store_Name", "Month"}, {{"Sales", each List.Sum([Sales]), type number},
                                                         {"Profit", each List.Sum([Profit]), type number},
                                                         {"Orders", each List.Count(List.Distinct([Order_ID])), Int64.Type}}),
    Margin  = Table.AddColumn(Grouped, "Margin", each [Profit] / [Sales], Percentage.Type)
in
    Margin'''
