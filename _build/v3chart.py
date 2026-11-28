"""Edition-3 chart blocks: real, editable PowerPoint charts drawn from the same data as the slide copy.

A topic gets a chart with  chart=dict(kind=..., ...)  (see SPEC below); v3kit turns it into a body block that pairs with a
grid / fx / steps block like any other. Every chart is a native PowerPoint chart (Insert › Chart in PowerPoint shows the
data), so learners can click it and see the series. Kinds python-pptx cannot write natively are built the way analysts
build them in older Excel versions:
  waterfall  stacked column with an invisible base series
  funnel     stacked bar with an invisible centring series, categories top-down
  histogram  column chart with a 5% gap
  box        stacked column (invisible base, Q1-median, median-Q3) with custom error bars for the whiskers
  sunburst   two-ring doughnut: inner ring = category totals, outer ring = subcategories, aligned
  combo      clustered column + line on a secondary axis (the second series is moved into a lineChart on new axes)
  treemap    squarified rectangles drawn as shapes (area proportional to value)
  sparklines a grid of tiny line charts, one per row label

SPEC (all optional except kind and the data):
  kind      column | stacked | bar | line | area | pie | doughnut | scatter | bubble | radar | combo | waterfall | funnel |
            histogram | box | sunburst | treemap | sparklines
  cats      category labels                       series  [(name, values), ...]
  points    scatter: [(name, xs, ys)]; bubble: [(name, xs, ys, sizes)]
  title     chart title (state the finding)       fmt     value number format, e.g. '₦#,##0,,"m"' or '0%'
  labels    True = data labels on every point, or a list of point indexes to label (first series)
  legend    True/False (default: True when more than one series)
  h         chart height in inches (default 2.75)  label   block label (default CHART FROM OUR DATA)
  note      grey note under the chart              colors  list of hex colours (per series, or per point for pie/bar)
  x_title / y_title  axis titles                  y2_fmt  secondary-axis number format (combo)
  box       box: [(name, lo, q1, med, q3, hi)]     tree    treemap: [(label, value)]
  rings     sunburst: [(category, [(sub, value), ...]), ...]
  rows      sparklines: [(label, values)]          ncol    sparklines: columns of the grid (default 2)
"""
from __future__ import annotations

import copy

from lxml import etree
from pptx.chart.data import BubbleChartData, CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION, XL_TICK_MARK
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

import slidekit as K
from slidekit import GREY, INK, NAVY, ORANGE, WHITE, wrap

TEAL, AMBER, SKY, PLUM, SAND = "13807A", "C98A00", "4A90C2", "7A4E8C", "B8A27A"
PALETTE = [NAVY, ORANGE, TEAL, AMBER, SKY, PLUM, SAND, "6B7B8C", "A3423C", "3E7D4F"]
LIGHT = "D9DEE4"
EMU = K.EMU


def _e(v):
    return Emu(int(v * EMU))


def _rgb(h):
    return RGBColor.from_string(h)


def height(spec):
    return spec.get("_h", spec.get("h", 2.75))


# ---------------------------------------------------------------- styling helpers
def _fill(fmt, color):
    fmt.fill.solid()
    fmt.fill.fore_color.rgb = _rgb(color)


def _nofill(ser):
    ser.format.fill.background()
    ser.format.line.fill.background()


def _axis_style(ax, fmt=None, title=None, grid=False):
    ax.has_major_gridlines = grid
    if grid:
        ax.major_gridlines.format.line.color.rgb = _rgb(LIGHT)
        ax.major_gridlines.format.line.width = Pt(0.75)
    ax.format.line.color.rgb = _rgb("BFC5CC")
    ax.tick_labels.font.size = Pt(10)
    ax.tick_labels.font.color.rgb = _rgb("44505C")
    if fmt:
        ax.tick_labels.number_format = fmt
        ax.tick_labels.number_format_is_linked = False
    if title:
        ax.has_title = True
        ax.axis_title.text_frame.text = title
        r = ax.axis_title.text_frame.paragraphs[0].runs[0]
        r.font.size, r.font.bold = Pt(10), False
        r.font.color.rgb = _rgb("44505C")


def _title(ch, text):
    if not text:
        ch.has_title = False
        return
    ch.has_title = True
    tf = ch.chart_title.text_frame
    tf.text = text
    r = tf.paragraphs[0].runs[0]
    r.font.size, r.font.bold = Pt(12), True
    r.font.color.rgb = _rgb(NAVY)


def _legend(ch, on):
    ch.has_legend = bool(on)
    if on:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(10)


def _labels(plot_or_ser, fmt=None, pos=None, size=9.5, color=INK):
    dl = plot_or_ser.data_labels
    dl.font.size = Pt(size)
    dl.font.color.rgb = _rgb(color)
    if fmt:
        dl.number_format = fmt
        dl.number_format_is_linked = False
    if pos is not None:
        try:
            dl.position = pos
        except Exception:  # noqa: BLE001  (not every chart type accepts every position)
            pass
    return dl


def _point_labels(ser, idxs, fmt=None, pos=None):
    for i in idxs:
        dl = ser.points[i].data_label
        dl.has_text_frame = False
        dl.show_value = True
        dl.font.size = Pt(9.5)
        dl.font.bold = True
        if pos is not None:
            try:
                dl.position = pos
            except Exception:  # noqa: BLE001
                pass
    if fmt:
        # point-level labels inherit the series number format
        d = _labels(ser, fmt)
        d.show_value = False


def _cat_data(cats, series, fmt=None):
    cd = CategoryChartData(number_format=fmt or "General")
    cd.categories = list(cats)
    for name, vals in series:
        cd.add_series(name, [None if v is None or v != v else float(v) for v in vals])   # v != v: NaN = a gap
    return cd


# ---------------------------------------------------------------- XML helpers (combo, error bars)
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _el(tag, **attrs):
    e = etree.SubElement(etree.Element("x"), qn(tag))
    for k, v in attrs.items():
        e.set(k, str(v))
    return e


def _num_lit(vals):
    lit = _el("c:numLit")
    fc = etree.SubElement(lit, qn("c:formatCode"))
    fc.text = "General"
    etree.SubElement(lit, qn("c:ptCount")).set("val", str(len(vals)))
    for i, v in enumerate(vals):
        pt = etree.SubElement(lit, qn("c:pt"))
        pt.set("idx", str(i))
        etree.SubElement(pt, qn("c:v")).text = repr(float(v))
    return lit


def _err_bars(ser_el, kind, vals, color="44505C"):
    """Custom error bars on a bar series: kind 'plus' or 'minus'."""
    eb = etree.Element(qn("c:errBars"))
    etree.SubElement(eb, qn("c:errBarType")).set("val", kind)
    etree.SubElement(eb, qn("c:errValType")).set("val", "cust")
    etree.SubElement(eb, qn("c:noEndCap")).set("val", "0")
    holder = etree.SubElement(eb, qn("c:plus" if kind == "plus" else "c:minus"))
    holder.append(_num_lit(vals))
    sp = etree.fromstring(f'<c:spPr xmlns:c="{C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                          f'<a:ln w="15875"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr>')
    eb.append(sp)
    cat = ser_el.find(qn("c:cat"))
    (cat if cat is not None else ser_el.find(qn("c:val"))).addprevious(eb)


def _to_secondary_line(ch, ser_index, color, fmt):
    """Move bar series #ser_index into a new lineChart on a secondary value axis (Excel's Combo chart)."""
    plot = ch._chartSpace.find(".//" + qn("c:plotArea"))
    bar = plot.find(qn("c:barChart"))
    sers = bar.findall(qn("c:ser"))
    ser = sers[ser_index]
    bar.remove(ser)
    inv = ser.find(qn("c:invertIfNegative"))
    if inv is not None:
        ser.remove(inv)
    for old in ser.findall(qn("c:spPr")):
        ser.remove(old)
    sp = etree.fromstring(f'<c:spPr xmlns:c="{C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                          f'<a:ln w="31750" cap="rnd"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:round/></a:ln></c:spPr>')
    ser.find(qn("c:tx")).addnext(sp)
    mk = etree.fromstring(f'<c:marker xmlns:c="{C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:symbol val="circle"/>'
                          f'<c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                          f'<a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker>')
    sp.addnext(mk)
    etree.SubElement(ser, qn("c:smooth")).set("val", "0")
    line = etree.Element(qn("c:lineChart"))
    etree.SubElement(line, qn("c:grouping")).set("val", "standard")
    etree.SubElement(line, qn("c:varyColors")).set("val", "0")
    line.append(ser)
    etree.SubElement(line, qn("c:marker")).set("val", "1")
    etree.SubElement(line, qn("c:axId")).set("val", "60001")
    etree.SubElement(line, qn("c:axId")).set("val", "60002")
    bar.addnext(line)
    # secondary axes: hidden category axis + right-hand value axis, placed after the existing axes
    last_ax = plot.findall(qn("c:valAx"))[-1]
    cat2 = etree.fromstring(
        f'<c:catAx xmlns:c="{C}"><c:axId val="60001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/>'
        f'<c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        f'<c:crossAx val="60002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
        f'<c:noMultiLvlLbl val="0"/></c:catAx>')
    val2 = etree.fromstring(
        f'<c:valAx xmlns:c="{C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:axId val="60002"/>'
        f'<c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="r"/>'
        f'<c:numFmt formatCode="{fmt}" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>'
        f'<c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:solidFill><a:srgbClr val="BFC5CC"/></a:solidFill></a:ln></c:spPr>'
        f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
        f'<c:crossAx val="60001"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>')
    last_ax.addnext(cat2)
    cat2.addnext(val2)


# ---------------------------------------------------------------- treemap (squarified)
def _squarify(vals, x, y, w, h):
    """Return rectangles (x, y, w, h) for vals (sorted desc), areas proportional to the values."""
    out = []
    vals = [float(v) for v in vals]
    tot = sum(vals)
    scale = w * h / tot
    areas = [v * scale for v in vals]

    def worst(row, side):
        s = sum(row)
        return max(max(side * side * r / (s * s), (s * s) / (side * side * r)) for r in row)

    i = 0
    while i < len(areas):
        side = min(w, h)
        row = [areas[i]]
        i += 1
        while i < len(areas) and worst(row + [areas[i]], side) <= worst(row, side):
            row.append(areas[i])
            i += 1
        s = sum(row)
        if w >= h:                       # lay the row as a column on the left
            cw = s / h
            yy = y
            for a in row:
                out.append((x, yy, cw, a / cw))
                yy += a / cw
            x, w = x + cw, w - cw
        else:                            # lay the row along the top
            rh = s / w
            xx = x
            for a in row:
                out.append((xx, y, a / rh, rh))
                xx += a / rh
            y, h = y + rh, h - rh
    return out


def _treemap(cv, x, y, w, h, spec):
    items = sorted(spec["tree"], key=lambda t: -t[1])
    tot = sum(v for _, v in items)
    ty = y
    if spec.get("title"):
        cv.text(x, y, w, 0.28, [spec["title"]], 12, NAVY, bold=True)
        ty += 0.34
    rects = _squarify([v for _, v in items], x, ty, w, h - (ty - y))
    cols = spec.get("colors") or PALETTE
    fmt = spec.get("value_fn") or (lambda v: f"{v / tot:.0%}")
    for i, ((lab, v), (rx, ry, rw, rh)) in enumerate(zip(items, rects)):
        cv.rect(rx + 0.015, ry + 0.015, rw - 0.03, rh - 0.03, cols[i % len(cols)], None)
        if rw > 0.75 and rh > 0.42:
            size = 11 if rw > 1.3 and rh > 0.6 else 9.5
            lines = wrap(lab, size, rw - 0.12, True)
            if len(lines) * size * 1.25 / 72 + size * 1.3 / 72 < rh - 0.08:
                cv.text(rx + 0.07, ry + 0.06, rw - 0.12, rh - 0.1, [f"**{lab}**", fmt(v)], size, WHITE)


# ---------------------------------------------------------------- the main entry point
def draw(cv, x, y, w, h, spec):
    """Draw the chart in the box (x, y, w, h) inches. cv is a slidekit Canvas (or the Dry measurer: then do nothing)."""
    slide = getattr(cv, "slide", None)
    kind = spec["kind"]
    # preview ops (the PNG previews cannot render charts: show a labelled placeholder)
    if slide is not None and hasattr(cv, "ops"):
        cv.ops.append(("rect", x, y, w, h, "F3F5F8", "C9D0D8", False))
        cv.ops.append(("text", x + 0.15, y + 0.12, w - 0.3, 0.3, [f"[{kind} chart] {spec.get('title', '')}"], 11, GREY, True, "l", "t", 0))
    if slide is None:
        return
    if kind == "treemap":
        _treemap(_Raw(cv), x, y, w, h, spec)
        return
    if kind == "sparklines":
        _sparklines(slide, x, y, w, h, spec, cv)
        return
    if kind == "dashboard":
        _dashboard(cv, x, y, w, h, spec)
        return
    cols = spec.get("colors") or PALETTE
    fmt = spec.get("fmt")
    series = spec.get("series", [])
    legend = spec.get("legend", len(series) > 1)
    X, Y, CX, CY = _e(x), _e(y), _e(w), _e(h)

    if kind in ("column", "stacked", "bar", "histogram", "line", "area", "radar", "combo"):
        ct = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "histogram": XL_CHART_TYPE.COLUMN_CLUSTERED, "combo": XL_CHART_TYPE.COLUMN_CLUSTERED,
              "stacked": XL_CHART_TYPE.COLUMN_STACKED, "bar": XL_CHART_TYPE.BAR_CLUSTERED, "line": XL_CHART_TYPE.LINE_MARKERS,
              "area": XL_CHART_TYPE.AREA, "radar": XL_CHART_TYPE.RADAR_MARKERS}[kind]
        if kind == "area" and spec.get("stacked"):
            ct = XL_CHART_TYPE.AREA_STACKED
        ch = slide.shapes.add_chart(ct, X, Y, CX, CY, _cat_data(spec["cats"], series, fmt)).chart
        _title(ch, spec.get("title"))
        ch.font.size = Pt(10)
        plot = ch.plots[0]
        if kind in ("column", "stacked", "bar", "histogram", "combo"):
            plot.gap_width = 5 if kind == "histogram" else spec.get("gap", 60)
            if kind == "stacked":
                plot.overlap = 100
        for i, s in enumerate(plot.series):
            c = cols[i % len(cols)]
            if kind in ("line", "radar"):
                s.format.line.color.rgb = _rgb(c)
                s.format.line.width = Pt(2.25)
                s.smooth = False
                s.marker.format.fill.solid()
                s.marker.format.fill.fore_color.rgb = _rgb(c)
                s.marker.size = 6 if kind == "line" else 5
            else:
                _fill(s.format, c)
        if spec.get("point_colors") and kind in ("column", "bar", "histogram"):
            s0 = plot.series[0]
            for i, c in enumerate(spec["point_colors"]):
                if c:
                    _fill(s0.points[i].format, c)
        if kind != "radar":
            _axis_style(ch.value_axis, fmt, spec.get("y_title"), grid=True)
            _axis_style(ch.category_axis, None, spec.get("x_title"))
            ch.category_axis.major_tick_mark = XL_TICK_MARK.NONE
            if kind == "bar":
                ch.category_axis.reverse_order = True      # first category at the top
                ch.value_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
            if spec.get("y_min") is not None:
                ch.value_axis.minimum_scale = spec["y_min"]
            if spec.get("y_max") is not None:
                ch.value_axis.maximum_scale = spec["y_max"]
        else:
            ch.value_axis.tick_labels.font.size = Pt(9)
            if fmt:
                ch.value_axis.tick_labels.number_format = fmt
                ch.value_axis.tick_labels.number_format_is_linked = False
            ch.category_axis.tick_labels.font.size = Pt(10)
        lab = spec.get("labels")
        if lab is True:
            plot.has_data_labels = True
            pos = {"bar": XL_LABEL_POSITION.OUTSIDE_END, "column": XL_LABEL_POSITION.OUTSIDE_END, "histogram": XL_LABEL_POSITION.OUTSIDE_END,
                   "combo": XL_LABEL_POSITION.OUTSIDE_END, "line": XL_LABEL_POSITION.ABOVE, "stacked": XL_LABEL_POSITION.CENTER}.get(kind)
            _labels(plot, spec.get("label_fmt", fmt), pos, color=WHITE if kind == "stacked" else INK)
        elif isinstance(lab, (list, tuple)):
            pos = XL_LABEL_POSITION.ABOVE if kind == "line" else XL_LABEL_POSITION.OUTSIDE_END
            _point_labels(plot.series[0], lab, spec.get("label_fmt", fmt), pos)
        _legend(ch, legend)
        if kind == "combo":
            _to_secondary_line(ch, 1, cols[1], spec.get("y2_fmt", "0%"))
            _legend(ch, True)
        return ch

    if kind in ("pie", "doughnut"):
        ct = XL_CHART_TYPE.PIE if kind == "pie" else XL_CHART_TYPE.DOUGHNUT
        ch = slide.shapes.add_chart(ct, X, Y, CX, CY, _cat_data(spec["cats"], series[:1], fmt)).chart
        _title(ch, spec.get("title"))
        ch.font.size = Pt(10)
        plot = ch.plots[0]
        plot.vary_by_categories = True
        for i, pt in enumerate(plot.series[0].points):
            _fill(pt.format, cols[i % len(cols)])
            pt.format.line.color.rgb = _rgb(WHITE)
        plot.has_data_labels = True
        dl = _labels(plot, spec.get("label_fmt", "0%"), XL_LABEL_POSITION.OUTSIDE_END if kind == "pie" else None, size=10)
        dl.show_percentage = spec.get("percent", True)
        dl.show_value = not spec.get("percent", True)
        dl.show_category_name = spec.get("cat_labels", False)
        if kind == "doughnut":
            hole = ch._chartSpace.find(".//" + qn("c:holeSize"))
            if hole is not None:
                hole.set("val", "55")
        _legend(ch, spec.get("legend", True))
        return ch

    if kind == "sunburst":
        rings = spec["rings"]
        cats, inner, outer, pcol = [], [], [], []
        for ci, (cat, subs) in enumerate(rings):
            tot = sum(v for _, v in subs)
            for j, (sub, v) in enumerate(subs):
                cats.append(sub)
                inner.append(tot if j == 0 else 0)
                outer.append(v)
                pcol.append(cols[ci % len(cols)])
        cd = _cat_data(cats, [("Category", inner), ("Subcategory", outer)], fmt)
        ch = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, X, Y, CX, CY, cd).chart
        _title(ch, spec.get("title"))
        plot = ch.plots[0]
        plot.vary_by_categories = True
        hole = ch._chartSpace.find(".//" + qn("c:holeSize"))
        if hole is not None:
            hole.set("val", "25")
        for si, s in enumerate(plot.series):
            for i, pt in enumerate(s.points):
                _fill(pt.format, pcol[i])
                pt.format.line.color.rgb = _rgb(WHITE)
                if si == 1:
                    pt.format.fill.fore_color.brightness = 0.35 if i % 2 else 0.15
        # label the inner ring with category names (only the non-zero points)
        s0 = plot.series[0]
        k = 0
        for ci, (cat, subs) in enumerate(rings):
            dl = s0.points[k].data_label
            dl.has_text_frame = True
            dl.text_frame.text = cat
            r = dl.text_frame.paragraphs[0].runs[0]
            r.font.size, r.font.bold = Pt(9.5), True
            r.font.color.rgb = _rgb(WHITE)
            k += len(subs)
        _legend(ch, False)
        return ch

    if kind == "waterfall":
        cats, steps = spec["cats"], spec["steps"]          # steps: list of numbers; last category = total (computed)
        base, up, down, total, run = [], [], [], [], 0.0
        if spec.get("start") is not None:                 # optional opening total (first category)
            run = float(spec["start"])
            base.append(0)
            up.append(0)
            down.append(0)
            total.append(run)
        for v in steps:
            nxt = run + v
            base.append(min(run, nxt))
            up.append(v if v >= 0 else 0)
            down.append(-v if v < 0 else 0)
            total.append(0)
            run = nxt
        base.append(0)
        up.append(0)
        down.append(0)
        total.append(run)
        ser = [("Base", base), ("Increase", up), ("Decrease", down), ("Total", total)]
        ch = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, X, Y, CX, CY, _cat_data(cats, ser, fmt)).chart
        _title(ch, spec.get("title"))
        plot = ch.plots[0]
        plot.gap_width, plot.overlap = 50, 100
        s = plot.series
        _nofill(s[0])
        _fill(s[1].format, spec.get("up", TEAL))
        _fill(s[2].format, spec.get("down", "B42B1E"))
        _fill(s[3].format, NAVY)
        _axis_style(ch.value_axis, fmt, spec.get("y_title"), grid=True)
        _axis_style(ch.category_axis)
        for si in (1, 3):
            for i in range(len(cats)):
                if (up if si == 1 else total)[i]:
                    dl = s[si].points[i].data_label
                    dl.has_text_frame = False
                    dl.show_value = True
                    dl.font.size = Pt(9.5)
                    dl.position = XL_LABEL_POSITION.INSIDE_END
                    dl.font.color.rgb = _rgb(WHITE)
            d = s[si].data_labels
            d.number_format = spec.get("label_fmt", fmt or "General")
            d.number_format_is_linked = False
            d.show_value = False
        _legend(ch, False)
        return ch

    if kind == "funnel":
        cats, vals = spec["cats"], [float(v) for v in spec["values"]]
        mx = max(vals)
        pad = [(mx - v) / 2 for v in vals]
        ch = slide.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, X, Y, CX, CY, _cat_data(cats, [("pad", pad), (spec.get("name", "Count"), vals)], fmt)).chart
        _title(ch, spec.get("title"))
        plot = ch.plots[0]
        plot.gap_width, plot.overlap = 15, 100
        _nofill(plot.series[0])
        s = plot.series[1]
        for i in range(len(vals)):
            _fill(s.points[i].format, cols[i % len(cols)])
        plot.series[1].data_labels  # noqa: B018  (create)
        d = _labels(s, spec.get("label_fmt", fmt or "#,##0"), XL_LABEL_POSITION.CENTER, size=10.5, color=WHITE)
        d.show_value = True
        ch.category_axis.reverse_order = True
        ch.category_axis.format.line.fill.background()
        ch.category_axis.tick_labels.font.size = Pt(10.5)
        ch.category_axis.major_tick_mark = XL_TICK_MARK.NONE
        ch.value_axis.visible = False
        ch.value_axis.has_major_gridlines = False
        ch.value_axis.maximum_scale = mx
        _legend(ch, False)
        return ch

    if kind == "box":
        rows = spec["box"]                                 # (name, lo, q1, med, q3, hi)
        cats = [r[0] for r in rows]
        base = [r[2] for r in rows]
        lower = [r[3] - r[2] for r in rows]
        upper = [r[4] - r[3] for r in rows]
        ch = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, X, Y, CX, CY,
                                    _cat_data(cats, [("Q1", base), ("Q1 to median", lower), ("Median to Q3", upper)], fmt)).chart
        _title(ch, spec.get("title"))
        plot = ch.plots[0]
        plot.gap_width, plot.overlap = 90, 100
        s = plot.series
        _nofill(s[0])
        _fill(s[1].format, spec.get("c1", SKY))
        _fill(s[2].format, spec.get("c2", NAVY))
        for sx in (s[1], s[2]):
            sx.format.line.color.rgb = _rgb(WHITE)
        _err_bars(s[0]._element, "minus", [r[2] - r[1] for r in rows])     # whisker down from Q1 to the low end
        _err_bars(s[2]._element, "plus", [r[5] - r[4] for r in rows])      # whisker up from Q3 to the high end
        _axis_style(ch.value_axis, fmt, spec.get("y_title"), grid=True)
        _axis_style(ch.category_axis)
        _legend(ch, False)
        return ch

    if kind == "scatter":
        cd = XyChartData()
        for name, xs, ys in spec["points"]:
            s = cd.add_series(name)
            for a, b in zip(xs, ys):
                s.add_data_point(float(a), float(b))
        ch = slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, X, Y, CX, CY, cd).chart
        _title(ch, spec.get("title"))
        for i, s in enumerate(ch.plots[0].series):
            c = cols[i % len(cols)]
            s.format.line.fill.background()
            s.marker.style = 8   # circle
            s.marker.size = spec.get("marker", 4)
            s.marker.format.fill.solid()
            s.marker.format.fill.fore_color.rgb = _rgb(c)
            s.marker.format.line.fill.background()
        _axis_style(ch.value_axis, fmt, spec.get("y_title"), grid=True)
        _axis_style(ch.category_axis, spec.get("x_fmt"), spec.get("x_title"))
        for k_, ax in (("x", ch.category_axis), ("y", ch.value_axis)):
            if spec.get(f"{k_}_min") is not None:
                ax.minimum_scale = spec[f"{k_}_min"]
            if spec.get(f"{k_}_max") is not None:
                ax.maximum_scale = spec[f"{k_}_max"]
        if spec.get("trend"):
            for s in [ch.plots[0].series[0]]:
                tl = etree.fromstring(f'<c:trendline xmlns:c="{C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                                      f'<c:spPr><a:ln w="22225"><a:solidFill><a:srgbClr val="{ORANGE}"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>'
                                      f'<c:trendlineType val="linear"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>')
                el = s._element
                anchor = el.find(qn("c:xVal"))
                anchor.addprevious(tl)
        _legend(ch, legend)
        return ch

    if kind == "bubble":
        cd = BubbleChartData()
        for name, xs, ys, ss in spec["points"]:
            s = cd.add_series(name)
            for a, b, z in zip(xs, ys, ss):
                s.add_data_point(float(a), float(b), float(z))
        ch = slide.shapes.add_chart(XL_CHART_TYPE.BUBBLE, X, Y, CX, CY, cd).chart
        _title(ch, spec.get("title"))
        plot = ch.plots[0]
        plot.bubble_scale = spec.get("scale", 45)
        for i, s in enumerate(plot.series):
            for j, pt in enumerate(s.points):
                _fill(pt.format, cols[(i + j) % len(cols)])
                pt.format.fill.fore_color.brightness = 0.1
                pt.format.line.color.rgb = _rgb(WHITE)
        if spec.get("names"):
            s = plot.series[0]
            for j, nm in enumerate(spec["names"]):
                dl = s.points[j].data_label
                dl.has_text_frame = True
                dl.text_frame.text = nm
                dl.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
                dl.position = XL_LABEL_POSITION.CENTER
        _axis_style(ch.value_axis, fmt, spec.get("y_title"), grid=True)
        _axis_style(ch.category_axis, spec.get("x_fmt"), spec.get("x_title"))
        for k_, ax in (("x", ch.category_axis), ("y", ch.value_axis)):
            if spec.get(f"{k_}_min") is not None:
                ax.minimum_scale = spec[f"{k_}_min"]
            if spec.get(f"{k_}_max") is not None:
                ax.maximum_scale = spec[f"{k_}_max"]
        _legend(ch, False)
        return ch

    raise ValueError(f"unknown chart kind {kind}")


class _Raw:
    """Canvas wrapper that draws shapes without adding preview ops twice."""

    def __init__(self, cv):
        self.cv = cv

    def rect(self, *a, **k):
        return self.cv.rect(*a, **k)

    def text(self, *a, **k):
        return self.cv.text(*a, **k)


def _sparklines(slide, x, y, w, h, spec, cv):
    rows = spec["rows"]
    ncol = spec.get("ncol", 2)
    nrow = -(-len(rows) // ncol)
    ty = y
    if spec.get("title"):
        cv.text(x, y, w, 0.28, [spec["title"]], 12, NAVY, bold=True)
        ty += 0.34
    cw = (w - 0.2 * (ncol - 1)) / ncol
    rh = (h - (ty - y)) / nrow
    lab_w = spec.get("label_w", 1.45)
    for i, (lab, vals) in enumerate(rows):
        c, r = i // nrow, i % nrow
        xx, yy = x + c * (cw + 0.2), ty + r * rh
        cv.text(xx, yy + rh / 2 - 0.1, lab_w, 0.22, [lab], spec.get("label_size", 9), INK)
        cd = _cat_data([str(k + 1) for k in range(len(vals))], [(lab, vals)])
        ch = slide.shapes.add_chart(XL_CHART_TYPE.LINE, _e(xx + lab_w), _e(yy + 0.02), _e(cw - lab_w), _e(rh - 0.04), cd).chart
        ch.has_title = False
        ch.has_legend = False
        s = ch.plots[0].series[0]
        s.format.line.color.rgb = _rgb(NAVY)
        s.format.line.width = Pt(1.5)
        s.smooth = False
        hi = max(range(len(vals)), key=lambda k: vals[k])
        p = s.points[hi]
        p.marker.style = 8
        p.marker.size = 5
        p.marker.format.fill.solid()
        p.marker.format.fill.fore_color.rgb = _rgb(ORANGE)
        p.marker.format.line.fill.background()
        ch.value_axis.visible = False
        ch.value_axis.has_major_gridlines = False
        ch.category_axis.visible = False
        # remove plot-area padding: give the plot a manual layout filling the frame
        pa = ch._chartSpace.find(".//" + qn("c:plotArea"))
        lay = pa.find(qn("c:layout"))
        if lay is None:
            lay = etree.Element(qn("c:layout"))
            pa.insert(0, lay)
        ml = etree.fromstring(f'<c:manualLayout xmlns:c="{C}"><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>'
                              f'<c:x val="0.02"/><c:y val="0.08"/><c:w val="0.96"/><c:h val="0.84"/></c:manualLayout>')
        for old in list(lay):
            lay.remove(old)
        lay.append(ml)


def _dashboard(cv, x, y, w, h, spec):
    """A one-screen dashboard mock: a row of KPI cards, then charts side by side (spec['charts'], each a chart spec)."""
    kpis, charts = spec.get("kpis", []), spec.get("charts", [])
    yy = y
    if spec.get("banner"):
        cv.rect(x, yy, w, 0.36, NAVY, None)
        cv.text(x + 0.15, yy + 0.07, w - 0.3, 0.26, [spec["banner"]], 12, WHITE, bold=True)
        yy += 0.44
    if kpis:
        kh, gap = 0.82, 0.14
        kw = (w - gap * (len(kpis) - 1)) / len(kpis)
        for i, k in enumerate(kpis):
            lab, val = k[0], k[1]
            sub = k[2] if len(k) > 2 else ""
            col = k[3] if len(k) > 3 else NAVY
            kx = x + i * (kw + gap)
            cv.rect(kx, yy, kw, kh, "F3F5F8", "D5DBE2")
            cv.rect(kx, yy, 0.07, kh, col, None)
            cv.text(kx + 0.18, yy + 0.07, kw - 0.25, 0.2, [lab.upper()], 9, GREY, bold=True)
            cv.text(kx + 0.18, yy + 0.27, kw - 0.25, 0.32, [val], 17, col, bold=True)
            if sub:
                cv.text(kx + 0.18, yy + 0.58, kw - 0.25, 0.2, [sub], 9, GREY)
        yy += kh + 0.16
    if charts:
        gap = 0.2
        widths = spec.get("widths") or [1] * len(charts)
        tot = sum(widths)
        avail = w - gap * (len(charts) - 1)
        cx = x
        ch_h = y + h - yy
        for c, k in zip(charts, widths):
            cw = avail * k / tot
            cv.rect(cx, yy, cw, ch_h, WHITE, "D5DBE2")
            draw(_NoOps(cv), cx + 0.05, yy + 0.05, cw - 0.1, ch_h - 0.1, c)
            cx += cw + gap


class _NoOps:
    """Pass-through canvas whose chart placeholders are still recorded for previews."""

    def __init__(self, cv):
        self.cv, self.slide, self.ops = cv, cv.slide, cv.ops

    def rect(self, *a, **k):
        return self.cv.rect(*a, **k)

    def text(self, *a, **k):
        return self.cv.text(*a, **k)
