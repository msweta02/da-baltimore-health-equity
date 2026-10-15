"""Build tableau/dashboard/Baltimore_Community_Health_Access.twbx -- the one-page
"Overview" dashboard from the design mockup, filled with real 2023 BNIA data.

A separate workbook: Baltimore_Service_Need_Dashboard.twbx is not touched. Reuses the
XML building blocks of build_tableau_workbook.py (data sources, worksheets, zones,
Hyper extract writer), so both workbooks are built the same way.

Run from the project root (needs pandas, openpyxl, tableauhyperapi):
    python build_overview_dashboard.py
"""
import os
import re
import sys
import uuid
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd

import build_tableau_workbook as B          # building blocks; its own build is not run
from build_tableau_workbook import (el, uid, I, full, iname, Field, Table, DataSource,
                                    Sheet, Z, H, V, zstyle, DERIV, KIND)

OUT = "tableau/dashboard/Baltimore_Community_Health_Access.twbx"
TWB = "Baltimore_Community_Health_Access.twb"
B.W_PX, B.H_PX = 1540, 1020                # design size the zones are laid out at
NAVY, INK, PANEL, PAGE = "#1f3350", "#1f3350", "#eef4fb", "#f6f8fb"
TEAL = "#3f8f8b"

# ======================================================================================
# Data (2023 cross-section, the same file every other output uses)
# ======================================================================================
a = pd.read_excel(f"{B.SRC}/{B.XLSX}", sheet_name="2023_Analysis")
hist = pd.read_excel(f"{B.SRC}/{B.XLSX}", sheet_name="Historical_Trends")
POV = "Percent of Family Households Living Below the Poverty Line"
CHILD = "Percent of Children Living Below the Poverty Line"
UNEMP = "Unemployment Rate"
VEH = "Percent of Households with No Vehicles Available"
NET = "Percent of Households with No Internet at Home"
EVICT = "Rate of Evictions per 1,000 Residents"
OD = "Number of Overdose Calls for Service per 1,000 Residents"
SCORE = "Service Need Score"


def need_bin(v):
    return ("Very High (80-100)" if v >= 80 else "High (60-79)" if v >= 60 else
            "Moderate-High (40-59)" if v >= 40 else "Moderate (20-39)" if v >= 20 else
            "Lower (0-19)")


BINS = ["Very High (80-100)", "High (60-79)", "Moderate-High (40-59)", "Moderate (20-39)",
        "Lower (0-19)"]
BIN_COLORS = ["#4b1d5c", "#9c3d77", "#e0703f", "#f2a65a", "#f7dc80"]   # dark -> light

geo = B.read_geo().dropna(subset=["CSA"])[["CSA", "Geometry"]]
csa = a[["CSA", SCORE, "Interpretation", POV, CHILD, UNEMP, VEH, NET, EVICT, OD]].copy()
csa["Need Index"] = csa[SCORE].map(need_bin)
csa["Need Rank"] = csa[SCORE].rank(ascending=False, method="first").astype(int)
csa["Eviction Rank"] = csa[EVICT].rank(ascending=False, method="first").astype(int)
csa["Map Label"] = csa.apply(lambda r: (f"{r.CSA}\n(Highest need)" if r["Need Rank"] == 1 else
                                        None),
                             axis=1)
csa = csa.merge(geo, on="CSA", how="left", validate="one_to_one")
assert len(csa) == 55 and csa.Geometry.notna().all()

# KPI + callout facts, computed so the text can never drift from the data
TOP = csa.nsmallest(1, "Need Rank").iloc[0]
OD_MAX = csa.loc[csa[OD].idxmax()]
EV = csa.nsmallest(5, "Eviction Rank")
EV_TOP = EV.iloc[0]
EV_HIGHEST_TIER = int((EV["Need Rank"] <= csa["Need Rank"].nsmallest(13).max()).sum())
R_POV_OD = csa[POV].corr(csa[OD])

drivers = pd.DataFrame([
    ("Poverty (family)", csa[POV].median(), "%"),
    ("Unemployment", csa[UNEMP].median(), "%"),
    ("No vehicle", csa[VEH].median(), "%"),
    ("Evictions /1,000", csa[EVICT].median(), ""),
    ("Overdose /1,000", csa[OD].median(), "")], columns=["Indicator", "Value", "u"])
drivers["Display"] = drivers.apply(lambda r: f"{r.Value:.1f}{r.u}", axis=1)
drivers["Order"] = range(1, len(drivers) + 1)
drivers = drivers.drop(columns="u")

IND = [("Poverty (family)", POV, "%"), ("Child poverty", CHILD, "%"),
       ("Unemployment", UNEMP, "%"), ("No vehicle", VEH, "%"), ("No internet", NET, "%"),
       ("Evictions /1,000", EVICT, ""), ("Overdose /1,000", OD, "")]
ind = pd.DataFrame([{"CSA": r.CSA, "Indicator": lab, "Order": k + 1,
                     "Display": f"{r[col]:.1f}{u}"}
                    for _, r in csa.iterrows() for k, (lab, col, u) in enumerate(IND)])

od = (hist[hist.Indicator == OD][["CSA", "Year", "Value"]].dropna()
      .sort_values(["CSA", "Year"]).reset_index(drop=True))
assert od.Year.min() == 2018 and od.CSA.nunique() == 55

HEAT = [("Internet", "Households With No Internet (Percentile)"),
        ("Vehicle", "Households With No Vehicle (Percentile)"),
        ("Jobless", "Unemployment Rate (Percentile)"),
        ("No diploma", "Adults Without HS Diploma (Percentile)")]
top10 = csa.nsmallest(10, "Need Rank")[["CSA", "Need Rank"]].merge(a, on="CSA")
heat = pd.DataFrame([{"CSA": r.CSA, "Need Rank": int(r["Need Rank"]), "Barrier": lab,
                      "Barrier Order": k + 1, "Percentile": round(100 * r[col], 0),
                      "Display": f"{100 * r[col]:.0f}%"}
                     for _, r in top10.iterrows() for k, (lab, col) in enumerate(HEAT)])

# ======================================================================================
# Data sources (one Hyper extract, one table per source -- the pattern that loads cleanly)
# ======================================================================================
PARAM = "Parameter 1"
PARAMREF = f"[Parameters].[{PARAM}]"


class ParamDS:
    """The workbook's Parameters data source: one string list parameter (Spotlight CSA)."""
    name, caption = "Parameters", None

    def __init__(self, members, default):
        self.members, self.default = members, default
        f = Field(PARAM, "string", "measure", "nominal", caption="Community")
        self.fields = {PARAM: f}

    def column(self):
        c = el("column", {"caption": "Community", "datatype": "string", "name": f"[{PARAM}]",
                          "param-domain-type": "list", "role": "measure", "type": "nominal",
                          "value": f'"{self.default}"'})
        c.append(el("calculation", {"class": "tableau", "formula": f'"{self.default}"'}))
        m = el("members")
        for v in self.members:
            m.append(el("member", {"value": f'"{v}"'}))
        c.append(m)
        return c

    def xml(self):
        d = el("datasource", {"hasconnection": "false", "inline": "true", "name": "Parameters",
                              "version": "18.1"})
        d.append(el("aliases", {"enabled": "yes"}))
        d.append(self.column())
        return d


params = ParamDS(sorted(csa.CSA), TOP.CSA)


def source(caption, table, df, types=None):
    d = DataSource(caption)
    d.add_table(Table(table, "excel", "derived", df, types=types or {}))
    return d


d_csa = source("Overview CSA 2023", "Overview_CSA", csa, {"Geometry": "spatial"})
d_drv = source("Overview Drivers", "Overview_Drivers", drivers)
d_ind = source("Overview Spotlight Indicators", "Overview_Indicators", ind)
d_od = source("Overview Overdose Trend", "Overview_Overdose", od)
d_heat = source("Overview Barrier Heatmap", "Overview_Heatmap", heat)
d_od.override("Year", "dimension", "ordinal")
SOURCES = [d_csa, d_drv, d_ind, d_od, d_heat]
d_csa.set_caption(POV, "Family poverty (%)")
d_csa.set_caption(OD, "Overdose calls per 1,000")
d_csa.set_caption(EVICT, "Evictions per 1,000")
d_od.set_caption("Value", "Overdose calls per 1,000")
d_heat.set_caption("Percentile", "Burden percentile")
d_drv.set_caption("Value", "Median value")

d_csa.palettes.append(("[none:Need Index:nk]", {f'"{b}"': c for b, c in zip(BINS, BIN_COLORS)}))
d_drv.palettes.append(("[none:Indicator:nk]", {f'"{r.Indicator}"': c for r, c in zip(
    drivers.itertuples(), ["#6f3f8f", "#3b7dd8", "#3fae9b", "#f0a33a", "#e8655a"])}))

SEL = {}
for d in (d_csa, d_ind, d_od):
    SEL[d] = d.calc("Is selected community", f"[CSA] = {PARAMREF}", "boolean", "dimension",
                    "nominal")
d_csa.palettes.append((f"[none:{SEL[d_csa]}:nk]", {"true": "#6f3f8f", "false": "#e3dcef"}))
C_EV5 = d_csa.calc("Top 5 evictions", "[Eviction Rank] <= 5", "boolean", "dimension", "nominal")
C_SPOT = d_csa.calc("Spotlight summary",
                    '[CSA] + ": " + [Interpretation] + " (Service Need Score " + '
                    'STR(ROUND([Service Need Score], 1)) + " of 100)."',
                    "string", "dimension", "nominal")

# ======================================================================================
# Worksheets
# ======================================================================================
SHEETS = []
PARAM_SHEETS = set()


def sheet(name, title=None, subtitle=None):
    s = Sheet(name, title, subtitle)
    SHEETS.append(s)
    return s


def map_sheet(name, title=None, subtitle=None):
    s = sheet(name, title, subtitle)
    s.is_map = True
    geom, c = I(d_csa, "Geometry", "clct", "nk"), I(d_csa, "CSA")
    s.use(geom, c)
    s.enc += [("lod", geom, None), ("lod", c, None), ("geometry", geom, None)]
    s.rows = f"[{d_csa.name}].[Latitude (generated)]"
    s.cols = f"[{d_csa.name}].[Longitude (generated)]"
    return s


def sel_filter(s, d):
    s.member_filter(I(d, SEL[d]), "true")
    PARAM_SHEETS.add(s.name)


NEED = I(d_csa, "Need Index")
m = map_sheet("OV_Map_NeedIndex", "Where Are Multiple Community Barriers Concentrated?",
              "Potential Service Need Index (higher = greater relative need)")
m.use(NEED, I(d_csa, "Map Label"))
m.enc += [("color", NEED, None), ("text", I(d_csa, "Map Label"), None)]
m.show_labels = "all"
m.manual = (NEED, BINS)

drv = sheet("OV_Bar_Drivers", "What's Driving the Need?",
            "Typical community (median of 55 CSAs), 2023. Percent unless per 1,000")
ind_i, val, ordr = I(d_drv, "Indicator"), I(d_drv, "Value", "sum", "qk"), I(d_drv, "Order", "sum", "qk")
drv.mark = "Bar"
drv.use(ind_i, val, ordr, I(d_drv, "Display", "attr", "nk"))
drv.rows, drv.cols = full(ind_i), full(val)
drv.enc += [("color", ind_i, None), ("text", I(d_drv, "Display", "attr", "nk"), None)]
drv.show_labels = "all"
drv.sort(ind_i, ordr, "ASC")
drv.styles.append(("header", {"attr": "width", "field": full(ind_i), "value": "190"}))

sm = map_sheet("OV_Spot_Map")
SELI = I(d_csa, SEL[d_csa])
sm.use(SELI)
sm.enc.insert(0, ("color", SELI, None))
sm.hide_tooltip = True
PARAM_SHEETS.add(sm.name)

si = sheet("OV_Spot_Indicators", "Key Indicators (2023)")
ii, io_ = I(d_ind, "Indicator"), I(d_ind, "Order", "sum", "qk")
si.mark = "Text"
si.use(ii, io_, I(d_ind, "Display", "attr", "nk"))
si.rows = full(ii)
si.enc.append(("text", I(d_ind, "Display", "attr", "nk"), None))
si.sort(ii, io_, "ASC")
si.styles.append(("header", {"attr": "width", "field": full(ii), "value": "190"}))
sel_filter(si, d_ind)

st = sheet("OV_Spot_Trend", "Overdose Calls Trend", "Per 1,000 residents, 2018-2023 (collected from 2018)")
yr, ov = I(d_od, "Year", "none", "ok"), I(d_od, "Value", "avg", "qk")
st.mark = "Line"
st.use(yr, ov)
st.cols, st.rows = full(yr), full(ov)
st.pane_styles.append({"attr": "mark-markers-mode", "value": "all"})
sel_filter(st, d_od)

sx = sheet("OV_Spot_Text")
sx.mark = "Text"
SPOT = I(d_csa, C_SPOT)
sx.use(SPOT)
sx.enc.append(("text", SPOT, None))
sx.label = [({"fontcolor": "#333333", "fontsize": "11"}, f"<{full(SPOT)}>")]
sx.show_labels = "all"
sx.hide_tooltip = True
sel_filter(sx, d_csa)

sc = sheet("OV_Scatter_PovertyOverdose", "Social Determinants & Health",
           f"Communities with higher poverty tend to have more overdose calls "
           f"(r = {R_POV_OD:+.2f}; association, not causation).")
x, y = I(d_csa, POV, "avg", "qk"), I(d_csa, OD, "avg", "qk")
sc.mark = "Circle"
sc.use(x, y, NEED, I(d_csa, "CSA"))
sc.rows, sc.cols = full(y), full(x)
sc.enc += [("color", NEED, None), ("lod", I(d_csa, "CSA"), None)]
sc.trend = True
sc.manual = (NEED, BINS)

ev = sheet("OV_Bar_Evictions", "Housing Instability", "Eviction rate (per 1,000 residents), top 5")
evc, evv = I(d_csa, "CSA"), I(d_csa, EVICT, "sum", "qk")
ev.mark = "Bar"
ev.use(evc, evv)
ev.rows, ev.cols = full(evc), full(evv)
ev.enc += [("color", evv, None), ("text", evv, None)]
ev.show_labels = "all"
ev.fmt(evv, "n0.0")
ev.sort(evc, evv, "DESC")
ev.member_filter(I(d_csa, C_EV5), "true")
ev.styles.append(("header", {"attr": "width", "field": full(evc), "value": "175"}))
ev.styles.append(("__encoding__", {"attr": "color", "field": full(evv),
                                   "palette": "purple_10_0", "type": "interpolated"}))

hm = sheet("OV_Heatmap_Barriers", "Access Barriers Heatmap",
           "10 highest-need communities. Percentile among 55 CSAs for lack of internet, "
           "lack of a vehicle, unemployment and no HS diploma (higher = greater burden)")
hc, hb = I(d_heat, "CSA"), I(d_heat, "Barrier")
hp, hd = I(d_heat, "Percentile", "avg", "qk"), I(d_heat, "Display", "attr", "nk")
hr, hbo = I(d_heat, "Need Rank", "sum", "qk"), I(d_heat, "Barrier Order", "sum", "qk")
hm.mark = "Square"
hm.use(hc, hb, hp, hd, hr, hbo)
hm.rows, hm.cols = full(hc), full(hb)
hm.enc += [("color", hp, None), ("text", hd, None)]
hm.show_labels = "all"
hm.sort(hc, hr, "ASC")
hm.sort(hb, hbo, "ASC")
hm.styles.append(("header", {"attr": "width", "field": full(hc), "value": "185"}))
hm.base_font = "9"
hm.styles.append(("__encoding__", {"attr": "color", "field": full(hp),
                                   "palette": "orange_gold_10_0", "type": "interpolated"}))


# ---- post-processing: parameter dependencies and manual legend order ------------------
def finish(s, ws):
    view = ws.find("table/view")
    if s.name in PARAM_SHEETS:
        view.find("datasources").append(el("datasource", {"name": "Parameters"}))
        dep = el("datasource-dependencies", {"datasource": "Parameters"})
        dep.append(params.column())
        idx = [i for i, c in enumerate(view) if c.tag == "datasource-dependencies"][-1] + 1
        view.insert(idx, dep)
    mo = getattr(s, "manual", None)
    if mo:
        ms = el("manual-sort", {"column": full(mo[0]), "direction": "ASC"})
        d = el("dictionary")
        for k in mo[1]:
            d.append(el("bucket", text=f'"{k}"'))
        ms.append(d)
        idx = [i for i, c in enumerate(view) if c.tag in ("slices", "aggregation")][0]
        view.insert(idx, ms)
    return ws


# ======================================================================================
# Dashboard
# ======================================================================================
class Page(B.Dashboard):
    def emit(self, z, x, y, w, h, axis=None):
        if z.kind == "param":
            u = lambda v, tot: str(round(v * 100000 / tot))
            a = {"h": u(h, B.H_PX), "id": str(self.nid()), "w": u(w, B.W_PX),
                 "x": u(x, B.W_PX), "y": u(y, B.H_PX), "mode": "compact",
                 "param": PARAMREF, "type-v2": "paramctrl", "show-title": "false"}
            if z.size is not None:
                a = dict({"fixed-size": str(z.size), "is-fixed": "true"}, **a)
            e = el("zone", a)
            e.append(zstyle())
            return e
        return super().emit(z, x, y, w, h, axis)

    def build(self, root):
        d = el("dashboard", {"enable-sort-zone-taborder": "true", "name": self.name})
        d.append(el("style"))
        d.append(el("size", {"sizing-mode": "automatic"}))
        dss = el("datasources")
        dss.append(el("datasource", {"name": "Parameters"}))
        d.append(dss)
        dep = el("datasource-dependencies", {"datasource": "Parameters"})
        dep.append(params.column())
        d.append(dep)
        zones = el("zones")
        outer = el("zone", {"h": "100000", "id": str(self.nid()), "type-v2": "layout-basic",
                            "w": "100000", "x": "0", "y": "0"})
        outer.append(self.emit(root, 0, 0, B.W_PX, B.H_PX))
        outer.append(zstyle(margin="0", bg=PAGE))
        zones.append(outer)
        d.append(zones)
        d.append(el("simple-id", {"uuid": self.uuid}))
        return d


def T(size, runs, bg=None, border=None, margin="6", w=1):
    return Z("text", size, runs=runs, bg=bg, border=border, margin=margin, w=w)


def P(*paras, size=None, bg=None, border=None, margin="4", w=1):
    """Stack of single-paragraph text zones: (txt, fontsize, bold, color, italic, height)."""
    kids = [T(p[5] if len(p) > 5 else None, [r(p[0], p[1], p[2] if len(p) > 2 else False,
                                               p[3] if len(p) > 3 else "#333333",
                                               p[4] if len(p) > 4 else False)],
              bg=bg, margin=margin) for p in paras]
    return V(*kids, size=size, bg=bg, border=border, w=w)


def r(txt, size=11, bold=False, color="#333333", italic=False):
    a = {"fontcolor": color, "fontsize": str(size)}
    if bold:
        a["bold"] = "true"
    if italic:
        a["italic"] = "true"
    return (a, txt)


def S(name, size=None, title=True, **kw):
    return Z("sheet", size, sheet=name, title=title, **kw)


CARD = dict(bg="#ffffff", border=("#dfe7f1", "1"))


def kpi(value, caption, size=None):
    return P((value, 22, True, INK, False, 46), (caption, 10, False, "#333333", False, 80),
             size=size)


page = Page("Overview", "Overview", "Overview")

sidebar = V(T(30, [r("", 4)], bg=NAVY),
            Z("button", 46, caption="Overview", target=page.win_uuid, bg=TEAL),
            Z("empty", bg=NAVY),
            P(("Data source:", 10, True, "#ffffff", False, 26),
              ("BNIA-JFI Vital Signs (2023)", 10, False, "#ffffff", False, 64),
              ("55 Community Statistical Areas (CSAs)", 10, False, "#ffffff", False, 90),
              ("Note:", 10, True, "#ffffff", False, 26),
              ("Highlights key indicators, not all 200+ variables in the dataset. "
               "Correlations compare communities, not individuals.", 10, False, "#ffffff",
               False, 330),
              size=540, bg=NAVY),
            size=165, bg=NAVY, name="Sidebar")

header = H(P(("Baltimore Community Health & Access", 24, True, INK, False, 50),
             ("Understanding where community conditions create the greatest need for "
              "healthcare navigation and patient advocacy.", 13)),
           P(("Stronger communities.", 13, False, INK, True, 44),
             ("Healthier futures.", 13, False, INK, True), size=330, bg="#e3ecf8", margin="10"),
           size=104, bg="#ffffff", name="Header")

kpis = H(H(kpi("55", "Community Statistical Areas (CSAs)"),
           kpi("2023", "Analysis Year"),
           kpi("10", "Communities with the highest combined burden (see heatmap)"),
           kpi(f"{OD_MAX[OD]:.1f}", f"Highest overdose calls per 1,000 ({OD_MAX.CSA})"),
           bg=PANEL, border=("#ffffff", "4")),
         P(("Key Takeaway", 13, True, INK, False, 30),
           ("Need concentrates in a few communities, from overlapping barriers - not one "
            "factor.", 10),
           size=420, bg=PANEL, border=("#ffffff", "4")),
         size=132, name="KPIs")

evict_note = (f"Evictions flag a partly different set of places: {EV_TOP.CSA} has the "
              f"highest rate ({EV_TOP[EVICT]:.1f}) but ranks {EV_TOP['Need Rank']} of 55 for "
              f"overall need; {EV_HIGHEST_TIER} of these 5 are in the highest-need tier.")

left = V(H(H(S("OV_Map_NeedIndex"), V(Z("color", 240, sheet="OV_Map_NeedIndex",
                                        param=full(NEED)), Z("empty"), size=265),
             w=1.6, **CARD),
           H(S("OV_Bar_Drivers"), **CARD), size=430),
         H(V(S("OV_Scatter_PovertyOverdose"),
             T(64, [r("Each dot is a CSA. The trend line shows the general relationship "
                      "between poverty and overdose calls (not a causal link).", 9)]),
             w=1.15, **CARD),
           V(S("OV_Bar_Evictions"), T(110, [r(evict_note, 10)], bg="#fdf7ea",
                                      border=("#f3e3bd", "1"), margin="8"), w=0.85, **CARD),
           V(S("OV_Heatmap_Barriers"),
             Z("color", 46, sheet="OV_Heatmap_Barriers", param=full(hp)), w=1.95, **CARD),
           w=1))

spot = V(T(40, [r("Community Spotlight", 15, True, INK)]),
         Z("param", 34),
         S("OV_Spot_Map", 150, title=False),
         S("OV_Spot_Indicators", 240),
         S("OV_Spot_Trend"),
         S("OV_Spot_Text", 120, title=False, bg="#fdf7ea", border=("#f3e3bd", "1")),
         size=330, **CARD)

root = H(sidebar,
         V(header, kpis, H(left, spot), bg=PAGE, name="Main"),
         name="Root")
DASH = page.build(root)


# ======================================================================================
# Assemble and package
# ======================================================================================
def build():
    wb = el("workbook", {"original-version": "18.1",
                         "source-build": "2026.2.1 (20262.26.0708.1337)",
                         "source-platform": "win", "version": "18.1",
                         "xmlns:user": "http://www.tableausoftware.com/xml/user"})
    man = el("document-format-change-manifest")
    for m_ in B.MANIFEST:
        man.append(el(m_))
    wb.append(man)
    dss = el("datasources")
    dss.append(params.xml())
    for d in SOURCES:
        dss.append(d.xml())
    wb.append(dss)
    wb.append(el("mapsources", None, el("mapsource", {"name": "Tableau"})))
    wss = el("worksheets")
    for s in SHEETS:
        wss.append(finish(s, s.xml()))
    wb.append(wss)
    wb.append(el("dashboards", None, DASH))
    wins = el("windows", {"source-height": "30"})
    w = el("window", {"class": "dashboard", "maximized": "true", "name": page.name})
    vps = el("viewpoints")
    for s in page.sheets:
        vps.append(el("viewpoint", {"name": s}, el("zoom", {"type": "entire-view"})))
    w.append(vps)
    w.append(el("active", {"id": "-1"}))
    w.append(el("simple-id", {"uuid": page.win_uuid}))
    wins.append(w)
    for s in SHEETS:
        wins.append(B.window_worksheet(s))
    wb.append(wins)
    ET.indent(wb, space="  ")
    return "<?xml version='1.0' encoding='utf-8' ?>\n\n" + ET.tostring(wb, encoding="unicode") + "\n"


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else OUT
    xml = build()
    B.DATASOURCES = SOURCES                 # the shared writer extracts these tables
    os.makedirs(os.path.dirname(out), exist_ok=True)
    tmp = os.path.join(os.path.dirname(out), "_overview_build.hyper")
    B.write_hyper(tmp)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(TWB, xml.encode("utf-8"))
        z.write(tmp, B.EXTRACT)
    os.remove(tmp)
    for f in os.listdir(os.path.dirname(out)):
        if f.startswith("hyperd") and f.endswith(".log"):
            os.remove(os.path.join(os.path.dirname(out), f))
    print(f"wrote {out}: {len(SHEETS)} worksheets, 1 dashboard, {len(SOURCES)} data sources")
    print(f"KPI: highest overdose {OD_MAX[OD]:.1f} ({OD_MAX.CSA}); top need {TOP.CSA}")
    print("callout:", evict_note)
