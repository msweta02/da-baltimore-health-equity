"""Build tableau/dashboard/Baltimore_Service_Need_Dashboard.twbx (Step 3).

Hand-authored TWB XML, modelled on the XML Tableau Public 2026.2 wrote for the existing
dashboard (T_Work/Baltimore Comm Health & Access.twbx). tableaudocumentapi cannot create
worksheets, dashboards or actions, so it is used only to re-open and check the result.

Tableau Public 2026.2 will not open a workbook whose data source is a live file
connection (it aborts the load with "Workbooks saved to Tableau Public must use
extracts" -- the existing T_Work dashboard fails the same way). So every table is
written into one Hyper file with Tableau's Hyper API, and the single data source
connects to it directly, the way Tableau's own World Indicators sample does.

Run from the project root after prepare_dashboard_data.py, with a Python that has
pandas, openpyxl and tableauhyperapi installed (pip install tableauhyperapi):
    python build_tableau_workbook.py
"""
import json
import os
import uuid
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd

SRC = "tableau"
OUT_DIR = "tableau/dashboard"
TWB_NAME = "Baltimore_Service_Need_Dashboard.twb"
TWBX = f"{OUT_DIR}/Baltimore_Service_Need_Dashboard.twbx"
XLSX = "Baltimore_Vital_Signs_Tableau.xlsx"
GEOJSON = "Community_Statistical_Areas_2020_8602756789985715764.geojson"
EXTRACT = "Data/Extracts/Service_Need.hyper"   # the only data file inside the .twbx

W_PX, H_PX = 1400, 1000               # design size used to lay out the zones
NAVY, SLATE, PALE, SECTION = "#2a3d5d", "#8fa6c4", "#ecf3f9", "#dde9f7"
TIER_COLORS = {"Highest": "#2a3d5d", "High": "#48638c", "Moderate": "#7089ad",
               "Lower": "#9fb3cc"}    # one hue, light->dark; passes the ordinal validator
TIER_ORDER = ["Highest", "High", "Moderate", "Lower"]
RENAMES = json.load(open(f"{SRC}/csa2020_renames.json"))


def uid():
    return "{" + str(uuid.uuid4()).upper() + "}"


def rid(n=28):
    return uuid.uuid4().hex[:n]


def el(tag, attrs=None, *children, text=None):
    e = ET.Element(tag, {k: str(v) for k, v in (attrs or {}).items() if v is not None})
    for c in children:
        if c is not None:
            e.append(c)
    if text is not None:
        e.text = text
    return e


def tstr(fmt_cols):
    """Tableau column datatype from a pandas dtype."""
    k = fmt_cols.kind
    return {"i": "integer", "f": "real", "b": "boolean"}.get(k, "string")


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ======================================================================================
# Data sources
# ======================================================================================
DERIV = {"none": "None", "sum": "Sum", "avg": "Avg", "ctd": "CountD", "usr": "User",
         "clct": "Collect", "min": "Min", "attr": "Attribute"}
KIND = {"nk": "nominal", "ok": "ordinal", "qk": "quantitative"}


class Field:
    def __init__(self, name, datatype, role, type_, caption=None, formula=None,
                 table=None, remote=None):
        self.name, self.datatype, self.role, self.type = name, datatype, role, type_
        self.caption, self.formula, self.table, self.remote = caption, formula, table, remote

    def column_xml(self):
        e = el("column", {"caption": self.caption, "datatype": self.datatype,
                          "name": f"[{self.name}]", "role": self.role, "type": self.type})
        if self.formula is not None:
            e.append(el("calculation", {"class": "tableau", "formula": self.formula}))
        return e


class Table:
    """One logical table: an Excel sheet, a CSV or the geojson."""

    def __init__(self, caption, kind, source, df=None, geo_cols=None, suffix="", types=None):
        self.caption, self.kind, self.source, self.df = caption, kind, source, df
        self.geo_cols, self.suffix, self.types = geo_cols, suffix, types or {}
        self.obj_id = caption.replace("_", "!") + "_" + uuid.uuid4().hex.upper()
        self.local = {}          # remote column -> local field name

    def columns(self):
        if self.geo_cols:
            return self.geo_cols
        return [(c, self.types.get(c, tstr(self.df[c].dtype))) for c in self.df.columns]

    def relation(self, conn_name):
        return el("relation", {"connection": conn_name, "name": self.caption,
                               "table": f"[Extract].[{self.caption}]", "type": "table"})


class DataSource:
    def __init__(self, caption):
        self.caption = caption
        self.name = "federated." + rid()
        self.tables, self.fields, self.calcs = [], {}, []
        self.relationships = []       # (left_expr, right_expr, left_table, right_table)
        self.palettes = []            # (instance, {value: color}) datasource-level colors
        self.aliases = {}             # field -> {member: alias}
        self.col_overrides = {}       # local name -> (role, type)
        self._n = 1000

    def add_table(self, t):
        self.tables.append(t)
        for cn, dt in t.columns():
            local = cn if cn not in self.fields else f"{cn}{t.suffix}"
            assert local not in self.fields, local
            t.local[cn] = local
            role, typ = (("measure", "quantitative") if dt in ("real", "integer")
                         else ("dimension", "nominal"))
            if dt == "spatial":
                role, typ = "measure", "nominal"
            self.fields[local] = Field(local, dt, role, typ, table=t, remote=cn)
        return t

    def override(self, local, role, typ, caption=None):
        f = self.fields[local]
        f.role, f.type = role, typ
        if caption:
            f.caption = caption
        self.col_overrides[local] = (role, typ)

    def set_caption(self, local, caption):
        f = self.fields[local]
        self.override(local, f.role, f.type, caption)

    def calc(self, caption, formula, datatype, role, typ):
        self._n += 1
        name = f"Calculation_{3889020000000000 + self._n}"
        f = Field(name, datatype, role, typ, caption=caption, formula=formula)
        self.fields[name] = f
        self.calcs.append(name)
        return name

    # -- XML ---------------------------------------------------------------------------
    def xml(self):
        ds = el("datasource", {"caption": self.caption, "inline": "true", "name": self.name,
                               "version": "18.1"})
        conn = el("connection", {"class": "federated"})
        named = el("named-connections")
        cname = "hyper." + rid()
        conn_of = {(t.kind, t.source): cname for t in self.tables}
        named.append(el("named-connection", {"caption": "Service_Need", "name": cname},
                        el("connection", {"authentication": "auth-none",
                                          "author-locale": "en_US", "class": "hyper",
                                          "dbname": EXTRACT, "default-settings": "yes",
                                          "port": "", "sslmode": "",
                                          "username": "tableau_internal_user"})))
        conn.append(named)

        rels = {t.caption: t.relation(conn_of[(t.kind, t.source)]) for t in self.tables}
        if len(self.tables) == 1:
            conn.append(rels[self.tables[0].caption])
        else:
            coll = el("relation", {"type": "collection"})
            for t in self.tables:
                coll.append(rels[t.caption])
            conn.append(coll)

        cols = el("cols")
        for local, f in sorted(self.fields.items()):
            if f.table is not None:
                cols.append(el("map", {"key": f"[{local}]",
                                       "value": f"[{f.table.caption}].[{f.remote}]"}))
        conn.append(cols)
        # Explicit column -> table ownership. Without it Tableau re-derives the mapping from
        # the Hyper schema and assigns suffixed names to the wrong tables.
        rtype = {"string": "130", "real": "5", "integer": "20", "boolean": "11", "spatial": "8"}
        agg = {"string": "Count", "real": "Sum", "integer": "Sum", "boolean": "Count",
               "spatial": "Collect"}
        recs = el("metadata-records")
        ordinal = 0
        for t in self.tables:
            for cn, dt in t.columns():
                r = el("metadata-record", {"class": "column"})
                for tag, val in [("remote-name", cn), ("remote-type", rtype[dt]),
                                 ("local-name", f"[{t.local[cn]}]"),
                                 ("parent-name", f"[{t.caption}]"), ("remote-alias", cn),
                                 ("ordinal", str(ordinal)), ("local-type", dt),
                                 ("aggregation", agg[dt]), ("contains-null", "true")]:
                    r.append(el(tag, text=val))
                recs.append(r)
                ordinal += 1
        conn.append(recs)
        ds.append(conn)
        ds.append(el("aliases", {"enabled": "yes"}))

        for fld, amap in self.aliases.items():
            c = el("column", {"datatype": "string", "name": fld, "role": "dimension",
                              "type": "nominal"})
            a = el("aliases")
            for k, v in amap.items():
                a.append(el("alias", {"key": k, "value": v}))
            c.append(a)
            ds.append(c)
        for local, f in sorted(self.fields.items()):
            if f.formula is not None or local in self.col_overrides:
                ds.append(f.column_xml())
        # A color assignment attaches to a datasource-level column-instance of the field;
        # without one Tableau silently falls back to its default palette.
        for field, _ in self.palettes:
            if field.startswith("[:"):
                continue
            deriv, name, kind = field.strip("[]").split(":")
            ds.append(el("column-instance", {"column": f"[{name}]",
                                             "derivation": DERIV[deriv], "name": field,
                                             "pivot": "key", "type": KIND[kind]}))
        # schema order: ... column*, extract?, layout?, style, semantic-values?, object-graph?
        ds.append(el("layout", {"dim-ordering": "alphabetic",
                                "measure-ordering": "alphabetic", "show-structure": "true"}))
        if self.palettes:
            st = el("style")
            rule = el("style-rule", {"element": "mark"})
            for field, cmap in self.palettes:
                enc = el("encoding", {"attr": "color", "field": field, "type": "palette"})
                for val, color in cmap.items():
                    enc.append(el("map", {"to": color}, el("bucket", text=val)))
                rule.append(enc)
            st.append(rule)
            ds.append(st)
        ds.append(el("semantic-values", None,
                     el("semantic-value", {"key": "[Country].[Name]",
                                           "value": '"United States"'})))

        og = el("object-graph")
        objs = el("objects")
        for t in self.tables:
            objs.append(el("object", {"caption": t.caption, "id": t.obj_id},
                           el("properties", {"context": ""},
                              t.relation(conn_of[(t.kind, t.source)]))))
        og.append(objs)
        if self.relationships:
            rs = el("relationships")
            for left, right, lt, rt in self.relationships:
                rs.append(el("relationship", None,
                             el("expression", {"op": "="}, el("expression", {"op": left}),
                                el("expression", {"op": right})),
                             el("first-end-point", {"object-id": lt.obj_id}),
                             el("second-end-point", {"object-id": rt.obj_id})))
            og.append(rs)
        ds.append(og)
        return ds


def read_xlsx(sheet):
    return pd.read_excel(f"{SRC}/{XLSX}", sheet_name=sheet)


def read_csv(name):
    return pd.read_csv(f"{SRC}/{name}")


def read_geo():
    feats = json.load(open(f"{SRC}/{GEOJSON}"))["features"]
    # The 2020 boundary names -> the CSA2010 key, using the same seven renames as the
    # cleaning notebook's CSA2020 column (and the existing dashboard's CASE). Done here
    # rather than as a calculated relationship because Tableau mis-planned map queries
    # through the calculated join. "Unassigned -- Jail" gets no CSA.
    to2010 = {v: k for k, v in RENAMES.items()}
    csas = set(pd.read_excel(f"{SRC}/{XLSX}", sheet_name="Service_Need_Profile").CSA)
    rows = []
    for f in feats:
        name = f["properties"]["CSA2020"]
        key = to2010.get(name, name)
        rows.append({"OBJECTID": f["properties"]["OBJECTID"], "SQMI": f["properties"]["SQMI"],
                     "CSA2020": name, "CSA": key if key in csas else None,
                     "Geometry": _wkt(f["geometry"])})
    geo = pd.DataFrame(rows)
    assert geo.CSA.notna().sum() == 55 and set(geo.CSA.dropna()) == csas
    return geo


def _wkt(geom):
    ring = lambda r: "(" + ", ".join(f"{x} {y}" for x, y, *_ in r) + ")"
    poly = lambda p: "(" + ", ".join(ring(r) for r in p) + ")"
    if geom["type"] == "Polygon":
        return "POLYGON " + poly(geom["coordinates"])
    return "MULTIPOLYGON (" + ", ".join(poly(p) for p in geom["coordinates"]) + ")"


# ---- main datasource: everything at CSA grain, as ONE extract table ---------------------
# The sources are named and suffixed as separate logical tables (so a collision such as
# "Median Household Income" in both 2023_Analysis and 2018_Analysis becomes
# "... (2023)" / "... (2018)"), then merged on CSA in Python. Tableau's relationship
# planner produced invalid queries for the filled maps when the geometry lived in a
# related table, so the extract carries one wide 55-row table instead.
staging = DataSource("staging")
t_prof = staging.add_table(Table("Service_Need_Profile", "excel", XLSX, read_xlsx("Service_Need_Profile")))
parts = [
    staging.add_table(Table("2023_Analysis", "excel", XLSX, read_xlsx("2023_Analysis"), suffix=" (2023)")),
    staging.add_table(Table("2018_Analysis", "excel", XLSX, read_xlsx("2018_Analysis"), suffix=" (2018)")),
    staging.add_table(Table("Prenatal_Care_Deep_Dive", "csv", "Prenatal_Care_Deep_Dive.csv",
                            read_csv("Prenatal_Care_Deep_Dive.csv"), suffix=" (prenatal)")),
    staging.add_table(Table("Population_2020", "csv", "csa_population_2020.csv",
                            read_csv("csa_population_2020.csv"), suffix=" (population)")),
    staging.add_table(Table("Boundaries_2020", "geo", GEOJSON, read_geo(), geo_cols=[
        ("OBJECTID", "integer"), ("SQMI", "real"), ("CSA2020", "string"), ("CSA", "string"),
        ("Geometry", "spatial")], suffix=" (geo)")),
]
wide = t_prof.df.rename(columns=t_prof.local)
for t in parts:
    key = t.local["CSA"]
    d = t.df.rename(columns=t.local)
    wide = wide.merge(d, left_on="CSA", right_on=key, how="left", validate="one_to_one")
    assert wide[key].notna().all(), f"{t.caption}: CSA without a match"
assert len(wide) == 55 and wide.Geometry.notna().all()
main = DataSource("Service Need (CSA)")
main.add_table(Table("CSA_Service_Need", "excel", "merged", wide,
                     types={"Geometry": "spatial", "OBJECTID": "integer"}))
M = main.fields
LE18 = "Life Expectancy"                                      # 2018_Analysis
HOUSING = "Percentage of Residential Properties with Housing Violations (Excluding Vacants)"   # 2023_Analysis
HOMICIDE = "Number of Gun-Related Homicides per 1,000 Residents"                          # 2023_Analysis
MHI23, MHI18 = "Median Household Income (2023)", "Median Household Income (2018)"
for f in (HOUSING, HOMICIDE, MHI23, MHI18, LE18, "Service Need Score", "Persistence_Count",
          "prenatal_care_pct", "income", "Total Population 2020", "persistence_count",
          "prenatal_rank"):
    assert f in M, f

main.palettes.append(("[none:Tier:nk]", {f'"{k}"': TIER_COLORS[k] for k in TIER_ORDER}))
main.set_caption("income", "Median income 2019 ($)")
main.set_caption(MHI18, "Median income 2018 ($)")
main.set_caption(MHI23, "Median income 2023 ($)")
main.set_caption(LE18, "Life expectancy (years)")
main.set_caption(HOUSING, "Housing violations (%)")
main.set_caption(HOMICIDE, "Gun homicides per 1,000")
main.set_caption("prenatal_care_pct", "Prenatal care (%)")
main.set_caption("Persistence_Count", "Years in worst 10")
# CSA selector: a workbook parameter, not a shared ("apply to worksheets") filter. Tableau
# Public dropped the data-source prefix from shared filters when the workbook was re-saved,
# breaking them on the next open; parameters belong to the workbook and survive saves.
CSA_PARAM_REF = "[Parameters].[Parameter CSA]"
CSA_ALL = "(All)"


def csa_param_column():
    c = el("column", {"caption": "CSA", "datatype": "string", "name": "[Parameter CSA]",
                      "param-domain-type": "list", "role": "measure", "type": "nominal",
                      "value": f'"{CSA_ALL}"'})
    c.append(el("calculation", {"class": "tableau", "formula": f'"{CSA_ALL}"'}))
    m = el("members")
    for v in [CSA_ALL] + sorted(wide.CSA):
        m.append(el("member", {"value": f'"{v}"'}))
    c.append(m)
    return c


def params_datasource():
    d = el("datasource", {"hasconnection": "false", "inline": "true", "name": "Parameters",
                          "version": "18.1"})
    d.append(el("aliases", {"enabled": "yes"}))
    d.append(csa_param_column())
    return d


C_CSAF = main.calc("CSA selected", f'{CSA_PARAM_REF} = "{CSA_ALL}" OR [CSA] = {CSA_PARAM_REF}',
                   "boolean", "dimension", "nominal")
C_IS_CSA = main.calc("Is a CSA polygon", '[CSA2020] <> "Unassigned -- Jail"',
                     "boolean", "dimension", "nominal")
C_LE_PERSIST = main.calc("Chronic (3+ of 8 LE years)", "[Persistence_Count] >= 3",
                         "boolean", "dimension", "nominal")
C_PN_PERSIST = main.calc("Prenatal persistent (3+ of 6)", "[persistence_count] >= 3",
                         "boolean", "dimension", "nominal")
C_LABEL = main.calc("Outlier label",
                    'IF [CSA (prenatal)] = "Orangeville/East Highlandtown" '
                    'OR [CSA (prenatal)] = "Cross-Country/Cheswolde" '
                    'THEN [CSA (prenatal)] END', "string", "dimension", "nominal")
C_YEAR = main.calc("KPI Analysis Year", '"2023"', "string", "dimension", "nominal")
C_HIGHEST = main.calc("KPI Highest Burden Community",
                      "{FIXED : MAX(IF [Service Need Score] = {FIXED : MAX([Service Need Score])} "
                      "THEN [CSA (2023)] END)}", "string", "dimension", "nominal")
C_MEDINC = main.calc("KPI Median Household Income (City)", f"MEDIAN([{MHI23}])",
                     "real", "measure", "quantitative")
C_POP = main.calc("KPI City Population", "SUM([Total Population 2020])",
                  "real", "measure", "quantitative")
C_POPSHARE = main.calc("KPI Highest-tier population share",
                       "SUM([Total Population 2020]) / MIN({FIXED : SUM([Total Population 2020])})",
                       "real", "measure", "quantitative")

# ---- tables without a CSA: their own data sources, reading the same Hyper file --------
# (Tableau Public accepts any number of data sources as long as each is an extract; it is
# live file connections it refuses.)
city = DataSource("Prenatal Trend (citywide)")
t_city = city.add_table(Table("prenatal_trend_citywide", "csv", "prenatal_trend_citywide.csv",
                              read_csv("prenatal_trend_citywide.csv")))
dom = DataSource("Domain Strength")
dom.add_table(Table("domain_strength", "csv", "domain_strength.csv",
                    read_csv("domain_strength.csv")))
bar = DataSource("Barrier Adjustment")
bar.add_table(Table("barrier_adjustment", "csv", "barrier_adjustment.csv",
                    read_csv("barrier_adjustment.csv")))
city.override("Year", "dimension", "ordinal")
city.set_caption("city_median", "City median (%)")
C_PERIOD = city.calc("Period", 'IF [is_break_year] THEN "2013-2016: likely collection change" '
                               'ELSE "Comparable years" END', "string", "dimension", "nominal")
C_SEGMENT = city.calc("Segment", 'IF [Year] <= 2012 THEN "2010-2012" '
                                 'ELSEIF [Year] <= 2016 THEN "2013-2016" ELSE "2017-2019" END',
                      "string", "dimension", "nominal")
city.palettes.append((f"[none:{C_PERIOD}:nk]", {'"Comparable years"': NAVY,
                                                '"2013-2016: likely collection change"': "#b3b9c2"}))

C_ISINC = dom.calc("Is Income", '[Domain] = "Income"', "boolean", "dimension", "nominal")
dom.palettes.append((f"[none:{C_ISINC}:nk]", {"true": NAVY, "false": SLATE}))

MN_ALONE = f'"[{bar.name}].[sum:CorrelationAlone:qk]"'
MN_ADJ = f'"[{bar.name}].[sum:CorrelationAdjusted:qk]"'
bar.aliases["[:Measure Names]"] = {MN_ALONE: "Correlation alone",
                                  MN_ADJ: "After adjusting for income"}
bar.palettes.append(("[:Measure Names]", {MN_ALONE: SLATE, MN_ADJ: NAVY}))

DATASOURCES = [main, city, dom, bar]


# ======================================================================================
# Worksheets
# ======================================================================================
def I(ds, field, deriv="none", kind="nk"):
    """A column-instance reference, e.g. [federated.x].[sum:Value:qk]."""
    return (ds, field, deriv, kind)


def iname(inst):
    ds, field, deriv, kind = inst
    return f"[{deriv}:{field}:{kind}]"


def full(inst):
    return f"[{inst[0].name}].{iname(inst)}"


CSA = I(main, "CSA")
CITY_YEAR = I(city, "Year", "none", "ok")


class Sheet:
    def __init__(self, name, title=None, subtitle=None):
        self.name, self.title, self.subtitle = name, title, subtitle
        self.uuid = uid()
        self.rows = self.cols = ""
        self.mark = "Automatic"
        self.enc = []                  # (encoding tag, inst, extra attrs)
        self.filters = []              # xml elements
        self.slices = []
        self.insts = []                # every instance used (for dependencies)
        self.ds = []
        self.sorts = []
        self.styles = []               # worksheet-level style-rule elements
        self.pane_styles = []
        self.label = None              # customized label runs
        self.trend = False
        self.is_map = False
        self.show_labels = None
        self.hide_tooltip = False

    def use(self, *insts):
        for i in insts:
            if i not in self.insts:
                self.insts.append(i)
            if i[0] not in self.ds:
                self.ds.append(i[0])
        return self

    def csa_filter(self):
        self.use(CSA)
        self.member_filter(I(main, C_CSAF), "true")
        self.uses_param = True

    def year_filter(self):
        self.use(CITY_YEAR)
        self.filters.append(el("filter", {"class": "categorical", "column": full(CITY_YEAR)},
                               el("groupfilter", {"function": "level-members",
                                                  "level": iname(CITY_YEAR),
                                                  "user:ui-enumeration": "all",
                                                  "user:ui-marker": "enumerate"})))
        self.slices.append(full(CITY_YEAR))

    def member_filter(self, inst, member):
        self.use(inst)
        self.filters.append(el("filter", {"class": "categorical", "column": full(inst)},
                               el("groupfilter", {"function": "member", "level": iname(inst),
                                                  "member": member,
                                                  "user:ui-domain": "database",
                                                  "user:ui-enumeration": "inclusive",
                                                  "user:ui-marker": "enumerate"})))
        self.slices.append(full(inst))

    def measure_names_filter(self, ds, members):
        mn = f"[{ds.name}].[:Measure Names]"
        gf = el("groupfilter", {"function": "union", "user:op": "manual"})
        for m in members:
            gf.append(el("groupfilter", {"function": "member", "level": "[:Measure Names]",
                                         "member": m}))
        self.filters.append(el("filter", {"class": "categorical", "column": mn}, gf))
        self.slices.append(mn)

    def sort(self, dim, by, direction):
        self.use(dim, by)
        self.sorts.append(el("computed-sort", {"column": full(dim), "direction": direction,
                                               "using": full(by)}))

    def fmt(self, inst, fmt):
        self.use(inst)
        self.styles.append(("cell", {"attr": "text-format", "field": full(inst), "value": fmt}))

    def xml(self):
        ws = el("worksheet", {"name": self.name})
        if self.title:
            ft = el("formatted-text")
            ft.append(el("run", {"bold": "true", "fontcolor": NAVY, "fontsize": "13"},
                         text=self.title))
            if self.subtitle:
                ft.append(el("run", {"fontcolor": "#3d3d3d", "fontsize": "11"},
                             text="\n" + self.subtitle))
            ws.append(el("layout-options", None, el("title", None, ft)))
        table = el("table")
        view = el("view")
        dss = el("datasources")
        for d in self.ds:
            dss.append(el("datasource", {"caption": d.caption, "name": d.name}))
        view.append(dss)
        if self.is_map:
            view.append(el("mapsources", None, el("mapsource", {"name": "Tableau"})))
        for d in self.ds:
            view.append(self._deps(d))
        for f in self.filters:
            view.append(f)
        for s in self.sorts:
            view.append(s)
        if self.slices:
            sl = el("slices")
            for s in dict.fromkeys(self.slices):
                sl.append(el("column", text=s))
            view.append(sl)
        view.append(el("aggregation", {"value": "true"}))
        table.append(view)

        style = el("style")
        rules = {}
        if self.mark != "Text" or not self.label:          # not the KPI cards
            # base font for every chart: darker and larger; cascades to axes, headers, labels
            rules["worksheet"] = [{"attr": "font-size", "value": getattr(self, "base_font", "11")},
                                  {"attr": "color", "value": "#262626"}]
        for elem, attrs in self.styles:
            rules.setdefault(elem, []).append(attrs)
        if self.is_map:
            rules.setdefault("map", []).append({"attr": "washout", "value": "0"})
        for elem, lst in rules.items():
            r = el("style-rule", {"element": elem})
            for a in lst:
                r.append(el("format", a))
            style.append(r)
        table.append(style)

        pane = el("pane", {"selection-relaxation-option": "selection-relaxation-allow"})
        pane.append(el("view", None, el("breakdown", {"value": "auto"})))
        pane.append(el("mark", {"class": self.mark}))
        encs = el("encodings")
        for tag, inst, extra in self.enc:
            encs.append(el(tag, dict({"column": inst if isinstance(inst, str) else full(inst)},
                                     **(extra or {}))))
        pane.append(encs)
        if self.label:
            ft = el("formatted-text")
            for attrs, txt in self.label:
                ft.append(el("run", attrs, text=txt))
            pane.append(el("customized-label", None, ft))
        if getattr(self, "refline", None):
            pane.append(el("reference-line", {
                "axis-column": full(self.cols_inst) if hasattr(self, "cols_inst") else self.cols,
                "enable-instant-analytics": "true", "formula": "average", "id": "refline0",
                "label-type": "none", "probability": "95",
                "scope": "per-table", "value-column": full(self.refline), "z-order": "1"}))
        if self.trend:
            pane.append(el("trendline", {"enable-confidence-bands": "false",
                                         "enable-instant-analytics": "true",
                                         "enabled": "true", "exclude-color": "true",
                                         "exclude-intercept": "false", "fit": "linear"}))
        pst = el("style")
        mark_rule = el("style-rule", {"element": "mark"})
        if self.show_labels is not None:
            mark_rule.append(el("format", {"attr": "mark-labels-show", "value": "true"}))
            mark_rule.append(el("format", {"attr": "mark-labels-cull",
                                           "value": "false" if self.show_labels == "all" else "true"}))
        for a in self.pane_styles:
            mark_rule.append(el("format", a))
        if len(mark_rule):
            pst.append(mark_rule)
        if self.mark == "Text" or self.label:
            pst.append(el("style-rule", {"element": "cell"},
                          el("format", {"attr": "text-align", "value": "center"})))
        if len(pst):
            pane.append(pst)
        table.append(el("panes", None, pane))
        table.append(el("rows", text=self.rows or None))
        table.append(el("cols", text=self.cols or None))
        if self.hide_tooltip:
            table.append(el("tooltip-style", {"tooltip-mode": "none"}))
        ws.append(table)
        ws.append(el("simple-id", {"uuid": self.uuid}))
        return ws

    def _deps(self, d):
        dep = el("datasource-dependencies", {"datasource": d.name})
        insts = [i for i in self.insts if i[0] is d]
        base = []

        def need(fname):
            if fname in base or fname not in d.fields:
                return
            base.append(fname)
            f = d.fields[fname]
            if f.formula:
                import re
                for ref in re.findall(r"\[([^\]]+)\]", f.formula):
                    need(ref)
        for i in insts:
            need(i[1])
        for fname in sorted(base):
            dep.append(d.fields[fname].column_xml())
        for i in insts:
            ds_, field, deriv, kind = i
            dep.append(el("column-instance", {"column": f"[{field}]",
                                              "derivation": DERIV[deriv],
                                              "name": iname(i), "pivot": "key",
                                              "type": KIND[kind]}))
        return dep


SHEETS = []


def sheet(*a, **k):
    s = Sheet(*a, **k)
    SHEETS.append(s)
    return s


# ---------------- KPI cards (Tab 1) --------------------------------------------------
def kpi(name, inst, caption, prefix="", fmt=None, extra=None):
    s = sheet(name)
    s.mark = "Text"
    s.use(inst)
    s.enc.append(("text", inst, None))
    big = "13" if name in ("KPI_HighestBurdenCommunity", "KPI_HighestTier") else "16"
    runs = [({"bold": "true", "fontcolor": NAVY, "fontsize": big},
             f"{prefix}<{full(inst)}>")]
    if extra:
        runs += extra(s)
    runs.append(({"fontcolor": "#333333", "fontsize": "10"}, "\n" + caption))
    s.label = runs
    s.show_labels = "all"
    s.hide_tooltip = True
    s.styles.append(("cell", {"attr": "height", "value": "110"}))
    s.styles.append(("cell", {"attr": "width", "value": "320" if name == "KPI_HighestTier" else "205"}))
    if fmt:
        s.fmt(inst, fmt)
    for elem, a in [("table", {"attr": "background-color", "value": "#00000000"}),
                    ("pane", {"attr": "band-color", "scope": "rows", "value": "#00000000"})]:
        s.styles.append((elem, a))
    return s


kpi("KPI_CSACount", I(main, "CSA", "ctd", "qk"), "Community Statistical Areas")
kpi("KPI_AnalysisYear", I(main, C_YEAR), "Analysis Year")
kpi("KPI_HighestBurdenCommunity", I(main, C_HIGHEST), "Highest Relative Burden")
kpi("KPI_MedianHouseholdIncome", I(main, C_MEDINC, "usr", "qk"),
    "Median Household Income (City, 2023)", prefix="$", fmt="n#,##0")
kpi("KPI_CityPopulation", I(main, C_POP, "usr", "qk"), "City Population (2020 Census)",
    fmt="n#,##0")


def _tier_extra(s):
    pop = I(main, "Total Population 2020", "sum", "qk")
    share = I(main, C_POPSHARE, "usr", "qk")
    s.use(pop, share)
    s.enc += [("text", pop, None), ("text", share, None)]
    s.fmt(pop, "n#,##0")
    s.fmt(share, "p0%")
    b = {"bold": "true", "fontcolor": NAVY, "fontsize": "13"}
    return [(b, " communities / "), (b, f"<{full(pop)}>"), (b, " residents ("),
            (b, f"<{full(share)}>"), (b, ")")]


k6 = kpi("KPI_HighestTier", I(main, "CSA", "ctd", "qk"), "in the Highest need tier",
         extra=_tier_extra)
k6.member_filter(I(main, "Tier"), '"Highest"')


# ---------------- maps ---------------------------------------------------------------
GEOM = I(main, "Geometry", "clct", "nk")
IS_CSA = I(main, C_IS_CSA)


def base_map(name, title, subtitle=None):
    s = sheet(name, title, subtitle)
    s.is_map = True
    s.use(GEOM, CSA)
    s.enc += [("lod", GEOM, None), ("lod", CSA, None), ("geometry", GEOM, None)]
    s.rows = f"[{main.name}].[Latitude (generated)]"
    s.cols = f"[{main.name}].[Longitude (generated)]"
    s.member_filter(IS_CSA, "true")
    s.csa_filter()
    return s


TIER = I(main, "Tier")
m1 = base_map("T1_Map_Tier", "Need tier by community",
              "Service_Need_Profile tier, 2023 conditions")
m1.use(TIER)
m1.enc.insert(0, ("color", TIER, None))

m2 = base_map("T1_Map_Persistence", "Chronic need: years in the worst 10 for life expectancy",
              "Years in the city's 10 lowest for life expectancy, of 8 (2011-2018). "
              "Darkest = all 8")
PERS = I(main, "Persistence_Count", "sum", "qk")
m2.use(PERS)
m2.enc.insert(0, ("color", PERS, None))
m2.styles.append(("mark", None))   # placeholder removed below
m2.styles.pop()

m3 = base_map("T3_Map_Prenatal", "Early prenatal care, 2019",
              "% of births with first-trimester care; centred on the city median (62.5%)")
PNPCT = I(main, "prenatal_care_pct", "avg", "qk")
m3.use(PNPCT)
m3.enc.insert(0, ("color", PNPCT, None))

# ---------------- Tab 1 bar: most persistent ---------------------------------------
b1 = sheet("T1_Bar_Persistence", "Lowest life expectancy, year after year (2011-2018)",
           "Years each community ranked in the city's 10 lowest, of 8. Communities with 3+ "
           "shown (two tie for 10th, so 11). Colored by need tier.")
b1.mark = "Bar"
b1.use(CSA, PERS, TIER)
b1.rows, b1.cols = full(CSA), full(PERS)
b1.enc += [("color", TIER, None), ("text", PERS, None)]
b1.show_labels = "all"
b1.sort(CSA, PERS, "DESC")
b1.member_filter(I(main, C_LE_PERSIST), "true")
b1.styles.append(("header", {"attr": "width", "field": full(CSA), "value": "200"}))
b1.csa_filter()

# ---------------- Tab 2 scatters -----------------------------------------------------
def scatter(name, title, x, y, subtitle=None, trend=False, yrange=None):
    s = sheet(name, title, subtitle)
    if yrange:
        s.styles.append(("__axis__", {"attr": "space", "class": "0", "field": full(y),
                                      "field-type": "quantitative", "max": str(yrange[1]),
                                      "min": str(yrange[0]), "range-type": "fixed",
                                      "scope": "rows", "type": "space"}))
    s.mark = "Circle"
    s.use(CSA, TIER, x, y)
    s.rows, s.cols = full(y), full(x)
    s.enc += [("color", TIER, None), ("lod", CSA, None)]
    s.trend = trend
    s.fmt(x, "c\"$\"#,##0")
    s.csa_filter()
    return s


scatter("T2_Scatter_LifeExpectancy", "Life expectancy (2018)",
        I(main, MHI18, "avg", "qk"), I(main, LE18, "avg", "qk"),
        "x: median household income, 2018", yrange=(60, 90))
scatter("T2_Scatter_Housing", "Housing violations, % of residential properties (2023)",
        I(main, MHI23, "avg", "qk"), I(main, HOUSING, "avg", "qk"),
        "x: median household income, 2023")
scatter("T2_Scatter_Homicide", "Gun-related homicides per 1,000 residents (2023)",
        I(main, MHI23, "avg", "qk"), I(main, HOMICIDE, "avg", "qk"),
        "x: median household income, 2023")
scatter("T2_Scatter_Prenatal", "Early prenatal care, % (2019)",
        I(main, "income", "avg", "qk"), I(main, "prenatal_care_pct", "avg", "qk"),
        "x: median household income, 2019", yrange=(30, 100))

# ---------------- Tab 2 domain bar ---------------------------------------------------
DOMAIN, STRENGTH = I(dom, "Domain"), I(dom, "TypicalStrength", "sum", "qk")
db = sheet("T2_Bar_DomainStrength", "What tracks life expectancy, by domain",
           "Typical strength = median |r| with life expectancy across the domain's measures, 2018")
dom.set_caption("TypicalStrength", "Typical strength (|r|)")
db.mark = "Bar"
db.use(DOMAIN, STRENGTH, I(dom, C_ISINC))
db.rows, db.cols = full(DOMAIN), full(STRENGTH)
db.enc += [("color", I(dom, C_ISINC), None), ("text", STRENGTH, None)]
db.show_labels = "all"
db.fmt(STRENGTH, "n0.00")
db.sort(DOMAIN, STRENGTH, "DESC")
db.styles.append(("header", {"attr": "width", "field": full(DOMAIN), "value": "175"}))

# ---------------- Tab 3 ranked bar: worst 10 and best 5 ------------------------------
# The map shows all 55; the bar names the extremes. Worst 10 matches the worst-10 logic of
# the persistence counts; best 5 is enough to show the size of the gap.
RANK = I(main, "prenatal_rank", "sum", "qk")
C_RGROUP = main.calc("Ranking group", 'IF [prenatal_rank] <= 10 THEN "Worst 10" '
                     'ELSEIF [prenatal_rank] >= 51 THEN "Best 5" END',
                     "string", "dimension", "nominal")
C_CITY19 = main.calc("City median 2019", "62.5", "real", "measure", "quantitative")
RGROUP, CITY19 = I(main, C_RGROUP), I(main, C_CITY19, "avg", "qk")
rb = sheet("T3_Bar_PrenatalRanked", "Worst 10 and best 5 communities (2019)",
           "% of births with first-trimester care. Colored as the map: orange below, blue above "
           "the city median (62.5%, the line)")
rb.mark = "Bar"
rb.use(CSA, PNPCT, RANK, RGROUP, CITY19)
rb.rows, rb.cols = f"({full(RGROUP)} / {full(CSA)})", full(PNPCT)
rb.enc += [("color", PNPCT, None), ("text", PNPCT, None), ("lod", CITY19, None)]
rb.show_labels = "all"
rb.fmt(PNPCT, "n0.0")
rb.sort(CSA, RANK, "ASC")
rb.filters.append(el("filter", {"class": "categorical", "column": full(RGROUP)},
                     el("groupfilter", {"function": "union", "user:op": "manual"},
                        el("groupfilter", {"function": "member", "level": iname(RGROUP),
                                           "member": '"Worst 10"'}),
                        el("groupfilter", {"function": "member", "level": iname(RGROUP),
                                           "member": '"Best 5"'}))))
rb.slices.append(full(RGROUP))
rb.manual_group_order = (RGROUP, ["Worst 10", "Best 5"])
rb.refline = CITY19
rb.styles.append(("header", {"attr": "width", "field": full(CSA), "value": "200"}))
rb.csa_filter()

# ---------------- Tab 3 scatter ------------------------------------------------------
sc = scatter("T3_Scatter_IncomePrenatal", "Income vs early prenatal care (2019)",
             I(main, "income", "avg", "qk"), I(main, "prenatal_care_pct", "avg", "qk"),
             "Labelled: the two communities furthest below what their income predicts",
             trend=True, yrange=(30, 100))
LBL = I(main, C_LABEL)
sc.use(LBL)
sc.enc.append(("text", LBL, None))   # labels are the text encoding
sc.show_labels = "all"

# ---------------- Tab 3 citywide trend -----------------------------------------------
CMED = I(city, "city_median", "avg", "qk")
ln = sheet("T3_Line_CitywideTrend", "City median early prenatal care, 2010-2019",
           "2013-2016: likely data collection change - do not compare levels across this span.")
ln.mark = "Line"
ln.use(CITY_YEAR, CMED, I(city, C_PERIOD), I(city, C_SEGMENT))
ln.cols, ln.rows = full(CITY_YEAR), full(CMED)
ln.enc += [("color", I(city, C_PERIOD), None), ("lod", I(city, C_SEGMENT), None),
           ("text", CMED, None)]
ln.show_labels = "some"
ln.fmt(CMED, "n0.0")
ln.pane_styles.append({"attr": "mark-markers-mode", "value": "all"})
ln.year_filter()

# ---------------- Tab 3 persistence table --------------------------------------------
PNPERS = I(main, "persistence_count", "sum", "qk")
tb = sheet("T3_Table_PrenatalPersistence", "Persistent prenatal-care gaps",
           "Years in the worst 10, of 6 comparable years (3+ shown)")
tb.mark = "Text"
tb.use(CSA, PNPERS)
tb.rows = full(CSA)
tb.enc.append(("text", PNPERS, None))
tb.sort(CSA, PNPERS, "DESC")
tb.member_filter(I(main, C_PN_PERSIST), "true")
tb.styles.append(("header", {"attr": "width", "field": full(CSA), "value": "200"}))
tb.csa_filter()

# ---------------- Tab 3 barrier paired bar -------------------------------------------
BARRIER = I(bar, "Barrier")
ALONE, ADJ = I(bar, "CorrelationAlone", "sum", "qk"), I(bar, "CorrelationAdjusted", "sum", "qk")
pb = sheet("T3_Bar_BarrierAdjustment", "Access barriers vs early prenatal care, 2019",
           "r across 55 CSAs. Light = barrier alone, dark = after adjusting for income")
pb.mark = "Bar"
pb.use(BARRIER, ALONE, ADJ)
MV = f"[{bar.name}].[Multiple Values]"
MNAMES = f"[{bar.name}].[:Measure Names]"
pb.rows = f"({full(BARRIER)} / {MNAMES})"
pb.cols = MV
pb.enc += [("color", MNAMES, None), ("text", MV, None)]
pb.show_labels = "all"
pb.styles.append(("cell", {"attr": "text-format", "field": MV, "value": "n+0.00;-0.00"}))
pb.measure_names_filter(bar, [MN_ALONE, MN_ADJ])
pb.styles.append(("label", {"attr": "display", "field": MNAMES, "value": "false"}))

# diverging color for the prenatal map, centred on the 2019 city median
m3.styles.append(("__encoding__", {"attr": "color", "center": "62.5", "field": full(PNPCT),
                                   "palette": "orange_blue_diverging_10_0",
                                   "type": "interpolated"}))
rb.styles.append(("__encoding__", {"attr": "color", "center": "62.5", "field": full(PNPCT),
                                   "palette": "orange_blue_diverging_10_0",
                                   "type": "interpolated"}))
m2.styles.append(("__encoding__", {"attr": "color", "field": full(PERS),
                                   "palette": "orange_10_0", "type": "interpolated"}))

# encodings inside a worksheet's style go in a 'mark' style-rule as <encoding>
_orig_xml = Sheet.xml


def _xml_with_encodings(self):
    if TIER in self.insts and not getattr(self, "_tier_done", False):
        self._tier_done = True
        self.styles.append(("__palette__", None))
    encs = [a for e, a in self.styles if e == "__encoding__"]
    axes = [a for e, a in self.styles if e == "__axis__"]
    self.styles = [(e, a) for e, a in self.styles if e != "__axis__"]
    pal = any(e == "__palette__" for e, _ in self.styles)
    self.styles = [(e, a) for e, a in self.styles if e != "__palette__"]
    self.styles = [(e, a) for e, a in self.styles if e != "__encoding__"]
    ws = _orig_xml(self)
    if axes:
        style = ws.find("table/style")
        rule = el("style-rule", {"element": "axis"})
        for a in axes:
            rule.append(el("encoding", a))
        style.insert(0, rule)
    if encs:
        style = ws.find("table/style")
        rule = el("style-rule", {"element": "mark"})
        for a in encs:
            rule.append(el("encoding", a))
        style.append(rule)
    mg = getattr(self, "manual_group_order", None)
    if mg:
        view = ws.find("table/view")
        ms = el("manual-sort", {"column": full(mg[0]), "direction": "ASC"})
        d = el("dictionary")
        for k in mg[1]:
            d.append(el("bucket", text=f'"{k}"'))
        ms.append(d)
        idx = [i for i, c in enumerate(view) if c.tag in ("slices", "aggregation")][0]
        view.insert(idx, ms)
    if pal:          # legend and sort order Highest -> Lower (colors: datasource level)
        view = ws.find("table/view")
        ms = el("manual-sort", {"column": full(TIER), "direction": "ASC"})
        d = el("dictionary")
        for k in TIER_ORDER:
            d.append(el("bucket", text=f'"{k}"'))
        ms.append(d)
        # schema: filters, then sorts, then slices
        idx = [i for i, c in enumerate(view) if c.tag in ("slices", "aggregation")][0]
        view.insert(idx, ms)
    return ws


def _xml_with_params(self):
    ws = _xml_with_encodings(self)
    if getattr(self, "uses_param", False):
        view = ws.find("table/view")
        view.find("datasources").append(el("datasource", {"name": "Parameters"}))
        dep = el("datasource-dependencies", {"datasource": "Parameters"})
        dep.append(csa_param_column())
        idx = [i for i, c in enumerate(view) if c.tag == "datasource-dependencies"][-1] + 1
        view.insert(idx, dep)
    return ws


Sheet.xml = _xml_with_params


# ======================================================================================
# Dashboards: a tiny layout engine producing Tableau's 0-100000 zone coordinates
# ======================================================================================
class Z:
    """A zone. kind: 'horz'/'vert' flow, or a leaf ('sheet','text','filter','color',
    'button','empty'). size = fixed px along the parent's axis, None = share the rest."""

    def __init__(self, kind, size=None, children=(), **kw):
        self.kind, self.size, self.children, self.kw = kind, size, list(children), kw


def H(*c, size=None, **kw):
    return Z("horz", size, c, **kw)


def V(*c, size=None, **kw):     # kw w= : flex weight within the parent
    return Z("vert", size, c, **kw)


def zstyle(bg=None, border=None, margin="4", pad=None):
    s = el("zone-style")
    if border:
        s.append(el("format", {"attr": "border-color", "value": border[0]}))
        s.append(el("format", {"attr": "border-style", "value": "solid"}))
        s.append(el("format", {"attr": "border-width", "value": border[1]}))
    else:
        s.append(el("format", {"attr": "border-color", "value": "#000000"}))
        s.append(el("format", {"attr": "border-style", "value": "none"}))
        s.append(el("format", {"attr": "border-width", "value": "0"}))
    s.append(el("format", {"attr": "margin", "value": margin}))
    if bg:
        s.append(el("format", {"attr": "background-color", "value": bg}))
    return s


class Dashboard:
    def __init__(self, name, tab_label, title):
        self.name, self.tab_label, self.title = name, tab_label, title
        self.uuid, self.win_uuid = uid(), uid()
        self._id = 0
        self.sheets = []

    def nid(self):
        self._id += 1
        return self._id

    def emit(self, z, x, y, w, h, axis=None):
        u = lambda v, tot: str(round(v * 100000 / tot))
        a = {"h": u(h, H_PX), "id": str(self.nid()), "w": u(w, W_PX),
             "x": u(x, W_PX), "y": u(y, H_PX)}
        if z.size is not None:
            a = dict({"fixed-size": str(z.size), "is-fixed": "true"}, **a)
        k = z.kind
        if k in ("horz", "vert"):
            a.update({"param": k, "type-v2": "layout-flow"})
            if z.kw.get("name"):
                a["friendly-name"] = z.kw["name"]
            e = el("zone", a)
            fixed = sum(c.size for c in z.children if c.size is not None)
            flex = [c for c in z.children if c.size is None]
            wsum = sum(c.kw.get("w", 1) for c in flex) or 1
            span = w if k == "horz" else h
            free = span - fixed
            pos = x if k == "horz" else y
            for c in z.children:
                s = c.size if c.size is not None else free * c.kw.get("w", 1) / wsum
                if k == "horz":
                    e.append(self.emit(c, pos, y, s, h, k))
                else:
                    e.append(self.emit(c, x, pos, w, s, k))
                pos += s
            e.append(zstyle(z.kw.get("bg"), z.kw.get("border"), z.kw.get("margin", "0")))
            return e
        if k == "sheet":
            a["name"] = z.kw["sheet"]
            if not z.kw.get("title", True):
                a["show-title"] = "false"
            e = el("zone", a)
            e.append(zstyle(z.kw.get("bg"), z.kw.get("border")))
            self.sheets.append(z.kw["sheet"])
            return e
        if k == "text":
            a.update({"forceUpdate": "true", "type-v2": "text"})
            e = el("zone", a)
            ft = el("formatted-text")
            for attrs, txt in z.kw["runs"]:
                ft.append(el("run", attrs, text=txt))
            e.append(ft)
            e.append(zstyle(z.kw.get("bg"), z.kw.get("border"), z.kw.get("margin", "4")))
            return e
        if k == "param":
            a.update({"mode": "compact", "param": CSA_PARAM_REF, "show-title": "false",
                      "type-v2": "paramctrl"})
            e = el("zone", a)
            e.append(zstyle())
            return e
        if k == "filter":
            # the control's own title is hidden: a text label beside it carries the name,
            # so label and dropdown share one baseline
            a.update({"mode": z.kw["mode"], "name": z.kw["sheet"], "param": z.kw["param"],
                      "show-title": "false",
                      "type-v2": "filter", "values": "database"})
            e = el("zone", a)
            e.append(zstyle())
            return e
        if k == "color":
            a.update({"name": z.kw["sheet"], "pane-specification-id": "0",
                      "param": z.kw["param"], "type-v2": "color"})
            e = el("zone", a)
            e.append(zstyle())
            return e
        if k == "button":
            a["type-v2"] = "dashboard-object"
            e = el("zone", a)
            b = el("button", {"action": f'tabdoc:goto-sheet window-id="{z.kw["target"]}"',
                              "button-type": "text"})
            vs = el("button-visual-state")
            vs.append(el("caption", text=z.kw["caption"]))
            vs.append(el("button-caption-font-style", {"fontcolor": "#ffffff",
                                                       "fontname": "Tableau Bold",
                                                       "fontsize": "10"}))
            vs.append(el("format", {"attr": "background-color", "value": z.kw["bg"]}))
            b.append(vs)
            e.append(b)
            e.append(zstyle())
            return e
        if k == "empty":
            a["type-v2"] = "empty"
            e = el("zone", a)
            e.append(zstyle())
            return e
        raise ValueError(k)

    def xml(self, content, filters, dash_deps):
        d = el("dashboard", {"enable-sort-zone-taborder": "true", "name": self.name})
        d.append(el("style"))
        d.append(el("size", {"sizing-mode": "automatic"}))   # fills the viewer's window
        dss = el("datasources")
        dss.append(el("datasource", {"name": "Parameters"}))
        for ds_ in dash_deps:
            dss.append(el("datasource", {"caption": ds_.caption, "name": ds_.name}))
        d.append(dss)
        pdep = el("datasource-dependencies", {"datasource": "Parameters"})
        pdep.append(csa_param_column())
        d.append(pdep)
        for ds_, insts in dash_deps.items():
            dep = el("datasource-dependencies", {"datasource": ds_.name})
            for i in insts:
                dep.append(ds_.fields[i[1]].column_xml())
            for i in insts:
                dep.append(el("column-instance", {"column": f"[{i[1]}]",
                                                  "derivation": DERIV[i[2]], "name": iname(i),
                                                  "pivot": "key", "type": KIND[i[3]]}))
            d.append(dep)

        buttons = [Z("button", 190, caption=db_.tab_label, target=db_.win_uuid,
                     bg=NAVY if db_ is self else SLATE) for db_ in DASHBOARDS]
        root = V(
            Z("text", 62, runs=[({"bold": "true", "fontalignment": "1", "fontcolor": "#ffffff",
                                  "fontsize": "18"},
                                 "Baltimore Community Service Need - " + self.title)],
              bg=NAVY),
            H(*filters, *([] if any(f.size is None for f in filters) else [Z("empty")]),
              *buttons, size=58, bg=PALE,
              border=("#ffffff", "3"), name="HC - Filter"),
            content,
            name="VC - Main", bg="#ffffff")
        zones = el("zones")
        outer = el("zone", {"h": "100000", "id": str(self.nid()), "type-v2": "layout-basic",
                            "w": "100000", "x": "0", "y": "0"})
        outer.append(self.emit(root, 0, 0, W_PX, H_PX))
        outer.append(zstyle(margin="0", bg="#ffffff"))
        zones.append(outer)
        d.append(zones)
        d.append(el("simple-id", {"uuid": self.uuid}))
        return d


def S(name, size=None, title=True, **kw):
    return Z("sheet", size, sheet=name, title=title, **kw)


def card(name):
    return Z("sheet", None, sheet=name, title=False, bg=PALE, border=("#ffffff", "3"))


def section(txt, note=None):
    runs = [({"bold": "true", "fontcolor": NAVY, "fontsize": "15"}, txt)]
    if note:          # data years: on the full-width section line, where there is room
        runs.append(({"bold": "true", "fontcolor": "#333333", "fontsize": "12"}, "      " + note))
    return Z("text", 38, runs=runs, bg=SECTION)


def callout(txt, size, big=False, w=1):
    return Z("text", size, w=w, runs=[({"bold": "true" if big else "false", "fontcolor": NAVY,
                                   "fontsize": "13" if big else "12"}, txt)],
             bg=PALE, border=(SLATE, "1"))


def filter_label(txt, size):
    return Z("text", size, runs=[({"bold": "true", "fontcolor": NAVY, "fontsize": "11"}, txt)],
             margin="8")


def csa_filter_zone(sheet_name):
    return Z("param", 250)          # sheet_name kept for call-site readability


def legend(sheet_name, inst, size=None):
    return Z("color", size, sheet=sheet_name, param=full(inst))


D1 = Dashboard("1. Where Need Concentrates", "1. Where Need", "Where Need Concentrates")
D2 = Dashboard("2. Income: The Common Thread", "2. Income", "Income: The Common Thread")
D4 = Dashboard("4. Prenatal Care: Income and Barriers", "4. Prenatal: Why",
               "Prenatal Care: Income, Barriers and Trend")
D3 = Dashboard("3. Prenatal Care: The Sharpest Symptom", "3. Prenatal: Gap",
               "Prenatal Care: The Sharpest Symptom")
DASHBOARDS = [D1, D2, D3, D4]

tab1 = V(section("1. Where need concentrates - and where it has persisted",
                 "Data years: need 2023 | life expectancy 2011-2018"),
         H(card("KPI_CSACount"), card("KPI_AnalysisYear"), card("KPI_HighestBurdenCommunity"),
           card("KPI_MedianHouseholdIncome"), card("KPI_CityPopulation"),
           Z("sheet", 330, sheet="KPI_HighestTier", title=False, bg=PALE,
             border=("#ffffff", "3")), size=118),
         H(S("T1_Map_Tier"),
           V(legend("T1_Map_Tier", TIER, 170), legend("T1_Map_Persistence", PERS, 150),
             Z("empty"), size=200),
           S("T1_Map_Persistence"),
           S("T1_Bar_Persistence", 480)),
         name="VC - Tab 1")

tab2 = V(section("2. Income: the common thread",
                 "Data years: 2018, 2019 or 2023 - each panel is labelled"),
         callout("Income is the only condition whose change over time tracked a change in "
                 "life expectancy (r = +0.49 vs others near 0).", 66, big=True),
         H(V(H(S("T2_Scatter_LifeExpectancy"), S("T2_Scatter_Housing")),
             H(S("T2_Scatter_Homicide"), S("T2_Scatter_Prenatal"))),
           V(legend("T2_Scatter_LifeExpectancy", TIER, 170), Z("empty"), size=165),
           V(S("T2_Bar_DomainStrength"),
             callout("These domains overlap heavily - no single one uniquely explains more "
                     "than ~3% of the life-expectancy gap the others don't already explain.",
                     185), size=500)),
         name="VC - Tab 2")

# Tab 3: WHERE the gap is (map <-> ranked bar highlight lives here); Tab 4: WHY / trend.
tab3 = V(section("3. Prenatal care: the sharpest symptom - where the gap is",
                 "Data year: 2019, the latest reliable year"),
         H(S("T3_Map_Prenatal"),
           V(legend("T3_Map_Prenatal", PNPCT, 150), Z("empty"), size=260),
           V(S("T3_Table_PrenatalPersistence"),
             callout("2019: early prenatal care ranges from 42.9% to 93.5% - a 50.6-point "
                     "spread. 21 of 55 communities are below 60%; in the worst, 57% of "
                     "mothers miss first-trimester care.", 235), size=360),
           S("T3_Bar_PrenatalRanked", 600)),
         name="VC - Tab 3")

tab4 = V(section("4. Prenatal care: income, barriers and the trend",
                 "Data years: 2019 (scatter, barriers) | 2010-2019 (trend)"),
         H(S("T3_Scatter_IncomePrenatal"),
           V(legend("T3_Scatter_IncomePrenatal", TIER, 150), Z("empty"), size=165),
           S("T3_Bar_BarrierAdjustment", 640)),
         H(S("T3_Line_CitywideTrend"), size=340),
         name="VC - Tab 4")

DASH_XML = [
    D1.xml(tab1, [filter_label("CSA", 60), csa_filter_zone("T1_Map_Tier")], {main: [CSA]}),
    D2.xml(tab2, [filter_label("CSA", 60), csa_filter_zone("T2_Scatter_Prenatal")], {main: [CSA]}),
    D3.xml(tab3, [filter_label("CSA", 60), csa_filter_zone("T3_Map_Prenatal")], {main: [CSA]}),
    D4.xml(tab4, [filter_label("CSA", 60), csa_filter_zone("T3_Scatter_IncomePrenatal"),
                  filter_label("Year", 60),
                  Z("filter", 200, mode="checkdropdown", sheet="T3_Line_CitywideTrend",
                    param=full(CITY_YEAR))], {main: [CSA], city: [CITY_YEAR]}),
]


# ---- actions: highlight CSA between the ranked bar and the map on Tab 3 --------------
def highlight(name, src, excl):
    a = el("action", {"caption": name, "name": f"[Action_{uuid.uuid4().hex[:8].upper()}]"})
    a.append(el("activation", {"auto-clear": "true", "type": "on-select"}))
    a.append(el("source", {"dashboard": D3.name, "type": "sheet", "worksheet": src}))
    cmd = el("command", {"command": "tsc:brush"})
    cmd.append(el("param", {"name": "exclude", "value": ",".join(excl)}))
    cmd.append(el("param", {"name": "field-captions", "value": "CSA"}))
    cmd.append(el("param", {"name": "target", "value": D3.name}))
    a.append(cmd)
    return a


others = [s for s in D3.sheets if s not in ("T3_Map_Prenatal", "T3_Bar_PrenatalRanked")]
ACTIONS = [highlight("Highlight CSA: ranked bar -> map", "T3_Bar_PrenatalRanked",
                     others + ["T3_Bar_PrenatalRanked"]),
           highlight("Highlight CSA: map -> ranked bar", "T3_Map_Prenatal",
                     others + ["T3_Map_Prenatal"])]


# ======================================================================================
# Assemble the workbook
# ======================================================================================
MANIFEST = ["AccessibleZoneTabOrder", "AnimationOnByDefault", "BasicButtonObject",
            "BasicButtonObjectTextSupport", "IntuitiveSorting", "IntuitiveSorting_SP2",
            "MapboxVectorStylesAndLayers", "MarkAnimation", "ObjectModelEncapsulateLegacy",
            "ObjectModelTableType", "RelationshipCalculations", "SchemaViewerObjectModel",
            "SetMembershipControl", "SheetIdentifierTracking", "SortTagCleanup",
            "SpatialFileLayer", "WindowsPersistSimpleIdentifiers", "ZoneBackgroundTransparency",
            "ZoneFriendlyName"]


def window_worksheet(s):
    w = el("window", {"class": "worksheet", "hidden": "true", "name": s.name})
    cards = el("cards")
    cards.append(el("edge", {"name": "left"}, el("strip", {"size": "160"},
                                                  el("card", {"type": "pages"}),
                                                  el("card", {"type": "filters"}),
                                                  el("card", {"type": "marks"}))))
    cards.append(el("edge", {"name": "top"},
                    el("strip", {"size": "2147483647"}, el("card", {"type": "columns"})),
                    el("strip", {"size": "2147483647"}, el("card", {"type": "rows"})),
                    el("strip", {"size": "31"}, el("card", {"type": "title"}))))
    w.append(cards)
    w.append(el("simple-id", {"uuid": uid()}))
    return w


def window_dashboard(d, active):
    a = {"class": "dashboard", "name": d.name}
    if active:
        a["maximized"] = "true"
    w = el("window", a)
    vps = el("viewpoints")
    for s in dict.fromkeys(d.sheets):
        zoom = "entire-view"
        vps.append(el("viewpoint", {"name": s}, el("zoom", {"type": zoom})))
    w.append(vps)
    w.append(el("active", {"id": "-1"}))
    w.append(el("simple-id", {"uuid": d.win_uuid}))
    return w


def build(active_dashboard=D1):
    wb = el("workbook", {"original-version": "18.1",
                         "source-build": "2026.2.1 (20262.26.0708.1337)",
                         "source-platform": "win", "version": "18.1",
                         "xmlns:user": "http://www.tableausoftware.com/xml/user"})
    man = el("document-format-change-manifest")
    for m in MANIFEST:
        man.append(el(m))
    wb.append(man)
    wb.append(el("preferences", None,
                 el("preference", {"name": "ui.encoding.shelf.height", "value": "24"}),
                 el("preference", {"name": "ui.shelf.height", "value": "26"})))
    dss = el("datasources")
    dss.append(params_datasource())
    for d in DATASOURCES:
        dss.append(d.xml())
    wb.append(dss)
    wb.append(el("mapsources", None, el("mapsource", {"name": "Tableau"})))
    acts = el("actions")
    for a in ACTIONS:
        acts.append(a)
    wb.append(acts)
    wss = el("worksheets")
    for s in SHEETS:
        wss.append(s.xml())
    wb.append(wss)
    dbs = el("dashboards")
    for d in DASH_XML:
        dbs.append(d)
    wb.append(dbs)
    wins = el("windows", {"source-height": "30"})
    for d in DASHBOARDS:
        wins.append(window_dashboard(d, d is active_dashboard))
    for s in SHEETS:
        wins.append(window_worksheet(s))
    wb.append(wins)
    ET.indent(wb, space="  ")
    body = ET.tostring(wb, encoding="unicode")
    return "<?xml version='1.0' encoding='utf-8' ?>\n\n" + body + "\n"


def write_hyper(path):
    """Every logical table as one table in schema Extract of a single Hyper file."""
    from tableauhyperapi import (Connection, CreateMode, HyperProcess, Inserter, SqlType,
                                 TableDefinition, TableName, Telemetry, escape_name)
    st = {"string": SqlType.text(), "real": SqlType.double(), "integer": SqlType.big_int(),
          "boolean": SqlType.bool(), "spatial": SqlType.geography()}
    conv = {"string": lambda v: None if pd.isna(v) else str(v),
            "real": lambda v: None if pd.isna(v) else float(v),
            "integer": lambda v: None if pd.isna(v) else int(v),
            "boolean": lambda v: None if pd.isna(v) else bool(v),
            "spatial": str}
    with HyperProcess(Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU,
                      parameters={"log_dir": os.path.dirname(os.path.abspath(path))}) as hp, \
            Connection(hp.endpoint, path, CreateMode.CREATE_AND_REPLACE) as con:
        con.catalog.create_schema_if_not_exists("Extract")
        for t in [t for d in DATASOURCES for t in d.tables]:
            cols = t.columns()
            td = TableDefinition(TableName("Extract", t.caption),
                                 [TableDefinition.Column(c, st[d]) for c, d in cols])
            con.catalog.create_table(td)
            fs = [conv[d] for _, d in cols]
            rows = [[f(v) for f, v in zip(fs, r)] for r in t.df[[c for c, _ in cols]]
                    .itertuples(index=False)]
            if any(d == "spatial" for _, d in cols):     # WKT text -> GEOGRAPHY on insert
                stage = [TableDefinition.Column(c, st[d] if d != "spatial" else SqlType.text())
                         for c, d in cols]
                maps = [c if d != "spatial" else
                        Inserter.ColumnMapping(c, f"CAST({escape_name(c)} AS GEOGRAPHY)")
                        for c, d in cols]
                ins = Inserter(con, td, maps, inserter_definition=stage)
            else:
                ins = Inserter(con, td)
            with ins:
                ins.add_rows(rows)
                ins.execute()
            n = con.execute_scalar_query(f"SELECT COUNT(*) FROM {td.table_name}")
            assert n == len(t.df), (t.caption, n)


def package(xml, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    hyper_tmp = os.path.join(os.path.dirname(path), "_extract_build.hyper")
    write_hyper(hyper_tmp)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(TWB_NAME, xml.encode("utf-8"))
        z.write(hyper_tmp, EXTRACT)
    os.remove(hyper_tmp)
    for leftover in os.listdir(os.path.dirname(path)):     # Hyper's own log files
        if leftover.startswith("hyperd") and leftover.endswith(".log"):
            os.remove(os.path.join(os.path.dirname(path), leftover))


if __name__ == "__main__":
    import sys
    active = {"1": D1, "2": D2, "3": D3, "4": D4}[sys.argv[1]] if len(sys.argv) > 1 else D1
    out = sys.argv[2] if len(sys.argv) > 2 else TWBX
    xml = build(active)
    package(xml, out)
    print(f"wrote {out}: {len(SHEETS)} worksheets, {len(DASHBOARDS)} dashboards, "
          f"{len(DATASOURCES)} datasources, {len(ACTIONS)} actions")
