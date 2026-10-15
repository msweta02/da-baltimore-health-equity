# Tableau Build Steps
### Baltimore Community Health & Access
Built against `Baltimore_Vital_Signs_Tableau.xlsx`

---

## Step 1 · Connect the data

**Data → New Data Source → Microsoft Excel**, select the workbook. Three sheets appear.

| Sheet | Grain | Used by |
|---|---|---|
| `2023_Analysis` | 55 rows, 1 per CSA, 125 columns | Panels 1, 2, 3, and profile side of 4 |
| `2018_Analysis` | 55 rows, 123 columns, has `Life Expectancy` | optional health-outcome evidence |
| `Historical_Trends` | 95,437 rows, CSA × Indicator × Year | the Trends chart in panel 4 only |

Drag **`2023_Analysis`** onto the canvas as your primary source. Then drag
`Historical_Trends` in as a **separate data source** — do not relate them. Only one
worksheet uses it, and relating two sheets of different grain creates duplicate rows.

**Two things to verify before you go further.**

1. The sheet has **55 rows and no `Baltimore City` row.** Confirm with
   Data → View Data. This is correct and deliberate, but it means every "Baltimore
   City" benchmark on the dashboard must be *computed*, not read from a row. Step 2.4
   handles that.
2. `Historical_Trends` stacks 191 indicators with different units into one `Value`
   column. **Any sheet built on it must have an `Indicator` filter applied first**, or
   Tableau will sum dollars, percentages and rates together.

---

## Step 2 · Calculated fields

Create all of these before building any worksheet. Analysis → Create Calculated Field.
Field names below are copied exactly from the workbook — including the capital **W**
in `With Less Than`, which is easy to mistype.

### 2.1 Barrier component z-scores

`{FIXED : ...}` computes across all 55 CSAs regardless of filters, which is what you
want for a stable index.

```
// z No Diploma
([Percent Population (25 Years and over) With Less Than a High School Diploma or GED]
 - {FIXED : AVG([Percent Population (25 Years and over) With Less Than a High School Diploma or GED])})
/ {FIXED : STDEV([Percent Population (25 Years and over) With Less Than a High School Diploma or GED])}
```

Repeat for three more, changing only the field:

- `z No Internet` → `Percent of Households with No Internet at Home`
- `z No Vehicle` → `Percent of Households with No Vehicles Available`
- `z Low Income` → `Percent of Households Earning Less than $25,000`

### 2.2 Barrier Burden Index

```
// Barrier Burden Index
([z No Diploma] + [z No Internet] + [z No Vehicle] + [z Low Income]) / 4
```

### 2.3 Caregiving Profile

Classifies each CSA by which care population dominates. Thresholds are the city
medians, computed as LODs since there is no city row.

```
// Caregiving Profile
IF   [Percent of Population 65 Years and over]
     >= {FIXED : MEDIAN([Percent of Population 65 Years and over])} + 3
THEN "Elder"
ELSEIF [Percent of Population 65 Years and over]
     >= {FIXED : MEDIAN([Percent of Population 65 Years and over])}
THEN "Elder-leaning"
ELSEIF [Percent of Households with Children Under 18]
     >= {FIXED : MEDIAN([Percent of Households with Children Under 18])} + 8
THEN "Family"
ELSEIF [Percent of Households with Children Under 18]
     >= {FIXED : MEDIAN([Percent of Households with Children Under 18])}
THEN "Family-leaning"
ELSE "Mixed"
END
```

The `+ 3` and `+ 8` offsets separate "dominant" from "leaning". Adjust them once you
see the distribution — they are a presentation choice, not a statistical one, and you
should be ready to say so.

### 2.4 City benchmarks for panel 1

There is no Baltimore City row, so each benchmark is a median across the 55 CSAs.
Create one per indicator shown in "What's Driving the Need?":

```
// City Median · No Vehicle
{FIXED : MEDIAN([Percent of Households with No Vehicles Available])}
```

Repeat for: `Percent of Children Living Below the Poverty Line`, `Unemployment Rate`,
`Rate of Evictions per 1,000 Residents`,
`Number of Overdose Calls for Service per 1,000 Residents`.

**Label these "City Median", not "Baltimore City"** on the dashboard. A median of 55
areas is not the same as a citywide rate, and calling it one is a factual error someone
will catch.

### 2.5 Map colour bins

`Service Need Score` already exists in the sheet. Bin it to match the reference legend:

```
// Need Index Bin
IF [Service Need Score] >= 80 THEN "> 80"
ELSEIF [Service Need Score] >= 60 THEN "60 – 80"
ELSEIF [Service Need Score] >= 40 THEN "40 – 60"
ELSEIF [Service Need Score] >= 20 THEN "20 – 40"
ELSE "< 20"
END
```

### 2.6 CSA selector parameter

Analysis → Create Parameter. Name `p.Selected CSA`, type String, Allowable values
**List**, then **Add values from → Field → CSA**. Default to any CSA.

```
// Is Selected CSA
[CSA] = [p.Selected CSA]
```

---

## Step 3 · Build the worksheets

Nine sheets. Name them exactly as below so the dashboard assembly step makes sense.

### `S1_KPI` — the six stat boxes
1. Marks → **Text**
2. Drag `CSA` to Text → change aggregation to **Count Distinct** (gives 55)
3. For the text KPIs, use ad-hoc calcs, e.g.
   `{FIXED : MAX(IF [Service Need Score] = {FIXED : MAX([Service Need Score])} THEN [CSA] END)}`
4. Format → remove all borders, gridlines, row and column dividers
5. Hide the sheet title

Make this **one sheet per KPI** (six sheets) rather than one wide sheet — it gives you
control over spacing in the container. Duplicate `S1_KPI` five times.

### `S2_Map` — Potential Service Need Index by CSA
1. You need CSA polygons. Download the **Community Statistical Areas** boundary file
   (shapefile or GeoJSON) from the BNIA Open Data portal.
2. Data → New Data Source → **Spatial file**, then relate it to `2023_Analysis` on
   `CSA` = the shapefile's CSA name field. Use a **relationship**, not a join.
3. Double-click the shapefile's **Geometry** field
4. Colour ← `Need Index Bin`. Edit Colours → assign the five hexes in Step 5
5. Border white, 1pt. Map → Background Maps → **None**
6. Tooltip: CSA, Service Need Score, Caregiving Profile, Total Number of Households
7. Expect 2–3 CSA name mismatches. Fix with right-click → **Group**, or aliases.

### `S3_Drivers` — What's Driving the Need?
1. Rows ← `Measure Names`, filtered to the five driver measures
2. Columns ← `Measure Values`
3. Marks → Bar
4. Filter ← `Is Selected CSA` = True
5. Add the five `City Median · …` fields to Columns as a **second axis**, Marks → Bar,
   colour `#C9D3E0`
6. **Right-click either axis → Synchronize Axis.** Without this the grey benchmark bars
   sit on their own scale and the comparison is meaningless — this is the single most
   common mistake in rebuilding this panel.
7. Label both bars, end of bar, 1 decimal

### `S4_CareScatter` — panel 2, the replacement
1. Columns ← `Percent of Population 65 Years and over` (Dimension, or AVG)
2. Rows ← `Barrier Burden Index`
3. Detail ← `CSA`
4. Size ← `Percent of Households with Children Under 18`
5. Colour ← `Caregiving Profile`
6. Analytics pane → **Constant Line** on X at the value of
   `{FIXED : MEDIAN([Percent of Population 65 Years and over])}` (15.4), and on Y at `0`
7. Format both reference lines: `#8FA6C4`, dashed, label "City Median (65+ 15.4%)" and
   "City Median Barrier Burden"
8. Label only the top five by `Barrier Burden Index` using a rank table calc:
   `IF RANK([Barrier Burden Index], 'desc') <= 5 THEN [CSA] END` computed along `CSA`
9. Edit Size → range roughly 60 to 400 so the largest bubbles read clearly
10. Sheet title: **"Older Adults vs. Barrier Burden (size = Households with Children)"**

### `S5_PovertyOverdose` — panel 3A
1. Columns ← `Percent of Family Households Living Below the Poverty Line`
2. Rows ← `Number of Overdose Calls for Service per 1,000 Residents`
3. Detail ← `CSA`, Colour ← `Need Index Bin`
4. Analytics → **Trend Line**, linear
5. Put the r value in the sheet title, not the tooltip

### `S6_Evictions` — panel 3B
1. Rows ← `CSA`, Columns ← `Rate of Evictions per 1,000 Residents`
2. Filter → Top 5 by that field, sort descending
3. Colour `#8E5FA8`, labels on, **hide the axis** once labelled

### `S7_BarrierHeatmap` — panel 3C
1. Rows ← `CSA` (filtered to your five communities)
2. Columns ← `Measure Names`, filtered to the four barrier percentages
3. Marks → **Square**, Colour ← `Measure Values`, Label ← `Measure Values`
4. **Colour → Edit Colours → tick "Use Separate Legends"** — without this, a 58%
   no-vehicle and a 30% no-diploma sit on one scale and the whole grid reads dark
5. Add a `City Median` row at the bottom so a reader can tell what is unusual

### `S8_Profile` — panel 4 Key Indicators
1. Rows ← `Measure Names` filtered to your profile indicators
2. Text ← `Measure Values`
3. Filter ← `Is Selected CSA` = True
4. Marks → Text, all borders off

### `S9_Trend` — panel 4 Trends
1. **Switch to the `Historical_Trends` data source**
2. **Filter `Indicator` first** — set it to `Percent of Population 65 Years and over`,
   then right-click → **Apply to Worksheets → This worksheet only**
3. Columns ← `Year` (continuous), Rows ← `Value` (AVG), Detail ← `CSA`
4. Colour all lines `#D6DEE9`, then highlight the selected CSA in `#4A7EBB` using
   `Is Selected CSA` on Colour
5. Annotate the 2020 point: **"estimation discontinuity, not a real jump"**

---

## Step 4 · Assemble the dashboard

**Dashboard → New Dashboard → Size → Fixed, 1600 × 1100.**

Container tree, all **Tiled**:

```
Vertical (outer)
├── Horizontal · HEADER                       height 90
├── Horizontal · BODY
│   ├── Vertical · NAV RAIL                   width 190
│   │   ├── Text: "Dashboard Pages" + 5 items
│   │   └── Vertical: filter controls
│   └── Vertical · PANELS                     width remainder
│       ├── Horizontal · TOP                  height 470
│       │   ├── Vertical: Panel 1             width 58%
│       │   │   ├── Text: "1. Executive Overview"
│       │   │   ├── Horizontal: S1_KPI ×6     height 80
│       │   │   └── Horizontal: S2_Map | S3_Drivers
│       │   └── Vertical: Panel 2             width 42%
│       │       ├── Text: "2. Community Needs"
│       │       └── S4_CareScatter
│       └── Horizontal · BOTTOM               height 470
│           ├── Vertical: Panel 3             width 58%
│           │   ├── Horizontal: S5 | S6
│           │   └── S7_BarrierHeatmap
│           └── Vertical: Panel 4             width 42%
│               ├── Horizontal: S8 | S9
│               └── Text: "Why this community stands out"
└── Text · FOOTER                             height 40
```

**The nav rail is a Blank object** inside a vertical container with Layout →
Background set to `#FFFFFF`, plus Text objects for the page list. The five "pages" are
Text objects; if you want them clickable, add each as a separate dashboard and use
**Dashboard → Actions → Go to Sheet** on a navigation button.

---

## Step 5 · Formatting

### Palette

| Use | Hex |
|---|---|
| Header, footer, section titles | `#1F3A5F` |
| Dashboard background | `#EDF1F7` |
| Panel card | `#FFFFFF` |
| Card border | `#DCE5F0` |
| Need Index `> 80` | `#6B2B52` |
| `60 – 80` | `#A33E6B` |
| `40 – 60` | `#D9534F` |
| `20 – 40` | `#E88B5A` |
| `< 20` | `#FBEFC4` |
| Profile · Elder | `#3C7A9E` |
| Profile · Elder-leaning | `#6FA3BE` |
| Profile · Family | `#8E5FA8` |
| Profile · Family-leaning | `#B48FC9` |
| Profile · Mixed | `#B9BFC7` |
| Benchmark bars | `#C9D3E0` |
| Text secondary | `#6E7783` |

### Every panel container
Outer padding **8px**, inner padding **12px**, background white, border `#DCE5F0` 1px.
Tableau has no corner radius — accept square corners rather than faking them with
transparent PNGs, which breaks on every resize.

### Every worksheet
Format → Lines → turn off gridlines, zero lines, row and column dividers. Hide any
axis whose marks are already labelled. Rewrite every tooltip in plain English — no
`SUM([...])`, no `AVG`.

### Type
Set once in Format → Workbook. Panel titles 15px semibold `#1F3A5F`. Finding lines
13px semibold. Chart titles 11.5px. Axis and labels 9.5px `#6E7783`.

---

## Step 6 · Interactivity

1. **Parameter control** — right-click `p.Selected CSA` → Show Parameter. Place it in
   panel 1 next to "Selected CSA:" and again in panel 4. One parameter drives both.
2. **Map → parameter** — Dashboard → Actions → **Change Parameter**, source `S2_Map`,
   target parameter `p.Selected CSA`, field `CSA`, run on **Select**. Clicking a
   community now updates both the drivers panel and the profile.
3. **Highlight** — Dashboard → Actions → Highlight, source `S2_Map`, targets
   `S4_CareScatter` and `S5_PovertyOverdose`, run on **Hover**.
4. **Filters** — Year, Caregiving Profile, and Indicator (map) as dropdowns in the nav
   rail. Right-click each → Apply to Worksheets → **Selected Worksheets**, and be
   deliberate about which.

Stop there. Every extra filter is one more control that can be left in a wrong state
during a live presentation.

---

## Step 7 · Before you publish

- [ ] `2023_Analysis` shows **55 rows** — no city total leaked in
- [ ] Every benchmark is labelled **"City Median"**, not "Baltimore City"
- [ ] `S3_Drivers` axes are **synchronized**
- [ ] `S7_BarrierHeatmap` uses **separate legends** per column
- [ ] `S9_Trend` has an `Indicator` filter applied, and it is pinned
- [ ] No sheet is named "Sheet 1"
- [ ] Footer states the source, the 2023 vintage, that `Service Need Score` and
      `Barrier Burden Index` are constructed composites, and that relationships are
      ecological correlations across 55 areas
- [ ] Tested at 1366 × 768, the laptop a board member will actually open it on

---

## Appendix · Values to expect

Sanity checks. If your build disagrees with these, something is wrong.

| Check | Value |
|---|---|
| CSAs | 55 |
| City median, 65+ | 15.4% |
| City median, households with children | 23.1% |
| City median, no vehicle | 27.1% |
| City median, no internet | 10.4% |
| City median, no diploma | 11.6% |
| City median, under $25k | 25.4% |
| Correlation, 65+ vs children | 0.09 |
| Correlation, barriers vs 65+ | 0.08 |
| Correlation, barriers vs children | 0.34 |
| Highest 65+ | Howard Park/West Arlington, 22.9% |
| Highest households with children | Cherry Hill, 47.1% |
