# Baltimore Community Health & Access: Where Needs Overlap

An interactive Tableau dashboard that shows where Baltimore's community conditions and health risks overlap, across all 55 Community Statistical Areas (CSAs).

**[View the live dashboard on Tableau Public →](https://public.tableau.com/app/profile/mary.sweta.kerketta/viz/BaltimoreCommunityHealthAccess/D_Overview)**

[![Overview dashboard](https://public.tableau.com/static/images/Ba/BaltimoreCommunityHealthAccess/D_Overview/1.png)](https://public.tableau.com/app/profile/mary.sweta.kerketta/viz/BaltimoreCommunityHealthAccess/D_Overview)

---

## Why this exists

The Lily Foundation & Advocates wanted to know **where in Baltimore** to place patient navigation, advocacy, health literacy and caregiver support services.

The city publishes rich neighbourhood data through BNIA's *Vital Signs*: more than 200 indicators per neighbourhood, every year since 2010. In that form it's hard to act on. This project turns it into one view that answers three questions:

1. **Which communities carry the highest overall need?** (Overview)
2. **Are the people who need care the same people who face barriers to it?** (Community Needs)
3. **How do health risks relate to the conditions around them, such as poverty, housing and access?** (Determinants)

The dashboard sits on top of an analysis written up in [`Service_Need_Findings.md`](Service_Need_Findings.md) and [`Prenatal_Care_Findings.md`](Prenatal_Care_Findings.md).

---

## The dashboard

| Page | What it shows |
|---|---|
| **1. Overview** | KPI cards, a choropleth of the Potential Service Need Index by CSA, *What's Driving the Need?* (the selected CSA against Baltimore City), and a Key Indicators panel across demographics, economics, access, housing and health |
| **2. Community Needs** | Older adults (65+) plotted against a Barrier Burden Index, with bubble size showing households with children and colour showing a caregiving profile (Elder / Family / Mixed) |
| **3. Social Determinants & Health** | Poverty against overdose calls with a trend line, the top communities by eviction rate, and an access-barriers heatmap (education, internet, vehicle, unemployment) |

Every page has **CSA** and **Year** filters. The tabs across the top switch between pages.

### How the two indices are built

- **Service Need Score (0–100):** each CSA is ranked against the other 54 on measures in four areas, and the ranks are averaged:
  - health literacy
  - navigation barriers
  - caregiver load
  - complex-care need

  On the map the score is grouped into five bands: `< 20`, `20–40`, `40–60`, `60–80` and `> 80`.
- **Barrier Burden Index:** the average of four z-scores, calculated across all 55 CSAs:
  - adults without a high school diploma
  - households with no internet
  - households with no vehicle
  - households earning under $25K

  Zero is the city average.

---

## By the numbers

All figures come from the dashboard's data (2023, 55 CSAs) unless a different year is shown.

| | |
|---|---|
| Communities covered | **55** CSAs |
| Years of history | **14** (2010–2023) |
| Indicators in the trend table | **191** (95,437 CSA × indicator × year rows) |
| Households | **250,608** |
| City population (2020 Census) | **585,708**, down from 620,961 in 2010 |
| City median household income (2023) | **$59,869**, up from $38,346 in 2010 |

**Where need concentrates.** **9 CSAs** score above 80 on the Service Need Index and **21** score above 60. The five highest:

| Community | Need score |
|---|---|
| Upton/Druid Heights | 89.9 |
| Midway/Coldstream | 86.2 |
| Sandtown-Winchester/Harlem Park | 84.2 |
| Cherry Hill | 83.0 |
| Greenmount East | 82.8 |

**The gaps are wide.**

| Indicator | City median | Highest CSA |
|---|---|---|
| Households with no vehicle | 27.1% | **63.0%** (Oldtown/Middle East) |
| Households with no internet | 10.4% | **29.6%** (Dorchester/Ashburton) |
| Children below the poverty line | 24.2% | **74.3%** (Poppleton/The Terraces/Hollins Market) |
| Evictions per 1,000 residents | 10.4 | **37.3** (Dickeyville/Franklintown) |
| Overdose calls per 1,000 residents | 7.7 | **36.7** (Upton/Druid Heights) |

**Two relationships behind the dashboard:**

- **Poverty and overdose go together.** Family poverty and overdose calls correlate at **r = 0.58** across CSAs.
- **Who needs care isn't who faces barriers.** The share of residents aged 65+ barely relates to the Barrier Burden Index (**r = 0.08**). Communities with many older adults aren't necessarily the ones facing access barriers, so one combined index would send services to the wrong places.

**Income has risen, but unevenly.** Citywide median household income rose 56% between 2010 and 2023 ($38,346 to $59,869). The typical CSA in 2023 sat at $55,871.

> The findings reports use their own scoring and tiering, so their community rankings differ slightly from the dashboard's Service Need Score.

---

## Repository layout

```
.
├── dashboard/
│   ├── Baltimore_Community_Health_Access.twbx   ← the published workbook (data embedded as an extract)
│   └── BUILD_STEPS.md                           ← how the dashboard was built, calculated fields included
├── raw/
│   └── Vital Signs 2010-2023.xlsx               ← BNIA source workbook, one sheet per year
├── tableau/                                     ← cleaned, Tableau-ready data
│   ├── Baltimore_Vital_Signs_Tableau.xlsx       ← 2023/2018 snapshots, historical trends, correlations
│   ├── Community_Statistical_Areas_2020_*.geojson  ← CSA boundaries for the map
│   └── *.csv / *.json                           ← prenatal-care and helper tables
├── Vital_Signs_Cleaning_and_EDA.ipynb           ← cleaning, EDA, Service Need Score, Tableau export
├── Income_Prenatal_Care_Deep_Dive.ipynb         ← follow-up: income and the early-prenatal-care gap
├── prepare_dashboard_data.py                    ← builds the prenatal-care tables in tableau/
├── build_tableau_workbook.py                    ← experimental: generates a Tableau workbook from code
├── build_overview_dashboard.py                  ← experimental: same approach for an Overview workbook
├── Service_Need_Findings.md                     ← written findings and recommendations
├── Prenatal_Care_Findings.md                    ← companion findings on prenatal care
├── Baltimore_Vital_Signs_Dashboard_Plan.html    ← original dashboard design plan
└── Docs/                                        ← presentations and dashboard write-ups
```

Working drafts (`T_Work/`), earlier workbook versions and workbooks generated by the scripts are left out of git on purpose (see `.gitignore`).

---

## Using it

**Just view it:** open the [Tableau Public link](https://public.tableau.com/app/profile/mary.sweta.kerketta/viz/BaltimoreCommunityHealthAccess/D_Overview).

**Open the workbook:** open `dashboard/Baltimore_Community_Health_Access.twbx` in Tableau Public or Tableau Desktop **2026.2 or newer**. The data is packed inside the file as an extract, so it works without the source files.

**Re-run the analysis:**

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
jupyter notebook Vital_Signs_Cleaning_and_EDA.ipynb      # writes tableau/Baltimore_Vital_Signs_Tableau.xlsx
jupyter notebook Income_Prenatal_Care_Deep_Dive.ipynb
python prepare_dashboard_data.py
```

> **Publishing note:** Tableau Public only accepts workbooks whose data is an *extract*, and it can't create one itself. If you rebuild the workbook against the Excel file, open it in Tableau Desktop, go to **Data Source → Extract**, save, and then publish.

---

## Known issues

- **"What's Driving the Need?" chart:** the Evictions Rate and Overdose Calls bars show a "%" sign, but both measures are rates per 1,000 residents.
- **"What's Driving the Need?" chart:** the Evictions Rate and Overdose Calls bars show a "%" sign, but both measures are rates per 1,000 residents.
- **Year label:** the "Analysis Year" label assumes exactly 14 years of data (`= 14` in the `Year Display` calculation), so it will need updating when 2024 data is added.
- **Health outcomes:** life expectancy is last available for 2018 and prenatal care for 2019. Current conditions use 2023.

---

## Data source

**Baltimore Neighborhood Indicators Alliance (BNIA), Jacob France Institute, University of Baltimore:** *Vital Signs*, 2010–2023, by Community Statistical Area. CSA boundaries are the 2020 Community Statistical Areas published by BNIA and Baltimore City. Basemap © Mapbox © OpenStreetMap.

Prepared for **The Lily Foundation & Advocates** by Mary Sweta Kerketta.
