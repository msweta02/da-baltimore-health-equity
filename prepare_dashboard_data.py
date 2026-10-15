"""Prepare the data files behind the prenatal-care Tableau dashboard (Step 2).

Reads the outputs of Vital_Signs_Cleaning_and_EDA.ipynb and
Income_Prenatal_Care_Deep_Dive.ipynb; writes only into tableau/.
Run from the project root:  python prepare_dashboard_data.py
"""
import json
import os

import numpy as np
import pandas as pd

OUT      = "tableau"
TABLEAU  = f"{OUT}/Baltimore_Vital_Signs_Tableau.xlsx"
PRENATAL = f"{OUT}/Prenatal_Care_Deep_Dive.csv"
GEOJSON  = f"{OUT}/Community_Statistical_Areas_2020_8602756789985715764.geojson"

PN  = "Percent of Births Where the Mother Received Early Prenatal Care (First Trimester)"
INC = "Median Household Income"
BREAK_YEARS = [2013, 2014, 2015, 2016]      # detected in the deep-dive notebook, section 3
LATEST = 2019

hist    = pd.read_excel(TABLEAU, sheet_name="Historical_Trends")
profile = pd.read_excel(TABLEAU, sheet_name="Service_Need_Profile")
a23     = pd.read_excel(TABLEAU, sheet_name="2023_Analysis")
a18     = pd.read_excel(TABLEAU, sheet_name="2018_Analysis")
corrs   = pd.read_excel(TABLEAU, sheet_name="Correlations")
pn      = pd.read_csv(PRENATAL)

# --- a) expected prenatal care from income, and the shortfall ------------------------
# Linear, raw income: the same np.polyfit fit as the deep-dive notebook's section 7c,
# which is what produced the findings doc's figures.
pn = pn.drop(columns=["expected_prenatal_pct", "shortfall"], errors="ignore")
beta = np.polyfit(pn["income"], pn["prenatal_care_pct"], 1)
pn["expected_prenatal_pct"] = np.polyval(beta, pn["income"]).round(1)
pn["shortfall"] = (pn["prenatal_care_pct"] - np.polyval(beta, pn["income"])).round(1)
pn.to_csv(PRENATAL, index=False)

o = pn.set_index("CSA").loc["Orangeville/East Highlandtown"]
print(f"a) linear fit: prenatal = {beta[1]:.2f} + {beta[0]*1000:.4f} per $1k income")
print(f"   Orangeville/East Highlandtown: actual {o.prenatal_care_pct}, expected "
      f"{o.expected_prenatal_pct}, shortfall {o.shortfall}  (findings doc: 42.9 / 66.9 / -24.0)")
assert (o.expected_prenatal_pct, o.shortfall) == (66.9, -24.0)

# --- b) citywide trend ------------------------------------------------------------------
p_long = hist[hist.Indicator == PN].dropna(subset=["Value"])
city = (p_long.groupby("Year").Value
        .agg(city_median="median", pct_below_60=lambda v: 100 * (v < 60).mean())
        .round(1).reset_index())
city["is_break_year"] = city.Year.isin(BREAK_YEARS)
city.to_csv(f"{OUT}/prenatal_trend_citywide.csv", index=False)
print(f"\nb) prenatal_trend_citywide.csv: {len(city)} years; 2019 median "
      f"{city.set_index('Year').city_median[2019]} (findings doc: 62.5)")

# --- c) by-CSA trend --------------------------------------------------------------------
by_csa = p_long[["CSA", "Year", "Value"]].sort_values(["CSA", "Year"])
by_csa.to_csv(f"{OUT}/prenatal_trend_by_csa.csv", index=False)
print(f"c) prenatal_trend_by_csa.csv: {len(by_csa)} rows "
      f"({by_csa.CSA.nunique()} CSAs x {by_csa.Year.nunique()} years)")

# --- d) domain strength (static, from Service_Need_Findings.md) -------------------------
dom = pd.DataFrame({
    "Domain": ["Income", "Safety", "Housing", "Family structure", "Civic participation",
               "Transport access", "Race composition", "Employment", "Environment",
               "Local business"],
    "TypicalStrength": [0.71, 0.55, 0.48, 0.39, 0.34, 0.32, 0.25, 0.23, 0.23, 0.09]})
dom.to_csv(f"{OUT}/domain_strength.csv", index=False)
print(f"d) domain_strength.csv: {len(dom)} rows")

# --- e) barrier vs income adjustment ----------------------------------------------------
print(f"\ne) Correlations sheet outcomes: {sorted(corrs.Outcome.unique())}")
print("   -> no prenatal-care outcome, and 2023_Analysis holds raw indicators only, so the "
      "income-adjusted figures are not pre-computed anywhere. Building them.")
BARRIERS = {
    "No internet at home":      "Percent of Households with No Internet at Home",
    "No HS diploma":            "Percent Population (25 Years and over) With Less Than a High School Diploma or GED",
    "No vehicle":               "Percent of Households with No Vehicles Available",
    "Public transport to work": "Percent of Population that Uses Public Transportation to Get to Work",
}
w = (hist[hist.Year == LATEST]
     .pivot_table(index="CSA", columns="Indicator", values="Value", aggfunc="first"))

def resid(v, z):                       # same as Income_Prenatal_Care_Deep_Dive.ipynb, 7b
    X = np.column_stack([np.ones(len(z)), z])
    return v - X @ np.linalg.lstsq(X, v, rcond=None)[0]

rows = []
for lab, col in BARRIERS.items():
    d = w[[col, PN, INC]].dropna()
    x, y, z = d[col].values, d[PN].values, d[INC].values
    rows.append({"Barrier": lab,
                 "CorrelationAlone": round(float(np.corrcoef(x, y)[0, 1]), 2),
                 "CorrelationAdjusted": round(float(np.corrcoef(resid(x, z), resid(y, z))[0, 1]), 2)})
bar = pd.DataFrame(rows)
bar.to_csv(f"{OUT}/barrier_adjustment.csv", index=False)
print(bar.to_string(index=False))
expected = {"No internet at home": (-0.66, -0.16), "No HS diploma": (-0.70, -0.36),
            "No vehicle": (-0.48, 0.29), "Public transport to work": (-0.51, 0.14)}
for _, r in bar.iterrows():
    assert (r.CorrelationAlone, r.CorrelationAdjusted) == expected[r.Barrier], r.Barrier

# --- helper: population for the KPI cards -----------------------------------------------
# 2023_Analysis has no Total Population (collected in census years only). The findings
# doc's "129,237 residents (22%)" is the 2020 Census figure, so carry that.
pop = (hist[(hist.Indicator == "Total Population") & (hist.Year == 2020)]
       [["CSA", "Value"]].rename(columns={"Value": "Total Population 2020"}))
pop.to_csv(f"{OUT}/csa_population_2020.csv", index=False)
hi = pop.merge(profile[["CSA", "Tier"]], on="CSA")
n_hi = int((hi.Tier == "Highest").sum())
p_hi = int(hi.loc[hi.Tier == "Highest", "Total Population 2020"].sum())
share = p_hi / hi["Total Population 2020"].sum()
print(f"\nhelper csa_population_2020.csv: Highest tier = {n_hi} CSAs, {p_hi:,} residents "
      f"({share:.0%}) of {int(hi['Total Population 2020'].sum()):,}")
assert (n_hi, p_hi) == (13, 129237)

# --- f) CSA key check across everything that joins --------------------------------------
# The geojson uses 2020 CSA names. Seven differ from the CSA2010 key; the mapping is the
# CSA2020 column the cleaning notebook carried (section 4), identical to the CASE used by
# the relationship in the existing Tableau dashboard.
geo = {f["properties"]["CSA2020"] for f in json.load(open(GEOJSON))["features"]}
ren = (hist[hist.Year == 2023][["CSA", "CSA2020"]].drop_duplicates("CSA")
       .set_index("CSA").CSA2020)
RENAMES = {k: v for k, v in ren.items() if k != v}

keys = {"Service_Need_Profile": profile.CSA, "Historical_Trends": hist.CSA,
        "2023_Analysis": a23.CSA, "2018_Analysis": a18.CSA,
        "Prenatal_Care_Deep_Dive.csv": pn.CSA, "prenatal_trend_by_csa.csv": by_csa.CSA,
        "csa_population_2020.csv": pop.CSA}
base = set(profile.CSA)
print("\nf) CSA join key (exact string equality against Service_Need_Profile):")
for k, s in keys.items():
    s = set(s)
    print(f"   {k:30s} {len(s)} CSAs  {'exact match' if s == base else 'MISMATCH'}")
    assert s == base, k
mapped = {RENAMES.get(c, c) for c in base}
print(f"   geojson CSA2020               {len(geo)} features; {len(base & geo)} match CSA "
      f"directly, {len(RENAMES)} match after the 2020 renames:")
for a, b in RENAMES.items():
    print(f"      {a}  ->  {b}")
extra = sorted(geo - mapped)
print(f"   after renames: all 55 CSAs matched = {mapped <= geo}; unmatched polygons: {extra}")
assert mapped <= geo and len(mapped) == 55 and extra == ["Unassigned -- Jail"]
json.dump(RENAMES, open(f"{OUT}/csa2020_renames.json", "w"), indent=1)

print("\nWritten to tableau/:")
for f in ["Prenatal_Care_Deep_Dive.csv", "prenatal_trend_citywide.csv",
          "prenatal_trend_by_csa.csv", "domain_strength.csv", "barrier_adjustment.csv",
          "csa_population_2020.csv", "csa2020_renames.json"]:
    print(f"   {f:32s} {os.path.getsize(os.path.join(OUT, f)):>7,} bytes")
