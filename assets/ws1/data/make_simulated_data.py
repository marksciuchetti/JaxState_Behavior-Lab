#!/usr/bin/env python3
"""
make_simulated_data.py
======================
SIMULATED DATA GENERATOR for Workshop: Mobile Eye Tracking Shelf Study (analysis session)
Martin's Family Clothing Consumer Behavior Research Laboratory, Jacksonville State University

EVERYTHING THIS SCRIPT WRITES IS SIMULATED. No shopper was recorded. The numbers are
drawn from the probability model documented below so that students can practice the
analysis steps on an export that looks like an iMotions "Individual AOI metrics" export.

The worked example mirrors the iMotions Enablement Services "In-Store Shopper Report"
(March 2020): Sensodyne point of sale material (POSM) versus no POSM on an oral care
shelf, eye tracking glasses, N = 15 per shelf, ages 18 to 59. That report found that
POSM drew first attention to Sensodyne, that shoppers spent almost 25 percent less time
on competitor products with POSM present (94.3 s versus 127.9 s), and that 100 percent
of shoppers explored the Sensodyne POSM wing. The model below reproduces those patterns
in direction and rough size; it does not reproduce the report's data.

Design
------
Between subjects: each respondent shops ONE shelf (POSM or NoPOSM). 30 respondents,
15 per condition, IDs P01 to P30. Glasses recordings are gaze mapped onto one reference
photograph per shelf, and the AOIs are drawn on that reference image (the "parent" in
iMotions language). TTFF parent (ms) therefore counts from the moment the gaze map
interval starts, which the lab defines as the first frame in which the shopper faces
the shelf. Because each AOI is active for the whole gaze map interval, TTFF AOI (ms)
equals TTFF parent (ms) in this design.

AOIs
----
Sensodyne_Products, Sensodyne_POSM_Wing (POSM shelf only), Competitor_A, Competitor_B,
Price_Tags, Shelf_Other.

Metric names
------------
Column names after the identifier columns are the metric names in the iMotions
"AOI Metrics" help article (48 metrics; Information, Gaze based, Fixation based,
Saccade based, Mouse based). Where the same name exists in two categories (for example
Dwell time (ms) exists as a gaze based and a fixation based metric) the category is
prefixed so the two cannot be confused. Fixation based metrics carry no prefix because
they are the main metrics in this workshop. The manuals do not print the exact header
layout of the Individual AOI metrics export, so the identifier columns follow the
layout iMotions uses for other respondent level metric exports (Study Name,
Respondent Name, Gender, Age, Group, then the stimulus and the metrics). Verify the
real header layout on the lab laptop before reusing this script's column order on
real data.

Missing values
--------------
iMotions writes NA when a metric cannot be calculated, for example when the respondent
never fixated the AOI. This generator does the same: the row for that AOI is present
and the fixation based metrics are NA. The analysis treats TTFF as "not observed"
(it stays NA) and dwell time as 0 ms for those rows. TTFF max. (ms) follows the
iMotions definition and substitutes the parent duration when there was no fixation.

Model assumptions (all in seconds unless noted; converted to ms on output)
-----------------------------------------------------------------------
A1  Shelf exposure (Stimulus duration): Normal(mean 150, SD 30), clipped to 80 to 240.
    The 2020 report shows competitor time alone of 94 to 128 s, so a total shelf
    exposure of about two and a half minutes is plausible.
A2  Dwell time share by AOI (fraction of Stimulus duration), condition means:
                          NoPOSM   POSM
    Sensodyne_Products    0.12     0.20   (+8 percentage points with POSM)
    Sensodyne_POSM_Wing   n/a      0.06
    Competitor_A          0.18     0.135  (minus 25 percent with POSM)
    Competitor_B          0.15     0.1125 (minus 25 percent with POSM)
    Price_Tags            0.08     0.08
    Shelf_Other           0.17     0.145
    The means sum to about 0.70 because gaze also falls between products, off the
    shelf, on the shopper's own hands, or outside the reference image (unmapped).
    Each respondent's share = condition mean x respondent propensity x LogNormal(0, 0.28),
    where the propensity ~ LogNormal(0, 0.30) is one value per respondent that scales
    all of that shopper's shares (some shoppers study the shelf closely, others glance;
    this is the shopper level tendency that the random intercept in the mixed model
    captures). If a respondent's shares sum to more than 0.95 they are rescaled to 0.95.
A3  TTFF parent to Sensodyne_Products: NoPOSM mean 3.6 s (SD about 0.95);
    POSM mean 2.4 s (SD about 0.90). The POSM effect is therefore about minus 1.2 s.
    Competitor_A is noticed first on the NoPOSM shelf (mean 2.2 s) and later with POSM
    (mean 3.2 s). Other AOI means are listed in TTFF_MEAN below. TTFF values are
    drawn from a log-normal distribution so they are right skewed like real TTFF.
A4  Probability that a respondent never fixates an AOI (row present, metrics NA):
    Sensodyne_Products 0.08 NoPOSM and 0.03 POSM; Competitor_A 0.05; Competitor_B 0.10;
    Price_Tags 0.20; Shelf_Other 0.02; Sensodyne_POSM_Wing 0.00 (the report found that
    100 percent of shoppers explored the wing). The resulting fixation based
    Respondent ratio for the brand AOIs is near 85 to 95 percent.
A5  Fixation count = dwell time / average fixation duration, where average fixation
    duration ~ LogNormal(log 280 ms, 0.20), minimum one fixation. First fixation
    duration ~ LogNormal(log 220 ms, 0.30). Revisit count ~ Poisson(dwell_s / 6),
    Dwells with fixations = Revisit count + 1. Gaze based Dwell count adds
    Poisson(0.4) visits without a fixation. Gaze based Hit time parent (ms) is 80 to
    300 ms before TTFF parent (ms) because gaze enters the AOI before a fixation is
    classified; when there is no fixation a glance is still recorded with probability
    0.5, otherwise the gaze based metrics are NA too.
A6  Valid data (%) is one value per respondent (an Information metric, the percentage
    of collected and interpolated samples): Normal(88, 6) clipped to 72 to 97, except
    respondent P17 (POSM) at 46 percent, whose glasses slipped during the session.
    The workshop exclusion rule is Valid data below 70 percent, so P17 is excluded.
A7  Survey (one row per respondent, collected after the glasses were removed):
    Purchase intent (1 to 7) = 3.0 + 9 x Sensodyne share + 0.4 x POSM + Normal(0, 1.0),
    rounded and clipped to 1 to 7. Brand recall of Sensodyne: logistic with
    log odds = minus 1.6 + 14 x Sensodyne share + 0.6 x POSM. Product chosen: one draw
    from Sensodyne, Competitor_A, Competitor_B, Other with weights exp(10 x share) for
    each brand AOI and a constant for Other. Attention and self report are therefore
    correlated but far from perfectly, which is what real shelf studies show.
A8  Session log: five collection days in October 2026, two student operators
    (Operator_1, Operator_2), a gaze check note per respondent, and Data complete
    = No for P17. One other respondent (P09) needed a second gaze check but has
    complete data, so the log shows that a repeated check is not by itself a reason
    to exclude.

Run:  python3 make_simulated_data.py            (optional: python3 make_simulated_data.py <seed>)
Outputs (same folder): SIMULATED_shelf_study_AOI_metrics.csv,
                       SIMULATED_shelf_study_survey.csv,
                       SIMULATED_shelf_study_session_log.csv
"""

import csv
import math
import os
import random
from datetime import date, timedelta

# Fixed seed so the files are reproducible. The seed was chosen from a scan of 60
# candidate seeds as the one whose realized sample sits closest to the assumed
# means (TTFF difference, pooled SD, dwell share difference, competitor reduction,
# and the survey correlation). Any single sample of 30 drifts from its model, which
# is itself a lesson: compare the realized descriptives with the assumptions above.
import sys
SEED = int(sys.argv[1]) if len(sys.argv) > 1 else 20261037
random.seed(SEED)

HERE = os.path.dirname(os.path.abspath(__file__))
STUDY_NAME = "SIMULATED_ShelfStudy_POSM_vs_NoPOSM"

AOIS_NOPOSM = ["Sensodyne_Products", "Competitor_A", "Competitor_B", "Price_Tags", "Shelf_Other"]
AOIS_POSM = ["Sensodyne_Products", "Sensodyne_POSM_Wing", "Competitor_A", "Competitor_B",
             "Price_Tags", "Shelf_Other"]

# A2: mean dwell share of the shelf exposure, per condition and AOI
SHARE_MEAN = {
    "NoPOSM": {"Sensodyne_Products": 0.12, "Competitor_A": 0.18, "Competitor_B": 0.15,
               "Price_Tags": 0.08, "Shelf_Other": 0.17},
    "POSM":   {"Sensodyne_Products": 0.20, "Sensodyne_POSM_Wing": 0.06, "Competitor_A": 0.135,
               "Competitor_B": 0.1125, "Price_Tags": 0.08, "Shelf_Other": 0.145},
}
SHARE_LOGSD = 0.28       # AOI specific spread of the share (log scale)
PROPENSITY_LOGSD = 0.30  # respondent level looking propensity (log scale), shared by all AOIs
SHARE_CAP = 0.95         # shares cannot use the whole exposure

# A3: TTFF parent in seconds: (mean, SD) on the natural scale, drawn log-normally
TTFF_MEAN = {
    "NoPOSM": {"Sensodyne_Products": (3.6, 0.95), "Competitor_A": (2.2, 0.8),
               "Competitor_B": (4.5, 1.6), "Price_Tags": (8.0, 3.0), "Shelf_Other": (1.4, 0.6)},
    "POSM":   {"Sensodyne_Products": (2.4, 0.90), "Sensodyne_POSM_Wing": (2.0, 1.0),
               "Competitor_A": (3.2, 1.1), "Competitor_B": (5.0, 1.8), "Price_Tags": (8.5, 3.2),
               "Shelf_Other": (1.4, 0.6)},
}

# A4: probability of never fixating the AOI
P_MISS = {
    "NoPOSM": {"Sensodyne_Products": 0.08, "Competitor_A": 0.05, "Competitor_B": 0.10,
               "Price_Tags": 0.20, "Shelf_Other": 0.02},
    "POSM":   {"Sensodyne_Products": 0.03, "Sensodyne_POSM_Wing": 0.00, "Competitor_A": 0.05,
               "Competitor_B": 0.10, "Price_Tags": 0.20, "Shelf_Other": 0.02},
}

AGE_BANDS = ["18-24", "25-34", "35-44", "45-59"]
BAD_RESPONDENT = "P17"   # A6: poor data quality, to be excluded


def lognormal_from_mean_sd(mean, sd):
    """Draw one log-normal value whose natural scale mean and SD match the arguments."""
    sigma2 = math.log(1 + (sd / mean) ** 2)
    mu = math.log(mean) - sigma2 / 2
    return random.lognormvariate(mu, math.sqrt(sigma2))


def poisson(lam):
    """Knuth's algorithm; fine for the small rates used here."""
    L = math.exp(-lam)
    k, p = 0, 1.0
    while True:
        p *= random.random()
        if p <= L:
            return k
        k += 1


def clip(x, lo, hi):
    return max(lo, min(hi, x))


def fmt(x, nd=0):
    """Format a number for the CSV; NA stays NA."""
    if x is None:
        return "NA"
    return f"{x:.{nd}f}"


# ---------------------------------------------------------------------------
# 1. Respondents
# ---------------------------------------------------------------------------
respondents = []
# Alternate conditions so that the two groups are collected in parallel over the
# five days (a good practice: it avoids confounding condition with collection day).
conditions = ["POSM", "NoPOSM"] * 15
sexes_posm = ["M"] * 8 + ["F"] * 7
sexes_noposm = ["F"] * 8 + ["M"] * 7
random.shuffle(sexes_posm)
random.shuffle(sexes_noposm)
start_day = date(2026, 10, 12)          # Monday
collection_days = [start_day + timedelta(days=d) for d in range(5)]  # Mon to Fri

ip, inp = 0, 0
for i in range(30):
    rid = f"P{i + 1:02d}"
    cond = conditions[i]
    if cond == "POSM":
        sex = sexes_posm[ip]; ip += 1
    else:
        sex = sexes_noposm[inp]; inp += 1
    age = random.choices(AGE_BANDS, weights=[0.35, 0.30, 0.20, 0.15])[0]
    day = collection_days[i // 6]       # six respondents per day
    operator = "Operator_1" if (i % 2 == 0) else "Operator_2"
    # A1 shelf exposure
    dur_s = clip(random.gauss(150, 30), 80, 240)
    # A6 valid data
    if rid == BAD_RESPONDENT:
        valid = 46.0
    else:
        valid = clip(random.gauss(88, 6), 72, 97)
    respondents.append(dict(id=rid, cond=cond, sex=sex, age=age, date=day, operator=operator,
                            dur_s=dur_s, valid=valid))

# ---------------------------------------------------------------------------
# 2. AOI metrics, one row per respondent by AOI
# ---------------------------------------------------------------------------
aoi_rows = []
shares_by_resp = {}   # kept for the survey model (A7)

for r in respondents:
    cond = r["cond"]
    aois = AOIS_POSM if cond == "POSM" else AOIS_NOPOSM
    dur_ms = r["dur_s"] * 1000.0

    # A2: draw shares (respondent propensity x AOI specific noise), then cap the total
    propensity = random.lognormvariate(0, PROPENSITY_LOGSD)
    shares = {a: SHARE_MEAN[cond][a] * propensity * random.lognormvariate(0, SHARE_LOGSD) for a in aois}
    total = sum(shares.values())
    if total > SHARE_CAP:
        shares = {a: s * SHARE_CAP / total for a, s in shares.items()}

    # A4: which AOIs were never fixated
    fixated = {a: (random.random() >= P_MISS[cond][a]) for a in aois}
    # For the bad respondent, gaze mapping covered little of the recording: thin data
    if r["id"] == BAD_RESPONDENT:
        shares = {a: s * 0.35 for a, s in shares.items()}

    # the survey model uses the share of Sensodyne products actually recorded
    shares_by_resp[r["id"]] = {a: (shares[a] if fixated[a] else 0.0) for a in aois}

    for a in aois:
        row = {
            "Study Name": STUDY_NAME,
            "Respondent Name": r["id"],
            "Gender": r["sex"],
            "Age": r["age"],
            "Group": cond,                       # the condition variable
            "Session Date": r["date"].isoformat(),
            "Stimulus": f"Shelf_{cond}_GazeMap",  # the gaze map reference image (the parent)
            "AOI Name": a,
            "Stimulus duration": fmt(dur_ms),    # Information: average duration presented (ms)
            "AOI duration (ms)": fmt(dur_ms),    # Information: AOI active for the whole interval
            "AOI duration (%)": "100.0",
            "Valid data": fmt(r["valid"], 1),    # Information: percent collected and interpolated
        }
        if fixated[a]:
            dwell_ms = shares[a] * dur_ms
            mean_fix = random.lognormvariate(math.log(280), 0.20)          # A5
            fix_count = max(1, int(round(dwell_ms / mean_fix)))
            fix_dur = dwell_ms / fix_count
            first_fix = clip(random.lognormvariate(math.log(220), 0.30), 60, 1500)
            if fix_count == 1:
                first_fix = fix_dur
            ttff_mean, ttff_sd = TTFF_MEAN[cond][a]
            ttff_s = lognormal_from_mean_sd(ttff_mean, ttff_sd)
            ttff_ms = clip(ttff_s * 1000.0, 120, dur_ms - dwell_ms)
            revisits = poisson(dwell_ms / 1000.0 / 6.0)
            dwells_fix = revisits + 1
            gaze_dwells = dwells_fix + poisson(0.4)
            hit_ms = max(0.0, ttff_ms - random.uniform(80, 300))
            gaze_dwell_ms = dwell_ms * random.uniform(1.03, 1.15)   # gaze dwell includes saccade samples
            row.update({
                "Gaze Hit time parent (ms)": fmt(hit_ms),
                "Gaze Dwell count": str(gaze_dwells),
                "Gaze Dwell time (ms)": fmt(gaze_dwell_ms),
                "Dwells with fixations": str(dwells_fix),
                "Revisit count": str(revisits),
                "Fixation count": str(fix_count),
                "TTFF AOI (ms)": fmt(ttff_ms),
                "TTFF parent (ms)": fmt(ttff_ms),
                "TTFF max. (ms)": fmt(ttff_ms),
                "Dwell time (ms)": fmt(dwell_ms),
                "Dwell time (%)": fmt(100.0 * dwell_ms / dur_ms, 2),
                "Fixation duration (ms)": fmt(fix_dur),
                "First fixation duration (ms)": fmt(first_fix),
            })
        else:
            glance = random.random() < 0.5
            hit_ms = lognormal_from_mean_sd(*TTFF_MEAN[cond][a]) * 1000.0 if glance else None
            row.update({
                "Gaze Hit time parent (ms)": fmt(clip(hit_ms, 100, dur_ms - 500)) if glance else "NA",
                "Gaze Dwell count": "1" if glance else "NA",
                "Gaze Dwell time (ms)": fmt(random.uniform(100, 180)) if glance else "NA",
                "Dwells with fixations": "NA",
                "Revisit count": "NA",
                "Fixation count": "NA",
                "TTFF AOI (ms)": "NA",
                "TTFF parent (ms)": "NA",
                "TTFF max. (ms)": fmt(dur_ms),     # iMotions: parent duration when never fixated
                "Dwell time (ms)": "NA",
                "Dwell time (%)": "NA",
                "Fixation duration (ms)": "NA",
                "First fixation duration (ms)": "NA",
            })
        aoi_rows.append(row)

AOI_COLUMNS = list(aoi_rows[0].keys())

# ---------------------------------------------------------------------------
# 3. Survey, one row per respondent (A7)
# ---------------------------------------------------------------------------
survey_rows = []
for r in respondents:
    sh = shares_by_resp[r["id"]]
    s_sens = sh.get("Sensodyne_Products", 0.0)
    posm = 1 if r["cond"] == "POSM" else 0
    pi = 3.0 + 9.0 * s_sens + 0.4 * posm + random.gauss(0, 1.0)
    pi = int(clip(round(pi), 1, 7))
    logit = -1.6 + 14.0 * s_sens + 0.6 * posm
    recall = "Yes" if random.random() < 1 / (1 + math.exp(-logit)) else "No"
    weights = {
        "Sensodyne": math.exp(10.0 * s_sens),
        "Competitor_A": math.exp(10.0 * sh.get("Competitor_A", 0.0)),
        "Competitor_B": math.exp(10.0 * sh.get("Competitor_B", 0.0)),
        "Other": math.exp(1.0),
    }
    chosen = random.choices(list(weights), weights=list(weights.values()))[0]
    survey_rows.append({
        "Study Name": STUDY_NAME,
        "Respondent Name": r["id"],
        "Group": r["cond"],
        "Purchase_intent_Sensodyne_1to7": pi,
        "Brand_recall_Sensodyne": recall,
        "Product_chosen": chosen,
    })

# ---------------------------------------------------------------------------
# 4. Session log (A8)
# ---------------------------------------------------------------------------
GAZE_OK = ["Gaze check passed first attempt (center and four corners)",
           "Gaze check passed first attempt",
           "Gaze check passed; glasses adjusted on nose pad before check"]
log_rows = []
for r in respondents:
    if r["id"] == BAD_RESPONDENT:
        note = ("Gaze check passed, then glasses slipped at about 40 s; shopper pushed them up "
                "twice; gaze offset visible in replay; Valid data 46 percent")
        complete = "No"
        extra = "EXCLUDE: below the 70 percent Valid data rule set before collection"
    elif r["id"] == "P09":
        note = "First gaze check off by about 2 cm at the periphery; repositioned glasses; second check passed"
        complete = "Yes"
        extra = "Keep: second check passed and Valid data above 70 percent"
    else:
        note = random.choice(GAZE_OK)
        complete = "Yes"
        extra = ""
    log_rows.append({
        "Respondent Name": r["id"],
        "Session Date": r["date"].isoformat(),
        "Operator": r["operator"],
        "Group": r["cond"],
        "Shelf exposure (s)": fmt(r["dur_s"], 1),
        "Calibration / gaze check note": note,
        "Valid data (%)": fmt(r["valid"], 1),
        "Data complete": complete,
        "Decision note": extra,
    })

# ---------------------------------------------------------------------------
# 5. Write the three files. Line 1 of each file is a comment marking it SIMULATED;
#    the analysis scripts skip that line.
# ---------------------------------------------------------------------------
BANNER = (f"# SIMULATED DATA for teaching. Generated by make_simulated_data.py (seed {SEED}). "
          "No real shopper was recorded. Metric names follow the iMotions AOI Metrics article.")


def write_csv(name, columns, rows):
    path = os.path.join(HERE, name)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        fh.write(BANNER + "\n")
        w = csv.DictWriter(fh, fieldnames=columns, lineterminator="\n")
        w.writeheader()
        for row in rows:
            w.writerow(row)
    print(f"wrote {path} ({len(rows)} rows)")


write_csv("SIMULATED_shelf_study_AOI_metrics.csv", AOI_COLUMNS, aoi_rows)
write_csv("SIMULATED_shelf_study_survey.csv", list(survey_rows[0].keys()), survey_rows)
write_csv("SIMULATED_shelf_study_session_log.csv", list(log_rows[0].keys()), log_rows)

# Quick check of the headline numbers so the person running the generator sees them
import statistics as st
for cond in ("NoPOSM", "POSM"):
    vals = [float(r["TTFF parent (ms)"]) / 1000 for r in aoi_rows
            if r["Group"] == cond and r["AOI Name"] == "Sensodyne_Products"
            and r["TTFF parent (ms)"] != "NA" and r["Respondent Name"] != BAD_RESPONDENT]
    dw = [0.0 if r["Dwell time (%)"] == "NA" else float(r["Dwell time (%)"]) for r in aoi_rows
          if r["Group"] == cond and r["AOI Name"] == "Sensodyne_Products"
          and r["Respondent Name"] != BAD_RESPONDENT]
    print(f"{cond}: TTFF Sensodyne mean {st.mean(vals):.2f} s (SD {st.stdev(vals):.2f}, n={len(vals)}); "
          f"dwell share mean {st.mean(dw):.1f}% (SD {st.stdev(dw):.1f}, n={len(dw)})")
