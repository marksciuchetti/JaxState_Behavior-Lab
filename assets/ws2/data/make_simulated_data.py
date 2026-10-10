#!/usr/bin/env python3
"""
make_simulated_data.py
======================
SIMULATED DATA GENERATOR for Workshop 2: Screen Based Eye Tracking of Television Commercials
(analysis session)
Martin's Family Clothing Consumer Behavior Research Laboratory, Jacksonville State University

EVERYTHING THIS SCRIPT WRITES IS SIMULATED. No viewer was recorded. The numbers are drawn
from the probability model documented below so that students can practice the analysis
steps on an export that looks like an iMotions "Individual AOI metrics" export.

The design follows the structure of the iMotions demonstration study "DEMO - SonyBravia"
(one brand, two television commercials, an eye tracker on a monitor). No data from that
study, and no property of the real commercials other than their working titles "Balls" and
"Paint", were used. Each commercial is 60 s long in this simulation.

Design (within subjects, the contrast to Workshop 1's between subjects design)
------------------------------------------------------------------------------
Each respondent watches BOTH commercials on the lab monitor with the Smart Eye Aurora
(60 Hz): Commercial_1_Balls and Commercial_2_Paint. A Baseline slide comes before each
commercial and a rating survey slide after each. Order is counterbalanced: Group A sees
Balls first, Group B sees Paint first. 30 respondents, R01 to R30, 15 per order.
Ages 19 to 45, Gender F or M, two collection days (2026-10-19, 2026-10-20), two operators.
The Group column therefore holds the ORDER (A or B), not a treatment; the analysis derives
the position of each commercial (first or second) from Group and Stimulus.

Dynamic AOIs (an AOI is active only while its element is on screen)
-------------------------------------------------------------------
AOI duration (ms) and AOI duration (%) record the active time. Dwell time (%) is
fixation based and relative to the time the AOI was active (AOI Metrics article).

    AOI            Active time Balls   Active time Paint   First activation (s into the ad)
    Product_TV     14 s                20 s                Balls 12.0, Paint 9.0
    Brand_Logo      5 s                 5 s                55.0 (end card, one interval)
    Tagline_Text    4 s                 4 s                56.0 (end card, one interval)
    Hero_Element   40 s                32 s                Balls 2.0, Paint 3.0

Product_TV and Hero_Element are active in several intervals; the end card AOIs have one.
TTFF AOI (ms) counts from the AOI's first activation; TTFF parent (ms) counts from the
start of the commercial (first activation + TTFF AOI).

Metric names
------------
Column names after the identifier columns are the metric names in the iMotions "AOI
Metrics" help article. Where the same name exists in two categories (Dwell time (ms) is
both a gaze based and a fixation based metric) the gaze based column is prefixed "Gaze" and
the fixation based column carries no prefix, as in Workshop 1. The manuals do not print the
exact header layout of the Individual AOI metrics export, so the identifier columns follow
the layout iMotions uses for other respondent level exports (Study Name, Respondent Name,
Gender, Age, Group, then the stimulus and the metrics). Verify the real header layout on
the lab laptop before reusing this script's column order on real data.

Missing values
--------------
iMotions writes NA when a metric cannot be calculated, for example when the respondent
never fixated the AOI. This generator does the same: the row for that AOI is present and
the fixation based metrics are NA. The analysis treats TTFF as "not observed" (it stays NA)
and dwell time as 0 for those rows.

TTFF max. (ms): the AOI Metrics article defines it from the start of the AOI's PARENT (here
the commercial) and says that a respondent who never fixated receives the parent's
presentation duration (here 60000 ms). The generator follows that definition (observed
rows: equal to TTFF parent; never fixated rows: 60000). The workshop specification asked
for the AOI active duration (5000 ms for Brand_Logo) instead; that coding is used as a
SENSITIVITY analysis in analysis.R (H2b), where it belongs, not in the export column. To
write the AOI duration into the export column instead, set TTFF_MAX_MODE = "aoi" below.

Model assumptions (percent of the AOI active time unless noted)
---------------------------------------------------------------
A1  Valid data is one value per respondent (an Information metric, the percentage of
    collected and interpolated samples): Normal(91, 5) clipped to 60 to 99, except
    respondent R23, who is forced to 58 (glasses worn, calibration stayed Poor, recorded
    anyway). The workshop exclusion rule is Valid data below 70 percent, so R23 is the
    only respondent excluded. R23's dwell shares are also scaled by 0.6 and the chance of
    never fixating the end card AOIs is raised to 0.35 (the tracker lost the eyes often).
A2  Product_TV Dwell time (%) (the primary outcome). Marginal means across the two
    orders: Balls 22 (SD about 7), Paint 29 (SD about 8). Each respondent has one
    propensity z_i ~ N(0, 1) shared by both commercials, with loadings that make the
    within respondent correlation about 0.5:
        dwell_ij = mean_j + order_adjustment + SD_j x (sqrt(0.5) z_i + sqrt(0.5) e_ij)
    The expected paired effect is therefore about 7 percentage points, with an SD of the
    differences near 8, i.e. d_z of about 0.9 (SD of differences = sqrt(49 + 64 - 2 x 28)
    = 7.5 before the order effect adds a little).
A3  Order effect (carryover): whichever commercial is seen SECOND gets +3 percentage
    points of Product_TV dwell share relative to the same commercial seen first. Because
    the design is counterbalanced the effect averages out of the marginal means, so the
    means in A2 are marginal: the cell means are Balls first 20.5, Balls second 23.5, Paint
    first 27.5, Paint second 30.5 (order_adjustment = -1.5 when first, +1.5 when second).
A4  Brand_Logo TTFF AOI (ms) from end card onset: Balls mean 950 (SD 350), Paint mean 780
    (SD 320), log normal (right skewed like real TTFF), clipped to 150 to 4200 ms, with a
    respondent speed factor that makes the two commercials correlate about 0.3 on the log
    scale. About 10 percent of respondent by commercial rows never fixate the logo (row
    present, metrics NA). TTFF max. follows the manual (see above).
A5  Brand_Logo Dwell time (%): mean 40 (SD 14) among respondents who fixated the logo,
    with a respondent propensity (correlation about 0.4 between the two commercials).
    Dwell time (ms) cannot exceed 90 percent of the active time that remains after the
    first fixation.
A6  Hero_Element Dwell time (%): Balls mean 55, Paint mean 48 (SD 10). Always fixated.
    TTFF AOI: Balls mean 700 ms (SD 300), Paint 600 ms (SD 280).
A7  Tagline_Text Dwell time (%): about 30 in both commercials (mean 31.5 among the 95
    percent who fixate it, SD 12; about 5 percent never fixate it). TTFF AOI mean 1200 ms
    (SD 450).
A8  Product_TV TTFF AOI: Balls mean 1100 ms (SD 500), Paint 900 ms (SD 450). Always fixated.
A9  Fixation count and fixation duration. For each respondent by AOI row the average
    fixation duration is drawn from Normal(290, 60) clipped to 150 to 600 ms and the
    Fixation count is Dwell time (ms) divided by it, rounded, at least 1. Fixation
    duration (ms) is then Dwell time (ms) divided by Fixation count, so the two columns
    agree exactly. That gives about 3.4 to 3.5 fixations per second of dwell. First
    fixation duration ~ LogNormal(log 250 ms, 0.30), equal to Fixation duration when
    there is one fixation. Revisit count ~ Poisson(dwell s / 6); Dwells with fixations =
    Revisit count + 1. Gaze Dwell count adds Poisson(0.4) visits without a fixation; Gaze
    Dwell time is 3 to 15 percent longer than the fixation based dwell; Gaze Hit time AOI
    is 80 to 300 ms before TTFF AOI. When the AOI was never fixated, a glance is still
    recorded with probability 0.5, otherwise the gaze based metrics are NA too.
A10 Survey (one row per respondent). Liking of each commercial (1 to 7): Balls mean 5.6
    (SD 1.0), Paint mean 5.1 (SD 1.2), within respondent correlation about 0.4 (bivariate
    normal, rounded, clipped). Balls is liked more even though Paint draws more
    Product_TV attention: a teaching point about attention versus attitude.
    Brand_recall (Yes or No) is logistic in the respondent's TOTAL Brand_Logo dwell time
    over the two commercials (z score, slope 1.5, an expected point biserial r of about
    0.43); the intercept is solved so that the expected share of Yes is 0.80.
    Ad_preferred is the commercial with the higher liking; when the two ratings are equal
    (common on a 7 point scale) the respondent had to choose, modeled as a fair coin.
A11 Session log. Calibration result is Excellent or Good (Excellent more likely at higher
    Valid data); R09 and R17 were Poor and needed one recalibration (then Good); R23
    stayed Poor after two recalibrations and was recorded anyway. Distance to screen ~
    Normal(60, 3) cm; about 20 percent wear glasses (R23 does). Data complete is Yes for
    all (all recordings exist); the Decision note for R23 applies the Valid data rule.

Run:  python3 make_simulated_data.py                 (default seed below)
      python3 make_simulated_data.py <seed>          (any other seed)
      python3 make_simulated_data.py --scan [n]      (rank n candidate seeds; writes nothing)
Outputs (same folder): SIMULATED_ad_test_AOI_metrics.csv,
                       SIMULATED_ad_test_survey.csv,
                       SIMULATED_ad_test_session_log.csv
"""

import csv
import math
import os
import random
import statistics as st
import sys
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))

# Fixed seed so the files are reproducible. The seed was chosen from a scan of 8000 candidate
# seeds, 20261101 to 20269100 (python3 make_simulated_data.py --scan 8000), as the one whose
# realized sample of 29 kept respondents sits closest to the assumed means, SDs,
# correlations, and paired effect sizes in A2 to A10 (see TARGETS below). The scan ranks
# seeds by closeness to the MODEL, never by a p value. Any single sample of 30 drifts from
# its model, which is itself a lesson: compare the realized descriptives with the
# assumptions above.
SEED = 20268242
SCAN_BASE = 20261101

TTFF_MAX_MODE = "parent"   # "parent" follows the AOI Metrics article; "aoi" writes the AOI duration

STUDY_NAME = "SIMULATED_AdTest_Balls_vs_Paint"
STIM = {"Balls": "Commercial_1_Balls", "Paint": "Commercial_2_Paint"}
STIM_MS = 60000
AOIS = ["Product_TV", "Brand_Logo", "Tagline_Text", "Hero_Element"]
ACTIVE_MS = {"Balls": {"Product_TV": 14000, "Brand_Logo": 5000, "Tagline_Text": 4000, "Hero_Element": 40000},
             "Paint": {"Product_TV": 20000, "Brand_Logo": 5000, "Tagline_Text": 4000, "Hero_Element": 32000}}
START_MS = {"Balls": {"Product_TV": 12000, "Brand_Logo": 55000, "Tagline_Text": 56000, "Hero_Element": 2000},
            "Paint": {"Product_TV": 9000, "Brand_Logo": 55000, "Tagline_Text": 56000, "Hero_Element": 3000}}
SINGLE_INTERVAL = {"Brand_Logo", "Tagline_Text"}   # end card AOIs: one interval, TTFF limits the dwell

# A2 / A5 / A6 / A7: Dwell time (%) means (marginal for Product_TV) and SDs, and the share of
# the SD that is shared by the respondent (rho); a "fixator" mean for the end card AOIs.
DWELL = {
    ("Product_TV", "Balls"):   dict(mean=22.0, sd=7.0,  rho=0.5, lo=2, hi=80),
    ("Product_TV", "Paint"):   dict(mean=29.0, sd=8.0,  rho=0.5, lo=2, hi=85),
    ("Brand_Logo", "Balls"):   dict(mean=40.0, sd=14.0, rho=0.4, lo=3, hi=90),
    ("Brand_Logo", "Paint"):   dict(mean=40.0, sd=14.0, rho=0.4, lo=3, hi=90),
    ("Tagline_Text", "Balls"): dict(mean=31.5, sd=12.0, rho=0.4, lo=3, hi=85),
    ("Tagline_Text", "Paint"): dict(mean=31.5, sd=12.0, rho=0.4, lo=3, hi=85),
    ("Hero_Element", "Balls"): dict(mean=55.0, sd=10.0, rho=0.4, lo=10, hi=95),
    ("Hero_Element", "Paint"): dict(mean=48.0, sd=10.0, rho=0.4, lo=10, hi=95),
}
ORDER_BONUS_PP = 3.0     # A3: second viewing adds 3 percentage points of Product_TV dwell share

# A4 / A6 / A7 / A8: TTFF AOI (ms): (mean, SD, lower clip, upper clip)
TTFF = {
    ("Product_TV", "Balls"):   (1100, 500, 150, 8000), ("Product_TV", "Paint"):   (900, 450, 150, 8000),
    ("Brand_Logo", "Balls"):   (950, 350, 150, 4200),  ("Brand_Logo", "Paint"):   (780, 320, 150, 4200),
    ("Tagline_Text", "Balls"): (1200, 450, 200, 3300), ("Tagline_Text", "Paint"): (1200, 450, 200, 3300),
    ("Hero_Element", "Balls"): (700, 300, 100, 4000),  ("Hero_Element", "Paint"): (600, 280, 100, 4000),
}
TTFF_RHO = 0.3           # share of the log TTFF variance that is shared across a respondent's AOIs
P_MISS = {"Product_TV": 0.0, "Brand_Logo": 0.10, "Tagline_Text": 0.05, "Hero_Element": 0.0}
BAD_RESPONDENT = "R23"   # A1
RECAL_RESPONDENTS = ("R09", "R17")   # A11
LOGO_RECALL_SLOPE = 1.5  # A10, per SD of total Brand_Logo dwell time


# ---------------------------------------------------------------------------
# helpers (all take the random.Random instance so a seed scan never touches global state)
# ---------------------------------------------------------------------------
def lognormal_params(mean, sd):
    sigma2 = math.log(1 + (sd / mean) ** 2)
    return math.log(mean) - sigma2 / 2, math.sqrt(sigma2)


def poisson(rng, lam):
    L = math.exp(-lam)
    k, p = 0, 1.0
    while True:
        p *= rng.random()
        if p <= L:
            return k
        k += 1


def clip(x, lo, hi):
    return max(lo, min(hi, x))


def fmt(x, nd=0):
    return "NA" if x is None else f"{x:.{nd}f}"


def pearson(x, y):
    mx, my = st.mean(x), st.mean(y)
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    den = math.sqrt(sum((a - mx) ** 2 for a in x) * sum((b - my) ** 2 for b in y))
    return sxy / den if den > 0 else 0.0     # a constant column (for example, everyone recalls) has no correlation


def plogis(x):
    return 1 / (1 + math.exp(-x))


# ---------------------------------------------------------------------------
# the generator
# ---------------------------------------------------------------------------
def generate(seed):
    rng = random.Random(seed)

    # ---- respondents ----------------------------------------------------
    ids = [f"R{i + 1:02d}" for i in range(30)]
    order = ["A" if i % 2 == 0 else "B" for i in range(30)]
    gender = {}
    for g, n_f in (("A", 8), ("B", 7)):
        members = [ids[i] for i in range(30) if order[i] == g]
        sexes = ["F"] * n_f + ["M"] * (len(members) - n_f)
        rng.shuffle(sexes)
        gender.update(dict(zip(members, sexes)))
    ages = [int(round(clip(rng.gauss(27, 7), 19, 45))) for _ in range(30)]
    ages[ages.index(min(ages))] = 19      # the stated range 19 to 45 holds exactly
    ages[ages.index(max(ages))] = 45
    resp = []
    for i, rid in enumerate(ids):
        valid = 58.0 if rid == BAD_RESPONDENT else clip(rng.gauss(91, 5), 60, 99)
        resp.append(dict(
            id=rid, order=order[i], sex=gender[rid], age=ages[i],
            date=date(2026, 10, 19) if i < 15 else date(2026, 10, 20),
            operator="Operator_1" if (i // 2) % 2 == 0 else "Operator_2",
            valid=round(valid, 1),
            z_prod=rng.gauss(0, 1), z_ttff=rng.gauss(0, 1), z_logo=rng.gauss(0, 1),
            z_tag=rng.gauss(0, 1), z_hero=rng.gauss(0, 1)))

    # ---- AOI rows -------------------------------------------------------
    rows = []                    # CSV rows (dicts)
    tidy = []                    # numeric copies for the scan
    for r in resp:
        bad = r["id"] == BAD_RESPONDENT
        for cname in ("Balls", "Paint"):
            first = (r["order"] == "A") == (cname == "Balls")      # position of this commercial
            for aoi in AOIS:
                active = ACTIVE_MS[cname][aoi]
                start = START_MS[cname][aoi]
                key = (aoi, cname)
                zname = {"Product_TV": "z_prod", "Brand_Logo": "z_logo", "Tagline_Text": "z_tag",
                         "Hero_Element": "z_hero"}[aoi]
                d = DWELL[key]
                adj = 0.0
                if aoi == "Product_TV":
                    adj = -ORDER_BONUS_PP / 2 if first else ORDER_BONUS_PP / 2
                share = (d["mean"] + adj + d["sd"] * (math.sqrt(d["rho"]) * r[zname]
                                                      + math.sqrt(1 - d["rho"]) * rng.gauss(0, 1)))
                share = clip(share, d["lo"], d["hi"])
                if bad:
                    share *= 0.6
                p_miss = 0.35 if (bad and aoi in SINGLE_INTERVAL) else P_MISS[aoi]
                fixated = rng.random() >= p_miss
                mu, sg = lognormal_params(*TTFF[key][:2])
                ttff = math.exp(mu + sg * (math.sqrt(TTFF_RHO) * r["z_ttff"]
                                           + math.sqrt(1 - TTFF_RHO) * rng.gauss(0, 1)))
                ttff = clip(ttff, TTFF[key][2], TTFF[key][3])
                row = {
                    "Study Name": STUDY_NAME, "Respondent Name": r["id"], "Gender": r["sex"],
                    "Age": r["age"], "Group": r["order"], "Session Date": r["date"].isoformat(),
                    "Stimulus": STIM[cname], "AOI Name": aoi,
                    "Stimulus duration": fmt(STIM_MS),
                    "AOI duration (ms)": fmt(active),
                    "AOI duration (%)": fmt(100.0 * active / STIM_MS, 2),
                    "Valid data": fmt(r["valid"], 1),
                }
                if fixated:
                    dwell = share / 100.0 * active
                    if aoi in SINGLE_INTERVAL:
                        dwell = min(dwell, 0.9 * (active - ttff))
                    dwell = max(dwell, 120.0)
                    dwell_ms = round(dwell)
                    mean_fix = clip(rng.gauss(290, 60), 150, 600)
                    fix_count = max(1, int(round(dwell_ms / mean_fix)))
                    fix_dur = dwell_ms / fix_count
                    first_fix = clip(rng.lognormvariate(math.log(250), 0.30), 80, 1500)
                    if fix_count == 1:
                        first_fix = fix_dur
                    revisits = poisson(rng, dwell_ms / 1000.0 / 6.0)
                    dwells_fix = revisits + 1
                    gaze_dwells = dwells_fix + poisson(rng, 0.4)
                    ttff_ms = round(ttff)
                    hit_ms = max(0.0, ttff_ms - rng.uniform(80, 300))
                    gaze_dwell_ms = min(dwell_ms * rng.uniform(1.03, 1.15), float(active))
                    dwell_pct = round(100.0 * dwell_ms / active, 2)
                    row.update({
                        "Gaze Hit time AOI (ms)": fmt(hit_ms),
                        "Gaze Dwell count": str(gaze_dwells),
                        "Gaze Dwell time (ms)": fmt(gaze_dwell_ms),
                        "Dwells with fixations": str(dwells_fix),
                        "Revisit count": str(revisits),
                        "Fixation count": str(fix_count),
                        "TTFF AOI (ms)": fmt(ttff_ms),
                        "TTFF parent (ms)": fmt(start + ttff_ms),
                        "TTFF max. (ms)": fmt(start + ttff_ms),
                        "Dwell time (ms)": fmt(dwell_ms),
                        "Dwell time (%)": fmt(dwell_pct, 2),
                        "Fixation duration (ms)": fmt(fix_dur),
                        "First fixation duration (ms)": fmt(first_fix),
                    })
                    tidy.append(dict(id=r["id"], order=r["order"], c=cname, aoi=aoi, first=first,
                                     fixated=True, ttff=ttff_ms, dwell_ms=dwell_ms, dwell_pct=dwell_pct,
                                     fix_count=fix_count, fix_dur=fix_dur, valid=r["valid"]))
                else:
                    glance = rng.random() < 0.5
                    row.update({
                        "Gaze Hit time AOI (ms)": fmt(rng.uniform(300, active - 200)) if glance else "NA",
                        "Gaze Dwell count": "1" if glance else "NA",
                        "Gaze Dwell time (ms)": fmt(rng.uniform(100, 180)) if glance else "NA",
                        "Dwells with fixations": "NA", "Revisit count": "NA", "Fixation count": "NA",
                        "TTFF AOI (ms)": "NA", "TTFF parent (ms)": "NA",
                        "TTFF max. (ms)": fmt(STIM_MS if TTFF_MAX_MODE == "parent" else active),
                        "Dwell time (ms)": "NA", "Dwell time (%)": "NA",
                        "Fixation duration (ms)": "NA", "First fixation duration (ms)": "NA",
                    })
                    tidy.append(dict(id=r["id"], order=r["order"], c=cname, aoi=aoi, first=first,
                                     fixated=False, ttff=None, dwell_ms=0.0, dwell_pct=0.0,
                                     fix_count=0, fix_dur=None, valid=r["valid"]))
                rows.append(row)

    # ---- survey (A10) ---------------------------------------------------
    logo_total = {r["id"]: sum(t["dwell_ms"] for t in tidy if t["id"] == r["id"] and t["aoi"] == "Brand_Logo")
                  for r in resp}
    vals = list(logo_total.values())
    mu_l, sd_l = st.mean(vals), st.stdev(vals)
    zlogo = {k: (v - mu_l) / sd_l for k, v in logo_total.items()}
    lo_a, hi_a = -10.0, 10.0           # solve the intercept so the expected share of Yes is 0.80
    for _ in range(80):
        mid = (lo_a + hi_a) / 2
        if st.mean(plogis(mid + LOGO_RECALL_SLOPE * z) for z in zlogo.values()) < 0.80:
            lo_a = mid
        else:
            hi_a = mid
    intercept = (lo_a + hi_a) / 2
    survey = []
    for r in resp:
        zb = rng.gauss(0, 1)
        zp = 0.4 * zb + math.sqrt(1 - 0.4 ** 2) * rng.gauss(0, 1)
        lb = int(clip(round(5.6 + 1.0 * zb), 1, 7))
        lp = int(clip(round(5.1 + 1.2 * zp), 1, 7))
        recall = "Yes" if rng.random() < plogis(intercept + LOGO_RECALL_SLOPE * zlogo[r["id"]]) else "No"
        if lb > lp:
            pref = "Balls"
        elif lp > lb:
            pref = "Paint"
        else:
            pref = rng.choice(["Balls", "Paint"])
        survey.append({"Study Name": STUDY_NAME, "Respondent Name": r["id"], "Group": r["order"],
                       "Liking_Balls_1to7": lb, "Liking_Paint_1to7": lp,
                       "Brand_recall": recall, "Ad_preferred": pref})

    # ---- session log (A11) ----------------------------------------------
    log = []
    for r in resp:
        rid = r["id"]
        if rid == BAD_RESPONDENT:
            calib, recal, glasses = "Poor", 2, "Yes"
            complete = "Yes"
            note = ("EXCLUDE: Valid data 58 percent is below the 70 percent rule set before collection; "
                    "calibration stayed Poor after two recalibrations; recorded in full anyway because the "
                    "session was scheduled")
        elif rid in RECAL_RESPONDENTS:
            calib, recal = "Poor then Good after recalibration", 1
            glasses = "Yes" if rng.random() < 0.2 else "No"
            complete = "Yes"
            note = "Keep: recalibration reached Good and Valid data is above 70 percent"
        else:
            p_exc = 0.7 if r["valid"] >= 92 else 0.2
            calib = "Excellent" if rng.random() < p_exc else "Good"
            recal = 0
            glasses = "Yes" if rng.random() < 0.2 else "No"
            complete = "Yes"
            note = ""
        log.append({
            "Respondent Name": rid, "Session Date": r["date"].isoformat(), "Operator": r["operator"],
            "Group": r["order"], "Calibration result": calib, "Recalibrations": recal,
            "Distance to screen (cm)": fmt(clip(rng.gauss(60, 3), 52, 70), 1),
            "Glasses worn": glasses, "Valid data (%)": fmt(r["valid"], 1),
            "Data complete": complete, "Decision note": note})
    return dict(seed=seed, resp=resp, rows=rows, tidy=tidy, survey=survey, log=log)


# ---------------------------------------------------------------------------
# realized statistics for the kept respondents (used by the scan and the console summary)
# ---------------------------------------------------------------------------
def realized(data):
    kept = {r["id"] for r in data["resp"] if r["valid"] >= 70}
    t = [x for x in data["tidy"] if x["id"] in kept]
    out = {}

    def get(aoi, c, field):
        return {x["id"]: x[field] for x in t if x["aoi"] == aoi and x["c"] == c}

    b, p = get("Product_TV", "Balls", "dwell_pct"), get("Product_TV", "Paint", "dwell_pct")
    ids = sorted(kept)
    bb, pp = [b[i] for i in ids], [p[i] for i in ids]
    diff = [y - x for x, y in zip(bb, pp)]
    out["prod_mean_balls"], out["prod_mean_paint"] = st.mean(bb), st.mean(pp)
    out["prod_sd_balls"], out["prod_sd_paint"] = st.stdev(bb), st.stdev(pp)
    out["prod_r"] = pearson(bb, pp)
    out["prod_dz"] = st.mean(diff) / st.stdev(diff)
    ordr = {r["id"]: r["order"] for r in data["resp"]}
    dA = [d for i, d in zip(ids, diff) if ordr[i] == "A"]     # Paint second
    dB = [d for i, d in zip(ids, diff) if ordr[i] == "B"]     # Paint first
    out["order_effect"] = (st.mean(dA) - st.mean(dB)) / 2
    for c in ("Balls", "Paint"):
        v = [x["ttff"] for x in t if x["aoi"] == "Brand_Logo" and x["c"] == c and x["fixated"]]
        out[f"logo_ttff_{c}"] = st.mean(v)
    for c in ("Balls", "Paint"):
        lg = [x for x in t if x["aoi"] == "Brand_Logo" and x["c"] == c]
        out[f"logo_miss_{c}"] = sum(not x["fixated"] for x in lg) / len(lg)
    lb_t = {x["id"]: x["ttff"] for x in t if x["aoi"] == "Brand_Logo" and x["c"] == "Balls"}
    lp_t = {x["id"]: x["ttff"] for x in t if x["aoi"] == "Brand_Logo" and x["c"] == "Paint"}
    both = [i for i in ids if lb_t[i] is not None and lp_t[i] is not None]
    d2 = [lp_t[i] - lb_t[i] for i in both]
    out["h2_dz"] = st.mean(d2) / st.stdev(d2)
    # interaction of commercial and order position = difference in the two order groups' total dwell
    totA = [b[i] + p[i] for i in ids if ordr[i] == "A"]
    totB = [b[i] + p[i] for i in ids if ordr[i] == "B"]
    out["interaction"] = st.mean(totA) - st.mean(totB)
    for c, tgt in (("Balls", "hero_balls"), ("Paint", "hero_paint")):
        out[tgt] = st.mean(x["dwell_pct"] for x in t if x["aoi"] == "Hero_Element" and x["c"] == c)
    for c, tgt in (("Balls", "tag_balls"), ("Paint", "tag_paint")):
        out[tgt] = st.mean(x["dwell_pct"] for x in t if x["aoi"] == "Tagline_Text" and x["c"] == c)
    fx = [x for x in t if x["fixated"]]
    out["fixdur_mean"] = st.mean(x["fix_dur"] for x in fx)
    out["fixdur_sd"] = st.stdev(x["fix_dur"] for x in fx)
    out["fix_per_s"] = sum(x["fix_count"] for x in fx) / (sum(x["dwell_ms"] for x in fx) / 1000)
    sv = [s for s in data["survey"] if s["Respondent Name"] in kept]
    lb, lp = [s["Liking_Balls_1to7"] for s in sv], [s["Liking_Paint_1to7"] for s in sv]
    out["lik_mean_balls"], out["lik_mean_paint"] = st.mean(lb), st.mean(lp)
    out["lik_sd_balls"], out["lik_sd_paint"] = st.stdev(lb), st.stdev(lp)
    out["lik_r"] = pearson(lb, lp)
    dl = [y - x for x, y in zip(lb, lp)]
    out["lik_dz"] = st.mean(dl) / st.stdev(dl)
    out["recall_yes"] = sum(s["Brand_recall"] == "Yes" for s in sv) / len(sv)
    tot = {i: sum(x["dwell_ms"] for x in t if x["id"] == i and x["aoi"] == "Brand_Logo") for i in ids}
    rec = {s["Respondent Name"]: 1.0 if s["Brand_recall"] == "Yes" else 0.0 for s in sv}
    out["recall_r"] = pearson([tot[i] for i in ids], [rec[i] for i in ids])
    return out


# target value and tolerance. The tolerance is about one standard error of the statistic in a
# sample of 29; the statistics that carry the three hypotheses (Product_TV dwell share and its
# order effect, Brand_Logo TTFF, liking) get a tighter tolerance so the realized sample stays
# close to the planned effects, in particular the planned d_z of about 0.9. The other planned
# paired effect sizes are the values the model implies: d_z of about -0.43 for the Brand_Logo
# TTFF difference (Paint minus Balls, about -170 ms with an SD of differences near 400 ms) and
# about -0.41 for liking (-0.5 with an SD of differences near 1.2). The commercial by order
# position interaction is planned to be zero.
TARGETS = {
    "prod_mean_balls": (22.0, 0.8), "prod_mean_paint": (29.0, 0.9),
    "prod_sd_balls": (7.0, 1.0), "prod_sd_paint": (8.0, 1.1),
    "prod_r": (0.5, 0.07), "prod_dz": (0.9, 0.08), "order_effect": (3.0, 0.6),
    "logo_ttff_Balls": (950, 60), "logo_ttff_Paint": (780, 50), "logo_miss_Balls": (0.10, 0.05), "logo_miss_Paint": (0.10, 0.05),
    "h2_dz": (-0.43, 0.12), "lik_dz": (-0.41, 0.08), "interaction": (0.0, 3.5),
    "hero_balls": (55.0, 2.0), "hero_paint": (48.0, 2.0), "tag_balls": (30.0, 2.0), "tag_paint": (30.0, 2.0),
    "fixdur_mean": (290, 6), "fixdur_sd": (60, 8), "fix_per_s": (3.45, 0.1),
    "lik_mean_balls": (5.6, 0.15), "lik_mean_paint": (5.1, 0.18),
    "lik_sd_balls": (1.0, 0.13), "lik_sd_paint": (1.2, 0.16), "lik_r": (0.4, 0.15),
    "recall_yes": (0.80, 0.05), "recall_r": (0.43, 0.12),
}


def loss(stats):
    return sum(((stats[k] - tgt) / tol) ** 2 for k, (tgt, tol) in TARGETS.items())


# ---------------------------------------------------------------------------
# writing
# ---------------------------------------------------------------------------
def write_csv(path_name, columns, rows, banner):
    path = os.path.join(HERE, path_name)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        fh.write(banner + "\n")
        w = csv.DictWriter(fh, fieldnames=columns, lineterminator="\n")
        w.writeheader()
        for row in rows:
            w.writerow(row)
    print(f"wrote {path} ({len(rows)} rows)")


def main():
    args = sys.argv[1:]
    if args and args[0] == "--scan":
        n = int(args[1]) if len(args) > 1 else 120
        res = []
        for s in range(SCAN_BASE, SCAN_BASE + n):
            d = generate(s)
            res.append((loss(realized(d)), s))
        res.sort()
        print("rank  seed        loss (sum of squared standardized gaps to the model)")
        for k, (l, s) in enumerate(res[:10], start=1):
            print(f"{k:>4}  {s}  {l:7.2f}")
        print(f"median loss over {n} seeds: {st.median(l for l, _ in res):.2f}")
        return
    seed = int(args[0]) if args else SEED
    data = generate(seed)
    banner = (f"# SIMULATED DATA for teaching. Generated by make_simulated_data.py (seed {seed}). "
              "No real viewer was recorded. Metric names follow the iMotions AOI Metrics article.")
    write_csv("SIMULATED_ad_test_AOI_metrics.csv", list(data["rows"][0].keys()), data["rows"], banner)
    write_csv("SIMULATED_ad_test_survey.csv", list(data["survey"][0].keys()), data["survey"], banner)
    write_csv("SIMULATED_ad_test_session_log.csv", list(data["log"][0].keys()), data["log"], banner)
    stats = realized(data)
    print(f"\nRealized statistics for the 29 kept respondents (target, tolerance) and loss {loss(stats):.2f}:")
    for k, (tgt, tol) in TARGETS.items():
        print(f"  {k:<18} {stats[k]:9.3f}   target {tgt}")


if __name__ == "__main__":
    main()
