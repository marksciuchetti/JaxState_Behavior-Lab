#!/usr/bin/env python3
"""
make_simulated_data.py
======================
SIMULATED EXPORT BUNDLE GENERATOR for Workshop 7: exporting, opening, checking, cleaning, and
merging iMotions Lab 11.1 data into an analysis ready data set.
Martin's Family Clothing Consumer Behavior Research Laboratory, Jacksonville State University

EVERYTHING THIS SCRIPT WRITES IS SIMULATED. No person was recorded, no camera or sensor was
switched on, and no file from a lab study was used. The numbers come from the probability model
documented below. The files imitate what the lab would receive after exporting one screen based
study from iMotions Lab 11.1, problems included, so that students can practice the data
preparation steps in prepare_data.R (R) and SIMULATED_data_prep.xlsx (Excel).

Study
-----
"SIMULATED Multimodal Ad Study", a screen based lab study in the style of Workshops 2 to 4.
24 participants, P01 to P24, at the lab desktop (Key 0002) with the Smart Eye Aurora (60 Hz), the
respondent camera with Affectiva AFFDEX (about 30 frames per second), and a Shimmer3 GSR+ (GSR and
PPG, 128 Hz). Slideshow: Instructions (self paced), Baseline (60 s gray screen with a fixation
cross), Ad_A and Ad_B (two 30 s video ads for two fictional brands of reusable water bottles,
ALTAVO in the hiking ad Ad_A and BRENLOW in the cycling ad Ad_B), then Questionnaire (an iMotions
survey slide). Order of the ads counterbalanced: odd participants AB, even participants BA; the
operator types the order into the respondent field Group (FACTS Section 11: Group is copied into
every export). Each ad has three AOIs, active for the whole ad: Logo (top left, small), Product
(center), Headline (bottom). Each ad carries one stimulus annotation, "EndCard", from 25,000 ms to
30,000 ms after the ad starts (the price and call to action).

Survey (Question IDs on the slide "Questionnaire"; values are the iMotions LABEL VALUE)
    Liking_A, Liking_B   How much did you like the hiking ad / the cycling ad? 1 to 7
    Recall_A, Recall_B   Which brand was advertised in the hiking ad / the cycling ad?
                         Label1 ALTAVO, Label2 BRENLOW, Label3 CASKEL, Label4 DORVANE,
                         Label5 I do not remember. Scored in prepare_data.R: correct = 1 (yes),
                         anything else = 0 (no). Correct: Recall_A = 1, Recall_B = 2.
    ATT_1, ATT_2, ATT_3R attitude toward reusable water bottles, 1 = strongly disagree to
                         7 = strongly agree; ATT_3R ("... are a waste of money") is reverse worded
    PI                   How likely are you to buy a reusable water bottle in the next three
                         months? 1 to 7

Time bases (FACTS Sections 2, 4, 6)
    Sensor Data "Timestamp": ms since the RECORDING started (documented).
    R metrics "Start (ms)": "relative to the start of the study". MODELED here as ms since the
        StartSlide event of the first slide (Instructions). The recording starts 18 to 32 s
        earlier (operator start and calibration), so the two clocks differ by a constant per
        recording. Verify on the lab computer (FACTS Section 16, item 11): if the offset found by
        prepare_data.R is zero on real data, "start of the study" means the recording start.
    Annotation export "Start Time(ms)": ms since the STIMULUS onset (documented).

The export bundle (written to exports/; every file starts with one "# SIMULATED DATA" line,
which a real export does not have)
    01_Sensor_Data/NNN_<Respondent>.csv   Sensor Data, one file per recording for P01 to P06
        (7 files, because P04 was recorded twice). File name pattern "<three digit number>_<name>"
        is an Inference from screenshots (FACTS Section 2); the number is the respondent's
        position in the study, NOT the participant number. Layout: #INFO block (row names from
        the Sensor Data article's color coded illustration), #METADATA block (row names from the
        iMotions 8.2 file shown in Opening Data Export Files in Excel), "#DATA", the header row,
        then one row per device sample, devices interleaved by time and not resampled. Comma
        separated with a dot decimal (CHOICE: the newer File Info page offers "Separator: Comma ,"
        and "Tabulator"; the description panel says tab separated; verify). Only the stimuli
        Baseline, Ad_A, Ad_B and the sensors ET, AFFDEX, GSR were selected (Stimuli and Sensors
        pages of the editor). Options ticked: Study properties, Channel headers, Channel meta
        data, Include AOI markers, Include Annotation markers, Repeat stimulus name (as in the
        article's screenshot). AOI columns are written "Logo_active" and "Logo_gazed" (CHOICE:
        the 8.2 file screenshot; the article text says "<AOI> (Active)" and "(Gazed)"; verify).
        AOI booleans: 1 or 0 on eye tracker rows while the AOI is active, empty elsewhere.
        Annotation column "EndCard": 1 or empty (CHOICE: the illustration; the Using Annotations
        article says 1 or 0; verify). Channel names: "ET_GazeLeftx" style gaze names as in the
        editor preview (Tobii example; Aurora names not stated), AFFDEX names as in the Keyboard
        and Mouse article preview, "GSR Conductance CAL" (Epoching article); "Internal ADC A13
        PPG CAL" is MODELED. A lost gaze sample or a frame without a face is written as empty
        cells (MODELED).
    02_AOI_Metrics/IndividualAOIMetrics.csv and AggregatedAOIMetrics.csv   AOI Metrics Export
        (R PROCESSING in 11.1). Metric names from the AOI Metrics article. MODELED: the file names,
        the "#METADATA"/"#DATA" block, the identifier columns (the shared R metrics columns plus
        "AOI Name"), the "Gaze " prefix that separates gaze based metrics from fixation based
        metrics with the same name (Workshop 2 convention), and "Respondent ratio (%)" written as
        100 or 0 for one respondent. Fixation metrics are NA when the AOI was never fixated.
    03_FEA_Metrics/FEAMetrics.csv   FEA Metrics Export (expression metrics), long by expression,
        threshold 50. MODELED: file name, the metric columns "Expression", "Threshold", "Percent
        of time present", "Number of segments" (the manual names no columns), with the shared R
        metrics columns in front.
    04_GSR_Metrics/GSRPeakMetrics.csv   GSR Peak Metrics, columns exactly as the Export R
        Processed GSR Files article lists them (as in Workshop 4). MODELED: file name (from the
        attachment "Example_GSRPeakMetrics.csv"), METADATA lines.
    05_Survey/MERGED_SURVEY_RESPONSE_MATRIX.txt and SurveyStatistics.txt (documented names),
        tab separated (CHOICE: survey description panel). Full response header pattern
        LABELID_<slide>_<Question ID>,Question1 "<text>" (Inference from the screenshot; verify).
        SurveyStatistics_student_copy.csv: a copy that a student opened and saved again in Excel on
        a laptop with German regional settings (semicolon separator, comma decimal). FACTS
        Section 12 supports this: Excel uses the regional "Decimal symbol" and "List separator".
    06_Data_Quality/SIMULATED Multimodal Ad Study_DataQuality.txt   rows "#Study Name : ...",
        "#Study Date : ...", then Respondent_Name, Stim_Name, <device>_SamplingRate,
        <device>_Quality (documented). MODELED: tab separator, device names "Eye tracking",
        "Affectiva AFFDEX", "Shimmer GSR" (from the manual examples).
    07_Annotations/Individual Annotation Data.csv and Aggregate Annotation Result.csv (CHOICE: the
        ".csv" names of the Export: Annotations article, not the ".txt" Marker names; verify).
        Columns as the article's screenshots show them; first row "Segmentation: ..." (MODELED
        text).
    08_Session_Log/SIMULATED_AdStudy_SessionLog.csv   the lab session log (MODELED; one row per
        recording; the columns follow Project doc 06: ID, date, seat or station, calibration, notes).

Problems built in
    The bundle contains realistic problems of the kind real exports have. Finding them is
    Exercise 2 of Workshop 7; prepare_data.R, qa_report.txt, and cleaning_log.csv show how the
    pipeline catches each one.

Usage
    python3 make_simulated_data.py              write the bundle with the chosen seed
    python3 make_simulated_data.py --calibrate  population values of the statistics (the plan)
    python3 make_simulated_data.py --scan 300   rank seeds by closeness to the plan (not by p)
    python3 make_simulated_data.py 12345        write the bundle with another seed
"""

import csv
import math
import os
import shutil
import sys
from datetime import datetime, timedelta, timezone

import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "exports")
SEED = 20271231           # chosen with --scan 400 (rank 1 of 400 by closeness to the plan); see the README
SCAN_BASE = 20271000

STUDY = "SIMULATED Multimodal Ad Study"
IMOTIONS_VERSION = "11.1.47136.14"
IDS = [f"P{i:02d}" for i in range(1, 25)]
ADS = ["Ad_A", "Ad_B"]
STIMS = ["Baseline", "Ad_A", "Ad_B"]                 # the stimuli of the analysis
ALL_SLIDES = ["Instructions", "Baseline", "Ad_A", "Ad_B", "Questionnaire"]
AOIS = {"Logo": (80, 60, 380, 220), "Product": (700, 260, 1220, 760), "Headline": (460, 860, 1460, 1000)}
AOI_ORDER = ["Logo", "Product", "Headline"]
SCREEN = (1920, 1080)
ET_HZ, CAM_HZ, GSR_HZ = 60.0, 30.0, 128.0
AD_MS, BASE_MS = 30000, 60000
ENDCARD = (25000, 30000)
THRESH = 50
EXPRESSIONS = ["Joy", "Engagement", "Brow Furrow"]
TZ = timezone(timedelta(hours=-6))                 # Central Standard Time (Alabama, November)
COLLECTION_DAYS = [datetime(2026, 11, d, tzinfo=TZ) for d in (16, 17, 18, 19, 20)]
DQ_EXPORT = datetime(2026, 11, 20, 16, 40, 12, tzinfo=TZ)
SURVEY_EXPORT = datetime(2026, 11, 20, 16, 55, 31, tzinfo=TZ)
LATE_EXPORT = datetime(2026, 11, 23, 14, 12, 40, tzinfo=TZ)
STUDY_CREATED = datetime(2026, 11, 2, 14, 31, 7, tzinfo=TZ)

BRANDS = ["ALTAVO", "BRENLOW", "CASKEL", "DORVANE", "I do not remember"]
CORRECT = {"Ad_A": 1, "Ad_B": 2}
QUESTIONS = [  # (Question ID, question text, number of labels, label texts)
    ("Liking_A", "How much did you like the hiking ad? (1 = not at all, 7 = very much)", 7, None),
    ("Liking_B", "How much did you like the cycling ad? (1 = not at all, 7 = very much)", 7, None),
    ("Recall_A", "Which brand was advertised in the hiking ad?", 5, BRANDS),
    ("Recall_B", "Which brand was advertised in the cycling ad?", 5, BRANDS),
    ("ATT_1", "I like reusable water bottles. (1 = strongly disagree, 7 = strongly agree)", 7, None),
    ("ATT_2", "Buying a reusable water bottle is a smart choice. (1 = strongly disagree, 7 = strongly agree)", 7, None),
    ("ATT_3R", "Reusable water bottles are a waste of money. (1 = strongly disagree, 7 = strongly agree)", 7, None),
    ("PI", "How likely are you to buy a reusable water bottle in the next three months? (1 = very unlikely, 7 = very likely)", 7, None),
]

# Participants with designed problems (fixed by design, not drawn at random)
P_RERUN, P_DROPOUT, P_DEFAULTS, P_GAP = "P04", "P05", "P22", "P12"
P_LOW_ET, P_LOW_FACE = "P18", "P21"
SENSOR_DATA_IDS = ["P01", "P02", "P03", "P04", "P05", "P06"]

# ---------------------------------------------------------------------------------------------
# The probability model (the plan). All rates and effects are invented for teaching.
# ---------------------------------------------------------------------------------------------
LOGO_VISITS = {"Ad_A": 1.25, "Ad_B": 0.95}   # mean number of visits to the Logo AOI per 30 s ad
LOGO_SD = 0.85                               # SD of the respondent's log attention to logos
RECALL_B0, RECALL_HIT, RECALL_SD = -0.55, 1.55, 0.45   # logit P(correct brand) = b0 + b_hit * hit + u
JOY_RATE = {"Instructions": 0.25, "Baseline": 0.15, "Ad_A": 4.6, "Ad_B": 3.1, "Questionnaire": 0.35}
JOY_EXPR_SD = 0.75                           # SD of the log expressiveness of a respondent
LIKE_BASE = {"Ad_A": 4.45, "Ad_B": 4.10}
LIKE_JOY = 0.037                             # liking points per percentage point of time with Joy
LIKE_U_SD, LIKE_E_SD = 0.55, 1.05
GSR_BASE_RATE, GSR_BASE_SD = 2.3, 0.35       # peaks per minute at Baseline (median), log SD
GSR_MULT = {"Instructions": 1.25, "Baseline": 1.0, "Ad_A": 2.75, "Ad_B": 2.5, "Questionnaire": 1.45}
GSR_MULT_SD = 0.20
ATT_LOAD = (1.15, 1.05, -1.0)                # item loadings on the latent attitude (ATT_3R reversed)
ATT_MEAN = (4.7, 4.9, 3.3)
ATT_ERR = (0.75, 0.75, 0.85)


def clip(x, lo, hi):
    return max(lo, min(hi, x))


def invlogit(x):
    return 1.0 / (1.0 + math.exp(-x))


def rng_for(seed, comp, idx):
    """One independent random stream per component and recording, so that the light mode (used
    for the seed scan) draws exactly the same latent values as the full mode."""
    return np.random.default_rng([seed, comp, idx])


def to7(x):
    return int(clip(round(x), 1, 7))


# ---------------------------------------------------------------------------------------------
# Recordings: 24 participants, 25 recordings (P04 twice)
# ---------------------------------------------------------------------------------------------
def make_recordings(seed):
    r0 = rng_for(seed, 0, 0)
    recs = []
    genders = r0.permutation(["FEMALE"] * 13 + ["MALE"] * 11)
    k = 0
    for i, pid in enumerate(IDS):
        day = COLLECTION_DAYS[i // 5]
        slot = i % 5
        start = day.replace(hour=9) + timedelta(minutes=90 * slot + int(r0.integers(0, 12)), seconds=int(r0.integers(0, 60)))
        base = {"pid": pid, "num": i + 1, "order": "AB" if (i + 1) % 2 == 1 else "BA",
                "gender": str(genders[i]), "age": int(r0.choice([19, 20, 20, 21, 21, 21, 22, 22, 23, 24, 26])),
                "shimmer": "Shimmer3-8956" if (i % 4) in (0, 3) else "Shimmer3-B", "start": start,
                "operator": "Operator_1" if i < 12 else "Operator_2", "use": True, "aborted": False,
                "rec_no": 1, "name": pid}
        if pid == P_DROPOUT:
            base["shimmer"] = "Shimmer3-B"
        if pid == P_RERUN:
            first = dict(base, aborted=True, use=False, idx=k)
            k += 1
            recs.append(first)
            second = dict(base, rec_no=2, name="P04_rerun", idx=k, start=start + timedelta(minutes=13, seconds=41))
            k += 1
            recs.append(second)
            continue
        base["idx"] = k
        k += 1
        recs.append(base)
    for r in recs:
        r["group_field"] = r["order"]
        r["gender_field"] = r["gender"]
        r["age_field"] = r["age"]
        if r["pid"] == P_DEFAULTS:            # respondent details left at the iMotions defaults
            r["group_field"], r["gender_field"], r["age_field"] = "Default", "OTHER", 0
        r["name_dq"] = r["name"] if r["pid"] != "P07" else "P07 "
        r["name_survey"] = r["name"] if r["pid"] != "P07" else "p07"
        rg9 = rng_for(seed, 9, r["idx"])
        r["guid"] = "%08x-%04x-%04x-%04x-%012x" % tuple(int(rg9.integers(0, 2 ** b)) for b in (32, 16, 16, 16, 48))
    return recs


def timeline(seed, rec):
    """Slide events in ms since the recording started (the Sensor Data clock)."""
    rg = rng_for(seed, 2, rec["idx"])
    lead = float(rg.uniform(18000, 32000))
    ads = ["Ad_A", "Ad_B"] if rec["order"] == "AB" else ["Ad_B", "Ad_A"]
    slides = ["Instructions", "Baseline"] + ads + ["Questionnaire"]
    ev = {}
    t = lead
    for s in slides:
        ss = t
        sm = ss + float(rg.uniform(45, 110) if s in ADS else rg.uniform(15, 40))
        if s == "Instructions":
            dur = float(rg.integers(12000, 28000))
        elif s == "Questionnaire":
            dur = float(rg.integers(70000, 160000))
        else:
            dur = float(BASE_MS if s == "Baseline" else AD_MS)
        em = sm + dur
        es = em + float(rg.uniform(3, 12))
        ev[s] = {"StartSlide": ss, "StartMedia": sm, "EndMedia": em, "EndSlide": es, "planned": int(dur) if s in STIMS else -1}
        t = es + float(rg.uniform(1, 6))
        if rec["aborted"] and s == "Baseline":
            stop = sm + 23460.0 + float(rg.uniform(0, 400))
            ev[s]["EndMedia"] = None
            ev[s]["EndSlide"] = None
            ev[s]["stop"] = stop
            break
    rec_end = (ev["Baseline"]["stop"] if rec["aborted"] else ev["Questionnaire"]["EndSlide"] + float(rg.uniform(500, 2000)))
    return {"lead": lead, "slides": [s for s in slides if s in ev], "ev": ev, "end": rec_end, "ads": ads}


def interval(tl, s):
    """[start, end) of a stimulus for the metrics: StartMedia to EndMedia (or the recording stop)."""
    e = tl["ev"][s]
    return e["StartMedia"], (e["EndMedia"] if e["EndMedia"] is not None else e["stop"])


# ---------------------------------------------------------------------------------------------
# Latent traits
# ---------------------------------------------------------------------------------------------
def traits(seed, rec):
    rg = rng_for(seed, 1, rec["num"])           # by participant: the rerun keeps the person's traits
    t = {"logo_u": float(rg.normal(0, LOGO_SD)), "recall_u": float(rg.normal(0, RECALL_SD)),
         "expr": float(math.exp(rg.normal(0, JOY_EXPR_SD))), "like_u": float(rg.normal(0, LIKE_U_SD)),
         "gsr_rate": float(GSR_BASE_RATE * math.exp(rg.normal(0, GSR_BASE_SD))),
         "gsr_mult": {s: float(GSR_MULT[s] * math.exp(rg.normal(0, GSR_MULT_SD))) for s in ALL_SLIDES},
         "scl0": float(rg.uniform(2.0, 9.0)), "hr": float(rg.uniform(64, 84)), "att": float(rg.normal(0, 1)),
         "et_q": float(rg.uniform(93.5, 99.0)), "face_q": float(rg.uniform(90.0, 99.0)),
         "blink_rate": float(rg.uniform(9, 20)), "brow_rate": float(math.exp(rg.normal(math.log(0.6), 0.6))),
         "move_rate": float(math.exp(rg.normal(math.log(1.6), 0.5)))}
    return t


def quality_targets(rec, tr, s):
    """Target percent of valid gaze samples and of frames with a face, per stimulus."""
    et, face = tr["et_q"], tr["face_q"]
    if rec["aborted"]:
        et = 31.0
    if rec["pid"] == P_LOW_ET and s in ("Ad_A", "Ad_B", "Questionnaire"):
        et = {"Ad_A": 90.4, "Ad_B": 77.0, "Questionnaire": 80.0}[s]
    if rec["pid"] == P_LOW_FACE and s in ADS:
        face = {"Ad_A": 66.0, "Ad_B": 71.5}[s]
    return et, face


# ---------------------------------------------------------------------------------------------
# Eye tracking: fixation sequence (latent) and gaze samples (full mode)
# ---------------------------------------------------------------------------------------------
def in_rect(x, y, r):
    return (x >= r[0]) & (x <= r[2]) & (y >= r[1]) & (y <= r[3])


def point_in(rg, r, margin=0.12):
    w, h = r[2] - r[0], r[3] - r[1]
    return (float(rg.uniform(r[0] + margin * w, r[2] - margin * w)), float(rg.uniform(r[1] + margin * h, r[3] - margin * h)))


def point_other(rg):
    while True:
        x, y = float(rg.uniform(30, SCREEN[0] - 30)), float(rg.uniform(30, SCREEN[1] - 30))
        if not any(in_rect(x, y, (r[0] - 40, r[1] - 40, r[2] + 40, r[3] + 40)) for r in AOIS.values()):
            return x, y


def fixations_ad(seed, rec, tr, ad, t0, t1):
    """Visits (dwells) to AOIs, each a run of fixations; returns fixations with times in ms."""
    rg = rng_for(seed, 3, rec["idx"] * 10 + (1 if ad == "Ad_A" else 2))
    mu = LOGO_VISITS[ad] * math.exp(tr["logo_u"] - LOGO_SD ** 2 / 2)
    n_logo = int(rg.poisson(mu))
    logo_times = sorted(float(t) for t in rg.uniform(t0 + 400, t1 - 600, n_logo))
    fix = []
    t = t0 + float(rg.uniform(160, 260))
    probs = np.array([0.46, 0.20, 0.34])          # Product, Headline, Other
    while t < t1 - 80:
        if logo_times and t >= logo_times[0]:
            aoi = "Logo"
            logo_times.pop(0)
        else:
            aoi = ["Product", "Headline", None][int(rg.choice(3, p=probs))]
        nf = 1 + int(rg.geometric(0.5)) - 1 if aoi != "Logo" else 1 + int(rg.geometric(0.6)) - 1
        nf = max(1, nf)
        for _ in range(nf):
            dur = 60.0 + float(rg.gamma(4.0, 55.0))
            x, y = point_in(rg, AOIS[aoi]) if aoi else point_other(rg)
            if t + dur > t1:
                dur = t1 - t
            if dur < 50:
                break
            fix.append({"t": t, "d": dur, "x": x, "y": y, "aoi": aoi})
            t += dur + float(rg.uniform(18, 34))
            if t >= t1 - 80:
                break
        t += float(rg.uniform(30, 55))
    return fix


def fixations_baseline(seed, rec, t0, t1):
    rg = rng_for(seed, 3, rec["idx"] * 10)
    fix, t = [], t0 + float(rg.uniform(150, 260))
    while t < t1 - 80:
        dur = 120.0 + float(rg.gamma(3.0, 140.0))
        if rg.random() < 0.85:
            x, y = 960 + float(rg.normal(0, 45)), 540 + float(rg.normal(0, 40))
        else:
            x, y = float(rg.uniform(200, 1720)), float(rg.uniform(150, 930))
        dur = min(dur, t1 - t)
        fix.append({"t": t, "d": dur, "x": x, "y": y, "aoi": None})
        t += dur + float(rg.uniform(25, 50))
    return fix


def fixation_metrics(fix, t0, aoi):
    """Fixation based metrics of one AOI (an AOI active for the whole ad)."""
    runs, prev_in = [], False
    for f in fix:
        inside = f["aoi"] == aoi
        if inside and not prev_in:
            runs.append([f])
        elif inside:
            runs[-1].append(f)
        prev_in = inside
    if not runs:
        return None
    dwell = sum(r[-1]["t"] + r[-1]["d"] - r[0]["t"] for r in runs)
    return {"count": sum(len(r) for r in runs), "dwells": len(runs), "revisits": len(runs) - 1,
            "ttff": runs[0][0]["t"] - t0, "dwell": dwell, "first": runs[0][0]["d"]}


def gaze_samples(seed, rec, tr, s, tl, fix):
    """Gaze samples at 60 Hz between StartSlide and EndSlide (or the stop) of one stimulus."""
    rg = rng_for(seed, 4, rec["idx"] * 10 + STIMS.index(s))
    e = tl["ev"][s]
    a = e["StartSlide"]
    b = e["EndSlide"] if e["EndSlide"] is not None else e["stop"]
    phase = float(rng_for(seed, 4, rec["idx"] * 10 + 9).uniform(0, 1000 / ET_HZ))
    k0 = math.ceil((a - phase) * ET_HZ / 1000)
    k1 = math.floor((b - phase) * ET_HZ / 1000)
    ts = phase + np.arange(k0, k1 + 1) * 1000 / ET_HZ + rg.normal(0, 0.25, k1 - k0 + 1)
    n = len(ts)
    gx, gy = np.full(n, 960.0), np.full(n, 540.0)
    if fix:
        ft = np.array([f["t"] for f in fix])
        fe = np.array([f["t"] + f["d"] for f in fix])
        fx_ = np.array([f["x"] for f in fix])
        fy_ = np.array([f["y"] for f in fix])
        idx = np.searchsorted(ft, ts, side="right") - 1
        for j in range(n):
            i = idx[j]
            if i < 0:
                gx[j], gy[j] = fx_[0], fy_[0]
            elif ts[j] < fe[i] or i == len(fix) - 1:
                gx[j], gy[j] = fx_[i], fy_[i]
            else:                                   # saccade: straight line to the next fixation
                w = (ts[j] - fe[i]) / max(ft[i + 1] - fe[i], 1.0)
                gx[j] = fx_[i] + w * (fx_[i + 1] - fx_[i])
                gy[j] = fy_[i] + w * (fy_[i + 1] - fy_[i])
    off = rg.normal(0, 6, 2)
    lx = gx - 7 + off[0] + rg.normal(0, 9, n)
    ly = gy + off[1] + rg.normal(0, 9, n)
    rx = gx + 7 + off[0] + rg.normal(0, 9, n)
    ry = gy + off[1] + rg.normal(0, 9, n)
    et_q, _ = quality_targets(rec, tr, s)
    valid = np.ones(n, bool)
    dur_s = (b - a) / 1000
    for _ in range(int(rg.poisson(tr["blink_rate"] / 60 * dur_s))):     # blinks
        c = float(rg.uniform(a, b))
        valid[(ts >= c) & (ts < c + float(rg.uniform(90, 220)))] = False
    target_invalid = (100 - et_q) / 100
    guard = 0
    while (1 - valid.mean()) < target_invalid - 0.004 and guard < 4000:  # data loss segments
        c = float(rg.uniform(a, b))
        valid[(ts >= c) & (ts < c + float(rg.uniform(40, 700 if et_q > 85 else 1600)))] = False
        guard += 1
    return {"t": ts, "lx": lx, "ly": ly, "rx": rx, "ry": ry, "valid": valid}


def gaze_metrics(gs, t0, t1, aoi):
    """Gaze based metrics from samples: runs of samples inside the AOI that last at least 100 ms."""
    m = (gs["t"] >= t0) & (gs["t"] < t1)
    x = (gs["lx"][m] + gs["rx"][m]) / 2
    y = (gs["ly"][m] + gs["ry"][m]) / 2
    ins = gs["valid"][m] & in_rect(x, y, AOIS[aoi])
    tt = gs["t"][m]
    runs, start = [], None
    for j in range(len(ins)):
        if ins[j] and start is None:
            start = j
        if (not ins[j] or j == len(ins) - 1) and start is not None:
            end = j if ins[j] else j - 1
            runs.append((start, end))
            start = None
    dt = 1000 / ET_HZ
    runs = [r for r in runs if (r[1] - r[0] + 1) * dt >= 100]
    if not runs:
        return None
    dwell = sum((r[1] - r[0] + 1) * dt for r in runs)
    return {"hit": tt[runs[0][0]] - t0, "count": len(runs), "revisits": len(runs) - 1, "dwell": dwell,
            "dwell_pct": 100 * dwell / (t1 - t0)}


# ---------------------------------------------------------------------------------------------
# AFFDEX frames (about 30 per second over the whole recording)
# ---------------------------------------------------------------------------------------------
def episodes(rg, rate_per_min, a, b, mean_s, hi_share):
    out = []
    for _ in range(int(rg.poisson(rate_per_min * (b - a) / 60000))):
        c = float(rg.uniform(a, b))
        d = 300 + float(rg.gamma(2.0, mean_s * 500))
        peak = float(rg.uniform(55, 99)) if rg.random() < hi_share else float(rg.uniform(15, 48))
        out.append((c, d, peak))
    return out


def affdex_frames(seed, rec, tr, tl):
    rg = rng_for(seed, 5, rec["idx"])
    phase = float(rg.uniform(0, 33))
    n = int((tl["end"] - phase) * CAM_HZ / 1000)
    ts = phase + np.arange(n) * 1000 / CAM_HZ + rg.normal(0, 1.5, n)
    joy = np.abs(rg.normal(0, 0.6, n))
    brow = np.abs(rg.normal(0, 1.2, n))
    move = np.abs(rg.normal(0, 3.0, n))
    face = np.ones(n, bool)

    def add(arr, eps):
        for c, d, pk in eps:
            m = (ts >= c) & (ts < c + d)
            if m.any():
                u = (ts[m] - c) / d
                arr[m] = np.maximum(arr[m], pk * np.sin(np.pi * u) ** 0.6)

    for s in tl["slides"]:
        a, b = interval(tl, s)
        add(joy, episodes(rg, JOY_RATE[s] * tr["expr"], a, b, 2.4, 0.78))
        add(brow, episodes(rg, tr["brow_rate"] * (1.4 if s == "Questionnaire" else 0.7), a, b, 1.6, 0.6))
        add(move, episodes(rg, tr["move_rate"] * (2.5 if s in ("Questionnaire", "Instructions") else 1.0), a, b, 1.2, 0.55))
        _, fq = quality_targets(rec, tr, s)
        mm = (ts >= a) & (ts < b)
        while face[mm].mean() > fq / 100 + 0.004:
            c = float(rg.uniform(a, b))
            face[(ts >= c) & (ts < c + float(rg.uniform(100, 900 if fq > 85 else 2500)))] = False
    eng = np.clip(np.maximum.reduce([joy * 1.02, brow * 0.95, move]) + np.abs(rg.normal(0, 1.5, n)), 0, 100)
    val = np.clip(joy * 0.92 - brow * 0.75 + rg.normal(0, 0.8, n), -100, 100)
    return {"t": ts, "Joy": np.clip(joy, 0, 100), "Engagement": eng, "Valence": val, "Brow Furrow": np.clip(brow, 0, 100), "face": face}


def fea_metrics(fr, a, b, expr):
    m = (fr["t"] >= a) & (fr["t"] < b)
    v = fr[expr][m]
    present = fr["face"][m] & (v >= THRESH)
    seg = int(np.sum(present[1:] & ~present[:-1]) + (1 if len(present) and present[0] else 0))
    return 100 * present.mean() if len(present) else float("nan"), seg


# ---------------------------------------------------------------------------------------------
# GSR (Shimmer3 GSR+, 128 Hz): tonic level plus phasic peaks; PPG pulse wave
# ---------------------------------------------------------------------------------------------
def gsr_plan(seed, rec, tr, tl):
    rg = rng_for(seed, 6, rec["idx"])
    peaks = []
    for s in tl["slides"]:
        a, b = interval(tl, s)
        lam = tr["gsr_rate"] * tr["gsr_mult"][s]
        for _ in range(int(rg.poisson(lam * (b - a) / 60000))):
            apex = float(rg.uniform(a, b))
            peaks.append({"apex": apex, "amp": float(math.exp(rg.normal(math.log(0.13), 0.65))), "rise": float(rg.uniform(0.9, 2.2))})
    gaps = []
    if rec["pid"] == P_DROPOUT:
        a, _ = interval(tl, "Ad_B")
        gaps.append((a + 7812.0, tl["end"] + 1))
    if rec["pid"] == P_GAP:
        a, _ = interval(tl, "Baseline")
        gaps.append((a + 21950.0, a + 21950.0 + 5203.0))
    small = []                                     # short Bluetooth losses (< 4 s, not removed)
    for _ in range(int(rg.poisson(2))):
        c = float(rg.uniform(0, tl["end"]))
        small.append((c, c + float(rg.uniform(30, 400))))
    return {"peaks": [p for p in peaks if p["amp"] >= 0.02], "gaps": gaps, "small": small, "phase": float(rg.uniform(0, 7.8))}


def in_gaps(t, gaps):
    return any(g0 <= t < g1 for g0, g1 in gaps)


def gsr_metrics(gp, a, b):
    """GSR Peak Metrics for an interval [a, b): Signal Duration without gaps longer than 4 s."""
    gap_ms = sum(max(0.0, min(b, g1) - max(a, g0)) for g0, g1 in gp["gaps"])
    sig = (b - a) - gap_ms
    pk = [p for p in gp["peaks"] if a <= p["apex"] < b and not in_gaps(p["apex"], gp["gaps"])]
    n = len(pk)
    return {"sig": int(round(sig)), "count": n, "ppm": (n / (sig / 60000) if sig > 0 else None),
            "amp": (float(np.mean([p["amp"] for p in pk])) if n else None)}


def gsr_samples(seed, rec, tr, tl, gp, t_from, t_to):
    rg = rng_for(seed, 7, rec["idx"])
    k0 = math.ceil((t_from - gp["phase"]) * GSR_HZ / 1000)
    k1 = math.floor((t_to - gp["phase"]) * GSR_HZ / 1000)
    ts = gp["phase"] + np.arange(k0, k1 + 1) * 1000 / GSR_HZ
    keep = np.array([not in_gaps(t, gp["gaps"]) and not in_gaps(t, gp["small"]) for t in ts])
    ts = ts[keep]
    drift = np.cumsum(rg.normal(0, 0.0009, len(ts)))
    scl = tr["scl0"] + 0.25 * np.sin(ts / 90000) + drift - drift.mean() * 0.5
    phasic = np.zeros(len(ts))
    for p in gp["peaks"]:
        on = p["apex"] - p["rise"] * 1000
        m = (ts >= on) & (ts < on + 25000)
        if m.any():
            u = (ts[m] - on) / 1000
            shape = (1 - np.exp(-u / (p["rise"] / 3))) * np.exp(-np.maximum(u - p["rise"], 0) / 3.8)
            shape = shape / max(shape.max(), 1e-9)
            phasic[m] += p["amp"] * shape
    gsr = scl + phasic + rg.normal(0, 0.003, len(ts))
    ph = (ts / 1000) * tr["hr"] / 60
    frac = ph - np.floor(ph)
    pulse = np.where(frac < 0.18, np.sin(np.pi * frac / 0.36), np.exp(-(frac - 0.18) * 4.2) * (1 + 0.12 * np.sin(2 * np.pi * (frac - 0.18) / 0.3)))
    ppg = 1650 + 210 * pulse + rg.normal(0, 4, len(ts))
    return ts, gsr, ppg


# ---------------------------------------------------------------------------------------------
# One full simulation
# ---------------------------------------------------------------------------------------------
def generate(seed, full=True):
    recs = make_recordings(seed)
    D = {"seed": seed, "recs": recs, "tl": {}, "tr": {}, "fix": {}, "fixm": {}, "gaze": {}, "gazem": {},
         "fr": {}, "fea": {}, "gp": {}, "gsr": {}, "survey": {}, "dq": {}}
    for rec in recs:
        k = rec["idx"]
        tl = timeline(seed, rec)
        tr = traits(seed, rec)
        D["tl"][k], D["tr"][k] = tl, tr
        # eye tracking
        for s in tl["slides"]:
            if s not in STIMS:
                continue
            a, b = interval(tl, s)
            fx = fixations_ad(seed, rec, tr, s, a, b) if s in ADS else fixations_baseline(seed, rec, a, b)
            D["fix"][(k, s)] = fx
            if s in ADS:
                for aoi in AOI_ORDER:
                    D["fixm"][(k, s, aoi)] = fixation_metrics(fx, a, aoi)
            if full:
                gs = gaze_samples(seed, rec, tr, s, tl, fx)
                D["gaze"][(k, s)] = gs
                if s in ADS:
                    for aoi in AOI_ORDER:
                        D["gazem"][(k, s, aoi)] = gaze_metrics(gs, a, b, aoi)
        # facial expressions
        fr = affdex_frames(seed, rec, tr, tl)
        D["fr"][k] = fr
        for s in tl["slides"]:
            a, b = interval(tl, s)
            for x in EXPRESSIONS:
                D["fea"][(k, "Stimulus", s, x)] = fea_metrics(fr, a, b, x)
            if s in ADS:
                for x in EXPRESSIONS:
                    D["fea"][(k, "Annotation", s, x)] = fea_metrics(fr, a + ENDCARD[0], a + ENDCARD[1], x)
        # GSR
        gp = gsr_plan(seed, rec, tr, tl)
        D["gp"][k] = gp
        for s in tl["slides"]:
            a, b = interval(tl, s)
            D["gsr"][(k, "Stimulus", s)] = gsr_metrics(gp, a, b)
            if s in ADS:
                D["gsr"][(k, "Annotation", s)] = gsr_metrics(gp, a + ENDCARD[0], a + ENDCARD[1])
        # survey (only recordings that reached the Questionnaire)
        if not rec["aborted"]:
            rg = rng_for(seed, 8, rec["num"])
            ans = {}
            for ad, q in (("Ad_A", "Liking_A"), ("Ad_B", "Liking_B")):
                joy = D["fea"][(k, "Stimulus", ad, "Joy")][0]
                ans[q] = to7(LIKE_BASE[ad] + LIKE_JOY * (joy - 10) + tr["like_u"] + rg.normal(0, LIKE_E_SD))
            for ad, q in (("Ad_A", "Recall_A"), ("Ad_B", "Recall_B")):
                hit = D["fixm"][(k, ad, "Logo")] is not None
                p = invlogit(RECALL_B0 + RECALL_HIT * hit + tr["recall_u"])
                if rg.random() < p:
                    ans[q] = CORRECT[ad]
                else:
                    other = 3 - CORRECT[ad]
                    ans[q] = int(rg.choice([other, 3, 4, 5], p=[0.25, 0.1, 0.1, 0.55]))
            for j, q in enumerate(("ATT_1", "ATT_2", "ATT_3R")):
                ans[q] = to7(ATT_MEAN[j] + ATT_LOAD[j] * tr["att"] + rg.normal(0, ATT_ERR[j]))
            ans["PI"] = to7(3.9 + 0.95 * tr["att"] + rg.normal(0, 1.15))
            D["survey"][k] = ans
        # data quality (percent and effective sampling rate per stimulus)
        rgq = rng_for(seed, 10, rec["idx"])
        for s in tl["slides"]:
            a, b = interval(tl, s)
            et_q, face_q = quality_targets(rec, tr, s)
            if full and s in STIMS:
                gs = D["gaze"][(k, s)]
                m = (gs["t"] >= a) & (gs["t"] < b)
                etq = 100 * gs["valid"][m].mean()
            else:
                etq = clip(et_q + rgq.normal(0, 0.6), 0, 100)
            mm = (fr["t"] >= a) & (fr["t"] < b)
            fq = 100 * fr["face"][mm].mean()
            gm = D["gsr"][(k, "Stimulus", s)]
            small = sum(max(0.0, min(b, g1) - max(a, g0)) for g0, g1 in gp["small"])
            gq = max(0.0, 100 * (gm["sig"] - small) / (b - a))
            D["dq"][(k, s)] = {"et_rate": int(round(60 - rgq.uniform(0, 0.8))), "et_q": int(round(etq)),
                               "fea_rate": int(round(30 - rgq.uniform(0, 1.2))), "fea_q": int(round(fq)),
                               "gsr_rate": int(round(GSR_HZ * gq / 100 - rgq.uniform(0, 1.5))) if gq > 0 else 0,
                               "gsr_q": int(round(gq))}
    return D


# ---------------------------------------------------------------------------------------------
# Realized statistics (the logic mirrors prepare_data.R and analysis.R)
# ---------------------------------------------------------------------------------------------
def analysis_sample(D, population=False):
    keep = []
    for rec in D["recs"]:
        if not rec["use"]:
            continue
        if not population and rec["pid"] in (P_LOW_ET, P_LOW_FACE):
            continue
        keep.append(rec)
    return keep


def realized(D, population=False):
    keep = analysis_sample(D, population)
    o = {"n": len(keep)}
    # H1: logo fixated by brand recall, pooled over the two ads
    tab = np.zeros((2, 2))
    hits, recs_ = [], []
    for rec in keep:
        for ad, q in (("Ad_A", "Recall_A"), ("Ad_B", "Recall_B")):
            h = int(D["fixm"][(rec["idx"], ad, "Logo")] is not None)
            c = int(D["survey"][rec["idx"]][q] == CORRECT[ad])
            tab[h, c] += 1
            hits.append(h)
            recs_.append(c)
    a, b, c, d = tab[1, 1], tab[1, 0], tab[0, 1], tab[0, 0]
    o["h1_logor"] = math.log(((a + 0.5) * (d + 0.5)) / ((b + 0.5) * (c + 0.5)))
    n = tab.sum()
    o["h1_phi"] = (a * d - b * c) / math.sqrt(max((a + b) * (c + d) * (a + c) * (b + d), 1e-9))
    o["hit_rate"] = float(np.mean(hits))
    o["recall_rate"] = float(np.mean(recs_))
    for ad in ADS:
        o["hit_" + ad] = float(np.mean([D["fixm"][(r["idx"], ad, "Logo")] is not None for r in keep]))
    # H2: joy and liking
    J, L = [], []
    for ad, q in (("Ad_A", "Liking_A"), ("Ad_B", "Liking_B")):
        j = np.array([D["fea"][(r["idx"], "Stimulus", ad, "Joy")][0] for r in keep])
        l_ = np.array([D["survey"][r["idx"]][q] for r in keep], float)
        o["r_" + ad] = float(np.corrcoef(j, l_)[0, 1])
        o["joy_" + ad] = float(j.mean())
        o["like_" + ad] = float(l_.mean())
        J.append(j)
        L.append(l_)
    # within respondent (centered) correlation as the stand in for the mixed model slope
    jc = np.concatenate([J[0] - (J[0] + J[1]) / 2, J[1] - (J[0] + J[1]) / 2])
    lc = np.concatenate([L[0] - (L[0] + L[1]) / 2, L[1] - (L[0] + L[1]) / 2])
    o["r_within"] = float(np.corrcoef(jc, lc)[0, 1])
    o["r_pooled"] = float(np.corrcoef(np.concatenate(J), np.concatenate(L))[0, 1])
    # H3: GSR peaks per minute, each ad versus Baseline (GSR valid only)
    for ad in ADS:
        dd, bb, aa = [], [], []
        for r in keep:
            gb = D["gsr"][(r["idx"], "Stimulus", "Baseline")]
            ga = D["gsr"][(r["idx"], "Stimulus", ad)]
            dur_a = interval(D["tl"][r["idx"]], ad)
            if ga["sig"] < 0.8 * (dur_a[1] - dur_a[0]) or gb["sig"] < 0.8 * BASE_MS:
                continue
            dd.append(ga["ppm"] - gb["ppm"])
            bb.append(gb["ppm"])
            aa.append(ga["ppm"])
        dd = np.array(dd)
        o["dz_" + ad] = float(dd.mean() / dd.std(ddof=1))
        o["ppm_" + ad] = float(np.mean(aa))
        o["n_gsr_" + ad] = len(dd)
    o["ppm_base"] = float(np.mean([D["gsr"][(r["idx"], "Stimulus", "Baseline")]["ppm"] for r in keep]))
    # attitude scale: Cronbach's alpha after recoding ATT_3R
    X = np.array([[D["survey"][r["idx"]]["ATT_1"], D["survey"][r["idx"]]["ATT_2"], 8 - D["survey"][r["idx"]]["ATT_3R"]]
                  for r in D["recs"] if r["use"]], float)
    k = X.shape[1]
    o["alpha"] = float(k / (k - 1) * (1 - X.var(axis=0, ddof=1).sum() / X.sum(axis=1).var(ddof=1)))
    return o


# The plan: the expected value of each statistic over 200 simulated studies of this size
# (python3 make_simulated_data.py --calibrate 200) and its SD, used as the tolerance. The seed scan
# picks the data set whose statistics sit closest to these values; p values play no part.
TARGETS = {
    "h1_logor": (1.432, 0.575), "h1_phi": (0.343, 0.126), "hit_rate": (0.548, 0.079), "recall_rate": (0.562, 0.073),
    "hit_Ad_A": (0.583, 0.104), "hit_Ad_B": (0.513, 0.103),
    "r_Ad_A": (0.304, 0.209), "r_Ad_B": (0.252, 0.208), "r_within": (0.298, 0.205), "r_pooled": (0.303, 0.151),
    "joy_Ad_A": (11.373, 2.437), "joy_Ad_B": (8.005, 2.114), "like_Ad_A": (4.487, 0.278), "like_Ad_B": (4.027, 0.272),
    "dz_Ad_A": (1.006, 0.235), "dz_Ad_B": (0.890, 0.221), "ppm_base": (2.473, 0.442), "ppm_Ad_A": (6.879, 1.076),
    "ppm_Ad_B": (6.263, 0.958), "alpha": (0.808, 0.075),
}


def loss(o):
    return sum(((o[k] - t) / tol) ** 2 for k, (t, tol) in TARGETS.items())


# ---------------------------------------------------------------------------------------------
# Writing helpers
# ---------------------------------------------------------------------------------------------
def banner(seed, extra):
    return (f"# SIMULATED DATA for teaching. Generated by make_simulated_data.py (seed {seed}). "
            f"No real person was recorded. {extra}")


def f_(x, nd):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "NA"
    return f"{x:.{nd}f}"


def ensure(path):
    os.makedirs(path, exist_ok=True)
    return path


def write_rows(path, lines_before, header, rows, delim=",", quote_all=False):
    with open(path, "w", newline="", encoding="utf-8") as fh:
        for ln in lines_before:
            fh.write(ln + "\n")
        w = csv.writer(fh, delimiter=delim, lineterminator="\n", quoting=csv.QUOTE_MINIMAL)
        if header is not None:
            w.writerow(header)
        for r in rows:
            w.writerow(r)


SHARED = ["Study Name", "Respondent Name", "Gender", "Age", "Group", "Type", "Label", "Start (ms)", "Duration (ms)"]


def shared_cols(rec, tl, typ, label, ad=None):
    study0 = tl["lead"]                          # study start = StartSlide of the first slide
    a, b = interval(tl, ad or label)
    if typ == "Annotation":
        a, b = a + ENDCARD[0], a + ENDCARD[1]
    return [STUDY, rec["name"], rec["gender_field"], rec["age_field"], rec["group_field"], typ, label,
            int(round(a - study0)), int(round(b - a))]


def rows_for_metrics(D):
    """Yield (rec, tl, Type, Label, slide) in export order: respondents in study order, the slides in
    alphabetical order of Label within Stimulus rows, then annotation rows."""
    for rec in D["recs"]:
        tl = D["tl"][rec["idx"]]
        for s in sorted(tl["slides"]):
            yield rec, tl, "Stimulus", s, s
        for s in sorted(x for x in tl["slides"] if x in ADS):
            yield rec, tl, "Annotation", "EndCard", s


def write_gsr(D, folder):
    meta = ["#METADATA",
            "#Export,GSR Peak Metrics (R PROCESSING > GSR Peak Detection and Aggregation)",
            "#Devices,Shimmer3-8956 GSR;Shimmer3-B GSR (placeholder name)",
            "#Signal,GSR Conductance CAL (microSiemens)",
            "#Start (ms),relative to the start of the study",
            "#Signal Duration (ms),duration without gaps longer than 4 seconds",
            "#Peaks Per Minute,Peak Count divided by Signal Duration in minutes",
            "#Average Peak Amplitude (microSiemens),mean amplitude of the detected peaks",
            "#DATA"]
    head = SHARED + ["Signal Duration (ms)", "Parent Stimulus", "Has Peaks", "Peak Count", "Peaks Per Minute",
                     "Average Peak Amplitude (microSiemens)"]
    rows = []
    for rec, tl, typ, lab, s in rows_for_metrics(D):
        g = D["gsr"][(rec["idx"], typ, s)]
        rows.append(shared_cols(rec, tl, typ, lab, s) + [g["sig"], s, int(g["count"] > 0), g["count"], f_(g["ppm"], 3), f_(g["amp"], 4)])
    ban = banner(D["seed"], "GSR Peak Metrics; columns as in the Export R Processed GSR Files article. MODELED: file name, METADATA lines; "
                 "Start (ms) counted from the StartSlide event of the first slide (verify).")
    write_rows(os.path.join(folder, "GSRPeakMetrics.csv"), [ban] + meta, head, rows)
    return len(rows)


def write_fea(D, folder):
    meta = ["#METADATA",
            "#Export,FEA Metrics Export (R PROCESSING > FEA thresholding); Expression metrics",
            "#Device,Affectiva AFFDEX (respondent camera)",
            "#Threshold,50 (scores of 50 or more count as present)",
            "#Percent of time present,percent of the interval's frames with the expression at or above the threshold",
            "#Number of segments,number of separate runs at or above the threshold",
            "#DATA"]
    head = SHARED + ["Parent Stimulus", "Expression", "Threshold", "Percent of time present", "Number of segments"]
    rows = []
    for rec, tl, typ, lab, s in rows_for_metrics(D):
        for x in EXPRESSIONS:
            pct, seg = D["fea"][(rec["idx"], typ, s, x)]
            rows.append(shared_cols(rec, tl, typ, lab, s) + [s, x, THRESH, f_(pct, 2), seg])
    ban = banner(D["seed"], "FEA Metrics Export, threshold 50. MODELED: file name, METADATA lines, and the metric columns (the manual names none); "
                 "the identifier columns are the shared R metrics columns.")
    write_rows(os.path.join(folder, "FEAMetrics.csv"), [ban] + meta, head, rows)
    return len(rows)


AOI_IND = SHARED + ["Parent Stimulus", "AOI Name", "Stimulus duration", "AOI duration (ms)", "Valid data",
                    "Gaze Hit time AOI (ms)", "Gaze Dwell count", "Gaze Revisit count", "Gaze Dwell time (ms)",
                    "Gaze Dwell time (%)", "Gaze Respondent ratio (%)", "Dwells with fixations", "Revisit count",
                    "Fixation count", "TTFF AOI (ms)", "Dwell time (ms)", "Dwell time (%)", "First fixation duration (ms)",
                    "Respondent ratio (%)"]


def write_aoi(D, folder):
    meta = ["#METADATA",
            "#Export,AOI Metrics Export (R PROCESSING); Individual AOI metrics",
            "#Device,Eyetracker Smart Eye Aurora (60 Hz); fixations by the I-VT filter",
            "#Threshold for dwell classification,100 ms",
            "#Gaze metrics,columns that start with Gaze are gaze based; the others are fixation based",
            "#NA,the AOI was never gazed at (Gaze columns) or never fixated (fixation columns)",
            "#DATA"]
    rows, agg = [], {}
    for rec in D["recs"]:
        tl = D["tl"][rec["idx"]]
        for ad in sorted(x for x in tl["slides"] if x in ADS):
            a, b = interval(tl, ad)
            vd = D["dq"][(rec["idx"], ad)]["et_q"]
            for aoi in AOI_ORDER:
                g = D["gazem"][(rec["idx"], ad, aoi)]
                fm = D["fixm"][(rec["idx"], ad, aoi)]
                row = shared_cols(rec, tl, "Stimulus", ad) + [ad, aoi, int(round(b - a)), int(round(b - a)), f_(float(vd), 1)]
                row += ([int(round(g["hit"])), g["count"], g["revisits"], int(round(g["dwell"])), f_(g["dwell_pct"], 2), 100] if g
                        else ["NA", "NA", "NA", "NA", "NA", 0])
                row += ([fm["dwells"], fm["revisits"], fm["count"], int(round(fm["ttff"])), int(round(fm["dwell"])),
                         f_(100 * fm["dwell"] / (b - a), 2), int(round(fm["first"])), 100] if fm
                        else ["NA", "NA", "NA", "NA", "NA", "NA", "NA", 0])
                rows.append(row)
                A = agg.setdefault((ad, aoi), {"n": 0, "g": [], "f": [], "gd": [], "fd": [], "fc": [], "tt": [], "rv": []})
                A["n"] += 1
                A["g"].append(g is not None)
                A["f"].append(fm is not None)
                if g:
                    A["gd"].append(g["dwell"])
                if fm:
                    A["fd"].append(fm["dwell"])
                    A["fc"].append(fm["count"])
                    A["tt"].append(fm["ttff"])
                    A["rv"].append(fm["revisits"])
    ban = banner(D["seed"], "Individual AOI metrics. Metric names from the AOI Metrics article. MODELED: file name, METADATA lines, identifier columns, "
                 "the Gaze prefix, and Respondent ratio (%) as 100 or 0 per respondent.")
    write_rows(os.path.join(folder, "IndividualAOIMetrics.csv"), [ban] + meta, AOI_IND, rows)
    head = ["Study Name", "Label", "AOI Name", "Respondent base", "Gaze Respondent count", "Gaze Respondent ratio (%)",
            "Respondent count", "Respondent ratio (%)", "Gaze Dwell time (ms)", "Dwell time (ms)", "Fixation count", "TTFF AOI (ms)", "Revisit count"]
    arows = []
    for (ad, aoi), A in sorted(agg.items(), key=lambda kv: (kv[0][0], AOI_ORDER.index(kv[0][1]))):
        arows.append([STUDY, ad, aoi, A["n"], sum(A["g"]), f_(100 * np.mean(A["g"]), 1), sum(A["f"]), f_(100 * np.mean(A["f"]), 1),
                      f_(np.mean(A["gd"]) if A["gd"] else None, 1), f_(np.mean(A["fd"]) if A["fd"] else None, 1),
                      f_(np.mean(A["fc"]) if A["fc"] else None, 2), f_(np.mean(A["tt"]) if A["tt"] else None, 1),
                      f_(np.mean(A["rv"]) if A["rv"] else None, 2)])
    meta2 = ["#METADATA", "#Export,AOI Metrics Export (R PROCESSING); Aggregated AOI metrics (the table of the AOI window)",
             "#Averaging,over all respondents in the analysis segment; time metrics over respondents with a value", "#DATA"]
    ban2 = banner(D["seed"], "Aggregated AOI metrics, averaged over every respondent in the iMotions analysis (before any exclusion). "
                  "MODELED: file name and layout; Respondent base from the AOI window screenshot (the article text says Respondent data).")
    write_rows(os.path.join(folder, "AggregatedAOIMetrics.csv"), [ban2] + meta2, head, arows)
    return len(rows), len(arows)


def write_survey(D, folder):
    def hdr(prefix, q, text):
        return f'{prefix}_Questionnaire_{q},Question1 "{text}"'
    head = ["STUDY", "RESPONDENT", "GENDER", "AGE", "GROUP"]
    for q, text, _, _ in QUESTIONS:
        head += [hdr("LABELID", q, text), hdr("LABELVALUE", q, text)]
    rows = []
    for rec in D["recs"]:
        if rec["aborted"]:
            continue
        ans = D["survey"][rec["idx"]]
        r = [STUDY, rec["name_survey"], rec["gender_field"], rec["age_field"], rec["group_field"]]
        for q, _, _, _ in QUESTIONS:
            r += [f"Label{ans[q]}", ans[q]]
        rows.append(r)
    ban = banner(D["seed"], "Full Survey response (MERGED_SURVEY_RESPONSE_MATRIX.txt), tab separated. Header pattern from the Export: Survey data "
                 "screenshot (verify); Recall labels: Label1 ALTAVO, Label2 BRENLOW, Label3 CASKEL, Label4 DORVANE, Label5 I do not remember.")
    path = os.path.join(folder, "MERGED_SURVEY_RESPONSE_MATRIX.txt")
    with open(path, "w", newline="", encoding="utf-8") as fh:
        fh.write(ban + "\n")
        fh.write("\t".join(head) + "\n")
        for r in rows:
            fh.write("\t".join(str(v) for v in r) + "\n")
    # SurveyStatistics.txt: all respondents in the analysis, raw label values
    st_rows = []
    for q, text, _, _ in QUESTIONS:
        v = np.array([D["survey"][rec["idx"]][q] for rec in D["recs"] if not rec["aborted"]], float)
        st_rows.append(["Questionnaire", q, text, f"{v.mean():.2f}", f"{np.median(v):.1f}", f"{v.std(ddof=1):.2f}", len(v)])
    sh = ["Stimulus Name", "Question ID", "Question Text", "Average", "Median", "StdDev", "Count"]
    ban2 = banner(D["seed"], "Survey Statistics (SurveyStatistics.txt), tab separated; columns from the Export: Survey data screenshot. "
                  "Computed by iMotions over every respondent in the analysis, on the raw label values.")
    with open(os.path.join(folder, "SurveyStatistics.txt"), "w", newline="", encoding="utf-8") as fh:
        fh.write(ban2 + "\n")
        fh.write("\t".join(sh) + "\n")
        for r in st_rows:
            fh.write("\t".join(str(v) for v in r) + "\n")
    # the student's copy, saved again from Excel with German regional settings
    ban3 = ("# SIMULATED DATA for teaching. A student opened SurveyStatistics.txt in Excel on a laptop with German regional settings "
            "and saved it as CSV: semicolon separator and comma decimal. Not an iMotions export.")
    with open(os.path.join(folder, "SurveyStatistics_student_copy.csv"), "w", newline="", encoding="utf-8") as fh:
        fh.write(ban3 + "\n")
        w = csv.writer(fh, delimiter=";", lineterminator="\n")
        w.writerow(sh)
        for r in st_rows:
            w.writerow([r[0], r[1], r[2], r[3].replace(".", ","), r[4].replace(".", ","), r[5].replace(".", ","), r[6]])
    return len(rows), len(st_rows)


def write_dq(D, folder):
    head = ["Respondent_Name", "Stim_Name", "Eye tracking_SamplingRate", "Eye tracking_Quality", "Affectiva AFFDEX_SamplingRate",
            "Affectiva AFFDEX_Quality", "Shimmer GSR_SamplingRate", "Shimmer GSR_Quality"]
    rows = []
    for rec in D["recs"]:
        tl = D["tl"][rec["idx"]]
        for s in tl["slides"]:
            q = D["dq"][(rec["idx"], s)]
            rows.append([rec["name_dq"], s, q["et_rate"], q["et_q"], q["fea_rate"], q["fea_q"], q["gsr_rate"], q["gsr_q"]])
    ban = banner(D["seed"], "Data quality file from RECORDED DATA > Export, exported 2026-11-20 16:40. Layout from the Export: Data Quality Statistics "
                 "screenshot; MODELED: tab separator and the device names.")
    lines = [ban, f"#Study Name : {STUDY}", f"#Study Date : {STUDY_CREATED.strftime('%d-%m-%Y %H:%M:%S')}"]
    path = os.path.join(folder, f"{STUDY}_DataQuality.txt")
    with open(path, "w", newline="", encoding="utf-8") as fh:
        for ln in lines:
            fh.write(ln + "\n")
        fh.write("\t".join(head) + "\n")
        for r in rows:
            fh.write("\t".join(str(v) for v in r) + "\n")
    return len(rows)


def write_annotations(D, folder):
    head = ["Study Name", "Stimulus Name", "Respondent Name", "Gender", "Age", "Group", "Marker Name", "Start Time(ms)",
            "End Time(ms)", "Duration(ms)", "Annotation", "Color"]
    rows = []
    for rec in D["recs"]:
        tl = D["tl"][rec["idx"]]
        for ad in sorted(x for x in tl["slides"] if x in ADS):
            rows.append([STUDY, ad, rec["name"], rec["gender_field"], rec["age_field"], rec["group_field"], "EndCard",
                         f"{ENDCARD[0]:.2f}", f"{ENDCARD[1]:.2f}", f"{ENDCARD[1] - ENDCARD[0]:.2f}",
                         "Price and call to action on screen", "#FF9BBB59"])
    ban = banner(D["seed"], "Annotation export, Individual file. Columns from the Export: Annotations screenshot; times are relative to the "
                 "stimulus onset. MODELED: the Segmentation line text.")
    write_rows(os.path.join(folder, "Individual Annotation Data.csv"), [ban, "Segmentation: Default (all respondents)", ""], head, rows)
    ages = [r["age_field"] for r in D["recs"] if not r["aborted"]]
    ahead = ["Study Name", "Stimuli Name", "Gender", "From Age", "To Age", "Group Name", "Marker Name", "Base", "Avg Duration (ms)",
             "Standard Deviation (ms)"]
    arows = [[STUDY, ad, "BOTH", f">={min(ages)}", f"<={max(ages)}", "All Groups", "EndCard",
              sum(1 for r in D["recs"] if ad in D["tl"][r["idx"]]["slides"]), "5000.00", "0.00"] for ad in ADS]
    ban2 = banner(D["seed"], "Annotation export, Aggregate file. Columns from the Export: Annotations screenshot. MODELED: the Segmentation line text.")
    write_rows(os.path.join(folder, "Aggregate Annotation Result.csv"), [ban2, "Segmentation: Default (all respondents)", ""], ahead, arows)
    return len(rows), len(arows)


def write_session_log(D, folder):
    head = ["Participant ID", "iMotions respondent name", "Recording", "Session date", "Start time", "Seat", "Computer", "Operator",
            "Order (typed in Group)", "Shimmer unit", "Aurora calibration", "Use recording", "Notes"]
    notes = {
        ("P04", 1): ("Failed", "No", "Aurora calibration failed (gaze not tracked in the live view). Recording stopped 23 s into Baseline. "
                                    "Recalibrated and recorded again as P04_rerun. Do not use this recording."),
        ("P04", 2): ("Good after recalibration", "Yes", "Rerun of P04 after recalibration (Baseline had not finished, so no ad was seen twice). "
                                                         "Use this recording as P04."),
        ("P05", 1): ("Good", "Yes", "Shimmer3-B low battery warning during Ad_B; GSR signal lost about 8 s into Ad_B and did not return. "
                                    "Keep; GSR from Ad_B on is missing."),
        ("P12", 1): ("Good", "Yes", "Short Bluetooth dropout of the Shimmer during Baseline (a few seconds); signal returned."),
        ("P18", 1): ("Fair (2 attempts)", "Yes", "Wears glasses; reflections on the left lens; gaze lost often in the second ad."),
        ("P21", 1): ("Good", "Yes", "Rested chin on hand during both ads; reminded once to sit back."),
        ("P22", 1): ("Good", "Yes", "Started in a hurry (next participant waiting)."),
        ("P09", 1): ("Good", "Yes", "Arrived 10 minutes late."),
        ("P14", 1): ("Good", "Yes", "Left handed; Shimmer on the right hand."),
    }
    rows = []
    for rec in D["recs"]:
        cal, use, note = notes.get((rec["pid"], rec["rec_no"]), ("Good", "Yes", ""))
        rows.append([rec["pid"], rec["name"], rec["rec_no"], rec["start"].strftime("%Y-%m-%d"), rec["start"].strftime("%H:%M"),
                     1, "Lab desktop (Key 0002)", rec["operator"], rec["order"], rec["shimmer"], cal, use, note])
    ban = banner(D["seed"], "Lab session log, one row per recording (MODELED layout after Project doc 06: ID, date, seat or station, calibration, notes). "
                 "Written by the operators during collection.")
    write_rows(os.path.join(folder, "SIMULATED_AdStudy_SessionLog.csv"), [ban], head, rows)
    return len(rows)


# ---------------------------------------------------------------------------------------------
# Sensor Data files
# ---------------------------------------------------------------------------------------------
SD_COLS = ["Row", "Timestamp", "EventSource", "SlideEvent", "StimType", "Duration", "CollectionPhase", "SourceStimuliName",
           "EventSource", "ET_GazeLeftx", "ET_GazeLefty", "ET_GazeRightx", "ET_GazeRighty",
           "EventSource", "SampleNumber", "Joy", "Engagement", "Valence", "Brow Furrow",
           "EventSource", "GSR Conductance CAL", "Internal ADC A13 PPG CAL",
           "Logo_active", "Logo_gazed", "EndCard"]


def sd_metadata(shimmer):
    et, fe, sh = "Eyetracker Smart Eye Aurora", "Affectiva AFFDEX", f"{shimmer} GSR"
    dev = ["", "", "SlideEvents", "SlideEvents", "SlideEvents", "SlideEvents", "SlideEvents", "SlideEvents",
           et, et, et, et, et, fe, fe, fe, fe, fe, fe, sh, sh, sh, "AOI", "AOI", "Annotation"]
    cat = ["", "Timeline", "Event source", "Slideshow", "Slideshow", "Slideshow", "Slideshow", "Slideshow",
           "Event source", "Gaze 2D", "Gaze 2D", "Gaze 2D", "Gaze 2D", "Event source", "Other", "Emotions", "Emotions",
           "Emotions", "Action Units", "Event source", "GSR", "PPG", "AOI", "AOI", "Annotation"]
    desc = ["", "iMotions synchronization timestamp. Relative to data recording start",
            "Boolean which marks slide events (1 = slide event)", "StartSlide, StartMedia, EndMedia, EndSlide", "Stimulus type",
            "Planned duration of the slide (ms)", "Collection phase", "User-defined name of the slide",
            "Boolean which marks data samples (1 = data sample)", "Gaze X left eye", "Gaze Y left eye", "Gaze X right eye",
            "Gaze Y right eye", "Boolean which marks data samples (1 = data sample)", "Frame number",
            "Probability of the expression", "Probability of the expression", "Valence (-100 to 100)", "Probability of the action unit",
            "Boolean which marks data samples (1 = data sample)", "Skin conductance", "Photoplethysmogram",
            "Gaze sample while the AOI is active", "Gaze sample inside the AOI", "Data sample inside the annotation"]
    unit = ["", "Milliseconds", "", "", "", "Milliseconds", "", "", "", "Pixels", "Pixels", "Pixels", "Pixels", "", "",
            "Probability (0-100)", "Probability (0-100)", "Score", "Probability (0-100)", "", "MicroSiemens", "MilliVolts", "", "", ""]
    short = ["", "ms", "", "", "", "ms", "", "", "", "px", "px", "px", "px", "", "", "", "", "", "", "", "uS", "mV", "", "", ""]
    group = ["", "Timeline", "Slideshow", "Slideshow", "Slideshow", "Slideshow", "Slideshow", "Slideshow",
             "Eye tracking", "Eye tracking", "Eye tracking", "Eye tracking", "Eye tracking", "Facial expressions",
             "Facial expressions", "Facial expressions", "Facial expressions", "Facial expressions", "Facial expressions",
             "GSR", "GSR", "GSR", "AOIs", "AOIs", "Annotations"]
    disp = [c if c != "EventSource" else f"EventSource ({d})" for c, d in zip(SD_COLS, dev)]
    prov = ["", "iMotions"] + ["iMotions"] * 6 + ["Smart Eye"] * 5 + ["Affectiva"] * 6 + ["Shimmer"] * 3 + ["iMotions"] * 3
    chan = ["", "iMotions_Synchronization_Timestamp", "Eventsource_Slideevents", "SE_SlideEvent", "SE_StimulusType", "SE_Duration",
            "SE_CollectionPhase", "SE_SourceStimuliName", "Eventsource_ET", "ET_GazeLeftx", "ET_GazeLefty", "ET_GazeRightx",
            "ET_GazeRighty", "Eventsource_AFFDEX", "AFFDEX_SampleNumber", "AFFDEX_Joy", "AFFDEX_Engagement", "AFFDEX_Valence",
            "AFFDEX_BrowFurrow", "Eventsource_GSR", "GSR_Conductance_CAL", "PPG_A13_CAL", "AOI_Logo_active", "AOI_Logo_gazed",
            "Annotation_EndCard"]
    rows = [("#Device", dev), ("#Category", cat), ("#Description", desc), ("#Unit", unit), ("#Group", group),
            ("#Display name", disp), ("#Provided By", prov), ("#Short unit", short), ("#Channel identifier", chan)]
    return [[lab] + vals[1:] for lab, vals in rows]


def unix_ms(dt):
    return dt.timestamp()


def write_sensor_data(D, folder):
    sizes = {}
    rec_pos = {r["idx"]: r["idx"] + 1 for r in D["recs"]}         # position in the study = file number
    for rec in D["recs"]:
        if rec["pid"] not in SENSOR_DATA_IDS:
            continue
        k = rec["idx"]
        tl, tr, fr, gp = D["tl"][k], D["tr"][k], D["fr"][k], D["gp"][k]
        width = len(SD_COLS)
        out = []                                  # (timestamp, order, row list)
        stims = [s for s in tl["slides"] if s in STIMS]
        for s in stims:
            e = tl["ev"][s]
            win0 = e["StartSlide"]
            win1 = e["EndSlide"] if e["EndSlide"] is not None else e["stop"]
            ad_on = e["StartMedia"]
            stimtype = "TestVideo" if s in ADS else "TestImage"
            for name in ("StartSlide", "StartMedia", "EndMedia", "EndSlide"):
                if e.get(name) is None:
                    continue
                row = [""] * width
                row[1], row[2], row[3], row[4], row[5], row[6], row[7] = e[name], "1", name, stimtype, str(e["planned"]), "StimuliDisplay", s
                out.append((e[name], 0, row))
            # eye tracker
            gs = D["gaze"][(k, s)]
            m = (gs["t"] >= win0) & (gs["t"] < win1)
            active = s in ADS
            for t, lx, ly, rx, ry, v in zip(gs["t"][m], gs["lx"][m], gs["ly"][m], gs["rx"][m], gs["ry"][m], gs["valid"][m]):
                row = [""] * width
                row[1], row[7], row[8] = t, s, "1"
                if v:
                    row[9], row[10], row[11], row[12] = f"{lx:.1f}", f"{ly:.1f}", f"{rx:.1f}", f"{ry:.1f}"
                if active:
                    row[22] = "1"
                    inside = v and in_rect((lx + rx) / 2, (ly + ry) / 2, AOIS["Logo"])
                    row[23] = "1" if inside else "0"
                    if ad_on + ENDCARD[0] <= t < ad_on + ENDCARD[1]:
                        row[24] = "1"
                out.append((t, 1, row))
            # AFFDEX
            mm = (fr["t"] >= win0) & (fr["t"] < win1)
            idxs = np.nonzero(mm)[0]
            for j in idxs:
                t = fr["t"][j]
                row = [""] * width
                row[1], row[7], row[13], row[14] = t, s, "1", str(int(j))
                if fr["face"][j]:
                    row[15], row[16], row[17], row[18] = (f"{fr['Joy'][j]:.2f}", f"{fr['Engagement'][j]:.2f}", f"{fr['Valence'][j]:.2f}",
                                                          f"{fr['Brow Furrow'][j]:.2f}")
                if s in ADS and ad_on + ENDCARD[0] <= t < ad_on + ENDCARD[1]:
                    row[24] = "1"
                out.append((t, 2, row))
            # GSR and PPG
            ts, gsr, ppg = gsr_samples(D["seed"], rec, tr, tl, gp, win0, win1)
            for t, g, p in zip(ts, gsr, ppg):
                row = [""] * width
                row[1], row[7], row[19], row[20], row[21] = t, s, "1", f"{g:.4f}", f"{p:.1f}"
                if s in ADS and ad_on + ENDCARD[0] <= t < ad_on + ENDCARD[1]:
                    row[24] = "1"
                out.append((t, 3, row))
        out.sort(key=lambda z: (z[0], z[1]))
        rec_start = rec["start"]
        info = [["#INFO"], ["#Study name", STUDY], ["#Respondent Name", rec["name"], f"Id: {rec['guid']}"],
                ["#Respondent Age", str(rec["age_field"])], ["#Respondent Gender", rec["gender_field"]],
                ["#Respondent Group", rec["group_field"]], ["#iMotions version", IMOTIONS_VERSION],
                ["#Current culture", "en-US English (United States)", "File encoding: System.Text.UTF8Encoding", "Date format: M/d/yyyy"],
                ["#Recording time", f"Date: {rec_start.month}/{rec_start.day}/{rec_start.year}",
                 f"Time: {rec_start.strftime('%H:%M:%S')}.{int(rng_for(D['seed'], 11, k).integers(0, 999)):03d} -06:00",
                 f"Unix time: {int(unix_ms(rec_start))}"],
                ["#Export time", f"Date: {LATE_EXPORT.month}/{LATE_EXPORT.day}/{LATE_EXPORT.year}",
                 f"Time: {(LATE_EXPORT + timedelta(seconds=7 * k)).strftime('%H:%M:%S')}.512 -06:00",
                 f"Unix time: {int(unix_ms(LATE_EXPORT + timedelta(seconds=7 * k)))}"],
                ["#Sensor info", "Eyetracker Smart Eye Aurora", "Latency: 0 milliseconds"],
                ["#Sensor info", "Affectiva AFFDEX", "Latency: 0 milliseconds"],
                ["#Sensor info", f"{rec['shimmer']} GSR", "Latency: 0 milliseconds"], ["#"]]
        meta = sd_metadata(rec["shimmer"]) + [["#"], ["#DATA"]]
        fname = f"{rec_pos[k]:03d}_{rec['name']}.csv"
        path = os.path.join(folder, fname)
        ban = banner(D["seed"], "Sensor Data export (one respondent). Layout: #INFO, #METADATA, #DATA, header, then one row per device sample, "
                     "devices interleaved by Timestamp (ms since the recording started), not resampled. MODELED channel names and values; see README.")
        with open(path, "w", newline="", encoding="utf-8") as fh:
            fh.write(ban + "\n")
            w = csv.writer(fh, lineterminator="\n")
            for r in info + meta:
                w.writerow(r + [""] * (width - len(r)))
            w.writerow(SD_COLS)
            for n_, (t, _, row) in enumerate(out, start=1):
                row[0] = str(n_)
                row[1] = f"{t:.4f}"
                w.writerow(row)
        sizes[fname] = (os.path.getsize(path), len(out))
    return sizes


def write_all(D):
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    folders = {k: ensure(os.path.join(OUT, k)) for k in ("01_Sensor_Data", "02_AOI_Metrics", "03_FEA_Metrics", "04_GSR_Metrics",
                                                          "05_Survey", "06_Data_Quality", "07_Annotations", "08_Session_Log")}
    print("wrote GSRPeakMetrics.csv rows:", write_gsr(D, folders["04_GSR_Metrics"]))
    print("wrote FEAMetrics.csv rows:", write_fea(D, folders["03_FEA_Metrics"]))
    print("wrote AOI metrics rows (individual, aggregated):", write_aoi(D, folders["02_AOI_Metrics"]))
    print("wrote survey rows (respondents, statistics):", write_survey(D, folders["05_Survey"]))
    print("wrote data quality rows:", write_dq(D, folders["06_Data_Quality"]))
    print("wrote annotation rows (individual, aggregate):", write_annotations(D, folders["07_Annotations"]))
    print("wrote session log rows:", write_session_log(D, folders["08_Session_Log"]))
    for f, (sz, n) in write_sensor_data(D, folders["01_Sensor_Data"]).items():
        print(f"wrote 01_Sensor_Data/{f}: {n} data rows, {sz / 1e6:.2f} MB")
        assert sz < 2_000_000, f"{f} is larger than 2 MB"


# ---------------------------------------------------------------------------------------------
# Plan (population) and seed scan
# ---------------------------------------------------------------------------------------------
def calibrate(reps=200):
    """Average the realized statistics over many simulated studies of the same size (light mode):
    the expected values of the plan. Seeds are drawn far away from the scan range."""
    acc = {}
    for s in range(reps):
        o = realized(generate(9_000_000 + s, full=False))
        for k, v in o.items():
            acc.setdefault(k, []).append(v)
    print(f"plan: mean (SD) of each statistic over {reps} simulated studies of 24 participants")
    for k, v in acc.items():
        print(f"  {k:<14} {np.mean(v):9.3f}  ({np.std(v):.3f})")


def _scan_one(s):
    return loss(realized(generate(s, full=False))), s


def scan(n):
    from multiprocessing import Pool
    with Pool(4) as pool:
        res = pool.map(_scan_one, range(SCAN_BASE, SCAN_BASE + n), chunksize=4)
    res.sort()
    print("rank  seed        loss (sum of squared standardized gaps to the plan; p values play no part)")
    for k, (l_, s) in enumerate(res[:10], start=1):
        print(f"{k:>4}  {s}  {l_:9.2f}")
    print(f"median loss of all {n} seeds: {np.median([l_ for l_, _ in res]):.2f}")


def main():
    args = sys.argv[1:]
    if args and args[0] == "--scan":
        scan(int(args[1]) if len(args) > 1 else 200)
        return
    if args and args[0] == "--calibrate":
        calibrate(int(args[1]) if len(args) > 1 else 200)
        return
    seed = int(args[0]) if args else SEED
    D = generate(seed, full=True)
    write_all(D)
    o = realized(D)
    print(f"\nRealized statistics (seed {seed}); loss {loss(o):.2f}")
    for k, v in o.items():
        t = TARGETS.get(k)
        print(f"  {k:<14} {v:10.4f}" + (f"   plan {t[0]}" if t else ""))


if __name__ == "__main__":
    main()
