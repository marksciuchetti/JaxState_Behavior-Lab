#!/usr/bin/env python3
"""
make_simulated_data.py
======================
SIMULATED DATA GENERATOR for Workshop 6: remote studies with iMotions. Webcam eye tracking and
webcam facial expression analysis (Affectiva AFFDEX) collected online through Remote Data Collection
(RDC), built in iMotions Lab 11.1, and processed in iMotions Lab after the download.
Martin's Family Clothing Consumer Behavior Research Laboratory, Jacksonville State University

EVERYTHING THIS SCRIPT WRITES IS SIMULATED. No person was recorded and no webcam was switched on.
The numbers are drawn from the probability model documented below so that students can practice the
data quality steps and the analysis of a remote study on files that look like what the lab would
download and export.

The design copies the logic of the iMotions vendor demo "(DEMO) Food Waste_webET_FEA_voice" (help
center article "Demo study: Food Waste (image, videos and talkaloud)"): a social marketing study in
which every respondent saw two one minute food waste videos, a narrative one (foodwaste_narrative,
60 s) and an emotional visual one (foodwaste_visual, 62 s), in random order, then an infographic
(15 s) with four tips to reduce food waste, then a questionnaire. The demo article reports that, with
AFFDEX thresholding at 50, the visual video showed engagement 57 percent of the time and the narrative
video 19 percent. Nothing from the demo's respondents was used. The demo also has a talkaloud task
with voice analysis; this simulated study leaves the microphone off (the lab rule is microphone off
unless the study needs voice, and the lab has no voice license).

Study
-----
Study Name "SIMULATED Food Waste Remote Study", type Remote Data Collection, sensors "Eye tracking
(Web Camera)" and "Respondent camera" (the RDC Select sensors page on the lab computer). Default
animated 13 point "Pre-study calibration" and "Post-study calibration" (39 s each). An iMotions
survey ("Questionnaire"). Screen recording and microphone off. Facial expressions were computed in
iMotions Lab by Post Processing > Affectiva after the study zip was downloaded and loaded, then FEA
thresholding at 50; gaze on the infographic was analyzed with five AOIs in iMotions Lab.

Flow, two study flows (the order of the two videos)
    Pre-study calibration (39 s), Instructions1 (30 s), first video, second video, infographic (15 s),
    Questionnaire (self paced), Post-study calibration (39 s)
    "Study flow 1": foodwaste_narrative first.   "Study flow 2": foodwaste_visual first.
The demo randomized the two videos inside one block. Here each order is its own flow, because the
documented "Block" column of the session data (Accessing RDC Study Data: "Block - Study flow name")
then records the order of EVERY session, including sessions that stopped before the videos and never
uploaded data. A differential attrition check needs that. MODELED: the invitation list was split at
random into two halves and each half received the link of one flow (the RDC Data Collection page
lists one LINK per BLOCK). Verify on the lab computer that an RDC study built in Lab 11.1 can hold two
flows, each with its own link.

Recruitment funnel (session level)
----------------------------------
260 sessions opened a link. Each session passes the respondent funnel steps exactly as the iMotions
Online data collection article names them (Respondent funnel, Screenshot): "STEP 1 Testing upload",
"STEP 2 Consent screen", "STEP 3 Screen recording setup", "STEP 4 Respondent camera setup", "STEP 5
Screen size", "STEP 6 Enter fullscreen", "STEP 7 Prepare positioning", "STEP 8 Positioning check",
"STEP 9 Stimuli presentation*", "STEP 10 Study data upload". The calibrations, both videos, the
infographic, and the questionnaire all sit inside STEP 9. MODELED: that an RDC study shows the same
ten steps (the RDC article shows a "Respondent funnel" tab but not its steps), that screen recording
off means nobody stops at STEP 3, and the stop probabilities in A1.

Files written (same folder)
---------------------------
SIMULATED_session_data.csv
    One row per session (260). DOCUMENTED column names, from the respondent list of an RDC study on
    my.imotions.com (Accessing RDC Study Data): Id, Started, Ended, Eye tracking, Est. face %, Block,
    Status. "Eye tracking" holds "accuracy, quality%" as the screen shows it (for example "0.4, 91%";
    the RDC article calls the first number Precision, the Online articles call it accuracy). Status
    values as documented: Processed, Abandoned, Errors. The on screen columns Respondent, Study, and
    Machine (Lab Management fields only) are left out. MODELED: the CSV layout of "Download session
    data" (the manuals show the screen, not the file), the Id format RESP0001 (the redirect token
    example is RESP0005), which calibration the session level accuracy summarizes (here the mean of
    the pre and post study values, one decimal) and which slides the quality covers (here the mean
    over all slides), Status "Abandoned" for every session that stopped before the end, and three
    MODELED columns that the manuals do not list for the session data: "Last funnel step (modeled)"
    (the step at which the session stopped, or "Completed"; the portal shows the funnel only as a
    chart), "Device type (modeled)", and "Browser (modeled)".
SIMULATED_quality.csv
    MODELED layout, one row per completed respondent and slide, in flow order. It copies the "Quality
    metrics" table of the iMotions Online study session page (columns SLIDE, EYE TRACKING, FACE %;
    calibration rows carry accuracy and percent, other rows percent only) with EYE TRACKING split into
    two numbers: Respondent Name, Block, Slide position, Slide, Eye tracking accuracy (degrees), Eye
    tracking quality (%), Face quality (%). Definitions from the Online data quality box: accuracy =
    mean distance between actual and expected gaze at calibration in degrees of visual angle; eye
    tracking quality = valid samples / all samples; face quality = share of samples with a face found.
    For an RDC study these numbers may have to be collected from the portal and from the Lab Exposure
    summary or data quality export; verify where each lives.
SIMULATED_fea_metrics.csv
    FEA Metrics Export style, the same MODELED header as Workshop 3: Study Name, Respondent Name,
    Interval type, Interval label, Interval duration (s), Expression, Threshold, Percent of time
    present, Number of segments, Mean raw score (0-100), Data quality (%). One row per completed
    respondent, video, and expression: Engagement, Joy, Brow Furrow (threshold 50), Positive Valence
    (+50), Negative Valence (-50). Workshop 3 reported Engagement as a raw mean because the
    thresholding article names no default for it; the Food Waste demo article reports Engagement as
    percent of time at a threshold of 50, so this workshop follows the demo (verify that the 11.1
    thresholding notebook offers Engagement, for example under "Other Measures thresholding"). Mean
    raw score is NA on thresholded rows (as in Workshop 3). Data quality (%) = face quality of that
    video, the same number as in SIMULATED_quality.csv.
SIMULATED_aoi_metrics.csv
    Individual AOI metrics export style, as in Workshop 2: identifier columns Study Name, Respondent
    Name, Stimulus, AOI Name (MODELED order), then metric names from the iMotions AOI Metrics article,
    which are also the names in the iMotions Online AOI menus: Stimulus duration, AOI duration (ms),
    AOI duration (%), Size (px2), Size (%), Valid data, Dwells with fixations, Revisit count, Fixation
    count, TTFF AOI (ms), TTFF parent (ms), Dwell time (ms), Dwell time (%), Fixation duration (ms),
    First fixation duration (ms), Last fixation duration (ms). One row per completed respondent and
    AOI on the infographic. A never fixated AOI keeps its row with the fixation metrics NA (Workshop 2
    convention). "Respondent ratio (%)" is an aggregate metric ("Percentage of respondents who fixated
    on the AOI"), so it is not a column here: the analysis computes it as the percent of respondents
    with a hit (TTFF AOI not NA). Valid data = eye tracking quality on the infographic.
SIMULATED_survey.csv
    MODELED survey export (the RDC survey article says only that results come from "the survey data
    export option"): Study Name, Respondent Name, Age, Gender, Intention_reduce_food_waste_1to7,
    More_persuasive_video (Narrative or Visual), Tips_recalled_0to4.

Screen and AOIs
---------------
RDC fixes the screen settings at 1920 x 1080 pixels and 47.4 x 29.6 cm at 65 cm (Using the iMotions
Study Builder), so 1 degree of visual angle is about 46 px. The infographic AOIs are large, with gaps
between them, as the Online introduction advises for webcam data ("large objects that have a buffer
of space between them"):
    Headline  x 160 to 1760, y 70 to 180     (1600 x 110 px, 176000 px2, 8.49 percent of the screen:
                                              the smallest AOI, a wide but thin banner about 2.4
                                              degrees tall)
    Tip_1     x 80 to 900,   y 300 to 620    (820 x 320 px, 262400 px2, 12.65 percent each; Tip_1 top
    Tip_2     x 1020 to 1840, y 300 to 620    left, Tip_2 top right, Tip_3 bottom left, Tip_4 bottom
    Tip_3     x 80 to 900,   y 700 to 1020    right; equal sizes so that their dwell times can be
    Tip_4     x 1020 to 1840, y 700 to 1020   compared fairly; 120 px, about 2.6 degrees, between
                                              the Headline and the top tips)

Model assumptions
-----------------
A1  Funnel. Device (MODELED): Laptop 63, Desktop 25, Phone 10, Tablet 2 percent. Browser (MODELED):
    computers Chrome 60, Edge 16, Safari 14, Firefox 10; phones and tablets Safari or Chrome. Stop
    probability at each step, given the step was reached: 1 Testing upload .012; 2 Consent screen
    .045; 3 Screen recording setup 0 (screen recording off); 4 Respondent camera setup .05 (camera
    permission refused or no camera); 5 Screen size .008 on computers and 1.0 on phones and tablets
    (the RDC article says "Participation can not be done on phones or tablets"; WHERE a phone is
    stopped is not stated, and STEP 5 is chosen because its screen type choices are Laptop and Desktop:
    verify with a phone during the pilot); 6 Enter fullscreen .015; 7 Prepare positioning .02;
    8 Positioning check .035 ("Allow skipping after timeout" off); 9 Stimuli presentation* .045 (left
    during a calibration, a video, the infographic, or the questionnaire; the same for both flows);
    10 Study data upload .03 (two thirds Abandoned, one third Errors). Everyone else is Processed.
    Planned completion: about 67.5 percent of the sessions that open a link.
A2  Calibration accuracy (completed respondents). A setup trait s ~ N(0, 1) (light, camera, movement;
    higher is worse) drives every quality value. Pre-study accuracy (degrees) is lognormal, median
    1.42, log SD 0.55, with 36 percent of its log variance from s. Drift: post = pre + d, with
    d ~ N(0.72, 0.52) and correlation .45 with s, d at least -0.7. Accuracy while the infographic was on
    screen is pre + 0.8 d (it came late in the flow). Values are rounded to 0.01 degree.
A3  Eye tracking quality per slide = 100 - missing percent, missing = exp(ln 3.4 + 0.80 s + 0.35 e),
    capped at 60; face quality per slide = 100 - exp(ln 2.4 + 0.55 s + 1.10 f + 0.35 e), with a face
    trait f ~ N(0, 1) (glasses glare, hands, hair), capped at 70 missing. Rounded to 0.1. Together
    with A2 the rule excludes about 19.5 percent of the completed respondents (planned: Pre-study
    accuracy 2.8, Post-study accuracy 6.1, eye tracking quality 6.6, face quality 4.0 percent, by
    first reason), and the 3.0 degree threshold removes about 10.8 percent more.
A4  Facial expressions. For each respondent, video, and expression, the share of time above the
    threshold is p = invlogit(a + b), a = the planned intercept of that video and expression (ALPHA
    below), b = sd x (rho x trait + sqrt(1 - rho^2) x noise) - 0.12 when the video was seen second
    (MODELED fatigue). Traits: Engagement uses a general expressiveness g; Joy and Positive Valence
    share a positive trait (correlation .45 with g); Brow Furrow and Negative Valence share a negative
    trait (correlation .45 with g). The 30 Hz frame sequence is a two state chain with stationary share
    p and mean episode length 3.0 s (Engagement), 1.6 s (Joy), 2.0 s (Brow Furrow), 2.0 s (Positive
    Valence), 1.8 s (Negative Valence), so low p gives many zeros, as real percent of time data do.
    Frames without a face (from A3) count as absent; Percent of time present divides by all frames of
    the video (assumption to verify, as in Workshop 3). Number of segments counts runs of present
    frames. Planned Engagement (kept respondents): narrative 20.4, visual 44.9 percent of the time,
    d_z 0.89; about 16 percent of respondents at exactly 0 during the narrative video. Planned Joy 2.7
    and 4.2, Brow Furrow 3.3 and 9.3, Positive Valence 3.7 and 6.4, Negative Valence 2.6 and 7.8
    (narrative and visual); like the demo, the visual video shows more of both valences.
A5  Infographic gaze. A fixation sequence over the 15 s (durations lognormal, median 300 ms; 30 to
    60 ms between fixations). The first fixation goes to the Headline (40 percent), Tip_1 (33), or
    elsewhere. A fixation stays on the same AOI with probability .68 (tips), .30 (Headline), .35
    (elsewhere); otherwise the next unvisited AOI in reading order (Headline, Tip_1, Tip_2, Tip_3,
    Tip_4) is chosen with probability .75, else an AOI by interest weights exp(mu + 0.55 x noise),
    mu = Tip_1 1.30, Tip_2 0.0, Tip_3 0.03, Tip_4 0.08, Headline -1.9, and "elsewhere" -0.6 - 1.7 x
    attention (attention ~ N(0, 1)); "elsewhere" points lie outside every AOI and below the Headline
    band. The recorded location is the intended point plus a respondent offset (random direction,
    length 0.85 x accuracy) plus noise (SD 0.45 x accuracy per axis), with the accuracy from A2 at the
    infographic; a recorded point outside every AOI is no hit, and a point inside another AOI is
    counted there (misattribution). Missing samples (from A3) remove the fixations whose middle falls
    in one of one to three gaps. Planned (kept respondents): dwell time Tip_1 3692, Tip_2 2123, Tip_3
    2119, Tip_4 2047 ms; Kendall's W 0.16; paired d_z of Tip_1 against each other tip about 0.67 to
    0.71 and of the other pairs 0.04 or less; Headline hit ratio 70 percent, falling from about 81
    percent at a mean accuracy of 1.5 degrees or less to about 55 percent above 3.5 degrees.
A6  Survey (completed respondents). Intention to reduce food waste (1 to 7): round(5.35 + 1.35 x z),
    z = 0.30 x the standardized visual engagement propensity + 0.954 x noise, clipped (planned M 5.30,
    SD 1.28 after rounding and clipping); NO effect of the flow (H3 is a true null). More persuasive
    video: Visual with probability invlogit(0.50 + 0.70 x the standardized engagement difference,
    visual minus narrative) (planned 61 percent Visual; point biserial r with the realized difference
    about .25). Tips recalled ~ Binomial(4, invlogit(0.15 + 1.30 x attention)) (planned Spearman rho
    with total tip dwell about .20). Age: 60 percent 18 to 24, the rest 25 to 64. Gender F 57, M 41,
    "Prefer not to say" 2 percent.
A7  Session times: ten weekdays from 2027-03-01, 09:00 to 23:00 (more in the evening); Ids RESP0001
    to RESP0260 in the order the link was opened. Ended is blank unless the upload finished.

Exclusion rule (written before collection; thresholds from the manuals)
    Exclude a completed respondent when the pre-study OR the post-study calibration accuracy is above
    4.0 degrees (iMotions Online Study Setup: "Generally acceptable range: 0-4 dva"), when eye
    tracking quality on the infographic is below 90 percent ("We recommend a quality score of at least
    90%"), or when face quality is below 80 percent during either video ("at least 80%"; also the lab
    rule of Workshop 3). Sensitivity: accuracy at or below 3.0 degrees at both calibrations (the Data
    collection and Analysis articles: "We recommend a threshold of 3.0").

Run:  python3 make_simulated_data.py                 (default seed below)
      python3 make_simulated_data.py <seed>          (any other seed)
      python3 make_simulated_data.py --scan [n]      (rank n candidate seeds; writes nothing)
      python3 make_simulated_data.py --calibrate     (population check of the model; writes nothing)
Requires: Python 3, numpy, scipy.
"""

import csv
import math
import os
import statistics as st
import sys
from datetime import datetime, timedelta, date

import numpy as np
from scipy import stats

HERE = os.path.dirname(os.path.abspath(__file__))

# Fixed seed so the files are reproducible. It was chosen from a scan of 2400 candidate seeds, 20276001
# to 20278400 (python3 make_simulated_data.py --scan 2400; 677 of them met the hard ranges of the plan:
# 170 to 180 completed, 135 to 145 kept, and 15 to 25 more removed at 3.0 degrees), as the one whose
# realized sample sits closest to the planned MODEL values in TARGETS below (loss 12.06; the median of
# the 677 usable seeds is 32.84). The scan ranks seeds by closeness to the model, never by a p value.
# Any single sample still drifts from its model (here, for example, intention is 0.30 points lower in
# the Narrative first order although the model has no order effect, and the recall correlation is
# weaker than planned), which is itself a lesson: compare the realized values with the assumptions above.
SEED = 20276200
SCAN_BASE = 20276001

STUDY_NAME = "SIMULATED Food Waste Remote Study"
N_OPEN = 260
FLOWS = ("Study flow 1", "Study flow 2")
NARR, VIS = "foodwaste_narrative", "foodwaste_visual"
VIDEOS = (NARR, VIS)
FIRST = {"Study flow 1": NARR, "Study flow 2": VIS}
VID_DUR = {NARR: 60.0, VIS: 62.0}
INFO_MS = 15000
FPS = 30


def flow_slides(flow):
    a, b = (NARR, VIS) if flow == "Study flow 1" else (VIS, NARR)
    return ["Pre-study calibration", "Instructions1", a, b, "infographic", "Questionnaire", "Post-study calibration"]


STEPS = ["STEP 1 Testing upload", "STEP 2 Consent screen", "STEP 3 Screen recording setup",
         "STEP 4 Respondent camera setup", "STEP 5 Screen size", "STEP 6 Enter fullscreen",
         "STEP 7 Prepare positioning", "STEP 8 Positioning check", "STEP 9 Stimuli presentation*",
         "STEP 10 Study data upload"]
HAZARD = [0.012, 0.045, 0.0, 0.05, 0.008, 0.015, 0.02, 0.035, 0.045, 0.03]
STEP_SECONDS = [4, 25, 0, 30, 15, 5, 20, 25, None, None]       # time spent on each setup step (median)
DEVICES = [("Laptop", 0.63), ("Desktop", 0.25), ("Phone", 0.10), ("Tablet", 0.02)]
BROWSER_PC = [("Chrome", 0.60), ("Edge", 0.16), ("Safari", 0.14), ("Firefox", 0.10)]
BROWSER_MOBILE = [("Safari", 0.60), ("Chrome", 0.40)]

# A2 and A3: quality model
PRE_MED, PRE_LOGSD, PRE_S = 1.42, 0.55, 0.6
DRIFT_MEAN, DRIFT_SD, DRIFT_S, DRIFT_MIN = 0.72, 0.52, 0.45, -0.7
INFO_DRIFT_SHARE = 0.8
ET_MISS_MED, ET_MISS_S, ET_MISS_E, ET_MISS_CAP = 3.4, 0.80, 0.35, 60.0
FACE_MISS_MED, FACE_MISS_S, FACE_MISS_F, FACE_MISS_E, FACE_MISS_CAP = 2.4, 0.55, 1.10, 0.35, 70.0
ACC_MAX, ACC_STRICT, ETQ_MIN, FACE_MIN = 4.0, 3.0, 90.0, 80.0

# A4: facial expressions
EXPR = ["Engagement", "Joy", "Brow Furrow", "Positive Valence", "Negative Valence"]
THRESH = {"Engagement": 50, "Joy": 50, "Brow Furrow": 50, "Positive Valence": 50, "Negative Valence": -50}
ALPHA = {  # planned logit intercepts (video, expression); set with --calibrate
    NARR: {"Engagement": -1.65, "Joy": -4.10, "Brow Furrow": -3.75, "Positive Valence": -3.70, "Negative Valence": -4.10},
    VIS:  {"Engagement": -0.11, "Joy": -3.55, "Brow Furrow": -2.65, "Positive Valence": -3.05, "Negative Valence": -2.80},
}
SIG = {"Engagement": 1.35, "Joy": 1.10, "Brow Furrow": 1.10, "Positive Valence": 1.10, "Negative Valence": 1.10}
RHO = {"Engagement": 0.66, "Joy": 0.70, "Brow Furrow": 0.70, "Positive Valence": 0.70, "Negative Valence": 0.70}
EPISODE_S = {"Engagement": 3.0, "Joy": 1.6, "Brow Furrow": 2.0, "Positive Valence": 2.0, "Negative Valence": 1.8}
POS_SECOND = -0.12
TRAIT_G = 0.45

# A5: infographic gaze
SCREEN_W, SCREEN_H = 1920, 1080
PX_PER_DEG = (1920 / 47.4) * 65.0 * math.tan(math.radians(1.0))      # about 45.96
AOIS = {"Headline": (160, 70, 1760, 180), "Tip_1": (80, 300, 900, 620), "Tip_2": (1020, 300, 1840, 620),
        "Tip_3": (80, 700, 900, 1020), "Tip_4": (1020, 700, 1840, 1020)}
AOI_NAMES = list(AOIS)
TIPS = AOI_NAMES[1:]
READ_ORDER = AOI_NAMES
MU = {"Headline": -1.9, "Tip_1": 1.30, "Tip_2": 0.0, "Tip_3": 0.03, "Tip_4": 0.08}
MU_ELSE, ATT_ELSE = -0.6, 1.70
W_NOISE = 0.55
P_STAY = {"Headline": 0.30, "Tip_1": 0.68, "Tip_2": 0.68, "Tip_3": 0.68, "Tip_4": 0.68, "else": 0.35}
P_FIRST = (0.40, 0.33)                 # Headline, Tip_1; the rest elsewhere
P_NEXT_UNVISITED = 0.75
BIAS_SHARE, NOISE_SHARE = 0.85, 0.45

# A6: survey
INT_MEAN, INT_SD, INT_W = 5.35, 1.35, 0.30
PREF_A, PREF_B = 0.50, 0.70
REC_A, REC_B = 0.15, 1.30

DAY0 = date(2027, 3, 1)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def clip(x, lo, hi):
    return max(lo, min(hi, x))


def invlogit(x):
    return 1.0 / (1.0 + np.exp(-x))


def pick(rng, pairs):
    names, probs = zip(*pairs)
    return names[int(rng.choice(len(names), p=np.array(probs) / sum(probs)))]


def weekdays(start, n):
    out, d = [], start
    while len(out) < n:
        if d.weekday() < 5:
            out.append(d)
        d += timedelta(days=1)
    return out


def area_pct(r):
    x0, y0, x1, y1 = r
    return (x1 - x0) * (y1 - y0), 100.0 * (x1 - x0) * (y1 - y0) / (SCREEN_W * SCREEN_H)


def classify(x, y):
    for name, (x0, y0, x1, y1) in AOIS.items():
        if x0 <= x <= x1 and y0 <= y <= y1:
            return name
    return None


def runs(mask):
    """Number of runs of True in a 1D boolean array."""
    if mask.size == 0:
        return 0
    m = mask.astype(np.int8)
    return int(m[0] + np.sum((m[1:] == 1) & (m[:-1] == 0)))


def missing_mask(rng, n, share):
    """A boolean mask of n frames with about share x n missing frames placed in a few runs; the
    realized share is what the export reports."""
    k_miss = int(round(share * n))
    mask = np.zeros(n, dtype=bool)
    if k_miss <= 0:
        return mask
    k_runs = 1 + int(rng.poisson(1.0 + k_miss / 120.0))
    cuts = np.sort(rng.choice(np.arange(1, k_miss), size=min(k_runs - 1, max(k_miss - 1, 0)), replace=False)) if k_miss > 1 else np.array([], int)
    lens = np.diff(np.concatenate([[0], cuts, [k_miss]]))
    for L in lens:
        free = np.flatnonzero(~mask)
        if L <= 0 or free.size == 0:
            continue
        start = int(rng.integers(0, max(1, n - L + 1)))
        mask[start:start + L] = True
    # top up or trim so the count is exact
    diff = k_miss - int(mask.sum())
    if diff > 0:
        free = np.flatnonzero(~mask)
        mask[rng.choice(free, size=diff, replace=False)] = True
    elif diff < 0:
        on = np.flatnonzero(mask)
        mask[rng.choice(on, size=-diff, replace=False)] = False
    return mask


# ---------------------------------------------------------------------------
# A4: frame level facial expression series for many respondents at once
# ---------------------------------------------------------------------------
def fea_series(rng, p, n_frames, episode_s):
    """p: array (n,) of stationary shares. Two state chain at FPS, mean on episode = episode_s.
    Returns a boolean array (n, n_frames)."""
    b = 1.0 / (episode_s * FPS)
    p = np.clip(p, 1e-6, 0.995)
    a = p * b / (1.0 - p)
    a = np.minimum(a, 0.999)
    n = p.size
    out = np.zeros((n, n_frames), dtype=bool)
    state = rng.random(n) < p
    u = rng.random((n, n_frames))
    for t in range(n_frames):
        out[:, t] = state
        turn_on = (~state) & (u[:, t] < a)
        turn_off = state & (u[:, t] < b)
        state = (state | turn_on) & ~turn_off
    return out


# ---------------------------------------------------------------------------
# A5: one infographic viewing
# ---------------------------------------------------------------------------
def rect_point(rng, r, margin=0.08):
    x0, y0, x1, y1 = r
    mx, my = margin * (x1 - x0), margin * (y1 - y0)
    return rng.uniform(x0 + mx, x1 - mx), rng.uniform(y0 + my, y1 - my)


def else_point(rng):
    """A point outside every AOI, below the Headline band (the gaps between the tips and the margins)."""
    for _ in range(100):
        x, y = rng.uniform(0, SCREEN_W), rng.uniform(240, SCREEN_H)
        if classify(x, y) is None:
            return x, y
    return 960.0, 230.0


def infographic(rng, w, acc_deg, et_missing_pct):
    """w: dict of interest weights (Tip_1..Tip_4, Headline, else). Returns a list of
    (start_ms, dur_ms, recorded_aoi) for the recorded fixations."""
    a_px = acc_deg * PX_PER_DEG
    ang = rng.uniform(0, 2 * math.pi)
    bx, by = BIAS_SHARE * a_px * math.cos(ang), BIAS_SHARE * a_px * math.sin(ang)
    sd = NOISE_SHARE * a_px
    t = rng.uniform(150, 350)
    fix = []
    cur, visited = None, set()
    while True:
        dur = clip(300.0 * math.exp(0.38 * rng.standard_normal()), 100, 1200)
        if t + dur > INFO_MS:
            dur = INFO_MS - t
            if dur < 80:
                break
        if cur is None:
            r = rng.random()
            target = "Headline" if r < P_FIRST[0] else ("Tip_1" if r < P_FIRST[0] + P_FIRST[1] else "else")
        elif rng.random() < P_STAY[cur]:
            target = cur
        else:
            nxt = next((a for a in READ_ORDER if a not in visited and a != cur), None)
            if nxt is not None and rng.random() < P_NEXT_UNVISITED:
                target = nxt
            else:
                opts = [a for a in AOI_NAMES + ["else"] if a != cur]
                ws = np.array([w[a] for a in opts])
                target = opts[int(rng.choice(len(opts), p=ws / ws.sum()))]
        visited.add(target)
        x, y = else_point(rng) if target == "else" else rect_point(rng, AOIS[target])
        rx, ry = x + bx + sd * rng.standard_normal(), y + by + sd * rng.standard_normal()
        fix.append((t, dur, classify(rx, ry)))
        cur = target
        t += dur + rng.uniform(30, 60)
        if t >= INFO_MS - 80:
            break
    # missing samples: one to three gaps; a fixation whose middle falls in a gap is lost
    miss_ms = et_missing_pct / 100.0 * INFO_MS
    gaps = []
    if miss_ms > 0:
        k = 1 + int(rng.integers(0, 3))
        parts = rng.dirichlet(np.ones(k)) * miss_ms
        for L in parts:
            s0 = rng.uniform(0, INFO_MS - L)
            gaps.append((s0, s0 + L))
    kept = [f for f in fix if not any(g0 <= f[0] + f[1] / 2 <= g1 for g0, g1 in gaps)]
    return kept


def aoi_metrics(fix):
    """Per AOI metrics from a recorded fixation list. A dwell is a run of consecutive recorded
    fixations in the same AOI. Dwell time (ms) = sum of the fixation durations (Workshop 2)."""
    out = {}
    for name in AOI_NAMES:
        idx = [k for k, f in enumerate(fix) if f[2] == name]
        if not idx:
            out[name] = None
            continue
        durs = [fix[k][1] for k in idx]
        dwells = 1 + sum(1 for a, b in zip(idx[:-1], idx[1:]) if b != a + 1)
        dwell = int(round(sum(durs)))
        out[name] = {"dwells": dwells, "revisits": dwells - 1, "count": len(idx), "ttff": int(round(fix[idx[0]][0])),
                     "dwell": dwell, "dwell_pct": 100.0 * dwell / INFO_MS, "fixdur": dwell / len(idx),
                     "first": int(round(durs[0])), "last": int(round(durs[-1]))}
    return out


# ---------------------------------------------------------------------------
# the generator
# ---------------------------------------------------------------------------
def generate(seed, n_open=N_OPEN, population=False):
    """population=True skips the funnel (everyone completes) for --calibrate."""
    rng = np.random.default_rng(seed)
    S = []                                          # sessions
    days = weekdays(DAY0, 10)
    for k in range(n_open):
        dev = pick(rng, DEVICES)
        brw = pick(rng, BROWSER_PC if dev in ("Laptop", "Desktop") else BROWSER_MOBILE)
        flow = FLOWS[int(rng.random() < 0.5)]
        day = days[int(rng.integers(0, len(days)))]
        hour = int(rng.choice(np.arange(9, 23), p=np.array([2, 2, 3, 3, 3, 3, 3, 4, 5, 6, 7, 7, 6, 4], float) / 58))
        started = datetime(day.year, day.month, day.day, hour, int(rng.integers(0, 60)), int(rng.integers(0, 60)))
        stop = None
        if not population:
            for j, h in enumerate(HAZARD):
                hz = 1.0 if (j == 4 and dev in ("Phone", "Tablet")) else h
                if rng.random() < hz:
                    stop = j
                    break
        S.append({"device": dev, "browser": brw, "flow": flow, "started": started, "stop": stop})
    S.sort(key=lambda s: s["started"])
    for k, s in enumerate(S):
        s["id"] = f"RESP{k + 1:04d}"
        if s["stop"] is None:
            s["status"], s["last"] = "Processed", "Completed"
        else:
            s["last"] = STEPS[s["stop"]]
            s["status"] = ("Errors" if rng.random() < 1 / 3 else "Abandoned") if s["stop"] == 9 else "Abandoned"
    comp = [s for s in S if s["status"] == "Processed"]
    n = len(comp)

    # ---- A2 and A3: quality --------------------------------------------------
    sq = rng.standard_normal(n)
    fq = rng.standard_normal(n)
    pre = np.exp(math.log(PRE_MED) + PRE_LOGSD * (PRE_S * sq + math.sqrt(1 - PRE_S ** 2) * rng.standard_normal(n)))
    drift = DRIFT_MEAN + DRIFT_SD * (DRIFT_S * sq + math.sqrt(1 - DRIFT_S ** 2) * rng.standard_normal(n))
    drift = np.maximum(drift, DRIFT_MIN)
    pre = np.clip(pre, 0.25, 9.5)
    post = np.clip(pre + drift, 0.25, 9.9)
    acc_info = np.clip(pre + INFO_DRIFT_SHARE * (post - pre), 0.25, 9.9)
    for i, s in enumerate(comp):
        s["pre"], s["post"] = round(float(pre[i]), 2), round(float(post[i]), 2)
        s["acc_info"] = float(acc_info[i])
        slides = flow_slides(s["flow"])
        s["slides"] = slides
        etq, fcq = {}, {}
        for sl in slides:
            m_et = min(ET_MISS_CAP, math.exp(math.log(ET_MISS_MED) + ET_MISS_S * sq[i] + ET_MISS_E * rng.standard_normal()))
            etq[sl] = round(100.0 - m_et, 1)
            m_fc = min(FACE_MISS_CAP, math.exp(math.log(FACE_MISS_MED) + FACE_MISS_S * sq[i] + FACE_MISS_F * fq[i] + FACE_MISS_E * rng.standard_normal()))
            fcq[sl] = 100.0 - m_fc                     # replaced by the realized frame share for the videos
        s["etq"], s["faceq"] = etq, fcq

    # ---- A4: facial expressions ------------------------------------------------
    g = rng.standard_normal(n)
    t_pos = TRAIT_G * g + math.sqrt(1 - TRAIT_G ** 2) * rng.standard_normal(n)
    t_neg = TRAIT_G * g + math.sqrt(1 - TRAIT_G ** 2) * rng.standard_normal(n)
    trait = {"Engagement": g, "Joy": t_pos, "Positive Valence": 0.85 * t_pos + math.sqrt(1 - 0.85 ** 2) * rng.standard_normal(n),
             "Brow Furrow": t_neg, "Negative Valence": 0.85 * t_neg + math.sqrt(1 - 0.85 ** 2) * rng.standard_normal(n)}
    second = {v: np.array([1.0 if FIRST[s["flow"]] != v else 0.0 for s in comp]) for v in VIDEOS}
    logit = {}
    for v in VIDEOS:
        for x in EXPR:
            e = rng.standard_normal(n)
            logit[(v, x)] = ALPHA[v][x] + SIG[x] * (RHO[x] * trait[x] + math.sqrt(1 - RHO[x] ** 2) * e) + POS_SECOND * second[v]
    fea = {}
    for v in VIDEOS:
        nfr = int(round(VID_DUR[v] * FPS))
        masks = []
        for i, s in enumerate(comp):
            mm = missing_mask(rng, nfr, (100.0 - s["faceq"][v]) / 100.0)
            masks.append(mm)
            s["faceq"][v] = round(100.0 * (1 - mm.mean()), 1)
        masks = np.array(masks) if n else np.zeros((0, nfr), bool)
        for x in EXPR:
            on = fea_series(rng, invlogit(logit[(v, x)]), nfr, EPISODE_S[x]) if n else np.zeros((0, nfr), bool)
            present = on & ~masks
            for i, s in enumerate(comp):
                fea[(s["id"], v, x)] = (round(100.0 * present[i].mean(), 2), runs(present[i]))
    for s in comp:
        for sl in s["slides"]:
            s["faceq"][sl] = round(s["faceq"][sl], 1)

    # ---- A5: infographic -------------------------------------------------------
    att = rng.standard_normal(n)
    aoi = {}
    for i, s in enumerate(comp):
        w = {a: math.exp(MU[a] + W_NOISE * rng.standard_normal()) for a in AOI_NAMES}
        w["else"] = math.exp(MU_ELSE - ATT_ELSE * att[i] + W_NOISE * rng.standard_normal())
        fx = infographic(rng, w, s["acc_info"], 100.0 - s["etq"]["infographic"])
        aoi[s["id"]] = aoi_metrics(fx)

    # ---- A6: survey ------------------------------------------------------------
    zv = (logit[(VIS, "Engagement")] - logit[(VIS, "Engagement")].mean()) / logit[(VIS, "Engagement")].std() if n > 1 else np.zeros(n)
    dl = logit[(VIS, "Engagement")] - logit[(NARR, "Engagement")]
    zd = (dl - dl.mean()) / dl.std() if n > 1 else np.zeros(n)
    sv = {}
    for i, s in enumerate(comp):
        zi = INT_W * zv[i] + math.sqrt(1 - INT_W ** 2) * rng.standard_normal()
        inten = int(clip(round(INT_MEAN + INT_SD * zi), 1, 7))
        pref = "Visual" if rng.random() < invlogit(PREF_A + PREF_B * zd[i]) else "Narrative"
        rec = int(rng.binomial(4, invlogit(REC_A + REC_B * att[i])))
        if rng.random() < 0.60:
            age = int(rng.integers(18, 25))
        else:
            age = int(clip(round(25 + rng.gamma(2.0, 9.0)), 25, 64))
        gen = pick(rng, [("F", 0.57), ("M", 0.41), ("Prefer not to say", 0.02)])
        sv[s["id"]] = {"int": inten, "pref": pref, "rec": rec, "age": age, "gender": gen}

    # ---- A7: times ---------------------------------------------------------------
    for s in S:
        s["ended"] = None
        if s["status"] in ("Processed", "Errors"):
            setup = sum(max(0.0, x * math.exp(0.35 * rng.standard_normal())) for x in STEP_SECONDS[:8] if x)
            stim = 39 + 30 + 60 + 62 + 15 + 75 * math.exp(0.3 * rng.standard_normal()) + 39
            upload = 40 + 60 * rng.random()
            s["ended"] = s["started"] + timedelta(seconds=int(round(setup + stim + upload)))
    return {"seed": seed, "S": S, "comp": comp, "fea": fea, "aoi": aoi, "sv": sv}


# ---------------------------------------------------------------------------
# realized statistics (the logic mirrors analysis.R)
# ---------------------------------------------------------------------------
def screen(s, acc_limit=ACC_MAX):
    reasons = []
    if s["pre"] > acc_limit:
        reasons.append("pre")
    if s["post"] > acc_limit:
        reasons.append("post")
    if s["etq"]["infographic"] < ETQ_MIN:
        reasons.append("etq")
    if s["faceq"][NARR] < FACE_MIN or s["faceq"][VIS] < FACE_MIN:
        reasons.append("face")
    return reasons


def dz_paired(a, b):
    d = np.asarray(b, float) - np.asarray(a, float)
    return d.mean() / d.std(ddof=1)


def realized(data):
    o = {}
    S, comp = data["S"], data["comp"]
    o["n_open"] = len(S)
    o["n_complete"] = len(comp)
    fl = {f: [s for s in S if s["flow"] == f] for f in FLOWS}
    o["compl_diff"] = (np.mean([s["status"] == "Processed" for s in fl[FLOWS[0]]]) -
                       np.mean([s["status"] == "Processed" for s in fl[FLOWS[1]]]))
    st9 = {f: [s["status"] == "Processed" for s in fl[f] if s["stop"] is None or s["stop"] >= 8] for f in FLOWS}
    o["compl9_diff"] = float(np.mean(st9[FLOWS[0]]) - np.mean(st9[FLOWS[1]])) if all(st9.values()) else 0.0
    kept = [s for s in comp if not screen(s)]
    o["n_kept"] = len(kept)
    first = {"pre": 0, "post": 0, "etq": 0, "face": 0}
    for s in comp:
        r = screen(s)
        if r:
            first[r[0]] += 1
    for k, v in first.items():
        o["ex_" + k] = v
    strict = [s for s in kept if max(s["pre"], s["post"]) <= ACC_STRICT]
    o["n_strict_extra"] = len(kept) - len(strict)
    pre = np.array([s["pre"] for s in comp])
    post = np.array([s["post"] for s in comp])
    o["drift"] = float((post - pre).mean())
    o["drift_sd"] = float((post - pre).std(ddof=1))
    F, A, SV = data["fea"], data["aoi"], data["sv"]

    def eng(sub):
        return (np.array([F[(s["id"], NARR, "Engagement")][0] for s in sub]), np.array([F[(s["id"], VIS, "Engagement")][0] for s in sub]))
    en, ev = eng(kept)
    o["eng_n"], o["eng_v"] = float(en.mean()), float(ev.mean())
    o["eng_n_sd"], o["eng_v_sd"] = float(en.std(ddof=1)), float(ev.std(ddof=1))
    o["dz_h1"] = float(dz_paired(en, ev))
    o["zero_n"] = float(np.mean(en == 0))
    en2, ev2 = eng(strict)
    o["dz_h1_strict"] = float(dz_paired(en2, ev2))

    def dwell(sub):
        return np.array([[A[s["id"]][t]["dwell"] if A[s["id"]][t] else 0 for t in TIPS] for s in sub], float)
    D = dwell(kept)
    for j, t in enumerate(TIPS):
        o["tip" + t[-1]] = float(D[:, j].mean())
    fr = stats.friedmanchisquare(*[D[:, j] for j in range(4)])
    o["W"] = float(fr.statistic / (len(kept) * 3))
    ps = []
    for a in range(4):
        for b in range(a + 1, 4):
            o[f"dz_{a + 1}{b + 1}"] = float(dz_paired(D[:, b], D[:, a]))
            ps.append(stats.wilcoxon(D[:, a], D[:, b], correction=True, method="approx").pvalue)
    ps = np.array(ps)
    order_ = np.argsort(ps)
    adj = np.empty(6)
    run = 0.0
    for rank, idx in enumerate(order_):
        run = max(run, min(1.0, ps[idx] * (6 - rank)))
        adj[idx] = run
    o["n_sig_pairs"] = int((adj < 0.05).sum())
    D2 = dwell(strict)
    fr2 = stats.friedmanchisquare(*[D2[:, j] for j in range(4)])
    o["W_strict"] = float(fr2.statistic / (len(strict) * 3))
    o["hit_head"] = float(np.mean([A[s["id"]]["Headline"] is not None for s in kept]))
    # the hit analysis uses every completed respondent whose eye tracking and face quality passed
    qpass = [s for s in comp if not ({"etq", "face"} & set(screen(s)))]
    hit = np.array([A[s["id"]]["Headline"] is not None for s in qpass], float)
    macc = np.array([(s["pre"] + s["post"]) / 2 for s in qpass])
    xm = macc - macc.mean()
    # logistic slope by a few Newton steps (no extra package)
    beta = np.zeros(2)
    X = np.column_stack([np.ones_like(xm), xm])
    for _ in range(25):
        p = invlogit(X @ beta)
        Wt = p * (1 - p)
        H = X.T @ (X * Wt[:, None])
        beta = beta + np.linalg.solve(H + 1e-9 * np.eye(2), X.T @ (hit - p))
    o["hit_slope"] = float(beta[1])
    inten = {f: np.array([SV[s["id"]]["int"] for s in kept if s["flow"] == f], float) for f in FLOWS}
    o["int_diff"] = float(inten[FLOWS[0]].mean() - inten[FLOWS[1]].mean())
    allint = np.concatenate(list(inten.values()))
    o["int_mean"], o["int_sd"] = float(allint.mean()), float(allint.std(ddof=1))
    pref = np.array([SV[s["id"]]["pref"] == "Visual" for s in kept], float)
    o["pref_vis"] = float(pref.mean())
    o["r_pref"] = float(np.corrcoef(ev - en, pref)[0, 1])
    rec = np.array([SV[s["id"]]["rec"] for s in kept], float)
    o["rho_recall"] = float(stats.spearmanr(D.sum(axis=1), rec).statistic)
    o["n_flow1_kept"] = sum(1 for s in kept if s["flow"] == FLOWS[0])
    return o


# Target value and tolerance (about one standard error of the statistic in a sample of about 140).
# The planned values follow from the model (python3 make_simulated_data.py --calibrate).
TARGETS = {
    "n_complete": (175.6, 4.0), "n_kept": (141.0, 3.0), "n_strict_extra": (19.0, 3.0),
    "ex_pre": (4.9, 2.2), "ex_post": (10.7, 3.2), "ex_etq": (11.6, 3.3), "ex_face": (7.0, 2.6),
    "compl_diff": (0.0, 0.06), "compl9_diff": (0.0, 0.035),
    "eng_n": (20.4, 1.7), "eng_v": (44.9, 2.2), "dz_h1": (0.89, 0.07),
    "tip1": (3692.0, 130.0), "tip2": (2123.0, 130.0), "tip3": (2119.0, 130.0), "tip4": (2047.0, 130.0),
    "W": (0.162, 0.025),
    "dz_12": (0.667, 0.085), "dz_13": (0.676, 0.085), "dz_14": (0.713, 0.085),
    "dz_23": (0.002, 0.085), "dz_24": (0.037, 0.085), "dz_34": (0.035, 0.085),
    "drift": (0.713, 0.039), "drift_sd": (0.513, 0.030),
    "int_diff": (0.0, 0.21), "int_mean": (5.30, 0.10), "int_sd": (1.28, 0.075),
    "hit_head": (0.703, 0.039), "hit_slope": (-0.50, 0.17),
    "r_pref": (0.249, 0.08), "rho_recall": (0.198, 0.08),
}


def loss(o):
    hard = (170 <= o["n_complete"] <= 180) and (135 <= o["n_kept"] <= 145) and (15 <= o["n_strict_extra"] <= 25)
    if not hard:
        return 1e9
    return sum(((o[k] - t) / tol) ** 2 for k, (t, tol) in TARGETS.items())


# ---------------------------------------------------------------------------
# writing
# ---------------------------------------------------------------------------
def banner(seed, extra):
    return (f"# SIMULATED DATA for teaching. Generated by make_simulated_data.py (seed {seed}). "
            f"No real person was recorded. {extra}")


def half_up(x, nd):
    """Round half up (2.15 -> 2.2), as R and Excel do for these values; Python's own rounding of a
    binary float can go down."""
    f = 10 ** nd
    return math.floor(x * f + 0.5 + 1e-9) / f


def fmt(x, nd):
    return "NA" if x is None else f"{x:.{nd}f}"


def write_csv(name, columns, rows, ban):
    path = os.path.join(HERE, name)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        fh.write(ban + "\n")
        w = csv.DictWriter(fh, fieldnames=columns, lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f"wrote {path} ({len(rows)} rows)")


B_SES = ("Columns Id, Started, Ended, Eye tracking, Est. face %, Block, Status are the RDC respondent list columns (Accessing RDC Study Data); "
         "Eye tracking = accuracy in degrees (session level, MODELED as the mean of the two calibrations, one decimal), quality percent. "
         "MODELED: the CSV layout, the Id format, Status Abandoned for every early stop, and the three columns marked (modeled). "
         "Last funnel step uses the step names of the Respondent funnel (Data collection - iMotions Online). Block = study flow: "
         "Study flow 1 = foodwaste_narrative first, Study flow 2 = foodwaste_visual first.")
B_QUAL = ("MODELED layout, verify on the lab computer: the Quality metrics table of a study session page (SLIDE, EYE TRACKING, FACE %) "
          "with EYE TRACKING split into accuracy (degrees of visual angle, calibration rows only) and quality (percent of valid samples). "
          "One row per completed respondent and slide, in flow order.")
B_FEA = ("FEA thresholding at 50 (Valence +50 and -50), as in the Food Waste demo article; post processed in iMotions Lab (Post Processing > Affectiva) "
         "after download. The column headers are the MODELED FEA Metrics Export header of Workshop 3 and must be verified on the lab machine. "
         "Data quality (%) = face quality of that video.")
B_AOI = ("Individual AOI metrics on the infographic (webcam eye tracking, 1920 x 1080). Metric names follow the iMotions AOI Metrics article; "
         "the identifier columns and their order are MODELED as in Workshop 2. NA = the AOI was never fixated. Valid data = eye tracking "
         "quality on the infographic.")
B_SV = ("MODELED survey export (iMotions survey, Questionnaire slide): intention to reduce food waste (1 to 7), the more persuasive video, "
        "and the number of the four infographic tips recalled correctly (0 to 4).")

SES_COLS = ["Id", "Started", "Ended", "Eye tracking", "Est. face %", "Block", "Status", "Last funnel step (modeled)",
            "Device type (modeled)", "Browser (modeled)"]
QUAL_COLS = ["Respondent Name", "Block", "Slide position", "Slide", "Eye tracking accuracy (degrees)", "Eye tracking quality (%)",
             "Face quality (%)"]
FEA_COLS = ["Study Name", "Respondent Name", "Interval type", "Interval label", "Interval duration (s)", "Expression", "Threshold",
            "Percent of time present", "Number of segments", "Mean raw score (0-100)", "Data quality (%)"]
AOI_COLS = ["Study Name", "Respondent Name", "Stimulus", "AOI Name", "Stimulus duration", "AOI duration (ms)", "AOI duration (%)",
            "Size (px2)", "Size (%)", "Valid data", "Dwells with fixations", "Revisit count", "Fixation count", "TTFF AOI (ms)",
            "TTFF parent (ms)", "Dwell time (ms)", "Dwell time (%)", "Fixation duration (ms)", "First fixation duration (ms)",
            "Last fixation duration (ms)"]
SV_COLS = ["Study Name", "Respondent Name", "Age", "Gender", "Intention_reduce_food_waste_1to7", "More_persuasive_video",
           "Tips_recalled_0to4"]


def write_all(data):
    seed = data["seed"]
    ses = []
    for s in data["S"]:
        done = s["status"] == "Processed"
        et = ""
        fc = ""
        if done:
            et = f"{half_up((s['pre'] + s['post']) / 2, 1):.1f}, {half_up(st.mean(s['etq'].values()), 0):.0f}%"
            fc = f"{half_up(st.mean(s['faceq'].values()), 0):.0f}"
        ses.append({"Id": s["id"], "Started": s["started"].strftime("%Y-%m-%d %H:%M:%S"),
                    "Ended": s["ended"].strftime("%Y-%m-%d %H:%M:%S") if s["ended"] else "",
                    "Eye tracking": et, "Est. face %": fc, "Block": s["flow"], "Status": s["status"],
                    "Last funnel step (modeled)": s["last"], "Device type (modeled)": s["device"], "Browser (modeled)": s["browser"]})
    write_csv("SIMULATED_session_data.csv", SES_COLS, ses, banner(seed, B_SES))
    qual, fea, aoi, sv = [], [], [], []
    for s in data["comp"]:
        for k, sl in enumerate(s["slides"], start=1):
            acc = s["pre"] if sl == "Pre-study calibration" else (s["post"] if sl == "Post-study calibration" else None)
            qual.append({"Respondent Name": s["id"], "Block": s["flow"], "Slide position": k, "Slide": sl,
                         "Eye tracking accuracy (degrees)": fmt(acc, 2), "Eye tracking quality (%)": f"{s['etq'][sl]:.1f}",
                         "Face quality (%)": f"{s['faceq'][sl]:.1f}"})
        for v in VIDEOS:
            for x in EXPR:
                pct, nseg = data["fea"][(s["id"], v, x)]
                fea.append({"Study Name": STUDY_NAME, "Respondent Name": s["id"], "Interval type": "Stimulus", "Interval label": v,
                            "Interval duration (s)": f"{VID_DUR[v]:.0f}", "Expression": x, "Threshold": THRESH[x],
                            "Percent of time present": f"{pct:.2f}", "Number of segments": nseg, "Mean raw score (0-100)": "NA",
                            "Data quality (%)": f"{s['faceq'][v]:.1f}"})
        m = data["aoi"][s["id"]]
        for a in AOI_NAMES:
            px, pc = area_pct(AOIS[a])
            r = m[a]
            row = {"Study Name": STUDY_NAME, "Respondent Name": s["id"], "Stimulus": "infographic", "AOI Name": a,
                   "Stimulus duration": INFO_MS, "AOI duration (ms)": INFO_MS, "AOI duration (%)": "100.00", "Size (px2)": px,
                   "Size (%)": f"{pc:.2f}", "Valid data": f"{s['etq']['infographic']:.1f}"}
            if r is None:
                for c in AOI_COLS[10:]:
                    row[c] = "NA"
            else:
                row.update({"Dwells with fixations": r["dwells"], "Revisit count": r["revisits"], "Fixation count": r["count"],
                            "TTFF AOI (ms)": r["ttff"], "TTFF parent (ms)": r["ttff"], "Dwell time (ms)": r["dwell"],
                            "Dwell time (%)": f"{r['dwell_pct']:.2f}", "Fixation duration (ms)": f"{r['fixdur']:.0f}",
                            "First fixation duration (ms)": r["first"], "Last fixation duration (ms)": r["last"]})
            aoi.append(row)
        q = data["sv"][s["id"]]
        sv.append({"Study Name": STUDY_NAME, "Respondent Name": s["id"], "Age": q["age"], "Gender": q["gender"],
                   "Intention_reduce_food_waste_1to7": q["int"], "More_persuasive_video": q["pref"], "Tips_recalled_0to4": q["rec"]})
    write_csv("SIMULATED_quality.csv", QUAL_COLS, qual, banner(seed, B_QUAL))
    write_csv("SIMULATED_fea_metrics.csv", FEA_COLS, fea, banner(seed, B_FEA))
    write_csv("SIMULATED_aoi_metrics.csv", AOI_COLS, aoi, banner(seed, B_AOI))
    write_csv("SIMULATED_survey.csv", SV_COLS, sv, banner(seed, B_SV))


# ---------------------------------------------------------------------------
# population check and seed scan
# ---------------------------------------------------------------------------
def calibrate(n_pop=4000):
    d = generate(1, n_open=n_pop, population=True)
    o = realized(d)
    comp = d["comp"]
    print(f"population check: {n_pop} completed respondents (no funnel)")
    keys = ["n_kept", "ex_pre", "ex_post", "ex_etq", "ex_face", "n_strict_extra", "drift", "drift_sd", "eng_n", "eng_v",
            "eng_n_sd", "eng_v_sd", "dz_h1", "zero_n", "dz_h1_strict", "tip1", "tip2", "tip3", "tip4", "W", "W_strict",
            "dz_12", "dz_13", "dz_14", "dz_23", "dz_24", "dz_34", "hit_head", "hit_slope", "int_mean", "int_sd", "int_diff",
            "pref_vis", "r_pref", "rho_recall"]
    for k in keys:
        v = o[k]
        if k in ("n_kept", "ex_pre", "ex_post", "ex_etq", "ex_face", "n_strict_extra"):
            print(f"  {k:<14} {v:8d}  ({100 * v / len(comp):.1f} percent of completers)")
        else:
            print(f"  {k:<14} {v:9.4f}")
    F = d["fea"]
    kept = [s for s in comp if not screen(s)]
    for x in EXPR:
        mn = np.mean([F[(s["id"], NARR, x)][0] for s in kept])
        mv = np.mean([F[(s["id"], VIS, x)][0] for s in kept])
        print(f"  {x:<18} narrative {mn:6.2f}  visual {mv:6.2f}")
    A = d["aoi"]
    for a in AOI_NAMES:
        hits = [A[s["id"]][a] for s in kept]
        tt = [h["ttff"] for h in hits if h]
        print(f"  {a:<9} hit ratio {np.mean([h is not None for h in hits]):.3f}  TTFF median {np.median(tt):7.0f}")
    macc = np.array([(s["pre"] + s["post"]) / 2 for s in comp])
    hit = np.array([A[s["id"]]["Headline"] is not None for s in comp])
    for lo, hi in ((0, 1.5), (1.5, 2.5), (2.5, 3.5), (3.5, 4.5), (4.5, 99)):
        m = (macc > lo) & (macc <= hi)
        print(f"  Headline hit, mean accuracy ({lo}, {hi}]: {hit[m].mean():.3f} (n = {m.sum()})")


def _scan_one(s):
    return loss(realized(generate(s))), s


def scan(n):
    from multiprocessing import Pool
    with Pool(2) as pool:
        res = pool.map(_scan_one, range(SCAN_BASE, SCAN_BASE + n), chunksize=8)
    res.sort()
    print("rank  seed        loss (sum of squared standardized gaps to the model)")
    for k, (l, s) in enumerate(res[:10], start=1):
        print(f"{k:>4}  {s}  {l:9.2f}")
    ok = [l for l, _ in res if l < 1e8]
    print(f"usable seeds: {len(ok)} of {n}; median loss of the usable seeds: {st.median(ok):.2f}")


def main():
    args = sys.argv[1:]
    if args and args[0] == "--scan":
        scan(int(args[1]) if len(args) > 1 else 200)
        return
    if args and args[0] == "--calibrate":
        calibrate(int(args[1]) if len(args) > 1 else 4000)
        return
    seed = int(args[0]) if args else SEED
    data = generate(seed)
    write_all(data)
    o = realized(data)
    print(f"\nRealized statistics (seed {seed}); loss {loss(o):.2f}")
    for k, v in o.items():
        tgt = TARGETS.get(k)
        print(f"  {k:<14} {v:10.4f}" + (f"   target {tgt[0]}" if tgt else ""))


if __name__ == "__main__":
    main()
