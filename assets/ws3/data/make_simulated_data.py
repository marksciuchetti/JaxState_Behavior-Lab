#!/usr/bin/env python3
"""
make_simulated_data.py
======================
SIMULATED DATA GENERATOR for Workshop 3: Facial Expression Analysis of a Video Commercial with
Affectiva AFFDEX and the Lab Webcam in iMotions Lab 11.1
Martin's Family Clothing Consumer Behavior Research Laboratory, Jacksonville State University

EVERYTHING THIS SCRIPT WRITES IS SIMULATED. No viewer was recorded. The numbers are drawn from the
probability model documented below so that students can practice the analysis steps on an export
that looks like the "FEA Metrics Export" of the iMotions FEA thresholding R Notebook.

Design (modeled on the structure of the lab's own library study "Nike Campaign Ad Evaluation")
---------------------------------------------------------------------------------------------
That study shows a 60 s Baseline, a 60 s commercial, a 15 s Recovery, and a short questionnaire, with
the respondent camera recording the face for AFFDEX. This simulation copies only that structure. No
data from that study were used, and nothing is said about the content of the commercial, which is not
known here. 32 respondents, V01 to V32, watched the sequence ONCE (a single viewing, a single fixed
order) on the lab monitor while the lab webcam recorded the face. The recordings were post processed
with Affectiva AFFDEX (every frame, n = 1) and then run through the FEA thresholding R Notebook at its
default thresholds (Emotion 50, Valence +50 and -50, Action Units 50). Sessions ran 2026-10-26 to
2026-10-29, eight respondents per day, two operators.

The commercial is divided by a binning stimulus annotation into six 10 s bins, Bin_01 to Bin_06 (the
iMotions Sony Bravia demo study describes this kind of binning annotation). The FEA Metrics Export
reports its metrics per respondent and per time interval, where the interval is a Stimulus, an
Annotation, or a Scene (Scenes are not used here):

    Interval type   Interval label                       Length
    Stimulus        Baseline                             60 s
    Stimulus        Commercial                           60 s
    Stimulus        Recovery                             15 s
    Annotation      Bin_01 ... Bin_06 (inside Commercial) 10 s each

Expressions (names follow the manuals)
--------------------------------------
Joy (an Emotion channel), Surprise (an Emotion channel), Brow Furrow (a Facial Expression or action
unit channel), Positive Valence and Negative Valence (the two thresholds, +50 and -50, applied to the
single Valence channel, which runs from -100 to +100). For each respondent and interval the thresholding
notebook reports the percent of time the expression was present ("Percent of time present") and the number
of segments (how often it was expressed; "Number of segments"). Engagement is NOT thresholded here: the
manuals list default thresholds for Emotion, Valence, Action Units, and the head behavior channels, but
they never name Engagement, so Engagement is reported as a raw mean score (0 to 100) per interval, the kind
of statistic the AFFDEX Raw Data Statistics Export gives. The generator writes that mean in the column
"Mean raw score (0-100)" on the Engagement rows; the two thresholded columns are NA on those rows.

HEADER LAYOUT: MODELED ON THE MANUAL'S DESCRIPTION, VERIFY ON THE LAB MACHINE
------------------------------------------------------------------------------
The manuals say what the FEA Metrics Export contains (how often and how long an expression was present
for each respondent and each time interval: Stimulus, Annotation, Scene) but they print no header row, no
file name, and no file format (one article calls R exports .txt, others .csv). The column names used here
are a clear choice, not a copy of a real export: Study Name, Respondent Name, Interval type, Interval label,
Interval duration (s), Expression, Threshold, Percent of time present, Number of segments, Mean raw score
(0-100), Data quality (%). The last column is respondent level and repeated on every row, which a real
export would not do; it is here so each file stands alone (Workshop 2 did the same with Valid data). Every
CSV carries a banner line saying so. Open a real export on the lab machine and adjust the column mapping in
analysis.R (the one place where headers are renamed) before reusing the scripts on real data.

Data quality
------------
The quality measure is named "data quality" in the manuals (Exposure Summary and Data Quality Statistics
exports) and "Quality Metrics" in the replay after post processing; it is the percentage of frames
successfully analyzed, "not an indicator of overall accuracy or face quality". Here it is the percent of
frames with a face found, taken after post processing (live AFFDEX renders at 15 Hz and shows lower values
until post processed). The metrics files call it "Data quality (%)"; the session log column, which holds the
operator's transcription of the same number, is called "Face detection percent". The manuals state NO
minimum for excluding a respondent (the only numbers are the 40 to 50 percent live range and the 40 percent
review trigger). The exclusion rule used here, Data quality below 80 percent, is therefore THE LAB'S OWN
RULE, fixed before collection, and not a manufacturer value.

Raw evidence and how thresholds work (why the sensitivity analysis behaves sensibly)
------------------------------------------------------------------------------------
The generator does not draw the percent of time directly. It simulates the RAW AFFDEX score for every video
frame (30 Hz after post processing, 4050 frames in all, 0 to 100 for Joy, Surprise, Brow Furrow, and
Engagement; -100 to +100 for Valence), applies the thresholding rule to those frames, and counts. Changing
the threshold therefore changes the metrics the way it would in the notebook:

  * Each expression is a set of EPISODES (a smile, a brow furrow). An episode has a peak score a, a width
    sigma (seconds), and a smooth shape a x exp(-|t / sigma|^2.2). The expression is "present" while the
    score is above the threshold, so a stronger episode stays above a given threshold longer: the time above
    a threshold th is 2 sigma ln(a / th)^(1/2.2) when a > th, and zero when a is at or below th.
  * Peaks are spread from weak to strong (Beta distributed, on average about 60 to 70 percent of the scale).
    A threshold of 25 therefore counts nearly every episode and stretches each one; a threshold of 75
    discards the weaker episodes and shortens the rest. Time percent falls as the threshold rises, and the
    Number of segments falls with it.
  * Small AR(1) noise (autocorrelated, a few score points) rides on every channel, so a score hovering near
    the threshold flickers across it and splits one episode into two segments, as real data do.
  * Frames without a detected face are missing (NA). The notebook's gap rule is applied: a gap with the
    expression absent on both sides stays absent, present on both sides stays present, and a mixed gap
    becomes absent. Percent of time present divides by the whole interval length, so missing frames lower it.
    (Assumption to verify: the manuals do not say what the denominator is.)
  * Number of segments counts runs of present frames inside the interval, so a run that crosses a bin boundary
    is counted in both bins. (Assumption to verify.)
  * The generator writes the file for the default thresholds (SIMULATED_fea_metrics.csv) and, from the same raw
    frames, two more files that a student would get by running the notebook again at a batch threshold of 25
    and of 75: SIMULATED_fea_metrics_threshold25.csv and SIMULATED_fea_metrics_threshold75.csv. The
    sensitivity analysis in analysis.R reads them. Valence uses +th and -th.

Model assumptions
-----------------
A1  Data quality. The 30 respondents who are kept draw Normal(94.8, 2.6) clipped to 86 to 99.5, shifted so
    that the mean of all 32 respondents is 93.0. V07 (glasses glare on the lenses, intermittent face loss)
    is set to 61.0 and V19 (looked down at a phone, one long loss during the commercial) to 72.0. Both fall
    below the lab's 80 percent rule and are excluded; nobody else is close to it. The missing frames are
    placed as dropout runs (typical run 0.9 s; V07 3.5 s; V19 9 s, mostly inside the commercial) until the
    missing count matches the respondent's value, so the reported Data quality (%) is the realized share.
A2  Joy, the primary outcome. Planned Percent of time present: Baseline mean about 2 (most viewers 0, zero
    inflated), Commercial mean about 11 (right skewed), Recovery about 4. Within the commercial the planned
    bin means are 5, 6, 8, 8, 24, and 15 percent for Bin_01 to Bin_06 (Joy concentrated in Bin_05 and Bin_06,
    3 to 8 percent elsewhere). The Commercial value is the mean of the six bins because the bins tile the
    commercial. Each respondent has a positive affect trait u ~ N(0, 1) that scales the Joy episode rate by
    exp(0.9 u - 0.405), which has mean 1, gives the right skew and zero inflation, and makes the viewers who
    smile in the commercial the viewers who smile at baseline (a modest within person correlation).
A3  Surprise peaks in Bin_03: planned Percent of time present Baseline 0.8, bins 1.2, 1.8, 7.0, 2.2, 2.0, 1.5,
    Recovery 0.8. Short episodes (sigma about 0.7 s).
A4  Brow Furrow is small and flat: planned 3.0 at Baseline, 3.2 to 3.8 in every bin, 3.0 at Recovery. Longer
    episodes (sigma about 1.8 s). A weak person trait.
A5  Valence. One Valence series per frame. Positive Valence follows the Joy episodes (valence = 0.92 x the Joy
    episode score) together with extra pleasant episodes at 40 percent of the Joy rate, so Positive Valence is
    a little more frequent than Joy and strongly related to it. Negative Valence is rare (planned 0.3 at
    Baseline, 0.4 to 0.6 in the bins, 0.3 at Recovery), mostly weak episodes, mostly zero.
A6  Engagement is the raw 0 to 100 score: a resting level plus a share of the Joy, Surprise, Brow Furrow, and
    pleasant episodes plus noise. Its mean per interval is reported (A raw mean, not a threshold metric).
A7  Survey (one row per respondent, all 32). Ad_liking_1to7 = round(5.25 + 1.15 z) clipped to 1 to 7, where z
    mixes the positive affect trait u with independent noise (weight LIKE_RHO on u), so that Positive Valence
    Percent of time present during the commercial correlates with Ad_liking_1to7 at about r = .45. Brand_attitude_1to7
    has mean about 5 and correlates about .6 with liking. Would_share_Yes_No is logistic in liking (about 45
    percent Yes).
A8  Session log. Glasses worn by about 15 percent (V07 does); Lighting note records the room (the manuals
    recommend even, diffused light with a lamp pointed at the wall behind the respondent and no direct light on
    the face); Data complete is Yes for everyone (all recordings exist and were post processed); the Decision
    note applies the 80 percent rule.

How the planned means are reached
---------------------------------
The episode rate for each channel and interval is (planned percent / 100) divided by the expected time above 50
of one episode, times a calibration factor (CAL below) and the respondent's trait multiplier. The CAL factors
were found with  python3 make_simulated_data.py --calibrate  (3000 synthetic respondents, the same pipeline,
data quality as in A1), which rescales each factor until the mean Percent of time present at the default
threshold matches the plan. They are constants, so a normal run does not recalibrate.

Seed
----
SEED (20265798) was chosen from a scan of 4000 consecutive candidate seeds (20262701 to 20266700; python3 make_simulated_data.py --scan 4000) as the one whose
realized sample of kept respondents sits closest to the planned means, bin pattern, zero fraction at baseline,
correlation, and survey statistics (see TARGETS below). The scan ranks seeds by closeness to the MODEL, never
by a p value. Any single sample of 30 drifts from its model, which is itself a lesson: compare the realized
descriptives with the assumptions above.

Run:  python3 make_simulated_data.py                  (default seed below)
      python3 make_simulated_data.py <seed>           (any other seed)
      python3 make_simulated_data.py --scan [n]       (rank n candidate seeds; writes nothing)
      python3 make_simulated_data.py --calibrate      (print suggested CAL factors; writes nothing)
Outputs (same folder): SIMULATED_fea_metrics.csv, SIMULATED_fea_metrics_threshold25.csv,
                       SIMULATED_fea_metrics_threshold75.csv, SIMULATED_fea_survey.csv,
                       SIMULATED_fea_session_log.csv
Requires: Python 3, numpy, scipy. Random numbers come from numpy RandomState (a frozen stream).
"""

import csv
import math
import os
import statistics as st
import sys
from datetime import date

import numpy as np
from scipy.signal import lfilter

HERE = os.path.dirname(os.path.abspath(__file__))

SEED = 20265798     # best of 4000 scanned seeds (20262701 to 20266700) by closeness to the model; loss 19.51 against a median of 234.32
SCAN_BASE = 20262701

STUDY_NAME = "SIMULATED_FEA_Commercial_Study"
FPS = 30
STIMULI = [("Baseline", 60), ("Commercial", 60), ("Recovery", 15)]
N_FR = {name: dur * FPS for name, dur in STIMULI}
TOTAL_FR = sum(N_FR.values())               # 4050 frames
BIN_FR = 10 * FPS
BIN_LABELS = [f"Bin_{j:02d}" for j in range(1, 7)]
THRESHOLD_SETS = (50, 25, 75)
EXCLUSION_RULE_PCT = 80.0                   # the lab's rule; the manuals state none
EXCLUDED = {"V07": 61.0, "V19": 72.0}
N_RESP = 32
DQ_MEAN_ALL = 93.0
EXPRESSIONS = ["Joy", "Surprise", "Brow Furrow", "Positive Valence", "Negative Valence", "Engagement"]
THRESHOLDED = EXPRESSIONS[:5]

# planned Percent of time present at the default thresholds (per interval, mean over typical respondents)
PLAN = {
    "Joy":              {"Baseline": 2.0, "bins": [5.0, 6.0, 8.0, 8.0, 24.0, 15.0], "Recovery": 4.0},
    "Surprise":         {"Baseline": 0.8, "bins": [1.2, 1.8, 7.0, 2.2, 2.0, 1.5],   "Recovery": 0.8},
    "Brow Furrow":      {"Baseline": 3.0, "bins": [3.2, 3.4, 3.8, 3.5, 3.6, 3.3],   "Recovery": 3.0},
    "Negative Valence": {"Baseline": 0.3, "bins": [0.4, 0.5, 0.6, 0.5, 0.4, 0.4],   "Recovery": 0.3},
}
POS_EXTRA_SHARE = 0.4        # extra pleasant episodes as a share of the Joy rate (A5)
JOY_TRAIT_SD = 0.9           # A2
LIKE_RHO = 0.62              # weight of the positive affect trait in liking (A7)

# episode shapes: sigma median (s), sd of log sigma, Beta(a, b) for the peak (x 100), shape exponent
EP = {
    "Joy":       dict(sig=1.3, sig_sd=0.45, amp=(1.4, 0.9), p=2.2),
    "Surprise":  dict(sig=0.7, sig_sd=0.40, amp=(1.3, 1.0), p=2.2),
    "Brow":      dict(sig=1.8, sig_sd=0.50, amp=(1.3, 1.2), p=2.2),
    "PosExtra":  dict(sig=1.5, sig_sd=0.45, amp=(1.3, 1.1), p=2.2),
    "NegVal":    dict(sig=1.2, sig_sd=0.45, amp=(0.9, 1.4), p=2.2),
}
NOISE_SD = dict(Joy=3.0, Surprise=2.5, Brow=3.5, Valence=3.5, Engagement=2.5)
NOISE_RHO = 0.9

# calibration factors (python3 make_simulated_data.py --calibrate). Keys: channel -> [Baseline, bins x 6, Recovery]
CAL = {
    "Joy":              [1.071, 1.099, 1.061, 1.108, 1.065, 1.375, 1.163, 1.140],
    "Surprise":         [1.012, 1.016, 1.074, 1.131, 1.021, 0.967, 1.095, 1.084],
    "Brow Furrow":      [0.999, 1.017, 0.922, 1.069, 1.017, 1.053, 1.062, 1.091],
    "Negative Valence": [1.089, 1.612, 1.091, 1.102, 1.532, 2.283, 1.834, 1.291],
}


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def plogis(x):
    return 1.0 / (1.0 + math.exp(-x))


def fmt(x, nd=0):
    return "NA" if x is None or (isinstance(x, float) and math.isnan(x)) else f"{x:.{nd}f}"


def ep_time_above(channel, th=50.0, n=200000):
    """Expected time (s) above th of ONE episode of this channel (Monte Carlo with a private stream)."""
    rs = np.random.RandomState(7)
    e = EP[channel]
    a = 100.0 * rs.beta(e["amp"][0], e["amp"][1], n)
    sig = e["sig"] * np.exp(e["sig_sd"] * rs.standard_normal(n))
    d = np.where(a > th, 2.0 * sig * np.log(np.maximum(a, th + 1e-9) / th) ** (1.0 / e["p"]), 0.0)
    return float(d.mean())


E_ABOVE = {k: ep_time_above(k) for k in EP}


def episodes(rs, rate_fr, e):
    """Smooth episode curve (0 to 100) for one stimulus. rate_fr = episodes per second at each frame."""
    n = len(rate_fr)
    out = np.zeros(n)
    lam = float(rate_fr.sum()) / FPS
    k = rs.poisson(lam)
    if k == 0:
        return out
    cdf = np.cumsum(rate_fr)
    cdf = cdf / cdf[-1]
    onsets = np.minimum(np.searchsorted(cdf, rs.random_sample(k)), n - 1)
    sig = e["sig"] * np.exp(e["sig_sd"] * rs.standard_normal(k))
    amp = 100.0 * rs.beta(e["amp"][0], e["amp"][1], k)
    for o, s, a in zip(onsets, sig, amp):
        half = int(5 * s * FPS) + 1
        lo, hi = max(0, o - half), min(n, o + half + 1)
        tt = (np.arange(lo, hi) - o) / (s * FPS)
        out[lo:hi] = np.maximum(out[lo:hi], a * np.exp(-np.abs(tt) ** e["p"]))
    return out


def ar_noise(rs, n_ch, n, sds):
    z = rs.standard_normal((n_ch, n))
    out = np.empty_like(z)
    for i in range(n_ch):
        out[i] = lfilter([math.sqrt(1 - NOISE_RHO ** 2) * sds[i]], [1.0, -NOISE_RHO], z[i])
    return out


def binarize(x, th, direction):
    """Binary present (1) or absent (0) per frame. NaN frames follow the notebook's gap rule."""
    n = len(x)
    valid = ~np.isnan(x)
    b = np.zeros(n, dtype=np.int8)
    if direction > 0:
        b[valid] = (x[valid] > th)
    else:
        b[valid] = (x[valid] < -th)
    if valid.all():
        return b
    ar = np.arange(n)
    last = np.maximum.accumulate(np.where(valid, ar, -1))
    nxt = np.minimum.accumulate(np.where(valid, ar, n)[::-1])[::-1]
    lb = np.where(last >= 0, b[np.clip(last, 0, n - 1)], 0)
    rb = np.where(nxt < n, b[np.clip(nxt, 0, n - 1)], 0)
    fill = (lb == 1) & (rb == 1) & (last >= 0) & (nxt < n)
    gap = ~valid
    b[gap] = fill[gap]
    return b


def interval_metrics(b, a, z):
    seg = b[a:z]
    pct = 100.0 * float(seg.sum()) / (z - a)
    nseg = int(seg[0]) + int(np.sum((seg[1:] == 1) & (seg[:-1] == 0)))
    return pct, nseg


def make_valid_mask(rs, dq, weights, mean_run_s):
    """Boolean mask over all 4050 frames; missing runs are placed until the missing count is exact."""
    n_miss = int(round((1.0 - dq / 100.0) * TOTAL_FR))
    valid = np.ones(TOTAL_FR, dtype=bool)
    starts, offs = [], 0
    for name, _ in STIMULI:
        starts.append(offs)
        offs += N_FR[name]
    sizes = [N_FR[name] for name, _ in STIMULI]
    w = np.array(weights, dtype=float)
    w = w / w.sum()
    guard = 0
    while (~valid).sum() < n_miss and guard < 20000:
        guard += 1
        s = int(rs.choice(3, p=w))
        L = int(max(3, rs.lognormal(math.log(mean_run_s * FPS) - 0.18, 0.6)))
        L = min(L, sizes[s] - 1, n_miss - int((~valid).sum()))
        if L < 1:
            break
        st0 = starts[s] + int(rs.randint(0, sizes[s] - L))
        valid[st0:st0 + L] = False
    return valid


# ---------------------------------------------------------------------------
# one respondent: raw frames -> thresholded metrics
# ---------------------------------------------------------------------------
def rates_for(channel, stim, mult, cal, plan_key=None):
    """Episode rate per frame for one stimulus. mult = the respondent's trait multiplier."""
    plan = PLAN[plan_key or channel]
    c = cal[plan_key or channel]
    per_ep = E_ABOVE[{"Joy": "Joy", "Surprise": "Surprise", "Brow Furrow": "Brow",
                      "Negative Valence": "NegVal", "PosExtra": "PosExtra"}[channel]]
    if stim == "Baseline":
        r = [plan["Baseline"] * c[0]]
        fr = np.full(N_FR["Baseline"], r[0])
    elif stim == "Recovery":
        r = [plan["Recovery"] * c[7]]
        fr = np.full(N_FR["Recovery"], r[0])
    else:
        fr = np.repeat([plan["bins"][j] * c[1 + j] for j in range(6)], BIN_FR)
    return fr / 100.0 / per_ep * mult


def simulate_respondent(rs, traits, valid_mask, thresholds, cal):
    """Return {threshold: {(itype, label, expr): (pct, nseg, mean_raw)}} for one respondent."""
    u, s_t, w_t, q_t, e0 = traits["u"], traits["s"], traits["w"], traits["q"], traits["e0"]
    m_joy = math.exp(JOY_TRAIT_SD * u - JOY_TRAIT_SD ** 2 / 2)
    m_sur = math.exp(0.5 * s_t - 0.125)
    m_brw = math.exp(0.7 * w_t - 0.245)
    m_neg = math.exp(0.8 * q_t - 0.32)
    out = {th: {} for th in thresholds}
    offs = 0
    for name, _dur in STIMULI:
        n = N_FR[name]
        vmask = valid_mask[offs:offs + n]
        offs += n
        rj = rates_for("Joy", name, m_joy, cal)
        rx = POS_EXTRA_SHARE * rates_for("Joy", name, m_joy, cal) * E_ABOVE["Joy"] / E_ABOVE["PosExtra"]
        rs_ = rates_for("Surprise", name, m_sur, cal)
        rb = rates_for("Brow Furrow", name, m_brw, cal)
        rn = rates_for("Negative Valence", name, m_neg, cal)
        J = episodes(rs, rj, EP["Joy"])
        X = episodes(rs, rx, EP["PosExtra"])
        S = episodes(rs, rs_, EP["Surprise"])
        B = episodes(rs, rb, EP["Brow"])
        N = episodes(rs, rn, EP["NegVal"])
        nz = ar_noise(rs, 5, n, [NOISE_SD["Joy"], NOISE_SD["Surprise"], NOISE_SD["Brow"], NOISE_SD["Valence"],
                                 NOISE_SD["Engagement"]])
        joy = np.clip(J + 1.0 + nz[0], 0, 100)
        sur = np.clip(S + 0.5 + nz[1], 0, 100)
        brw = np.clip(B + 1.5 + nz[2], 0, 100)
        val = np.clip(np.maximum(0.92 * J, X) + nz[3] - N, -100, 100)
        eng = np.clip(e0 + 0.5 * J + 0.6 * S + 0.4 * B + 0.3 * X + nz[4], 0, 100)
        series = {"Joy": (joy, 1), "Surprise": (sur, 1), "Brow Furrow": (brw, 1),
                  "Positive Valence": (val, 1), "Negative Valence": (val, -1)}
        for key in series:
            series[key] = (np.where(vmask, series[key][0], np.nan), series[key][1])
        eng_m = np.where(vmask, eng, np.nan)
        for th in thresholds:
            for expr in THRESHOLDED:
                x, direction = series[expr]
                b = binarize(x, float(th), direction)
                out[th][("Stimulus", name, expr)] = interval_metrics(b, 0, n) + (None,)
                if name == "Commercial":
                    for j in range(6):
                        out[th][("Annotation", BIN_LABELS[j], expr)] = interval_metrics(b, j * BIN_FR, (j + 1) * BIN_FR) + (None,)
        for th in thresholds:
            out[th][("Stimulus", name, "Engagement")] = (None, None, _nanmean(eng_m))
            if name == "Commercial":
                for j in range(6):
                    out[th][("Annotation", BIN_LABELS[j], "Engagement")] = (None, None, _nanmean(eng_m[j * BIN_FR:(j + 1) * BIN_FR]))
    return out


def _nanmean(x):
    v = x[~np.isnan(x)]
    return float(v.mean()) if len(v) else None


# ---------------------------------------------------------------------------
# the full study
# ---------------------------------------------------------------------------
def generate(seed, thresholds=(50,), cal=None):
    cal = cal or CAL
    rs = np.random.RandomState(seed)
    ids = [f"V{i + 1:02d}" for i in range(N_RESP)]
    sexes = np.array(["F"] * 18 + ["M"] * 14)
    sexes = list(sexes[rs.permutation(N_RESP)])
    ages = [int(round(min(max(rs.normal(26.5, 6.5), 19), 45))) for _ in range(N_RESP)]
    ages[ages.index(min(ages))] = 19
    ages[ages.index(max(ages))] = 45
    u = rs.standard_normal(N_RESP)
    s_t = rs.standard_normal(N_RESP)
    w_t = rs.standard_normal(N_RESP)
    q_t = rs.standard_normal(N_RESP)
    e0 = np.maximum(1.0, rs.normal(4.0, 1.5, N_RESP))

    # A1 data quality targets
    dq_t = np.clip(rs.normal(94.8, 2.6, N_RESP), 86.0, 99.5)
    for rid, v in EXCLUDED.items():
        dq_t[ids.index(rid)] = v
    kept_idx = [i for i, rid in enumerate(ids) if rid not in EXCLUDED]
    shift = (DQ_MEAN_ALL * N_RESP - sum(EXCLUDED.values()) - dq_t[kept_idx].sum()) / len(kept_idx)
    for i in kept_idx:
        dq_t[i] = min(99.5, max(85.5, dq_t[i] + shift))

    resp = []
    metrics = {th: {} for th in thresholds}
    for i, rid in enumerate(ids):
        if rid == "V07":
            wts, run = (1.0, 1.0, 0.3), 3.5
        elif rid == "V19":
            wts, run = (0.08, 0.84, 0.08), 9.0
        else:
            wts, run = (0.444, 0.444, 0.112), 0.9
        valid = make_valid_mask(rs, float(dq_t[i]), wts, run)
        dq = round(100.0 * float(valid.sum()) / TOTAL_FR, 1)
        res = simulate_respondent(rs, dict(u=float(u[i]), s=float(s_t[i]), w=float(w_t[i]), q=float(q_t[i]),
                                           e0=float(e0[i])), valid, thresholds, cal)
        for th in thresholds:
            metrics[th][rid] = res[th]
        resp.append(dict(id=rid, sex=sexes[i], age=ages[i], u=float(u[i]), dq=dq,
                         date=date(2026, 10, 26 + i // 8),
                         operator="Operator_1" if (i // 2) % 2 == 0 else "Operator_2"))

    # A7 survey
    survey = []
    for r in resp:
        eps = rs.standard_normal()
        z = LIKE_RHO * r["u"] + math.sqrt(1 - LIKE_RHO ** 2) * eps
        like = int(min(7, max(1, round(5.25 + 1.15 * z))))
        att = int(min(7, max(1, round(5.0 + 1.2 * (0.6 * z + 0.8 * rs.standard_normal())))))
        share = "Yes" if rs.random_sample() < plogis(-0.25 + 1.0 * z) else "No"
        r["like"] = like
        survey.append({"Respondent Name": r["id"], "Gender": r["sex"], "Age": r["age"],
                       "Ad_liking_1to7": like, "Brand_attitude_1to7": att, "Would_share_Yes_No": share})

    # A8 session log
    lights = ["Diffused light, lamp against wall behind respondent",
              "Diffused light, lamp against wall; blinds lowered",
              "Overhead lights dimmed, lamp against wall",
              "Diffused light, lamp against wall; monitor tilted slightly down"]
    log = []
    kept_dq = sorted(r["dq"] for r in resp if r["id"] not in EXCLUDED)
    lowest_kept = kept_dq[0]
    for r in resp:
        rid = r["id"]
        glasses = "Yes" if (rid == "V07" or rs.random_sample() < 0.15) else "No"
        light = lights[int(rs.randint(0, len(lights)))]
        note = ""
        if rid == "V07":
            light = "Overhead light reflecting on glasses lenses (glare)"
            note = ("EXCLUDE: Data quality 61 percent is below the lab's 80 percent rule set before collection; "
                    "glare on the glasses lenses hid the eyes and brows for long stretches; recorded in full anyway "
                    "because the session was scheduled")
        elif rid == "V19":
            note = ("EXCLUDE: Data quality 72 percent is below the lab's 80 percent rule set before collection; "
                    "respondent looked down at a phone during the commercial; recorded in full anyway because the "
                    "session was scheduled")
        elif r["dq"] == lowest_kept:
            note = (f"Keep: Data quality {r['dq']:.1f} percent is above the lab's 80 percent rule; "
                    "brief face losses when the respondent adjusted position")
        log.append({"Respondent Name": rid, "Session Date": r["date"].isoformat(), "Operator": r["operator"],
                    "Glasses": glasses, "Lighting note": light, "Face detection percent": fmt(r["dq"], 1),
                    "Data complete": "Yes", "Decision note": note})
    return dict(seed=seed, resp=resp, metrics=metrics, survey=survey, log=log)


def metric_rows(data, th):
    """Long format rows of the FEA Metrics Export for one threshold set."""
    rows = []
    order = [("Stimulus", n) for n, _ in STIMULI] + [("Annotation", b) for b in BIN_LABELS]
    dur = {n: d for n, d in STIMULI}
    for r in data["resp"]:
        m = data["metrics"][th][r["id"]]
        for itype, label in order:
            seconds = dur[label] if itype == "Stimulus" else 10
            for expr in EXPRESSIONS:
                pct, nseg, raw = m[(itype, label, expr)]
                if expr == "Engagement":
                    thr = "NA"
                elif expr == "Negative Valence":
                    thr = f"-{th}"
                else:
                    thr = str(th)
                rows.append({
                    "Study Name": STUDY_NAME, "Respondent Name": r["id"], "Interval type": itype,
                    "Interval label": label, "Interval duration (s)": seconds, "Expression": expr, "Threshold": thr,
                    "Percent of time present": fmt(pct, 2) if expr != "Engagement" else "NA",
                    "Number of segments": str(nseg) if expr != "Engagement" else "NA",
                    "Mean raw score (0-100)": fmt(raw, 2) if expr == "Engagement" else "NA",
                    "Data quality (%)": fmt(r["dq"], 1)})
    return rows


# ---------------------------------------------------------------------------
# realized statistics (kept respondents) and the scan loss
# ---------------------------------------------------------------------------
def pearson(x, y):
    mx, my = st.mean(x), st.mean(y)
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    den = math.sqrt(sum((a - mx) ** 2 for a in x) * sum((b - my) ** 2 for b in y))
    return sxy / den if den > 0 else 0.0


def rank_avg(v):
    order = sorted(range(len(v)), key=lambda i: v[i])
    r = [0.0] * len(v)
    i = 0
    while i < len(v):
        j = i
        while j + 1 < len(v) and v[order[j + 1]] == v[order[i]]:
            j += 1
        for k in range(i, j + 1):
            r[order[k]] = (i + j) / 2 + 1
        i = j + 1
    return r


def realized(data):
    kept = [r for r in data["resp"] if r["dq"] >= EXCLUSION_RULE_PCT]
    m = data["metrics"][50]

    def vals(expr, itype, label):
        return [m[r["id"]][(itype, label, expr)][0] for r in kept]

    out = {}
    joy_b, joy_c, joy_r = (vals("Joy", "Stimulus", s) for s in ("Baseline", "Commercial", "Recovery"))
    out["joy_base"], out["joy_comm"], out["joy_rec"] = st.mean(joy_b), st.mean(joy_c), st.mean(joy_r)
    out["joy_base_zero"] = sum(v == 0 for v in joy_b) / len(joy_b)
    out["joy_comm_any"] = sum(v > 0 for v in joy_c) / len(joy_c)
    for j, lab in enumerate(BIN_LABELS):
        out[f"joy_bin{j + 1}"] = st.mean(vals("Joy", "Annotation", lab))
        out[f"sur_bin{j + 1}"] = st.mean(vals("Surprise", "Annotation", lab))
        out[f"brw_bin{j + 1}"] = st.mean(vals("Brow Furrow", "Annotation", lab))
    out["brow_comm"] = st.mean(vals("Brow Furrow", "Stimulus", "Commercial"))
    out["pv_comm"] = st.mean(vals("Positive Valence", "Stimulus", "Commercial"))
    out["neg_comm"] = st.mean(vals("Negative Valence", "Stimulus", "Commercial"))
    pv = vals("Positive Valence", "Stimulus", "Commercial")
    like = [float(r["like"]) for r in kept]
    out["r_pv_like"] = pearson(pv, like)
    out["rho_pv_like"] = pearson(rank_avg(pv), rank_avg(like))
    out["like_mean"], out["like_sd"] = st.mean(like), st.stdev(like)
    sv = {s["Respondent Name"]: s for s in data["survey"]}
    out["att_mean"] = st.mean(sv[r["id"]]["Brand_attitude_1to7"] for r in kept)
    out["share_yes"] = sum(sv[r["id"]]["Would_share_Yes_No"] == "Yes" for r in kept) / len(kept)
    out["bin_sur_peak"] = out["sur_bin3"] - max(out[f"sur_bin{j}"] for j in (1, 2, 4, 5, 6))
    out["brow_range"] = max(out[f"brw_bin{j}"] for j in range(1, 7)) - min(out[f"brw_bin{j}"] for j in range(1, 7))
    out["joy_bin5_peak"] = out["joy_bin5"] - max(out[f"joy_bin{j}"] for j in (1, 2, 3, 4, 6))
    out["joy_low_min"] = min(out[f"joy_bin{j}"] for j in (1, 2, 3, 4))
    out["joy_low_max"] = max(out[f"joy_bin{j}"] for j in (1, 2, 3, 4))
    out["brow_min"] = min(out[f"brw_bin{j}"] for j in range(1, 7))
    out["brow_max"] = max(out[f"brw_bin{j}"] for j in range(1, 7))
    return out


TARGETS = {
    "joy_base": (2.0, 0.5), "joy_comm": (11.0, 1.0), "joy_rec": (4.0, 0.9), "joy_base_zero": (0.67, 0.07),
    "joy_comm_any": (0.90, 0.06),
    "joy_bin1": (5.0, 1.2), "joy_bin2": (6.0, 1.4), "joy_bin3": (8.0, 1.6), "joy_bin4": (8.0, 1.6),
    "joy_bin5": (24.0, 2.5), "joy_bin6": (15.0, 2.0),
    "sur_bin1": (1.2, 0.6), "sur_bin2": (1.8, 0.7), "sur_bin3": (7.0, 1.2), "sur_bin4": (2.2, 0.8),
    "sur_bin5": (2.0, 0.8), "sur_bin6": (1.5, 0.7),
    "brow_comm": (3.5, 0.5), "neg_comm": (0.5, 0.25), "pv_comm": (12.5, 1.5),
    "r_pv_like": (0.45, 0.04), "rho_pv_like": (0.42, 0.08),
    "like_mean": (5.2, 0.25), "like_sd": (1.2, 0.15), "att_mean": (5.0, 0.25), "share_yes": (0.45, 0.10),
}
# hard shape requirements (penalized when violated)
SHAPE_MIN = {"bin_sur_peak": 3.0, "joy_bin5_peak": 6.0, "joy_low_min": 3.0, "brow_min": 2.3}
SHAPE_MAX = {"brow_range": 2.0, "joy_low_max": 8.6, "brow_max": 4.7}


def loss(stats):
    total = sum(((stats[k] - tgt) / tol) ** 2 for k, (tgt, tol) in TARGETS.items())
    for k, lo in SHAPE_MIN.items():
        if stats[k] < lo:
            total += 25.0 * (lo - stats[k])
    for k, hi in SHAPE_MAX.items():
        if stats[k] > hi:
            total += 25.0 * (stats[k] - hi)
    return total


def _scan_one(seed):
    return loss(realized(generate(seed))), seed


# ---------------------------------------------------------------------------
# calibration
# ---------------------------------------------------------------------------
def calibrate(n_resp=3000, iters=4):
    cal = {k: list(v) for k, v in CAL.items()}
    for it in range(iters):
        rs = np.random.RandomState(1000)      # the same stream every iteration (common random numbers)
        acc = {(ch, k): [] for ch in cal for k in range(8)}
        for i in range(n_resp):
            tr = dict(u=float(rs.standard_normal()), s=float(rs.standard_normal()), w=float(rs.standard_normal()),
                      q=float(rs.standard_normal()), e0=4.0)
            dq = float(min(99.5, max(86.0, rs.normal(94.8, 2.6))))
            valid = make_valid_mask(rs, dq, (0.444, 0.444, 0.112), 0.9)
            res = simulate_respondent(rs, tr, valid, (50,), cal)[50]
            for ch in cal:
                keys = [("Stimulus", "Baseline")] + [("Annotation", b) for b in BIN_LABELS] + [("Stimulus", "Recovery")]
                for k, (itype, lab) in enumerate(keys):
                    acc[(ch, k)].append(res[(itype, lab, ch)][0])
        print(f"iteration {it + 1}")
        for ch in cal:
            plan = [PLAN[ch]["Baseline"]] + PLAN[ch]["bins"] + [PLAN[ch]["Recovery"]]
            obs = [st.mean(acc[(ch, k)]) for k in range(8)]
            for k in range(8):
                cal[ch][k] *= plan[k] / max(obs[k], 1e-6)
            print(f"  {ch:<17} observed " + " ".join(f"{o:5.2f}" for o in obs) + "  plan " + " ".join(f"{p:5.2f}" for p in plan))
    print("\nCAL = {")
    for ch in cal:
        print(f'    "{ch}": [' + ", ".join(f"{c:.3f}" for c in cal[ch]) + "],")
    print("}")


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
    if args and args[0] == "--calibrate":
        calibrate()
        return
    if args and args[0] == "--scan":
        n = int(args[1]) if len(args) > 1 else 120
        from multiprocessing import Pool
        with Pool(2) as pool:
            res = pool.map(_scan_one, range(SCAN_BASE, SCAN_BASE + n), chunksize=4)
        res.sort()
        print("rank  seed        loss (sum of squared standardized gaps to the model)")
        for k, (l, s) in enumerate(res[:10], start=1):
            print(f"{k:>4}  {s}  {l:7.2f}")
        print(f"median loss over {n} seeds: {st.median(l for l, _ in res):.2f}")
        return
    seed = int(args[0]) if args else SEED
    data = generate(seed, THRESHOLD_SETS)
    cols = ["Study Name", "Respondent Name", "Interval type", "Interval label", "Interval duration (s)", "Expression",
            "Threshold", "Percent of time present", "Number of segments", "Mean raw score (0-100)", "Data quality (%)"]
    tail = ("The column headers are modeled on the manual's description of the FEA Metrics Export (percent of time present "
            "and number of segments per respondent and interval) and must be verified on the lab machine.")
    for th, fname in ((50, "SIMULATED_fea_metrics.csv"), (25, "SIMULATED_fea_metrics_threshold25.csv"),
                      (75, "SIMULATED_fea_metrics_threshold75.csv")):
        what = "default thresholds (Emotion 50, Valence +50 and -50, Action Units 50)" if th == 50 else f"batch threshold {th} (Valence +{th} and -{th})"
        banner = (f"# SIMULATED DATA for teaching. Generated by make_simulated_data.py (seed {seed}). No real viewer was "
                  f"recorded. FEA thresholding at the {what}. {tail}")
        write_csv(fname, cols, metric_rows(data, th), banner)
    banner = (f"# SIMULATED DATA for teaching. Generated by make_simulated_data.py (seed {seed}). No real viewer was "
              "recorded. Survey answers modeled on a short questionnaire after the commercial.")
    write_csv("SIMULATED_fea_survey.csv", list(data["survey"][0].keys()), data["survey"], banner)
    banner = (f"# SIMULATED DATA for teaching. Generated by make_simulated_data.py (seed {seed}). No real viewer was "
              "recorded. Face detection percent is the operator's copy of the Data quality value (percent of frames analyzed) "
              "after post processing.")
    write_csv("SIMULATED_fea_session_log.csv", list(data["log"][0].keys()), data["log"], banner)
    stats = realized(data)
    print(f"\nRealized statistics for the kept respondents (target, tolerance) and loss {loss(stats):.2f}:")
    for k, (tgt, tol) in TARGETS.items():
        print(f"  {k:<14} {stats[k]:9.3f}   target {tgt}")
    for k in list(SHAPE_MIN) + list(SHAPE_MAX):
        print(f"  {k:<14} {stats[k]:9.3f}")


if __name__ == "__main__":
    main()
