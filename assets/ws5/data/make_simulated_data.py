#!/usr/bin/env python3
"""
make_simulated_data.py
======================
SIMULATED DATA GENERATOR for Workshop 5: EEG frontal alpha asymmetry (FA) with the Neuroelectrics
Enobio 8 (NIC2 and Lab Streaming Layer) in iMotions Lab 11.1, two video commercials
Martin's Family Clothing Consumer Behavior Research Laboratory, Jacksonville State University

EVERYTHING THIS SCRIPT WRITES IS SIMULATED. No person was recorded. The numbers are drawn from the
probability model documented below so that students can practice the analysis steps on files that
look like the exports of the iMotions EEG Frontal Asymmetry and EEG Power Spectral Density R
Notebooks, plus three lab files (an artifact report table, a session log, and a survey).

The design copies the logic of the iMotions demo study "DEMO - SonyBravia" (iMotions 9.0 demo
article), which compared a commercial with a key scene (the frog scene) against the same commercial
without it (between subjects), compared two commercials (within subjects), used two neutral 1 minute
nature videos as baselines, and analyzed the last 10 seconds of each commercial. Nothing about the
content of the Sony commercials was used. The commercials here are unnamed, and every time profile
inside them is invented for teaching.

Design
------
48 respondents, P01 to P48, block randomized (blocks of 4) to four groups of 12:
    A_C1first  Version A (sees Commercial_1A, WITH the key scene), Commercial_1 first
    A_C2first  Version A, Commercial_2 first
    B_C1first  Version B (sees Commercial_1B, the same commercial with the key scene cut)
    B_C2first  Version B, Commercial_2 first
Flow for every respondent: instruction slide, Baseline_1 (60 s neutral nature video), first
commercial, rating slide, Baseline_2 (60 s neutral nature video), second commercial, rating slide.
Each commercial is compared with the baseline that immediately precedes it.
Lengths: Commercial_1A 60 s (KeyScene 32.0 to 38.0 s); Commercial_1B 54 s (the 6 s cut); Commercial_2
60 s. Stimulus annotations: "KeyScene" (6 s, Commercial_1A only) and "EndScene" (the last 10 s of
each commercial; Label is "EndScene" on all three commercials, so the Parent Stimulus column tells
them apart).
Device: Enobio 8 through NIC2 and LSL, iMotions device name "Enobio EEG", 500 Hz (lab manual).
Channel map (ASSUMED from the lab manual montage order, NOT verified): Ch1 Fz, Ch2 Cz, Ch3 Pz, Ch4 Oz,
Ch5 F3, Ch6 F4, Ch7 P3, Ch8 P4.

Files written (same folder)
---------------------------
SIMULATED_eeg_fa_metrics.csv
    EEG Frontal Asymmetry Metrics export as if the FA notebook had been run with left = Ch5 (F3),
    right = Ch6 (F4), Alpha band only, Welch 1 s windows, notch 60 Hz, band-pass 0.5 - 100 Hz,
    artifact threshold 120 microvolts. Columns exactly as the R Processed EEG Export Files article
    lists them (only the Alpha band was selected): Study Name, Respondent Name, Gender, Age, Group,
    Type, Label, Start (ms), Duration (ms), Parent Stimulus, Mean Frontal Asymmetry Alpha, Algorithm
    Device Name. MODELED: Type is "Stimulus" or "Annotation" (no "Annotation (Instance)" rows);
    Parent Stimulus repeats the label on stimulus rows; Algorithm Device Name is one column holding
    "Frontal Asymmetry (Enobio EEG)"; the stimulus mean FA is the mean of the 1 s window FA values.
SIMULATED_eeg_fa_metrics_default_channels.csv
    The same export as if the notebook had been run with its Enobio DEFAULTS (left Ch1, right Ch7),
    which in the assumed montage are Fz and P3: a midline frontal and a left parietal site, not a
    left and right frontal pair. Teaching contrast only.
SIMULATED_eeg_psd_metrics.csv
    EEG Power Spectral Density Metrics export (Theta and Alpha selected, all 8 channels in the
    cluster). Columns: the same first ten, then "Mean PSD Ch1 Theta (dB)", "Mean PSD Ch1 Alpha (dB)",
    ... "Mean PSD Ch8 Alpha (dB)", "Mean PSD Electrode Cluster Average Theta (dB)", "Mean PSD
    Electrode Cluster Average Alpha (dB)", Algorithm Device Name. The name pattern follows the
    article; the column ORDER and the dB base (10 log10 of power in microvolts squared per Hz) are
    MODELED.
SIMULATED_eeg_artifact_report.csv
    MODELED table copied from the per respondent notebook report (box "Percentage of data filtered
    (nb of samples filtered / nb total of samples)"), one row per respondent, stimulus, and channel
    Ch5 (F3) and Ch6 (F4): Samples expected (duration x 500), Samples recorded, Samples removed (at
    or above 120 microvolts), Percent removed (%) = removed / recorded x 100. The real report layout,
    and whether it gives values per stimulus, must be verified on the lab computer.
SIMULATED_eeg_session_log.csv
    MODELED operator log: Respondent Name, Session Date, Operator, Version, Order, Cap size,
    Connection, NIC2 quality F3 before recording, NIC2 quality F4 before recording, Gel used, Setup
    minutes, Incidents, Data complete, Decision note.
SIMULATED_eeg_survey.csv
    MODELED questionnaire: Liking of each commercial (1 to 7), key scene recall (Version A only),
    brand named correctly, Gender, Age.

Model assumptions
-----------------
A1  Window level model. For every respondent, stimulus, channel, and band (Theta, Alpha) the
    natural log of power in each 1 s Welch window (50 percent overlap, one window every 0.5 s, so a
    60 s stimulus has 119 windows, a 10 s EndScene 19, and the 6 s KeyScene 11) is
        ln P = channel level + person level + person by channel level + phase effect
               + common fluctuation (AR(1), phi 0.85, SD 0.30, shared by all channels of a band)
               + channel noise (AR(1), phi 0.45, SD 0.16) + asymmetry term (Alpha at F3 and F4 only).
    Channel levels (dB): Alpha Fz 8.5, Cz 9.5, Pz 12.0, Oz 13.0, F3 8.5, F4 8.5, P3 11.5, P4 11.8;
    Theta Fz 9.5, Cz 9.3, Pz 8.8, Oz 8.6, F3 9.2, F4 9.2, P3 8.7, P4 8.7. Person level SD 0.46 (Alpha)
    and 0.35 (Theta) in ln units; person by channel SD 0.12 (F3 and F4 share theirs, so all of the
    stable left right difference sits in the trait below).
A2  Asymmetry term. ln P(F4, Alpha) gets +F/2 and ln P(F3, Alpha) gets -F/2, so the window FA
    ln(F4 alpha) - ln(F3 alpha) = F + noise (noise SD 0.16 x sqrt(2) = 0.23 per window).
    F = trait + commercial profile + stimulus state + individual responses, where
      trait          ~ Normal(-0.05, 0.12) per respondent (most scores negative, as in the demo chart)
      stimulus state ~ Normal(0, 0.035) per respondent and stimulus (slow drift)
      commercial response ~ 0.045 x (sqrt(.5) u_i + sqrt(.5) v_ik), shared part u_i across commercials
      key scene response  ~ Normal(0, 0.06) per Version A respondent, scaled by the key scene shape
      end scene response  ~ Normal(0, 0.04) per respondent and commercial, scaled by the end scene shape
      position: -0.012 for the commercial seen second (fatigue), 0 for the first.
A3  Commercial profiles (mean FA relative to the baseline level 0):
      Commercial_1A: level -0.04; KeyScene bump of 0.153 that rises over the first second (32 to
                     33 s), holds to 38 s, and decays with a 1.5 s time constant (mean over the 11
                     KeyScene windows about +0.10 once the position effect is averaged in);
                     EndScene +0.072 above the level (mean +0.03).
      Commercial_1B: level -0.04; EndScene +0.026 above the level (mean -0.015).
      Commercial_2:  level +0.005; EndScene +0.02 above the level.
    Planned effects: KeyScene change +0.10 (change SD about 0.14, d_z about 0.70); EndScene change of
    Commercial 1, Version A minus Version B, 0.045 (SD about 0.118, d about 0.38); whole commercial
    change, Commercial_2 minus Commercial_1, about 0.032 (SD about 0.108, d_z about 0.30).
A4  Phase effects on power: during commercials Alpha falls by 0.15 ln units (about 0.65 dB) at
    every channel (attention), frontal midline Theta (Fz) rises by 0.08 ln units (about 0.35 dB) and
    Theta elsewhere by 0.03.
A5  Artifacts. Supra-threshold events (samples at or above 120 microvolts) arrive at a respondent
    rate ~ lognormal(median 4 per minute, log SD 0.55): 85 percent blinks (0.15 s median, frontal
    channels Fz, F3, F4, each hit with probability 0.85) and 15 percent movements (0.9 s median,
    every channel hit with probability 0.95); the duration at each channel varies by up to
    15 percent. Percent removed at a channel = union of its event
    intervals / recorded time x 100. A 1 s window that overlaps a removed sample is dropped
    (FA needs F3 and F4 clean; PSD needs its own channel clean); a segment with no clean window is NA.
A6  Four "restless" respondents (rate x 2.5 and one stimulus with extra movement until the larger
    of F3 and F4 reaches 15.5 to 24 percent) are kept by the 25 percent rule and dropped by the
    15 percent sensitivity rule.
A7  Exactly three exclusions under the rule written before collection (more than 25 percent of F3
    or F4 samples removed in any baseline or commercial, or a Red NIC2 quality status at F3 or F4
    that could not be fixed):
      lost F4 contact (Version A): F4 Red before recording after re-seating, gel, and a lead swap;
        F4 contact pops remove about 40 percent of F4 samples; the surviving F4 windows carry
        broadband noise (+0.80 ln units), so this respondent's FA values are far too positive.
      heavy movement and blinking (Version B): event rate x 4; extra movement in Commercial_1B
        until F4 reaches 29.4 percent and in Commercial_2 until 21 percent.
      LSL stream gap (Version A): 8.6 s gap in Commercial_2 from 21.4 s (Samples recorded falls
        short of Samples expected) and a drift of about 15 s on all channels after the stream
        reconnected.
A8  Survey. Liking (1 to 7): Commercial_1A mean 5.3, Commercial_1B 5.0, Commercial_2 5.1 (SD 1.2),
    from the standardized latent commercial approach (commercial response plus stimulus state,
    weight 0.42), a response style shared by the two ratings (30 percent), and noise. Key scene
    recall (Version A) from the key scene response (about 75 percent Yes). Brand named correctly
    about 70 percent Yes.
A9  Session log. Cap size S 20, M 60, L 20 percent ("Most adults will be medium"); Connection USB
    (the iMotions article recommends it); NIC2 quality before recording Green (82 percent) or
    Yellow; Gel used Yes (the condensed lab manual says always gel); Setup minutes ~ Normal(16, 3)
    clipped to 10 to 25; three sessions per weekday from 2027-02-01.

Run:  python3 make_simulated_data.py                 (default seed below)
      python3 make_simulated_data.py <seed>          (any other seed)
      python3 make_simulated_data.py --scan [n]      (rank n candidate seeds; writes nothing)
      python3 make_simulated_data.py --calibrate     (population check of the model; writes nothing)
"""

import csv
import math
import os
import statistics as st
import sys
from datetime import date, timedelta

import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))

# Fixed seed so the files are reproducible. It was chosen from a scan of 3000 candidate seeds,
# 20265001 to 20268000 (python3 make_simulated_data.py --scan 3000; 2267 of them gave exactly the
# three planned exclusions with 22 and 23 kept and no NA for a kept respondent), as the one whose
# realized sample of kept respondents sits closest to the planned MODEL values in TARGETS below
# (loss 3.13; the median of the 2267 usable seeds is 67.5). The scan ranks seeds by closeness to the
# model, never by a p value. Any single sample still drifts from its model (here, for example, the
# liking correlation for Commercial 1 is lower than planned), which is itself a lesson: compare the
# realized descriptives with the assumptions above.
SEED = 20267088
SCAN_BASE = 20265001

STUDY_NAME = "SIMULATED Commercial FA Study"
DEVICE_NAME = "Enobio EEG"
FA_DEVICE_CELL = "Frontal Asymmetry (Enobio EEG)"
PSD_DEVICE_CELL = "Power Spectral Density (Enobio EEG)"
FS = 500
N_RESP = 48
CH = ["Ch1", "Ch2", "Ch3", "Ch4", "Ch5", "Ch6", "Ch7", "Ch8"]
SITE = {"Ch1": "Fz", "Ch2": "Cz", "Ch3": "Pz", "Ch4": "Oz", "Ch5": "F3", "Ch6": "F4", "Ch7": "P3", "Ch8": "P4"}
FRONTAL = {0, 4, 5}                    # channel indexes of Fz, F3, F4 (blinks reach them)
I_F3, I_F4, I_FZ, I_P3 = 4, 5, 0, 6
BANDS = ["Theta", "Alpha"]
LEVEL_DB = {"Alpha": [8.5, 9.5, 12.0, 13.0, 8.5, 8.5, 11.5, 11.8],
            "Theta": [9.5, 9.3, 8.8, 8.6, 9.2, 9.2, 8.7, 8.7]}
LN10_10 = 10.0 / math.log(10.0)        # dB = LN10_10 x ln(P), about 4.343

DUR = {"Baseline_1": 60.0, "Baseline_2": 60.0, "Commercial_1A": 60.0, "Commercial_1B": 54.0, "Commercial_2": 60.0}
KEY_START, KEY_END = 32.0, 38.0
END_LEN = 10.0
WIN, STEP = 1.0, 0.5
NWMAX = int(round((60.0 - WIN) / STEP)) + 1          # 119 windows in a 60 s stimulus
LEAD_S = (20, 40)                     # instruction slide before Baseline_1
RATING_S = (25, 60)                   # rating slide after each commercial

# A2 and A3: FA model (ln units)
TRAIT_MEAN, TRAIT_SD = -0.05, 0.12
STATE_SD = 0.035
RESP_SD, RESP_SHARED = 0.045, 0.5
KEY_RESP_SD, END_RESP_SD = 0.06, 0.04
POS_SECOND = -0.012
LEVEL = {"Commercial_1A": -0.04, "Commercial_1B": -0.04, "Commercial_2": 0.005}
KEY_AMP = 0.153
END_AMP = {"Commercial_1A": 0.0719, "Commercial_1B": 0.0257, "Commercial_2": 0.02}
# A1 and A4: power model
PERSON_SD = {"Alpha": 0.46, "Theta": 0.35}
PERSON_CH_SD = 0.12
COMMON_PHI, COMMON_SD = 0.85, 0.30
NOISE_PHI, NOISE_SD = 0.45, 0.16
ALPHA_COMM = -0.15
THETA_COMM = {0: 0.08}                # Fz; other channels THETA_COMM_OTHER
THETA_COMM_OTHER = 0.03
# A5 to A7: artifacts
RATE_MEDIAN, RATE_LOGSD = 4.0, 0.55
P_MOVE = 0.15
BLINK_MED, BLINK_LOGSD = 0.15, 0.40
MOVE_MED, MOVE_LOGSD = 0.90, 0.60
N_RESTLESS = 4
EXCLUDE_PCT, STRICT_PCT = 25.0, 15.0
GAP_START, GAP_LEN, DRIFT_LEN = 21.4, 8.6, 15.0
F4_NOISE = 0.80
# A8 survey
LIKE_MEAN = {"Commercial_1A": 5.3, "Commercial_1B": 5.0, "Commercial_2": 5.1}
LIKE_SD, LIKE_W, LIKE_STYLE = 1.2, 0.42, 0.30

DAYS0 = date(2027, 2, 1)
OPERATORS = ["Operator_1", "Operator_2"]


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def clip(x, lo, hi):
    return max(lo, min(hi, x))


def weekdays(start, n):
    out, d = [], start
    while len(out) < n:
        if d.weekday() < 5:
            out.append(d)
        d += timedelta(days=1)
    return out


def win_starts(d_s):
    return np.arange(int(round((d_s - WIN) / STEP)) + 1) * STEP


def key_shape(c):
    """Shape of the key scene response at window center c (seconds into Commercial_1A)."""
    g = np.zeros_like(c)
    rise = (c >= KEY_START) & (c < KEY_START + 1.0)
    g[rise] = c[rise] - KEY_START
    hold = (c >= KEY_START + 1.0) & (c <= KEY_END)
    g[hold] = 1.0
    after = c > KEY_END
    g[after] = np.exp(-(c[after] - KEY_END) / 1.5)
    return g


def end_shape(c, d_s):
    a = d_s - END_LEN
    g = np.zeros_like(c)
    rise = (c >= a) & (c < a + 1.0)
    g[rise] = c[rise] - a
    g[c >= a + 1.0] = 1.0
    return g


def union_len(iv):
    """Total length of the union of intervals [(a, b), ...]."""
    if not iv:
        return 0.0
    iv = sorted(iv)
    tot, (ca, cb) = 0.0, iv[0]
    for a, b in iv[1:]:
        if a > cb:
            tot += cb - ca
            ca, cb = a, b
        else:
            cb = max(cb, b)
    return tot + cb - ca


def ar1_all(rng, n_resp):
    """All window noise of the study at once: shape (respondent, stimulus slot, series, window).
    Series 0 and 1 are the common fluctuations (Theta, Alpha); 2 to 9 the Theta channel noise;
    10 to 17 the Alpha channel noise. One loop over the 119 windows keeps the scan fast."""
    phi = np.array([COMMON_PHI] * 2 + [NOISE_PHI] * 16)[None, None, :]
    sdv = np.array([COMMON_SD] * 2 + [NOISE_SD] * 16)[None, None, :]
    z = rng.standard_normal((n_resp, 4, 18, NWMAX))
    out = np.empty_like(z)
    out[..., 0] = z[..., 0] * sdv
    innov = sdv * np.sqrt(1 - phi ** 2)
    for t in range(1, NWMAX):
        out[..., t] = phi * out[..., t - 1] + innov * z[..., t]
    return out


def lognorm(rng, med, logsd):
    return med * math.exp(logsd * rng.standard_normal())


def pearson(x, y):
    mx, my = st.mean(x), st.mean(y)
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    den = math.sqrt(sum((a - mx) ** 2 for a in x) * sum((b - my) ** 2 for b in y))
    return sxy / den if den > 0 else 0.0


# ---------------------------------------------------------------------------
# artifact events
# ---------------------------------------------------------------------------
def draw_events(rng, rate, d_s):
    """Supra-threshold events for one stimulus: list of (start, end, set of channel indexes). Each
    channel an event reaches gets its own entry, with the duration jittered by up to 15 percent,
    because the same blink does not cross the threshold for exactly as long at every electrode."""
    ev = []
    for _ in range(rng.poisson(rate * d_s / 60.0)):
        if rng.random() < P_MOVE:
            dur = clip(lognorm(rng, MOVE_MED, MOVE_LOGSD), 0.2, 4.0)
            chans, p_hit = range(8), 0.95
        else:
            dur = clip(lognorm(rng, BLINK_MED, BLINK_LOGSD), 0.05, 0.6)
            chans, p_hit = sorted(FRONTAL), 0.85
        a = rng.uniform(0, d_s - dur)
        for c in chans:
            if rng.random() < p_hit:
                ev.append((a, min(d_s, a + dur * rng.uniform(0.85, 1.15)), {c}))
    return ev


def pct_removed(ev, chan, recorded_s, extra=()):
    iv = [(a, b) for a, b, cs in ev if chan in cs] + list(extra)
    return 100.0 * union_len(iv) / recorded_s


def add_movement_until(rng, ev, d_s, target, recorded_s, gap=None):
    """Add movement events until the larger of F3 and F4 percent removed reaches target."""
    for _ in range(400):
        cur = max(pct_removed(ev, I_F3, recorded_s), pct_removed(ev, I_F4, recorded_s))
        if cur >= target:
            break
        dur = clip(lognorm(rng, 1.2, 0.5), 0.3, 4.0)
        a = rng.uniform(0, d_s - dur)
        if gap and not (a + dur <= gap[0] or a >= gap[1]):
            continue
        for c in range(8):
            if rng.random() < 0.95:
                ev.append((a, min(d_s, a + dur * rng.uniform(0.9, 1.1)), {c}))
    return ev


# ---------------------------------------------------------------------------
# the generator
# ---------------------------------------------------------------------------
def generate(seed, n_resp=N_RESP, special=True):
    rng = np.random.default_rng(seed)
    ids = [f"P{i + 1:02d}" for i in range(n_resp)]

    # ---- respondents: block randomized groups ------------------------------
    groups = []
    for _ in range(n_resp // 4):
        blk = ["A_C1first", "A_C2first", "B_C1first", "B_C2first"]
        rng.shuffle(blk)
        groups += blk
    group = dict(zip(ids, groups))
    version = {i: group[i][0] for i in ids}
    order = {i: group[i][2:] for i in ids}
    sexes = ["F"] * round(n_resp * 0.56) + ["M"] * (n_resp - round(n_resp * 0.56))
    rng.shuffle(sexes)
    sex = dict(zip(ids, sexes))
    age = {i: int(round(clip(rng.normal(21.0, 2.6), 18, 34))) for i in ids}

    # ---- special respondents (A6, A7) --------------------------------------
    sp = {}
    if special:
        a_ids = [i for i in ids if version[i] == "A"]
        b_ids = [i for i in ids if version[i] == "B"]
        pick = rng.choice(a_ids, size=2, replace=False)
        sp["lostF4"], sp["lslgap"] = str(pick[0]), str(pick[1])
        sp["movement"] = str(rng.choice(b_ids))
        rest_pool = [i for i in ids if i not in sp.values()]
        sp["restless"] = [str(x) for x in rng.choice(rest_pool, size=N_RESTLESS, replace=False)]
    restless = set(sp.get("restless", []))

    # ---- person parameters ---------------------------------------------------
    trait = rng.normal(TRAIT_MEAN, TRAIT_SD, n_resp)
    u_shared = rng.standard_normal(n_resp)
    key_resp = rng.normal(0, KEY_RESP_SD, n_resp)
    p_lev = {b: rng.normal(0, PERSON_SD[b], n_resp) for b in BANDS}
    p_ch = {b: rng.normal(0, PERSON_CH_SD, (n_resp, 8)) for b in BANDS}
    for b in BANDS:
        p_ch[b][:, I_F4] = p_ch[b][:, I_F3]                 # F3 and F4 share (A1)
    rate = np.exp(math.log(RATE_MEDIAN) + RATE_LOGSD * rng.standard_normal(n_resp))
    style = rng.standard_normal(n_resp)
    noise_all = ar1_all(rng, n_resp)

    # ---- per stimulus: timing, noise, FA, power, artifacts -------------------
    rows = {}            # (rid, label) -> dict with FA, default FA, PSD, artifacts
    timing = {}
    for k, rid in enumerate(ids):
        c1 = "Commercial_1A" if version[rid] == "A" else "Commercial_1B"
        seq = ["Baseline_1", c1, "Baseline_2", "Commercial_2"] if order[rid] == "C1first" else \
              ["Baseline_1", "Commercial_2", "Baseline_2", c1]
        t = float(rng.integers(*LEAD_S))
        starts = {}
        for j, lab in enumerate(seq):
            starts[lab] = t
            t += DUR[lab]
            if j in (1, 3):
                t += float(rng.integers(*RATING_S))
        timing[rid] = (seq, starts)
        r_rate = rate[k] * (2.5 if rid in restless else 1.0) * (4.0 if rid == sp.get("movement") else 1.0)
        resp_ik = {lab: RESP_SD * (math.sqrt(RESP_SHARED) * u_shared[k] + math.sqrt(1 - RESP_SHARED) * rng.standard_normal())
                   for lab in (c1, "Commercial_2")}
        end_ik = {lab: rng.normal(0, END_RESP_SD) for lab in (c1, "Commercial_2")}
        restless_stim = seq[int(rng.integers(0, 4))] if rid in restless else None
        restless_target = float(rng.uniform(15.5, 24.0)) if rid in restless else None
        for j, lab in enumerate(seq):
            d_s = DUR[lab]
            ws = win_starts(d_s)
            cen = ws + WIN / 2
            nw = len(ws)
            is_comm = lab.startswith("Commercial")
            # --- FA state F at each window (A2, A3)
            state = rng.normal(0, STATE_SD)
            F = np.full(nw, trait[k] + state)
            if is_comm:
                F += LEVEL[lab] + resp_ik[lab] + (POS_SECOND if j == 3 else 0.0)
                F += (END_AMP[lab] + end_ik[lab]) * end_shape(cen, d_s)
                if lab == "Commercial_1A":
                    F += (KEY_AMP + key_resp[k]) * key_shape(cen)
            # --- ln power per band, channel, window (A1, A4)
            lnp = {}
            for b in BANDS:
                base = np.array(LEVEL_DB[b]) / LN10_10 + p_lev[b][k] + p_ch[b][k]
                if is_comm:
                    if b == "Alpha":
                        base = base + ALPHA_COMM
                    else:
                        base = base + np.array([THETA_COMM.get(c, THETA_COMM_OTHER) for c in range(8)])
                bi = BANDS.index(b)
                common = noise_all[k, j, bi, :nw][None, :]
                noise = noise_all[k, j, 2 + 8 * bi: 10 + 8 * bi, :nw]
                arr = base[:, None] + common + noise
                if b == "Alpha":
                    arr[I_F4] += F / 2
                    arr[I_F3] -= F / 2
                if special and rid == sp.get("lostF4"):
                    arr[I_F4] += F4_NOISE
                lnp[b] = arr
            # --- artifacts (A5 to A7)
            ev = draw_events(rng, r_rate, d_s)
            recorded = d_s
            gap = None
            if special and rid == sp.get("lslgap") and lab == "Commercial_2":
                gap = (GAP_START, GAP_START + GAP_LEN)
                recorded = d_s - GAP_LEN
                ev = [e for e in ev if e[1] <= gap[0] or e[0] >= gap[1]]
                for c in range(8):
                    ev.append((gap[1], min(d_s, gap[1] + DRIFT_LEN * rng.uniform(0.95, 1.05)), {c}))
            if special and rid == sp.get("lostF4"):
                for _ in range(rng.poisson(30.0 * d_s / 60.0)):
                    dur = clip(lognorm(rng, 0.8, 0.5), 0.2, 3.0)
                    a = rng.uniform(0, d_s - dur)
                    ev.append((a, a + dur, {I_F4}))
            if special and rid == sp.get("movement") and lab in ("Commercial_1B", "Commercial_2"):
                ev = add_movement_until(rng, ev, d_s, 29.4 if lab == "Commercial_1B" else 21.0, recorded)
            if lab == restless_stim:
                ev = add_movement_until(rng, ev, d_s, restless_target, recorded)
            # --- clean window masks per channel
            clean = np.ones((8, nw), dtype=bool)
            for a, b_, cs in ev:
                hit = (a < ws + WIN) & (b_ > ws)
                for c in cs:
                    clean[c] &= ~hit
            if gap:
                clean &= ~((gap[0] < ws + WIN) & (gap[1] > ws))[None, :]
            pct = {c: round(pct_removed(ev, c, recorded), 2) for c in (I_F3, I_F4)}
            # --- segments: the stimulus and its annotations
            segs = [("Stimulus", lab, 0.0, d_s)]
            if lab == "Commercial_1A":
                segs.append(("Annotation", "KeyScene", KEY_START, KEY_END))
            if is_comm:
                segs.append(("Annotation", "EndScene", d_s - END_LEN, d_s))
            fa_w = lnp["Alpha"][I_F4] - lnp["Alpha"][I_F3]
            dfa_w = lnp["Alpha"][I_P3] - lnp["Alpha"][I_FZ]          # default channels Ch7 minus Ch1
            for typ, name, a, b_ in segs:
                inside = (ws >= a - 1e-9) & (ws + WIN <= b_ + 1e-9)
                m_fa = inside & clean[I_F3] & clean[I_F4]
                m_dfa = inside & clean[I_FZ] & clean[I_P3]
                psd = {}
                mean_pow = {b: [] for b in BANDS}
                for b in BANDS:
                    for c in range(8):
                        m = inside & clean[c]
                        if m.sum() == 0:
                            psd[(b, c)] = None
                            continue
                        pw = float(np.exp(lnp[b][c][m]).mean())
                        psd[(b, c)] = LN10_10 * math.log(pw)
                        mean_pow[b].append(pw)
                    psd[(b, "cluster")] = LN10_10 * math.log(st.mean(mean_pow[b])) if mean_pow[b] else None
                key = (rid, name if typ == "Stimulus" else f"{name}@{lab}")
                rows[key] = dict(
                    rid=rid, type=typ, label=name, parent=lab, start_ms=int(round((starts[lab] + a) * 1000)),
                    dur_ms=int(round((b_ - a) * 1000)),
                    fa=float(fa_w[m_fa].mean()) if m_fa.any() else None, n_fa=int(m_fa.sum()), n_win=int(inside.sum()),
                    dfa=float(dfa_w[m_dfa].mean()) if m_dfa.any() else None, psd=psd)
            rows[(rid, lab)]["pct"] = pct
            rows[(rid, lab)]["recorded_s"] = recorded
            rows[(rid, lab)]["latent"] = (resp_ik[lab] + state) if is_comm else state
            rows[(rid, lab)]["order_pos"] = j

    # ---- survey (A8) ---------------------------------------------------------
    survey = {}
    lat_sd = math.sqrt(RESP_SD ** 2 + STATE_SD ** 2)
    for k, rid in enumerate(ids):
        c1 = "Commercial_1A" if version[rid] == "A" else "Commercial_1B"
        rec = {}
        for lab, key in ((c1, "Liking_Commercial_1_1to7"), ("Commercial_2", "Liking_Commercial_2_1to7")):
            z = rows[(rid, lab)]["latent"] / lat_sd
            w = LIKE_W * z + math.sqrt(1 - LIKE_W ** 2) * (math.sqrt(LIKE_STYLE) * style[k] + math.sqrt(1 - LIKE_STYLE) * rng.standard_normal())
            rec[key] = int(clip(round(LIKE_MEAN[lab] + LIKE_SD * w), 1, 7))
        if version[rid] == "A":
            p = 1 / (1 + math.exp(-(1.45 + 1.2 * key_resp[k] / KEY_RESP_SD)))
            rec["KeyScene_Recall"] = "Yes" if rng.random() < p else "No"
        else:
            rec["KeyScene_Recall"] = "NA"
        rec["Brand_Named_Correctly"] = "Yes" if rng.random() < 0.70 else "No"
        survey[rid] = rec

    data = dict(seed=seed, ids=ids, group=group, version=version, order=order, sex=sex, age=age,
                rows=rows, timing=timing, survey=survey, sp=sp, trait=dict(zip(ids, trait)))
    if special:
        data["log"] = session_log(rng, data)
    return data


# ---------------------------------------------------------------------------
# session log (A9)
# ---------------------------------------------------------------------------
def session_log(rng, data):
    ids, sp, rows = data["ids"], data["sp"], data["rows"]
    days = weekdays(DAYS0, (len(ids) + 2) // 3)
    log = []
    routine = ["None", "None", "None", "None", "None", "None",
               "Participant sneezed during Baseline_2; live marker set",
               "Door noise in the hallway during the second commercial",
               "Participant adjusted the chin strap between the two commercials",
               "Phone vibrated in a bag at the start of Baseline_1; bag moved out of the room",
               "Participant asked a question during the first rating slide"]
    for k, rid in enumerate(ids):
        q = {}
        for site in ("F3", "F4"):
            q[site] = "Green" if rng.random() < 0.82 else "Yellow"
        cap = rng.choice(["S", "M", "M", "M", "L"])
        setup = int(clip(round(rng.normal(16, 3)), 10, 25))
        incident = str(rng.choice(routine))
        if rng.random() < 0.10:
            site = str(rng.choice(["F3", "F4"]))
            incident = f"{site} Red at the first check; hair parted, gel added, electrode pressed for 10 s; {site} Green before recording"
            q[site] = "Green"
            setup += 4
        complete, note = "Yes", ""
        maxp = max(max(rows[(rid, lab)]["pct"].values()) for lab in data["timing"][rid][0])
        worst = max(((lab, ch, v) for lab in data["timing"][rid][0] for ch, v in
                     (("F3", rows[(rid, lab)]["pct"][I_F3]), ("F4", rows[(rid, lab)]["pct"][I_F4]))), key=lambda x: x[2])
        if rid == sp["lostF4"]:
            q["F4"] = "Red"
            setup = 34
            incident = ("F4 Red in NIC2 Liveview at every check: electrode re-seated, hair parted, gel added twice, F4 and F3 leads "
                        "swapped (the fault stayed at the F4 site); recorded because the participant was present; F4 Red before recording")
            note = (f"EXCLUDE: Red NIC2 quality at F4 that could not be fixed (rule written before collection); the artifact report "
                    f"also shows {worst[2]:.2f} percent of {worst[1]} samples removed in {worst[0]}")
        elif rid == sp["movement"]:
            incident = ("Frequent large blinks and head movement from the first commercial on (participant reported dry contact "
                        "lenses); reminded twice to sit still")
            note = (f"EXCLUDE: {worst[2]:.2f} percent of {worst[1]} samples removed in {worst[0]}, above the 25 percent rule "
                    "written before collection")
        elif rid == sp["lslgap"]:
            complete = "No"
            incident = (f"LSL stream gap of about {GAP_LEN:.1f} s during Commercial_2 (no samples in the iMotions sensor widget); "
                        f"the stream reconnected without operator action; slow drift on all channels for about {DRIFT_LEN:.0f} s after it")
            note = (f"EXCLUDE: {worst[2]:.2f} percent of {worst[1]} samples removed in {worst[0]}, above the 25 percent rule "
                    "written before collection; Samples recorded is short of Samples expected for that stimulus")
        elif rid in sp["restless"]:
            incident = "Fidgeting and frequent blinking; reminded once to sit still"
            note = (f"Keep: largest percent removed {worst[2]:.2f} ({worst[1]}, {worst[0]}) is below the 25 percent rule; "
                    "dropped in the 15 percent sensitivity analysis")
        elif "Yellow" in q.values():
            note = "Keep: Yellow at the last impedance check is within the lab target (below 10 kOhm)"
        log.append({"Respondent Name": rid, "Session Date": days[k // 3].isoformat(), "Operator": OPERATORS[(k // 3) % 2],
                    "Version": data["version"][rid], "Order": data["order"][rid], "Cap size": str(cap), "Connection": "USB",
                    "NIC2 quality F3 before recording": q["F3"], "NIC2 quality F4 before recording": q["F4"],
                    "Gel used": "Yes", "Setup minutes": setup, "Incidents": incident, "Data complete": complete,
                    "Decision note": note})
    return log


# ---------------------------------------------------------------------------
# realized statistics for the kept respondents (the scan and the console summary use them; the
# logic mirrors analysis.R)
# ---------------------------------------------------------------------------
def r4(x):
    return None if x is None else round(x, 4)


def excluded(data, pct_limit=EXCLUDE_PCT):
    bad = set()
    for rid in data["ids"]:
        for lab in data["timing"][rid][0]:
            if max(data["rows"][(rid, lab)]["pct"].values()) > pct_limit:
                bad.add(rid)
    for lg in data.get("log", []):
        if "Red" in (lg["NIC2 quality F3 before recording"], lg["NIC2 quality F4 before recording"]):
            bad.add(lg["Respondent Name"])
    return bad


def change_table(data, kept, field="fa"):
    R = data["rows"]
    out = {}
    for rid in kept:
        seq, _ = data["timing"][rid]
        c1 = "Commercial_1A" if data["version"][rid] == "A" else "Commercial_1B"
        pre = {seq[1]: seq[0], seq[3]: seq[2]}
        def g(key):
            v = R[key][field]
            return r4(v) if field != "theta" else v
        rec = dict(version=data["version"][rid], order=data["order"][rid])
        b1, b2 = g((rid, pre[c1])), g((rid, pre["Commercial_2"]))
        rec["bl_c1"], rec["bl_c2"] = b1, b2
        rec["c1"], rec["c2"] = g((rid, c1)), g((rid, "Commercial_2"))
        rec["es1"], rec["es2"] = g((rid, f"EndScene@{c1}")), g((rid, "EndScene@Commercial_2"))
        rec["ks"] = g((rid, "KeyScene@Commercial_1A")) if rec["version"] == "A" else None
        rec["b_1"], rec["b_2"] = g((rid, "Baseline_1")), g((rid, "Baseline_2"))
        out[rid] = rec
    return out


def sd(x):
    return st.stdev(x) if len(x) > 1 else float("nan")


def realized(data):
    bad = excluded(data)
    kept = [i for i in data["ids"] if i not in bad]
    T = change_table(data, kept)
    o = {}
    A = [i for i in kept if T[i]["version"] == "A"]
    B = [i for i in kept if T[i]["version"] == "B"]
    o["n_kept"], o["n_A"], o["n_B"] = len(kept), len(A), len(B)
    if any(T[i]["ks"] is None for i in A) or any(T[i][f] is None for i in kept for f in ("c1", "c2", "es1", "es2", "b_1", "b_2")):
        o["na_ok"] = 0.0
        return o
    o["na_ok"] = 1.0
    bl = [(T[i]["b_1"] + T[i]["b_2"]) / 2 for i in kept]
    o["bl_mean"], o["bl_sd"] = st.mean(bl), sd(bl)
    o["bl_r"] = pearson([T[i]["b_1"] for i in kept], [T[i]["b_2"] for i in kept])
    o["bl_diff"] = st.mean(T[i]["b_2"] - T[i]["b_1"] for i in kept)
    ks = [T[i]["ks"] - T[i]["bl_c1"] for i in A]
    o["ks_mean"], o["ks_sd"] = st.mean(ks), sd(ks)
    o["dz_h1"] = o["ks_mean"] / o["ks_sd"]
    esA = [T[i]["es1"] - T[i]["bl_c1"] for i in A]
    esB = [T[i]["es1"] - T[i]["bl_c1"] for i in B]
    sp_ = math.sqrt(((len(esA) - 1) * sd(esA) ** 2 + (len(esB) - 1) * sd(esB) ** 2) / (len(esA) + len(esB) - 2))
    o["es_diff"], o["es_sd"] = st.mean(esA) - st.mean(esB), sp_
    o["d_h2"] = o["es_diff"] / sp_
    c1c = {i: T[i]["c1"] - T[i]["bl_c1"] for i in kept}
    c2c = {i: T[i]["c2"] - T[i]["bl_c2"] for i in kept}
    dif = [c2c[i] - c1c[i] for i in kept]
    o["h3_diff"], o["h3_sd"] = st.mean(dif), sd(dif)
    o["dz_h3"] = o["h3_diff"] / o["h3_sd"]
    o["inter"] = st.mean(c2c[i] - c1c[i] for i in A) - st.mean(c2c[i] - c1c[i] for i in B)
    o["c_change_sd"] = (sd(list(c1c.values())) + sd(list(c2c.values()))) / 2
    o["cov_c"] = st.covariance(list(c1c.values()), list(c2c.values()))
    sv = data["survey"]
    o["like_r1"] = pearson([c1c[i] for i in kept], [sv[i]["Liking_Commercial_1_1to7"] for i in kept])
    o["like_r2"] = pearson([c2c[i] for i in kept], [sv[i]["Liking_Commercial_2_1to7"] for i in kept])
    R = data["rows"]
    th = []
    for i in kept:
        seq, _ = data["timing"][i]
        for cl, bl_ in ((seq[1], seq[0]), (seq[3], seq[2])):
            th.append(R[(i, cl)]["psd"][("Theta", 0)] - R[(i, bl_)]["psd"][("Theta", 0)])
    o["theta_chg"] = st.mean(th)
    D = change_table(data, kept, "dfa")
    if any(D[i][f] is None for i in kept for f in ("bl_c1", "c1", "c2", "es1")) or any(D[i]["ks"] is None for i in A):
        o["na_ok"] = 0.0
        return o
    dks = [D[i]["ks"] - D[i]["bl_c1"] for i in A]
    o["dks_mean"] = st.mean(dks)
    o["dks_t"] = st.mean(dks) / (sd(dks) / math.sqrt(len(dks)))
    strict = excluded(data, STRICT_PCT)
    o["n_strict_extra"] = len(strict - bad)
    o["excl_ok"] = 1.0 if (len(bad) == 3 and o["n_A"] == 22 and o["n_B"] == 23) else 0.0
    return o


# Target value and tolerance (about one standard error of the statistic in a sample of 45, or of
# 22 and 23 per version). The planned values follow from the model in A2 and A3 (see --calibrate).
TARGETS = {
    "bl_mean": (-0.05, 0.02), "bl_sd": (0.127, 0.015), "bl_r": (0.85, 0.04), "bl_diff": (0.0, 0.012),
    "ks_mean": (0.10, 0.025), "ks_sd": (0.142, 0.02), "dz_h1": (0.70, 0.10),
    "es_diff": (0.045, 0.02), "es_sd": (0.118, 0.012), "d_h2": (0.38, 0.08),
    "h3_diff": (0.032, 0.012), "h3_sd": (0.108, 0.012), "dz_h3": (0.29, 0.08),
    "inter": (-0.023, 0.03), "c_change_sd": (0.083, 0.008),
    "like_r1": (0.28, 0.10), "like_r2": (0.28, 0.10), "theta_chg": (0.35, 0.10),
    "dks_mean": (0.0, 0.03),
}


def loss(o):
    if o.get("na_ok", 0) < 1 or o.get("excl_ok", 0) < 1:
        return 1e9
    base = sum(((o[k] - t) / tol) ** 2 for k, (t, tol) in TARGETS.items())
    base += 1000.0 * (o["n_strict_extra"] != N_RESTLESS)
    base += 1000.0 * (o["cov_c"] <= 0.0004)          # keeps the random intercept away from zero
    return base


# ---------------------------------------------------------------------------
# writing
# ---------------------------------------------------------------------------
def banner_text(seed, extra):
    return (f"# SIMULATED DATA for teaching. Generated by make_simulated_data.py (seed {seed}). "
            f"No real person was recorded. {extra}")


B_FA = ("Columns follow the EEG Frontal Asymmetry Metrics export (R Processed EEG Export Files article; only the Alpha band selected). "
        "MODELED: Type values, Parent Stimulus on stimulus rows, and the single Algorithm Device Name column. ASSUMED channel map, verify in NIC2: "
        "Ch5 = F3 (left), Ch6 = F4 (right).")
M_FA = ("# METADATA (modeled): Flow = EEG Frontal Asymmetry R Notebook; Device = Lab Streaming Layer, Enobio EEG, profile Enobio, 500 Hz; "
        "Left frontal electrode set = Ch5; Right frontal electrode set = Ch6; Powerbands = Alpha (8 - 12 Hz); Welch PSD (50% overlap), "
        "window 1 s; Band-pass 0.5 - 100 Hz; Notch 60 Hz; artifact threshold 120 microvolts; FA = ln(right alpha) - ln(left alpha)")
B_DFA = ("TEACHING CONTRAST: the same export as if the FA notebook had been run with its Enobio DEFAULTS (left Ch1, right Ch7), which in the "
         "ASSUMED montage are Fz and P3, not a left and right frontal pair. Do not analyze these values as frontal asymmetry.")
M_DFA = ("# METADATA (modeled): Flow = EEG Frontal Asymmetry R Notebook (factory defaults); Left frontal electrode set = Ch1; "
         "Right frontal electrode set = Ch7; all other settings as in SIMULATED_eeg_fa_metrics.csv")
B_PSD = ("Columns follow the EEG Power Spectral Density Metrics export (Theta and Alpha selected, all 8 channels in the cluster). MODELED: "
         "column order, Type values, Parent Stimulus on stimulus rows, dB = 10 log10(mean power), and the single Algorithm Device Name column. "
         "ASSUMED channel map: Ch1 Fz, Ch2 Cz, Ch3 Pz, Ch4 Oz, Ch5 F3, Ch6 F4, Ch7 P3, Ch8 P4.")
B_ART = ("MODELED layout: values copied by the lab from the notebook report box Percentage of data filtered (nb of samples filtered / "
         "nb total of samples), one row per respondent, stimulus, and channel. Percent removed (%) = Samples removed / Samples recorded x 100. "
         "Verify the real report layout (and whether it gives values per stimulus) on the lab computer.")
B_LOG = ("MODELED operator log. NIC2 quality = the Red, Yellow, or Green Liveview status at the last check before recording. "
         "Connection USB as the iMotions Enobio article recommends.")
B_SV = "MODELED questionnaire: liking of each commercial (1 to 7) on its rating slide; recall and brand questions at the end; KeyScene_Recall is NA for Version B."

ID_COLS = ["Study Name", "Respondent Name", "Gender", "Age", "Group", "Type", "Label", "Start (ms)", "Duration (ms)", "Parent Stimulus"]


def fmt(x, nd):
    return "NA" if x is None else f"{x:.{nd}f}"


def ordered_keys(data, rid):
    seq, starts = data["timing"][rid]
    keys = [(rid, lab) for lab in seq]
    ann = []
    for lab in seq:
        if lab == "Commercial_1A":
            ann.append((rid, "KeyScene@Commercial_1A"))
        if lab.startswith("Commercial"):
            ann.append((rid, f"EndScene@{lab}"))
    ann.sort(key=lambda kk: data["rows"][kk]["start_ms"])
    return keys + ann


def id_part(data, row):
    rid = row["rid"]
    return {"Study Name": STUDY_NAME, "Respondent Name": rid, "Gender": data["sex"][rid], "Age": data["age"][rid],
            "Group": data["group"][rid], "Type": row["type"], "Label": row["label"], "Start (ms)": row["start_ms"],
            "Duration (ms)": row["dur_ms"], "Parent Stimulus": row["parent"]}


def write_csv(name, columns, rows, banner, extra_comments=()):
    path = os.path.join(HERE, name)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        fh.write(banner + "\n")
        for c in extra_comments:
            fh.write(c + "\n")
        w = csv.DictWriter(fh, fieldnames=columns, lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f"wrote {path} ({len(rows)} rows)")


def write_all(data):
    seed = data["seed"]
    fa_rows, dfa_rows, psd_rows, art_rows = [], [], [], []
    psd_cols = [f"Mean PSD {c} {b} (dB)" for c in CH for b in BANDS] + [f"Mean PSD Electrode Cluster Average {b} (dB)" for b in BANDS]
    for rid in data["ids"]:
        for key in ordered_keys(data, rid):
            row = data["rows"][key]
            base = id_part(data, row)
            fa_rows.append({**base, "Mean Frontal Asymmetry Alpha": fmt(row["fa"], 4), "Algorithm Device Name": FA_DEVICE_CELL})
            dfa_rows.append({**base, "Mean Frontal Asymmetry Alpha": fmt(row["dfa"], 4), "Algorithm Device Name": FA_DEVICE_CELL})
            pr = dict(base)
            for ci, c in enumerate(CH):
                for b in BANDS:
                    pr[f"Mean PSD {c} {b} (dB)"] = fmt(row["psd"][(b, ci)], 3)
            for b in BANDS:
                pr[f"Mean PSD Electrode Cluster Average {b} (dB)"] = fmt(row["psd"][(b, "cluster")], 3)
            pr["Algorithm Device Name"] = PSD_DEVICE_CELL
            psd_rows.append(pr)
            if row["type"] == "Stimulus":
                exp_n = int(round(DUR[row["label"]] * FS))
                rec_n = int(round(row["recorded_s"] * FS))
                for ci in (I_F3, I_F4):
                    pct = row["pct"][ci]
                    rem = int(round(pct / 100.0 * rec_n))
                    art_rows.append({"Study Name": STUDY_NAME, "Respondent Name": rid, "Group": data["group"][rid],
                                     "Label": row["label"], "Channel": CH[ci], "Electrode (assumed)": SITE[CH[ci]],
                                     "Duration (ms)": row["dur_ms"], "Samples expected": exp_n, "Samples recorded": rec_n,
                                     "Samples removed": rem, "Percent removed (%)": f"{100.0 * rem / rec_n:.2f}"})
    write_csv("SIMULATED_eeg_fa_metrics.csv", ID_COLS + ["Mean Frontal Asymmetry Alpha", "Algorithm Device Name"], fa_rows,
              banner_text(seed, B_FA), extra_comments=[M_FA])
    write_csv("SIMULATED_eeg_fa_metrics_default_channels.csv", ID_COLS + ["Mean Frontal Asymmetry Alpha", "Algorithm Device Name"],
              dfa_rows, banner_text(seed, B_DFA), extra_comments=[M_DFA])
    write_csv("SIMULATED_eeg_psd_metrics.csv", ID_COLS + psd_cols + ["Algorithm Device Name"], psd_rows, banner_text(seed, B_PSD))
    write_csv("SIMULATED_eeg_artifact_report.csv", list(art_rows[0].keys()), art_rows, banner_text(seed, B_ART))
    write_csv("SIMULATED_eeg_session_log.csv", list(data["log"][0].keys()), data["log"], banner_text(seed, B_LOG))
    sv_rows = []
    for rid in data["ids"]:
        s = data["survey"][rid]
        sv_rows.append({"Study Name": STUDY_NAME, "Respondent Name": rid, "Group": data["group"][rid], "Gender": data["sex"][rid],
                        "Age": data["age"][rid], "Liking_Commercial_1_1to7": s["Liking_Commercial_1_1to7"],
                        "Liking_Commercial_2_1to7": s["Liking_Commercial_2_1to7"], "KeyScene_Recall": s["KeyScene_Recall"],
                        "Brand_Named_Correctly": s["Brand_Named_Correctly"]})
    write_csv("SIMULATED_eeg_survey.csv", list(sv_rows[0].keys()), sv_rows, banner_text(seed, B_SV))


def calibrate():
    """Population check of the model: a large simulated sample without the special respondents."""
    d = generate(1, n_resp=4000, special=False)
    kept = d["ids"]
    T = change_table(d, kept)
    A = [i for i in kept if T[i]["version"] == "A"]
    B = [i for i in kept if T[i]["version"] == "B"]
    ks = [T[i]["ks"] - T[i]["bl_c1"] for i in A if T[i]["ks"] is not None]
    print(f"population check, n = {len(kept)} (no special respondents)")
    print(f"  KeyScene change: mean {st.mean(ks):.4f} SD {sd(ks):.4f} d_z {st.mean(ks) / sd(ks):.3f}   (plan 0.10, 0.142, 0.70)")
    esA = [T[i]["es1"] - T[i]["bl_c1"] for i in A]
    esB = [T[i]["es1"] - T[i]["bl_c1"] for i in B]
    spd = math.sqrt((sd(esA) ** 2 + sd(esB) ** 2) / 2)
    print(f"  EndScene C1 change: A {st.mean(esA):.4f}, B {st.mean(esB):.4f}, diff {st.mean(esA) - st.mean(esB):.4f}, SD {spd:.4f}, d {(st.mean(esA) - st.mean(esB)) / spd:.3f}")
    c1c = [T[i]["c1"] - T[i]["bl_c1"] for i in kept]
    c2c = [T[i]["c2"] - T[i]["bl_c2"] for i in kept]
    dif = [b - a for a, b in zip(c1c, c2c)]
    print(f"  whole commercial change: C1 {st.mean(c1c):.4f} (SD {sd(c1c):.4f}), C2 {st.mean(c2c):.4f} (SD {sd(c2c):.4f}); "
          f"C2 - C1 {st.mean(dif):.4f}, SD {sd(dif):.4f}, d_z {st.mean(dif) / sd(dif):.3f}; cov {st.covariance(c1c, c2c):.5f}")
    bl = [T[i]["b_1"] for i in kept]
    print(f"  Baseline_1 FA mean {st.mean(bl):.4f} SD {sd(bl):.4f}; r(Baseline_1, Baseline_2) {pearson(bl, [T[i]['b_2'] for i in kept]):.3f}")
    sv = d["survey"]
    print(f"  liking r: C1 {pearson(c1c, [sv[i]['Liking_Commercial_1_1to7'] for i in kept]):.3f}, "
          f"C2 {pearson(c2c, [sv[i]['Liking_Commercial_2_1to7'] for i in kept]):.3f}")
    rec = [sv[i]["KeyScene_Recall"] == "Yes" for i in A]
    print(f"  KeyScene recall: {100 * st.mean(rec):.1f} percent Yes")
    pct = [v for (rid, lab), r in d["rows"].items() if "pct" in r for v in r["pct"].values()]
    print(f"  percent removed (F3, F4 rows): median {st.median(pct):.2f}, 90th percentile {np.percentile(pct, 90):.2f}, max {max(pct):.2f}")
    nfa = [r["n_fa"] / r["n_win"] for (rid, lab), r in d["rows"].items() if r["label"] == "KeyScene"]
    print(f"  share of KeyScene windows kept for FA: mean {st.mean(nfa):.3f}")
    D = change_table(d, kept, "dfa")
    dks = [D[i]["ks"] - D[i]["bl_c1"] for i in A if D[i]["ks"] is not None]
    print(f"  default channel (Ch7 - Ch1) KeyScene change: mean {st.mean(dks):.4f} SD {sd(dks):.4f}; "
          f"Baseline_1 default FA mean {st.mean(D[i]['b_1'] for i in kept):.3f}")


def main():
    args = sys.argv[1:]
    if args and args[0] == "--scan":
        n = int(args[1]) if len(args) > 1 else 200
        res = []
        for s in range(SCAN_BASE, SCAN_BASE + n):
            res.append((loss(realized(generate(s))), s))
        res.sort()
        print("rank  seed        loss (sum of squared standardized gaps to the model)")
        for k, (l, s) in enumerate(res[:10], start=1):
            print(f"{k:>4}  {s}  {l:9.2f}")
        ok = [l for l, _ in res if l < 1e8]
        print(f"usable seeds: {len(ok)} of {n}; median loss of the usable seeds: {st.median(ok):.2f}")
        return
    if args and args[0] == "--calibrate":
        calibrate()
        return
    seed = int(args[0]) if args else SEED
    data = generate(seed)
    write_all(data)
    o = realized(data)
    print(f"\nRealized statistics for the {o['n_kept']} kept respondents (A {o['n_A']}, B {o['n_B']}); loss {loss(o):.2f}")
    for k, (tgt, tol) in TARGETS.items():
        print(f"  {k:<14} {o[k]:9.4f}   target {tgt}")
    print(f"  extra exclusions under the {STRICT_PCT:.0f} percent rule: {o['n_strict_extra']}; covariance of the commercial changes {o['cov_c']:.5f}")
    print(f"  special respondents: {data['sp']}")


if __name__ == "__main__":
    main()
