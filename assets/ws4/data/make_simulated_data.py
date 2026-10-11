#!/usr/bin/env python3
"""
make_simulated_data.py
======================
SIMULATED DATA GENERATOR for Workshop 4: Skin Conductance (GSR) and Heart Rate (PPG) with the
Shimmer3 GSR+ in iMotions Lab 11.1, arousal during two video commercials
Martin's Family Clothing Consumer Behavior Research Laboratory, Jacksonville State University

EVERYTHING THIS SCRIPT WRITES IS SIMULATED. No person was recorded. The numbers are drawn from
the probability model documented below so that students can practice the analysis steps on
exports that look like the iMotions GSR R Notebook exports and a Sensor Data export.

The design is modeled on the lab's own library studies "Nike Campaign Ad Evaluation" and
"e.l.f. Campaign Ad Evaluation": a 60 s Baseline, the commercial, a 15 s Recovery, and a short
questionnaire. The Nike commercial is 60 s and the e.l.f. commercial is 30 s. Nothing about the
content of either commercial was used, and the time profile of peaks inside each commercial
(the onset response and the "late segment" bumps in KEY_MOMENTS) is invented for teaching.

Design (within subjects)
------------------------
30 respondents, P01 to P30. Every respondent watched BOTH sequences, in counterbalanced order
(Group A: Nike first; Group B: e.l.f. first; 15 each), with a 2 minute rest between the two
sequences. A sequence is Baseline_x, x_Commercial, Recovery_x, then the questionnaire. Stimulus
labels: Baseline_Nike, Nike_Commercial, Recovery_Nike, Baseline_elf, elf_Commercial,
Recovery_elf. Respondents were run two at a time (one Group A and one Group B respondent per
pair, 15 pairs, three pairs per day, 2026-11-02 to 2026-11-06). The lab owns two Shimmer3 GSR+
units: Seat 1 uses Shimmer3-8956 (the unit seen in the lab install) and Seat 2 uses
Shimmer3-B (a PLACEHOLDER name; verify the real name on the lab machine). The analysis pools
both devices, as the "Multiple device aggregation" option of the GSR Peak Detection and
Aggregation R Notebook does. Seat is assigned at random within each pair.

Files written (same folder)
---------------------------
SIMULATED_gsr_peak_metrics.csv   GSR Peak Metrics export, one row per respondent and stimulus.
                                 Columns exactly as the Export R Processed GSR Files article lists
                                 them: Study Name, Respondent Name, Gender, Age, Group, Type, Label,
                                 Start (ms), Duration (ms), Signal Duration (ms), Parent Stimulus,
                                 Has Peaks, Peak Count, Peaks Per Minute, Average Peak Amplitude
                                 (microSiemens). Type is "Stimulus" and Parent Stimulus repeats the
                                 label (MODELED: the manual does not print a Type or Parent Stimulus
                                 example row; verify on a real export). The export has no Device
                                 column, so the analysis joins Device from the session log.
SIMULATED_gsr_signal_quality.csv GSR Signal Quality export: Signal-to-Noise ratio in decibels and
                                 Signal missing in percent per stimulus and respondent. The two
                                 measures are named in the GSR Signal Quality article; the other
                                 columns (Device, Flow Name, Group) are MODELED. The real export is a
                                 .txt file; verify its layout.
SIMULATED_gsr_peak_table.csv     GSR Peak Table, one row per detected peak, with the columns the
                                 Export R Processed GSR Files article lists (the optional Scene and
                                 Annotation columns are left out). Peak times are in ms from the
                                 start of the recording. M2M Bin names the FIRST of the (overlapping)
                                 5 s bins that holds the peak apex (MODELED).
SIMULATED_heart_rate.csv         Mean heart rate (beats per minute) per respondent and stimulus.
                                 The manuals name NO PPG heart rate notebook and NO column name for
                                 the PPG beats per minute channel (FACTS section 7). This file is
                                 simulated as if the continuously calculated beats per minute value
                                 of the Shimmer3 PPG channel had been exported in Sensor Data and
                                 averaged over each stimulus. The column name "Mean PPG Heart Rate
                                 (BPM)" is MODELED and MUST BE VERIFIED on the lab machine.
SIMULATED_gsr_survey.csv         Questionnaire: Respondent Name, Group, Arousal_Nike_1to9,
                                 Arousal_elf_1to9 (self assessment manikin style), Liking_Nike_1to7,
                                 Liking_elf_1to7.
SIMULATED_gsr_session_log.csv    Operator log: Respondent Name, Session Date, Operator, Seat, Device,
                                 Electrode placement (fingers), Skin note, Warm up minutes, Data
                                 complete, Decision note.

Model assumptions
-----------------
A1  Peaks Per Minute (the primary outcome), overall means and SDs across all 30 respondents
    (nonresponders included), planned:
        Baseline_Nike, Baseline_elf   mean 2.4, SD 1.3
        Nike_Commercial               mean 4.6, SD 2.0
        elf_Commercial                mean 3.7, SD 1.8
        Recovery_Nike, Recovery_elf   mean about 3.0 (SD about 2.6; a 15 s window gives only whole
                                      peak counts, so Peaks Per Minute moves in steps of 4)
    The planned paired effects follow: Nike minus its baseline 2.2 (d_z about 1.25), e.l.f. minus
    its baseline 1.3 (d_z about 0.8), Nike minus e.l.f. 0.9 (d_z about 0.5).
A2  Respondent propensity. Each respondent has one arousal propensity z_i ~ N(0, 1) shared by all
    six stimuli, so a respondent's values correlate. The shared share RHO_LAT of the latent
    variance was calibrated (python3 make_simulated_data.py --calibrate) so that the correlation
    among all respondents, including the nonresponders, is about .5 for the baseline and
    commercial pairs.
A3  Counts, not rates, are what iMotions detects. The latent rate (peaks per minute of signal) is
    turned into a whole number of peaks by stochastic rounding of rate x signal minutes (floor
    plus a Bernoulli draw for the fraction). This makes the counts more regular than a Poisson
    process, which suits GSR peaks (detected peaks cannot overlap), and it is the only way to
    reach the planned SDs: a Poisson count at a mean of 1.85 in the 30 s commercial would give an
    SD of Peaks Per Minute near 2.7, not 1.8. Peaks Per Minute = Peak Count / (Signal Duration in
    minutes), as the manual states.
A4  Nonresponders. Exactly 3 of the 30 respondents (10 percent) are nonresponders: their peak rate
    is a tenth to a quarter of the others (means 0.2 to 0.45 peaks per minute), so Has Peaks = 0
    in most stimuli. They are kept in the data. One of them has dry skin. The overall means in
    A1 include them, so the responders' latent means are somewhat higher (solved in
    latent_params()). SNORESP_RULE in the analysis (Has Peaks = 0 in at least 4 of the 6
    stimuli) flags them from the Peak Metrics file alone.
A5  Order effect (habituation): the commercial seen SECOND gets 0.4 fewer peaks per minute and the
    commercial seen first 0.4 more (a second minus first difference of -0.8; baselines and
    recoveries are not adjusted). Because the design is counterbalanced this averages out of the
    marginal means in A1.
A6  Average Peak Amplitude. Each peak amplitude is lognormal around a stimulus mean: Baseline 0.20,
    Nike_Commercial 0.28, Recovery_Nike 0.22, elf_Commercial 0.25, Recovery_elf 0.21
    microSiemens (slightly larger for Nike), times a respondent amplitude factor (log SD 0.28),
    times peak to peak noise (log SD 0.30), clipped to 0.05 to 0.90. Nonresponders' rare peaks are
    30 percent smaller. The row value is the mean of the peaks in that stimulus (NA when the
    stimulus has no peak, as the metric cannot be calculated).
A7  Heart rate (mean beats per minute per stimulus, from PPG). Resting level per respondent
    ~ Normal(74, 9) clipped to 52 to 100. Stimulus mean = resting level + stimulus offset +
    sequence level noise (SD 1.2, shared by the three stimuli of a sequence) + stimulus noise
    (SD 1.3 baseline, 1.6 commercial and recovery). Offsets: Baseline 0; Nike_Commercial -1.5;
    elf_Commercial -0.8; Recovery_Nike -0.4; Recovery_elf -0.3. The offsets stand for a small
    deceleration in the first part of each commercial (the orienting response); the file holds
    only the whole stimulus mean, so the commercial mean is about 1.5 BPM below its baseline for
    Nike and 0.8 BPM for e.l.f.
A8  Signal quality. Signal-to-Noise ratio (dB) ~ Normal(24.5, 3.4) per row (respondent factor SD 2.8,
    row noise SD 2.0), 1.0 dB lower at Seat 2, clipped to 9 to 34, so most rows lie between 18
    and 32. Two respondents have dry skin (about 7 dB lower, rows between 0 and 20 dB that need
    a look). Signal missing (%) is 0 for about two thirds of the rows and 0.1 to 4 percent
    otherwise (short gaps that are interpolated). Three rows carry one Bluetooth gap longer than
    4 s (4.2 to 5.6 s); a gap longer than 4 s shortens Signal Duration (ms) by the gap.
A9  Planned exclusions, rule written in advance (the lab's rule built from the GSR Signal Quality
    article, which calls an SNR below 0 dB poor and above 20 dB good): exclude a respondent whose
    SNR is below 0 dB or whose Signal missing exceeds 20 percent in any commercial.
      P11  loose electrode lead: SNR -3.4 dB in Nike_Commercial (and 2.6 dB in Recovery_Nike);
           the few "peaks" detected there are spurious (9 in the commercial, small amplitude).
      P24  Bluetooth dropout: Signal missing 38.0 percent in elf_Commercial (11.4 s gap), so
           Signal Duration is 18.6 s.
    All other respondents are kept, which leaves 28.
A10 Survey. Self reported arousal (1 to 9): Nike mean 5.9, e.l.f. mean 5.2 (SD 1.6), correlated
    with the Peaks Per Minute of the same commercial at about r = .35 for responders (the
    nonresponders report arousal like everyone else, so their rows lower the full sample r a
    little), plus a respondent response style shared by the two ratings. Liking (1 to 7): Nike
    mean 5.3 (SD 1.1), e.l.f. mean 5.0 (SD 1.2), correlation .4.
A11 Session log. Warm up minutes: the manuals state no warm up time (FACTS section 5), so the lab
    rule modeled here is 5 minutes (3 to 8 allowed; P11 had only 2). Electrode placement follows
    the Shimmer 3 GSR+ placement article: non-dominant hand, distal phalanges preferred, medial
    when the fingertips are callused, PPG probe on another free finger.

Run:  python3 make_simulated_data.py                 (default seed below)
      python3 make_simulated_data.py <seed>          (any other seed)
      python3 make_simulated_data.py --scan [n]      (rank n candidate seeds; writes nothing)
      python3 make_simulated_data.py --calibrate     (population check of the model; writes nothing)
"""

import csv
import math
import os
import random
import statistics as st
import sys
from datetime import date, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))

# Fixed seed so the files are reproducible. It was chosen from a scan of 20000 candidate seeds,
# 20264001 to 20284000 (python3 make_simulated_data.py --scan 20000), as the one whose realized
# sample of 28 kept respondents sits closest to the planned means, SDs, correlations, effect sizes,
# order effect, and survey values (see TARGETS below). The scan ranks seeds by closeness to the
# MODEL, never by a p value. Any single sample of 28 drifts from its model, which is itself a
# lesson: compare the realized descriptives with the assumptions above.
SEED = 20278504
SCAN_BASE = 20264001

STUDY_NAME = "SIMULATED_Arousal_Nike_vs_elf"
N_RESP = 30
N_NONRESP = 3
Q_NR = N_NONRESP / N_RESP
DEVICE = {1: "Shimmer3-8956", 2: "Shimmer3-B"}
FLOW = {1: "GSR_Seat1", 2: "GSR_Seat2"}
DAYS = [date(2026, 11, 2) + timedelta(days=k) for k in range(5)]
OPERATORS = ["Operator_1", "Operator_2", "Operator_1", "Operator_2", "Operator_1"]

# label, brand, phase, duration in ms
STIMS = [("Baseline_Nike", "Nike", "Baseline", 60000), ("Nike_Commercial", "Nike", "Commercial", 60000),
         ("Recovery_Nike", "Nike", "Recovery", 15000), ("Baseline_elf", "elf", "Baseline", 60000),
         ("elf_Commercial", "elf", "Commercial", 30000), ("Recovery_elf", "elf", "Recovery", 15000)]
LABELS = [s[0] for s in STIMS]
DUR = {s[0]: s[3] for s in STIMS}
BRAND = {s[0]: s[1] for s in STIMS}
PHASE = {s[0]: s[2] for s in STIMS}
QUESTIONNAIRE_MS = (35000, 80000)     # time on the questionnaire slide
REST_MS = 120000                      # the 2 minute rest between the two sequences
LEAD_MS = (20000, 40000)              # instruction slide before the first Baseline

# A1: planned overall Peaks Per Minute (mean, SD) including nonresponders
PPM_PLAN = {"Baseline_Nike": (2.4, 1.3), "Baseline_elf": (2.4, 1.3), "Nike_Commercial": (4.6, 2.0),
            "elf_Commercial": (3.7, 1.8), "Recovery_Nike": (3.0, 2.6), "Recovery_elf": (3.0, 2.6)}
# A4: nonresponders' mean Peaks Per Minute
NR_MEAN = {"Baseline_Nike": 0.20, "Baseline_elf": 0.20, "Nike_Commercial": 0.45, "elf_Commercial": 0.40,
           "Recovery_Nike": 0.20, "Recovery_elf": 0.20}
ROUND_VAR = 1.0 / 6.0     # variance (in squared counts) that stochastic rounding adds, on average
RHO_LAT = 0.30            # A2: share of the responders' latent variance that is shared (calibrated)
POS_ADJ = 0.4             # A5: +0.4 for the commercial seen first, -0.4 for the one seen second
# A6: stimulus mean peak amplitude (microSiemens)
AMP_MEAN = {"Baseline_Nike": 0.20, "Baseline_elf": 0.20, "Nike_Commercial": 0.28, "elf_Commercial": 0.25,
            "Recovery_Nike": 0.22, "Recovery_elf": 0.21}
AMP_RESP_SD, AMP_PEAK_SD, AMP_LO, AMP_HI = 0.28, 0.30, 0.05, 0.90
# A7: heart rate offsets (BPM) relative to the respondent's resting level
HR_OFFSET = {"Baseline_Nike": 0.0, "Baseline_elf": 0.0, "Nike_Commercial": -1.5, "elf_Commercial": -0.8,
             "Recovery_Nike": -0.4, "Recovery_elf": -0.3}
HR_NOISE = {"Baseline": 1.3, "Commercial": 1.6, "Recovery": 1.6}
HR_SEQ_SD = 1.2
# A8 / A9: signal quality and the two planned exclusions
SNR_MEAN, SNR_RESP_SD, SNR_ROW_SD, SEAT2_DB, DRY_DB = 24.5, 2.8, 2.0, -1.0, -7.0
P11, P24 = "P11", "P24"
P11_SNR = {"Nike_Commercial": -3.4, "Recovery_Nike": 2.6}
P11_SPURIOUS = {"Nike_Commercial": 9, "Recovery_Nike": 2}      # spurious peak counts from the loose lead
P24_MISSING_PCT = 38.0
EXCLUDE_SNR_DB, EXCLUDE_MISSING_PCT = 0.0, 20.0
# invented time profile inside the commercials: (fraction of the commercial, strength) of a late bump
KEY_MOMENTS = {"Nike_Commercial": [(0.40, 0.7), (0.78, 1.0)], "elf_Commercial": [(0.55, 0.9)]}
BIN_LEN_S, BIN_OVERLAP_S = 5, 1
BIN_STEP_S = BIN_LEN_S - BIN_OVERLAP_S
NR_RULE_MIN_ZERO = 4      # nonresponder rule: Has Peaks = 0 in at least 4 of the 6 stimuli


# ---------------------------------------------------------------------------
# helpers (all take the random.Random instance so a seed scan never touches global state)
# ---------------------------------------------------------------------------
def clip(x, lo, hi):
    return max(lo, min(hi, x))


def fmt(x, nd=0):
    return "NA" if x is None else f"{x:.{nd}f}"


def pearson(x, y):
    mx, my = st.mean(x), st.mean(y)
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    den = math.sqrt(sum((a - mx) ** 2 for a in x) * sum((b - my) ** 2 for b in y))
    return sxy / den if den > 0 else 0.0


def bins_for(d_s):
    """Complete 5 s bins with 1 s overlap inside a stimulus of d_s seconds: (start, end) in seconds."""
    out, k = [], 0
    while BIN_STEP_S * k + BIN_LEN_S <= d_s + 1e-9:
        out.append((BIN_STEP_S * k, BIN_STEP_S * k + BIN_LEN_S))
        k += 1
    return out


def first_bin_label(t_s, d_s):
    for k, (a, b) in enumerate(bins_for(d_s), start=1):
        if a <= t_s < b:
            return f"{k}[{a}-{b}s]"
    return "NA"


def latent_params(label):
    """Responders' latent mean and SD of Peaks Per Minute, solved so that the mixture of responders
    and nonresponders reproduces the planned overall mean and SD (A1, A3, A4)."""
    mean, sd = PPM_PLAN[label]
    t = DUR[label] / 60000.0
    m_n = NR_MEAN[label]
    mu_n = m_n * t
    var_n = mu_n * (1 - mu_n) / t ** 2          # Bernoulli count -> Peaks Per Minute variance
    m_r = (mean - Q_NR * m_n) / (1 - Q_NR)
    ex2_r = ((mean ** 2 + sd ** 2) - Q_NR * (var_n + m_n ** 2)) / (1 - Q_NR)
    var_r = ex2_r - m_r ** 2
    var_lat = var_r - ROUND_VAR / t ** 2
    return m_r, math.sqrt(max(var_lat, 0.05))


LATENT = {lab: latent_params(lab) for lab in LABELS}


def stoch_round(rng, mu):
    base = math.floor(mu)
    return int(base + (1 if rng.random() < mu - base else 0))


def draw_amps(rng, label, n, resp_factor, nonresp, spurious=False):
    out = []
    for _ in range(n):
        if spurious:
            out.append(round(rng.uniform(0.04, 0.11), 4))
            continue
        m = AMP_MEAN[label] * (0.70 if nonresp else 1.0)
        v = m * math.exp(resp_factor - AMP_RESP_SD ** 2 / 2 + AMP_PEAK_SD * rng.gauss(0, 1) - AMP_PEAK_SD ** 2 / 2)
        out.append(round(clip(v, AMP_LO, AMP_HI), 4))
    return out


# ---------------------------------------------------------------------------
# the generator
# ---------------------------------------------------------------------------
def generate(seed, n_resp=N_RESP, detail=False):
    rng = random.Random(seed)
    big = n_resp != N_RESP                     # population runs for calibration skip the named special cases
    ids = [f"P{i + 1:02d}" for i in range(n_resp)]
    n_pairs = n_resp // 2

    # ---- respondents: pairs, group, seat, day -----------------------------
    group, seat, day_i = {}, {}, {}
    for p in range(n_pairs):
        a_first = rng.random() < 0.5
        seat1_first = rng.random() < 0.5
        a, b = ids[2 * p], ids[2 * p + 1]
        group[a], group[b] = ("A", "B") if a_first else ("B", "A")
        seat[a], seat[b] = (1, 2) if seat1_first else (2, 1)
        day_i[a] = day_i[b] = min(p // 3, 4)
    sexes = ["F"] * round(n_resp * 17 / 30) + ["M"] * (n_resp - round(n_resp * 17 / 30))
    rng.shuffle(sexes)
    sex = dict(zip(ids, sexes))
    ages = [int(round(clip(rng.gauss(27, 7), 19, 45))) for _ in range(n_resp)]
    if not big:
        ages[ages.index(min(ages))] = 19      # the range 19 to 45 holds exactly
        ages[ages.index(max(ages))] = 45
    age = dict(zip(ids, ages))

    pool = [i for i in ids if i not in (P11, P24)] if not big else list(ids)
    n_nr = N_NONRESP if not big else round(Q_NR * n_resp)
    nonresp = set(rng.sample(pool, n_nr))
    nr_sorted = sorted(nonresp)
    others = [i for i in pool if i not in nonresp]
    dry = {nr_sorted[0], rng.choice(others)}
    gap_resp = rng.sample(others, 3) if not big else []
    gap_plan = {}
    if gap_resp:
        gap_plan[gap_resp[0]] = ("Baseline_elf", round(rng.uniform(4.6, 5.6), 1))
        gap_plan[gap_resp[1]] = ("Baseline_Nike", round(rng.uniform(4.6, 5.6), 1))
        gap_plan[gap_resp[2]] = ("Nike_Commercial", round(rng.uniform(4.2, 4.8), 1))
    if not big:
        gap_plan[P24] = ("elf_Commercial", 11.4)

    resp = {}
    for rid in ids:
        resp[rid] = dict(id=rid, group=group[rid], seat=seat[rid], device=DEVICE[seat[rid]],
                         sex=sex[rid], age=age[rid], day=day_i[rid], nonresp=rid in nonresp, dry=rid in dry,
                         z=rng.gauss(0, 1), u_snr=rng.gauss(0, 1), u_amp=rng.gauss(0, 1),
                         hr_rest=clip(rng.gauss(74, 9), 52, 100),
                         u_style=rng.gauss(0, 1), w_like=rng.gauss(0, 1))

    # ---- stimulus rows ------------------------------------------------------
    metrics, quality, hr_rows, tidy = [], [], [], {}
    peaks_plan = []                                   # what the detail pass needs
    for rid in ids:
        r = resp[rid]
        order = ["Nike", "elf"] if r["group"] == "A" else ["elf", "Nike"]
        lead = rng.randint(*LEAD_MS)
        t_cursor = lead
        starts = {}
        for seq, brand in enumerate(order):
            for lab in (f"Baseline_{brand}", f"{brand}_Commercial", f"Recovery_{brand}"):
                starts[lab] = t_cursor
                t_cursor += DUR[lab]
            t_cursor += rng.randint(*QUESTIONNAIRE_MS)      # questionnaire slide
            if seq == 0:
                t_cursor += REST_MS                         # 2 minute rest
        r["starts"] = starts
        seq_noise = {brand: rng.gauss(0, HR_SEQ_SD) for brand in ("Nike", "elf")}
        for lab in LABELS:
            brand, phase, dur_ms = BRAND[lab], PHASE[lab], DUR[lab]
            first = order[0] == brand
            # --- signal quality: gap, missing, SNR
            gap_s = 0.0
            if rid in gap_plan and gap_plan[rid][0] == lab:
                gap_s = gap_plan[rid][1]
            if rid == P24 and lab == "elf_Commercial" and not big:
                missing_pct = P24_MISSING_PCT
            elif gap_s > 0:
                missing_pct = round(100 * gap_s / (dur_ms / 1000), 1)
            else:
                u = rng.random()
                missing_pct = 0.0 if u < 0.68 else round(rng.uniform(0.1, 1.5), 1) if u < 0.93 else round(rng.uniform(1.5, 4.0), 1)
            snr = (SNR_MEAN + (SEAT2_DB if r["seat"] == 2 else 0.0) + (DRY_DB if r["dry"] else 0.0)
                   + SNR_RESP_SD * r["u_snr"] + SNR_ROW_SD * rng.gauss(0, 1) - (3.0 if gap_s > 0 else 0.0))
            snr = round(clip(snr, 9.0, 34.0), 1)
            if rid == P11 and not big and lab in P11_SNR:
                snr = P11_SNR[lab]
            long_gap_ms = 0
            if gap_s > 4.0 or (rid == P24 and lab == "elf_Commercial" and not big):
                long_gap_ms = int(round((P24_MISSING_PCT / 100 * dur_ms) if (rid == P24 and lab == "elf_Commercial" and not big) else gap_s * 1000))
            sig_ms = dur_ms - long_gap_ms
            t_min = sig_ms / 60000.0
            # --- peak count
            plan_mean, _ = PPM_PLAN[lab]
            if r["nonresp"]:
                mu = NR_MEAN[lab] * t_min
                lam = None
            else:
                m_r, sd_lat = LATENT[lab]
                adj = 0.0
                if phase == "Commercial":
                    adj = POS_ADJ if first else -POS_ADJ
                lam = m_r + adj + sd_lat * (math.sqrt(RHO_LAT) * r["z"] + math.sqrt(1 - RHO_LAT) * rng.gauss(0, 1))
                lam = max(lam, 0.4)
                mu = lam * t_min
            cap = int((sig_ms / 1000.0 - 3.0) / 2.2)
            true_count = min(stoch_round(rng, mu), cap)
            spurious = (rid == P11 and not big and lab in P11_SPURIOUS)
            count = P11_SPURIOUS[lab] if spurious else true_count
            amps = draw_amps(rng, lab, count, AMP_RESP_SD * r["u_amp"], r["nonresp"], spurious=spurious)
            has = 1 if count > 0 else 0
            ppm = count / t_min
            avg_amp = round(sum(amps) / len(amps), 4) if amps else None
            # --- heart rate
            hr = r["hr_rest"] + HR_OFFSET[lab] + seq_noise[brand] + rng.gauss(0, HR_NOISE[phase])
            hr = round(hr, 1)
            # --- rows
            metrics.append({
                "Study Name": STUDY_NAME, "Respondent Name": rid, "Gender": r["sex"], "Age": r["age"],
                "Group": r["group"], "Type": "Stimulus", "Label": lab, "Start (ms)": starts[lab],
                "Duration (ms)": dur_ms, "Signal Duration (ms)": sig_ms, "Parent Stimulus": lab,
                "Has Peaks": has, "Peak Count": count, "Peaks Per Minute": fmt(ppm, 3),
                "Average Peak Amplitude (microSiemens)": fmt(avg_amp, 4)})
            quality.append({
                "Study Name": STUDY_NAME, "Respondent Name": rid, "Group": r["group"], "Device": r["device"],
                "Flow Name": FLOW[r["seat"]], "Stimulus": lab, "Signal-to-Noise ratio (dB)": fmt(snr, 1),
                "Signal missing (%)": fmt(missing_pct, 1)})
            hr_rows.append({
                "Study Name": STUDY_NAME, "Respondent Name": rid, "Group": r["group"], "Device": r["device"],
                "Stimulus": lab, "Start (ms)": starts[lab], "Duration (ms)": dur_ms,
                "Mean PPG Heart Rate (BPM)": fmt(hr, 1)})
            tidy[(rid, lab)] = dict(id=rid, lab=lab, group=r["group"], count=count, true_count=true_count,
                                    sig_ms=sig_ms, ppm=ppm, has=has, amp=avg_amp, amps=amps, hr=hr, snr=snr,
                                    missing=missing_pct, nonresp=r["nonresp"], long_gap_ms=long_gap_ms)
            peaks_plan.append((rid, lab, count, amps, starts[lab], dur_ms, gap_s, spurious,
                               long_gap_ms, rid == P24 and lab == "elf_Commercial"))

    # ---- survey (A10) ------------------------------------------------------
    survey = []
    for rid in ids:
        r = resp[rid]
        row = {"Study Name": STUDY_NAME, "Respondent Name": rid, "Group": r["group"]}
        for brand, a_mean, a_sd, l_mean, l_sd in (("Nike", 5.72, 1.6, 5.3, 1.1), ("elf", 5.04, 1.6, 5.0, 1.2)):
            t = tidy[(rid, f"{brand}_Commercial")]
            plan_m, plan_s = PPM_PLAN[f"{brand}_Commercial"]
            true_ppm = t["true_count"] / (t["sig_ms"] / 60000.0)
            z_ppm = (true_ppm - plan_m) / plan_s if not r["nonresp"] else rng.gauss(0, 1)
            w = math.sqrt(0.4) * r["u_style"] + math.sqrt(0.6) * rng.gauss(0, 1)
            a = a_mean + a_sd * (SURVEY_R[brand] * z_ppm + math.sqrt(1 - SURVEY_R[brand] ** 2) * w)
            row[f"Arousal_{brand}_1to9"] = int(clip(round(a), 1, 9))
        zl_n = r["w_like"]
        zl_e = 0.4 * zl_n + math.sqrt(1 - 0.4 ** 2) * rng.gauss(0, 1)
        row["Liking_Nike_1to7"] = int(clip(round(5.3 + 1.1 * zl_n), 1, 7))
        row["Liking_elf_1to7"] = int(clip(round(5.0 + 1.2 * zl_e), 1, 7))
        survey.append(row)

    out = dict(seed=seed, resp=resp, ids=ids, metrics=metrics, quality=quality, hr=hr_rows, tidy=tidy,
               survey=survey, gap_plan=gap_plan, nonresp=nr_sorted, dry=sorted(dry))

    # ---- session log (A11) -------------------------------------------------
    log = []
    for rid in ids:
        r = resp[rid]
        rows_q = [tidy[(rid, lab)] for lab in LABELS]
        medial = rng.random() < 0.18
        phal = "medial" if medial else "distal"
        placement = f"Index and middle fingers, {phal} phalanges, non-dominant hand; PPG probe on ring finger"
        if r["dry"]:
            skin = "Dry, hand cream applied"
        elif medial:
            skin = "Callused fingertips"
        else:
            skin = rng.choice(["Normal", "Normal", "Normal", "Slightly moist", "Normal"])
        warm = rng.choice([4, 5, 5, 5, 5, 5, 6, 6, 8, 5])
        complete, note = "Yes", ""
        if not big and rid == P11:
            warm = 2
            order_txt = "Nike was the second sequence" if r["group"] == "B" else "the e.l.f. sequence that followed was clean after the lead was re-seated"
            note = (f"EXCLUDE: SNR {P11_SNR['Nike_Commercial']:.1f} dB in Nike_Commercial is below 0 dB, the rule written before "
                    f"collection; electrode lead loosened during the commercial (warm up cut to 2 minutes); {order_txt}; "
                    "recorded in full anyway because the session was scheduled")
        elif not big and rid == P24:
            complete = "No"
            note = (f"EXCLUDE: Signal missing {P24_MISSING_PCT:.1f} percent in elf_Commercial is above 20 percent, the rule written "
                    "before collection; Bluetooth dropout of 11.4 s; GSR and PPG both lost for the gap")
        elif rid in gap_plan:
            lab, gs = gap_plan[rid]
            mp = tidy[(rid, lab)]["missing"]
            note = (f"Keep: one Bluetooth gap of {gs:.1f} s in {lab} (Signal missing {mp:.1f} percent, below the 20 percent rule); "
                    "Signal Duration is shorter than Duration for that stimulus")
        elif r["nonresp"]:
            nz = sum(1 for t in rows_q if t["has"] == 0)
            note = (f"Keep: Has Peaks = 0 in {nz} of 6 stimuli (nonresponder) but signal quality is good; analyzed as recorded "
                    "and checked in the sensitivity analysis")
        elif r["dry"]:
            lo_s = min(t["snr"] for t in rows_q)
            hi_s = max(t["snr"] for t in rows_q)
            note = f"Keep: dry skin, SNR {lo_s:.1f} to {hi_s:.1f} dB needs a look at the signal but no commercial is below 0 dB"
        log.append({"Respondent Name": rid, "Session Date": DAYS[r["day"]].isoformat(), "Operator": OPERATORS[r["day"]],
                    "Seat": r["seat"], "Device": r["device"], "Electrode placement (fingers)": placement,
                    "Skin note": skin, "Warm up minutes": warm, "Data complete": complete, "Decision note": note})
    out["log"] = log
    if detail:
        out["peaks"] = make_peaks(seed, resp, peaks_plan)
    return out


SURVEY_R = {"Nike": 0.56, "elf": 0.49}   # A10: weight of the standardized Peaks Per Minute in self reported arousal for responders (calibrated so the full sample r is about .35 for each commercial)


# ---------------------------------------------------------------------------
# peak times (detail pass; a separate random stream per respondent and stimulus, so the realized
# statistics used by the scan never depend on it)
# ---------------------------------------------------------------------------
def profile(label, d_s):
    n = int(round(d_s * 10))
    ts = [(k + 0.5) / 10.0 for k in range(n)]
    ph = PHASE[label]
    w = []
    for t in ts:
        if ph == "Baseline":
            v = 1 + 0.6 * math.exp(-t / 5.0)
        elif ph == "Recovery":
            v = 0.3 + 1.7 * math.exp(-t / 6.0)
        else:
            v = 1 + 1.4 * math.exp(-t / 4.0)
            for frac, s in KEY_MOMENTS.get(label, []):
                v += s * math.exp(-0.5 * ((t - frac * d_s) / (0.045 * d_s)) ** 2)
        w.append(v)
    cum, c = [], 0.0
    for v in w:
        c += v
        cum.append(c)
    return ts, cum


def make_peaks(seed, resp, plan):
    rows = []
    prof_cache = {}
    for rid, lab, count, amps, start_ms, dur_ms, gap_s, spurious, long_gap_ms, is_p24 in plan:
        if count == 0:
            continue
        r = random.Random(f"{seed}|{rid}|{lab}")
        d_s = dur_ms / 1000.0
        if lab not in prof_cache:
            prof_cache[lab] = profile(lab, d_s)
        ts, cum = prof_cache[lab]
        last_bin_end = bins_for(d_s)[-1][1]
        gap = None
        if long_gap_ms:
            g0 = 8.0 if is_p24 else r.uniform(6.0, d_s - 6.0 - long_gap_ms / 1000.0)
            gap = (g0, g0 + long_gap_ms / 1000.0)
        chosen = []
        sep = 2.0
        for k in range(count):
            for attempt in range(600):
                if attempt == 300:
                    sep = 1.2
                rise, decay = r.uniform(1.0, 2.8), r.uniform(1.5, 5.5)
                if spurious:
                    apex = r.uniform(rise + 0.5, last_bin_end - 0.1)
                else:
                    x = r.random() * cum[-1]
                    lo, hi = 0, len(cum) - 1
                    while lo < hi:
                        mid = (lo + hi) // 2
                        if cum[mid] < x:
                            lo = mid + 1
                        else:
                            hi = mid
                    apex = ts[lo]
                on, off = apex - rise, min(apex + decay, d_s)
                if on < 0.5 or apex > last_bin_end - 0.05 or apex + 0.2 > off:
                    continue
                if any(abs(apex - a) < sep for a, _, _ in chosen):
                    continue
                if gap and not (off < gap[0] or on > gap[1]):
                    continue
                chosen.append((apex, on, off))
                break
            else:
                raise RuntimeError(f"could not place peak {k + 1} of {count} for {rid} {lab}")
        chosen.sort()
        a_list = amps[:]
        r.shuffle(a_list)
        for (apex, on, off), amp in zip(chosen, a_list):
            rows.append((rid, lab, start_ms, dur_ms, start_ms + int(round(on * 1000)), start_ms + int(round(apex * 1000)),
                         start_ms + int(round(off * 1000)), amp, first_bin_label(apex, d_s)))
    return rows


# ---------------------------------------------------------------------------
# realized statistics for the kept respondents (used by the scan and the console summary)
# ---------------------------------------------------------------------------
def excluded_by_rule(data):
    bad = set()
    for (rid, lab), t in data["tidy"].items():
        if PHASE[lab] == "Commercial" and (t["snr"] < EXCLUDE_SNR_DB or t["missing"] > EXCLUDE_MISSING_PCT):
            bad.add(rid)
    return bad


def flagged_nonresponders(data):
    out = set()
    for rid in data["ids"]:
        zeros = sum(1 for lab in LABELS if data["tidy"][(rid, lab)]["has"] == 0)
        if zeros >= NR_RULE_MIN_ZERO:
            out.add(rid)
    return out


def realized(data):
    bad = excluded_by_rule(data)
    kept = [i for i in data["ids"] if i not in bad]
    T = data["tidy"]
    out = {}

    def col(lab, f="ppm"):
        return [T[(i, lab)][f] for i in kept]

    for lab in LABELS:
        out[f"m_{lab}"] = st.mean(col(lab))
        out[f"sd_{lab}"] = st.stdev(col(lab))
    pairs = [("Baseline_Nike", "Nike_Commercial"), ("Baseline_elf", "elf_Commercial"), ("Nike_Commercial", "elf_Commercial")]
    out["corr_mean"] = st.mean(pearson(col(a), col(b)) for a, b in pairs)

    def dz(x, y):
        d = [b - a for a, b in zip(x, y)]
        return st.mean(d) / st.stdev(d)
    out["dz_h1_nike"] = dz(col("Baseline_Nike"), col("Nike_Commercial"))
    out["dz_h1_elf"] = dz(col("Baseline_elf"), col("elf_Commercial"))
    out["dz_h2"] = dz(col("elf_Commercial"), col("Nike_Commercial"))
    grp = {i: data["resp"][i]["group"] for i in kept}
    d2 = {i: T[(i, "Nike_Commercial")]["ppm"] - T[(i, "elf_Commercial")]["ppm"] for i in kept}
    dA = [d2[i] for i in kept if grp[i] == "A"]
    dB = [d2[i] for i in kept if grp[i] == "B"]
    out["order_effect"] = (st.mean(dB) - st.mean(dA)) / 2
    for key, labs in (("amp_base", ("Baseline_Nike", "Baseline_elf")), ("amp_nike", ("Nike_Commercial",)),
                      ("amp_elf", ("elf_Commercial",))):
        v = [T[(i, lab)]["amp"] for i in kept for lab in labs if T[(i, lab)]["amp"] is not None]
        out[key] = st.mean(v)
    hb = [T[(i, "Baseline_Nike")]["hr"] for i in kept] + [T[(i, "Baseline_elf")]["hr"] for i in kept]
    out["hr_base_mean"], out["hr_base_sd"] = st.mean(hb), st.stdev(hb)
    out["hr_chg_nike"] = st.mean(T[(i, "Nike_Commercial")]["hr"] - T[(i, "Baseline_Nike")]["hr"] for i in kept)
    out["hr_chg_elf"] = st.mean(T[(i, "elf_Commercial")]["hr"] - T[(i, "Baseline_elf")]["hr"] for i in kept)
    out["hr_chg_diff"] = out["hr_chg_nike"] - out["hr_chg_elf"]
    sv = {s["Respondent Name"]: s for s in data["survey"]}
    for b in ("Nike", "elf"):
        a = [sv[i][f"Arousal_{b}_1to9"] for i in kept]
        out[f"ar_mean_{b}"], out[f"ar_sd_{b}"] = st.mean(a), st.stdev(a)
        out[f"ar_r_{b}"] = pearson(a, col(f"{b}_Commercial"))
        l_ = [sv[i][f"Liking_{b}_1to7"] for i in kept]
        out[f"lk_mean_{b}"] = st.mean(l_)
    rows = [T[(i, lab)]["snr"] for i in kept for lab in LABELS]
    out["snr_in_18_32"] = sum(1 for v in rows if 18 <= v <= 32) / len(rows)
    out["n_kept"] = len(kept)
    out["nr_ok"] = 1.0 if (flagged_nonresponders(data) == set(data["nonresp"]) and len(data["nonresp"]) == N_NONRESP) else 0.0
    out["excl_ok"] = 1.0 if bad == {P11, P24} else 0.0
    return out


# target value and tolerance (about one standard error of the statistic in a sample of 28; the
# statistics that carry the three hypotheses get a tighter tolerance). The planned paired effect
# sizes are the values the planned means, SDs and correlation of .5 imply:
#   Nike minus Baseline_Nike   2.2 / sqrt(1.3^2 + 2.0^2 - 2 x .5 x 1.3 x 2.0) = 2.2 / 1.76 = 1.25
#   elf minus Baseline_elf     1.3 / sqrt(1.3^2 + 1.8^2 - 2 x .5 x 1.3 x 1.8) = 1.3 / 1.61 = 0.81
#   Nike minus elf             0.9 / sqrt(2.0^2 + 1.8^2 - 2 x .5 x 2.0 x 1.8) = 0.9 / 1.91 = 0.47
TARGETS = {
    "m_Baseline_Nike": (2.4, 0.20), "m_Baseline_elf": (2.4, 0.20), "m_Nike_Commercial": (4.6, 0.25),
    "m_elf_Commercial": (3.7, 0.22), "m_Recovery_Nike": (3.0, 0.45), "m_Recovery_elf": (3.0, 0.45),
    "sd_Baseline_Nike": (1.3, 0.20), "sd_Baseline_elf": (1.3, 0.20), "sd_Nike_Commercial": (2.0, 0.28),
    "sd_elf_Commercial": (1.8, 0.28),
    "corr_mean": (0.5, 0.06), "dz_h1_nike": (1.25, 0.18), "dz_h1_elf": (0.81, 0.15), "dz_h2": (0.47, 0.10),
    "order_effect": (-0.8, 0.25),
    "amp_base": (0.20, 0.025), "amp_nike": (0.28, 0.025), "amp_elf": (0.25, 0.025),
    "hr_base_mean": (74.0, 2.0), "hr_base_sd": (9.0, 1.1), "hr_chg_nike": (-1.5, 0.25), "hr_chg_elf": (-0.8, 0.25),
    "hr_chg_diff": (-0.7, 0.25),
    "ar_mean_Nike": (5.9, 0.30), "ar_mean_elf": (5.2, 0.30), "ar_r_Nike": (0.35, 0.06), "ar_r_elf": (0.35, 0.06),
    "lk_mean_Nike": (5.3, 0.25), "lk_mean_elf": (5.0, 0.25), "snr_in_18_32": (0.93, 0.04),
}


def loss(stats):
    base = sum(((stats[k] - tgt) / tol) ** 2 for k, (tgt, tol) in TARGETS.items())
    base += 1000.0 * (1.0 - stats["nr_ok"]) + 1000.0 * (1.0 - stats["excl_ok"])
    return base


# ---------------------------------------------------------------------------
# writing
# ---------------------------------------------------------------------------
def banner_text(seed, extra):
    return (f"# SIMULATED DATA for teaching. Generated by make_simulated_data.py (seed {seed}). "
            f"No real person was recorded. {extra}")


B_METRICS = ("Columns follow the GSR Peak Metrics export (Export R Processed GSR Files article); Type and Parent Stimulus "
             "values are modeled. The Shimmer units are Shimmer3-8956 (Seat 1) and Shimmer3-B (Seat 2, placeholder name, verify).")
B_QUALITY = ("Signal-to-Noise ratio (dB) and Signal missing (%) are the two measures of the GSR Signal Quality export; the other "
             "columns are modeled. Device Shimmer3-B (Seat 2) is a placeholder name, verify on the lab machine.")
B_PEAKS = ("Columns follow the GSR Peak Table (Export R Processed GSR Files article); M2M Bin is modeled. Device Shimmer3-B "
           "(Seat 2) is a placeholder name, verify on the lab machine.")
B_HR = ("The column Mean PPG Heart Rate (BPM) is a MODELED name: the manuals name no PPG heart rate notebook or column. It stands for "
        "the beats per minute channel of the Shimmer3 PPG averaged per stimulus from a Sensor Data export. VERIFY the real "
        "column name, calculation, and units on the lab machine.")
B_SURVEY = "Self assessment manikin style arousal (1 to 9) and liking (1 to 7) after each sequence."
B_LOG = ("Operator log. Seat 1 uses Shimmer3-8956; Seat 2 uses Shimmer3-B (placeholder name, verify). The 5 minute warm up is a "
         "lab rule because the manuals state no warm up time.")


def write_csv(path_name, columns, rows, banner, extra_comments=()):
    path = os.path.join(HERE, path_name)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        fh.write(banner + "\n")
        for c in extra_comments:
            fh.write(c + "\n")
        w = csv.DictWriter(fh, fieldnames=columns, lineterminator="\n")
        w.writeheader()
        for row in rows:
            w.writerow(row)
    print(f"wrote {path} ({len(rows)} rows)")


def peak_table_rows(data):
    resp = data["resp"]
    rows, idx_all = [], {}
    by_resp = {}
    for p in data["peaks"]:
        by_resp.setdefault(p[0], []).append(p)
    out = []
    for rid in data["ids"]:
        plist = sorted(by_resp.get(rid, []), key=lambda p: p[5])
        within = {}
        r = resp[rid]
        for k, (rid_, lab, st_ms, d_ms, on, ap, off, amp, binlab) in enumerate(plist, start=1):
            within[lab] = within.get(lab, 0) + 1
            out.append({
                "Study Name": STUDY_NAME, "Respondent Name": rid, "Respondent Gender": r["sex"],
                "Respondent Age": r["age"], "Respondent Group": r["group"], "Stimulus Label": lab,
                "Stimulus Start": st_ms, "Stimulus Duration": d_ms, "Peak Index": k,
                "Peak Index by Stimulus": within[lab], "Peak Onset Time": on, "Peak Time": ap, "Peak Offset Time": off,
                "Peak Amplitude": f"{amp:.4f}", "Device": r["device"], "Flow Name": FLOW[r["seat"]], "M2M Bin": binlab})
    return out


def main():
    args = sys.argv[1:]
    if args and args[0] == "--scan":
        n = int(args[1]) if len(args) > 1 else 200
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
    if args and args[0] == "--calibrate":
        calibrate()
        return
    seed = int(args[0]) if args else SEED
    data = generate(seed, detail=True)
    banner = banner_text(seed, "")
    write_csv("SIMULATED_gsr_peak_metrics.csv", list(data["metrics"][0].keys()), data["metrics"],
              banner_text(seed, B_METRICS))
    write_csv("SIMULATED_gsr_signal_quality.csv", list(data["quality"][0].keys()), data["quality"], banner_text(seed, B_QUALITY))
    prow = peak_table_rows(data)
    meta = ("# METADATA (modeled): Units = microSiemens; Devices = Shimmer3-8956 (Seat 1), Shimmer3-B (Seat 2, placeholder name); "
            "Algorithm = GSR Peak Detection (parameter values are not stated in the manuals); peak times in ms from the start of the recording")
    write_csv("SIMULATED_gsr_peak_table.csv", list(prow[0].keys()), prow, banner_text(seed, B_PEAKS), extra_comments=[meta])
    write_csv("SIMULATED_heart_rate.csv", list(data["hr"][0].keys()), data["hr"], banner_text(seed, B_HR))
    write_csv("SIMULATED_gsr_survey.csv", list(data["survey"][0].keys()), data["survey"], banner_text(seed, B_SURVEY))
    write_csv("SIMULATED_gsr_session_log.csv", list(data["log"][0].keys()), data["log"], banner_text(seed, B_LOG))
    stats = realized(data)
    print(f"\nRealized statistics for the {stats['n_kept']} kept respondents (target, tolerance) and loss {loss(stats):.2f}:")
    for k, (tgt, tol) in TARGETS.items():
        print(f"  {k:<20} {stats[k]:9.3f}   target {tgt}")
    print(f"  nonresponders flagged by the rule: {sorted(flagged_nonresponders(data))}; excluded by the rule: {sorted(excluded_by_rule(data))}")


def calibrate():
    """Population check of the model: a large simulated sample (no exclusions, no named special cases)."""
    d = generate(1, n_resp=6000)
    T, ids = d["tidy"], d["ids"]

    def col(lab):
        return [T[(i, lab)]["ppm"] for i in ids]
    print("population check, n = 6000 (all respondents, nonresponders included)")
    for lab in LABELS:
        print(f"  {lab:<16} mean {st.mean(col(lab)):5.2f}  SD {st.stdev(col(lab)):5.2f}   plan {PPM_PLAN[lab]}")
    pairs = [("Baseline_Nike", "Nike_Commercial"), ("Baseline_elf", "elf_Commercial"), ("Nike_Commercial", "elf_Commercial"),
             ("Baseline_Nike", "Baseline_elf"), ("Baseline_Nike", "elf_Commercial"), ("Baseline_elf", "Nike_Commercial")]
    cs = [pearson(col(a), col(b)) for a, b in pairs]
    print("  correlations:", ", ".join(f"{a[:6]}/{b[:6]} {c:.2f}" for (a, b), c in zip(pairs, cs)), f"  mean {st.mean(cs[:3]):.2f}")
    for b in ("Nike", "elf"):
        sv = d["survey"]
        a = [s[f"Arousal_{b}_1to9"] for s in sv]
        print(f"  arousal {b}: mean {st.mean(a):.2f} SD {st.stdev(a):.2f}  r with PPM {pearson(a, col(f'{b}_Commercial')):.3f}")
    hb = [T[(i, 'Baseline_Nike')]['hr'] for i in ids]
    print(f"  HR baseline mean {st.mean(hb):.1f} SD {st.stdev(hb):.1f}; change Nike {st.mean(T[(i,'Nike_Commercial')]['hr']-T[(i,'Baseline_Nike')]['hr'] for i in ids):.2f},"
          f" elf {st.mean(T[(i,'elf_Commercial')]['hr']-T[(i,'Baseline_elf')]['hr'] for i in ids):.2f}")
    for lab in LABELS:
        v = [T[(i, lab)]['amp'] for i in ids if T[(i, lab)]['amp'] is not None]
        print(f"  amplitude {lab:<16} {st.mean(v):.3f}")
    sn = [T[(i, lab)]['snr'] for i in ids for lab in LABELS]
    print(f"  SNR mean {st.mean(sn):.1f} SD {st.stdev(sn):.1f}, share in 18 to 32: {sum(18 <= v <= 32 for v in sn) / len(sn):.3f}")
    print("  latent parameters (responders): " + "; ".join(f"{k} {v[0]:.2f}/{v[1]:.2f}" for k, v in LATENT.items()))


if __name__ == "__main__":
    main()
