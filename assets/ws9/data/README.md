# Workshop 9 data: planning the follow up study (power, counterbalancing, sessions, consent, and the analysis plan)

**PLANNED STUDY, SIMULATED PLANNING DATA.** Nothing in this folder comes from a real participant. Every power value was computed from simulated data sets (`plan_power.R`, seed 20270208), every planning value comes from the SIMULATED Multimodal Ad Study of Workshops 7 and 8 (seed 20271231) or is labeled as a judgment, and every date is a placeholder. ALTAVO and BRENLOW are invented brands, the client is fictional, and participants and staff carry codes only (P01, R01, X01; S1 to S6). The consent form is a TEACHING MODEL that must be replaced by the current JSU IRB template and approved before any use. No number here may be quoted as a finding.

Martin's Family Clothing Consumer Behavior Research Laboratory, Jacksonville State University. Workshop 9 turns the series back to the start of a study: from a question to a comparison, a design, a sample size, a counterbalanced order, a session schedule, a consent form, an IRB submission, a pilot, and a preregistered analysis plan.

---

## 1. Continuity with Workshops 7 and 8

| Workshop | What it did | What Workshop 9 takes from it |
|---|---|---|
| 7 | Recorded the SIMULATED Multimodal Ad Study: 24 participants (22 analyzed), one 30 s ad for ALTAVO and one for BRENLOW, Aurora, AFFDEX, Shimmer3 GSR+; rules R1 to R3; session log | The variance of liking (participant SD 0.80, residual SD 1.09, fitted in STEP 0), the exclusion rate (2 of 24), the 90 minute session slots, the measures, the AOIs, rules R1 to R3 (adapted), the session log columns |
| 8 | Full analysis; no difference between the two ads; sensitivity power analysis; "What the next study needs" | The paired sample sizes (34 and 45 for d_z = 0.50; 90 and 120 for d_z = 0.30), reproduced exactly in STEP 1; the H1 planning values (recall .35 when the Logo is not fixated, odds ratio 4.34); the call for several ads per brand, purchase intention per ad, a named smallest effect, and a preregistration |
| 6 | Remote study with webcam eye tracking: 142 of 260 who opened the link were kept | The lesson that remote studies lose far more participants (STEP 4: the same 48 would need about 88 link openers) |

`plan_power.R` reads these files from `../../ws6/data`, `../../ws7/data`, and `../../ws8/data` when the folders exist, and otherwise from `prior_inputs/`, which holds byte identical copies (STEP 0 compares the MD5 checksums when both exist). A reader who downloaded only Workshop 9 can run everything.

---

## 2. The planned study and the planning result

**Design (PLANNED).** Within subjects lab study, one participant at a time at the single Smart Eye Aurora seat, with AFFDEX and one Shimmer3 GSR+. Every participant sees **12 ads (6 per brand)** in a Williams order, rates liking and purchase intention after each ad, and answers brand recall for each ad at the end. Ads are a random factor so that a brand difference can be separated from the executions. **48 participants** are analyzed.

**Why 12 ads and not 6.** Workshop 8 asked for "several ads per brand", and three per brand is the obvious first guess. The simulation shows that the number of ads, not the number of participants, limits the power of a brand claim: with 3 ads per brand, power for a 0.50 point brand difference in liking is .47 at 48 participants and .51 at 72; with 4 ads per brand it never reaches .80 in the grid; with 6 ads per brand it is .88 at 48 participants. The plan therefore asks the client for six ads per brand. This is the main lesson of the workshop: plan the stimuli, not only the people (Westfall et al., 2014).

| Result (all SIMULATED or PLANNED) | Value | Where |
|---|---|---|
| Paired comparison, d_z = 0.50: participants for 80% power | 34; 45 at the strictest Holm step of three | STEP 1 (power.t.test; identical to Workshop 8) |
| Paired comparison, d_z = 0.30 | 90; 120 with Holm | STEP 1 |
| Brand test on liking (P1), ads random, 6 per brand, 48 participants | power .880 (Monte Carlo SE .015); formula .866 | STEP 2, `SIMULATED_power_grid.csv` |
| Same with 3 per brand (six ads in all) | .474 at 48; .510 at 72 | STEP 2 |
| Type I error of the planned test at the chosen design | .052 (MC SE .010) | STEP 2, check C1 |
| Power if the ad SD were 0.10 / 0.30 instead of 0.20 | .976 / .640 | STEP 2, checks C2 and C3 |
| With a participant by brand slope (SD 0.50) not in the model | Type I error .060, power .886 | STEP 2, checks C4 and C5 |
| Chance of concluding equivalence (90% CI inside plus or minus 0.50) when the brands do not differ | 88.2% | STEP 2, check C1 |
| H1 (Logo fixation and recall), 48 participants, 12 ads each | power .975 for odds ratio 2.0; 1.000 for 4.34 | STEP 3, `SIMULATED_power_h1.csv` |
| Sessions to book | 53 (48 main + 5 reserve) + 2 open slots; cap 60 | STEP 4 |
| Counterbalancing | Williams design, 12 rows; 4 blocks of 12; brand pattern ABBA... or BAAB... | STEP 5 |
| Session length | 60 min for the participant, 90 min slots, 5 per weekday | STEP 6 |
| Schedule | pilot day Monday, February 1, 2027 (2 sessions); data collection on 11 weekdays, February 8 to 22, 2027 (placeholder dates) | STEP 6, `SIMULATED_session_schedule.csv` |

---

## 3. What is in this folder

| File or folder | What it is |
|---|---|
| `plan_power.R` | The planning script in STEP blocks 0 to 8 (setup and assumptions; analytic power; simulation of the brand test; simulation of H1; the chosen design; the Williams design; the session schedule; figures; tables). |
| `plan_output.txt` | The full console log of `plan_power.R` (written with `sink`). Every number in the documents below is copied from it. |
| `SIMULATED_planning_assumptions.csv` | Every planning value (A1 to A19) with its source file and line, or the label JUDGMENT or decision. |
| `SIMULATED_power_analytic.csv` | STEP 1: n and power of a paired t test for four d_z values, alpha .05 and .05/3. |
| `SIMULATED_power_grid.csv` | STEP 2: power by participants (24 to 72) and ads per brand (2, 3, 4, 6), with Monte Carlo SE and interval, power at the Holm step, the formula, the naive power, the share of singular fits, and the lmerTest verification columns. |
| `SIMULATED_power_checks.csv` | STEP 2, part B: checks C1 to C5 at the chosen design. |
| `SIMULATED_power_h1.csv` | STEP 3: power for H1 by participants and odds ratio. |
| `SIMULATED_latin_square.csv` | The Williams design for 12 ads (row, ad in each position, brand sequence). |
| `SIMULATED_order_schedule.csv` | One row per participant code (pilot X01 and X02, main P01 to P48, reserve R01 to R05): block, Latin square row, the ad in each of the 12 positions, the brand sequence, the first brand. |
| `SIMULATED_session_schedule.csv` | One row per session slot: date, weekday, arrival, recording start and end, participant leaves, room ready, participant code, Latin square row, lead operator, assistant, notes. Codes only. |
| `figures/` | `SIMULATED_fig1_power_curves.png` (power by participants and ads per brand), `SIMULATED_fig2_power_by_ads.png` (power at 48 participants by ads per brand and ad SD), `SIMULATED_fig3_power_h1.png`, `SIMULATED_fig4_williams_square.png` (all 1400 px wide, 150 dpi, Okabe and Ito colors, titled SIMULATED or PLANNED). |
| `analysis_plan_prereg.md` | The preregistration style analysis plan: hypotheses (P1; S1 to S3; R1 to R3), design, sampling plan and stopping rule, variables, models, inference criteria, exclusion rules written in advance, exploratory analyses. |
| `consent_form_model.md` | TEACHING MODEL consent form (eighth grade reading level; Flesch-Kincaid grade about 2.5 for the prose of the form, approximate syllable count), a model assent paragraph, a note on parental permission, and the source of every element. |
| `planning_checklist.md` | The study planning checklist from the question to the pilot (18 items), each with what "done" looks like, where to look, and the worked example. |
| `make_excel_workbook.py` | Builds `SIMULATED_planning.xlsx` and `excel_walkthrough.md` from one registry of cells and verifies them (openpyxl and LibreOffice). |
| `SIMULATED_planning.xlsx` | The Excel route: SampleSize, Attrition, LatinSquare, Assignment, Calendar, Check. |
| `excel_walkthrough.md` | Step by step Excel instructions with exact sheet names, cell addresses, formulas, and values. |
| `prior_inputs/` | Byte identical copies of the eight Workshop 6, 7, and 8 files that `plan_power.R` reads. |

---

## 4. How to run (from this folder)

```
Rscript plan_power.R              # about 9 minutes on one core; writes plan_output.txt, the CSV files, figures/
Rscript plan_power.R quick        # about 2.5 minutes: 100 simulated studies per cell (numbers differ slightly from plan_output.txt)
python3 make_excel_workbook.py    # about 5 seconds; writes and verifies the workbook and excel_walkthrough.md
```

Run `make_excel_workbook.py` after `plan_power.R`, because it reads the CSV files and `plan_output.txt`. It needs LibreOffice (`soffice`) on the path for the recalculation check.

**Runtime (last run).** 9.0 minutes on one core: STEP 2 part A (10,000 mixed model fits) 5.1 min, part B (2,500 fits) 1.3 min, STEP 3 (1,200 mixed logistic fits) 2.5 min. On macOS or Linux with several free cores, `N_CORES <- 2L` in STEP 0 runs two cells at a time; the results do not change, because every simulated study has its own seed (SEED + 10000 x cell + study). Windows always runs on one core.

**Simulation sizes.** 500 simulated studies per cell for the brand test grid (20 cells) and for each of the five checks; 200 per cell for H1 (6 cells), because a mixed logistic fit takes about four times as long and 500 would have pushed the run past 10 minutes. The Monte Carlo SE is printed for every power value.

**Software.** R 4.3.3 with lme4 1.1.35.1, lmerTest 3.1.3, and ggplot2 3.4.4 (and parallel, grid, and tools from base R); Python 3 with openpyxl 3.1.5; LibreOffice 24.2.7.2 (headless). **No package was installed for this workshop.** If a package is missing on a lab computer: `install.packages(c("lme4", "lmerTest", "ggplot2"))`.

**Excel check (last run).** openpyxl confirmed that 1,028 registered cells and 64 Check rows hold the registered formulas (2,084 formula cells). LibreOffice recalculated a copy: all 64 Check rows are OK (`Check!B69` = 0), no cell shows an error value, the 40 exact results (counts, sample sizes, dates, codes, rows, staff) agree with R exactly or at the printed decimals, and the 24 approximations agree within their tolerances: the corrected normal approximation of the sample size within 0.13 participants of `power.t.test` (and every rounded up n is identical), the approximate paired power within 0.0004, and the approximate power of the brand test within 0.011 of R's exact noncentral t formula.

---

## 5. Things that look like errors but are not

1. **Power curves that are not monotone.** With 500 simulated studies per cell, the Monte Carlo SE is about .011 to .022, so neighboring cells can be out of order (for example 4 ads per brand: .726 at 60 participants and .688 at 72). The dashed formula lines in figure 1 show the trend; 3 of 20 cells differ from the formula by more than 1.96 Monte Carlo SE (about 1 expected by chance), and the mean difference is +.002.
2. **The test is computed on the ad means.** In a complete, balanced design, the lmerTest Satterthwaite test of the brand effect is exactly the t test of the ALTAVO ad means against the BRENLOW ad means (df = ads - 2), whenever the ad variance is estimated above zero. The script uses that identity to save time and proves it: the first 20 studies of every cell (500 in all) were also analyzed with lmerTest; in the 416 fits with the ad variance clearly above zero, the t and df agree (largest differences 1.5e-03 and 3.8e-03 in the grid), the naive p of the 83 singular fits agrees, and all 500 decisions at .05 agree.
3. **Singular fits.** With few ads, the ad variance is often estimated at zero (up to 38 percent of fits with 2 ads per brand). lmerTest then reports several hundred df, and a test that accepted that would be too liberal (the "naive" column). The planned rule tests the brand effect on the ad means in that case, which keeps the Type I error at .052.
4. **H1 power of 1.000.** With the Workshop 8 odds ratio (4.34) and 12 observations per participant, every one of the 200 simulated studies was significant; the Monte Carlo SE is then reported as 0.
5. **A brand pattern of ABBA, not ABAB.** No 12 x 12 Williams square alternates the brands in every row; the labeling used (even symbols ALTAVO, odd symbols BRENLOW) is the only one of 462 that keeps every row to at most two ads of one brand in a row.
6. **The pilot participants (X01, X02) have rows 1 and 2, and the reserve codes have no row.** Pilot data are not analyzed; a reserve participant receives the row of the participant it replaces.

---

## 6. Items to verify

1. **The JSU IRB.** The current consent template and its required wording, the review level this study needs, the submission forms and system, required training, and the timelines (none were taken from JSU web pages for this folder; read them on JSU's own pages and cite the URLs, or ask the IRB office). The lab procedure manual requires approval or exemption before any participant is scheduled (section 2.1).
2. **The IRB decisions left in brackets** in `consent_form_model.md`: retention period, who may see the data, compensation and any course credit alternative, age of consent, the IRB contact text, and whether face video may be used in teaching with a separate written release.
3. **The ad SD.** The plan assumes that ads of one brand differ in mean liking with an SD of 0.20 points. No lab estimate exists, and the pilot cannot estimate it. If the client's ads differ more (0.30), power falls to about .64; the analysis plan names this risk.
4. **The client can supply six comparable ads per brand** with the same layout (for the AOIs). With fewer ads the brand claim is underpowered and the plan must be revised before the IRB submission.
5. **Session timing.** The sensor setup without the EEG cap (planned 10 minutes) and the recording (computed 15.7 minutes, planned 20) must be measured in the pilot.
6. **Implementing the order in iMotions Lab 11.1.** One study flow per Latin square row (Using the iMotions Study Builder) or a test plan CSV (Using a Test Plan for Study Design, which says respondents cannot be added after recording starts, so the reserve codes must be in the file from the beginning). The test plan article's examples come from older versions; try it on the lab computer. `SIMULATED_order_schedule.csv` has a banner line that a test plan file must not have.
7. **The dates.** Placeholder dates (pilot Monday, February 1, 2027; data collection February 8 to 22, 2027) must be set after IRB approval and checked against the JSU academic calendar; `WORKDAY` in the workbook skips weekends only.
8. **Everything in the Workshop 7 and 8 lists of items to verify** (export layouts, the fixation filter, the AFFDEX threshold, the GSR notebook parameters), because the analysis plan rests on them.
9. **Staffing.** Two trained people per session (codes S1 to S6) in half day shifts; who these are, and their CITI training, is decided by the lab.

## Sources cited in this folder

The full references are listed at the end of `analysis_plan_prereg.md`, `excel_walkthrough.md`, and `planning_checklist.md`. The key source for the main lesson is: Westfall, J., Kenny, D. A., & Judd, C. M. (2014). Statistical power and optimal design in experiments in which samples of participants respond to samples of stimuli. *Journal of Experimental Psychology: General, 143*(5), 2020-2045. https://doi.org/10.1037/xge0000014
