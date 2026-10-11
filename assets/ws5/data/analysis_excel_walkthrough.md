# EEG frontal alpha asymmetry analysis in Excel: step by step walkthrough

**SIMULATED DATA.** Every number in this walkthrough comes from the SIMULATED files that `make_simulated_data.py` generated (seed 20267088). No real person was recorded. The design copies the logic of the iMotions demo study "DEMO - SonyBravia" (a commercial with a key scene against the same commercial without it, a second commercial for everyone, neutral 1 minute nature videos as baselines, and the last 10 seconds of each commercial); nothing about the content of the Sony commercials was used.

This is the same analysis as `analysis.R`, done in Excel for students who do not use R. The finished workbook is `SIMULATED_eeg_analysis.xlsx`; the sheet names below refer to it. **Every cell address in this walkthrough is written as `Sheet!Cell` and was generated from the workbook itself.** `make_excel_workbook.py` writes the workbook and this file from one list of cells, reads the workbook back with openpyxl to confirm that every address holds the formula printed here, and recalculates a copy to confirm the results. A copy was recalculated with LibreOffice 24.2.7.2 420(Build:2) (headless): all 199 Check rows are OK, and the 215 results printed in this walkthrough agree with analysis.R at the printed decimals.

Companion files: `analysis.R` (R version), `analysis_output.txt` (its console output), `results_text.md` (model paragraphs), `figures/` (PNG figures).

---

## 0. What you are analyzing

**Design.** 48 respondents (P01 to P48) were randomized to two versions of Commercial 1. **Version A** saw `Commercial_1A` (60 s) with a key scene at 32 to 38 s; **Version B** saw `Commercial_1B` (54 s), the same commercial with the key scene cut. Everyone also saw `Commercial_2` (60 s). Half of each version saw Commercial 1 first. The flow was `Baseline_1` (a 60 s neutral nature video), the first commercial, a rating slide, `Baseline_2` (another 60 s nature video), the second commercial, and a rating slide. **Each commercial is compared with the baseline that immediately precedes it.** Two stimulus annotations mark the analysis windows: `KeyScene` (32 to 38 s of `Commercial_1A`) and `EndScene` (the last 10 s of each commercial).

**Device.** One Neuroelectrics Enobio 8 streamed through NIC2 and Lab Streaming Layer (iMotions device name `Enobio EEG`, 500 Hz). iMotions reports the Enobio channels as Ch1 to Ch8; the scalp site of each channel is set in NIC2. This workbook **assumes** the lab manual montage order (Ch1 Fz, Ch2 Cz, Ch3 Pz, Ch4 Oz, Ch5 F3, Ch6 F4, Ch7 P3, Ch8 P4) and must be checked against the NIC2 protocol before real data are analyzed.

**The metric.** The EEG Frontal Asymmetry R Notebook computes, in each 1 s Welch window, **FA = ln(alpha power at the right frontal electrode) - ln(alpha power at the left frontal electrode)**, here ln(F4) - ln(F3). Alpha power goes **down** when a region is more active, so a higher FA means relatively more **left** frontal activity, which the notebook article reads as approach motivation; a lower FA means relatively more right frontal activity (withdrawal). Approach is not the same as liking. FA scores are read **relative to each other** (here, relative to a baseline), never as absolute values.

**Files.**

| File | What it is | Columns used here | Used for |
|---|---|---|---|
| `SIMULATED_eeg_fa_metrics.csv` | EEG Frontal Asymmetry Metrics export (R Processing), notebook run with left Ch5 and right Ch6 | `Respondent Name`, `Group`, `Type`, `Label`, `Parent Stimulus`, `Mean Frontal Asymmetry Alpha` | H1 to H3 |
| `SIMULATED_eeg_fa_metrics_default_channels.csv` | The same export made with the notebook **default** channels (left Ch1, right Ch7) | the same | Teaching contrast |
| `SIMULATED_eeg_artifact_report.csv` | Percent of F3 and F4 samples removed by the artifact threshold, per respondent and stimulus (**modeled layout**) | `Respondent Name`, `Electrode (assumed)`, `Samples expected`, `Samples recorded`, `Percent removed (%)` | Exclusion rule |
| `SIMULATED_eeg_session_log.csv` | Operator log (**modeled layout**) | `NIC2 quality F3 before recording`, `NIC2 quality F4 before recording`, `Incidents` | Exclusion rule |
| `SIMULATED_eeg_survey.csv` | Questionnaire (**modeled layout**) | `Liking_Commercial_1_1to7`, `Liking_Commercial_2_1to7` | Exploratory correlation |

A sixth file, `SIMULATED_eeg_psd_metrics.csv` (EEG Power Spectral Density Metrics export), is used in R for the theta descriptives and a consistency check; the workbook does not need it.

**Hypotheses.**
- H1 (within subjects, Version A): FA during the `KeyScene` minus FA during the baseline before `Commercial_1A` is greater than zero.
- H2 (between subjects): the baseline corrected FA during the `EndScene` of Commercial 1 is higher for Version A than for Version B.
- H3 (within subjects, all kept respondents): the baseline corrected whole commercial FA differs between `Commercial_2` and Commercial 1.
- Exploratory: whole commercial FA change and liking. Sensitivity: H1 and H2 with a 15 percent threshold and without baseline correction. Teaching contrast: H1 with the notebook default channels.

**Decision rules fixed before collection.** Exclude a respondent when more than 25 percent of the F3 or F4 samples were removed in any baseline or commercial, or when the session log records a Red NIC2 quality status at F3 or F4 that could not be fixed. All tests are two tailed with alpha = .05 (the one tailed p values in the blocks are for discussion only).

---

## 1. Import the files and build the Data sheet

iMotions exports are comma separated with a dot decimal. Double clicking a file uses your regional defaults, which can put everything in one column (Opening Data Export Files in Excel article). Use the import route instead.

**Route A, Get Data (Excel 2016 and later):**
1. **Data > From Text/CSV**, select `SIMULATED_eeg_fa_metrics.csv`, click **Import**.
2. In the preview set **Delimiter: Comma** and **File Origin: 65001 Unicode (UTF-8)**.
3. Click **Transform Data**. The two FA files start with **two** comment lines (`# SIMULATED ...` and `# METADATA ...`); the other files start with one. In Power Query choose **Home > Remove Rows > Remove Top Rows** (2 or 1), then **Home > Use First Row as Headers**, then **Close & Load**. (A real iMotions export has no comment lines; the simulated files carry them so nobody mistakes them for real data. Keep the METADATA line somewhere: it records the notebook settings.)
4. **You should see** 12 columns with `Study Name` in the first column and 312 data rows: 7 rows for each Version A respondent (four stimuli, `KeyScene`, and two `EndScene` rows) and 6 for each Version B respondent.

**Route B, Text Import Wizard (any version):** **File > Open > Browse**, set the file type to **All Files**, open the .csv, choose **Delimited**, start at row 3 for the FA files (row 2 for the others), tick **Comma**, click **Advanced** and set the decimal symbol to a dot, then **Finish**.

**Common mistake.** The FA column can hold the text `NA` (a segment in which every 1 s window was removed by the artifact threshold). Leave it as `NA`; do not type 0. A zero FA would be a real value (equal left and right alpha).

In the workbook this is the **Data** sheet, header in row 1 and data in rows 2 to 313. Column letters: `B` = Respondent Name, `E` = Group, `F` = Type, `G` = Label, `J` = Parent Stimulus, `K` = Mean Frontal Asymmetry Alpha. The **Default** sheet has the same layout for the default channel export, and the **Artifact**, **Log**, and **Survey** sheets hold the other three files as imported.

---

## 2. Helper columns: Version, Order, Role, Keep, Keep15

Add five columns to the right of the export (`M` to `Q`, headers in orange) and fill them down. The formulas below are the ones in row 2.

| Header | Cell | Formula | Meaning |
|---|---|---|---|
| Version | `Data!M2` | `=LEFT(E2,1)` | A or B, the first letter of `Group` (for example `A_C1first`) |
| Order | `Data!N2` | `=MID(E2,3,10)` | `C1first` or `C2first` |
| Role | `Data!O2` | `=IF(F2="Stimulus",IF(LEFT(G2,8)="Baseline",G2,IF(LEFT(G2,12)="Commercial_1","C1","C2")),IF(G2="KeyScene","KeyScene",IF(LEFT(J2,12)="Commercial_1","EndScene_C1","EndScene_C2")))` | Which part of the design the row is: `Baseline_1`, `Baseline_2`, `C1`, `C2`, `KeyScene`, `EndScene_C1`, `EndScene_C2` |
| Keep | `Data!P2` | `=IF(COUNTIFS(Artifact!$B$2:$B$385,B2,Artifact!$K$2:$K$385,">25")+COUNTIFS(Log!$A$2:$A$49,B2,Log!$H$2:$H$49,"Red")+COUNTIFS(Log!$A$2:$A$49,B2,Log!$I$2:$I$49,"Red")>0,"No","Yes")` | The exclusion rule: No when any F3 or F4 row of this respondent in the artifact report is above 25 percent, or the log shows Red at F3 or F4 |
| Keep15 | `Data!Q2` | `=IF(COUNTIFS(Artifact!$B$2:$B$385,B2,Artifact!$K$2:$K$385,">15")+COUNTIFS(Log!$A$2:$A$49,B2,Log!$H$2:$H$49,"Red")+COUNTIFS(Log!$A$2:$A$49,B2,Log!$I$2:$I$49,"Red")>0,"No","Yes")` | The same rule at 15 percent (sensitivity analysis only) |

**Why Role needs Parent Stimulus.** The label `EndScene` appears on all three commercials, so a row is identified by `Label` **and** `Parent Stimulus`. The Role formula sends an `EndScene` row whose parent starts with `Commercial_1` to `EndScene_C1` and the one on `Commercial_2` to `EndScene_C2`, and it gives `Commercial_1A` and `Commercial_1B` the same role, `C1`, because they are the two versions of the same commercial.

**Common mistake.** Averaging every `EndScene` row of a respondent, or filtering on `Label` alone. That mixes the two commercials.

---

## 3. The exclusion rule: Screening sheet

The **Screening** sheet lists every respondent once and shows why each one is kept or excluded. Columns C and D are **array formulas** (`MAX(IF(...))`): in Excel 365 press **Enter**; in Excel 2019 or older press **Ctrl+Shift+Enter**.

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Version | `Screening!B5` | `=INDEX(Data!$M$2:$M$313,MATCH($A5,Data!$B$2:$B$313,0))` |
| Largest percent removed, F3 | `Screening!C5` | `=MAX(IF((Artifact!$B$2:$B$385=$A5)*(Artifact!$F$2:$F$385="F3"),Artifact!$K$2:$K$385))` |
| Largest percent removed, F4 | `Screening!D5` | `=MAX(IF((Artifact!$B$2:$B$385=$A5)*(Artifact!$F$2:$F$385="F4"),Artifact!$K$2:$K$385))` |
| Largest of the two | `Screening!E5` | `=MAX(C5:D5)` |
| Red at F3 or F4 (count) | `Screening!F5` | `=COUNTIFS(Log!$A$2:$A$49,$A5,Log!$H$2:$H$49,"Red")+COUNTIFS(Log!$A$2:$A$49,$A5,Log!$I$2:$I$49,"Red")` |
| Decision (25 percent rule) | `Screening!G5` | `=IF(OR(E5>25,F5>0),"Exclude","Keep")` |
| Decision (15 percent rule) | `Screening!H5` | `=IF(OR(E5>15,F5>0),"Exclude","Keep")` |

The counts below the table:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Respondents excluded (25 percent rule) | `Screening!E54` | `=COUNTIF(G5:G52,"Exclude")` | 3 |
| Respondents excluded (15 percent rule) | `Screening!E55` | `=COUNTIF(H5:H52,"Exclude")` | 7 |
| Kept, Version A | `Screening!E56` | `=COUNTIFS(B5:B52,"A",G5:G52,"Keep")` | 22 |
| Kept, Version B | `Screening!E57` | `=COUNTIFS(B5:B52,"B",G5:G52,"Keep")` | 23 |
| Kept, Version A (15 percent rule) | `Screening!E58` | `=COUNTIFS(B5:B52,"A",H5:H52,"Keep")` | 20 |
| Kept, Version B (15 percent rule) | `Screening!E59` | `=COUNTIFS(B5:B52,"B",H5:H52,"Keep")` | 21 |
| Artifact report rows with a stream gap | `Screening!E60` | `=COUNTIF(Artifact!$M$2:$M$385,">0")` | 2 |

**You should see** three exclusions, each with a cause in the session log (**Log** sheet, column `Incidents`):

- `P27` (Version A): Red NIC2 status at F4 that could not be fixed (lost electrode contact); largest percent removed 41.78.
- `P41` (Version A): LSL stream gap and drift; 30.42 percent removed.
- `P45` (Version B): heavy movement and blinking; 30.64 percent removed.

That leaves 22 Version A and 23 Version B respondents. The respondent with the stream gap (P41) also shows `Samples recorded` below `Samples expected` on the **Artifact** sheet (column `M` holds the difference). Nine other respondents (P01, P03, P07, P15, P31, P38, P39, P40, and P42) had a Red status at the first impedance check that was fixed before recording; the log records Green for them, so the rule keeps them. Under the stricter 15 percent rule `P04` (Version A, 22.83 percent), `P18` (Version A, 17.28 percent), `P22` (Version B, 19.79 percent), and `P32` (Version B, 21.01 percent) are also dropped.

**Common mistake.** Applying the rule to the annotation rows or to a single channel. The rule looks at both frontal channels and at every baseline and commercial, because a respondent whose F3 or F4 signal is unreliable in any of them cannot give a trustworthy baseline corrected score.

---

## 4. Descriptive table (Summary sheet)

The **Summary** sheet gives n, mean, and SD of the raw FA for each role and version (`*` = both versions) among kept respondents. The formulas in row 5 (`Baseline_1`, Version A):

| Statistic | Cell | Formula |
|---|---|---|
| n | `Summary!C5` | `=COUNTIFS(Data!$O$2:$O$313,$A5,Data!$M$2:$M$313,$B5,Data!$P$2:$P$313,"Yes")` |
| Mean FA | `Summary!D5` | `=AVERAGEIFS(Data!$K$2:$K$313,Data!$O$2:$O$313,$A5,Data!$M$2:$M$313,$B5,Data!$P$2:$P$313,"Yes")` |
| SD FA (array formula) | `Summary!E5` | `=STDEV.S(IF((Data!$O$2:$O$313=$A5)*(Data!$M$2:$M$313=$B5)*(Data!$P$2:$P$313="Yes"),Data!$K$2:$K$313))` |

**You should see** raw FA values mostly below zero (as in the demo study chart): `Baseline_1` M = -0.052 and `Baseline_2` M = -0.054 over all kept respondents, with SDs of about 0.13. The SD of the raw values is large because every respondent carries a stable asymmetry of their own; that stable part is what the baseline correction removes. A PivotTable gives the same numbers: **Insert > PivotTable**, **Keep** in **Filters** set to **Yes**, **Role** in **Rows**, **Version** in **Columns**, and **Mean Frontal Asymmetry Alpha** in **Values** set to **Average** (Creating Excel Pivot Tables article).

---

## 5. H1: key scene versus the baseline before Commercial_1A (Version A)

The **H1** sheet has one row per kept Version A respondent (rows 5 to 26). The first respondents (rows 5 to 24) are also kept under the 15 percent rule, and the last 2 (rows 25 to 26) are dropped by it, so the sensitivity block only shortens the range.

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Order | `H1!B5` | `=INDEX(Data!$N$2:$N$313,MATCH($A5,Data!$B$2:$B$313,0))` |
| Keep15 | `H1!C5` | `=INDEX(Data!$Q$2:$Q$313,MATCH($A5,Data!$B$2:$B$313,0))` |
| Baseline before Commercial_1A | `H1!D5` | `=IF($B5="C1first",SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"Baseline_1"),SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"Baseline_2"))` |
| KeyScene | `H1!E5` | `=SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"KeyScene")` |
| Change (KeyScene minus baseline) | `H1!F5` | `=E5-D5` |

The baseline formula is the heart of the design: for a `C1first` respondent the baseline right before `Commercial_1A` is `Baseline_1`; for a `C2first` respondent it is `Baseline_2`.

The test sits in column L (formula text in column M). A paired t test **is** a one sample t test on the change scores:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Baseline mean (FA) | `H1!L6` | `=AVERAGE(D5:D26)` | -0.090 |
| KeyScene mean (FA) | `H1!L8` | `=AVERAGE(E5:E26)` | 0.008 |
| n | `H1!L10` | `=COUNT(F5:F26)` | 22 |
| Mean change (KeyScene minus baseline) | `H1!L11` | `=AVERAGE(F5:F26)` | 0.097 |
| SD of the changes | `H1!L12` | `=STDEV.S(F5:F26)` | 0.142 |
| SE of the mean change | `H1!L13` | `=L12/SQRT(L10)` | 0.0303 |
| t | `H1!L14` | `=L11/L13` | 3.21 |
| df | `H1!L15` | `=L10-1` | 21 |
| **p, two tailed (T.TEST type 1)** | `H1!L16` | `=T.TEST(D5:D26,E5:E26,2,1)` | 0.00421 |
| p from the hand computed t | `H1!L17` | `=T.DIST.2T(ABS(L14),L15)` | 0.00421 |
| p, one tailed (teaching only) | `H1!L18` | `=T.DIST.RT(L14,L15)` | 0.00211 |
| Mean change 95% CI lower | `H1!L20` | `=L11-L19*L13` | 0.034 |
| Mean change 95% CI upper | `H1!L21` | `=L11+L19*L13` | 0.160 |

`T.TEST(array1, array2, tails, type)`: `tails = 2` for a two tailed test; `type = 1` for **paired** data. `type = 2` and `type = 3` treat the two columns as different people and are wrong here. **You should see** the hand computed p equal the `T.TEST` p.

### Cohen's d_z with a 95 percent confidence interval

For paired data the effect size is **d_z**, the mean change in units of the SD of the changes (t = d_z x sqrt(n)).

| Step | Cell | Formula | Result |
|---|---|---|---|
| d_z | `H1!L22` | `=L11/L12` | 0.68 |
| SE of d_z (approximation) | `H1!L23` | `=SQRT(1/L10+L22^2/(2*L10))` | 0.237 |
| d_z 95% CI lower (approximation) | `H1!L24` | `=L22-1.96*L23` | 0.22 |
| d_z 95% CI upper (approximation) | `H1!L25` | `=L22+1.96*L23` | 1.15 |
| Correlation of the two columns | `H1!L26` | `=CORREL(D5:D26,E5:E26)` | 0.46 |
| Respondents with a positive change | `H1!L27` | `=COUNTIF(F5:F26,">0")` | 17 |

In words: SE(d_z) = sqrt(1/n + d_z^2 / (2 n)), the large sample approximation. `analysis.R` also reports the exact interval from the noncentral t distribution, [0.21, 1.14] (Lakens, 2013); say which one you report. Rule of thumb for d_z: 0.2 small, 0.5 medium, 0.8 large. The Wilcoxon signed rank test, the check that the result does not depend on normality, has no Excel function; in R: V = 210, p = .007.

**Why the key scene scatters more.** The 6 s key scene holds only 11 one second Welch windows, the 60 s baseline 119, so a key scene value averages out much less noise. The correlation of the two columns (0.46) is lower than the correlation of the whole `Commercial_1A` with the same baseline (r = 0.63 in step 6 of `analysis_output.txt`), and the SD of the key scene changes (0.142) is larger than the SD of the whole `Commercial_1A` changes (0.098). A short window needs either more respondents or a larger effect.

The checks at the end of the column:

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| Kept Version A respondents with a KeyScene row in Data | `H1!L96` | `=COUNTIFS(Data!$O$2:$O$313,"KeyScene",Data!$P$2:$P$313,"Yes")` | 22 |
| Roster rows dropped under the 15 percent rule | `H1!L97` | `=COUNTIF(C5:C26,"No")` | 2 |
| Keep15 = No inside the 15 percent block (must be 0) | `H1!L98` | `=COUNTIF(C5:C24,"No")` | 0 |

**Common mistake.** Comparing the key scene with `Baseline_1` for everyone. For `C2first` respondents, `Baseline_1` came before `Commercial_2`, and several minutes and another commercial separate it from the key scene.

---

## 6. H2: end scene of Commercial 1, Version A versus Version B (Welch t test)

The **H2** sheet lists all 45 kept respondents: Version A in rows 5 to 26 and Version B in rows 27 to 49. Within each version, the respondents dropped by the 15 percent rule sit at the bottom (A rows 25 to 26, B rows 48 to 49), so the sensitivity block uses rows 5 to 24 and 27 to 47.

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Version | `H2!B5` | `=INDEX(Data!$M$2:$M$313,MATCH($A5,Data!$B$2:$B$313,0))` |
| Baseline before Commercial 1 | `H2!E5` | `=IF($C5="C1first",SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"Baseline_1"),SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"Baseline_2"))` |
| EndScene of Commercial 1 | `H2!F5` | `=SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"EndScene_C1")` |
| Change | `H2!G5` | `=F5-E5` |

Different respondents saw the two versions, so this is a **two group** comparison. Welch's t test does not assume equal variances and is the recommended default (Delacre et al., 2017). The block sits in column J:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Version A mean change | `H2!J6` | `=AVERAGE(G5:G26)` | 0.033 |
| Version A SD | `H2!J7` | `=STDEV.S(G5:G26)` | 0.124 |
| Version A n | `H2!J8` | `=COUNT(G5:G26)` | 22 |
| Version B mean change | `H2!J9` | `=AVERAGE(G27:G49)` | -0.012 |
| Version B SD | `H2!J10` | `=STDEV.S(G27:G49)` | 0.123 |
| Version B n | `H2!J11` | `=COUNT(G27:G49)` | 23 |
| Difference, A minus B | `H2!J12` | `=J6-J9` | 0.045 |
| SE of the difference (Welch) | `H2!J13` | `=SQRT(J7^2/J8+J10^2/J11)` | 0.0367 |
| Welch t | `H2!J14` | `=J12/J13` | 1.22 |
| Welch df | `H2!J15` | `=(J7^2/J8+J10^2/J11)^2/((J7^2/J8)^2/(J8-1)+(J10^2/J11)^2/(J11-1))` | 42.87 |
| **p, two tailed (T.TEST type 3)** | `H2!J16` | `=T.TEST(G5:G26,G27:G49,2,3)` | 0.22864 |
| p from the hand computed t | `H2!J17` | `=T.DIST.2T(ABS(J14),J15)` | 0.2288 |
| p, one tailed (teaching only) | `H2!J18` | `=IF(J14>0,J16/2,1-J16/2)` | 0.11432 |
| Difference 95% CI lower | `H2!J20` | `=J12-J19*J13` | -0.029 |
| Difference 95% CI upper | `H2!J21` | `=J12+J19*J13` | 0.119 |
| Pooled SD | `H2!J22` | `=SQRT(((J8-1)*J7^2+(J11-1)*J10^2)/(J8+J11-2))` | 0.123 |
| Cohen's d | `H2!J23` | `=J12/J22` | 0.36 |
| Hedges' g | `H2!J25` | `=J24*J23` | 0.36 |
| g 95% CI lower (approximation) | `H2!J27` | `=J25-1.96*J26` | -0.23 |
| g 95% CI upper (approximation) | `H2!J28` | `=J25+1.96*J26` | 0.95 |

`T.TEST(..., 2, 3)`: `type = 3` is the unequal variance (Welch) test for two groups. The Welch df is not a whole number. Spreadsheet functions such as `T.DIST.2T` and `T.INV.2T` may round a fractional df down to a whole number (LibreOffice does), so the hand computed p and the interval can differ from R in the third or fourth decimal; **report the `T.TEST` p**, which matched R to six decimals in the recalculated copy. Hedges' g is Cohen's d times a small sample correction; `analysis.R` also gives the exact interval, g = 0.36, 95% CI [-0.22, 0.94], and the Mann Whitney check, W = 313, p = .177.

**Reading a nonsignificant result.** The difference points the predicted way (Version A higher), but the 95 percent interval of the difference includes zero, so the data are compatible with no difference and with a difference as large as g = 0.94. That is not evidence that the key scene has no effect on the end scene. The power analysis in R (step 8 of `analysis_output.txt`) shows why: at the observed d = 0.36, a new study would need about 120 respondents **per version** for 80 percent power, and this study had a power of about 0.22. Between subjects comparisons need far more respondents than within subjects ones, because the differences between people stay in the error term.

The checks at the end of the column:

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| Kept respondents with an EndScene_C1 row in Data | `H2!J87` | `=COUNTIFS(Data!$O$2:$O$313,"EndScene_C1",Data!$P$2:$P$313,"Yes")` | 45 |
| Version B rows inside the Version A range (must be 0) | `H2!J88` | `=COUNTIF(B5:B26,"B")` | 0 |
| Version A rows inside the Version B range (must be 0) | `H2!J89` | `=COUNTIF(B27:B49,"A")` | 0 |

---

## 7. H3: Commercial_2 versus Commercial_1 (paired), and what Excel cannot do

The **H3** sheet uses the same roster order as H2. Each commercial is compared with the baseline right before it, so the two changes of a respondent use **different** baselines:

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Baseline before C1 | `H3!D5` | `=IF($C5="C1first",SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"Baseline_1"),SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"Baseline_2"))` |
| Commercial_1 (FA) | `H3!E5` | `=SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"C1")` |
| C1 change | `H3!F5` | `=E5-D5` |
| Baseline before C2 | `H3!G5` | `=IF($C5="C1first",SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"Baseline_2"),SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"Baseline_1"))` |
| Commercial_2 (FA) | `H3!H5` | `=SUMIFS(Data!$K$2:$K$313,Data!$B$2:$B$313,$A5,Data!$O$2:$O$313,"C2")` |
| C2 change | `H3!I5` | `=H5-G5` |
| Difference (C2 change minus C1 change) | `H3!J5` | `=I5-F5` |

The paired block sits in column O:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| C1 change mean | `H3!O6` | `=AVERAGE(F5:F49)` | -0.029 |
| C2 change mean | `H3!O8` | `=AVERAGE(I5:I49)` | 0.004 |
| n | `H3!O10` | `=COUNT(J5:J49)` | 45 |
| Mean difference (C2 change minus C1 change) | `H3!O11` | `=AVERAGE(J5:J49)` | 0.033 |
| SD of the differences | `H3!O12` | `=STDEV.S(J5:J49)` | 0.112 |
| t | `H3!O14` | `=O11/O13` | 1.98 |
| df | `H3!O15` | `=O10-1` | 44 |
| **p, two tailed (T.TEST type 1)** | `H3!O16` | `=T.TEST(F5:F49,I5:I49,2,1)` | 0.05455 |
| Mean difference 95% CI lower | `H3!O20` | `=O11-O19*O13` | -0.001 |
| Mean difference 95% CI upper | `H3!O21` | `=O11+O19*O13` | 0.066 |
| d_z | `H3!O22` | `=O11/O12` | 0.29 |
| d_z 95% CI lower (approximation) | `H3!O24` | `=O22-1.96*O23` | -0.00 |
| d_z 95% CI upper (approximation) | `H3!O25` | `=O22+1.96*O23` | 0.59 |

**Order and version.** Every respondent saw both commercials, so the one seen second may be affected by the first (carryover or fatigue). For `C1first` respondents the difference C2 change minus C1 change is the commercial effect **plus** the position effect (Commercial_2 came second); for `C2first` respondents it is the commercial effect **minus** the position effect. Half the gap between the two order groups estimates the position effect:

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| Mean difference, C1first | `H3!O30` | `=AVERAGEIFS(J5:J49,C5:C49,"C1first")` | 0.0339 |
| Mean difference, C2first | `H3!O31` | `=AVERAGEIFS(J5:J49,C5:C49,"C2first")` | 0.0317 |
| Position (second minus first), simple estimate | `H3!O34` | `=(O30-O31)/2` | 0.0011 |
| Commercial effect, position balanced | `H3!O35` | `=(O30+O31)/2` | 0.0328 |
| Commercial x version, simple estimate (A minus B) | `H3!O38` | `=O36-O37` | -0.0450 |

`analysis.R` fits the linear mixed model `FA_change ~ commercial * version + position + (1 | respondent)` with lme4 (Bates et al., 2015) and lmerTest, which Excel cannot do. Its estimates differ slightly from the simple ones because the four version by order cells are unequal after the exclusions: commercial b = 0.0323 (p = .059), position b = 0.0006 (p = .971), commercial x version b = -0.0450 (p = .184), and the **intraclass correlation** ICC = 0.110, the share of the remaining variance that lies between respondents. The ICC is low because the baseline correction already removed each respondent's stable asymmetry.

---

## 8. Exploratory: FA change and liking

Columns K and L of the H3 sheet pull the liking ratings from the **Survey** sheet with `INDEX` and `MATCH`. The correlation blocks use `CORREL`, then the usual t and Fisher z formulas:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Commercial_1: Pearson r | `H3!O44` | `=CORREL(F5:F49,K5:K49)` | 0.17 |
| Commercial_1: t for r | `H3!O45` | `=O44*SQRT((O43-2)/(1-O44^2))` | 1.14 |
| Commercial_1: p, two tailed | `H3!O46` | `=T.DIST.2T(ABS(O45),O43-2)` | 0.26067 |
| Commercial_1: r 95% CI lower (Fisher z) | `H3!O47` | `=TANH(ATANH(O44)-NORM.S.INV(0.975)/SQRT(O43-3))` | -0.13 |
| Commercial_1: r 95% CI upper (Fisher z) | `H3!O48` | `=TANH(ATANH(O44)+NORM.S.INV(0.975)/SQRT(O43-3))` | 0.44 |
| Commercial_2: Pearson r | `H3!O54` | `=CORREL(I5:I49,L5:L49)` | 0.29 |
| Commercial_2: p, two tailed | `H3!O56` | `=T.DIST.2T(ABS(O55),O53-2)` | 0.05103 |

Approach motivation and liking are different things (approach can come with anger or curiosity), and a rating and a brain signal are different kinds of measure, so a small to moderate correlation is the expected result. The Spearman rank correlations are in R.

---

## 9. Sensitivity analyses

**Stricter threshold (15 percent).** The sensitivity blocks use the shorter row ranges described in sections 5 and 6:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| H1, 15 percent: n | `H1!L36` | `=COUNT(F5:F24)` | 20 |
| H1, 15 percent: mean change | `H1!L37` | `=AVERAGE(F5:F24)` | 0.100 |
| H1, 15 percent: t | `H1!L40` | `=L37/L39` | 3.03 |
| H1, 15 percent: p (T.TEST type 1) | `H1!L42` | `=T.TEST(D5:D24,E5:E24,2,1)` | 0.00683 |
| H1, 15 percent: d_z | `H1!L48` | `=L37/L38` | 0.68 |
| H2, 15 percent: difference | `H2!J39` | `=J33-J36` | 0.040 |
| H2, 15 percent: Welch t | `H2!J41` | `=J39/J40` | 1.01 |
| H2, 15 percent: p (T.TEST type 3) | `H2!J43` | `=T.TEST(G5:G24,G27:G47,2,3)` | 0.31816 |
| H2, 15 percent: Hedges' g | `H2!J52` | `=J51*J50` | 0.31 |

**Without baseline correction.** H1 then asks whether the raw `KeyScene` FA is above zero, and H2 compares the raw `EndScene` FA of the two versions:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| H1 raw: mean KeyScene FA | `H1!L59` | `=AVERAGE(E5:E26)` | 0.008 |
| H1 raw: t against 0 | `H1!L62` | `=L59/L61` | 0.23 |
| H1 raw: p, two tailed | `H1!L64` | `=T.DIST.2T(ABS(L62),L63)` | 0.81897 |
| H2 raw: difference, A minus B | `H2!J66` | `=J60-J63` | -0.021 |
| H2 raw: Welch t | `H2!J68` | `=J66/J67` | -0.50 |
| H2 raw: p (T.TEST type 3) | `H2!J70` | `=T.TEST(F5:F26,F27:F49,2,3)` | 0.61873 |
| H2 raw: pooled SD | `H2!J76` | `=SQRT(((J62-1)*J61^2+(J65-1)*J64^2)/(J62+J65-2))` | 0.141 |

FA has no meaningful zero point: each respondent carries a stable asymmetry, and most raw scores are negative. The raw key scene FA therefore says nothing about the key scene, and the raw end scene comparison carries a chance difference between the baselines of the two groups (`H2!J90`, -0.066) and a larger SD (pooled 0.141 against 0.123 for the change scores). This is why the planned tests use baseline corrected scores.

---

## 10. Teaching contrast: the notebook default channels

The FA notebook's Enobio defaults are **left Ch1, right Ch7**. With the lab manual montage those are Fz (midline frontal) and P3 (left parietal), not a left and right frontal pair. The **Default** sheet holds the export made with those defaults, and columns G to I of the H1 sheet repeat H1 on it:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Default channels: baseline mean | `H1!L70` | `=AVERAGE(G5:G26)` | 0.744 |
| Default channels: KeyScene mean | `H1!L72` | `=AVERAGE(H5:H26)` | 0.734 |
| Default channels: mean change | `H1!L75` | `=AVERAGE(I5:I26)` | -0.011 |
| Default channels: t | `H1!L78` | `=L75/L77` | -0.40 |
| Default channels: p (T.TEST type 1) | `H1!L80` | `=T.TEST(G5:G26,H5:H26,2,1)` | 0.69070 |
| Default channels: d_z | `H1!L86` | `=L75/L76` | -0.09 |
| Correlation of the default channel change with the F3/F4 change | `H1!L99` | `=CORREL(F5:F26,I5:I26)` | 0.17 |

**You should see** "FA" values near +0.7 in every row, a first warning sign: ln(P3 alpha) - ln(Fz alpha) is large and positive because alpha is stronger over the back of the head. The key scene effect disappears, and the default channel changes are unrelated to the F3/F4 changes. Nothing in the export flags the mistake; only the channel map does. Read the NIC2 protocol (the order in which the electrodes were dragged onto the montage) before setting the left and right electrode sets.

---

## 11. Charts with 95 percent confidence intervals

1. Make a small chart table: one row per mean, columns **Mean** (point to the mean cells of a block) and **CI half width** = `CONFIDENCE.T(0.05, SD cell, n cell)`. The workbook has one on the H1 sheet at `H1!K104:M105` (half widths in `H1!M104` and `H1!M105`) and one on the H2 sheet at `H2!I95:K96` (half widths in `H2!K95` and `H2!K96`).
2. Select the labels and the means, **Insert > Column chart > Clustered Column**.
3. Click the bars, **Chart Design > Add Chart Element > Error Bars > More Error Bars Options**, choose **Custom > Specify Value**, and select the CI half width cells for both **Positive** and **Negative**.

**Common mistake.** Reading two overlapping error bars on the H1 chart as "no difference". The paired test uses the interval of the **changes** (`H1!L20` and `H1!L21`), which is narrower because it removes the differences between respondents. Bars around negative means also hang below zero; FA is a difference score, so a bar chart starting at zero is less informative than the dot and line figures in `figures/`.

---

## 12. Check your numbers and know what Excel cannot do

The **Check** sheet puts each live Excel result next to the value `analysis.R` produced and prints OK or CHECK (status in column F, the count of CHECK rows in `Check!B204`; it should be 0). Several parts of the R analysis have no Excel equivalent:

- **Wilcoxon signed rank** (H1, H3) and **Mann Whitney** (H2) tests, the check that a conclusion does not depend on normality.
- **Exact noncentral t intervals** for d_z, d, and g.
- The **mixed model** with Satterthwaite tests and the **ICC**.
- The **power analysis** (`power.t.test`) for H2.
- **Shapiro Wilk** tests, QQ plots, and **Spearman** correlations.
- The **frontal midline theta** descriptives and the FA against PSD consistency check, which use the PSD export.

If you have the **Analysis ToolPak** add in, **Data > Data Analysis > t-Test: Paired Two Sample for Means** (H1, H3) and **t-Test: Two-Sample Assuming Unequal Variances** (H2) are good cross checks.

---

## 13. Common mistakes, collected

1. Opening the .csv by double clicking and getting one column, or forgetting that the FA files have two comment lines.
2. Running the FA notebook with its default channels without checking the NIC2 montage (section 10).
3. Identifying a segment by `Label` alone; `EndScene` exists on all three commercials (section 2).
4. Comparing a commercial with the wrong baseline; each commercial goes with the baseline right before it (section 5).
5. Leaving excluded respondents in a roster, or applying the 25 percent rule to one channel only.
6. Using `T.TEST` type 1 for the two versions (H2) or type 2 or 3 for the within subjects tests (H1, H3).
7. Reading the raw FA sign as approach or withdrawal. FA is interpreted relative to a baseline or to another stimulus.
8. Reading a nonsignificant H2 as "no effect". Report the interval and the power.
9. Treating 312 rows as 312 independent observations. Each respondent contributes several rows; the tests use one value per respondent.
10. Reading approach as liking. FA indexes approach or withdrawal motivation, not whether the viewer liked the commercial.

## Sources

- iMotions help center, R Notebook: EEG Frontal Asymmetry (FA): https://help.imotions.com/lab/analysis/eeg/r-notebook-eeg-frontal-asymmetry-fa
- iMotions help center, R Processed EEG Export Files: https://help.imotions.com/lab/analysis/eeg/r-processed-eeg-export-files
- iMotions help center, R Notebook: EEG Power Spectral Density (PSD): https://help.imotions.com/lab/analysis/eeg/r-notebook-eeg-power-spectral-density-psd
- iMotions help center, How to check EEG data quality: https://help.imotions.com/lab/sensor-setup/eeg/how-to-check-eeg-data-quality
- iMotions help center, Setting up NeuroElectrics Enobio: https://help.imotions.com/lab/sensor-setup/eeg/neuroelectrics/setting-up-neuroelectrics-enobio
- iMotions help center, Annotations: https://help.imotions.com/lab/analysis/my-library/annotations
- iMotions help center, Opening Data Export Files in Excel: https://help.imotions.com/lab/exports/additional-export-information/opening-data-export-files-in-excel
- iMotions help center, Creating Excel Pivot Tables: https://help.imotions.com/lab/exports/additional-export-information/creating-excel-pivot-tables
- Bates, D., Mächler, M., Bolker, B., & Walker, S. (2015). Fitting linear mixed-effects models using lme4. *Journal of Statistical Software, 67*(1), 1-48. https://doi.org/10.18637/jss.v067.i01
- Delacre, M., Lakens, D., & Leys, C. (2017). Why psychologists should by default use Welch's t-test instead of Student's t-test. *International Review of Social Psychology, 30*(1), 92-101. https://doi.org/10.5334/irsp.82
- Lakens, D. (2013). Calculating and reporting effect sizes to facilitate cumulative science: A practical primer for t-tests and ANOVAs. *Frontiers in Psychology, 4*, Article 863. https://doi.org/10.3389/fpsyg.2013.00863

## Items to verify on the lab computer

- The **channel map**: which 10-10 site each of Ch1 to Ch8 holds in the lab NIC2 protocol (this workbook assumes Ch5 = F3 and Ch6 = F4), and therefore which channels to enter as the left and right frontal electrode sets.
- The **artifact report**: the layout of the notebook report box "Percentage of data filtered (nb of samples filtered / nb total of samples)", whether it gives values per stimulus or for the whole recording, and whether the threshold is 120 or 100 microvolts in Lab 11.1. The simulated file is a modeled table.
- The **FA and PSD export details**: the Type values written for annotation rows ("Annotation" or "Annotation (Instance)"), what Parent Stimulus holds on a stimulus row, whether "Algorithm Device Name" is one column or two, the file type (.csv or .txt), and the PSD column order.
- How the notebook computes a **stimulus mean FA** (the mean of the 1 s FA values, as simulated here, or ln of the mean powers) and how it handles a segment in which every window was removed (NA in the simulation).
- The **session log** and **survey** layouts are lab designs, not iMotions exports.
- The sampling rate shown in the LSL widget (500 Hz expected) and whether the Enobio device profile was selected before recording.
