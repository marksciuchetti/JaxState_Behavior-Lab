# GSR and heart rate analysis in Excel: step by step walkthrough

**SIMULATED DATA.** Every number in this walkthrough comes from the SIMULATED files that `make_simulated_data.py` generated (seed 20278504). No real person was recorded. The design follows the structure of the lab's own library studies "Nike Campaign Ad Evaluation" and "e.l.f. Campaign Ad Evaluation" (a 60 s Baseline, the commercial, a 15 s Recovery, and a short questionnaire; the Nike commercial is 60 s and the e.l.f. commercial 30 s); nothing about the content of either commercial was used.

This is the same analysis as `analysis.R`, done in Excel for students who do not use R. The finished workbook is `SIMULATED_gsr_analysis.xlsx`; the sheet names below refer to it. You can follow along in that workbook or rebuild it from the CSV files. **Every cell address in this walkthrough is written as `Sheet!Cell` and was generated from the workbook itself** (`make_excel_workbook.py` writes both and then reads the workbook back with openpyxl to confirm that every address holds the formula printed here), so the addresses match the file.

Companion files: `analysis.R` (R version), `analysis_output.txt` (its console output), `results_text.md` (model paragraphs), `figures/` (PNG figures).

---

## 0. What you are analyzing

**Design.** Within subjects. Each respondent (P01 to P30) watched two sequences on the lab monitor with a Shimmer3 GSR+ on the non-dominant hand: a Baseline (60 s), the commercial, a Recovery (15 s), and a short questionnaire, for Nike and for e.l.f., with a 2 minute rest between the two sequences. The order was counterbalanced: Group A saw Nike first, Group B saw e.l.f. first (15 each). Respondents were run two at a time at Seat 1 (unit `Shimmer3-8956`) and Seat 2 (unit `Shimmer3-B`, a placeholder name to verify on the lab machine). The analysis pools both devices, as the **Multiple device aggregation** option of the GSR R Notebooks does. The six stimulus labels are `Baseline_Nike`, `Nike_Commercial`, `Recovery_Nike`, `Baseline_elf`, `elf_Commercial`, and `Recovery_elf`.

**Exports.** Four files feed the workbook.

| File | iMotions export | Columns used here | Used for |
|---|---|---|---|
| `SIMULATED_gsr_peak_metrics.csv` | GSR Peak Metrics (Export: R Processed GSR Files), one row per respondent and stimulus | `Respondent Name`, `Group`, `Label`, `Duration (ms)`, `Signal Duration (ms)`, `Has Peaks`, `Peak Count`, `Peaks Per Minute`, `Average Peak Amplitude (microSiemens)` | H1, H2, descriptives |
| `SIMULATED_gsr_signal_quality.csv` | GSR Signal Quality export (the layout is modeled) | `Signal-to-Noise ratio (dB)`, `Signal missing (%)`, `Device` | Exclusion rule, device check |
| `SIMULATED_heart_rate.csv` | Mean of the PPG beats per minute channel per stimulus (the column name `Mean PPG Heart Rate (BPM)` is **modeled**; the manuals name no PPG heart rate column, so verify it on the lab machine) | `Mean PPG Heart Rate (BPM)` | H3 |
| `SIMULATED_gsr_survey.csv` | Questionnaire | `Arousal_Nike_1to9`, `Arousal_elf_1to9` | Exploratory correlation |

A fifth file, `SIMULATED_gsr_peak_table.csv` (one row per detected peak), is used in R for the moment to moment figure and for consistency checks; the workbook does not need it. `SIMULATED_gsr_session_log.csv` is the operator log (seat, device, electrode placement, skin note, warm up minutes, decision notes).

`Peaks Per Minute` is `Peak Count` divided by `Signal Duration` in minutes. `Signal Duration (ms)` leaves out gaps longer than 4 seconds, so a Bluetooth dropout shortens it (section 8). `Has Peaks` is 1 when the stimulus held at least one peak and 0 when it held none. `Average Peak Amplitude` is NA when there is no peak, because it cannot be calculated.

**Hypotheses.**
- H1 (arousal response): Peaks Per Minute is higher during each commercial than during its own baseline.
- H2: Peaks Per Minute differs between the Nike and the e.l.f. commercial, checked for an order (carryover) effect.
- H3: mean heart rate falls during each commercial relative to its baseline (a deceleration), and the change differs between the commercials.
- Exploratory: self reported arousal is related to Peaks Per Minute.
- Sensitivity: the H1 and H2 results without the nonresponders.

**Decision rules fixed before looking at the data.** Exclude a respondent whose Signal-to-Noise ratio is below 0 dB or whose Signal missing is above 20 percent in any commercial (the lab's rule, built from the GSR Signal Quality article, which calls an SNR below 0 dB poor and above 20 dB good). A **nonresponder** has `Has Peaks` = 0 in at least 4 of the 6 stimuli; nonresponders stay in the main analysis. H1 and H3 differences are **commercial minus baseline**; H2 is **Nike minus e.l.f.** Two tailed tests, alpha = .05.

---

## 1. Import the files and build the Data sheet

iMotions exports are comma separated with a dot decimal. Double clicking a file uses your regional defaults, which can put everything in one column (Opening Data Export Files in Excel article). Use the import route instead.

**Route A, Get Data (Excel 2016 and later):**
1. **Data > From Text/CSV**, select `SIMULATED_gsr_peak_metrics.csv`, click **Import**.
2. In the preview set **Delimiter: Comma** and **File Origin: 65001 Unicode (UTF-8)**. **You should see** 15 columns with `Study Name` in the first column.
3. Click **Transform Data** if the first row shows the `# SIMULATED ...` comment line: in Power Query choose **Home > Remove Rows > Remove Top Rows > 1**, then **Home > Use First Row as Headers**, then **Close & Load**. (A real iMotions export has no comment line; the simulated files carry one so nobody mistakes them for real data. The Peak Table has a second comment line, the METADATA block, so remove 2 rows there.)
4. **You should see** a table with 180 data rows (30 respondents x 6 stimuli).

**Route B, Text Import Wizard (any version):** **File > Open > Browse**, set the file type to **All Files**, open the .csv, choose **Delimited**, start at row 2 (to skip the comment line), tick **Comma**, click **Advanced** and set the decimal symbol to a dot, then **Finish**.

**Joining the other files.** The Peak Metrics file has no Device, SNR, or heart rate. Each of those files has one row per respondent and stimulus, so add a key column to each (`=B2&"|"&G2` on the Peak Metrics sheet, where B is `Respondent Name` and G is `Label`; the key of the other files uses their own respondent and stimulus columns) and pull the values across with `INDEX` and `MATCH` (or `XLOOKUP` in Excel 365). In the finished workbook this is already done: columns P to T of the **Data** sheet hold the matched values as pasted values (gray headers).

**Common mistake.** The NA cells in `Average Peak Amplitude` must stay as the text `NA` (or be blank). Do not replace NA with 0: a zero amplitude would pull the mean down for every respondent who had no peak.

Save as **.xlsx** right away; a .csv does not keep formulas or extra sheets. Keep the original .csv files read only.

In the workbook this is the **Data** sheet, header in row 1 and data in rows 2 to 181. Column letters there: `B` = Respondent Name, `E` = Group (the order), `G` = Label, `I` = Duration (ms), `J` = Signal Duration (ms), `L` = Has Peaks, `M` = Peak Count, `N` = Peaks Per Minute, `O` = Average Peak Amplitude, `P` = Device, `Q` = Signal-to-Noise ratio (dB), `R` = Signal missing (%), `S` = Mean PPG Heart Rate (BPM), `T` = Arousal_1to9 (survey; filled on the commercial rows only).

---

## 2. Apply the rules with helper columns

Add seven columns to the right of the data (`U` to `AA` in the workbook, headers in orange) and fill them down to the last row. The formulas below are the ones in row 2.

| Header | Cell | Formula | Meaning |
|---|---|---|---|
| Phase | `Data!U2` | `=IF(LEFT(G2,8)="Baseline","Baseline",IF(LEFT(G2,8)="Recovery","Recovery","Commercial"))` | Baseline, Commercial, or Recovery, read from the label |
| Brand | `Data!V2` | `=IF(ISNUMBER(SEARCH("Nike",G2)),"Nike","elf")` | Nike or elf |
| Position | `Data!W2` | `=IF(OR(AND(E2="A",V2="Nike"),AND(E2="B",V2="elf")),"first","second")` | Was this sequence the first or the second one this respondent saw? Group A saw Nike first, Group B saw e.l.f. first |
| Keep | `Data!X2` | `=IF(COUNTIFS(Data!$B$2:$B$181,B2,Data!$U$2:$U$181,"Commercial",Data!$Q$2:$Q$181,"<0")+COUNTIFS(Data!$B$2:$B$181,B2,Data!$U$2:$U$181,"Commercial",Data!$R$2:$R$181,">20")>0,"No","Yes")` | The exclusion rule: No when any commercial row of this respondent has SNR below 0 dB or Signal missing above 20 percent |
| Nonresponder | `Data!Y2` | `=IF(COUNTIFS(Data!$B$2:$B$181,B2,Data!$L$2:$L$181,0)>=4,1,0)` | 1 when `Has Peaks` is 0 in at least 4 of this respondent's 6 stimuli |
| PPM_recalc | `Data!Z2` | `=M2/(J2/60000)` | Peak Count divided by Signal Duration in minutes, to compare with the exported Peaks Per Minute |
| PPM_if_Duration | `Data!AA2` | `=M2/(I2/60000)` | Peak Count divided by Duration in minutes (the wrong denominator, for the check in section 8) |

**You should see** `Keep = No` on all 12 rows of respondents P11 and P24 and `Yes` on every other row. P11 is excluded for an SNR below 0 dB in the Nike commercial (a loose electrode lead), and P24 for Signal missing of 38.0 percent in the e.l.f. commercial (a Bluetooth dropout); the session log records both decisions. Respondents P03, P16, and P21 have `Nonresponder = 1` on all six of their rows. 3 of the 28 kept respondents (10.7 percent) are nonresponders; they are kept and checked in the sensitivity analysis. 15 kept respondents are in Group A and 13 in Group B.

`COUNTIFS` inside `Keep` and `Nonresponder` counts rows of the same respondent on the whole Data range, so the rule is evaluated per respondent, not per row. A low SNR in a Baseline or Recovery row does not exclude anyone; the rule looks at the commercials only.

**Common mistake.** Applying the SNR rule to every stimulus. The rule was written for the commercials, because they carry the hypotheses.

---

## 3. Descriptive table by stimulus

### 3a. With a PivotTable (the iMotions recommended route)

The Creating Excel Pivot Tables article gives the generic steps: click the first cell of the data, **Insert > PivotTable > OK**, drag a category to **Rows**, drag a metric to **Values**, right click the value, **Value Field Settings**, change **Sum** to **Average**, **OK**. For this study:

1. Click cell A1 of the Data sheet, **Insert > PivotTable**, **New Worksheet**, **OK**.
2. Drag **Keep** to **Filters** and set it to **Yes**.
3. Drag **Label** to **Rows**.
4. Drag **Peaks Per Minute** to **Values** twice. Set the first to **Average** and the second to **StdDev** (the sample standard deviation, the same as `STDEV.S`).
5. Drag **Has Peaks** to **Values** and set it to **Average**; multiply by 100 beside the table to get the percent of respondents with at least one peak.

**You should see** `Nike_Commercial` with an average Peaks Per Minute of 4.658 and `Baseline_Nike` with 2.296.

**Common mistake.** Leaving **Keep** out of the filter. P11 and P24 would then add twelve rows, and the loose electrode lead of P11 would add nine spurious peaks to the Nike commercial.

### 3b. With formulas (the Summary sheet)

A formula table updates when the data change and shows exactly what is being averaged. The stimulus label is in column A of each row (row 5 is the first, `Baseline_Nike`). The formulas in row 5:

| Statistic | Cell | Formula |
|---|---|---|
| n | `Summary!B5` | `=COUNTIFS(Data!$G$2:$G$181,$A5,Data!$X$2:$X$181,"Yes")` |
| PPM mean | `Summary!C5` | `=AVERAGEIFS(Data!$N$2:$N$181,Data!$G$2:$G$181,$A5,Data!$X$2:$X$181,"Yes")` |
| PPM SD | `Summary!D5` | `=STDEV.S(IF((Data!$G$2:$G$181=$A5)*(Data!$X$2:$X$181="Yes"),Data!$N$2:$N$181))` |
| Has Peaks (%) | `Summary!E5` | `=100*AVERAGEIFS(Data!$L$2:$L$181,Data!$G$2:$G$181,$A5,Data!$X$2:$X$181,"Yes")` |
| Peak Count mean | `Summary!F5` | `=AVERAGEIFS(Data!$M$2:$M$181,Data!$G$2:$G$181,$A5,Data!$X$2:$X$181,"Yes")` |
| Peak amplitude mean (microSiemens) | `Summary!G5` | `=AVERAGEIFS(Data!$O$2:$O$181,Data!$G$2:$G$181,$A5,Data!$X$2:$X$181,"Yes")` |
| n with amplitude | `Summary!H5` | `=COUNTIFS(Data!$G$2:$G$181,$A5,Data!$X$2:$X$181,"Yes",Data!$O$2:$O$181,">0")` |
| Heart rate mean (BPM) | `Summary!I5` | `=AVERAGEIFS(Data!$S$2:$S$181,Data!$G$2:$G$181,$A5,Data!$X$2:$X$181,"Yes")` |
| Heart rate SD (BPM) | `Summary!J5` | `=STDEV.S(IF((Data!$G$2:$G$181=$A5)*(Data!$X$2:$X$181="Yes"),Data!$S$2:$S$181))` |
| SNR mean (dB) | `Summary!K5` | `=AVERAGEIFS(Data!$Q$2:$Q$181,Data!$G$2:$G$181,$A5,Data!$X$2:$X$181,"Yes")` |
| Signal missing mean (%) | `Summary!L5` | `=AVERAGEIFS(Data!$R$2:$R$181,Data!$G$2:$G$181,$A5,Data!$X$2:$X$181,"Yes")` |

The two `STDEV.S(IF(...))` formulas (columns D and J) are **array formulas** because Excel has no `STDEVIFS`. In Excel 365 type them and press **Enter**. In Excel 2019 or older press **Ctrl+Shift+Enter**; Excel then shows the formula in braces. `AVERAGEIFS` and `COUNTIFS` ignore the text NA in the amplitude column, which is exactly the "amplitude not calculable" rule.

**You should see** the same numbers as the PivotTable, in rows 5 to 10 (one row per stimulus), and a column chart of mean Peaks Per Minute by stimulus under the tables. In the simulation the mean Peaks Per Minute is 2.296 and 2.471 in the baselines, 4.658 in the Nike commercial, 3.5 in the e.l.f. commercial, and 3.286 and 3.286 in the recoveries. The recoveries last only 15 s, so Peaks Per Minute there moves in steps of 4 (one peak is 4 peaks per minute) and is too coarse for a hypothesis test; the workshop reports it descriptively only.

---

## 4. H1: arousal response, each commercial against its own baseline

A paired test needs the two measurements of the same respondent side by side, so first reshape the long export into one row per respondent. The **H1** sheet does this for the 28 kept respondents (listed in column A, rows 5 to 32); the 25 responders come first (rows 5 to 29) and the 3 nonresponders last (rows 30 to 32). Each of columns D, E, G, and H is a `SUMIFS` that picks the one matching Peaks Per Minute row from the Data sheet:

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Group (order) | `H1!B5` | `=INDEX(Data!$E$2:$E$181,MATCH($A5,Data!$B$2:$B$181,0))` |
| Nonresponder | `H1!C5` | `=INDEX(Data!$Y$2:$Y$181,MATCH($A5,Data!$B$2:$B$181,0))` |
| Baseline_Nike PPM | `H1!D5` | `=SUMIFS(Data!$N$2:$N$181,Data!$B$2:$B$181,$A5,Data!$G$2:$G$181,"Baseline_Nike")` |
| Nike_Commercial PPM | `H1!E5` | `=SUMIFS(Data!$N$2:$N$181,Data!$B$2:$B$181,$A5,Data!$G$2:$G$181,"Nike_Commercial")` |
| Difference Nike | `H1!F5` | `=E5-D5` |
| Baseline_elf PPM | `H1!G5` | `=SUMIFS(Data!$N$2:$N$181,Data!$B$2:$B$181,$A5,Data!$G$2:$G$181,"Baseline_elf")` |
| elf_Commercial PPM | `H1!H5` | `=SUMIFS(Data!$N$2:$N$181,Data!$B$2:$B$181,$A5,Data!$G$2:$G$181,"elf_Commercial")` |
| Difference e.l.f. | `H1!I5` | `=H5-G5` |

(By hand you would get the same table from the PivotTable of section 3a: put **Respondent Name** in Rows, **Label** in Columns, **Peaks Per Minute** in Values as Sum, and filter **Keep** to `Yes`.)

Now the test for **Nike**. The statistics sit in column P of the H1 sheet with the formula text shown beside them in column Q:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Baseline_Nike mean (peaks per minute) | `H1!P6` | `=AVERAGE(D5:D32)` | 2.30 |
| Baseline_Nike SD | `H1!P7` | `=STDEV.S(D5:D32)` | 1.09 |
| Nike_Commercial mean | `H1!P8` | `=AVERAGE(E5:E32)` | 4.66 |
| Nike_Commercial SD | `H1!P9` | `=STDEV.S(E5:E32)` | 2.04 |
| n pairs | `H1!P10` | `=COUNT(F5:F32)` | 28 |
| Mean difference, commercial minus baseline | `H1!P11` | `=AVERAGE(F5:F32)` | 2.36 |
| SD of the differences | `H1!P12` | `=STDEV.S(F5:F32)` | 1.77 |
| **p, two tailed (T.TEST type 1)** | `H1!P16` | `=T.TEST(D5:D32,E5:E32,2,1)` | **0.000000** |

`T.TEST(array1, array2, tails, type)`: `tails = 2` for a two tailed test; `type = 1` for **paired** data, which is the type for this within subjects design. `type = 2` (equal variances) and `type = 3` (unequal variances) treat the two columns as different people and give the wrong answer for paired data. `T.TEST` returns only the p value, so compute t and df by hand with the paired formulas (a paired t test is a one sample t test on the differences):

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| SE of the mean difference | `H1!P13` | `=P12/SQRT(P10)` | 0.334 |
| Paired t | `H1!P14` | `=P11/P13` | 7.06 |
| df | `H1!P15` | `=P10-1` | 27 |
| p from the hand computed t | `H1!P17` | `=T.DIST.2T(ABS(P14),P15)` | 0.000000, matches T.TEST |
| Critical t (95%) | `H1!P18` | `=T.INV.2T(0.05,P15)` | 2.052 |
| Mean difference 95% CI lower | `H1!P19` | `=P11-P18*P13` | 1.68 |
| Mean difference 95% CI upper | `H1!P20` | `=P11+P18*P13` | 3.05 |

**You should see** the hand computed p equal the `T.TEST` p. If they differ, the `type` argument or a column range is wrong.

### Cohen's d_z by hand, with a 95 percent confidence interval

For paired data the effect size is **d_z**: the mean difference in units of the SD of the *differences*. It is linked to t by t = d_z x sqrt(n), so t, n, and d_z always agree.

| Step | Cell | Formula | Result |
|---|---|---|---|
| d_z (Nike) | `H1!P21` | `=P11/P12` | 1.33 |
| SE of d_z (approximation) | `H1!P22` | `=SQRT(1/P10+P21^2/(2*P10))` | 0.260 |
| d_z 95% CI lower | `H1!P23` | `=P21-1.96*P22` | 0.83 |
| d_z 95% CI upper | `H1!P24` | `=P21+1.96*P22` | 1.84 |
| Correlation of the two columns | `H1!P25` | `=CORREL(D5:D32,E5:E32)` | 0.50 |

In words: d_z = mean difference / SD of differences; SE(d_z) = sqrt(1/n + d_z^2 / (2 n)). The interval is the usual large sample approximation. `analysis.R` reports the interval from the noncentral t distribution (effectsize package), which is more exact at this sample size; d_z itself matches, and the limits differ slightly ([0.82, 1.84] in R). Say which method you used. Rule of thumb for reading d_z: 0.2 small, 0.5 medium, 0.8 large.

### The same test for e.l.f.

The e.l.f. block repeats the recipe on columns G, H, and I (the baseline, the commercial, and their difference). The e.l.f. commercial lasts 30 s, so each peak adds 2 peaks per minute to a respondent's rate and the values move in steps of 2 (apart from the rows where a gap shortened `Signal Duration`).

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Baseline_elf mean | `H1!P30` | `=AVERAGE(G5:G32)` | 2.47 |
| elf_Commercial mean | `H1!P32` | `=AVERAGE(H5:H32)` | 3.50 |
| n pairs | `H1!P34` | `=COUNT(I5:I32)` | 28 |
| Mean difference, commercial minus baseline | `H1!P35` | `=AVERAGE(I5:I32)` | 1.03 |
| SD of the differences | `H1!P36` | `=STDEV.S(I5:I32)` | 1.55 |
| Paired t | `H1!P38` | `=P35/P37` | 3.52 |
| df | `H1!P39` | `=P34-1` | 27 |
| p, two tailed (T.TEST type 1) | `H1!P40` | `=T.TEST(G5:G32,H5:H32,2,1)` | 0.001552 |
| d_z (e.l.f.) | `H1!P45` | `=P35/P36` | 0.67 |
| d_z 95% CI lower | `H1!P47` | `=P45-1.96*P46` | 0.26 |
| d_z 95% CI upper | `H1!P48` | `=P45+1.96*P46` | 1.07 |

The Wilcoxon signed rank tests, a check that the conclusion does not depend on normality, have no Excel function; in R: Nike V = 317, e.l.f. V = 190. The differences are built from whole peak counts, so they cannot be perfectly normal; the two tests agree in this sample.

The block below the e.l.f. blocks, with the checks, counts the kept respondents again from the Data sheet:

| Statistic | Cell | Formula |
|---|---|---|
| Kept respondents with a Nike_Commercial row in Data | `H1!P152` | `=COUNTIFS(Data!$G$2:$G$181,"Nike_Commercial",Data!$X$2:$X$181,"Yes")` |
| Kept responders with a Nike_Commercial row in Data | `H1!P153` | `=COUNTIFS(Data!$G$2:$G$181,"Nike_Commercial",Data!$X$2:$X$181,"Yes",Data!$Y$2:$Y$181,0)` |
| Nonresponders listed in the roster | `H1!P154` | `=SUM(C5:C32)` |
| Nonresponders inside the responder rows | `H1!P155` | `=SUM(C5:C29)` |

**You should see** 28 in the first (it must equal n pairs), 25 in the second, 3 in the third, and 0 in the fourth. If the last one is not 0, a nonresponder is listed among the responders and the sensitivity blocks below are wrong.

**Common mistake.** Pairing the wrong rows. The two ranges must list the respondents in the same order, which is why every column of the H1 sheet is a lookup on the respondent name in column A rather than a pasted column.

---

## 5. H2: Nike versus e.l.f., and the order effect

The **H2** sheet lists the two commercial values per respondent (column D e.l.f., column E Nike, column F the difference **Nike minus e.l.f.**, column G whether Nike was the first or the second sequence) in the same roster order as H1. The paired recipe sits in column J (formula text in column K):

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| elf_Commercial mean (peaks per minute) | `H2!J6` | `=AVERAGE(D5:D32)` | 3.50 |
| Nike_Commercial mean | `H2!J8` | `=AVERAGE(E5:E32)` | 4.66 |
| n pairs | `H2!J10` | `=COUNT(F5:F32)` | 28 |
| Mean difference, Nike minus elf | `H2!J11` | `=AVERAGE(F5:F32)` | 1.16 |
| SD of the differences | `H2!J12` | `=STDEV.S(F5:F32)` | 2.00 |
| Paired t | `H2!J14` | `=J11/J13` | 3.07 |
| df | `H2!J15` | `=J10-1` | 27 |
| p, two tailed (T.TEST type 1) | `H2!J16` | `=T.TEST(D5:D32,E5:E32,2,1)` | 0.004832 |
| d_z | `H2!J21` | `=J11/J12` | 0.58 |
| d_z 95% CI lower | `H2!J23` | `=J21-1.96*J22` | 0.18 |
| d_z 95% CI upper | `H2!J24` | `=J21+1.96*J22` | 0.98 |
| Correlation of the two columns | `H2!J25` | `=CORREL(D5:D32,E5:E32)` | 0.48 |

Because every respondent saw both commercials, the second one may be affected by the first (carryover; for GSR the usual suspect is habituation, a smaller response to a repeated kind of stimulus). The simulated viewers were built with about 0.4 fewer peaks per minute for the commercial seen second and 0.4 more for the one seen first. The paired difference carries this: for Group A (Nike first) the difference Nike minus e.l.f. is commercial effect minus order effect, and for Group B (Nike second) it is commercial effect plus order effect. So half the gap between the two group means is the order effect:

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| Mean difference, Group A (Nike seen first) | `H2!J28` | `=AVERAGEIFS(F5:F32,B5:B32,"A")` | 1.67 |
| Mean difference, Group B (Nike seen second) | `H2!J29` | `=AVERAGEIFS(F5:F32,B5:B32,"B")` | 0.57 |
| n, Group A | `H2!J30` | `=COUNTIF(B5:B32,"A")` | 15 |
| n, Group B | `H2!J31` | `=COUNTIF(B5:B32,"B")` | 13 |
| Order effect (second minus first) | `H2!J32` | `=(J29-J28)/2` | -0.55 |
| Commercial effect (Nike minus elf), order adjusted | `H2!J33` | `=(J28+J29)/2` | 1.12 |

`analysis.R` fits `ppm ~ commercial * order_position + (1 | respondent)` with lme4 and lmerTest and gets the order effect b = -0.55, SE = 0.37, p = 0.150, and the commercial effect 1.12; the simple Excel estimate coincides with the model coefficients because each order group's pair of means is free in the model. Excel cannot supply the standard error with the right degrees of freedom, the interaction test, or the **intraclass correlation** (ICC = 0.48 after the fixed effects): that is the share of the remaining variance that lies between respondents, the reason the two rows of a respondent are not independent. Use R for those.

The check at the end of the block:

| Statistic | Cell | Formula |
|---|---|---|
| Kept respondents with a Nike_Commercial row in Data | `H2!J36` | `=COUNTIFS(Data!$G$2:$G$181,"Nike_Commercial",Data!$X$2:$X$181,"Yes")` |

**You should see** 28, equal to n pairs.

### Sensitivity without the nonresponders

A second block on the H2 sheet (column J, below the first) repeats the test on rows 5 to 29 only, the 25 responders:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| n pairs (responders) | `H2!J45` | `=COUNT(F5:F29)` | 25 |
| Mean difference, Nike minus elf | `H2!J46` | `=AVERAGE(F5:F29)` | 1.22 |
| Paired t | `H2!J49` | `=J46/J48` | 2.89 |
| p, two tailed (T.TEST type 1) | `H2!J51` | `=T.TEST(D5:D29,E5:E29,2,1)` | 0.007963 |
| d_z | `H2!J56` | `=J46/J47` | 0.58 |
| Order effect, responders only | `H2!J67` | `=(J64-J63)/2` | -0.58 |

---

## 6. H3: heart rate change from baseline

The **H3** sheet holds the mean heart rate per stimulus (`SUMIFS` on the Data sheet) and the change (commercial minus the baseline that preceded it) in the same roster order:

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Baseline_Nike HR | `H3!D5` | `=SUMIFS(Data!$S$2:$S$181,Data!$B$2:$B$181,$A5,Data!$G$2:$G$181,"Baseline_Nike")` |
| Nike_Commercial HR | `H3!E5` | `=SUMIFS(Data!$S$2:$S$181,Data!$B$2:$B$181,$A5,Data!$G$2:$G$181,"Nike_Commercial")` |
| Change Nike | `H3!F5` | `=E5-D5` |
| Baseline_elf HR | `H3!G5` | `=SUMIFS(Data!$S$2:$S$181,Data!$B$2:$B$181,$A5,Data!$G$2:$G$181,"Baseline_elf")` |
| elf_Commercial HR | `H3!H5` | `=SUMIFS(Data!$S$2:$S$181,Data!$B$2:$B$181,$A5,Data!$G$2:$G$181,"elf_Commercial")` |
| Change e.l.f. | `H3!I5` | `=H5-G5` |
| Difference of changes | `H3!J5` | `=F5-I5` |

Three paired blocks sit in column M (formula text in column N): the Nike commercial against its baseline, the e.l.f. commercial against its baseline, and the **difference between the two changes** (Nike change minus e.l.f. change). In the first two, the mean difference **is** the heart rate change.

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Nike: baseline mean (BPM) | `H3!M6` | `=AVERAGE(D5:D32)` | 76.17 |
| Nike: commercial mean (BPM) | `H3!M8` | `=AVERAGE(E5:E32)` | 74.66 |
| Nike: mean change (BPM) | `H3!M11` | `=AVERAGE(F5:F32)` | -1.51 |
| Nike: SD of the changes | `H3!M12` | `=STDEV.S(F5:F32)` | 1.92 |
| Nike: paired t | `H3!M14` | `=M11/M13` | -4.16 |
| Nike: p (T.TEST type 1) | `H3!M16` | `=T.TEST(D5:D32,E5:E32,2,1)` | 0.000287 |
| Nike: d_z | `H3!M21` | `=M11/M12` | -0.79 |
| e.l.f.: mean change (BPM) | `H3!M35` | `=AVERAGE(I5:I32)` | -0.98 |
| e.l.f.: paired t | `H3!M38` | `=M35/M37` | -2.82 |
| e.l.f.: p (T.TEST type 1) | `H3!M40` | `=T.TEST(G5:G32,H5:H32,2,1)` | 0.008896 |
| e.l.f.: d_z | `H3!M45` | `=M35/M36` | -0.53 |
| Difference of the changes, Nike minus e.l.f. (BPM) | `H3!M59` | `=AVERAGE(J5:J32)` | -0.53 |
| Difference of the changes: SD | `H3!M60` | `=STDEV.S(J5:J32)` | 2.39 |
| Difference of the changes: paired t | `H3!M62` | `=M59/M61` | -1.18 |
| Difference of the changes: p (T.TEST type 1) | `H3!M64` | `=T.TEST(I5:I32,F5:F32,2,1)` | 0.248640 |
| Difference of the changes: d_z | `H3!M69` | `=M59/M60` | -0.22 |

A negative change means the heart rate was lower during the commercial than during its baseline. Heart rate comes from the PPG channel of the Shimmer3, a different sensor from the GSR electrodes, so a nonresponder in GSR can still show a normal heart rate change. The heart rate numbers rest on a modeled column name (see section 0); on real data first confirm in a pilot export that the beats per minute channel exists and what it contains.

The checks at the end of the H3 sheet:

| Statistic | Cell | Formula |
|---|---|---|
| Kept respondents with a Nike_Commercial row in Data | `H3!M78` | `=COUNTIFS(Data!$G$2:$G$181,"Nike_Commercial",Data!$X$2:$X$181,"Yes")` |
| Respondents whose heart rate fell during the Nike commercial | `H3!M79` | `=COUNTIF(F5:F32,"<0")` |
| Respondents whose heart rate fell during the e.l.f. commercial | `H3!M80` | `=COUNTIF(I5:I32,"<0")` |

**You should see** 28, then 21, then 17.

---

## 7. Exploratory correlation and the nonresponder sensitivity analysis

### Self reported arousal and Peaks Per Minute

Self reported arousal (a 1 to 9 self assessment manikin style rating) is collected after each sequence. Columns J and K of the H1 sheet hold it for each respondent, and columns L and M hold the respondent means over the two commercials. The correlation blocks in column P use `CORREL`, then the usual t and Fisher z formulas:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Nike: Pearson r (arousal with Peaks Per Minute) | `H1!P103` | `=CORREL(J5:J32,E5:E32)` | 0.37 |
| Nike: t for r | `H1!P104` | `=P103*SQRT((P102-2)/(1-P103^2))` | 2.00 |
| Nike: p, two tailed | `H1!P105` | `=T.DIST.2T(ABS(P104),P102-2)` | 0.0558 |
| Nike: r 95% CI lower (Fisher z) | `H1!P106` | `=TANH(ATANH(P103)-NORM.S.INV(0.975)/SQRT(P102-3))` | -0.01 |
| Nike: r 95% CI upper (Fisher z) | `H1!P107` | `=TANH(ATANH(P103)+NORM.S.INV(0.975)/SQRT(P102-3))` | 0.65 |
| e.l.f.: Pearson r | `H1!P113` | `=CORREL(K5:K32,H5:H32)` | 0.34 |
| Respondent means: Pearson r | `H1!P123` | `=CORREL(L5:L32,M5:M32)` | 0.30 |

The Spearman rank correlations and the check that no correlation hides behind the outlying nonresponders are in R. Because a rating and Peaks Per Minute are different kinds of measure (what the viewer says and what the skin did), a modest correlation is the expected result; GSR also measures only arousal, not whether the feeling is positive or negative.

### Sensitivity analysis: without the nonresponders

Nonresponders (P03, P16, and P21) show no peaks in most stimuli. They are a physiological finding (some people have very little phasic skin conductance activity), not a data fault, so the main analysis keeps them and the sensitivity analysis removes them. The roster puts them in the last 3 rows (rows 30 to 32), so the sensitivity blocks only change the row range from `5:32` to `5:29`:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Nike, responders only: n pairs | `H1!P58` | `=COUNT(F5:F29)` | 25 |
| Nike, responders only: mean difference | `H1!P59` | `=AVERAGE(F5:F29)` | 2.65 |
| Nike, responders only: paired t | `H1!P62` | `=P59/P61` | 8.11 |
| Nike, responders only: p (T.TEST type 1) | `H1!P64` | `=T.TEST(D5:D29,E5:E29,2,1)` | 0.000000 |
| Nike, responders only: d_z | `H1!P69` | `=P59/P60` | 1.62 |
| e.l.f., responders only: mean difference | `H1!P83` | `=AVERAGE(I5:I29)` | 1.15 |
| e.l.f., responders only: paired t | `H1!P86` | `=P83/P85` | 3.61 |
| e.l.f., responders only: p (T.TEST type 1) | `H1!P88` | `=T.TEST(G5:G29,H5:H29,2,1)` | 0.001392 |
| e.l.f., responders only: d_z | `H1!P93` | `=P83/P84` | 0.72 |

Compare the effect sizes with the full sample: removing the nonresponders raises d_z for Nike from 1.33 to 1.62 and for e.l.f. from 0.67 to 0.72. The conclusions do not change (every test stays on the same side of .05), and that is what a sensitivity analysis is for. The nonresponders also raise the correlation between a respondent's baseline and commercial values; in R, the correlations among responders only are lower.

---

## 8. Why Peaks Per Minute divides by Signal Duration

`Signal Duration (ms)` is the duration of the recorded signal without gaps longer than 4 seconds, and `Peaks Per Minute` divides `Peak Count` by it. A Bluetooth dropout hides a stretch of the stimulus from the sensor; no peak can be detected in the gap, so dividing by the full `Duration` would count the gap as time without arousal. The Summary sheet tests this:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Rows with Signal Duration shorter than Duration (kept respondents) | `Summary!B14` | `=SUMPRODUCT(--(Data!$J$2:$J$181<Data!$I$2:$I$181),--(Data!$X$2:$X$181="Yes"))` | 3 |
| Largest understatement if Duration were used (peaks per minute) | `Summary!B15` | `=SUMPRODUCT(MAX((Data!$X$2:$X$181="Yes")*(Data!$Z$2:$Z$181-Data!$AA$2:$AA$181)))` | 0.415 |
| Largest gap between exported Peaks Per Minute and PPM_recalc | `Summary!B16` | `=SUMPRODUCT(MAX(ABS(Data!$N$2:$N$181-Data!$Z$2:$Z$181)))` | 0.0004 or less (the export rounds to 3 decimals) |

In this data set the biggest effect is on the excluded respondent P24, whose e.l.f. commercial row has a Signal Duration of 18600 ms out of 30000 ms: 2 peaks are 6.452 peaks per minute by Signal Duration and 4.000 by Duration. Among the kept respondents the gaps are shorter (4.6 to 5.1 s of a 60 s stimulus), so the understatement is at most about 0.4 peaks per minute, but it would be a systematic error in the same direction for everyone who had a dropout.

**You should see** the exported `Peaks Per Minute` (column N) and `PPM_recalc` (column Z) agree on every row.

**Common mistake.** Recomputing Peaks Per Minute yourself from `Peak Count` and `Duration (ms)`. Use `Signal Duration (ms)`.

The device check on the Summary sheet (a small table below the Signal Duration block, columns A to D) shows the mean SNR and the mean Peaks Per Minute of the baselines and commercials for each device; the two devices are close (R tests this in section 9 of `analysis_output.txt`), so pooling them is reasonable here.

| Statistic | Cell | Formula |
|---|---|---|
| Rows by device (first device) | `Summary!B19` | `=COUNTIFS(Data!$P$2:$P$181,$A19,Data!$X$2:$X$181,"Yes")` |
| SNR mean by device (first device) | `Summary!C19` | `=AVERAGEIFS(Data!$Q$2:$Q$181,Data!$P$2:$P$181,$A19,Data!$X$2:$X$181,"Yes")` |
| PPM mean by device (first device) | `Summary!D19` | `=AVERAGEIFS(Data!$N$2:$N$181,Data!$P$2:$P$181,$A19,Data!$X$2:$X$181,"Yes",Data!$U$2:$U$181,"<>Recovery")` |

---

## 9. Chart with 95 percent confidence intervals

1. Make a small chart table: one row per baseline and commercial, columns **Mean PPM** (point to the mean cells of the paired blocks) and **CI half width** = `CONFIDENCE.T(0.05, SD cell, n cell)`. `CONFIDENCE.T` returns half the width of a t based 95 percent interval around a mean. The workbook has this table on the H1 sheet at `H1!O160:Q163`; the half widths are in `H1!Q160`, `H1!Q161`, `H1!Q162`, and `H1!Q163`.
2. Select the stimulus names and the means, **Insert > Column chart > Clustered Column**.
3. Click the bars, **Chart Design > Add Chart Element > Error Bars > More Error Bars Options**, choose **Custom > Specify Value**, and select the four CI half width cells for both **Positive** and **Negative**.
4. Set the bar fill to navy (`#1F3864`) and lighten the gridlines.

**Common mistake.** Using the default error bars (**Standard Error** or a fixed 5 percent). Those are not 95 percent confidence intervals. And remember that error bars around the two means are *not* the interval that goes with the paired test: the paired test uses the interval of the *differences* (`H1!P19` and `H1!P20` for Nike), which is much narrower because it removes the between respondent variation.

---

## 10. Check your numbers and know what Excel cannot do

The **Check** sheet puts each live Excel result next to the value `analysis.R` produced and prints OK or CHECK (status in column F, the count of CHECK rows in `Check!B192`). Your Excel results should match to rounding; the d_z limits differ by method and have a looser tolerance. Several parts of the R analysis have no Excel equivalent:

- **Wilcoxon signed rank tests** (a check that the conclusion does not depend on normality). In R: H1 Nike V = 317, H1 e.l.f. V = 190, H2 V = 179, H3 Nike V = 44.5, H3 e.l.f. V = 95.5, H3 difference V = 153.5.
- **Mixed effects model** with a random intercept per respondent and the **ICC**: `lme4::lmer(ppm ~ commercial * order_position + (1 | respondent))`.
- **Shapiro Wilk** tests and QQ plots of the paired differences, **Spearman** correlations, and the **noncentral t** interval for d_z.
- **Moment to moment view** (percent of respondents with a peak in each 5 s bin with 1 s overlap): it needs the GSR Peak Table (or the Moment To Moment Statistics export) and is computed in R.

If you have the **Analysis ToolPak** add in (**File > Options > Add-ins > Manage Excel Add-ins > Go > Analysis ToolPak**), **Data > Data Analysis > t-Test: Paired Two Sample for Means** prints t, df, and p in one table and is a good cross check of sections 4 to 6.

---

## 11. Common mistakes, collected

1. Opening the .csv by double clicking and getting one column: use **Data > From Text/CSV** or the Text Import Wizard.
2. Replacing the NA in `Average Peak Amplitude` with 0. NA means no peak, so the amplitude does not exist.
3. Forgetting the exclusion filter (Keep = Yes) in the PivotTable or in `COUNTIFS`, or leaving P11 and P24 in a roster.
4. Applying the exclusion rule to the baselines and recoveries as well as the commercials.
5. `T.TEST` with `type = 2` or `type = 3` on a within subjects design. Use `type = 1` and make sure the two columns are in the same respondent order.
6. Reporting `T.TEST` alone. Always give M, SD, n, t, df, p, and d_z with its interval.
7. Treating 180 rows as 180 independent observations. Each respondent contributes six rows; tests use one value per respondent and stimulus, and an all rows analysis needs the mixed model.
8. Dividing Peak Count by Duration instead of Signal Duration, or reading a 15 s Recovery or a 30 s commercial Peaks Per Minute as if it had a one minute resolution.
9. Dropping the nonresponders without saying so, or keeping them without checking what they do to the result. Report both.
10. Reading a GSR peak as a positive or a negative feeling. GSR shows arousal (intensity), not valence.

## Sources

- iMotions help center, GSR R-Notebooks: Peak Detection and Aggregation: https://help.imotions.com/lab/analysis/eda-gsr/gsr-r-notebooks-peak-detection-and-aggregation (local copy: `C:\JaxState_Behavior Lab\iMotion_Manual\GSR RNotebooks Peak Detection and Aggegation.pdf`)
- iMotions help center, GSR R-Notebooks: GSR Signal Quality: https://help.imotions.com/lab/analysis/eda-gsr/gsr-r-notebooks-gsr-signal-quality (local copy: `C:\JaxState_Behavior Lab\iMotion_Manual\GSR RNotebooks GSR Signal Quality.pdf`)
- iMotions help center, Export: R Processed GSR Files: https://help.imotions.com/lab/exports/r-processed-exports/export-r-processed-gsr-files (local copy: `C:\JaxState_Behavior Lab\iMotion_Manual\Export R Processed GSR Files.pdf`)
- iMotions help center, PPG, BVP, and Optical Heart Rate: https://help.imotions.com/lab/sensor-setup/shimmer/ppg-bvp-and-optical-heart-rate
- iMotions help center, Opening Data Export Files in Excel: https://help.imotions.com/lab/exports/additional-export-information/opening-data-export-files-in-excel (local copy: `iMotion_Manual\Opening Data Export Files in Excel.pdf`)
- iMotions help center, Creating Excel Pivot Tables: https://help.imotions.com/lab/exports/additional-export-information/creating-excel-pivot-tables (local copy: `iMotion_Manual\Making Excel Pivot Tables.pdf`)

## Items to verify on the lab laptop

- The name, calculation, and units of the PPG beats per minute column in a Sensor Data export, and whether any notebook accepts it. The manuals name no PPG heart rate notebook or column; the column `Mean PPG Heart Rate (BPM)` in the simulated file is a modeled name.
- The exact layout of the GSR Signal Quality export in Lab 11.1 (the manual names Signal-to-Noise ratio in decibels and Signal missing in percent per stimulus and respondent but prints no header row) and whether it is a .txt or a .csv file.
- The Type and Parent Stimulus values that the GSR Peak Metrics export writes for a plain stimulus row.
- The name of the second Shimmer unit (`Shimmer3-B` is a placeholder), the flow names, and whether the **Multiple device aggregation** box appears with two Shimmers in Lab 11.1.
- The peak detection parameter values used (gap interpolation, phasic filter length, low pass cutoff, onset, offset, amplitude threshold, minimum peak duration); the manuals do not state defaults, so record them from the parameter panel.
- Whether the real Moment To Moment Statistics file labels the bins the way the simulation does (the manual does not show the 1 s overlap in the file).
