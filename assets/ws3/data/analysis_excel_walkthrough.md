# Facial expression analysis in Excel: step by step walkthrough

**SIMULATED DATA.** Every number in this walkthrough comes from `SIMULATED_fea_metrics.csv`, which `make_simulated_data.py` generated (seed 20265798). No real viewer was recorded. The design follows the structure of the lab's library study "Nike Campaign Ad Evaluation" (a 60 s Baseline, a 60 s commercial, a 15 s Recovery, a short questionnaire, the respondent camera recorded for AFFDEX); no data from that study were used, and nothing is said about what the commercial shows.

This is the same analysis as `analysis.R`, done in Excel for students who do not use R. The finished workbook is `SIMULATED_fea_analysis.xlsx`; the sheet names below refer to it. You can follow along in that workbook or rebuild it from the CSV files. **Every cell address in this walkthrough is written as `Sheet!Cell` and was generated from the workbook itself** (`make_excel_workbook.py` writes both), so the addresses match the file. The sheets `H1` and `H3` have names that look like cell addresses, so when you type a formula that points at them, put the name in single quotes (`'H1'!L10`); the walkthrough leaves the quotes out.

Companion files: `analysis.R` (R version), `analysis_output.txt` (its console output), `results_text.md` (model paragraphs), `figures/` (PNG figures).

---

## 0. What you are analyzing

**Design.** One viewing. Each of 32 respondents (V01 to V32) watched the same sequence once on the lab monitor while the lab webcam recorded the face: a 60 s Baseline, a 60 s commercial, a 15 s Recovery, and a short questionnaire. The video was post processed with Affectiva AFFDEX (every frame), then the **FEA thresholding** R Notebook turned each expression into present (1) or absent (0) frame by frame at its default thresholds (Emotion 50, Valence +50 and -50, Action Units 50). A binning stimulus annotation cut the commercial into six 10 s bins, `Bin_01` to `Bin_06`. Because everyone saw the Baseline first and only once, the Baseline is a comparison period and not a counterbalanced control.

**Export.** The **FEA Metrics Export** (analysis > Export > R Processing > FEA Metrics Export; choose all or selected expressions) reports, for each respondent and each time interval (Stimulus, Annotation, or Scene), how often and how long an expression was present. The file here has one row per respondent by interval by expression: 32 respondents x 9 intervals x 6 expressions = 1728 data rows. **The manuals print no header row for this export, so the column names below are modeled on their description and must be checked against a real export on the lab machine.** The columns used:

| Column | Meaning | Used for |
|---|---|---|
| `Data quality (%)` | Percentage of frames successfully analyzed after post processing (the Quality Metrics value); not a measure of accuracy | Exclusion rule |
| `Interval type`, `Interval label` | Stimulus (`Baseline`, `Commercial`, `Recovery`) or Annotation (`Bin_01` to `Bin_06`) | Selecting rows |
| `Expression` | Joy, Surprise, Brow Furrow, Positive Valence, Negative Valence, Engagement | Selecting rows |
| `Threshold` | The threshold the notebook applied (50; -50 for Negative Valence) | Documentation |
| `Percent of time present` | Percent of the interval in which the expression was present | H1, H2, H3 |
| `Number of segments` | How many separate stretches of presence the interval holds (how often it was expressed) | Descriptives |
| `Mean raw score (0-100)` | Engagement only: mean AFFDEX score in the interval | Descriptives |

Engagement is not thresholded: the notebook article lists default thresholds for Emotion, Valence, Action Units, and the head position channels, but does not name Engagement, so it is reported as a raw mean (the kind of statistic the AFFDEX Raw Data Statistics Export gives). iMotions eye tracking exports write **NA** when a metric cannot be calculated; the simulated file uses NA the same way (the threshold, percent, and segment columns are NA on Engagement rows, and the raw mean is NA on all other rows). Whether the FEA export does this is an item to verify.

**Hypotheses.**
- H1: the share of time with Joy present is larger during the commercial than during the Baseline. Joy is zero inflated (most viewers show none at Baseline), so the primary test is the **Wilcoxon signed rank test**; the paired t test and d_z are reported beside it.
- H2: Joy `Percent of time present` differs across the six bins; the plan is a concentration in `Bin_05` and `Bin_06`. Tested with a repeated measures ANOVA (the mixed model F test in R), the proportion of viewers with Joy present per bin, and a Friedman test.
- H3: viewers with more Positive Valence time during the commercial like the ad more (Pearson and Spearman correlation with `Ad_liking_1to7`).

**Decision rules fixed before looking at the data.** Exclude a respondent whose `Data quality (%)` is below 80 percent. The manuals state no minimum; 80 percent is the lab's own rule. A viewer has an expression **present** in an interval when `Percent of time present` is above 0 (that is the quantity the thresholding Aggregation counts per time window). The H1 difference is **Commercial minus Baseline**. Two tailed tests, alpha = .05.

---

## 1. Import the export into Excel

iMotions exports are comma separated with a dot decimal. Double clicking the file uses your regional defaults, which can put everything in one column (Opening Data Export Files in Excel article). Use the import route instead.

**Route A, Get Data (Excel 2016 and later):**
1. **Data > From Text/CSV**, select `SIMULATED_fea_metrics.csv`, click **Import**.
2. In the preview set **Delimiter: Comma** and **File Origin: 65001 Unicode (UTF-8)**. **You should see** 11 columns with `Study Name` in the first column.
3. Click **Transform Data** if the first row shows the `# SIMULATED ...` comment line: in Power Query choose **Home > Remove Rows > Remove Top Rows > 1**, then **Home > Use First Row as Headers**, then **Close & Load**. (A real iMotions export has no comment line; the simulated file carries one so nobody mistakes it for real data.)
4. **You should see** a table with 1728 data rows.

**Route B, Text Import Wizard (any version):** **File > Open > Browse**, set the file type to **All Files**, open the .csv, choose **Delimited**, start at row 2 (to skip the comment line), tick **Comma**, click **Advanced** and set the decimal symbol to a dot, then **Finish**.

**Common mistake.** The NA cells must stay as the text `NA` (or be blank). Do not replace NA with 0: a 0 in `Percent of time present` would mean "the expression was never present", which is a measurement, while NA means "no value exists". Step 2 handles both correctly.

Save as **.xlsx** right away; a .csv does not keep formulas or extra sheets. Keep the original .csv read only.

In the workbook this is the **Data** sheet, header in row 1 and data in rows 2 to 1729. Column letters there: `B` = Respondent Name, `C` = Interval type, `D` = Interval label, `F` = Expression, `H` = Percent of time present, `I` = Number of segments, `J` = Mean raw score, `K` = Data quality (%).

---

## 2. Clean with helper columns

Add two columns to the right of the data (`L` and `M` in the workbook, headers in orange) and fill them down to the last row. The formulas below are the ones in row 2.

| Header | Cell | Formula | Meaning |
|---|---|---|---|
| Keep | `Data!L2` | `=IF(K2>=80,"Yes","No")` | Exclusion rule on Data quality |
| Present | `Data!M2` | `=IF(ISNUMBER(H2),IF(H2>0,1,0),"")` | 1 when the expression was present in the interval; blank for NA |

**You should see** `Keep = No` on all 54 rows for each excluded respondent, V07 (61.0 percent) and V19 (72.0 percent); every other respondent is `Yes`. V07 and V19 were recorded in full and post processed; the session log explains each exclusion (glare on glasses lenses for V07; looking down at a phone during the commercial for V19). The same (30) respondents appear in every analysis below.

`ISNUMBER` is the key in the second formula: it is TRUE for a number and FALSE for the text NA or a blank cell, so one formula handles both ways an export can mark a missing metric.

---

## 3. Descriptive table by expression and interval

### 3a. With a PivotTable (the iMotions recommended route)

The Creating Excel Pivot Tables article gives the generic steps: click the first cell of the data, **Insert > PivotTable > OK**, drag a category to **Rows**, drag a metric to **Values**, right click the value, **Value Field Settings**, change **Sum** to **Average**, **OK**, then **PivotChart** if you want a quick picture. For this study:

1. Click cell A1 of the Data sheet, **Insert > PivotTable**, **New Worksheet**, **OK**.
2. Drag **Keep** to **Filters** and set it to **Yes**.
3. Drag **Expression** to **Filters** too and choose **Joy** (or put it in **Rows** to see all six).
4. Drag **Interval label** to **Rows**.
5. Drag **Percent of time present** to **Values** twice: **Average** and **StdDev** (the sample standard deviation, the same as STDEV.S).
6. Drag **Present** to **Values** and set it to **Sum** for the number of viewers with the expression present, or **Average** and multiply by 100 for the percent of viewers.
7. Drag **Number of segments** to **Values** and set it to **Average**.

**You should see** for `Joy`: `Baseline` average 1.61, `Commercial` average 10.37, and `Bin_05` average 24.36.

**Common mistake.** Leaving **Keep** out of the filter. V07 and V19 would then add rows of thin data (many missing frames, so low percentages).

### 3b. With formulas (the Summary sheet)

A formula table updates when the data change and shows exactly what is being averaged. The expression name is in column A and the interval in column B of each row (row 5 is the first). The formulas in row 5 (`Joy`, `Baseline`):

| Statistic | Cell | Formula |
|---|---|---|
| n | `Summary!C5` | `=COUNTIFS(Data!$F$2:$F$1729,$A5,Data!$D$2:$D$1729,$B5,Data!$L$2:$L$1729,"Yes")` |
| Mean time present (%) | `Summary!D5` | `=AVERAGEIFS(Data!$H$2:$H$1729,Data!$F$2:$F$1729,$A5,Data!$D$2:$D$1729,$B5,Data!$L$2:$L$1729,"Yes")` |
| SD (%) | `Summary!E5` | `=STDEV.S(IF((Data!$F$2:$F$1729=$A5)*(Data!$D$2:$D$1729=$B5)*(Data!$L$2:$L$1729="Yes"),Data!$H$2:$H$1729))` |
| Median (%) | `Summary!F5` | `=MEDIAN(IF((Data!$F$2:$F$1729=$A5)*(Data!$D$2:$D$1729=$B5)*(Data!$L$2:$L$1729="Yes"),Data!$H$2:$H$1729))` |
| Viewers present (n) | `Summary!G5` | `=COUNTIFS(Data!$F$2:$F$1729,$A5,Data!$D$2:$D$1729,$B5,Data!$L$2:$L$1729,"Yes",Data!$H$2:$H$1729,">0")` |
| Viewers present (%) | `Summary!H5` | `=100*G5/C5` |
| Mean segments | `Summary!I5` | `=AVERAGEIFS(Data!$I$2:$I$1729,Data!$F$2:$F$1729,$A5,Data!$D$2:$D$1729,$B5,Data!$L$2:$L$1729,"Yes")` |

The two array formulas (`STDEV.S(IF(...))` and `MEDIAN(IF(...))`, columns E and F) are needed because Excel has no STDEVIFS or MEDIANIFS. In Excel 365 type them and press **Enter**. In Excel 2019 or older press **Ctrl+Shift+Enter**; Excel then shows the formula in braces. `AVERAGEIFS` and `COUNTIFS` ignore the text NA.

**You should see** the same numbers as the PivotTable, in rows 5 to 49 (nine rows per expression: the three stimuli, then the six bins) and a column chart of mean Joy by interval under the Engagement table. The **Engagement** table starts in row 53 (`Summary!D53` is the mean raw score for the first interval):

| Statistic | Cell | Formula |
|---|---|---|
| n | `Summary!C53` | `=COUNTIFS(Data!$F$2:$F$1729,$A53,Data!$D$2:$D$1729,$B53,Data!$L$2:$L$1729,"Yes",Data!$J$2:$J$1729,">=0")` |
| Mean raw score | `Summary!D53` | `=AVERAGEIFS(Data!$J$2:$J$1729,Data!$F$2:$F$1729,$A53,Data!$D$2:$D$1729,$B53,Data!$L$2:$L$1729,"Yes")` |
| SD | `Summary!E53` | `=STDEV.S(IF((Data!$F$2:$F$1729=$A53)*(Data!$D$2:$D$1729=$B53)*(Data!$L$2:$L$1729="Yes"),Data!$J$2:$J$1729))` |
| Median | `Summary!F53` | `=MEDIAN(IF((Data!$F$2:$F$1729=$A53)*(Data!$D$2:$D$1729=$B53)*(Data!$L$2:$L$1729="Yes"),Data!$J$2:$J$1729))` |

---

## 4. H1: Joy, Commercial versus Baseline

A paired test needs the two measurements of the same respondent side by side, so first reshape the long export into one row per respondent. The **H1** sheet does this for the 30 kept respondents (listed in column A, rows 5 to 34); each of columns C and D is a `SUMIFS` that picks the one matching Joy row from the Data sheet:

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Data quality (%) | `H1!B5` | `=INDEX(Data!$K$2:$K$1729,MATCH($A5,Data!$B$2:$B$1729,0))` |
| Baseline Joy (%) | `H1!C5` | `=SUMIFS(Data!$H$2:$H$1729,Data!$B$2:$B$1729,$A5,Data!$F$2:$F$1729,"Joy",Data!$D$2:$D$1729,"Baseline")` |
| Commercial Joy (%) | `H1!D5` | `=SUMIFS(Data!$H$2:$H$1729,Data!$B$2:$B$1729,$A5,Data!$F$2:$F$1729,"Joy",Data!$D$2:$D$1729,"Commercial")` |
| Difference (Commercial minus Baseline) | `H1!E5` | `=ROUND(D5-C5,2)` |
| Absolute difference | `H1!F5` | `=ABS(E5)` |
| Rank of absolute difference | `H1!G5` | `=IF(E5=0,"",RANK.AVG(F5,$F$5:$F$34,1)-COUNTIF($E$5:$E$34,0))` |
| Sign of difference | `H1!H5` | `=SIGN(E5)` |
| Tie term (c^2 - 1) | `H1!I5` | `=IF(E5=0,0,COUNTIF($F$5:$F$34,F5)^2-1)` |

(By hand you would get the same table from the PivotTable of section 3a: put **Respondent Name** in Rows, **Interval label** in Columns, **Percent of time present** in Values as Sum, and filter **Expression** to `Joy` and **Keep** to `Yes`.)

The difference is rounded to two decimals (`ROUND`) because the export has two decimals; without it, floating point noise can make two equal differences look unequal, which breaks the tie handling in the rank columns.

Now the test. The statistics sit in column L of the H1 sheet with the formula text shown beside them in column M:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Baseline mean (%) | `H1!L6` | `=AVERAGE(C5:C34)` | 1.61 |
| Baseline SD (%) | `H1!L7` | `=STDEV.S(C5:C34)` | 3.40 |
| Commercial mean (%) | `H1!L8` | `=AVERAGE(D5:D34)` | 10.37 |
| Commercial SD (%) | `H1!L9` | `=STDEV.S(D5:D34)` | 13.57 |
| n pairs | `H1!L10` | `=COUNT(E5:E34)` | 30 |
| Mean difference, Commercial minus Baseline (%) | `H1!L11` | `=AVERAGE(E5:E34)` | 8.76 |
| SD of the differences (%) | `H1!L12` | `=STDEV.S(E5:E34)` | 11.84 |
| **p, two tailed (T.TEST type 1)** | `H1!L16` | `=T.TEST(C5:C34,D5:D34,2,1)` | **0.000346** |

`T.TEST(array1, array2, tails, type)`: `tails = 2` for a two tailed test; `type = 1` for **paired** data, which is the type for this within subjects design. `type = 2` and `type = 3` treat the two columns as different people and give the wrong answer for paired data. `T.TEST` returns only the p value, so compute t and df by hand with the paired formulas (a paired t test is a one sample t test on the differences):

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| SE of the mean difference | `H1!L13` | `=L12/SQRT(L10)` | 2.161 |
| Paired t | `H1!L14` | `=L11/L13` | 4.05 |
| df | `H1!L15` | `=L10-1` | 29 |
| p from the hand computed t | `H1!L17` | `=T.DIST.2T(ABS(L14),L15)` | 0.000346, matches T.TEST |
| Critical t (95%) | `H1!L18` | `=T.INV.2T(0.05,L15)` | 2.045 |
| Mean difference 95% CI lower | `H1!L19` | `=L11-L18*L13` | 4.34 |
| Mean difference 95% CI upper | `H1!L20` | `=L11+L18*L13` | 13.18 |

**You should see** the hand computed p equal the `T.TEST` p. If they differ, the `type` argument or a column range is wrong.

### Cohen's d_z by hand, with a 95 percent confidence interval

For paired data the effect size is **d_z**: the mean difference in units of the SD of the *differences*. It is linked to t by t = d_z x sqrt(n), so t, n, and d_z always agree.

| Step | Cell | Formula | Result |
|---|---|---|---|
| Cohen's d_z | `H1!L21` | `=L11/L12` | 0.74 |
| SE of d_z (approximation) | `H1!L22` | `=SQRT(1/L10+L21^2/(2*L10))` | 0.206 |
| d_z 95% CI lower | `H1!L23` | `=L21-1.96*L22` | 0.34 |
| d_z 95% CI upper | `H1!L24` | `=L21+1.96*L22` | 1.14 |
| Correlation of the two columns | `H1!L25` | `=CORREL(C5:C34,D5:D34)` | 0.60 |

In words: d_z = mean difference / SD of differences; SE(d_z) = sqrt(1/n + d_z^2 / (2 n)). The interval is the usual large sample approximation. `analysis.R` reports the interval from the noncentral t distribution (effectsize package), which is more exact at this sample size; d_z itself matches, and the limits differ slightly ([0.33, 1.14] in R). Say which method you used. Rule of thumb for reading d_z: 0.2 small, 0.5 medium, 0.8 large.

### Viewers with any Joy, and the exact McNemar test

Because most Baseline values are 0, the count of viewers who showed any Joy is part of the result.

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| Viewers with Joy, Baseline | `H1!L28` | `=COUNTIF(C5:C34,">0")` | 8 |
| Viewers with Joy, Commercial | `H1!L29` | `=COUNTIF(D5:D34,">0")` | 25 |
| Percent of viewers with Joy, Baseline | `H1!L30` | `=100*L28/L10` | 26.7 |
| Percent of viewers with Joy, Commercial | `H1!L31` | `=100*L29/L10` | 83.3 |
| Joy at Baseline only | `H1!L32` | `=SUMPRODUCT((C5:C34>0)*(D5:D34=0))` | 1 |
| Joy in the Commercial only | `H1!L33` | `=SUMPRODUCT((C5:C34=0)*(D5:D34>0))` | 18 |
| Exact McNemar p, two tailed | `H1!L34` | `=MIN(1,2*BINOM.DIST(MIN(L32,L33),L32+L33,0.5,TRUE))` | 7.63e-05 |

The McNemar test looks only at the viewers whose status changed between the two periods (Joy in one, none in the other) and asks whether the changes favor one direction; with the exact binomial the p value is `MIN(1, 2 x BINOM.DIST(smaller count, total changed, 0.5, TRUE))`.

### The Wilcoxon signed rank test by hand (the primary test)

Excel has no signed rank function, but the test is a short calculation on the difference column. The steps: drop the zero differences; rank the absolute differences from smallest to largest (tied values share the average rank, `RANK.AVG`); add up the ranks of the positive differences (that sum is the V that R prints); compare V with its expected value under no difference, using the normal approximation with the tie and continuity corrections that R applies when there are ties (`wilcox.test(exact = FALSE)`).

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| Differences above 0 | `H1!L37` | `=COUNTIF(E5:E34,">0")` | 23 |
| Differences below 0 | `H1!L38` | `=COUNTIF(E5:E34,"<0")` | 3 |
| Differences equal to 0 | `H1!L39` | `=COUNTIF(E5:E34,0)` | 4 |
| Nonzero differences | `H1!L40` | `=L37+L38` | 26 |
| V = sum of ranks of positive differences | `H1!L41` | `=SUMIFS(G5:G34,H5:H34,1)` | 331.0 |
| Sum of ranks of negative differences | `H1!L42` | `=SUMIFS(G5:G34,H5:H34,-1)` | 20.0 |
| V plus negative rank sum minus n(n+1)/2 | `H1!L43` | `=L41+L42-L40*(L40+1)/2` | 0.0 |
| Expected V under H0 | `H1!L44` | `=L40*(L40+1)/4` | 175.50 |
| Sum of (t^3 - t) over tie groups | `H1!L45` | `=SUM(I5:I34)` | 0 |
| SD of V under H0 | `H1!L46` | `=SQRT(L40*(L40+1)*(2*L40+1)/24-L45/48)` | 39.373 |
| z with continuity correction | `H1!L47` | `=(L41-L44-0.5*SIGN(L41-L44))/L46` | 3.937 |
| **Wilcoxon p, two tailed** | `H1!L48` | `=2*(1-NORM.S.DIST(ABS(L47),TRUE))` | **8.26e-05** |
| Matched pairs rank biserial correlation | `H1!L49` | `=(L41-L42)/(L41+L42)` | 0.89 |
| Sign test p, two tailed | `H1!L50` | `=MIN(1,2*BINOM.DIST(MIN(L37,L38),L37+L38,0.5,TRUE))` | 0.000088 |

**You should see** V equal to 331 and the p value equal to the one in `analysis_output.txt` (8.26e-05). The rank biserial correlation is the effect size that goes with the signed rank test: +1 means every nonzero difference is positive and large relative to the negative ones. The sign test uses only the direction of each difference and is the simplest check of all.

The two check rows at the bottom of the block:

| Statistic | Cell | Formula |
|---|---|---|
| Kept respondents with a Joy Baseline row in Data | `H1!L53` | `=COUNTIFS(Data!$F$2:$F$1729,"Joy",Data!$D$2:$D$1729,"Baseline",Data!$L$2:$L$1729,"Yes")` |
| Smallest Data quality in column B | `H1!L54` | `=MIN(B5:B34)` |

**You should see** 30 in the first (it must equal n pairs) and a number of at least 80 in the second.

---

## 5. H2: Joy across the six bins

The **Bins** sheet lists the 30 kept respondents (rows 5 to 34) with one column per bin (`B` to `G`), each a `SUMIFS` on the Joy rows of the Data sheet; the bin names in row 4 are the lookup keys, so do not rename them.

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Bin_01 Joy (%) | `Bins!B5` | `=SUMIFS(Data!$H$2:$H$1729,Data!$B$2:$B$1729,$A5,Data!$F$2:$F$1729,"Joy",Data!$D$2:$D$1729,B$4)` |
| Commercial Joy (%) | `Bins!H5` | `=SUMIFS(Data!$H$2:$H$1729,Data!$B$2:$B$1729,$A5,Data!$F$2:$F$1729,"Joy",Data!$D$2:$D$1729,"Commercial")` |
| Respondent mean over the bins | `Bins!I5` | `=AVERAGE(B5:G5)` |
| Rank of Bin_01 | `Bins!J5` | `=RANK.AVG(B5,$B5:$G5,1)` |
| Tie term | `Bins!P5` | `=SUMPRODUCT(COUNTIF($B5:$G5,$B5:$G5)^2-1)` |

### The proportion of viewers with Joy present per bin

This is the quantity the FEA thresholding **Aggregation** shows in the Aggregate Replay (how many respondents showed an expression in a time window). Rows 37 to 44 under the roster hold, for each bin, the count of viewers, the count with Joy present (`COUNTIFS` with the criterion `">0"` on Percent of time present), the proportion, and the mean. The formulas in column B (`Bin_01`):

| Statistic | Cell | Formula | Results across the bins (simulated) |
|---|---|---|---|
| Viewers (n) | `Bins!B37` | `=COUNTIFS(Data!$F$2:$F$1729,"Joy",Data!$D$2:$D$1729,B$4,Data!$L$2:$L$1729,"Yes")` | 30 |
| Viewers with Joy present | `Bins!B38` | `=COUNTIFS(Data!$F$2:$F$1729,"Joy",Data!$D$2:$D$1729,B$4,Data!$L$2:$L$1729,"Yes",Data!$H$2:$H$1729,">0")` | 7, 8, 7, 9, 20, 16 |
| Proportion with Joy present | `Bins!B39` | `=B38/B37` | 0.233, 0.267, 0.233, 0.300, 0.667, 0.533 |
| Percent of viewers with Joy present | `Bins!B40` | `=100*B39` | 23.3, 26.7, 23.3, 30.0, 66.7, 53.3 |
| Mean Joy (%) | `Bins!B41` | `=AVERAGE(B5:B34)` | 4.83, 7.73, 6.42, 5.41, 24.36, 13.43 |
| SD (%) | `Bins!B42` | `=STDEV.S(B5:B34)` | 11.59, 15.14, 17.18, 11.84, 28.69, 21.86 |
| 95% CI half width | `Bins!B43` | `=CONFIDENCE.T(0.05,B42,B37)` | 4.33, 5.65, 6.41, 4.42, 10.71, 8.16 |
| Friedman rank sum | `Bins!B44` | `=SUM(J5:J34)` | 89.0, 97.5, 91.0, 94.0, 144.0, 114.5 |

**You should see** the proportion and the mean both highest in `Bin_05` (column F) and next highest in `Bin_06` (column G).

### Repeated measures ANOVA by hand

`analysis.R` fits `pct ~ bin + (1 | respondent)` with lme4 and lmerTest. With every respondent measured in every bin (a balanced design) and a positive respondent variance, that model's F test for the bin factor is the classic repeated measures ANOVA, which Excel can compute from sums of squares. The block sits in column S (formula text in column T):

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Respondents (n) | `Bins!S6` | `=COUNT(I5:I34)` | 30 |
| Bins (k) | `Bins!S7` | `=COUNTA(B4:G4)` | 6 |
| SS total | `Bins!S9` | `=DEVSQ(B5:G34)` | 69377.6 |
| SS bins | `Bins!S10` | `=S6*DEVSQ(B41:G41)` | 8483.3 |
| SS respondents | `Bins!S11` | `=S7*DEVSQ(I5:I34)` | 32034.0 |
| SS error | `Bins!S12` | `=S9-S10-S11` | 28860.4 |
| df error | `Bins!S14` | `=(S6-1)*(S7-1)` | 145 |
| MS bins | `Bins!S15` | `=S10/S13` | 1696.65 |
| MS error | `Bins!S16` | `=S12/S14` | 199.04 |
| **F** | `Bins!S17` | `=S15/S16` | **8.52** |
| p (F.DIST.RT) | `Bins!S18` | `=F.DIST.RT(S17,S13,S14)` | 4.25e-07 |
| Respondent variance (ANOVA estimate) | `Bins!S19` | `=(((S11)/(S6-1))-S16)/S7` | 150.93 |
| ICC (adjusted) | `Bins!S20` | `=S19/(S19+S16)` | 0.431 |

`DEVSQ` returns the sum of squared deviations from the mean. The error term uses (n - 1)(k - 1) = 145 degrees of freedom, which is also the Satterthwaite denominator df that lmerTest reports for this balanced design. The respondent variance and ICC are the ANOVA estimates, which equal the REML estimates of the mixed model here.

### Friedman test by hand (the nonparametric check)

Joy is zero inflated and right skewed, so the F test's normality assumption is doubtful (see the assumption plots from `analysis.R`). The Friedman test ranks each respondent's six bin values from 1 to 6 (ties share an average rank) and asks whether the bins differ in their rank sums. The tie correction matters here because many bins hold 0.

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| 12 x sum of squared deviations of the rank sums | `Bins!S23` | `=12*SUMPRODUCT((B44:G44-S6*(S7+1)/2)^2)` | 26886.00 |
| Denominator with tie correction | `Bins!S24` | `=S6*S7*(S7+1)-SUM(P5:P34)/(S7-1)` | 775.20 |
| **Friedman chi-squared** | `Bins!S25` | `=S23/S24` | **34.68** |
| p (CHISQ.DIST.RT) | `Bins!S26` | `=CHISQ.DIST.RT(S25,S7-1)` | 1.74e-06 |
| Kendall's W | `Bins!S27` | `=S25/(S6*(S7-1))` | 0.231 |

Kendall's W is the effect size for the Friedman test (0 = no agreement among respondents about the order of the bins, 1 = complete agreement).

### Contrasts: where is the Joy?

An omnibus test says the bins differ, not where. Two descriptive contrasts compare a late bin with the mean of `Bin_01` to `Bin_04`; the standard error uses the error mean square and the squared weights (1 and four times -1/4 give 1.25):

| Statistic | Cell | Formula |
|---|---|---|
| Bin_05 minus mean(Bin_01..Bin_04) (percentage points) | `Bins!S30` | `=F41-AVERAGE(B41:E41)` |
| SE of the contrast | `Bins!S31` | `=SQRT(S16*1.25/S6)` |
| t | `Bins!S32` | `=S30/S31` |
| p, two tailed | `Bins!S33` | `=T.DIST.2T(ABS(S32),S14)` |
| Bin_06 minus mean(Bin_01..Bin_04) (percentage points) | `Bins!S34` | `=G41-AVERAGE(B41:E41)` |
| SE of the contrast | `Bins!S35` | `=SQRT(S16*1.25/S6)` |
| t | `Bins!S36` | `=S34/S35` |
| p, two tailed | `Bins!S37` | `=T.DIST.2T(ABS(S36),S14)` |

R reports the same contrasts with Satterthwaite degrees of freedom and a Holm adjustment for the two contrasts; the unadjusted Excel p values match the unadjusted R p values.

The two check rows at the bottom of the block:

| Statistic | Cell | Formula |
|---|---|---|
| Largest gap between Commercial Joy and the mean of the six bins (percentage points) | `Bins!S40` | `=SUMPRODUCT(MAX(ABS(H5:H34-I5:I34)))` |
| Rank sums add up to n k (k + 1) / 2 | `Bins!S41` | `=SUM(B44:G44)-S6*S7*(S7+1)/2` |

**You should see** a gap of at most 0.01 percentage points (the bins tile the commercial, so the stimulus value is the mean of the bin values, up to rounding) and 0 for the rank sum check.

---

## 6. H3: Positive Valence and ad liking

The **H3** sheet holds `SIMULATED_fea_survey.csv` (one row per respondent, rows 4 to 35), with the kept respondents first (rows 4 to 33) and the excluded respondents last (rows 34 to 35), so the test ranges below leave them out. Column G flags Keep from the Data sheet, column H pulls Positive Valence `Percent of time present` during the commercial, and columns I and J rank the two variables for the Spearman correlation:

| Column | Cell (first row) | Formula |
|---|---|---|
| Keep | `H3!G4` | `=IF(INDEX(Data!$K$2:$K$1729,MATCH($A4,Data!$B$2:$B$1729,0))>=80,"Yes","No")` |
| Positive Valence during the Commercial (%) | `H3!H4` | `=SUMIFS(Data!$H$2:$H$1729,Data!$B$2:$B$1729,$A4,Data!$F$2:$F$1729,"Positive Valence",Data!$D$2:$D$1729,"Commercial")` |
| Rank of Ad_liking_1to7 | `H3!I4` | `=RANK.AVG(D4,$D$4:$D$33,1)` |
| Rank of Positive Valence (%) | `H3!J4` | `=RANK.AVG(H4,$H$4:$H$33,1)` |

Positive Valence is the Valence channel above +50. The statistics sit in column M (formula text in column N):

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| n (kept respondents) | `H3!M5` | `=COUNT(H4:H33)` | 30 |
| Positive Valence mean (%) | `H3!M6` | `=AVERAGE(H4:H33)` | 11.44 |
| **Pearson r (CORREL)** | `H3!M12` | `=CORREL(H4:H33,D4:D33)` | **0.45** |
| t for r | `H3!M13` | `=M12*SQRT((M5-2)/(1-M12^2))` | 2.70 |
| df | `H3!M14` | `=M5-2` | 28 |
| p, two tailed | `H3!M15` | `=T.DIST.2T(ABS(M13),M14)` | 0.011560 |
| r 95% CI lower (Fisher z) | `H3!M16` | `=TANH(ATANH(M12)-NORM.S.INV(0.975)/SQRT(M5-3))` | 0.11 |
| r 95% CI upper (Fisher z) | `H3!M17` | `=TANH(ATANH(M12)+NORM.S.INV(0.975)/SQRT(M5-3))` | 0.70 |
| Least squares slope (percentage points per liking point) | `H3!M18` | `=SLOPE(H4:H33,D4:D33)` | 4.98 |
| **Spearman rho** | `H3!M23` | `=CORREL(I4:I33,J4:J33)` | **0.41** |
| t for rho | `H3!M24` | `=M23*SQRT((M5-2)/(1-M23^2))` | 2.41 |
| p, two tailed | `H3!M25` | `=T.DIST.2T(ABS(M24),M5-2)` | 0.022730 |
| rho 95% CI lower (Bonett and Wright) | `H3!M26` | `=TANH(ATANH(M23)-NORM.S.INV(0.975)*SQRT((1+M23^2/2)/(M5-3)))` | 0.05 |
| rho 95% CI upper (Bonett and Wright) | `H3!M27` | `=TANH(ATANH(M23)+NORM.S.INV(0.975)*SQRT((1+M23^2/2)/(M5-3)))` | 0.68 |

A Pearson correlation uses the raw values, so a few viewers with a lot of Positive Valence can pull it; Spearman uses ranks (`RANK.AVG` gives tied values the average rank, and ratings on a 7 point scale tie often), so it is the check against that. `Ad_liking_1to7` is a whole number rating, and it is correlated with a time percentage: a correlation says the two move together, not that one causes the other, and AFFDEX scores show what the face expressed, not what the viewer felt.

---

## 7. Chart with 95 percent confidence intervals

1. Make a small chart table: two rows (Baseline, Commercial), columns **Mean Joy (%)** (point to the mean cells) and **CI half width** = `CONFIDENCE.T(0.05, SD cell, n cell)`. `CONFIDENCE.T` returns half the width of a t based 95 percent interval around a mean. The workbook has this table on the H1 sheet at `H1!K59:M60`; the half widths are in `H1!M59` and `H1!M60`.
2. Select the interval names and the means, **Insert > Column chart > Clustered Column**.
3. Click the bars, **Chart Design > Add Chart Element > Error Bars > More Error Bars Options**, choose **Custom > Specify Value**, and select the two CI half width cells for both **Positive** and **Negative**.
4. Set the bar fill to navy (`#1F3864`) and lighten the gridlines.

**Common mistake.** Using the default error bars (**Standard Error** or a fixed 5 percent). Those are not 95 percent confidence intervals. And remember that error bars around the two means are *not* the interval that goes with the paired test: the paired test uses the interval of the *differences* (`H1!L19` and `H1!L20`), which is much narrower because it removes the between respondent variation. With zero inflated data, a bar chart of means also hides the shape: show the individual values (the R figure does).

---

## 8. Check your numbers and know what Excel cannot do

The **Check** sheet puts each live Excel result next to the value `analysis.R` produced and prints OK or CHECK (status in column F, the count of CHECK rows in `Check!B57`). Your Excel results should match to rounding; the d_z limits differ by method and have a looser tolerance. Several parts of the R analysis have no Excel equivalent:

- **Mixed model estimates** (random intercept variance as fitted by REML, Satterthwaite degrees of freedom for unbalanced or missing data, the model on the square root of the percent).
- **Hodges-Lehmann estimate** and its interval for the median difference.
- **Exact Clopper-Pearson intervals** for the proportions of viewers per bin.
- **The sensitivity analysis at thresholds of 25 and 75.** In iMotions you rerun the thresholding notebook with a batch threshold of 25 and then 75 and export the FEA Metrics Export again. In Excel, import `SIMULATED_fea_metrics_threshold25.csv` into a copy of the Data sheet (same columns, same helper formulas) and the H1 sheet recalculates for that threshold. Expect the mean Joy percent to rise at 25 and fall at 75, because a lower threshold counts weaker expressions, and longer stretches of them, as present.

If you have the **Analysis ToolPak** add in (**File > Options > Add-ins > Manage Excel Add-ins > Go > Analysis ToolPak**), **Data > Data Analysis > t-Test: Paired Two Sample for Means** prints t, df, and p in one table and is a good cross check of section 4; **Anova: Two-Factor Without Replication** gives the repeated measures table of section 5 (rows = respondents, columns = bins).

---

## 9. Common mistakes, collected

1. Opening the .csv by double clicking and getting one column: use **Data > From Text/CSV** or the Text Import Wizard.
2. Replacing NA with 0. NA means no value exists (for example, the thresholded columns on Engagement rows).
3. Forgetting the exclusion filter (Keep = Yes) in the PivotTable or in COUNTIFS, or leaving V07 and V19 in a roster.
4. `T.TEST` with `type = 2` or `type = 3` on a within subjects design. Use `type = 1` and make sure the two columns are in the same respondent order.
5. Reporting the paired t test alone for a zero inflated outcome. Report the signed rank test, the count of viewers with any Joy, and d_z with its interval.
6. Treating 1728 rows as 1728 viewers. Each respondent contributes 54 rows; tests use one value per respondent and interval.
7. Adding up bin percentages as if they were separate observations of 60 s: each bin is 10 s of the 60 s commercial, so the commercial value is the **mean** of the six bins, not the sum.
8. Reading the Baseline as a counterbalanced control. Everyone saw the Baseline first, so any difference between Baseline and Commercial also includes time in the chair.
9. Reading `Percent of time present` as the strength of an emotion. FEA reports whether the face showed an expression above the threshold, not how intensely the viewer felt it.

## Sources

- iMotions help center, R Notebook (FEA): Facial expression analysis thresholding and aggregation: https://help.imotions.com/lab/analysis/facial-expression-analysis/r-notebook-fea-facial-expression-analysis-thresholding-and-aggregation (local copy: `C:\JaxState_Behavior Lab\iMotion_Manual\R Notebook FEA Facial expression analysis thresholding and aggregation.pdf`)
- iMotions help center, Facial Expression Analysis: Importing and Post Processing Videos: https://help.imotions.com/lab/import-and-post-processing/facial-expression-analysis-importing-and-post-processing-videos (local copy: `C:\JaxState_Behavior Lab\iMotion_Manual\Facial Expression Analysis Post Processing face recordings.pdf`)
- iMotions help center, Affectiva (AFFDEX) Signals Overview: https://help.imotions.com/lab/analysis/facial-expression-analysis/affdex/affectiva-affdex-signals-overview (local copy: `C:\JaxState_Behavior Lab\iMotion_Manual\AFFDEX SDK 52 Signals.pdf`)
- iMotions help center, R Notebooks (FEA): AFFDEX Raw Data Statistics Export: https://help.imotions.com/lab/analysis/facial-expression-analysis/affdex/r-notebooks-fea-affdex-raw-data-statistics-export (local copy: `C:\JaxState_Behavior Lab\iMotion_Manual\R Notebooks FEA AFFDEX Raw Data Statistics Export.pdf`)
- iMotions help center, Facial Expression Analysis: Why does AFFDEX show low QA in subject replay?: https://help.imotions.com/lab/analysis/facial-expression-analysis/affdex/facial-expression-analysis-why-does-affdex-show-low-qa-in-subject-replay (local copy: `C:\JaxState_Behavior Lab\iMotion_Manual\Facial Expression Analysis Why does AFFDEX show low QA in subject replay.pdf`)
- iMotions help center, Demo study: Sony Bravia (Testing of video-commercials) - 9.0 (binning annotation): https://help.imotions.com/lab/getting-started/demo-studies/9-0-demo-studies/demo-study-sony-bravia-testing-of-video-commercials-9-0 (local copy: `C:\JaxState_Behavior Lab\iMotion_Manual\1Printed\Demo study Sony Bravia Testing of videocommercials  90.pdf`)
- iMotions help center, Opening Data Export Files in Excel: https://help.imotions.com/lab/exports/additional-export-information/opening-data-export-files-in-excel (local copy: `iMotion_Manual\Opening Data Export Files in Excel.pdf`)
- iMotions help center, Creating Excel Pivot Tables: https://help.imotions.com/lab/exports/additional-export-information/creating-excel-pivot-tables (local copy: `iMotion_Manual\Making Excel Pivot Tables.pdf`)

## Items to verify on the lab computer

- The exact column header layout, file name, and file format of the FEA Metrics Export in iMotions Lab 11.1 (the manuals describe the content but print no header row; one article calls R exports .txt and others .csv). Adjust the column letters in sections 1 and 2 once a real export is open.
- Whether the real export writes `NA` or leaves cells blank where a metric does not apply; `ISNUMBER` handles both.
- Whether the denominator of the percent of time present is the whole interval or only the frames with a face found, and how a segment that crosses a bin boundary is counted (the simulated file divides by the whole interval and counts the segment in both bins).
- Whether the thresholding notebook offers Engagement as a thresholded expression (the article lists Emotion, Valence, Action Units, and head position channels but does not name Engagement); until then Engagement stays a raw mean.
- What the thresholding Aggregation counts as "showed the expression" in a time window, and its default bin length and overlap (the manual states neither); the workbook uses Percent of time present above 0 in the 10 s bin.
- The default for the short or long classification time threshold (not stated), and the per channel threshold dialog.
- The 80 percent data quality exclusion rule is the lab's; no manual states a minimum. Decide the rule after viewing Quality Metrics on pilot data.
- Which AFFDEX version is installed in 11.1.13 and whether the rolling baseline setting changes the numbers.
