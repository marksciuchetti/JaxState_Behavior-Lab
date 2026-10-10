# Ad test analysis in Excel: step by step walkthrough

**SIMULATED DATA.** Every number in this walkthrough comes from `SIMULATED_ad_test_AOI_metrics.csv`, which `make_simulated_data.py` generated (seed 20268242). No real viewer was recorded. The design follows the structure of the iMotions demonstration study "DEMO - SonyBravia" (one brand, two television commercials, an eye tracker on a monitor); no data from that study were used.

This is the same analysis as `analysis.R`, done in Excel for students who do not use R. The finished workbook is `SIMULATED_ad_test_analysis.xlsx`; the sheet names below refer to it. You can follow along in that workbook or rebuild it from the CSV files. **Every cell address in this walkthrough is written as `Sheet!Cell` and was generated from the workbook itself** (`make_excel_workbook.py` writes both), so the addresses match the file.

Companion files: `analysis.R` (R version), `analysis_output.txt` (its console output), `results_text.md` (model paragraphs), `figures/` (PNG figures).

---

## 0. What you are analyzing

**Design.** Within subjects. Each respondent watched both television commercials for the same brand, `Commercial_1_Balls` and `Commercial_2_Paint` (60 s each in the simulation), on the lab monitor with the Smart Eye Aurora (60 Hz). A Baseline slide came before each commercial and a rating survey slide after each. The order was counterbalanced: Group A saw Balls first, Group B saw Paint first. 30 respondents, R01 to R30, 15 per order. In Workshop 1 the groups were different people (a Welch t test); here the two commercials are measured on the *same* people, so the analysis is a **paired** test and the order of viewing must be checked.

**Export.** One row per respondent by stimulus by AOI, the layout of the iMotions **Individual AOI metrics** export, which iMotions recommends for statistics because it keeps each respondent's values instead of averaging across the segment (Areas of Interest article, Export options). There are 240 data rows: 30 respondents x 2 commercials x 4 AOIs. The columns after the identifiers are iMotions AOI metrics with the names from the **AOI Metrics** help article. The ones used here:

| Column | iMotions definition (AOI Metrics article) | Used for |
|---|---|---|
| `Valid data` | Percentage of collected and interpolated samples; below 100 means missing data | Exclusion rule |
| `AOI duration (ms)` and `AOI duration (%)` | How long the (dynamic) AOI was active, and that time as a percentage of the stimulus duration | Interpreting Dwell time (%) |
| `TTFF AOI (ms)` | Time from the AOI's start until the first fixation on the AOI | H2 (Brand_Logo) |
| `TTFF parent (ms)` | The same, counted from the start of the AOI's parent (here the commercial) | Not used (see below) |
| `TTFF max. (ms)` | Counted from the start of the parent; a respondent who never fixated the AOI gets the parent's full presentation duration (60000 ms here) | Not used (see below) |
| `Dwell time (ms)` | Total time fixating on the AOI | Brand_Logo total dwell and recall |
| `Dwell time (%)` | Dwell time relative to the time during which the AOI was active | H1 (Product_TV) |
| `Fixation count` | Number of fixations inside the AOI | Descriptives |

Where the same metric name exists as a gaze based and a fixation based metric (Dwell time, Revisit count), the gaze based column carries the prefix `Gaze`. iMotions writes **NA** when a metric cannot be calculated, for example when the respondent never fixated the AOI. The simulated file does the same.

**Dynamic AOIs.** An AOI that follows an element on screen is active only while the element is shown (Areas of Interest article, dynamic AOIs). Product_TV (the television set in the product shots) is active for 14 s of the 60 s in Balls and 20 s in Paint; Brand_Logo for the last 5 s; Tagline_Text for the last 4 s; Hero_Element (the bouncing balls or the paint explosion) for 40 s in Balls and 32 s in Paint. `Dwell time (%)` divides by this active time, which is why it can be compared across the two commercials even though the AOIs are on screen for different lengths.

**Hypotheses.**
- H1: the share of dwell time on `Product_TV` is predicted to be higher in Paint than in Balls (the simulation built in a larger share for Paint, which shows the set for 20 s against 14 s); the test is two tailed.
- H2: `TTFF AOI (ms)` on `Brand_Logo` (from the start of the end card) is predicted to be shorter in Paint than in Balls; the test is two tailed.
- H3: liking (1 to 7) differs between the commercials.
- Order effect: the commercial seen second gets extra Product_TV dwell share (carryover from the first viewing).

**Decision rules fixed before looking at the data.** Exclude a respondent whose `Valid data` is below 70 percent. For an AOI the respondent never fixated, TTFF is *not observed* (leave it out of TTFF statistics) and dwell time is *0* (keep it in dwell statistics). Every difference is **Paint minus Balls**. Two tailed tests, alpha = .05.

---

## 1. Import the export into Excel

iMotions exports are comma separated with a dot decimal. Double clicking the file uses your regional defaults, which can put everything in one column (Opening Data Export Files in Excel article). Use the import route instead.

**Route A, Get Data (Excel 2016 and later):**
1. **Data > From Text/CSV**, select `SIMULATED_ad_test_AOI_metrics.csv`, click **Import**.
2. In the preview set **Delimiter: Comma** and **File Origin: 65001 Unicode (UTF-8)**. **You should see** 25 columns with `Study Name` in the first column.
3. Click **Transform Data** if the first row shows the `# SIMULATED ...` comment line: in Power Query choose **Home > Remove Rows > Remove Top Rows > 1**, then **Home > Use First Row as Headers**, then **Close & Load**. (A real iMotions export has no comment line; the simulated file carries one so nobody mistakes it for real data.)
4. **You should see** a table with 240 data rows.

**Route B, Text Import Wizard (any version):** **File > Open > Browse**, set the file type to **All Files**, open the .csv, choose **Delimited**, start at row 2 (to skip the comment line), tick **Comma**, click **Advanced** and set the decimal symbol to a dot, then **Finish**.

**Common mistake.** The NA cells must stay as the text `NA` (or be blank). Do not replace NA with 0 in a TTFF column: a zero TTFF would mean the respondent fixated the logo at the very first instant, which is false. Step 2 handles NA correctly.

Save as **.xlsx** right away; a .csv does not keep formulas or extra sheets. Keep the original .csv read only.

In the workbook this is the **Data** sheet, header in row 1 and data in rows 2 to 241. Column letters there: `B` = Respondent Name, `E` = Group (the order), `G` = Stimulus, `H` = AOI Name, `L` = Valid data, `R` = Fixation count, `S` = TTFF AOI (ms), `V` = Dwell time (ms), `W` = Dwell time (%).

---

## 2. Clean with helper columns

Add five columns to the right of the data (`Z` to `AD` in the workbook, headers in orange) and fill them down to the last row. The formulas below are the ones in row 2.

| Header | Cell | Formula | Meaning |
|---|---|---|---|
| Keep | `Data!Z2` | `=IF(L2>=70,"Yes","No")` | Exclusion rule on Valid data |
| Position | `Data!AA2` | `=IF(OR(AND(E2="A",G2="Commercial_1_Balls"),AND(E2="B",G2="Commercial_2_Paint")),"first","second")` | Was this commercial the first or the second one this respondent saw? Group A saw Balls first, Group B saw Paint first |
| Dwell_pct_clean | `Data!AB2` | `=IF(ISNUMBER(W2),W2,0)` | Never fixated AOI counts as 0 percent |
| Fixated | `Data!AC2` | `=IF(ISNUMBER(S2),1,0)` | 1 when at least one fixation hit the AOI |
| Fix_count_clean | `Data!AD2` | `=IF(ISNUMBER(R2),R2,0)` | Fixation count with NA counted as 0 |

**You should see** `Keep = No` on all eight rows of respondent R23 (Valid data 58 percent; the session log says calibration stayed Poor after two recalibrations and the session was recorded anyway). Every other respondent is `Yes`. Respondents R09 and R17 needed one recalibration (Poor, then Good) but have complete data and stay in; a recalibration is not by itself a reason to exclude.

`ISNUMBER` is the key: it is TRUE for a number and FALSE for the text NA or a blank cell, so one formula handles both ways an export can mark a missing metric.

---

## 3. Descriptive table by commercial and AOI

### 3a. With a PivotTable (the iMotions recommended route)

The Creating Excel Pivot Tables article gives the generic steps: click the first cell of the data, **Insert > PivotTable > OK**, drag a category to **Rows**, drag a metric to **Values**, right click the value, **Value Field Settings**, change **Sum** to **Average**, **OK**, then **PivotChart** if you want a quick picture. For this study:

1. Click cell A1 of the Data sheet, **Insert > PivotTable**, **New Worksheet**, **OK**.
2. Drag **Keep** to **Filters** and set it to **Yes**.
3. Drag **AOI Name** to **Rows** and **Stimulus** to **Columns**.
4. Drag **TTFF AOI (ms)** to **Values** three times. Set the first to **Average**, the second to **StdDev** (the sample standard deviation, the same as STDEV.S), the third to **Count**. The text NA is ignored by Average and StdDev, but **Count** counts every non blank cell including the text NA; use **Count Numbers** (Value Field Settings > Summarize Values By > Count Numbers) for the number of fixators, or the formula table in 3b.
5. Drag **Dwell_pct_clean** to **Values** twice: **Average** and **StdDev**.
6. Drag **Fixated** to **Values** and set it to **Average**; multiply by 100 beside the table. This is the fixation based **Respondent ratio (%)** from the iMotions AOI table.

**You should see** `Product_TV`: Balls average Dwell_pct_clean 21.24 and Paint 28.49; `Brand_Logo`: respondent ratio 93.1 percent for Balls and 93.1 percent for Paint.

**Common mistake.** Leaving **Keep** out of the filter. R23 would then add eight rows of thin data.

### 3b. With formulas (the Summary sheet)

A formula table updates when the data change and shows exactly what is being averaged. The AOI name is in column A and the stimulus in column B of each row (row 5 is the first). The formulas in row 5 (`Product_TV`, `Commercial_1_Balls`):

| Statistic | Cell | Formula |
|---|---|---|
| n | `Summary!C5` | `=COUNTIFS(Data!$G$2:$G$241,$B5,Data!$H$2:$H$241,$A5,Data!$Z$2:$Z$241,"Yes")` |
| Respondent ratio (%) | `Summary!D5` | `=100*AVERAGEIFS(Data!$AC$2:$AC$241,Data!$G$2:$G$241,$B5,Data!$H$2:$H$241,$A5,Data!$Z$2:$Z$241,"Yes")` |
| n with TTFF | `Summary!E5` | `=COUNTIFS(Data!$G$2:$G$241,$B5,Data!$H$2:$H$241,$A5,Data!$Z$2:$Z$241,"Yes",Data!$S$2:$S$241,">0")` |
| TTFF AOI mean (ms) | `Summary!F5` | `=AVERAGEIFS(Data!$S$2:$S$241,Data!$G$2:$G$241,$B5,Data!$H$2:$H$241,$A5,Data!$Z$2:$Z$241,"Yes")` |
| TTFF AOI SD (ms) | `Summary!G5` | `=STDEV.S(IF((Data!$G$2:$G$241=$B5)*(Data!$H$2:$H$241=$A5)*(Data!$Z$2:$Z$241="Yes")*ISNUMBER(Data!$S$2:$S$241),Data!$S$2:$S$241))` |
| Dwell time mean (%) | `Summary!H5` | `=AVERAGEIFS(Data!$AB$2:$AB$241,Data!$G$2:$G$241,$B5,Data!$H$2:$H$241,$A5,Data!$Z$2:$Z$241,"Yes")` |
| Dwell time SD (%) | `Summary!I5` | `=STDEV.S(IF((Data!$G$2:$G$241=$B5)*(Data!$H$2:$H$241=$A5)*(Data!$Z$2:$Z$241="Yes"),Data!$AB$2:$AB$241))` |
| Fixation count mean | `Summary!J5` | `=AVERAGEIFS(Data!$AD$2:$AD$241,Data!$G$2:$G$241,$B5,Data!$H$2:$H$241,$A5,Data!$Z$2:$Z$241,"Yes")` |

The two `STDEV.S(IF(...))` formulas (columns G and I) are **array formulas** because Excel has no STDEVIFS. In Excel 365 type them and press **Enter**. In Excel 2019 or older press **Ctrl+Shift+Enter**; Excel then shows the formula in braces. `AVERAGEIFS` and `COUNTIFS` ignore the text NA, which is exactly the "TTFF not observed" rule.

**You should see** the same numbers as the PivotTable, in rows 5 to 12 (two rows per AOI), and a clustered bar chart of mean dwell share by AOI and commercial under the table.

---

## 4. H1: Product_TV dwell share, paired t test

A paired test needs the two measurements of the same respondent side by side, so first reshape the long export into one row per respondent. The **ProductTV** sheet does this for the 29 kept respondents (listed in column A, rows 5 to 33); each of columns D and E is a `SUMIFS` that picks the one matching Product_TV row from the Data sheet:

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Group (order) | `ProductTV!B5` | `=INDEX(Data!$E$2:$E$241,MATCH($A5,Data!$B$2:$B$241,0))` |
| Valid data | `ProductTV!C5` | `=INDEX(Data!$L$2:$L$241,MATCH($A5,Data!$B$2:$B$241,0))` |
| Balls Dwell time (%) | `ProductTV!D5` | `=SUMIFS(Data!$AB$2:$AB$241,Data!$B$2:$B$241,$A5,Data!$H$2:$H$241,"Product_TV",Data!$G$2:$G$241,"Commercial_1_Balls")` |
| Paint Dwell time (%) | `ProductTV!E5` | `=SUMIFS(Data!$AB$2:$AB$241,Data!$B$2:$B$241,$A5,Data!$H$2:$H$241,"Product_TV",Data!$G$2:$G$241,"Commercial_2_Paint")` |
| Difference (Paint minus Balls) | `ProductTV!F5` | `=E5-D5` |
| Paint seen | `ProductTV!G5` | `=IF(B5="A","second","first")` |

(By hand you would get the same table from the PivotTable of section 3a: put **Respondent Name** in Rows, **Stimulus** in Columns, **Dwell_pct_clean** in Values as Sum, and filter **AOI Name** to `Product_TV` and **Keep** to `Yes`.)

Now the test. The statistics sit in column J of the ProductTV sheet with the formula text shown beside them in column K:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Balls mean (%) | `ProductTV!J6` | `=AVERAGE(D5:D33)` | 21.24 |
| Balls SD (%) | `ProductTV!J7` | `=STDEV.S(D5:D33)` | 6.62 |
| Paint mean (%) | `ProductTV!J8` | `=AVERAGE(E5:E33)` | 28.49 |
| Paint SD (%) | `ProductTV!J9` | `=STDEV.S(E5:E33)` | 7.75 |
| n pairs | `ProductTV!J10` | `=COUNT(F5:F33)` | 29 |
| Mean difference, Paint minus Balls (%) | `ProductTV!J11` | `=AVERAGE(F5:F33)` | 7.25 |
| SD of the differences (%) | `ProductTV!J12` | `=STDEV.S(F5:F33)` | 7.40 |
| **p, two tailed (T.TEST type 1)** | `ProductTV!J16` | `=T.TEST(D5:D33,E5:E33,2,1)` | **0.000013** |

`T.TEST(array1, array2, tails, type)`: `tails = 2` for a two tailed test; `type = 1` for **paired** data, which is the type for this within subjects design. `type = 2` (equal variances) and `type = 3` (unequal variances, the Welch test of Workshop 1) treat the two columns as different people and give the wrong answer for paired data. `T.TEST` returns only the p value, so compute t and df by hand with the paired formulas (a paired t test is a one sample t test on the differences):

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| SE of the mean difference | `ProductTV!J13` | `=J12/SQRT(J10)` | 1.374 |
| Paired t | `ProductTV!J14` | `=J11/J13` | 5.27 |
| df | `ProductTV!J15` | `=J10-1` | 28 |
| p from the hand computed t | `ProductTV!J17` | `=T.DIST.2T(ABS(J14),J15)` | 0.000013, matches T.TEST |
| Critical t (95%) | `ProductTV!J18` | `=T.INV.2T(0.05,J15)` | 2.048 |
| Mean difference 95% CI lower | `ProductTV!J19` | `=J11-J18*J13` | 4.43 |
| Mean difference 95% CI upper | `ProductTV!J20` | `=J11+J18*J13` | 10.06 |

**You should see** the hand computed p equal the `T.TEST` p. If they differ, the `type` argument or a column range is wrong.

### Cohen's d_z by hand, with a 95 percent confidence interval

For paired data the effect size is **d_z**: the mean difference in units of the SD of the *differences*. It is linked to t by t = d_z x sqrt(n), so t, n, and d_z always agree.

| Step | Cell | Formula | Result |
|---|---|---|---|
| Cohen's d_z | `ProductTV!J21` | `=J11/J12` | 0.98 |
| SE of d_z (approximation) | `ProductTV!J22` | `=SQRT(1/J10+J21^2/(2*J10))` | 0.226 |
| d_z 95% CI lower | `ProductTV!J23` | `=J21-1.96*J22` | 0.54 |
| d_z 95% CI upper | `ProductTV!J24` | `=J21+1.96*J22` | 1.42 |
| Correlation of the two columns | `ProductTV!J25` | `=CORREL(D5:D33,E5:E33)` | 0.48 |

In words: d_z = mean difference / SD of differences; SE(d_z) = sqrt(1/n + d_z^2 / (2 n)). The interval is the usual large sample approximation. `analysis.R` reports the interval from the noncentral t distribution (effectsize package), which is more exact at this sample size; d_z itself matches, and the limits differ slightly ([0.53, 1.42] in R). Say which method you used. Rule of thumb for reading d_z: 0.2 small, 0.5 medium, 0.8 large. The correlation in the last row is the within respondent correlation: respondents who looked at the television a lot in one commercial also did in the other, which is why a paired test is more powerful than a test of two independent groups.

---

## 5. The order effect

Because every respondent saw both commercials, the second commercial may have benefited from the first (carryover). The simulated viewers were built with about +3 percentage points of Product_TV dwell share for whichever commercial came second. The paired differences carry this: for Group A (Paint seen second) the difference is commercial effect + order effect, and for Group B (Paint seen first) it is commercial effect - order effect. So half the gap between the two group means is the order effect:

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| Mean difference, Group A (Paint seen second) | `ProductTV!J28` | `=AVERAGEIFS(F5:F33,B5:B33,"A")` | 10.98 |
| Mean difference, Group B (Paint seen first) | `ProductTV!J29` | `=AVERAGEIFS(F5:F33,B5:B33,"B")` | 3.76 |
| n, Group A | `ProductTV!J30` | `=COUNTIF(B5:B33,"A")` | 14 |
| n, Group B | `ProductTV!J31` | `=COUNTIF(B5:B33,"B")` | 15 |
| Order effect (second minus first) | `ProductTV!J32` | `=(J28-J29)/2` | 3.61 |
| Commercial effect (Paint minus Balls), order adjusted | `ProductTV!J33` | `=(J28+J29)/2` | 7.37 |

`analysis.R` fits `dwell_pct ~ commercial * order_position + (1 | respondent)` with lme4 and lmerTest and gets the order effect b = 3.61, SE = 1.21, p = 0.006, and the same commercial effect 7.37; with 14 and 15 respondents in the two order groups the simple Excel estimate coincides with the model estimate. Excel cannot supply the standard error with the right degrees of freedom, the interaction test, or the **intraclass correlation** (ICC = 0.57 after the fixed effects): that is the share of the remaining variance that lies between respondents, the reason the two rows of a respondent are not independent. Use R for those.

The two check rows at the bottom of the block:

| Statistic | Cell | Formula |
|---|---|---|
| Kept respondents with a Product_TV row for Balls in Data | `ProductTV!J36` | `=COUNTIFS(Data!$G$2:$G$241,"Commercial_1_Balls",Data!$H$2:$H$241,"Product_TV",Data!$Z$2:$Z$241,"Yes")` |
| Smallest Valid data in column C | `ProductTV!J37` | `=MIN(C5:C33)` |

**You should see** 29 in the first (it must equal n pairs) and a number of at least 70 in the second.

---

## 6. H2: Brand_Logo time to first fixation

TTFF has a complication that dwell share does not: a respondent who never fixated the logo has **NA**, and a pair with an NA cannot enter a paired test. The **BrandLogo** sheet therefore lists only the 25 kept respondents who have a TTFF in **both** commercials (rows 5 to 29); the other 4 kept respondents have an NA in at least one commercial and leave the paired test. Columns C and D use the same `SUMIFS` idea on `TTFF AOI (ms)` (SUMIFS ignores the text NA, and a roster respondent has a number in both rows). `TTFF AOI (ms)` counts from the start of the end card, which is what the question is about; `TTFF max. (ms)` and `TTFF parent (ms)` count from the start of the commercial, so they are not used here.

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Balls TTFF AOI (ms) | `BrandLogo!C5` | `=SUMIFS(Data!$S$2:$S$241,Data!$B$2:$B$241,$A5,Data!$H$2:$H$241,"Brand_Logo",Data!$G$2:$G$241,"Commercial_1_Balls")` |
| Paint TTFF AOI (ms) | `BrandLogo!D5` | `=SUMIFS(Data!$S$2:$S$241,Data!$B$2:$B$241,$A5,Data!$H$2:$H$241,"Brand_Logo",Data!$G$2:$G$241,"Commercial_2_Paint")` |
| Difference (Paint minus Balls) | `BrandLogo!E5` | `=D5-C5` |

The block in column H repeats the paired recipe of section 4 on this table:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Balls mean (ms) | `BrandLogo!H6` | `=AVERAGE(C5:C29)` | 906.76 |
| Paint mean (ms) | `BrandLogo!H8` | `=AVERAGE(D5:D29)` | 786.96 |
| n pairs | `BrandLogo!H10` | `=COUNT(E5:E29)` | 25 |
| Mean difference, Paint minus Balls (ms) | `BrandLogo!H11` | `=AVERAGE(E5:E29)` | -119.80 |
| SD of the differences (ms) | `BrandLogo!H12` | `=STDEV.S(E5:E29)` | 330.69 |
| Paired t | `BrandLogo!H14` | `=H11/H13` | -1.81 |
| df | `BrandLogo!H15` | `=H10-1` | 24 |
| p, two tailed (T.TEST type 1) | `BrandLogo!H16` | `=T.TEST(C5:C29,D5:D29,2,1)` | 0.082620 |
| Cohen's d_z | `BrandLogo!H21` | `=H11/H12` | -0.36 |
| d_z 95% CI lower | `BrandLogo!H23` | `=H21-1.96*H22` | -0.77 |
| d_z 95% CI upper | `BrandLogo!H24` | `=H21+1.96*H22` | 0.04 |

The respondent ratio (the fixation based **Respondent ratio (%)**) is below the block:

| Statistic | Cell | Formula | Result |
|---|---|---|---|
| Fixated the logo, Balls | `BrandLogo!H29` | `=COUNTIFS(Data!$H$2:$H$241,"Brand_Logo",Data!$Z$2:$Z$241,"Yes",Data!$G$2:$G$241,"Commercial_1_Balls",Data!$AC$2:$AC$241,1)` | 27 |
| Respondent ratio, Balls (%) | `BrandLogo!H30` | `=100*H29/H28` | 93.1 |
| Fixated the logo, Paint | `BrandLogo!H32` | `=COUNTIFS(Data!$H$2:$H$241,"Brand_Logo",Data!$Z$2:$Z$241,"Yes",Data!$G$2:$G$241,"Commercial_2_Paint",Data!$AC$2:$AC$241,1)` | 27 |
| Respondent ratio, Paint (%) | `BrandLogo!H33` | `=100*H32/H31` | 93.1 |

**Common mistake.** Treating the NA as 0 ms: that would make the respondent the fastest viewer in the sample. The other tempting shortcut, giving non fixators the longest possible value (5000 ms, the length of the end card), answers a different question; `analysis.R` runs it as a sensitivity analysis (H2b) and the result is in `analysis_output.txt`. Excel has no built in Wilcoxon signed rank test either; in this sample the signed rank test (p = 0.046) and the paired t test (p = 0.083) fall on opposite sides of .05, so run it in R before you draw a conclusion.

---

## 7. H3: Liking

The **Survey** sheet holds `SIMULATED_ad_test_survey.csv` (one row per respondent, rows 4 to 33), with the kept respondents first (rows 4 to 32) and the excluded respondent last (row 33), so the test ranges below leave it out. Column H flags Keep from the Data sheet, column I is the difference, column J sums the Brand_Logo dwell time over both commercials, and column K codes recall as 1 or 0:

| Column | Cell (first row) | Formula |
|---|---|---|
| Keep | `Survey!H4` | `=IF(INDEX(Data!$L$2:$L$241,MATCH($B4,Data!$B$2:$B$241,0))>=70,"Yes","No")` |
| Difference (Paint minus Balls) | `Survey!I4` | `=E4-D4` |
| Brand_Logo Dwell time (ms), both commercials | `Survey!J4` | `=SUMIFS(Data!$V$2:$V$241,Data!$B$2:$B$241,$B4,Data!$H$2:$H$241,"Brand_Logo")` |
| Recall (1 = Yes) | `Survey!K4` | `=IF(F4="Yes",1,0)` |

The paired recipe sits in column N (formula text in column O):

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Balls mean (points) | `Survey!N5` | `=AVERAGE(D4:D32)` | 5.52 |
| Paint mean (points) | `Survey!N7` | `=AVERAGE(E4:E32)` | 5.00 |
| n pairs | `Survey!N9` | `=COUNT(I4:I32)` | 29 |
| Mean difference, Paint minus Balls (points) | `Survey!N10` | `=AVERAGE(I4:I32)` | -0.52 |
| SD of the differences (points) | `Survey!N11` | `=STDEV.S(I4:I32)` | 1.09 |
| Paired t | `Survey!N13` | `=N10/N12` | -2.56 |
| df | `Survey!N14` | `=N9-1` | 28 |
| p, two tailed (T.TEST type 1) | `Survey!N15` | `=T.TEST(D4:D32,E4:E32,2,1)` | 0.016290 |
| Cohen's d_z | `Survey!N20` | `=N10/N11` | -0.47 |
| d_z 95% CI lower | `Survey!N22` | `=N20-1.96*N21` | -0.86 |
| d_z 95% CI upper | `Survey!N23` | `=N20+1.96*N21` | -0.09 |

The negative sign means Paint was liked *less*. Balls is liked more even though Paint draws more attention to the television set: attention (what the eyes did) and attitude (what the respondent says) are different measures, which is the main teaching point of this study. The counts below the test block (`COUNTIF` of Ad_preferred, `SUMPRODUCT` of the ratings) show that 17 of the 29 respondents gave the two commercials the same rating (a tied pair carries no information for a signed rank test, which drops it); the Wilcoxon test is in R.

### Brand recall and logo attention (point biserial correlation)

`Brand_recall` is one answer per respondent, so the logo attention measure is one number per respondent too: the total fixation based dwell time on Brand_Logo over both commercials (column J; the text NA is ignored by `SUMIFS`, so a never fixated logo counts as 0 ms). A Pearson correlation between a continuous variable and a 0/1 variable is the **point biserial** correlation:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Point biserial r (CORREL) | `Survey!N39` | `=CORREL(J4:J32,K4:K32)` | 0.42 |
| t for r | `Survey!N40` | `=N39*SQRT((N35-2)/(1-N39^2))` | 2.38 |
| p, two tailed | `Survey!N41` | `=T.DIST.2T(ABS(N40),N35-2)` | 0.0245 |
| r 95% CI lower (Fisher z) | `Survey!N42` | `=TANH(ATANH(N39)-NORM.S.INV(0.975)/SQRT(N35-3))` | 0.06 |
| r 95% CI upper (Fisher z) | `Survey!N43` | `=TANH(ATANH(N39)+NORM.S.INV(0.975)/SQRT(N35-3))` | 0.68 |

---

## 8. Chart with 95 percent confidence intervals

1. Make a small chart table: two rows (Balls, Paint), columns **Mean dwell (%)** (point to the mean cells) and **CI half width** = `CONFIDENCE.T(0.05, SD cell, n cell)`. `CONFIDENCE.T` returns half the width of a t based 95 percent interval around a mean. The workbook has this table on the ProductTV sheet at `ProductTV!I42:K43`; the half widths are in `ProductTV!K42` and `ProductTV!K43`.
2. Select the commercial names and the means, **Insert > Column chart > Clustered Column**.
3. Click the bars, **Chart Design > Add Chart Element > Error Bars > More Error Bars Options**, choose **Custom > Specify Value**, and select the two CI half width cells for both **Positive** and **Negative**.
4. Set the bar fill to navy (`#1F3864`) and lighten the gridlines.

**Common mistake.** Using the default error bars (**Standard Error** or a fixed 5 percent). Those are not 95 percent confidence intervals. And remember that error bars around the two means are *not* the interval that goes with the paired test: the paired test uses the interval of the *differences* (`ProductTV!J19` and `ProductTV!J20`), which is much narrower because it removes the between respondent variation.

---

## 9. Check your numbers and know what Excel cannot do

The **Check** sheet puts each live Excel result next to the value `analysis.R` produced and prints OK or CHECK (status in column F, the count of CHECK rows in `Check!B47`). Your Excel results should match to rounding; the d_z limits differ by method and have a looser tolerance. Several parts of the R analysis have no Excel equivalent:

- **Wilcoxon signed rank tests** (a check that the conclusion does not depend on normality). In R: H1 V = 402, H2 V = 88, H3 V = 9.
- **Mixed effects model** with a random intercept per respondent and the **ICC**: `lme4::lmer(dwell_pct ~ commercial * order_position + (1 | respondent))`.
- **Sensitivity analysis** for never fixators (H2b), **exact McNemar** test of the two respondent ratios, and **exact binomial** test of the preference split.

If you have the **Analysis ToolPak** add in (**File > Options > Add-ins > Manage Excel Add-ins > Go > Analysis ToolPak**), **Data > Data Analysis > t-Test: Paired Two Sample for Means** prints t, df, and p in one table and is a good cross check of sections 4, 6, and 7.

---

## 10. Common mistakes, collected

1. Opening the .csv by double clicking and getting one column: use **Data > From Text/CSV** or the Text Import Wizard.
2. Replacing NA with 0 in a TTFF column. NA means not observed.
3. Forgetting the exclusion filter (Keep = Yes) in the PivotTable or in COUNTIFS, or leaving R23 in a roster.
4. `T.TEST` with `type = 2` or `type = 3` on a within subjects design. Use `type = 1` and make sure the two columns are in the same respondent order.
5. Reporting `T.TEST` alone. Always give M, SD, n, t, df, p, and d_z with its interval.
6. Pairing the wrong rows: the two ranges must list the respondents in the same order, and a respondent with an NA has to leave the roster entirely.
7. Treating 240 rows as 240 viewers. Each respondent contributes eight rows; per AOI tests use one value per respondent and commercial, and the all rows analysis needs the mixed model.
8. Ignoring the order of viewing. The commercial seen second gains extra dwell share; counterbalancing makes the commercial comparison fair, but the order effect still adds noise and should be checked.

## Sources

- iMotions help center, AOI Metrics: https://help.imotions.com/lab/analysis/eye-tracking/areas-of-interest/aoi-metrics (local copy: `C:\JaxState_Behavior Lab\Eye Tracking Glasses (ETG)\AOI Metrics.pdf`)
- iMotions help center, Areas of Interest (AOIs), including dynamic AOIs and export options: https://help.imotions.com/lab/analysis/eye-tracking/areas-of-interest/areas-of-interest-aois (local copy: `C:\JaxState_Behavior Lab\Eye Tracking Glasses (ETG)\Areas of Interest AOIs.pdf`)
- iMotions help center, Opening Data Export Files in Excel: https://help.imotions.com/lab/exports/additional-export-information/opening-data-export-files-in-excel (local copy: `iMotion_Manual\Opening Data Export Files in Excel.pdf`)
- iMotions help center, Creating Excel Pivot Tables: https://help.imotions.com/lab/exports/additional-export-information/creating-excel-pivot-tables (local copy: `iMotion_Manual\Making Excel Pivot Tables.pdf`)

## Items to verify on the lab laptop

- The exact column header layout of the Individual AOI metrics export in iMotions Lab 11.1 (the manuals name the metrics but do not print the header row). Adjust the column letters in sections 1 and 2 once a real export is open.
- Whether the real export writes `NA` or leaves cells blank for never fixated AOIs; `ISNUMBER` handles both.
- How the export reports `AOI duration (ms)`, `TTFF AOI (ms)`, and `Dwell time (%)` for a dynamic AOI that is active in several separate intervals (Product_TV and Hero_Element here). The AOI Metrics article defines the start time and the active time but does not show a multi interval example.
- How `TTFF max. (ms)` is written for a dynamic AOI. The AOI Metrics article says it counts from the start of the parent and uses the parent's presentation duration when there was no fixation; the simulated file follows that wording (60000 ms).
- Whether the dwell classification threshold (default 100 ms) and the first fixation exclusion were left at their defaults when the export was made; both change the numbers (Areas of Interest article, Export options).
