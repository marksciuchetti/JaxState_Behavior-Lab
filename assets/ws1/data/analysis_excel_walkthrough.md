# Shelf study analysis in Excel: step by step walkthrough

**SIMULATED DATA.** Every number in this walkthrough comes from `SIMULATED_shelf_study_AOI_metrics.csv`, which `make_simulated_data.py` generated (seed 20261037). No real shopper was recorded. The design mirrors the iMotions In-Store Shopper Report (Sensodyne point of sale material, POSM, versus no POSM on an oral care shelf, eye tracking glasses) in direction and rough size only.

This is the same analysis as `analysis.R`, done in Excel for students who do not use R. The finished workbook is `SIMULATED_shelf_study_analysis.xlsx`; the sheet names below refer to it. You can follow along in that workbook or rebuild it from the CSV files.

Companion files: `analysis.R` (R version), `analysis_output.txt` (its console output), `results_text.md` (model paragraphs), `figures/` (PNG figures).

---

## 0. What you are analyzing

**Design.** Between subjects. Each shopper wore eye tracking glasses and shopped one shelf: the POSM shelf (Sensodyne wing display present) or the NoPOSM shelf. 30 respondents, 15 per shelf, IDs P01 to P30.

**Export.** One row per respondent by AOI, the layout of the iMotions **Individual AOI metrics** export, which iMotions recommends for statistics because it keeps each respondent's values instead of averaging across the segment (Areas of Interest article, Export options). The columns after the identifiers are iMotions AOI metrics with the names from the **AOI Metrics** help article. The ones used here:

| Column | iMotions definition (AOI Metrics article) | Used for |
|---|---|---|
| `Valid data` | Percentage of collected and interpolated samples; below 100 means missing data | Exclusion rule |
| `TTFF parent (ms)` | Time from the start of the AOI's parent (here the gaze map of the shelf) until the first fixation on the AOI | Primary hypothesis |
| `TTFF max. (ms)` | As TTFF parent, but a respondent who never fixated the AOI gets the parent's full duration | Check that includes non lookers |
| `Dwell time (ms)` | Total time fixating on the AOI | Attention |
| `Dwell time (%)` | Dwell time relative to the time the AOI was active | Secondary hypothesis (dwell share) |
| `Fixation count` | Number of fixations inside the AOI | Descriptives |
| `Revisit count` | Look backs with a fixation after the first dwell | Descriptives |
| `Gaze Hit time parent (ms)` | Gaze based: time until gaze first entered the AOI (no fixation needed) | Comparison with TTFF |

Where the same metric name exists as a gaze based and a fixation based metric (Dwell time, Revisit count), the gaze based column carries the prefix `Gaze`. iMotions writes **NA** when a metric cannot be calculated, for example when the respondent never fixated the AOI. The simulated file does the same.

**Hypotheses.**
- H1 (primary): POSM shortens the time to first fixation on `Sensodyne_Products`.
- H2 (secondary): POSM raises the dwell time share on `Sensodyne_Products`.
- Descriptive: POSM lowers time on `Competitor_A` plus `Competitor_B` (the 2020 report's headline, almost 25 percent less).

**Decision rules fixed before looking at the data.** Exclude a respondent whose `Valid data` is below 70 percent. For an AOI the respondent never fixated, TTFF is *not observed* (leave it out of TTFF statistics) and dwell time is *0* (keep it in dwell statistics). Two tailed tests, alpha = .05.

---

## 1. Import the export into Excel

iMotions exports are comma separated with a dot decimal. Double clicking the file uses your regional defaults, which can put everything in one column (Opening Data Export Files in Excel article). Use the import route instead.

**Route A, Get Data (Excel 2016 and later):**
1. **Data > From Text/CSV**, select `SIMULATED_shelf_study_AOI_metrics.csv`, click **Import**.
2. In the preview set **Delimiter: Comma** and **File Origin: 65001 Unicode (UTF-8)**. **You should see** 25 columns with `Study Name` in the first column.
3. Click **Transform Data** if the first row shows the `# SIMULATED ...` comment line: in Power Query choose **Home > Remove Rows > Remove Top Rows > 1**, then **Home > Use First Row as Headers**, then **Close & Load**. (A real iMotions export has no comment line; the simulated file carries one so nobody mistakes it for real data.)
4. **You should see** a table with 165 data rows (30 respondents, six AOIs on the POSM shelf and five on the NoPOSM shelf).

**Route B, Text Import Wizard (any version):** **File > Open > Browse**, set the file type to **All Files**, open the .csv, choose **Delimited**, start at row 2 (to skip the comment line), tick **Comma**, click **Advanced** and set the decimal symbol to a dot, then **Finish**.

**Common mistake.** The NA cells must stay as the text `NA` (or be blank). Do not replace NA with 0 in the TTFF columns: a zero TTFF would mean the shopper fixated the product at the very first instant, which is false. Step 2 handles NA correctly.

Save as **.xlsx** right away; a .csv does not keep formulas or extra sheets. Keep the original .csv read only.

In the workbook this is the **Data** sheet. Column letters there: `E` = Group (condition), `H` = AOI Name, `L` = Valid data, `R` = Fixation count, `T` = TTFF parent (ms), `W` = Dwell time (%). Data rows run from 2 to 166.

---

## 2. Clean with helper columns

Add four columns to the right of the data (Z to AC in the workbook) and fill them down to the last row.

| Column | Header | Formula in row 2 | Meaning |
|---|---|---|---|
| Z | Keep | `=IF(L2>=70,"Yes","No")` | Exclusion rule on Valid data |
| AA | TTFF_s | `=IF(ISNUMBER(T2),T2/1000,"")` | TTFF in seconds; blank when not observed |
| AB | Dwell_pct_clean | `=IF(ISNUMBER(W2),W2,0)` | Never fixated AOI counts as 0 percent |
| AC | Fixated | `=IF(ISNUMBER(T2),1,0)` | 1 when at least one fixation hit the AOI |

**You should see** `Keep = No` on all rows of respondent P17 (Valid data 46 percent; the session log says the glasses slipped at about 40 seconds and the gaze offset is visible in the replay). Every other respondent is `Yes`. Respondent P09 needed a second gaze check but passed it and stays in; a repeated check is not by itself a reason to exclude.

`ISNUMBER` is the key: it is TRUE for a number and FALSE for the text NA or a blank cell, so one formula handles both ways an export can mark a missing metric.

---

## 3. Descriptive table by condition and AOI

### 3a. With a PivotTable (the iMotions recommended route)

The Creating Excel Pivot Tables article gives the generic steps: click the first cell of the data, **Insert > PivotTable > OK**, drag a category to **Rows**, drag a metric to **Values**, right click the value, **Value Field Settings**, change **Sum** to **Average**, **OK**, then **PivotChart** if you want a quick picture. For this study:

1. Click cell A1 of the Data sheet, **Insert > PivotTable**, **New Worksheet**, **OK**.
2. Drag **Keep** to **Filters** and set it to **Yes**.
3. Drag **AOI Name** to **Rows** and **Group** to **Columns**.
4. Drag **TTFF_s** to **Values** three times. Set the first to **Average**, the second to **StdDev** (this is the sample standard deviation, the same as STDEV.S), the third to **Count** (Count ignores the blank cells, so it gives the number of respondents who fixated the AOI).
5. Drag **Dwell_pct_clean** to **Values** twice: **Average** and **StdDev**.
6. Drag **Fixated** to **Values** and set it to **Average**; multiply by 100 in your head or in a cell beside the table. This is the fixation based **Respondent ratio (%)** from the iMotions AOI table.

**You should see** `Sensodyne_Products`: NoPOSM average TTFF_s 3.59 (count 15), POSM 2.25 (count 14); average Dwell_pct_clean 10.76 versus 18.43.

**Common mistake.** Leaving **Keep** out of the filter. P17 would then add six rows of thin data to the POSM column.

### 3b. With formulas (the Summary sheet)

A formula table updates when the data change and shows exactly what is being averaged. On a sheet with the AOI name in column A and the condition in column B of each row (row 5 is the first), the Summary sheet uses:

| Statistic | Formula in row 5 |
|---|---|
| n (kept rows for this AOI) | `=COUNTIFS(Data!$E$2:$E$166,$B5,Data!$H$2:$H$166,$A5,Data!$Z$2:$Z$166,"Yes")` |
| Respondent ratio (%) | `=100*AVERAGEIFS(Data!$AC$2:$AC$166,Data!$E$2:$E$166,$B5,Data!$H$2:$H$166,$A5,Data!$Z$2:$Z$166,"Yes")` |
| n with TTFF | `=COUNTIFS(Data!$E$2:$E$166,$B5,Data!$H$2:$H$166,$A5,Data!$Z$2:$Z$166,"Yes",Data!$AA$2:$AA$166,">0")` |
| TTFF mean (s) | `=AVERAGEIFS(Data!$AA$2:$AA$166,Data!$E$2:$E$166,$B5,Data!$H$2:$H$166,$A5,Data!$Z$2:$Z$166,"Yes")` |
| TTFF SD (s) | `=STDEV.S(IF((Data!$E$2:$E$166=$B5)*(Data!$H$2:$H$166=$A5)*(Data!$Z$2:$Z$166="Yes")*ISNUMBER(Data!$AA$2:$AA$166),Data!$AA$2:$AA$166))` |
| Dwell time mean (%) | `=AVERAGEIFS(Data!$AB$2:$AB$166,Data!$E$2:$E$166,$B5,Data!$H$2:$H$166,$A5,Data!$Z$2:$Z$166,"Yes")` |
| Dwell time SD (%) | `=STDEV.S(IF((Data!$E$2:$E$166=$B5)*(Data!$H$2:$H$166=$A5)*(Data!$Z$2:$Z$166="Yes"),Data!$AB$2:$AB$166))` |
| Fixation count mean | `=AVERAGEIFS(Data!$R$2:$R$166,Data!$E$2:$E$166,$B5,Data!$H$2:$H$166,$A5,Data!$Z$2:$Z$166,"Yes")` |

The two `STDEV.S(IF(...))` formulas are **array formulas** because Excel has no STDEVIFS. In Excel 365 type them and press **Enter**. In Excel 2019 or older press **Ctrl+Shift+Enter**; Excel then shows the formula in braces. `AVERAGEIFS` and `COUNTIFS` ignore the blank TTFF_s cells, which is exactly the "TTFF not observed" rule.

**You should see** the same numbers as the PivotTable, and the Summary sheet includes a clustered bar chart of mean dwell share by AOI and condition.

---

## 4. The primary test: TTFF to Sensodyne_Products, Welch's t test

Excel's `T.TEST` needs two blocks of numbers, one per condition. Build them:

1. On the Data sheet, **Data > Filter**. Filter **AOI Name** to `Sensodyne_Products` and **Keep** to `Yes`. **You should see** 29 rows.
2. Select the visible rows, copy, and paste into a new sheet (**Sensodyne** in the workbook). Sort by **Group** so the POSM rows are together (rows 5 to 18) and the NoPOSM rows are together (rows 19 to 33). Column `E` holds TTFF_s and column `G` holds Dwell_pct_clean.
3. Now the test. In the workbook the formulas sit in column K with the formula text shown beside them in column L:

| Statistic | Formula | Result (simulated) |
|---|---|---|
| POSM mean | `=AVERAGE(E5:E18)` | 2.25 s |
| POSM SD | `=STDEV.S(E5:E18)` | 0.86 |
| POSM n | `=COUNT(E5:E18)` | 14 |
| NoPOSM mean | `=AVERAGE(E19:E33)` | 3.59 s |
| NoPOSM SD | `=STDEV.S(E19:E33)` | 0.98 |
| NoPOSM n | `=COUNT(E19:E33)` | 15 |
| **p, two tailed Welch** | `=T.TEST(E5:E18,E19:E33,2,3)` | **0.0006** |

`T.TEST(array1, array2, tails, type)`: `tails = 2` for a two tailed test; `type = 3` for two samples with unequal variances, which is Welch's test, the version `analysis.R` uses. `type = 2` would assume equal variances and `type = 1` is for paired data (never for a between subjects design). `COUNT`, `AVERAGE`, `STDEV.S`, and `T.TEST` all ignore blank cells, so a never fixated respondent drops out of the TTFF test automatically.

`T.TEST` gives only the p value. To report t and df, compute them by hand with the Welch formulas (M = mean, s = SD, n = count; subscript 1 = POSM, 2 = NoPOSM):

| Statistic | Formula (K6 = M1, K7 = s1, K8 = n1, K9 = M2, K10 = s2, K11 = n2) | Result |
|---|---|---|
| t | `=(K6-K9)/SQRT(K7^2/K8+K10^2/K11)` | -3.91 |
| df (Welch Satterthwaite) | `=(K7^2/K8+K10^2/K11)^2/((K7^2/K8)^2/(K8-1)+(K10^2/K11)^2/(K11-1))` | 26.89 |
| p from t and df | `=T.DIST.2T(ABS(K13),K14)` | 0.0006, matches T.TEST |

**You should see** the hand computed p equal the `T.TEST` p. If they differ, the `type` argument or a block range is wrong.

---

## 5. Cohen's d by hand, with a 95 percent confidence interval

Cohen's d expresses the difference in pooled standard deviation units, which a reader can compare across studies.

| Step | Formula | Result |
|---|---|---|
| Pooled SD | `=SQRT(((K8-1)*K7^2+(K11-1)*K10^2)/(K8+K11-2))` | 0.92 |
| d | `=(K6-K9)/K17` | -1.45 |
| SE of d (approximation) | `=SQRT((K8+K11)/(K8*K11)+K18^2/(2*(K8+K11)))` | 0.42 |
| 95% CI lower | `=K18-1.96*K19` | -2.26 |
| 95% CI upper | `=K18+1.96*K19` | -0.63 |

In words: pooled SD = sqrt(((n1 - 1) s1^2 + (n2 - 1) s2^2) / (n1 + n2 - 2)); d = (M1 - M2) / pooled SD; SE(d) = sqrt((n1 + n2) / (n1 n2) + d^2 / (2 (n1 + n2))). The negative sign means POSM is *lower* (faster). The interval formula is the usual large sample approximation. `analysis.R` reports the interval from the noncentral t distribution (effectsize package), which is more exact at this sample size; d itself matches, the limits differ slightly ([-2.26, -0.61] in R). Say which method you used.

Rule of thumb for reading d: 0.2 small, 0.5 medium, 0.8 large. A d near 1.4 is large, and the simulated data were built that way so that the effect is easy to see with 15 per group. Real shelf effects are often smaller.

---

## 6. The secondary test: dwell time share

Repeat Sections 4 and 5 on column `G` (Dwell_pct_clean) of the Sensodyne sheet. Here every kept respondent has a value (0 for never fixated), so n = 14 and 15 again.

| Statistic | Formula | Result |
|---|---|---|
| POSM mean, SD | `=AVERAGE(G5:G18)`, `=STDEV.S(G5:G18)` | 18.43 percent, 5.98 |
| NoPOSM mean, SD | `=AVERAGE(G19:G33)`, `=STDEV.S(G19:G33)` | 10.76 percent, 4.42 |
| p, Welch | `=T.TEST(G5:G18,G19:G33,2,3)` | 0.0007 |
| t, df | same hand formulas | 3.91, 23.87 |
| d, 95% CI | same hand formulas | 1.47 [0.65, 2.29] |

---

## 7. Chart with 95 percent confidence intervals

1. Make a small chart table: two rows (NoPOSM, POSM), columns **Mean TTFF (s)** (point to the mean cells) and **CI half width** = `=CONFIDENCE.T(0.05, SD cell, n cell)`. `CONFIDENCE.T` returns half the width of a t based 95 percent interval around a mean. **You should see** about 0.54 for NoPOSM and 0.49 for POSM.
2. Select the condition names and the means, **Insert > Column chart > Clustered Column**.
3. Click the bars, **Chart Design > Add Chart Element > Error Bars > More Error Bars Options**, choose **Custom > Specify Value**, and select the two CI half width cells for both **Positive** and **Negative**.
4. Set the bar fill to navy (`#1F3864`) and remove gridlines or lighten them. Title: "Mean TTFF to Sensodyne_Products by condition (SIMULATED)".

The workbook's Sensodyne sheet has this chart ready made from the chart table at J42:L45.

**Common mistake.** Using the default error bars (**Standard Error** or a fixed 5 percent). Those are not 95 percent confidence intervals.

---

## 8. Survey

The **Survey** sheet joins `SIMULATED_shelf_study_survey.csv` to each respondent's Sensodyne dwell share with `INDEX/MATCH` against the Sensodyne sheet, and marks Keep from the same lookup. Purchase intent (1 to 7) uses the same Welch recipe: `=T.TEST(POSM block, NoPOSM block, 2, 3)` gives p = 0.116 and d = 0.60; the difference (4.86 versus 4.13) is in the expected direction but the test does not reach alpha with 14 and 15 respondents. Brand recall and product chosen are counts: `=COUNTIFS(Group range,"POSM",Recall range,"Yes",Keep range,"Yes")` and so on; R runs Fisher's exact test on these tables (p = 0.109 and 0.120). Excel has no Fisher test built in; report the counts and percentages, or use R. `=CORREL(dwell range, purchase intent range)` gives r = 0.49 between Sensodyne dwell share and purchase intent.

---

## 9. Check your numbers and know what Excel cannot do

The **Check** sheet holds the values `analysis.R` produced. Your Excel results should match to rounding. Two parts of the R analysis have no Excel equivalent:

- **Wilcoxon rank sum test** (a check that the conclusion does not depend on normality). Excel has no built in version. With 15 per group and a right skewed metric such as TTFF, it is worth running in R; in this sample it agrees with the t tests (p = 0.0006 and p = 0.0007).
- **Mixed effects model** of dwell share across AOIs with a random intercept per respondent. This is how to analyze all six AOIs at once without treating a shopper's six rows as six independent people. Excel cannot fit it; `analysis.R` does with `lme4::lmer(dwell_pct ~ aoi * condition + (1 | respondent))`. In the simulated data the AOI by condition interaction is clear, F(4, 108) = 8.14, p < .001, and about 31 percent of the dwell variance sits between respondents.

If you have the **Analysis ToolPak** add in (**File > Options > Add-ins > Manage Excel Add-ins > Go > Analysis ToolPak**), **Data > Data Analysis > t-Test: Two-Sample Assuming Unequal Variances** prints t, df, and p in one table and is a good cross check of Sections 4 and 6.

---

## 10. Common mistakes, collected

1. Opening the .csv by double clicking and getting one column: use **Data > From Text/CSV** or the Text Import Wizard.
2. Replacing NA with 0 in a TTFF column. NA means not observed.
3. Forgetting the exclusion filter (Keep = Yes) in the PivotTable or in COUNTIFS.
4. `T.TEST` with `type = 1` (paired) or `type = 2` (equal variances) when the design is between subjects with unequal spread.
5. Reporting `T.TEST` alone. Always give M, SD, n, t, df, p, and d with its interval.
6. Non contiguous blocks in `T.TEST`. Sort by Group first.
7. Treating 165 rows as 165 shoppers. Each shopper contributes five or six rows; per AOI tests use one row per shopper, and the all AOI analysis needs the mixed model.

## Sources

- iMotions help center, AOI Metrics: https://help.imotions.com/lab/analysis/eye-tracking/areas-of-interest/aoi-metrics (local copy: `C:\JaxState_Behavior Lab\Eye Tracking Glasses (ETG)\AOI Metrics.pdf`)
- iMotions help center, Areas of Interest (AOIs): https://help.imotions.com/lab/analysis/eye-tracking/areas-of-interest/areas-of-interest-aois (local copy: `C:\JaxState_Behavior Lab\Eye Tracking Glasses (ETG)\Areas of Interest AOIs.pdf`)
- iMotions help center, Opening Data Export Files in Excel: https://help.imotions.com/lab/exports/additional-export-information/opening-data-export-files-in-excel
- iMotions help center, Creating Excel Pivot Tables: https://help.imotions.com/lab/exports/additional-export-information/creating-excel-pivot-tables
- iMotions Enablement Services, In-Store Shopper Report, March 2020 (local copy: `C:\JaxState_Behavior Lab\iMotion_Manual\1Printed\In-Store Shopper Sample Report 2020.pdf`)

## Items to verify on the lab laptop

- The exact column header layout of the Individual AOI metrics export in iMotions Lab 11.1 (the manuals name the metrics but do not print the header row). Adjust the column letters in Sections 2 and 3 once a real export is open.
- Whether the real export writes `NA` or leaves cells blank for never fixated AOIs; `ISNUMBER` handles both.
- Whether the dwell classification threshold (default 100 ms) and the first fixation exclusion were left at their defaults when the export was made; both change the numbers (Areas of Interest article, Export options).
