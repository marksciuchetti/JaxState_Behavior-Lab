# Excel walkthrough: descriptive statistics, paired tests, and the Holm correction

**SIMULATED DATA.** Every number in this walkthrough comes from the SIMULATED Multimodal Ad Study of Workshop 7 (`analysis_ready_long.csv`, seed 20271231). No real person was recorded, ALTAVO and BRENLOW are invented brands, and the numbers must never be quoted as findings.

This is the Excel route through three parts of `analysis_full.R`: the descriptive statistics by ad and AOI (STEP 3), the Q5 paired tests of the facial expressions (STEP 6), and the Holm correction. The finished workbook is `SIMULATED_analysis_report.xlsx`. **Every cell address below is written as `Sheet!Cell` and was generated from the workbook itself:** `make_excel_workbook.py` writes the workbook and this file from one list of cells, reads the workbook back with openpyxl to confirm that each address holds the formula printed here, and recalculates a copy to confirm the results. A copy was recalculated with LibreOffice 24.2.7.2 420(Build:2) (headless): all 256 Check rows are OK (`Check!B261` = 0), no cell shows an error value, the largest absolute difference from R is 4.7e-05, and the 254 results printed here agree with R at the printed decimals.

Companion files: `analysis_full.R` and `analysis_output.txt` (the R analysis and its full output), `SIMULATED_descriptives.csv` and `SIMULATED_all_tests.csv` (the R values the `Check` sheet uses), `results_text.md` (model paragraphs), `README.md`.

---

## 0. What you are working with

| Sheet | What it holds | Rows |
|---|---|---|
| `Data` | `analysis_ready_long.csv` (Workshop 7) as the Text Import leaves it; helper column `AD` (orange) | 72 |
| `AOI_long` | one row per participant, ad, and AOI; the dwell values stacked into one column | 144 |
| `Descriptives` | n, mean, SD, median, quartiles, IQR, percent at zero | 20 |
| `Pivot` | Average of dwell by AOI and ad (the PivotTable, as formulas) | 4 |
| `Q5` | Joy, Engagement, Brow Furrow: Ad_A versus Ad_B, paired | 22 participants |
| `Holm` | the Holm correction of the three Q5 p values | 3 tests |
| `Check` | every Excel result next to the R value | 256 |

Blue headers mark data as imported; orange headers mark columns you add; yellow cells are typed by you (the participant, ad, and AOI grid and the roster). **Analysis sample:** include = 1 (22 of 24 participants; P18 and P21 were excluded by rule R2 in Workshop 7).

---

## 1. Import the analysis ready table (Data sheet)

`analysis_ready_long.csv` has one row per participant and stimulus (Baseline, Ad_A, Ad_B). Line 1 is the SIMULATED banner; the column names are on line 2.

1. **Data > From Text/CSV** (Excel 2016 and later), select the file, **Transform Data**; in Power Query **Home > Remove Rows > Remove Top Rows** (1), then **Home > Use First Row as Headers**, set `participant_id`, `stimulus`, and `order` to **Text**, **Close & Load**. (Older Excel: **File > Open**, Text Import Wizard, **Delimited**, **Start import at row: 2**, **Comma**.)
2. **You should see** 29 columns (`A` = participant_id to `AC` = recall_correct) and 72 data rows. Cells that read `NA` are text: the value does not exist (for example liking on a Baseline row). Leave them; AVERAGEIFS, STDEV.S, and MEDIAN skip text.
3. Rename the sheet `Data`. Add the key in column `AD` and fill it down to row 73:

| Result | Cell | Formula | Value |
|---|---|---|---|
| Key of the first row | `Data!AD2` | `=A2&"\|"&B2` | `P01\|Baseline` |

**Common mistake:** importing with the banner as the header row. Every column is then named `# SIMULATED ...` or `Column1`, and the numbers come in as text.

---

## 2. Stack the three AOIs into one column (AOI_long sheet)

The table has three dwell columns (`Y` Logo, `Z` Product, `AA` Headline). AVERAGEIFS and a PivotTable need ONE value column and one column that says which AOI the value belongs to ("long" format). Type the grid in columns `A` to `C` (yellow: P01 Ad_A Logo, P01 Ad_A Product, P01 Ad_A Headline, P01 Ad_B Logo, ...; 144 rows), then fill these two formulas down to row 145:

| Column | Result | Cell | Formula |
|---|---|---|---|
| `D` | include, looked up from Data | `AOI_long!D2` | `=INDEX(Data!$G$2:$G$73,MATCH(A2&"\|"&B2,Data!$AD$2:$AD$73,0))` |
| `E` | gaze dwell of the AOI named in column C | `AOI_long!E2` | `=IF(C2="Logo",INDEX(Data!$Y$2:$Y$73,MATCH(A2&"\|"&B2,Data!$AD$2:$AD$73,0)),IF(C2="Product",INDEX(Data!$Z$2:$Z$73,MATCH(A2&"\|"&B2,Data!$AD$2:$AD$73,0)),INDEX(Data!$AA$2:$AA$73,MATCH(A2&"\|"&B2,Data!$AD$2:$AD$73,0))))` |

**Excel 365:** `=XLOOKUP(A2&"|"&B2,Data!$AD$2:$AD$73,Data!$G$2:$G$73)` does the same as the INDEX and MATCH pair in column `D`.

**Check the stacking** (cells on the right of `AOI_long`):

| Check | Cell | Formula | Value |
|---|---|---|---|
| Rows (24 x 2 x 3) | `AOI_long!H3` | `=COUNTA(A2:A145)` | 144 |
| Lookups that failed (#N/A) | `AOI_long!H4` | `=SUMPRODUCT(--ISNA(D2:E145))` | 0 |
| Rows with include = 1 (22 x 6) | `AOI_long!H5` | `=COUNTIF(D2:D145,1)` | 132 |

---

## 3. Descriptive statistics (Descriptives sheet)

Report n, a center, and a spread for every measure before any test. Percent of time and dwell pile up at zero (a participant who never looked at the Logo, or never reached the Joy threshold), so the median, the quartiles, and the percent at zero say things the mean hides.

**The formulas for the first row** (gaze dwell, Ad_A, Logo, row 5); the other rows are the same with their own row number. `B` and `C` hold the ad and the AOI, so the formulas never contain a typed name:

| Result | Cell | Formula |
|---|---|---|
| n | `Descriptives!D5` | `=COUNTIFS(AOI_long!$B$2:$B$145,B5,AOI_long!$C$2:$C$145,C5,AOI_long!$D$2:$D$145,1)` |
| Mean (AVERAGEIFS) | `Descriptives!E5` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$B$2:$B$145,B5,AOI_long!$C$2:$C$145,C5,AOI_long!$D$2:$D$145,1)` |
| SD (STDEV.S, array) | `Descriptives!F5` | `{=STDEV.S(IF((AOI_long!$B$2:$B$145=B5)*(AOI_long!$C$2:$C$145=C5)*(AOI_long!$D$2:$D$145=1),AOI_long!$E$2:$E$145))}` |
| Median (array) | `Descriptives!G5` | `{=MEDIAN(IF((AOI_long!$B$2:$B$145=B5)*(AOI_long!$C$2:$C$145=C5)*(AOI_long!$D$2:$D$145=1),AOI_long!$E$2:$E$145))}` |
| Q1 (QUARTILE.INC, array) | `Descriptives!H5` | `{=QUARTILE.INC(IF((AOI_long!$B$2:$B$145=B5)*(AOI_long!$C$2:$C$145=C5)*(AOI_long!$D$2:$D$145=1),AOI_long!$E$2:$E$145),1)}` |
| Q3 (QUARTILE.INC, array) | `Descriptives!I5` | `{=QUARTILE.INC(IF((AOI_long!$B$2:$B$145=B5)*(AOI_long!$C$2:$C$145=C5)*(AOI_long!$D$2:$D$145=1),AOI_long!$E$2:$E$145),3)}` |
| IQR | `Descriptives!J5` | `=I5-H5` |
| % at zero | `Descriptives!K5` | `=100*COUNTIFS(AOI_long!$B$2:$B$145,B5,AOI_long!$C$2:$C$145,C5,AOI_long!$D$2:$D$145,1,AOI_long!$E$2:$E$145,0)/D5` |

**Array formulas.** STDEV.S, MEDIAN, and QUARTILE.INC have no "IFS" version, so `IF(...)` builds the list of values that meet the conditions: `(range = B5)*(range = C5)*(include = 1)` is 1 only where all three conditions hold, and `IF` returns the dwell value there and FALSE elsewhere, which STDEV.S skips. In **Excel 365** type the formula and press **Enter**; in **Excel 2019 or older** press **Ctrl+Shift+Enter**, and Excel shows braces `{...}` around it (do not type the braces). **Common mistake:** pressing Enter in an older Excel: the formula then uses only one row and returns a wrong number without an error.

**Excel 365 alternative** without an array formula: `=STDEV.S(FILTER(AOI_long!$E$2:$E$145,(AOI_long!$B$2:$B$145=B5)*(AOI_long!$C$2:$C$145=C5)*(AOI_long!$D$2:$D$145=1)))`.

**A. Gaze dwell by ad and AOI** (percent of the 30 s ad)

| Measure | Ad or stimulus | n | Mean | SD | Median | Q1 | Q3 | IQR | % at zero |
|---|---|---|---|---|---|---|---|---|---|
| Logo | Ad_A | 22 | 1.782 | 2.820 | 0.525 | 0.000 | 2.095 | 2.095 | 45.45 |
| Logo | Ad_B | 22 | 1.396 | 1.891 | 0.750 | 0.000 | 2.040 | 2.040 | 45.45 |
| Product | Ad_A | 22 | 42.584 | 8.422 | 43.885 | 38.400 | 48.155 | 9.755 | 0.00 |
| Product | Ad_B | 22 | 41.344 | 6.168 | 41.530 | 38.515 | 45.610 | 7.095 | 0.00 |
| Headline | Ad_A | 22 | 17.844 | 7.119 | 19.195 | 13.393 | 22.930 | 9.537 | 0.00 |
| Headline | Ad_B | 22 | 20.267 | 6.442 | 21.750 | 15.930 | 23.790 | 7.860 | 0.00 |

**B. Facial expressions** (percent of time at or above threshold 50)

| Measure | Ad or stimulus | n | Mean | SD | Median | Q1 | Q3 | IQR | % at zero |
|---|---|---|---|---|---|---|---|---|---|
| Joy | Baseline | 22 | 0.379 | 1.422 | 0.000 | 0.000 | 0.000 | 0.000 | 90.91 |
| Joy | Ad_A | 22 | 12.131 | 10.808 | 9.665 | 3.192 | 18.975 | 15.783 | 13.64 |
| Joy | Ad_B | 22 | 8.329 | 9.464 | 3.780 | 0.000 | 16.475 | 16.475 | 36.36 |
| Engagement | Baseline | 22 | 2.330 | 2.360 | 1.385 | 0.140 | 4.122 | 3.982 | 27.27 |
| Engagement | Ad_A | 22 | 14.516 | 11.719 | 11.445 | 4.393 | 20.170 | 15.777 | 4.55 |
| Engagement | Ad_B | 22 | 9.950 | 10.014 | 6.110 | 0.360 | 19.135 | 18.775 | 27.27 |
| Brow Furrow | Baseline | 22 | 0.437 | 1.180 | 0.000 | 0.000 | 0.000 | 0.000 | 81.82 |
| Brow Furrow | Ad_A | 22 | 0.646 | 1.517 | 0.000 | 0.000 | 0.000 | 0.000 | 81.82 |
| Brow Furrow | Ad_B | 22 | 0.384 | 1.310 | 0.000 | 0.000 | 0.000 | 0.000 | 90.91 |

**C. GSR Peaks Per Minute** (valid GSR only, rule R3; the formulas add the condition `gsr_valid = 1`)

| Measure | Ad or stimulus | n | Mean | SD | Median | Q1 | Q3 | IQR | % at zero |
|---|---|---|---|---|---|---|---|---|---|
| GSR Peaks Per Minute | Baseline | 22 | 2.063 | 1.859 | 2.000 | 0.250 | 3.000 | 2.750 | 27.27 |
| GSR Peaks Per Minute | Ad_A | 22 | 7.273 | 6.805 | 6.000 | 2.000 | 8.000 | 6.000 | 4.55 |
| GSR Peaks Per Minute | Ad_B | 21 | 6.000 | 4.940 | 6.000 | 2.000 | 8.000 | 6.000 | 14.29 |

**D. Liking** (1 to 7)

| Measure | Ad or stimulus | n | Mean | SD | Median | Q1 | Q3 | IQR | % at zero |
|---|---|---|---|---|---|---|---|---|---|
| Liking | Ad_A | 22 | 4.591 | 1.403 | 5.000 | 4.000 | 5.000 | 1.000 |  |
| Liking | Ad_B | 22 | 4.045 | 1.290 | 4.000 | 3.000 | 5.000 | 2.000 |  |

**Read the tables.** The Product held 41.3 to 42.6 percent of the ad on average and the Logo less than 2 percent; 45 percent of participants never had their gaze on the Logo in each ad, so its median is below its mean. Brow Furrow was zero for most participants (its median and both quartiles are 0): a mean of 0.65 percent describes almost nobody.

---

## 4. The PivotTable (Pivot sheet)

The iMotions article Creating Excel Pivot Tables (local copy: Making Excel Pivot Tables) gives the steps for an AOI metrics table; on `AOI_long` they are:

1. Click any cell of `AOI_long`, then **Insert > PivotTable**, choose **New Worksheet**, **OK**.
2. In the **PivotTable Fields** pane drag `AOI` to **Rows**, `ad` to **Columns**, and `gaze dwell (% of ad)` to **Values**.
3. **You should see** "Sum of gaze dwell (% of ad)". Right click a number, **Value Field Settings**, choose **Average**, **OK**.
4. Drag `include (from Data)` to **Filters**, open the filter, tick only **1**, **OK**.
5. **You should see** the averages below, with the rows sorted A to Z (Headline, Logo, Product). Add a second copy of the dwell field to Values with **Count** to see n = 22 in every cell.

**Common mistakes:** leaving the value on **Sum** (it adds the participants' values together); forgetting the include filter (24 participants instead of 22); building the PivotTable from `Data` (three dwell columns cannot go into Columns as one field; that is why section 2 stacked them).

A PivotTable cannot be checked cell by cell, so the `Pivot` sheet builds the same table with AVERAGEIFS and COUNTIFS:

| Result | Cell | Formula | Value |
|---|---|---|---|
| Average dwell, Logo, Ad_A | `Pivot!B4` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$C$2:$C$145,$A4,AOI_long!$B$2:$B$145,B$3,AOI_long!$D$2:$D$145,1)` | 1.78 |
| Average dwell, Logo, Ad_B | `Pivot!C4` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$C$2:$C$145,$A4,AOI_long!$B$2:$B$145,C$3,AOI_long!$D$2:$D$145,1)` | 1.40 |
| Average dwell, Product, Ad_A | `Pivot!B5` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$C$2:$C$145,$A5,AOI_long!$B$2:$B$145,B$3,AOI_long!$D$2:$D$145,1)` | 42.58 |
| Average dwell, Product, Ad_B | `Pivot!C5` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$C$2:$C$145,$A5,AOI_long!$B$2:$B$145,C$3,AOI_long!$D$2:$D$145,1)` | 41.34 |
| Average dwell, Headline, Ad_A | `Pivot!B6` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$C$2:$C$145,$A6,AOI_long!$B$2:$B$145,B$3,AOI_long!$D$2:$D$145,1)` | 17.84 |
| Average dwell, Headline, Ad_B | `Pivot!C6` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$C$2:$C$145,$A6,AOI_long!$B$2:$B$145,C$3,AOI_long!$D$2:$D$145,1)` | 20.27 |
| Grand average, Logo (both ads) | `Pivot!D4` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$C$2:$C$145,$A4,AOI_long!$D$2:$D$145,1)` | 1.59 |
| Grand average, Product (both ads) | `Pivot!D5` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$C$2:$C$145,$A5,AOI_long!$D$2:$D$145,1)` | 41.96 |
| Grand average, Headline (both ads) | `Pivot!D6` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$C$2:$C$145,$A6,AOI_long!$D$2:$D$145,1)` | 19.06 |
| Grand average, Ad_A (all AOIs) | `Pivot!B7` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$B$2:$B$145,B$3,AOI_long!$D$2:$D$145,1)` | 20.74 |
| Grand average, Ad_B (all AOIs) | `Pivot!C7` | `=AVERAGEIFS(AOI_long!$E$2:$E$145,AOI_long!$B$2:$B$145,C$3,AOI_long!$D$2:$D$145,1)` | 21.00 |
| Count, Logo, Ad_A | `Pivot!E4` | `=COUNTIFS(AOI_long!$C$2:$C$145,$A4,AOI_long!$B$2:$B$145,"Ad_A",AOI_long!$D$2:$D$145,1)` | 22 |

The grand averages equal the estimated marginal means of the Q4 mixed model in `SIMULATED_q4_marginal_means.csv`, because every participant has every cell. The PivotTable shows WHAT the averages are; whether the AOI pattern differs between the ads is the Q4 interaction test, which needs R (`analysis_output.txt`, STEP 5: F(2, 126) = 1.14, p = .323).

---

## 5. Q5 with T.TEST (Q5 sheet)

Q5 asks whether Joy, Engagement, and Brow Furrow differ between Ad_A and Ad_B. Each participant saw both ads, so the tests are **paired**: each one works on the participant's own difference (Ad_A minus Ad_B). The roster in `Q5!A4:A25` lists the 22 participants with include = 1 (typed, yellow; the `Check` sheet confirms it). The values come from `Data` by key:

| Result | Cell | Formula |
|---|---|---|
| Joy, Ad_A, first participant | `Q5!B4` | `=INDEX(Data!$Q$2:$Q$73,MATCH($A4&"\|Ad_A",Data!$AD$2:$AD$73,0))` |
| Joy, Ad_B, first participant | `Q5!C4` | `=INDEX(Data!$Q$2:$Q$73,MATCH($A4&"\|Ad_B",Data!$AD$2:$AD$73,0))` |
| Difference | `Q5!D4` | `=B4-C4` |

Engagement and Brow Furrow use the same formulas in their own columns. **The results block for Joy** (the blocks for Engagement and Brow Furrow sit to its right and use their own columns):

| Result | Cell | Formula |
|---|---|---|
| Ad_A mean | `Q5!B29` | `=AVERAGE(B4:B25)` |
| Ad_A SD | `Q5!B30` | `=STDEV.S(B4:B25)` |
| Ad_B mean | `Q5!B31` | `=AVERAGE(C4:C25)` |
| Ad_B SD | `Q5!B32` | `=STDEV.S(C4:C25)` |
| n (participants) | `Q5!B33` | `=COUNT(D4:D25)` |
| Mean difference | `Q5!B34` | `=AVERAGE(D4:D25)` |
| SD of the differences | `Q5!B35` | `=STDEV.S(D4:D25)` |
| SE of the mean difference | `Q5!B36` | `=B35/SQRT(B33)` |
| t | `Q5!B37` | `=B34/B36` |
| df | `Q5!B38` | `=B33-1` |
| p, two tailed (T.TEST, paired) | `Q5!B39` | `=T.TEST(B4:B25,C4:C25,2,1)` |
| p from t (T.DIST.2T) | `Q5!B40` | `=T.DIST.2T(ABS(B37),B38)` |
| Critical t (95%) | `Q5!B41` | `=T.INV.2T(0.05,B38)` |
| Mean difference, 95% CI lower | `Q5!B42` | `=B34-B41*B36` |
| Mean difference, 95% CI upper | `Q5!B43` | `=B34+B41*B36` |
| Cohen's d_z | `Q5!B44` | `=B34/B35` |
| SE of d_z (approximation) | `Q5!B45` | `=SQRT(1/B33+B44^2/(2*B33))` |
| d_z 95% CI lower (approximation) | `Q5!B46` | `=B44-1.96*B45` |
| d_z 95% CI upper (approximation) | `Q5!B47` | `=B44+1.96*B45` |
| Ad_A higher (count) | `Q5!B48` | `=COUNTIF(D4:D25,">0")` |
| Ad_B higher (count) | `Q5!B49` | `=COUNTIF(D4:D25,"<0")` |
| No difference (count) | `Q5!B50` | `=COUNTIF(D4:D25,0)` |
| Zero in both ads (count) | `Q5!B51` | `=COUNTIFS(B4:B25,0,C4:C25,0)` |

**All three tests:**

| Result | Joy cell | Joy | Engagement cell | Engagement | Brow Furrow cell | Brow Furrow |
|---|---|---|---|---|---|---|
| Ad_A mean | `Q5!B29` | 12.131 | `Q5!F29` | 14.516 | `Q5!J29` | 0.646 |
| Ad_A SD | `Q5!B30` | 10.808 | `Q5!F30` | 11.719 | `Q5!J30` | 1.517 |
| Ad_B mean | `Q5!B31` | 8.329 | `Q5!F31` | 9.950 | `Q5!J31` | 0.384 |
| Ad_B SD | `Q5!B32` | 9.464 | `Q5!F32` | 10.014 | `Q5!J32` | 1.310 |
| n | `Q5!B33` | 22 | `Q5!F33` | 22 | `Q5!J33` | 22 |
| Mean difference | `Q5!B34` | 3.803 | `Q5!F34` | 4.566 | `Q5!J34` | 0.262 |
| SD of the differences | `Q5!B35` | 12.078 | `Q5!F35` | 12.847 | `Q5!J35` | 1.648 |
| SE of the mean difference | `Q5!B36` | 2.5751 | `Q5!F36` | 2.7389 | `Q5!J36` | 0.3513 |
| t | `Q5!B37` | 1.477 | `Q5!F37` | 1.667 | `Q5!J37` | 0.747 |
| df | `Q5!B38` | 21 | `Q5!F38` | 21 | `Q5!J38` | 21 |
| p (T.TEST, paired) | `Q5!B39` | 0.15459 | `Q5!F39` | 0.11032 | `Q5!J39` | 0.46359 |
| p from t (T.DIST.2T) | `Q5!B40` | 0.15459 | `Q5!F40` | 0.11032 | `Q5!J40` | 0.46359 |
| Critical t | `Q5!B41` | 2.0796 | `Q5!F41` | 2.0796 | `Q5!J41` | 2.0796 |
| 95% CI lower | `Q5!B42` | -1.553 | `Q5!F42` | -1.130 | `Q5!J42` | -0.468 |
| 95% CI upper | `Q5!B43` | 9.158 | `Q5!F43` | 10.262 | `Q5!J43` | 0.993 |
| Cohen's d_z | `Q5!B44` | 0.315 | `Q5!F44` | 0.355 | `Q5!J44` | 0.159 |
| SE of d_z | `Q5!B45` | 0.2184 | `Q5!F45` | 0.2198 | `Q5!J45` | 0.2145 |
| d_z CI lower (approx.) | `Q5!B46` | -0.113 | `Q5!F46` | -0.075 | `Q5!J46` | -0.261 |
| d_z CI upper (approx.) | `Q5!B47` | 0.743 | `Q5!F47` | 0.786 | `Q5!J47` | 0.580 |
| Ad_A higher | `Q5!B48` | 12 | `Q5!F48` | 15 | `Q5!J48` | 3 |
| Ad_B higher | `Q5!B49` | 7 | `Q5!F49` | 6 | `Q5!J49` | 2 |
| No difference | `Q5!B50` | 3 | `Q5!F50` | 1 | `Q5!J50` | 17 |
| Zero in both ads | `Q5!B51` | 3 | `Q5!F51` | 1 | `Q5!J51` | 17 |

**Read the results.** T.TEST returns only the p value; its arguments are (range 1, range 2, tails = 2, type = 1 for paired). The hand computed t and T.DIST.2T must give the same p, which confirms that the two ranges are aligned row by row. **Common mistake:** type = 2 or 3 (two independent groups) ignores the pairing and gives a different, wrong p. R adds what Excel has no function for: the Wilcoxon signed rank checks (Joy p = .184, Engagement p = .099, Brow Furrow p = .590) and the exact noncentral t intervals of d_z (Joy [-0.12, 0.74], Engagement [-0.08, 0.78], Brow Furrow [-0.26, 0.58]). Brow Furrow was zero in both ads for 17 of 22 participants: its t test rests on 5 nonzero differences, and the Wilcoxon test is the better guide.

---

## 6. The Holm correction (Holm sheet)

Three tests on the same question (do the ads differ in facial expression?) give three chances of a false positive. Holm's procedure (Holm, 1979) keeps the chance of at least one false positive in the family at .05: sort the p values from smallest to largest, multiply the smallest by m (the number of tests), the next by m minus 1, and so on, and never let an adjusted value fall below the one before it.

**The step table** (rows 11 to 13; m in `Holm!K3` = `=COUNT(B4:B6)`):

| Step k | Cell | k-th smallest p | Value | Multiplier | p x multiplier | Value | Holm adjusted | Value |
|---|---|---|---|---|---|---|---|---|
| 1 | `Holm!B11` | `=SMALL($B$4:$B$6,A11)` | 0.11032 | `=$K$3-A11+1` | `=B11*C11` | 0.33096 | `=MIN(1,D11)` | 0.33096 |
| 2 | `Holm!B12` | `=SMALL($B$4:$B$6,A12)` | 0.15459 | `=$K$3-A12+1` | `=B12*C12` | 0.30919 | `=MIN(1,MAX(E11,D12))` | 0.33096 |
| 3 | `Holm!B13` | `=SMALL($B$4:$B$6,A13)` | 0.46359 | `=$K$3-A13+1` | `=B13*C13` | 0.46359 | `=MIN(1,MAX(E12,D13))` | 0.46359 |

**Back to each test** (rows 4 to 6):

| Result | Cell | Formula |
|---|---|---|
| p, linked from Q5 | `Holm!B4` | `=Q5!B39` |
| Rank (1 = smallest) | `Holm!C4` | `=RANK.EQ(B4,$B$4:$B$6,1)` |
| Multiplier | `Holm!D4` | `=$K$3-C4+1` |
| p x multiplier | `Holm!E4` | `=B4*D4` |
| Holm adjusted p | `Holm!F4` | `=INDEX($E$11:$E$13,C4)` |
| Bonferroni p | `Holm!G4` | `=MIN(1,$K$3*B4)` |
| Decision | `Holm!H4` | `=IF(F4<0.05,"significant","not significant")` |

| Test | p | Rank | Multiplier | p x multiplier | Holm p | Bonferroni p | Decision |
|---|---|---|---|---|---|---|---|
| Joy | 0.15459 | 2 | 2 | 0.30919 | 0.33096 | 0.46378 | not significant |
| Engagement | 0.11032 | 1 | 3 | 0.33096 | 0.33096 | 0.33096 | not significant |
| Brow Furrow | 0.46359 | 3 | 1 | 0.46359 | 0.46359 | 1.00000 | not significant |

**Read the table.** The order from the smallest p is Engagement (p = 0.11032, multiplied by 3), then Joy (p = 0.15459, multiplied by 2), then Brow Furrow (p = 0.46359, multiplied by 1). The running maximum matters here: Joy's product 0.30919 is raised to 0.33096, the adjusted value before it. **Common mistake:** reporting p x multiplier without the running maximum, which can make a later test look stronger than an earlier one. None of the three differences survives the correction (and none was significant before it). That does NOT show that the ads produce the same expressions: across the three expressions the 95% intervals of d_z run from -0.26 to 0.78, so they include both no difference and a moderate advantage for Ad_A; 22 participants cannot tell these apart.

The R function `p.adjust(p, method = "holm")` gives the same numbers in one line (`analysis_output.txt`, STEP 6).

---

## 7. Check your numbers (Check sheet)

Every result printed above sits on the `Check` sheet next to the value that `analysis_full.R` wrote to `SIMULATED_descriptives.csv`, `SIMULATED_all_tests.csv`, `SIMULATED_q4_marginal_means.csv`, or `analysis_output.txt`, with the difference and a tolerance of half a unit in the last printed decimal. `Check!B261` counts the rows with status CHECK and must read 0; in the LibreOffice recalculation it reads 0 for 256 rows, and the largest absolute difference between Excel and R is 4.7e-05. The last two rows confirm the Q5 roster: every listed participant has include = 1, and the roster holds all 22.

R only: H1 to H3 (Workshop 7), the Q4 mixed model and its checks, the Q4 and Q6 paired tests, the order checks, the Wilcoxon checks, the exact d_z intervals, the Spearman correlations, and the power analysis.

---

## 8. Common mistakes, collected

1. Averaging over all 24 participants: every COUNTIFS, AVERAGEIFS, and IF condition needs `include = 1`.
2. Pressing Enter on an array formula in Excel 2019 or older (no braces): the result is wrong without an error.
3. Reporting only the mean of a measure that piles up at zero: add the median and the percent at zero.
4. A PivotTable left on **Sum**, or built from the wide `Data` sheet instead of the stacked `AOI_long` sheet.
5. T.TEST with type 2 or 3 for a paired design.
6. Holm without the running maximum, or Bonferroni when Holm was planned (Bonferroni is never more powerful).
7. Reading "not significant" as "no difference": check the confidence interval of the difference and of d_z.
8. Typing values by hand instead of linking them (`=Q5!B39`, not `0.15459`): a typed value does not update when the data change (Broman & Woo, 2018).

---

## Sources

- Broman, K. W., & Woo, K. H. (2018). Data organization in spreadsheets. *The American Statistician, 72*(1), 2-10. https://doi.org/10.1080/00031305.2017.1375989
- Holm, S. (1979). A simple sequentially rejective multiple test procedure. *Scandinavian Journal of Statistics, 6*(2), 65-70.
- Lakens, D. (2013). Calculating and reporting effect sizes to facilitate cumulative science: A practical primer for t-tests and ANOVAs. *Frontiers in Psychology, 4*, Article 863. https://doi.org/10.3389/fpsyg.2013.00863
- iMotions help center, Creating Excel Pivot Tables (local PDF: Making Excel Pivot Tables): https://help.imotions.com/lab/exports/additional-export-information/creating-excel-pivot-tables

## Items to verify on the lab computer

1. The Excel version on the lab computers (Excel 365 takes array formulas with Enter and has FILTER and XLOOKUP; Excel 2019 or older needs Ctrl+Shift+Enter and has neither).
2. The Windows regional settings (expected: dot decimal, comma list separator); with other settings the formulas use semicolons between arguments.
