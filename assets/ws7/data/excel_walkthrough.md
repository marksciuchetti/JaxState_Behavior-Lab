# Excel walkthrough: from iMotions exports to one analysis ready table

**SIMULATED DATA.** Every number in this walkthrough comes from the SIMULATED export bundle that `make_simulated_data.py` generated (seed 20271231). No real person was recorded, and the numbers must never be quoted as findings. The study is "SIMULATED Multimodal Ad Study": 24 participants (P01 to P24) at the lab desktop with the Smart Eye Aurora, the respondent camera with AFFDEX, and a Shimmer3 GSR+; slides Instructions, Baseline (60 s), Ad_A and Ad_B (30 s video ads, order AB or BA), and a survey.

This is the Excel route through the core of `prepare_data.R` and the H3 test of `analysis.R`. The finished workbook is `SIMULATED_data_prep.xlsx`. **Every cell address below is written as `Sheet!Cell` and was generated from the workbook itself:** `make_excel_workbook.py` writes the workbook and this file from one list of cells, reads the workbook back with openpyxl to confirm that each address holds the formula printed here, and recalculates a copy to confirm the results. A copy was recalculated with LibreOffice 24.2.7.2 420(Build:2) (headless): all 52 Check rows are OK, no cell shows an error value, and the 50 results printed here agree with R at the printed decimals.

Companion files: `prepare_data.R` and `prepare_output.txt` (the R pipeline and its full output), `analysis.R` and `analysis_output.txt` (the tests), `README.md` (the bundle), `results_text.md` (model paragraphs), `figures/`.

---

## 0. What you are working with

| Sheet | What it holds | Rows |
|---|---|---|
| `GSR_Import` | `exports/04_GSR_Metrics/GSRPeakMetrics.csv` as the Text Import leaves it; helper columns `P` to `U` (orange) | 170 |
| `SessionLog` | `exports/08_Session_Log/SIMULATED_AdStudy_SessionLog.csv`, one row per recording; helper column `N` | 25 |
| `Survey` | `exports/05_Survey/MERGED_SURVEY_RESPONSE_MATRIX.txt` (tab separated); helper columns `V` to `AC`, Cronbach's alpha in `AF2` and `AF3` | 24 |
| `DataQuality` | `exports/06_Data_Quality/SIMULATED Multimodal Ad Study_DataQuality.txt` (tab separated); helper columns `I` to `M` | 122 |
| `Keys` | the problem keys made visible | 5 |
| `Merged` | one row per participant and stimulus, joined with INDEX and MATCH | 72 |
| `Pivot` | the PivotTable summary, built with AVERAGEIFS and COUNTIFS | 3 |
| `H3` | GSR Peaks Per Minute, each ad versus Baseline, with T.TEST | 22 and 21 |
| `Check` | every Excel result next to the R value | 52 |

Blue headers mark data as exported; orange headers mark columns you add; yellow cells are typed by you (the participant and stimulus grid).

**Rules written before collection (study protocol).** R1: use only the recordings that the session log marks Use recording = Yes. R2: exclude a participant from all analyses when Eye tracking_Quality is below 90 percent or Affectiva AFFDEX_Quality is below 80 percent on Ad_A or Ad_B. R3: for GSR analyses only, a stimulus is missing when Signal Duration (ms) is below 80 percent of Duration (ms).

---

## 1. Text Import of one metrics file (GSRPeakMetrics.csv)

Row numbers in this walkthrough count the SIMULATED banner in line 1 of every file; a real export has no banner, so subtract 1. iMotions writes its metrics files with information ABOVE the table. In `GSRPeakMetrics.csv` line 1 is the SIMULATED banner (a real export has no banner), lines 2 to 9 are the `#METADATA` block, line 10 reads `#DATA`, and the column names are on **line 11**. Double clicking the file makes Excel guess, and on a computer with other regional settings everything lands in column A (Opening Data Export Files in Excel article).

**Route A, Text Import Wizard (any Excel version).**
1. **File > Open > Browse**, set the file type to **All Files (\*.\*)**, select `GSRPeakMetrics.csv`, **Open**.
2. Step 1 of 3: choose **Delimited**, set **Start import at row: 11**, File origin **65001 : Unicode (UTF-8)**, tick **My data has headers**. **You should see** `Study Name,Respondent Name,...` as the first preview line.
3. Step 2 of 3: tick **Comma** only. **You should see** the preview split into columns.
4. Step 3 of 3: select the `Respondent Name` column and set it to **Text** (so that nothing reformats a key); click **Advanced...** and set **Decimal separator** to `.` and **Thousands separator** to `,`; **Finish**.
5. **You should see** 15 columns (`A` = Study Name to `O` = Average Peak Amplitude (microSiemens)) and 170 data rows, header in row 1. Cells that read `NA` are text: iMotions writes NA where a value does not exist (no peaks, or no signal). Leave them as they are.

**Route B, Get Data (Excel 2016 and later).** **Data > From Text/CSV**, select the file, **Transform Data**; in Power Query **Home > Remove Rows > Remove Top Rows** (10), then **Home > Use First Row as Headers**, set `Respondent Name` to **Text**, **Close & Load**.

**The other three files** go on their own sheets the same way: `MERGED_SURVEY_RESPONSE_MATRIX.txt` with **Tab** as the delimiter, **Text qualifier: {none}**, start at row 2 (its headers contain commas and quotation marks, so a comma delimiter breaks them apart); `SIMULATED Multimodal Ad Study_DataQuality.txt` with **Tab**, start at row 4 (two `#Study ...` lines and the banner sit above the header); `SIMULATED_AdStudy_SessionLog.csv` with **Comma**, start at row 2.

**Common mistakes.** (1) Starting the import at row 1: the `#METADATA` lines become data and every column turns to text. (2) Leaving the delimiter on Tab for a comma file: one column. (3) Saving the result back as .csv: formulas, helper columns, and extra sheets are lost; save as `.xlsx` and keep the raw export untouched in its folder. (4) Opening the student's copy `SurveyStatistics_student_copy.csv`: it was saved again on a laptop with German regional settings (semicolons, decimal commas). Never analyze a resaved copy; import the raw export.

---

## 2. The session log is the key table

Every iMotions export names people by the **respondent name** typed at the lab computer. The study is about **participants** (P01 to P24). The session log connects the two, one row per recording, and says which recording to use. P04 was recorded twice: `P04` (stopped 23 s into Baseline after the Aurora calibration failed; Use recording = No) and `P04_rerun` (Use recording = Yes).

Add one helper column to `SessionLog`, a clean version of the respondent name:

| Result | Cell | Formula | Value |
|---|---|---|---|
| Key of the first recording | `SessionLog!N2` | `=UPPER(TRIM(B2))` | P01 |

Fill it down to row 26. **You should see** `P04` and `P04_RERUN` as two different keys in rows 5 and 6.

---

## 3. Fix the keys with TRIM and UPPER (Keys sheet)

Before fixing a key, **look at it**. The `Keys` sheet pulls five raw keys from the imported sheets by reference and tests each one:

| Raw key | Where | LEN | CODE of last character | EXACT in the session log | MATCH finds it | UPPER(TRIM()) | Participant | Use |
|---|---|---|---|---|---|---|---|---|
| `"P07 "` | DataQuality, column A | 4 | 32 | FALSE | FALSE | P07 | P07 | Yes |
| `"p07"` | Survey, column B | 3 | 55 | FALSE | TRUE | P07 | P07 | Yes |
| `"P04"` | GSR_Import, column B | 3 | 52 | TRUE | TRUE | P04 | P04 | No |
| `"P04_rerun"` | GSR_Import, column B | 9 | 110 | TRUE | TRUE | P04_RERUN | P04 | Yes |
| `"P01"` | GSR_Import, column B (a normal key) | 3 | 49 | TRUE | TRUE | P01 | P01 | Yes |

The formulas in row 4 (the other rows are the same with their own row number):

| Test | Cell | Formula |
|---|---|---|
| The raw key (a reference) | `Keys!B4` | `=DataQuality!A34` |
| Number of characters | `Keys!C4` | `=LEN(B4)` |
| Code of the last character (32 = space) | `Keys!D4` | `=CODE(RIGHT(B4,1))` |
| Case sensitive match | `Keys!E4` | `=SUMPRODUCT(--EXACT(B4,SessionLog!$B$2:$B$26))>0` |
| What MATCH does | `Keys!F4` | `=ISNUMBER(MATCH(B4,SessionLog!$B$2:$B$26,0))` |
| The standardized key | `Keys!G4` | `=UPPER(TRIM(B4))` |
| Participant from the session log | `Keys!H4` | `=IFERROR(INDEX(SessionLog!$A$2:$A$26,MATCH(G4,SessionLog!$N$2:$N$26,0)),"NOT FOUND")` |
| Use recording | `Keys!I4` | `=IFERROR(INDEX(SessionLog!$L$2:$L$26,MATCH(G4,SessionLog!$N$2:$N$26,0)),"NOT FOUND")` |

**Read the table.** `"P07 "` has 4 characters and its last character is code 32, a space: invisible on screen, fatal for a lookup. `"p07"` fails the case sensitive EXACT test, yet MATCH finds it: **Excel lookups (MATCH, VLOOKUP, XLOOKUP) ignore upper and lower case, R joins do not.** A trailing space breaks both. `P04_rerun` is a correct name of the second recording, and the session log maps it to participant P04.

Now give every imported sheet a clean key and the participant (orange columns, filled down):

| Sheet | Result | Cell | Formula |
|---|---|---|---|
| `GSR_Import` | Key | `GSR_Import!P2` | `=UPPER(TRIM(B2))` |
| `GSR_Import` | Participant | `GSR_Import!Q2` | `=IFERROR(INDEX(SessionLog!$A$2:$A$26,MATCH(P2,SessionLog!$N$2:$N$26,0)),"NOT FOUND")` |
| `GSR_Import` | Use recording | `GSR_Import!R2` | `=IFERROR(INDEX(SessionLog!$L$2:$L$26,MATCH(P2,SessionLog!$N$2:$N$26,0)),"NOT FOUND")` |
| `GSR_Import` | Lookup key | `GSR_Import!S2` | `=IF(R2="Yes",Q2&"\|"&F2&"\|"&G2&"\|"&K2,"DROP")` |
| `Survey` | Key | `Survey!V2` | `=UPPER(TRIM(B2))` |
| `Survey` | Participant | `Survey!W2` | `=IFERROR(INDEX(SessionLog!$A$2:$A$26,MATCH(V2,SessionLog!$N$2:$N$26,0)),"NOT FOUND")` |
| `DataQuality` | Key | `DataQuality!I2` | `=UPPER(TRIM(A2))` |
| `DataQuality` | Participant | `DataQuality!J2` | `=IFERROR(INDEX(SessionLog!$A$2:$A$26,MATCH(I2,SessionLog!$N$2:$N$26,0)),"NOT FOUND")` |
| `DataQuality` | Use recording | `DataQuality!K2` | `=IFERROR(INDEX(SessionLog!$L$2:$L$26,MATCH(I2,SessionLog!$N$2:$N$26,0)),"NOT FOUND")` |
| `DataQuality` | Lookup key | `DataQuality!L2` | `=IF(K2="Yes",J2&"\|"&B2,"DROP")` |
| `DataQuality` | LEN of the raw name | `DataQuality!M2` | `=LEN(A2)` |

The **lookup key** joins the participant with the other key columns. In `GSR_Import` it is participant, Type, Label, and Parent Stimulus, because Label alone is not unique: the annotation `EndCard` appears once in Ad_A and once in Ad_B. Rows of the dropped recording get the key `DROP`, so no lookup can ever return them (rule R1). **You should see** `NOT FOUND` nowhere: filter the Participant columns for it. **You should see** `DROP` in exactly 2 rows of `GSR_Import` and 2 rows of `DataQuality` (the Instructions and Baseline rows of the aborted recording).

**Excel 365 and Excel 2021:** each INDEX and MATCH pair can be written as XLOOKUP, for example `=XLOOKUP(P2,SessionLog!$N$2:$N$26,SessionLog!$A$2:$A$26,"NOT FOUND")` for the participant. The workbook uses INDEX and MATCH because every Excel version and LibreOffice understand it.

---

## 4. Join the survey and the data quality file to the GSR metrics (Merged sheet)

The `Merged` sheet is the analysis ready table: one row per participant and stimulus, 24 participants times 3 stimuli = 72 rows. Type the grid first (yellow columns `A` and `B`: P01 Baseline, P01 Ad_A, P01 Ad_B, P02 Baseline, ...), then fill these formulas down from row 2 to row 73:

| Column | Result | Cell | Formula |
|---|---|---|---|
| `C` | GSR lookup key | `Merged!C2` | `=A2&"\|Stimulus\|"&B2&"\|"&B2` |
| `D` | GSR Peaks Per Minute (as exported) | `Merged!D2` | `=INDEX(GSR_Import!$N$2:$N$171,MATCH(C2,GSR_Import!$S$2:$S$171,0))` |
| `E` | GSR valid (R3) | `Merged!E2` | `=INDEX(GSR_Import!$U$2:$U$171,MATCH(C2,GSR_Import!$S$2:$S$171,0))` |
| `F` | GSR Peaks Per Minute (analysis) | `Merged!F2` | `=IF(E2=1,D2,"")` |
| `G` | DQ lookup key | `Merged!G2` | `=A2&"\|"&B2` |
| `H` | Eye tracking quality (%) | `Merged!H2` | `=INDEX(DataQuality!$D$2:$D$123,MATCH(G2,DataQuality!$L$2:$L$123,0))` |
| `I` | Face quality (%) | `Merged!I2` | `=INDEX(DataQuality!$F$2:$F$123,MATCH(G2,DataQuality!$L$2:$L$123,0))` |
| `J` | Eye tracking minimum, ads | `Merged!J2` | `=MIN(INDEX(DataQuality!$D$2:$D$123,MATCH(A2&"\|Ad_A",DataQuality!$L$2:$L$123,0)),INDEX(DataQuality!$D$2:$D$123,MATCH(A2&"\|Ad_B",DataQuality!$L$2:$L$123,0)))` |
| `K` | Face minimum, ads | `Merged!K2` | `=MIN(INDEX(DataQuality!$F$2:$F$123,MATCH(A2&"\|Ad_A",DataQuality!$L$2:$L$123,0)),INDEX(DataQuality!$F$2:$F$123,MATCH(A2&"\|Ad_B",DataQuality!$L$2:$L$123,0)))` |
| `L` | include (R2) | `Merged!L2` | `=IF(AND(J2>=90,K2>=80),1,0)` |
| `M` | liking | `Merged!M2` | `=IF(B2="Ad_A",INDEX(Survey!$G$2:$G$25,MATCH(A2,Survey!$W$2:$W$25,0)),IF(B2="Ad_B",INDEX(Survey!$I$2:$I$25,MATCH(A2,Survey!$W$2:$W$25,0)),""))` |
| `N` | recall_correct | `Merged!N2` | `=IF(B2="Ad_A",INDEX(Survey!$X$2:$X$25,MATCH(A2,Survey!$W$2:$W$25,0)),IF(B2="Ad_B",INDEX(Survey!$Y$2:$Y$25,MATCH(A2,Survey!$W$2:$W$25,0)),""))` |

Values in the first ad row (P01, Ad_A, row 3):

| Result | Cell | Formula | Value |
|---|---|---|---|
| GSR Peaks Per Minute (as exported) | `Merged!D3` | `=INDEX(GSR_Import!$N$2:$N$171,MATCH(C3,GSR_Import!$S$2:$S$171,0))` | 6 |
| Liking of Ad_A (survey) | `Merged!M3` | `=IF(B3="Ad_A",INDEX(Survey!$G$2:$G$25,MATCH(A3,Survey!$W$2:$W$25,0)),IF(B3="Ad_B",INDEX(Survey!$I$2:$I$25,MATCH(A3,Survey!$W$2:$W$25,0)),""))` | 4 |
| Brand recalled, Ad_A (survey, scored in section 6) | `Merged!N3` | `=IF(B3="Ad_A",INDEX(Survey!$X$2:$X$25,MATCH(A3,Survey!$W$2:$W$25,0)),IF(B3="Ad_B",INDEX(Survey!$Y$2:$Y$25,MATCH(A3,Survey!$W$2:$W$25,0)),""))` | 0 |

**XLOOKUP version** (Excel 365) of column `D`: `=XLOOKUP(C2,GSR_Import!$S$2:$S$171,GSR_Import!$N$2:$N$171)`; of column `M`: `=IF(B2="Ad_A",XLOOKUP(A2,Survey!$W$2:$W$25,Survey!$G$2:$G$25),IF(B2="Ad_B",XLOOKUP(A2,Survey!$W$2:$W$25,Survey!$I$2:$I$25),""))`.

**Check the join** (cells on the right of `Merged`):

| Check | Cell | Formula | Value |
|---|---|---|---|
| Rows (participant x stimulus) | `Merged!Q3` | `=COUNTA(A2:A73)` | 72 |
| Lookups that failed (#N/A) | `Merged!Q4` | `=SUMPRODUCT(--ISNA(D2:D73))+SUMPRODUCT(--ISNA(H2:H73))+SUMPRODUCT(--ISNA(I2:I73))` | 0 |
| Participants with include = 1 | `Merged!Q5` | `=COUNTIFS(B2:B73,"Baseline",L2:L73,1)` | 22 |
| Stimulus rows with GSR not valid | `Merged!Q6` | `=COUNTIF(E2:E73,0)` | 1 |

**You should see** 0 failed lookups. Try it the wrong way once, in a copy: build the data quality lookup key in `DataQuality!L2` from the raw name in column `A` instead of the participant in column `J`. P07 (raw name `"P07 "`) and P04 (whose data quality rows are named `P04_rerun`) then return `#N/A`, and a filter on quality would quietly drop them.

**Rules R2 and R3 in the table.** Columns `J` and `K` take the LOWER of the two ads for each participant; column `L` applies rule R2. P18, P21 get include = 0. Column `E` applies rule R3 through the helper columns of `GSR_Import`:

| Result | Cell | Formula |
|---|---|---|
| Signal share | `GSR_Import!T2` | `=IF(I2>0,J2/I2,0)` |
| GSR valid (R3) | `GSR_Import!U2` | `=IF(T2>=0.8,1,0)` |

P05's Shimmer lost its battery 7.8 s into Ad_B, so row 16 of `Merged` (P05, Ad_B) shows GSR valid = 0 and an empty cell in column `F`: the exported 7.680 peaks per minute were computed on 7.8 s of signal. P05 stays in every other analysis.

---

## 5. A pivot table summary (Pivot sheet)

The iMotions article Creating Excel Pivot Tables gives the steps: click a cell of the data, **Insert > PivotTable > New Worksheet > OK**, drag `stimulus` to **Rows**, drag `GSR Peaks Per Minute (analysis)` to **Values**, right click the value, **Value Field Settings**, change **Sum** to **Average**, **OK**. Then drag `include (R2)` to **Filters** and keep only 1. **Common mistake:** the default is **Sum**, which adds the participants' values together; always switch to Average, and use Count to check n.

A PivotTable cannot be checked cell by cell, so the `Pivot` sheet builds the same table with AVERAGEIFS and COUNTIFS:

| Result | Cell | Formula | Value |
|---|---|---|---|
| n (GSR), Baseline | `Pivot!B4` | `=COUNTIFS(Merged!$B$2:$B$73,A4,Merged!$L$2:$L$73,1,Merged!$E$2:$E$73,1)` | 22 |
| Mean GSR Peaks Per Minute, Baseline | `Pivot!C4` | `=AVERAGEIFS(Merged!$D$2:$D$73,Merged!$B$2:$B$73,A4,Merged!$L$2:$L$73,1,Merged!$E$2:$E$73,1)` | 2.063 |
| n (GSR), Ad_A | `Pivot!B5` | `=COUNTIFS(Merged!$B$2:$B$73,A5,Merged!$L$2:$L$73,1,Merged!$E$2:$E$73,1)` | 22 |
| Mean GSR Peaks Per Minute, Ad_A | `Pivot!C5` | `=AVERAGEIFS(Merged!$D$2:$D$73,Merged!$B$2:$B$73,A5,Merged!$L$2:$L$73,1,Merged!$E$2:$E$73,1)` | 7.273 |
| n (GSR), Ad_B | `Pivot!B6` | `=COUNTIFS(Merged!$B$2:$B$73,A6,Merged!$L$2:$L$73,1,Merged!$E$2:$E$73,1)` | 21 |
| Mean GSR Peaks Per Minute, Ad_B | `Pivot!C6` | `=AVERAGEIFS(Merged!$D$2:$D$73,Merged!$B$2:$B$73,A6,Merged!$L$2:$L$73,1,Merged!$E$2:$E$73,1)` | 6.000 |
| Mean liking, Ad_A | `Pivot!E5` | `=AVERAGEIFS(Merged!$M$2:$M$73,Merged!$B$2:$B$73,A5,Merged!$L$2:$L$73,1)` | 4.591 |
| Brand recalled (proportion), Ad_A | `Pivot!F5` | `=AVERAGEIFS(Merged!$N$2:$N$73,Merged!$B$2:$B$73,A5,Merged!$L$2:$L$73,1)` | 0.545 |
| Mean liking, Ad_B | `Pivot!E6` | `=AVERAGEIFS(Merged!$M$2:$M$73,Merged!$B$2:$B$73,A6,Merged!$L$2:$L$73,1)` | 4.045 |
| Brand recalled (proportion), Ad_B | `Pivot!F6` | `=AVERAGEIFS(Merged!$N$2:$N$73,Merged!$B$2:$B$73,A6,Merged!$L$2:$L$73,1)` | 0.545 |

**You should see** n = 21 for Ad_B and 22 for the other two stimuli: rule R3 removed P05's Ad_B value. Your PivotTable must show the same means.

---

## 6. Score the survey (Survey sheet)

The survey export holds the **label value** that was chosen, not a score. Two items need scoring before analysis:

| Result | Cell | Formula |
|---|---|---|
| Recall_A correct (Label1 = ALTAVO) | `Survey!X2` | `=IF(K2=1,1,0)` |
| Recall_B correct (Label2 = BRENLOW) | `Survey!Y2` | `=IF(M2=2,1,0)` |
| ATT_3R reverse coded (8 minus the answer) | `Survey!Z2` | `=8-S2` |
| Attitude (mean of three items) | `Survey!AA2` | `=AVERAGE(O2,Q2,Z2)` |
| Sum of the three items (recoded) | `Survey!AB2` | `=O2+Q2+Z2` |
| Sum of the three items (not recoded) | `Survey!AC2` | `=O2+Q2+S2` |

**Cronbach's alpha** = k / (k minus 1) x (1 minus the sum of the item variances / the variance of the sum), with k = 3 items:

| Result | Cell | Formula | Value |
|---|---|---|---|
| Cronbach's alpha, ATT_3R recoded | `Survey!AF2` | `=3/2*(1-(VAR.S(O2:O25)+VAR.S(Q2:Q25)+VAR.S(Z2:Z25))/VAR.S(AB2:AB25))` | 0.839 |
| Cronbach's alpha, ATT_3R NOT recoded (the mistake) | `Survey!AF3` | `=3/2*(1-(VAR.S(O2:O25)+VAR.S(Q2:Q25)+VAR.S(S2:S25))/VAR.S(AC2:AC25))` | -1.101 |

A negative alpha is the classic sign of a reverse worded item that was not recoded. **Common mistake:** averaging the Recall label values (SurveyStatistics.txt does exactly that: an "Average" of 2.42 for Recall_A means nothing, because the labels are brand names).

---

## 7. H3 with T.TEST (H3 sheet)

H3: GSR Peaks Per Minute are higher during each ad than during Baseline. Each participant gives a Baseline value and an ad value, so the test is **paired**. The rosters in `H3!A5:A26` (Ad_A, 22 participants) and `H3!G5:G25` (Ad_B, 21 participants; P05 is missing because of rule R3) list the participants with include = 1 and valid GSR in both stimuli; the `Check` sheet confirms each list. The values come from `Merged` column `F`:

| Result | Cell | Formula |
|---|---|---|
| Baseline, first participant | `H3!B5` | `=INDEX(Merged!$F$2:$F$73,MATCH(A5&"\|Stimulus\|Baseline\|Baseline",Merged!$C$2:$C$73,0))` |
| Ad_A, first participant | `H3!C5` | `=INDEX(Merged!$F$2:$F$73,MATCH(A5&"\|Stimulus\|Ad_A\|Ad_A",Merged!$C$2:$C$73,0))` |
| Difference | `H3!D5` | `=C5-B5` |

**Ad_A versus Baseline**

| Result | Cell | Formula | Value |
|---|---|---|---|
| Baseline mean | `H3!B30` | `=AVERAGE(B5:B26)` | 2.063 |
| Baseline SD | `H3!B31` | `=STDEV.S(B5:B26)` | 1.859 |
| Ad_A mean | `H3!B32` | `=AVERAGE(C5:C26)` | 7.273 |
| Ad_A SD | `H3!B33` | `=STDEV.S(C5:C26)` | 6.805 |
| n | `H3!B34` | `=COUNT(D5:D26)` | 22 |
| Mean difference | `H3!B35` | `=AVERAGE(D5:D26)` | 5.210 |
| SD of the differences | `H3!B36` | `=STDEV.S(D5:D26)` | 6.308 |
| SE of the mean difference | `H3!B37` | `=B36/SQRT(B34)` | 1.345 |
| t | `H3!B38` | `=B35/B37` | 3.874 |
| df | `H3!B39` | `=B34-1` | 21 |
| p (T.TEST, paired) | `H3!B40` | `=T.TEST(C5:C26,B5:B26,2,1)` | 0.00088 |
| p from t (T.DIST.2T) | `H3!B41` | `=T.DIST.2T(ABS(B38),B39)` | 0.00088 |
| Critical t | `H3!B42` | `=T.INV.2T(0.05,B39)` | 2.080 |
| 95% CI lower | `H3!B43` | `=B35-B42*B37` | 2.413 |
| 95% CI upper | `H3!B44` | `=B35+B42*B37` | 8.007 |
| Cohen's d_z | `H3!B45` | `=B35/B36` | 0.826 |
| SE of d_z | `H3!B46` | `=SQRT(1/B34+B45^2/(2*B34))` | 0.247 |
| d_z CI lower (approx.) | `H3!B47` | `=B45-1.96*B46` | 0.342 |
| d_z CI upper (approx.) | `H3!B48` | `=B45+1.96*B46` | 1.310 |
| Increases | `H3!B49` | `=COUNTIF(D5:D26,">0")` | 19 |

**Ad_B versus Baseline**

| Result | Cell | Formula | Value |
|---|---|---|---|
| Baseline mean | `H3!H29` | `=AVERAGE(H5:H25)` | 1.875 |
| Baseline SD | `H3!H30` | `=STDEV.S(H5:H25)` | 1.678 |
| Ad_B mean | `H3!H31` | `=AVERAGE(I5:I25)` | 6.000 |
| Ad_B SD | `H3!H32` | `=STDEV.S(I5:I25)` | 4.940 |
| n | `H3!H33` | `=COUNT(J5:J25)` | 21 |
| Mean difference | `H3!H34` | `=AVERAGE(J5:J25)` | 4.125 |
| SD of the differences | `H3!H35` | `=STDEV.S(J5:J25)` | 4.347 |
| SE of the mean difference | `H3!H36` | `=H35/SQRT(H33)` | 0.949 |
| t | `H3!H37` | `=H34/H36` | 4.349 |
| df | `H3!H38` | `=H33-1` | 20 |
| p (T.TEST, paired) | `H3!H39` | `=T.TEST(I5:I25,H5:H25,2,1)` | 0.00031 |
| p from t (T.DIST.2T) | `H3!H40` | `=T.DIST.2T(ABS(H37),H38)` | 0.00031 |
| Critical t | `H3!H41` | `=T.INV.2T(0.05,H38)` | 2.086 |
| 95% CI lower | `H3!H42` | `=H34-H41*H36` | 2.146 |
| 95% CI upper | `H3!H43` | `=H34+H41*H36` | 6.103 |
| Cohen's d_z | `H3!H44` | `=H34/H35` | 0.949 |
| SE of d_z | `H3!H45` | `=SQRT(1/H33+H44^2/(2*H33))` | 0.263 |
| d_z CI lower (approx.) | `H3!H46` | `=H44-1.96*H45` | 0.434 |
| d_z CI upper (approx.) | `H3!H47` | `=H44+1.96*H45` | 1.464 |
| Increases | `H3!H48` | `=COUNTIF(J5:J25,">0")` | 16 |

T.TEST returns only the p value; its arguments are (range 1, range 2, tails = 2, type = 1 for paired). The hand computed t and T.DIST.2T must give the same p, which is a good check that the two ranges are aligned row by row. R also reports the Wilcoxon signed rank test (Ad_A p < .001, Ad_B p < .001), the exact noncentral t intervals for d_z (Lakens, 2013) (Ad_A [0.33, 1.30], Ad_B [0.42, 1.46]), and Holm adjusted p values (Ad_A p < .001, Ad_B p < .001); Excel has no function for these.

---

## 8. Check your numbers (Check sheet)

Every result printed above sits on the `Check` sheet next to the value that R produced, with the difference and a tolerance of half a unit in the last printed decimal. The cell `Check!B57` counts rows with status CHECK and must read 0. The last two rows confirm the H3 rosters: every listed participant has include = 1 and valid GSR in that ad.

H1 (Fisher's exact test and the mixed logistic model), H2 (the linear mixed model), the Wilcoxon checks, and the Sensor Data consistency check are R only; see `analysis_output.txt` and `prepare_output.txt`.

---

## 9. Common mistakes, collected

1. Double clicking an export instead of importing it: wrong delimiter or decimal mark, and `#` lines inside the data.
2. Trusting a fixed row number for the header: it moves between iMotions releases. Find the `#DATA` line (or the first column name) every time.
3. Joining on the raw respondent name: `"P07 "` with its trailing space never matches. TRIM and UPPER first, and LEN to see what you cannot see.
4. Forgetting that Excel lookups ignore case while R joins do not: the same file can join in Excel and fail in R.
5. Keeping both recordings of P04, or keeping the first one: the session log decides (rule R1).
6. Using Label alone as a key: `EndCard` exists in both ads.
7. Leaving the PivotTable on Sum, or forgetting the include filter.
8. Treating `NA` as 0: an AOI never fixated has no time to first fixation (NA stays NA), but its dwell time is 0.
9. Averaging the reverse worded item as it was answered (alpha goes negative) or averaging multiple choice label values.
10. Analyzing a copy that someone opened and saved again in Excel (`SurveyStatistics_student_copy.csv`). Spreadsheet programs can also silently turn codes into dates; keep the raw export read only and work in an `.xlsx` copy (Broman & Woo, 2018; Ziemann et al., 2016).

---

## Sources

- Broman, K. W., & Woo, K. H. (2018). Data organization in spreadsheets. *The American Statistician*, *72*(1), 2-10. https://doi.org/10.1080/00031305.2017.1375989
- Ziemann, M., Eren, Y., & El-Osta, A. (2016). Gene name errors are widespread in the scientific literature. *Genome Biology*, *17*, Article 177. https://doi.org/10.1186/s13059-016-1044-7
- Lakens, D. (2013). Calculating and reporting effect sizes to facilitate cumulative science: A practical primer for t-tests and ANOVAs. *Frontiers in Psychology*, *4*, Article 863. https://doi.org/10.3389/fpsyg.2013.00863
- iMotions help center, Opening Data Export Files in Excel: https://help.imotions.com/lab/exports/additional-export-information/opening-data-export-files-in-excel
- iMotions help center, Creating Excel Pivot Tables: https://help.imotions.com/lab/exports/additional-export-information/creating-excel-pivot-tables
- iMotions help center, Export: R Processed GSR Files: https://help.imotions.com/lab/exports/r-processed-exports/export-r-processed-gsr-files
- iMotions help center, Export: Survey data: https://help.imotions.com/lab/exports/export-survey-data
- iMotions help center, Export: Data Quality Statistics: https://help.imotions.com/lab/exports/export-data-quality-statistics

## Items to verify on the lab computer

1. The header line of a real `GSR Peak Metrics` export in Lab 11.1: file name, `.csv` or `.txt`, separator, and the lines above the header (modeled here as a `#METADATA` block ending in `#DATA`).
2. The separator of `_DataQuality.txt` and of the survey files (modeled as tab), and the device names in the data quality headers.
3. The header pattern of `MERGED_SURVEY_RESPONSE_MATRIX.txt` for the lab's survey elements (modeled on the screenshot pattern `LABELVALUE_<slide>_<Question ID>,Question1 "<text>"`).
4. Which Excel version runs on the lab computers (XLOOKUP needs Excel 365 or 2021) and the Windows regional settings (expected: dot decimal, comma list separator).
