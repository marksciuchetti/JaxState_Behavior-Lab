# Remote study analysis in Excel: step by step walkthrough

**SIMULATED DATA.** Every number in this walkthrough comes from the SIMULATED files that `make_simulated_data.py` generated (seed 20276200). No real person was recorded and no webcam was switched on. The design copies the logic of the iMotions vendor demo "(DEMO) Food Waste_webET_FEA_voice" (help center article "Demo study: Food Waste (image, videos and talkaloud)"): two one minute food waste videos, a narrative one and an emotional visual one, in random order, then an infographic with four tips, then a questionnaire. Nothing from the demo's respondents was used.

This is the Excel route through the analysis in `analysis.R`, for students who do not use R. It covers the funnel, differential attrition, the exclusion rule, calibration drift, H1, and H3. The finished workbook is `SIMULATED_remote_analysis.xlsx`; the sheet names below refer to it. **Every cell address in this walkthrough is written as `Sheet!Cell` and was generated from the workbook itself.** `make_excel_workbook.py` writes the workbook and this file from one list of cells, reads the workbook back with openpyxl to confirm that every address holds the formula printed here, and recalculates a copy to confirm the results. A copy was recalculated with LibreOffice 24.2.7.2 420(Build:2) (headless): all 127 Check rows are OK, and the 134 results printed in this walkthrough agree with analysis.R at the printed decimals.

Companion files: `analysis.R` (R version), `analysis_output.txt` (its console output), `results_text.md` (model paragraphs), `figures/` (PNG figures).

---

## 0. What you are analyzing

**Study.** "SIMULATED Food Waste Remote Study", a **Remote Data Collection** (RDC) study built in iMotions Lab 11.1 with the sensors "Eye tracking (Web Camera)" and "Respondent camera", uploaded to my.imotions.com, and shared by link. Respondents took part at home on their own computers. The flow: `Pre-study calibration` (13 points, 39 s), `Instructions1` (30 s), the two videos (`foodwaste_narrative`, 60 s, and `foodwaste_visual`, 62 s), `infographic` (15 s, four tips), `Questionnaire` (iMotions survey), and `Post-study calibration` (13 points, 39 s). After collection the study zip was downloaded and loaded into iMotions Lab, the faces were scored with **Post Processing > Affectiva**, and the FEA thresholding notebook was run at a threshold of 50 (as in the demo article).

**Order.** Each order of the two videos is its own study flow, recorded in the documented `Block` column: `Study flow 1` = narrative first, `Study flow 2` = visual first. The flow was assigned at random before the link was opened (modeled: the invitation list was split at random into two halves). The demo randomized the two videos inside one block instead; that hides the order of every session whose data never uploaded, and a differential attrition check needs it.

**Webcam quality has two parts** (iMotions Online data quality box). **Accuracy** is "the mean distance between the actual and expected gaze positions calculated during calibration", in degrees of visual angle, so **smaller is better**. **Quality** is "the ratio of valid eye tracking samples in relation to the total number of samples collected"; face quality is the ratio of samples "where the face can be detected".

**Files.**

| File | What it is | Documented or modeled | Used for |
|---|---|---|---|
| `SIMULATED_session_data.csv` | One row per session that opened a link (260) | `Id`, `Started`, `Ended`, `Eye tracking`, `Est. face %`, `Block`, `Status` are the RDC respondent list columns (Accessing RDC Study Data). **Modeled:** the CSV layout and the columns `Last funnel step (modeled)`, `Device type (modeled)`, `Browser (modeled)` | Funnel, attrition |
| `SIMULATED_quality.csv` | One row per completed respondent and slide (1246 rows) | **Modeled layout** (the Quality metrics table of a study session page, with EYE TRACKING split into accuracy and quality) | Exclusion rule, drift |
| `SIMULATED_fea_metrics.csv` | FEA Metrics Export, one row per completed respondent, video, and expression (1780 rows) | **Modeled header** (the Workshop 3 layout) | H1 |
| `SIMULATED_survey.csv` | Survey export (178 rows) | **Modeled layout** | H3 |
| `SIMULATED_aoi_metrics.csv` | Individual AOI metrics on the infographic | Metric names from the AOI Metrics article; identifier columns modeled | H2 (R only) |

**Hypotheses.**
- H1 (within subjects): percent of time Engaged is higher for `foodwaste_visual` than for `foodwaste_narrative`.
- H2 (within subjects): dwell time differs among the four tip AOIs on the infographic (R only; section 10).
- H3 (between subjects; randomization check): intention to reduce food waste does not differ by which video came first.
- Data quality: the funnel, completion by order (differential attrition), calibration drift, and accuracy against the hit on the smallest AOI (R only).

**Decision rules fixed before collection.** Exclude a completed respondent when the Pre-study **or** the Post-study calibration accuracy is above 4.0 degrees (iMotions Online Study Setup: "Generally acceptable range: 0-4 dva"), when eye tracking quality on the infographic is below 90 percent ("We recommend a quality score of at least 90%"), or when face quality is below 80 percent during either video ("at least 80%"). Sensitivity: accuracy at or below 3.0 degrees at both calibrations (the Data collection and Analysis articles: "We recommend a threshold of 3.0"). All tests are two tailed with alpha = .05.

---

## 1. Import the files

iMotions exports are comma separated with a dot decimal. Double clicking a file uses your regional defaults, which can put everything in one column (Opening Data Export Files in Excel article). Use the import route instead.

**Route A, Get Data (Excel 2016 and later):**
1. **Data > From Text/CSV**, select `SIMULATED_session_data.csv`, click **Import**.
2. In the preview set **Delimiter: Comma** and **File Origin: 65001 Unicode (UTF-8)**.
3. Click **Transform Data**. Every simulated file starts with **one** comment line (`# SIMULATED ...`). In Power Query choose **Home > Remove Rows > Remove Top Rows** (1), then **Home > Use First Row as Headers**. Select the `Eye tracking` column and set **Data Type: Text** (Power Query may otherwise try to read "1.9, 96%" as something else), then **Close & Load**.
4. **You should see** 10 columns with `Id` in the first column and 260 data rows, one per session. Repeat for the other files (1246, 1780, and 178 data rows).

**Route B, Text Import Wizard (any version):** **File > Open > Browse**, set the file type to **All Files**, open the .csv, choose **Delimited**, start at row 2, tick **Comma**, set the `Eye tracking` column to **Text** in the last step, click **Advanced** and set the decimal symbol to a dot, then **Finish**.

**Common mistakes.** (1) The `Eye tracking` cell holds **two numbers in one text cell** ("1.9, 96%": accuracy in degrees, then eye tracking quality). It cannot be averaged until it is split (section 2). (2) Empty cells are real: a session that never finished has no `Ended` time and no quality values. Leave them empty; do not type 0. (3) The quality table writes `NA` in the accuracy column of every slide that is not a calibration; leave it as text.

In the workbook the files sit on the **Session**, **Quality**, **FEA**, and **Survey** sheets, header in row 1 and data from row 2. Session column letters: `A` = Id, `D` = Eye tracking, `F` = Block, `G` = Status, `H` = Last funnel step (modeled), `I` = Device type (modeled).

---

## 2. Helper columns on the Session sheet

Add five columns to the right of the session data (`K` to `O`, headers in orange) and fill them down. The formulas below are the ones in row 2.

| Header | Cell | Formula | Meaning |
|---|---|---|---|
| Step number | `Session!K2` | `=IF(H2="Completed",11,VALUE(MID(H2,6,FIND(" ",H2,6)-6)))` | The number inside "STEP 5 Screen size" (characters 6 up to the next space); 11 for a session that completed |
| Completed (1 = yes) | `Session!L2` | `=IF(G2="Processed",1,0)` | 1 when Status is Processed (the data arrived and were processed) |
| Order | `Session!M2` | `=IF(F2="Study flow 1","Narrative first","Visual first")` | The flow translated into the order of the videos |
| Session accuracy (degrees) | `Session!N2` | `=IF(D2="","",VALUE(LEFT(D2,FIND(",",D2)-1)))` | The text before the comma, as a number |
| Session quality (%) | `Session!O2` | `=IF(D2="","",VALUE(SUBSTITUTE(MID(D2,FIND(",",D2)+2,10),"%","")))` | The text after ", " without the % sign |

The first completed session, `RESP0001`, is in row 2: its `Eye tracking` cell reads "1.3, 95%" and `Session!N2` returns 1.3. A session that stopped early has an empty `Eye tracking` cell, so its last two helper cells stay empty (the first one is `RESP0008` in row 9).

**Why step numbers instead of step names.** The ninth step is called "STEP 9 Stimuli presentation*" (with an asterisk, as the portal writes it). In `COUNTIF` and `COUNTIFS` an asterisk in the criterion is a **wildcard** that matches any text, so a criterion typed from step names can count the wrong rows. Numbers avoid the problem.

**Common mistake.** Screening respondents on the session level `Eye tracking` number. It is rounded to one decimal and (in this simulation) is the mean of the two calibrations, so it hides a bad calibration: 6 completed respondents show 4.0 or less there although one calibration was above 4.0 (step 4 of `analysis_output.txt`). The rule uses the quality table instead (section 5).

---

## 3. The respondent funnel (Funnel sheet)

The portal draws the respondent funnel as a chart ("View information on where respondents dropped off. Sessions that are in progress are not included."). The **Funnel** sheet rebuilds it as a table. A session **entered** a step when its step number is that step or a later one, and it **stopped** at the step where it ended.

| Column | Cell (STEP 1) | Formula |
|---|---|---|
| Entered | `Funnel!C5` | `=COUNTIF(Session!$K$2:$K$261,">="&A5)` |
| Stopped here | `Funnel!D5` | `=COUNTIF(Session!$K$2:$K$261,A5)` |
| Continued | `Funnel!E5` | `=C5-D5` |
| Percent of opened | `Funnel!F5` | `=100*C5/COUNTA(Session!$A$2:$A$261)` |
| Percent of previous step (STEP 2) | `Funnel!G6` | `=100*C6/C5` |

**You should see** (simulated):

| Step | Entered | Stopped here | Percent of opened |
|---|---|---|---|
| STEP 1 Testing upload | 260 | 2 | 100.0 |
| STEP 2 Consent screen | 258 | 14 | 99.2 |
| STEP 3 Screen recording setup | 244 | 0 | 93.8 |
| STEP 4 Respondent camera setup | 244 | 12 | 93.8 |
| STEP 5 Screen size | 232 | 36 | 89.2 |
| STEP 6 Enter fullscreen | 196 | 3 | 75.4 |
| STEP 7 Prepare positioning | 193 | 2 | 74.2 |
| STEP 8 Positioning check | 191 | 6 | 73.5 |
| STEP 9 Stimuli presentation* | 185 | 4 | 71.2 |
| STEP 10 Study data upload | 181 | 3 | 69.6 |
| **Completed (Status Processed)** | **178** | | **68.5** |

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Completed (Status Processed) | `Funnel!C15` | `=COUNTIF(Session!$G$2:$G$261,"Processed")` | 178 |
| Continued after STEP 10 minus Completed (must be 0) | `Funnel!C16` | `=E14-C15` | 0 |
| Status Abandoned | `Funnel!C17` | `=COUNTIF(Session!$G$2:$G$261,"Abandoned")` | 81 |
| Status Errors | `Funnel!C18` | `=COUNTIF(Session!$G$2:$G$261,"Errors")` | 1 |

**Reading the funnel.** Of 260 sessions that opened a link, 178 (68.5 percent) completed. The largest single loss is at **STEP 5 Screen size** (36 sessions: 34 phones or tablets and 2 computers). Every phone or tablet session that reached STEP 5 stopped there (2 of the 36 phone and tablet sessions had left earlier). This is modeled: the RDC article says "Participation can not be done on phones or tablets" but not at which step a phone is stopped. Nobody stopped at STEP 3 because screen recording was off in this study. STEP 9 holds both calibrations, both videos, the infographic, and the questionnaire; 4 sessions stopped there. At STEP 10 the upload failed for 3 sessions (1 with Status Errors).

**By device** (rows 53 to 56, modeled column): `Funnel!B53` = `=COUNTIF(Session!$I$2:$I$261,A53)`, `Funnel!C53` = `=COUNTIFS(Session!$I$2:$I$261,A53,Session!$L$2:$L$261,1)`. Laptop 124 of 158 completed, Desktop 54 of 66, Phone 0 of 29, Tablet 0 of 7. Put "desktop or laptop computer only" in the invitation.

**Chart.** The sheet has a bar chart of the Entered column (**Insert > Bar > Clustered Bar** on `Funnel!B4:C15`; right click the vertical axis, **Format Axis > Categories in reverse order**, so STEP 1 is on top).

**Common mistake.** Reporting only the final number. A funnel shows **where** people were lost, and each kind of loss has a different fix: consent wording (STEP 2), camera help (STEP 4), "computer only" in the invitation (STEP 5), a shorter study (STEP 9), the data upload alternative (STEP 10).

---

## 4. Differential attrition: did completion depend on the order?

If one order made more people quit (for example, the visual video upset some viewers when it came first), the people who finished each order would no longer be comparable, and an order difference could appear that randomization did not create (Zhou & Fishbach, 2016). Because the flow was assigned before the link was opened, the `Block` column gives the order of **every** session.

| Column | Cell (Narrative first) | Formula |
|---|---|---|
| Opened | `Funnel!B22` | `=COUNTIF(Session!$M$2:$M$261,A22)` |
| Reached STEP 9 | `Funnel!C22` | `=COUNTIFS(Session!$M$2:$M$261,A22,Session!$K$2:$K$261,">=9")` |
| Completed | `Funnel!D22` | `=COUNTIFS(Session!$M$2:$M$261,A22,Session!$L$2:$L$261,1)` |
| Percent completed | `Funnel!F22` | `=100*D22/B22` |
| Expected completed (if order and completion were unrelated) | `Funnel!H22` | `=B22*$D$24/$B$24` |

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Opened, Narrative first | `Funnel!B22` | `=COUNTIF(Session!$M$2:$M$261,A22)` | 126 |
| Opened, Visual first | `Funnel!B23` | `=COUNTIF(Session!$M$2:$M$261,A23)` | 134 |
| Completed, Narrative first | `Funnel!D22` | `=COUNTIFS(Session!$M$2:$M$261,A22,Session!$L$2:$L$261,1)` | 87 |
| Completed, Visual first | `Funnel!D23` | `=COUNTIFS(Session!$M$2:$M$261,A23,Session!$L$2:$L$261,1)` | 91 |
| Percent completed, Narrative first | `Funnel!F22` | `=100*D22/B22` | 69.0 |
| Percent completed, Visual first | `Funnel!F23` | `=100*D23/B23` | 67.9 |
| Chi square (Pearson, no continuity correction) | `Funnel!L23` | `=SUMPRODUCT((D22:E23-H22:I23)^2/H22:I23)` | 0.04 |
| p (CHISQ.DIST.RT) | `Funnel!L25` | `=CHISQ.DIST.RT(L23,L24)` | 0.844 |
| p (CHISQ.TEST) | `Funnel!L26` | `=CHISQ.TEST(D22:E23,H22:I23)` | 0.844 |
| phi | `Funnel!L27` | `=SQRT(L23/B24)` | 0.01 |
| Smallest expected count | `Funnel!L29` | `=MIN(H22:I23)` | 39.74 |

`CHISQ.TEST(observed, expected)` needs both tables, so the expected counts are built first (row total x column total / N). R's `chisq.test` applies the Yates continuity correction to a 2 x 2 table by default; `analysis.R` reports the uncorrected value as the main result because it matches Excel, and prints the corrected one beside it (X2 = 0.004).

**The sharper check.** Before STEP 9 nobody had seen a video, so only losses from STEP 9 on can depend on the order. Rows 37 to 38 repeat the test for the sessions that reached STEP 9:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Percent completed after STEP 9, Narrative first | `Funnel!E37` | `=100*C37/B37` | 95.6 |
| Percent completed after STEP 9, Visual first | `Funnel!E38` | `=100*C38/B38` | 96.8 |
| Chi square (Pearson) | `Funnel!L38` | `=SUMPRODUCT((C37:D38-H37:I38)^2/H37:I38)` | 0.18 |
| p (CHISQ.DIST.RT) | `Funnel!L39` | `=CHISQ.DIST.RT(L38,1)` | 0.668 |
| Smallest expected count | `Funnel!L40` | `=MIN(H37:I38)` | 3.44 |

The smallest expected count is below 5, so report the **Fisher exact test**, which Excel does not have: in R, p = 0.718, odds ratio = 0.72, 95% CI [0.10, 4.38].

**You should see** completion rates within a few points of each other and p values well above .05 in both tests: no sign of differential attrition (simulated). A nonsignificant test is not proof of equal attrition (the test has little power for small losses), so report the counts by order as well.

---

## 5. The exclusion rule (Screening sheet)

The **Screening** sheet lists every completed respondent once (178 respondents in rows 5 to 182). The thresholds sit in the yellow cells `Screening!B2` (4.0 degrees), `Screening!D2` (3.0, sensitivity), `Screening!F2` (90 percent), and `Screening!H2` (80 percent); change one and every decision updates. The five quality values come from the **Quality** sheet with `SUMIFS` on respondent and slide (each pair occurs once, so the sum is the value).

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Order | `Screening!B5` | `=INDEX(Session!$M$2:$M$261,MATCH($A5,Session!$A$2:$A$261,0))` |
| Pre-study accuracy | `Screening!C5` | `=SUMIFS(Quality!$E$2:$E$1247,Quality!$A$2:$A$1247,$A5,Quality!$D$2:$D$1247,"Pre-study calibration")` |
| Post-study accuracy | `Screening!D5` | `=SUMIFS(Quality!$E$2:$E$1247,Quality!$A$2:$A$1247,$A5,Quality!$D$2:$D$1247,"Post-study calibration")` |
| Eye tracking quality, infographic | `Screening!E5` | `=SUMIFS(Quality!$F$2:$F$1247,Quality!$A$2:$A$1247,$A5,Quality!$D$2:$D$1247,"infographic")` |
| Face quality, narrative | `Screening!F5` | `=SUMIFS(Quality!$G$2:$G$1247,Quality!$A$2:$A$1247,$A5,Quality!$D$2:$D$1247,"foodwaste_narrative")` |
| Face quality, visual | `Screening!G5` | `=SUMIFS(Quality!$G$2:$G$1247,Quality!$A$2:$A$1247,$A5,Quality!$D$2:$D$1247,"foodwaste_visual")` |
| Fail: Pre-study accuracy | `Screening!H5` | `=IF(C5>$B$2,1,0)` |
| Fail: Post-study accuracy | `Screening!I5` | `=IF(D5>$B$2,1,0)` |
| Fail: eye tracking quality | `Screening!J5` | `=IF(E5<$F$2,1,0)` |
| Fail: face quality | `Screening!K5` | `=IF(OR(F5<$H$2,G5<$H$2),1,0)` |
| Decision | `Screening!L5` | `=IF(SUM(H5:K5)>0,"Exclude","Keep")` |
| First reason | `Screening!M5` | `=IF(H5=1,"Pre-study accuracy",IF(I5=1,"Post-study accuracy",IF(J5=1,"Eye tracking quality",IF(K5=1,"Face quality",""))))` |
| Largest accuracy | `Screening!N5` | `=MAX(C5:D5)` |
| Decision at 3.0 | `Screening!O5` | `=IF(AND(L5="Keep",N5<=$D$2),"Keep","Exclude")` |
| Drift (Post minus Pre) | `Screening!P5` | `=D5-C5` |

The counts (column S):

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Completed respondents screened | `Screening!S7` | `=COUNTA(A5:A182)` | 178 |
| Excluded | `Screening!S8` | `=COUNTIF($L$5:$L$182,"Exclude")` | 36 |
| Kept | `Screening!S9` | `=COUNTIF($L$5:$L$182,"Keep")` | 142 |
| First reason: Pre-study accuracy | `Screening!S10` | `=COUNTIF($M$5:$M$182,"Pre-study accuracy")` | 5 |
| First reason: Post-study accuracy | `Screening!S11` | `=COUNTIF($M$5:$M$182,"Post-study accuracy")` | 12 |
| First reason: eye tracking quality | `Screening!S12` | `=COUNTIF($M$5:$M$182,"Eye tracking quality")` | 14 |
| First reason: face quality | `Screening!S13` | `=COUNTIF($M$5:$M$182,"Face quality")` | 5 |
| Failed (any): Pre-study accuracy | `Screening!S14` | `=SUM(H5:H182)` | 5 |
| Failed (any): Post-study accuracy | `Screening!S15` | `=SUM(I5:I182)` | 17 |
| Failed (any): eye tracking quality | `Screening!S16` | `=SUM(J5:J182)` | 20 |
| Failed (any): face quality | `Screening!S17` | `=SUM(K5:K182)` | 14 |
| Failed more than one criterion | `Screening!S18` | `=SUMPRODUCT(--((H5:H182+I5:I182+J5:J182+K5:K182)>1))` | 14 |
| Failed only at the Post-study calibration | `Screening!S19` | `=COUNTIFS(H5:H182,0,I5:I182,1)` | 12 |
| Kept, Narrative first | `Screening!S20` | `=COUNTIFS($B$5:$B$182,"Narrative first",$L$5:$L$182,"Keep")` | 69 |
| Kept, Visual first | `Screening!S21` | `=COUNTIFS($B$5:$B$182,"Visual first",$L$5:$L$182,"Keep")` | 73 |
| Kept at 3.0 degrees | `Screening!S22` | `=COUNTIF($O$5:$O$182,"Keep")` | 125 |
| Further removed at 3.0 degrees | `Screening!S23` | `=S9-S22` | 17 |

**Two ways to count reasons.** "Failed (any)" counts every respondent who failed a criterion, so a respondent can appear twice and the column adds up to more than the number excluded. "First reason" gives each excluded respondent one reason, the first one failed in the order of the rule, so it adds up to the total. Report both, as the lab convention asks ("report respondents kept per stimulus and the number excluded for each reason").

**You should see** 36 exclusions and 142 kept respondents (69 Narrative first, 73 Visual first). Of the excluded respondents, 12 calibrated well at the start and failed only at the end: a single calibration at the start would have kept them. The RDC Testing with Webcam Eye Tracking article calls a calibration at the start and another at the end "currently the best practice that we recommend for webcam eye tracking".

**Common mistake.** Reading accuracy like a score where bigger is better. Accuracy is an **error** in degrees: the rule excludes values **above** 4.0, while the two quality rules exclude values **below** their minimum.

---

## 6. Calibration drift (Screening sheet, column S)

The same calibration ran at the start and at the end, so a paired comparison shows how much the gaze estimate worsened while the study ran (people move, lean back, or the light changes).

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Pre-study mean (degrees) | `Screening!S28` | `=AVERAGE(C5:C182)` | 1.71 |
| Post-study mean (degrees) | `Screening!S30` | `=AVERAGE(D5:D182)` | 2.42 |
| n | `Screening!S32` | `=COUNT(P5:P182)` | 178 |
| Mean difference, Post minus Pre | `Screening!S33` | `=AVERAGE(P5:P182)` | 0.72 |
| SD of the differences | `Screening!S34` | `=STDEV.S(P5:P182)` | 0.50 |
| t | `Screening!S36` | `=S33/S35` | 19.05 |
| df | `Screening!S37` | `=S32-1` | 177 |
| p, two tailed (T.TEST type 1) | `Screening!S38` | `=T.TEST(C5:C182,D5:D182,2,1)` | 9.96e-45 |
| Mean difference 95% CI lower | `Screening!S41` | `=S33-S40*S35` | 0.64 |
| Mean difference 95% CI upper | `Screening!S42` | `=S33+S40*S35` | 0.79 |
| Cohen's d_z | `Screening!S43` | `=S33/S34` | 1.43 |
| Correlation of the two calibrations | `Screening!S47` | `=CORREL(C5:C182,D5:D182)` | 0.91 |
| Respondents with a worse Post-study accuracy | `Screening!S50` | `=COUNTIF(P5:P182,">0")` | 165 |

The p value is far below .001; write "p < .001", never "p = .000". `analysis.R` also gives the exact interval for d_z, [1.22, 1.64], and the Wilcoxon signed rank test (V = 15415, p < .001).

---

## 7. H1: Engagement, visual versus narrative video (H1 sheet)

The **H1** sheet has one row per kept respondent (rows 5 to 146). The respondents kept under the 3.0 degree sensitivity rule come first (rows 5 to 129) and the 17 removed by it last (rows 130 to 146), so the sensitivity block only shortens the range.

| Column | Cell (first respondent) | Formula |
|---|---|---|
| Order | `H1!B5` | `=INDEX(Session!$M$2:$M$261,MATCH($A5,Session!$A$2:$A$261,0))` |
| Decision at 3.0 | `H1!C5` | `=INDEX(Screening!$O$5:$O$182,MATCH($A5,Screening!$A$5:$A$182,0))` |
| Engagement, narrative (%) | `H1!D5` | `=SUMIFS(FEA!$H$2:$H$1781,FEA!$B$2:$B$1781,$A5,FEA!$D$2:$D$1781,"foodwaste_narrative",FEA!$F$2:$F$1781,"Engagement")` |
| Engagement, visual (%) | `H1!E5` | `=SUMIFS(FEA!$H$2:$H$1781,FEA!$B$2:$B$1781,$A5,FEA!$D$2:$D$1781,"foodwaste_visual",FEA!$F$2:$F$1781,"Engagement")` |
| Difference (visual minus narrative) | `H1!F5` | `=E5-D5` |

`SUMIFS` with three conditions (respondent, `Interval label`, `Expression`) picks exactly one row of the FEA export. The test sits in column I (formula text in column J). A paired t test **is** a one sample t test on the differences:

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Narrative mean (% of time) | `H1!I7` | `=AVERAGE(D5:D146)` | 19.49 |
| Narrative SD | `H1!I8` | `=STDEV.S(D5:D146)` | 19.97 |
| Visual mean (% of time) | `H1!I9` | `=AVERAGE(E5:E146)` | 43.63 |
| Visual SD | `H1!I10` | `=STDEV.S(E5:E146)` | 26.39 |
| n | `H1!I11` | `=COUNT(F5:F146)` | 142 |
| Mean difference, visual minus narrative | `H1!I12` | `=AVERAGE(F5:F146)` | 24.14 |
| SD of the differences | `H1!I13` | `=STDEV.S(F5:F146)` | 28.83 |
| t | `H1!I15` | `=I12/I14` | 9.98 |
| df | `H1!I16` | `=I11-1` | 141 |
| **p, two tailed (T.TEST type 1)** | `H1!I17` | `=T.TEST(D5:D146,E5:E146,2,1)` | 4.57e-18 |
| p from the hand computed t | `H1!I18` | `=T.DIST.2T(ABS(I15),I16)` | 4.57e-18 |
| Mean difference 95% CI lower | `H1!I20` | `=I12-I19*I14` | 19.36 |
| Mean difference 95% CI upper | `H1!I21` | `=I12+I19*I14` | 28.92 |
| Cohen's d_z | `H1!I22` | `=I12/I13` | 0.84 |
| d_z 95% CI lower (approximation) | `H1!I24` | `=I22-1.96*I23` | 0.65 |
| d_z 95% CI upper (approximation) | `H1!I25` | `=I22+1.96*I23` | 1.03 |
| Respondents with a positive difference | `H1!I27` | `=COUNTIF(F5:F146,">0")` | 113 |
| Respondents at exactly 0, narrative | `H1!I29` | `=COUNTIF(D5:D146,0)` | 21 |
| Respondents at exactly 0, visual | `H1!I30` | `=COUNTIF(E5:E146,0)` | 4 |
| Median, narrative | `H1!I31` | `=MEDIAN(D5:D146)` | 13.22 |
| Median, visual | `H1!I32` | `=MEDIAN(E5:E146)` | 42.96 |
| Median difference | `H1!I33` | `=MEDIAN(F5:F146)` | 23.82 |

`T.TEST(array1, array2, tails, type)`: `tails = 2` for a two tailed test; `type = 1` for **paired** data. `type = 2` and `type = 3` treat the two columns as different people and are wrong here. **You should see** the hand computed p equal the `T.TEST` p (both far below .001, so report p < .001). For paired data the effect size is **d_z**, the mean difference in units of the SD of the differences; SE(d_z) = sqrt(1/n + d_z^2 / (2 n)) is the large sample approximation, and `analysis.R` also reports the exact interval from the noncentral t distribution, [0.64, 1.03] (Lakens, 2013). Rule of thumb for d_z: 0.2 small, 0.5 medium, 0.8 large.

**Why the Wilcoxon check matters here.** Percent of time data pile up at zero and have a long right tail: 21 respondents were never above the threshold during the narrative video, and its mean (19.49) is above its median (13.22). With n = 142 the t test on the mean difference is robust to this (the differences themselves are close to symmetric), and the Wilcoxon signed rank test, which compares ranks, is the check. It has no Excel function; in R: V = 8852, p < .001.

**Common mistake.** Reading "percent of time Engaged" as "percent of viewers engaged". It is, for each respondent, the share of the video's frames in which the Engagement score was above 50; the mean is the average of those shares.

---

## 8. Sensitivity: H1 at the 3.0 degree threshold (H1 sheet)

Accuracy does not affect facial expression scores, but the rule removes respondents, so the sensitivity analysis asks whether H1 depends on who was removed. The second block uses rows 5 to 129.

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| n | `H1!I42` | `=COUNT(F5:F129)` | 125 |
| Narrative mean (% of time) | `H1!I38` | `=AVERAGE(D5:D129)` | 18.08 |
| Visual mean (% of time) | `H1!I40` | `=AVERAGE(E5:E129)` | 42.43 |
| Mean difference, visual minus narrative | `H1!I43` | `=AVERAGE(F5:F129)` | 24.36 |
| t | `H1!I46` | `=I43/I45` | 9.87 |
| p (T.TEST type 1) | `H1!I48` | `=T.TEST(D5:D129,E5:E129,2,1)` | 2.62e-17 |
| Mean difference 95% CI lower | `H1!I51` | `=I43-I50*I45` | 19.47 |
| Mean difference 95% CI upper | `H1!I52` | `=I43+I50*I45` | 29.24 |
| Cohen's d_z | `H1!I53` | `=I43/I44` | 0.88 |

| Roster check | Cell | Formula | Result |
|---|---|---|---|
| Kept respondents on the Screening sheet | `H1!I64` | `=COUNTIF(Screening!$L$5:$L$182,"Keep")` | 142 |
| Rows in the sensitivity block marked Exclude at 3.0 (must be 0) | `H1!I65` | `=COUNTIF(C5:C129,"Exclude")` | 0 |
| Rows below the sensitivity block marked Exclude at 3.0 | `H1!I66` | `=COUNTIF(C130:C146,"Exclude")` | 17 |

**You should see** the same conclusion with 17 fewer respondents: d_z 0.84 with all kept respondents and 0.88 at 3.0 degrees.

---

## 9. H3: intention by order, the randomization check (H3 sheet)

The **H3** sheet lists the kept respondents by order: Narrative first in rows 5 to 73, Visual first in rows 74 to 146. `H3!C5` = `=INDEX(Survey!$E$2:$E$179,MATCH($A5,Survey!$B$2:$B$179,0))` looks up the answer.

| Statistic | Cell | Formula | Result (simulated) |
|---|---|---|---|
| Narrative first mean | `H3!F7` | `=AVERAGE(C5:C73)` | 5.04 |
| Narrative first SD | `H3!F8` | `=STDEV.S(C5:C73)` | 1.37 |
| Narrative first n | `H3!F9` | `=COUNT(C5:C73)` | 69 |
| Visual first mean | `H3!F10` | `=AVERAGE(C74:C146)` | 5.34 |
| Visual first SD | `H3!F11` | `=STDEV.S(C74:C146)` | 1.22 |
| Visual first n | `H3!F12` | `=COUNT(C74:C146)` | 73 |
| Difference, Narrative first minus Visual first | `H3!F13` | `=F7-F10` | -0.30 |
| Welch t | `H3!F15` | `=F13/F14` | -1.37 |
| Welch df | `H3!F16` | `=(F8^2/F9+F11^2/F12)^2/((F8^2/F9)^2/(F9-1)+(F11^2/F12)^2/(F12-1))` | 135.97 |
| **p, two tailed (T.TEST type 3)** | `H3!F17` | `=T.TEST(C5:C73,C74:C146,2,3)` | 0.171 |
| Difference 95% CI lower | `H3!F20` | `=F13-F19*F14` | -0.73 |
| Difference 95% CI upper | `H3!F21` | `=F13+F19*F14` | 0.13 |
| Pooled SD | `H3!F22` | `=SQRT(((F9-1)*F8^2+(F12-1)*F11^2)/(F9+F12-2))` | 1.29 |
| Cohen's d | `H3!F23` | `=F13/F22` | -0.23 |
| Hedges' g | `H3!F25` | `=F24*F23` | -0.23 |
| g 95% CI lower (approximation) | `H3!F27` | `=F25-1.96*F26` | -0.56 |
| g 95% CI upper (approximation) | `H3!F28` | `=F25+1.96*F26` | 0.10 |

`type = 3` in `T.TEST` is Welch's test, which does not assume equal variances (Delacre et al., 2017). Excel's `T.DIST.2T` and `T.INV.2T` may round the Welch df down to a whole number, so the hand computed p and the interval can differ from R in the third decimal; report the `T.TEST` p. The exact interval for g from R is [-0.56, 0.10]; the Mann Whitney check gives W = 2243, p = .249.

**What this test means.** The two groups differ only in which video came first, and the order was assigned at random. A test of something measured after both videos asks two things at once: did randomization give comparable groups, and did the order carry over to the questionnaire? A **nonsignificant** result is what randomization predicts. A significant one would mean either a chance imbalance (expected in 5 percent of such tests) or a real order effect; the analysis would then report intention separately by order or add order to the model, and the attrition check in section 4 would matter even more. **You should see** a small difference with a confidence interval that includes zero (simulated).

**Common mistake.** Calling a nonsignificant randomization check "proof" that the groups are equal. It only means the data give no reason to think they differ; the confidence interval shows how large a difference is still compatible with the data.

---

## 10. What Excel cannot do here: the R results

**H2: dwell time on the four tips** (`SIMULATED_aoi_metrics.csv`, `Dwell time (ms)`, with NA counted as 0 ms because the respondent never fixated the AOI). Within each respondent the four tips are ranked 1 to 4 and Friedman's test asks whether the mean ranks differ; Kendall's W = Friedman chi square / (n (k - 1)) runs from 0 (no agreement) to 1 (every respondent ranks the tips the same way). In R: Friedman chi square(3) = 78.95, p < .001, n = 142, W = 0.185. Holm adjusted Wilcoxon signed rank comparisons (mean difference in ms, first AOI minus second):

| Pair | Mean difference (ms) | d_z | V | p | p (Holm) |
|---|---|---|---|---|---|
| Tip_1 vs Tip_2 | 1643 | 0.66 | 8420.5 | < .001 | < .001 |
| Tip_1 vs Tip_3 | 1710 | 0.73 | 8602 | < .001 | < .001 |
| Tip_1 vs Tip_4 | 1739 | 0.71 | 8599 | < .001 | < .001 |
| Tip_2 vs Tip_3 | 68 | 0.03 | 5169.5 | .851 | 1.000 |
| Tip_2 vs Tip_4 | 96 | 0.05 | 5246 | .731 | 1.000 |
| Tip_3 vs Tip_4 | 29 | 0.01 | 5379.5 | .538 | 1.000 |

Three of the six pairs differ after the adjustment (Tip_1 vs Tip_2, Tip_1 vs Tip_3, and Tip_1 vs Tip_4); the other pairs (Tip_2 vs Tip_3, Tip_2 vs Tip_4, and Tip_3 vs Tip_4) do not. At 3.0 degrees: chi square(3) = 73.18, W = 0.195, 3 pairs. A PivotTable of the AOI file (**Insert > PivotTable**, `AOI Name` in **Rows**, `Dwell time (ms)` in **Values** set to **Average**) gives descriptive means, but it skips the NA rows, so its means are for the respondents who fixated each tip, not for everyone (Creating Excel Pivot Tables article).

**Accuracy and the hit on the smallest AOI** (the Headline banner). Logistic regression of hit (at least one fixation) on the mean of the two calibration accuracies, for every completed respondent whose eye tracking and face quality passed (n = 151): odds ratio per degree = 0.61, 95% CI [0.40, 0.89], p = .013. Hit ratio by mean accuracy: 90.0 percent at 1.5 degrees or less (n = 60); 67.7 percent at 1.5 to 2.5 degrees (n = 62); 63.6 percent at 2.5 to 3.5 degrees (n = 22); 71.4 percent above 3.5 degrees (n = 7); the last band holds too few respondents for a stable percentage.

**Exploratory.** Engagement difference (visual minus narrative) with naming the visual video more persuasive: point biserial r = .25, 95% CI [.09, .40], p = .003 (in Excel, `CORREL` of the difference column and a 0/1 column gives the same r). Total tip dwell time with tips recalled: r = .16, 95% CI [-.01, .32], p = .058; Spearman rho = .16.

Also R only: the Wilcoxon, Mann Whitney, and Fisher exact tests, Shapiro Wilk tests, and the exact noncentral t intervals.

---

## 11. Check your numbers (Check sheet)

The **Check** sheet puts each live Excel result next to the value `analysis.R` produced and prints OK or CHECK (status in column F, the count of CHECK rows in `Check!B132`; it should be 0). If you rebuild the analysis yourself and a row says CHECK, the usual causes are a wrong range (an excluded respondent left in a roster), the wrong `T.TEST` type, or a lookup on the wrong video or expression.

If you have the **Analysis ToolPak** add in, **Data > Data Analysis > t-Test: Paired Two Sample for Means** (H1, drift) and **t-Test: Two-Sample Assuming Unequal Variances** (H3) are good cross checks.

---

## 12. Common mistakes, collected

1. Opening the .csv by double clicking, or letting the `Eye tracking` text become a number or a date.
2. Using the step names with an asterisk in `COUNTIF` (the asterisk is a wildcard); use the step numbers.
3. Screening on the rounded session level accuracy instead of both calibrations (section 2).
4. Treating accuracy like a score: bigger accuracy numbers are worse.
5. Counting reasons only once or only "any": report both columns (section 5).
6. Testing attrition only among completed respondents; the order of every session is in `Block`, so use all sessions (section 4).
7. Using `T.TEST` type 2 or 3 for H1 or drift (paired data), or type 1 for H3 (two groups).
8. Reading a nonsignificant randomization check as proof of equal groups (section 9).
9. Treating the 1780 FEA rows as independent observations: each respondent contributes ten rows, and the tests use one value per respondent and video.
10. Forgetting that every number here is simulated.

## Sources

- iMotions help center, Accessing RDC Study Data: https://help.imotions.com/lab/remote-data-collection/accessing-rdc-study-data
- iMotions help center, Data Collection - Behavioral Research Online (respondent funnel and data quality box): https://help.imotions.com/online/behavioral-research/data-collection-behavioral-research-online
- iMotions help center, Analysis - Behavioral Research Online: https://help.imotions.com/online/behavioral-research/analysis-behavioral-research-online
- iMotions help center, RDC Testing with Webcam Eye Tracking: https://help.imotions.com/lab/remote-data-collection/getting-started-with-online-data-collection/rdc-testing-with-webcam-eye-tracking
- iMotions help center, Participating in RDC Studies: https://help.imotions.com/lab/remote-data-collection/participating-in-rdc-studies
- iMotions help center, Demo study: Food Waste (image, videos and talkaloud): https://help.imotions.com/lab/getting-started/demo-studies/10-0-demo-studies/demo-study-food-waste-image-videos-and-talkaloud
- iMotions help center, R Notebook (FEA): Facial expression analysis thresholding and aggregation: https://help.imotions.com/lab/analysis/facial-expression-analysis/r-notebook-fea-facial-expression-analysis-thresholding-and-aggregation
- iMotions help center, AOI Metrics: https://help.imotions.com/lab/analysis/eye-tracking/areas-of-interest/aoi-metrics
- iMotions help center, Opening Data Export Files in Excel: https://help.imotions.com/lab/exports/additional-export-information/opening-data-export-files-in-excel
- iMotions help center, Creating Excel Pivot Tables: https://help.imotions.com/lab/exports/additional-export-information/creating-excel-pivot-tables
- iMotions Online Study Setup (local manual PDF, "Generally acceptable range: 0-4 dva")
- Delacre, M., Lakens, D., & Leys, C. (2017). Why psychologists should by default use Welch's t-test instead of Student's t-test. *International Review of Social Psychology*, *30*(1), 92-101. https://doi.org/10.5334/irsp.82
- Lakens, D. (2013). Calculating and reporting effect sizes to facilitate cumulative science: A practical primer for t-tests and ANOVAs. *Frontiers in Psychology*, *4*, Article 863. https://doi.org/10.3389/fpsyg.2013.00863
- Zhou, H., & Fishbach, A. (2016). The pitfall of experimenting on the web: How unattended selective attrition leads to surprising (yet false) research conclusions. *Journal of Personality and Social Psychology*, *111*(4), 493-504. https://doi.org/10.1037/pspa0000056

## Items to verify on the lab computer

- The **session data download**: the CSV layout of "Download session data" for an RDC study, its column names (the manuals show the on screen list), the Id format, and which calibration the `Eye tracking` accuracy summarizes. Whether it holds any per session funnel step, device, or browser column; the three columns marked (modeled) are not documented.
- That an RDC study built in Lab 11.1 can hold **two flows** with one link each, and that the `Block` column names the flow.
- The **Status** of a session that stopped before the end (modeled here as Abandoned) and the step at which a **phone** is stopped (modeled as STEP 5); test with a phone during the pilot.
- Where the **per slide quality values** live for an RDC study (portal session page, Lab Exposure summary, or the data quality export); `SIMULATED_quality.csv` is a modeled table.
- The **FEA Metrics Export** header in Lab 11.1, whether the thresholding notebook offers Engagement (for example under "Other Measures thresholding"), and the denominator of "Percent of time present".
- The **survey export** layout for an RDC study ("the survey data export option").
