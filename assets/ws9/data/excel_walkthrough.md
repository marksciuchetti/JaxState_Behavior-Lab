# Excel walkthrough: sample size, attrition, the Williams design, and the session calendar

**PLANNED STUDY, SIMULATED PLANNING DATA.** Every number in this walkthrough comes from `plan_power.R` (Workshop 9, seed 20270208), which plans the follow up study that Workshop 8 called for. No real person was recorded, ALTAVO and BRENLOW are invented brands, participant and staff codes are placeholders, the dates are placeholders, and no number may be quoted as a finding.

This is the Excel route through the parts of `plan_power.R` that a spreadsheet can do: the analytic sample size (STEP 1), the reason why the number of ads matters (STEP 2, as a formula rather than a simulation), the attrition allowance (STEP 4), the Williams design (STEP 5), and the session calendar (STEP 6). The finished workbook is `SIMULATED_planning.xlsx`. **Every cell address below is written as `Sheet!Cell` and was generated from the workbook itself:** `make_excel_workbook.py` writes the workbook and this file from one list of cells, reads the workbook back with openpyxl to confirm that each address holds the formula printed here, and recalculates a copy to confirm the results. A copy was recalculated with LibreOffice 24.2.7.2 420(Build:2) (headless): all 64 Check rows are OK (`Check!B69` = 0), and no cell shows an error value. The 40 exact results (counts, sample sizes, dates, codes) agree with R exactly or at the printed decimals (largest difference 3.3e-05); the 24 approximations agree within their stated tolerances: 0.25 participants for the corrected sample sizes (8 values, largest difference 0.1246); 0.02 for the approximate power of the brand test (8 values, largest difference 0.0110); 0.005 for the approximate paired power (8 values, largest difference 0.0004).

Companion files: `plan_power.R` and `plan_output.txt` (the R planning script and its full output), the `SIMULATED_*.csv` files (the R values the `Check` sheet uses), `analysis_plan_prereg.md`, `planning_checklist.md`, `README.md`.

---

## 0. What you are working with

| Sheet | What it holds |
|---|---|
| `SampleSize` | A: sample size and power of a paired comparison for four values of d_z, with and without a Holm family of three. B: approximate power of the brand test when ads are a random factor |
| `Attrition` | sessions to book from the Workshop 7 exclusion rate; reserve, cap, days, open slots; the Workshop 6 remote comparison |
| `LatinSquare` | the Williams design for 12 ads, the ad labels, the brand sequence of each row, and the balance checks |
| `Assignment` | the Latin square row of each participant code (P01 to P48) and the block check |
| `Calendar` | one row per session slot (55 slots): date, times, phase, code, row, staff pair |
| `Check` | every Excel result next to the R value (64 rows) |

Yellow cells are typed by you (inputs). Gray cells hold values copied from R so that the `Check` sheet can compare them. Excel 2010 or later is needed for `NORM.S.INV`, `NORM.S.DIST`, and `T.INV.2T`.

---

## 1. Sample size of a paired comparison (SampleSize, part A)

A paired comparison means that every participant gives one score in each of two conditions, for example the mean liking of the six ALTAVO ads and the mean liking of the six BRENLOW ads. The effect size d_z is the mean of the participants' differences divided by the SD of those differences (Workshop 8). The ads are treated as fixed: the answer is about THESE ads (section 2 explains the difference).

Type the inputs in `SampleSize!B4:B7` (alpha .05, power .80, three tests in the Holm family, 48 participants for the power columns) and the four effect sizes in `SampleSize!B14:B17`. The three z values:

| Result | Cell | Formula | Value | R |
|---|---|---|---|---|
| z for alpha (two tailed) | `SampleSize!B9` | `=NORM.S.INV(1-B4/2)` | 1.9600 |  |
| z at the strictest Holm step (alpha / 3) | `SampleSize!B10` | `=NORM.S.INV(1-B4/B6/2)` | 2.3940 |  |
| z for the power target | `SampleSize!B11` | `=NORM.S.INV(B5)` | 0.8416 |  |

The formulas of the first effect row (fill them down to row 17):

| Result | Cell | Formula |
|---|---|---|
| n, normal approximation | `SampleSize!C14` | `=(($B$9+$B$11)/B14)^2` |
| n, with the small sample correction | `SampleSize!D14` | `=C14+$B$9^2/2` |
| n needed (always round UP) | `SampleSize!E14` | `=ROUNDUP(D14,0)` |
| Holm: n, normal approximation | `SampleSize!F14` | `=(($B$10+$B$11)/B14)^2` |
| Holm: n, corrected | `SampleSize!G14` | `=F14+$B$10^2/2` |
| Holm: n needed | `SampleSize!H14` | `=ROUNDUP(G14,0)` |
| Power at the n in B7 | `SampleSize!I14` | `=NORM.S.DIST(B14*SQRT($B$7-$B$9^2/2)-$B$9,TRUE)` |
| Holm: power at the n in B7 | `SampleSize!J14` | `=NORM.S.DIST(B14*SQRT($B$7-$B$10^2/2)-$B$10,TRUE)` |

**You should see** (Excel next to R's `power.t.test`, STEP 1 of `plan_output.txt`):

| Effect | d_z | n, normal | n, corrected | R n exact | n needed | R | Holm n, corrected | R Holm n exact | Holm n needed | R |
|---|---|---|---|---|---|---|---|---|---|---|
| small (0.30) | 0.300 | 87.21 | 89.13 | 89.15 | 90 | 90 | 119.19 | 119.20 | 120 | 120 |
| moderate (0.50) | 0.500 | 31.40 | 33.32 | 33.37 | 34 | 34 | 44.74 | 44.78 | 45 | 45 |
| Workshop 8 liking estimate (Q6) | 0.355 | 62.13 | 64.05 | 64.07 | 65 | 65 | 85.73 | 85.75 | 86 | 86 |
| Workshop 7 GSR estimate (H3, Ad_A) | 0.826 | 11.50 | 13.42 | 13.55 | 14 | 14 | 18.21 | 18.30 | 19 | 19 |

| Effect | Power at 48 (Excel) | R | Holm power at 48 (Excel) | R |
|---|---|---|---|---|
| small (0.30) | 0.5305 | 0.5303 | 0.3525 | 0.3525 |
| moderate (0.50) | 0.9242 | 0.9240 | 0.8328 | 0.8324 |
| Workshop 8 liking estimate (Q6) | 0.6747 | 0.6744 | 0.4976 | 0.4974 |
| Workshop 7 GSR estimate (H3, Ad_A) | 0.9999 | 0.9999 | 0.9992 | 0.9992 |

Why the correction: the normal approximation alone pretends that the SD of the differences is known, so it gives about two participants too few; adding half of the squared critical z brings it within a fraction of a participant of the exact t based answer (Guenther, 1981), and the rounded up numbers agree with R in every row. The power formula is the same equation solved for the power. **Common mistake:** rounding the sample size to the nearest whole number instead of UP. With 33.37 participants needed, 33 participants give less than 80 percent power.

**Read:** 34 participants give 80 percent power for one paired test at d_z = 0.50, and 45 at the strictest step of a Holm family of three, the numbers Workshop 8 reported. The secondary paired comparisons of the follow up study are planned as such a family, which is one reason the plan uses 48.

---

## 2. Why more ads beat more participants (SampleSize, part B)

The key planned test of the follow up study asks whether ALTAVO advertising is liked more than BRENLOW advertising IN GENERAL, not only for the twelve ads shown. Then the ads are a sample, and the brand difference must be judged against how much ads of the same brand differ. In a complete, balanced design (every participant rates every ad) the test has 2 x ads per brand - 2 degrees of freedom, and the standard error of the brand difference is

SE = SQRT(2 x (ad SD^2 + residual SD^2 / participants) / ads per brand).

More participants shrink only the second term; only more ads shrink the first. Type the brand difference (0.50), the ad SD (0.20, a planning judgment), and the residual SD from Workshop 7 in `SampleSize!B21:B23`, and the participants and ads per brand in columns A and B of rows 26 to 33. The formulas of row 26:

| Result | Cell | Formula |
|---|---|---|
| df | `SampleSize!C26` | `=2*B26-2` |
| SE of the brand difference | `SampleSize!D26` | `=SQRT(2*($B$22^2+$B$23^2/A26)/B26)` |
| Noncentrality (difference / SE) | `SampleSize!E26` | `=$B$21/D26` |
| Critical t | `SampleSize!F26` | `=T.INV.2T($B$4,C26)` |
| Approximate power | `SampleSize!G26` | `=1-NORM.S.DIST((F26*(1-1/(4*C26))-E26)/SQRT(1+F26^2/(2*C26)),TRUE)` |

**You should see** (R: the exact formula, column `power_formula` of `SIMULATED_power_grid.csv`; the simulated power of `plan_output.txt` agrees with it within Monte Carlo error):

| Participants | Ads per brand | df | SE | Noncentrality | Critical t | Power (Excel) | R (exact) |
|---|---|---|---|---|---|---|---|
| 48 | 2 | 2 | 0.2540 | 1.968 | 4.303 | 0.224 | 0.213 |
| 48 | 3 | 4 | 0.2074 | 2.411 | 2.776 | 0.445 | 0.451 |
| 48 | 4 | 6 | 0.1796 | 2.784 | 2.447 | 0.640 | 0.643 |
| 48 | 6 | 10 | 0.1467 | 3.409 | 2.228 | 0.866 | 0.866 |
| 72 | 2 | 2 | 0.2374 | 2.106 | 4.303 | 0.242 | 0.235 |
| 72 | 3 | 4 | 0.1938 | 2.580 | 2.776 | 0.493 | 0.500 |
| 72 | 4 | 6 | 0.1679 | 2.979 | 2.447 | 0.698 | 0.701 |
| 72 | 6 | 10 | 0.1371 | 3.648 | 2.228 | 0.907 | 0.907 |

Excel has no noncentral t distribution, so the power uses a normal approximation to it (Abramowitz & Stegun, 1964, chapter 26); the `Check` sheet accepts a difference of up to 0.02 from R's exact value. **Read:** with 3 ads per brand, going from 48 to 72 participants hardly moves the power; going from 3 to 6 ads per brand at 48 participants roughly doubles it. That is the reason the plan asks the client for six ads per brand.

---

## 3. Attrition, reserve, and the cap (Attrition sheet)

Type the inputs in `Attrition!B4:B12` (yellow): the target of 48 analyzable participants, the Workshop 7 exclusions (2 of 24), the 12 rows of the Williams design, the cap multiple 1.25, the Workshop 6 funnel (260 opened the link, 178 completed, 142 kept), and 5 slots per day.

| Result | Cell | Formula | Value | R |
|---|---|---|---|---|
| Exclusion rate in the lab (Workshop 7) | `Attrition!B15` | `=B5/B6` | 0.0833 | 0.0833 |
| Sessions expected to end with the target | `Attrition!B16` | `=B4/(1-B15)` | 52.36 |  |
| Sessions to book (main + reserve) | `Attrition!B17` | `=ROUNDUP(B16,0)` | 53 | 53 |
| Reserve sessions | `Attrition!B18` | `=B17-B4` | 5 | 5 |
| Hard cap on sessions | `Attrition!B19` | `=B4*B8` | 60 | 60 |
| Weekdays of data collection | `Attrition!B20` | `=ROUNDUP(B17/B12,0)` | 11 | 11 |
| Slots on those days | `Attrition!B21` | `=B20*B12` | 55 | 55 |
| Open slots for rebooking no shows | `Attrition!B22` | `=B21-B17` | 2 | 2 |
| Remote: completed share of those who opened the link | `Attrition!B25` | `=B10/B9` | 0.6846 | 0.6846 |
| Remote: share of completed respondents excluded | `Attrition!B26` | `=(B10-B11)/B10` | 0.2022 | 0.2022 |
| Remote: kept share of those who opened the link | `Attrition!B27` | `=B11/B9` | 0.5462 | 0.5462 |
| Remote: people who must open the link to keep the target | `Attrition!B28` | `=ROUNDUP(B4/B27,0)` | 88 | 88 |

**Read:** in the lab, 53 sessions are expected to end with 48 usable participants, so 5 reserve sessions are booked; no shows are rebooked into the 2 open slots of the last day; data collection never runs beyond 60 sessions. The same 48 participants in a remote study would need about 88 people to open the link (Workshop 6), so moving this study online would need a new plan. **Common mistake:** multiplying the target by 1 plus the exclusion rate (48 x 1.083 = 52.0, so 52 sessions) instead of dividing it by the share that is kept (48 / 0.917 = 52.4, so 53). Of 52 sessions, 8.3 percent excluded leaves only 47.7 participants on average.

---

## 4. The Williams design (LatinSquare sheet)

A Latin square shows every ad once in every row and once in every position. A Williams design (Williams, 1949) is a Latin square in which, in addition, every ad is IMMEDIATELY followed by every other ad exactly once, so whatever one ad does to the next is balanced as well. For an even number of ads, one square is enough.

1. Type the positions 1 to 12 in `LatinSquare!C5:N5`. Row 1 of the square, in symbols 0 to 11, is 0, 1, 11, 2, 10, 3, ... (`LatinSquare!C6`, filled right to `N6`):

| Result | Cell | Formula |
|---|---|---|
| Row 1, position 1 | `LatinSquare!C6` | `=IF(C$5=1,0,IF(MOD(C$5,2)=0,C$5/2,$B$3-(C$5-1)/2))` |
| Row r: add r - 1 and wrap at 12 | `LatinSquare!C8` | `=MOD(C$6+$B8-1,$B$3)` |
| Ad label (even symbol ALTAVO, odd BRENLOW) | `LatinSquare!C22` | `=IF(MOD(C8,2)=0,"ALTAVO_"&(C8/2+1),"BRENLOW_"&((C8+1)/2))` |

   **You should see** row 1 = 0, 1, 11, 2, 10, 3, 9, 4, 8, 5, 7, 6. Fill the symbol formula over `C8:N19` (row numbers 1 to 12 in `B8:B19`) and the label formula over `C22:N33`.
2. Why the labels go by even and odd symbols: which symbols belong to which brand does not change the balance of the ADS, but it decides the order of the BRANDS. `plan_power.R` tried all 462 ways to split the 12 symbols into two brands; only the even and odd split keeps every row to at most two ads of one brand in a row. Column `O` shows the brand sequence: row 1 is `ABBAABBAABBA` and row 2 is `BAABBAABBAAB`.
3. The checks (all must hold):

| Check | Cell | Formula | Value | Must be |
|---|---|---|---|---|
| Each ad once per row (smallest count) | `LatinSquare!C50` | `=MIN(C37:N48)` | 1 | 1 |
| Each ad once per row (largest count) | `LatinSquare!C51` | `=MAX(C37:N48)` | 1 | 1 |
| Each ad once per position (smallest) | `LatinSquare!C67` | `=MIN(C54:N65)` | 1 | 1 |
| Each ad once per position (largest) | `LatinSquare!C68` | `=MAX(C54:N65)` | 1 | 1 |
| Ordered pair count, ad i then ad j | `LatinSquare!C72` | `=SUMPRODUCT(($C$8:$M$19=$B72)*($D$8:$N$19=C$71))` | (a 12 x 12 table) | 0 on the diagonal, 1 elsewhere |
| Sum of all pair counts | `LatinSquare!C85` | `=SUM(C72:N83)` | 132 | 132 |
| Diagonal (an ad after itself) | `LatinSquare!C86` | `=C72+D73+...+N83` | 0 | 0 |
| Largest pair count | `LatinSquare!C87` | `=MAX(C72:N83)` | 1 | 1 |
| ALTAVO ads in position 1 | `LatinSquare!C89` | `=SUMPRODUCT(--(MOD(C8:C19,2)=0))` | 6 | 6 |
| Labels that differ from the R square | `LatinSquare!C107` | `=SUMPRODUCT(--(C22:N33<>C94:N105))` | 0 | 0 |

   With 132 adjacent pairs in all, none on the diagonal, and no count above 1, each of the 132 ordered pairs of different ads occurs exactly once. **Common mistake:** a plain cyclic Latin square (1 2 3 ... 12, then 2 3 ... 12 1) is balanced for position but not for carryover: ad 2 then always follows ad 1.

---

## 5. Who gets which row (Assignment sheet)

`plan_power.R` assigned the 48 main participants in 4 blocks of 12: each block uses every row once, in a random order drawn with a fixed seed. If data collection stopped after any complete block, every row would still be used equally often. Excel cannot reproduce R's random numbers, so the codes, blocks, and rows are typed (yellow) in `Assignment!A4:C51`. To draw your own order in Excel: list rows 1 to 12 for a block, put `=RAND()` next to each, copy and Paste Special > Values, and sort by that column (the values must be pasted, because RAND() changes on every recalculation).

| Result | Cell | Formula | Value | R |
|---|---|---|---|---|
| First ad of P01 | `Assignment!D4` | `=INDEX(LatinSquare!$C$22:$C$33,C4)` | BRENLOW_2 | |
| Every row once in every block (smallest count) | `Assignment!H18` | `=MIN(H5:K16)` | 1 | 1 |
| Every row once in every block (largest count) | `Assignment!H19` | `=MAX(H5:K16)` | 1 | 1 |
| Participants who see an ALTAVO ad first | `Assignment!H20` | `=COUNTIF(E4:E51,"ALTAVO")` | 24 | 24 |

Reserve participants (R01 to R05) are not in this table: each one takes the row of the participant it replaces, so that the rows stay balanced among the analyzed participants.

---

## 6. The session calendar (Calendar sheet)

The lab has one Smart Eye Aurora, so sessions run one at a time. A slot is 90 minutes: 15 for arrival and consent, 10 for the sensors and calibration, 20 for the recording, 15 for removing the sensors and the debriefing, and 30 for the backup, the session log, and cleaning (the lab procedure manual's times; see STEP 6 of `plan_output.txt`). Type the first day (a placeholder) in `Calendar!B3`, the first slot (09:00) in `B5`, and the slot length (90) in `B6`; `B4`, `B7`, and `B8` come from the Attrition sheet. Type k = 1 to 55 in column A and fill the formulas of row 12 down:

| Result | Cell | Formula |
|---|---|---|
| Day index (0 = first day) | `Calendar!B12` | `=INT((A12-1)/$B$4)` |
| Date (weekdays only) | `Calendar!C12` | `=WORKDAY($B$3,B12)` |
| Weekday | `Calendar!D12` | `=TEXT(C12,"dddd")` |
| Slot of the day (1 to 5) | `Calendar!E12` | `=MOD(A12-1,$B$4)+1` |
| Arrival time | `Calendar!F12` | `=$B$5+(E12-1)*$B$6/1440` |
| Recording starts (25 min later) | `Calendar!G12` | `=F12+25/1440` |
| Room ready for the next participant | `Calendar!J12` | `=F12+$B$6/1440` |
| Phase | `Calendar!K12` | `=IF(A12<=$B$7,"main",IF(A12<=$B$7+$B$8,"reserve","open"))` |
| Participant code | `Calendar!L12` | `=IF(K12="main","P"&TEXT(A12,"00"),IF(K12="reserve","R"&TEXT(A12-$B$7,"00"),""))` |
| Latin square row | `Calendar!M12` | `=IF(K12="main",INDEX(Assignment!$C$4:$C$51,MATCH(L12,Assignment!$A$4:$A$51,0)),"")` |
| Half day shift (morning = slots 1 to 3) | `Calendar!N12` | `=2*B12+IF(E12>=4,1,0)` |
| Staff pair (1, 2, 3 in turn) | `Calendar!O12` | `=MOD(N12,3)+1` |
| Lead operator | `Calendar!P12` | `=CHOOSE(O12,"S1","S3","S5")` |
| Assistant | `Calendar!Q12` | `=CHOOSE(O12,"S2","S4","S6")` |

Times are fractions of a day in Excel, so 25 minutes is `25/1440`. Format the date column as `yyyy-mm-dd` and the time columns as `hh:mm`. `WORKDAY` skips Saturdays and Sundays only; university closures must be added as its third argument (a range of holiday dates) once the JSU academic calendar for the term is known. **Common mistake:** formatting the time cells as numbers, which shows 0.375 instead of 09:00.

The columns `T` to `Y` (gray) hold R's schedule, and these cells count the differences:

| Check | Cell | Formula | Value | Must be |
|---|---|---|---|---|
| dates that differ from R | `Calendar!T3` | `=SUMPRODUCT(--(C12:C66<>T12:T66))` | 0 | 0 |
| arrival times that differ from R | `Calendar!T4` | `=SUMPRODUCT(--(ROUND(F12:F66*1440,0)<>ROUND(U12:U66*1440,0)))` | 0 | 0 |
| participant codes that differ from R | `Calendar!T5` | `=SUMPRODUCT(--(L12:L66&""<>V12:V66&""))` | 0 | 0 |
| Latin square rows that differ from R | `Calendar!T6` | `=SUMPRODUCT(--(M12:M66&""<>W12:W66&""))` | 0 | 0 |
| lead operators that differ from R | `Calendar!T7` | `=SUMPRODUCT(--(P12:P66<>X12:X66))` | 0 | 0 |
| assistants that differ from R | `Calendar!T8` | `=SUMPRODUCT(--(Q12:Q66<>Y12:Y66))` | 0 | 0 |
| Last day of data collection | `Calendar!T9` | `=MAX(C12:C66)` | 2027-02-22 | 2027-02-22 |

---

## 7. Check your numbers (Check sheet)

Every result with an R value sits on the `Check` sheet next to that value, with the difference and a tolerance: half a unit in the last printed decimal for exact results (0 for counts, sample sizes, dates, and codes), 0.25 participants for the corrected sample sizes, 0.005 for the paired power, and 0.02 for the approximate power of the brand test. `Check!B69` counts the rows with status CHECK and must read 0; in the LibreOffice recalculation it reads 0 for 64 rows. What Excel cannot do is listed at the bottom of the sheet: the mixed model simulations (STEP 2 and STEP 3), the exact noncentral t power, and R's random draw of the rows.

---

## Sources

- Abramowitz, M., & Stegun, I. A. (Eds.). (1964). *Handbook of mathematical functions with formulas, graphs, and mathematical tables* (National Bureau of Standards Applied Mathematics Series 55). U.S. Government Printing Office.
- Guenther, W. C. (1981). Sample size formulas for normal theory t tests. *The American Statistician, 35*(4), 243-244. https://doi.org/10.1080/00031305.1981.10479363
- Williams, E. J. (1949). Experimental designs balanced for the estimation of residual effects of treatments. *Australian Journal of Scientific Research, Series A, 2*(2), 149-168. https://doi.org/10.1071/CH9490149
