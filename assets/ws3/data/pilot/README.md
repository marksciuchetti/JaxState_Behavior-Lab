# Workshop 3 pilot practice files (REAL data, shared with permission)

These files come from a real pilot recorded on the lab computer on October 13, 2026 (iMotions Lab 11.1.47136.14, study "PILOT1_2026-10_AFFDEX_Aurora", a copy of the lab's e.l.f. Creative Evaluation). The respondent is the lab director, who agreed to share his own recording, including the age and gender in the file header, as practice data for students. A real participant's export is identifiable data and is never published.

One respondent, the respondent camera with Affectiva AFFDEX (post processed), and the Smart Eye Aurora. The analysis PILOT1_Master used the FEA thresholding flow at its factory defaults. The flow: a 60 s baseline, Creative A (a still advertisement) for 15 s, a 15 s recovery, and a short questionnaire, then the same for Creative B. No face video and no advertisement is included; the files hold scores and times only. The numbers describe a recording; they test nothing.

| File | What it is |
|---|---|
| 001_PILOT01.csv | The Sensor Data export as iMotions wrote it (#INFO, #METADATA, 20,554 data rows, 168 columns; 17 MB; the same file as in Workshop 2) |
| FEAMetrics.csv | iMotions' own FEA Metrics Export, Expression metrics: one row per stimulus with Count Frames and, for each expression, Threshold, Frames >= Threshold, and Time Percentage |
| FEAExpressionTable.csv | iMotions' own FEA Metrics Export, Expression table: one row per episode with onset, offset, duration, and Truncation Type (494 rows) |
| pilot_fea_analysis.R | The analysis in base R: header and FEA thresholding settings, codebook, device filter, stimuli, frame rate and face found, time above the threshold, raw score against the threshold, episodes, Brow Furrow at thresholds of 25 and 75, figure |
| pilot_fea_analysis_output.txt | What the script prints |
| pilot_fea_codebook.csv | The #METADATA device and unit for every column |
| pilot_fea_summary_by_stimulus.csv | Frames, frame rate, face found, Attention, Interocular Distance, and the time above the threshold for six expressions per stimulus |
| pilot_fea_threshold_sensitivity.csv | Brow Furrow on the two creatives at thresholds of 25, 50, and 75 |
| pilot_fea_figure_R.png | The script's figure |

Walkthrough: Workshop 3, section 7.5, https://marksciuchetti.github.io/JaxState_Behavior-Lab/workshops/workshop-03-affdex-facial-expression.html

The same recording, read for the eye tracking data, is in Workshop 2, section 7.7.
