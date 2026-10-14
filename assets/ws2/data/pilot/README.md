# Workshop 2 pilot practice files (REAL data, shared with permission)

These files come from a real pilot recorded on the lab computer on October 13, 2026 (iMotions Lab 11.1.47136.14, study "PILOT1_2026-10_AFFDEX_Aurora", a copy of the lab's e.l.f. Creative Evaluation). The respondent is the lab director, who agreed to share his own recording, including the age and gender in the file header, as practice data for students. A real participant's export is identifiable data and is never published.

One respondent, the Smart Eye Aurora at 60 Hz on a 2560 x 1440 pixel screen at 65 cm, and the respondent camera with AFFDEX. The flow: a 60 s baseline, Creative A (a still advertisement) for 15 s, a 15 s recovery, and a short questionnaire, then the same for Creative B. The advertisements are not included; the files hold coordinates and times only. The numbers describe a recording; they test nothing.

| File | What it is |
|---|---|
| 001_PILOT01.csv | The Sensor Data export as iMotions wrote it (#INFO, #METADATA, 20,554 data rows, 168 columns; 17 MB) |
| pilot_et_analysis.R | The analysis in base R: header and I-VT settings, codebook, device filter, stimuli, sampling rate and gaze data, fixation counts, AOI metrics rebuilt from the fixation table, questionnaire AOIs, figure |
| pilot_et_analysis_output.txt | What the script prints |
| pilot_et_codebook.csv | The #METADATA device and unit for every column |
| pilot_et_summary_by_stimulus.csv | Rows, gaze percentage, eye to tracker distance, and fixation counts per stimulus |
| pilot_aoi_metrics_rebuilt.csv | Time to first fixation, fixation count, dwell time, and revisits per AOI, each beside iMotions' own value |
| pilot_et_figure_R.png | The script's figure |
| IndividualAOImetrics.csv | iMotions' own AOI metrics per stimulus (Export > R PROCESSING > Individual metrics exports) |
| AggregatedAOImetrics.csv | iMotions' own AOI metrics aggregated over respondents (one respondent here) |
| FixationTable.csv | iMotions' own table of every fixation, with Truncation Type and the AOI it hit |
| SaccadeTable.csv | iMotions' own table of every saccade |

Walkthrough: Workshop 2, section 7.7, https://marksciuchetti.github.io/JaxState_Behavior-Lab/workshops/workshop-02-aurora-ad-test.html

The same recording, read for the face data, is in Workshop 3, section 7.5.
