# Workshop 4 pilot practice files (REAL data, shared with permission)

These files come from a real pilot recorded on the lab computer on October 12, 2026 (iMotions Lab 11.1.47136.14, study "W4 Pilot - Shimmer test"). The respondent is the lab director, who agreed to share his own recording, including the age and gender in the file header, as practice data for students. A real participant's export is identifiable data and is never published.

One respondent, one Shimmer3 GSR+ at 128 Hz, a 60 s baseline, a 30 s e.l.f. commercial, and a 15 s recovery. The numbers describe a recording; they test nothing.

| File | What it is |
|---|---|
| 001_PILOT01.csv | The Sensor Data export as iMotions wrote it (#INFO, #METADATA, 13,578 data rows, 43 columns) |
| pilot_analysis.R | The analysis in base R: header, codebook, device filter, stimuli, heart rate cleaning, the artifact rule of Kleckner et al. (2018), peak table, summary, settings, figure |
| pilot_analysis_output.txt | What the script prints |
| pilot_gsr_hr_workbook.xlsx | The same analysis as live Excel formulas, with a Check sheet against the script |
| pilot_codebook.csv | The #METADATA rows for every column |
| pilot_peak_table.csv | One row per detected peak |
| pilot_summary_by_stimulus.csv | Peaks, Peaks Per Minute, tonic level, and heart rate per stimulus |
| pilot_peak_detection_settings.csv | The GSR Peak Detection settings from the export header |
| pilot_figure_R.png | The script's figure |
| GSRPeakMetrics.csv | iMotions' own peak metrics per stimulus (Export > R PROCESSING > GSR Metrics Export > Peak metrics) |
| GSRPeakTable.csv | iMotions' own table of the 14 peaks (Peak table) |
| 001_PILOT01_peak_detection_only.csv | A second Sensor Data export, from the analysis, holding only the GSR Peak Detection signals |

Walkthrough: Workshop 4, section 7.7, https://marksciuchetti.github.io/JaxState_Behavior-Lab/workshops/workshop-04-shimmer-gsr-heart-rate.html

Reference for the artifact rule: Kleckner, I. R., Jones, R. M., Wilder-Smith, O., Wormwood, J. B., Akcakaya, M., Quigley, K. S., Lord, C., & Goodwin, M. S. (2018). Simple, transparent, and flexible automated quality assessment procedures for ambulatory electrodermal activity data. IEEE Transactions on Biomedical Engineering, 65(7), 1460-1467. https://doi.org/10.1109/TBME.2017.2758643
