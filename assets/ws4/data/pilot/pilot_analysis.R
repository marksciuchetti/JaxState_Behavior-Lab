# =============================================================================
# pilot_analysis.R
# Workshop 4: reading and analyzing a REAL iMotions Sensor Data export
# Martin's Family Clothing Consumer Behavior Research Laboratory, JSU
#
# REAL DATA, shared with permission. 001_PILOT01.csv is the Sensor Data export of a pilot
# recorded on the lab computer on October 12, 2026 (iMotions Lab 11.1.47136.14, study
# "W4 Pilot - Shimmer test"). The respondent is the lab director, who agreed to share his own
# recording, including the age and gender in the file header, as practice data. A real
# participant's export is identifiable data and is never published.
#
# Design of the pilot: one seated respondent wearing one Shimmer3 GSR+ (unit Shimmer3-8956,
# 128 Hz, finger electrodes and a pulse probe) watched a 60 s quiet baseline, a 30 s e.l.f.
# commercial, and a 15 s quiet recovery. The analysis "W4 Pilot - Shimmer test-1" ran GSR
# Signal Quality, GSR Peak Detection, and GSR Peak Binned Aggregation; the export includes the
# raw Shimmer channels and the nine GSR Peak Detection channels.
#
# ONE RESPONDENT: every number below DESCRIBES this recording. Nothing is tested, and nothing
# can be concluded about the commercial.
#
# What the script does, in order:
#   0. Settings: the file name and the artifact rule (Kleckner et al., 2018)
#   1. Read the header: the #INFO block and the '#Sensor info' rows
#   2. Build a codebook from the #METADATA block (unit, description, data format per column)
#   3. Read the data rows and give the repeated EventSource columns clear names
#   4. Split the rows by device (Shimmer, GSR Peak Detection, System Load, slide events)
#   5. Find each stimulus from the slide events (StartMedia to EndMedia)
#   6. Clean heart rate and IBI: -1 means "no value"
#   7. Apply the artifact rule to skin conductance
#   8. Build a peak table from the Peak Detected and Peak Occurrences columns
#   9. Summarize each stimulus
#  10. Record the GSR Peak Detection settings from the header
#  11. Draw a two panel figure (base R graphics, no packages needed)
#  12. Write the output files
#  13. Check the results against iMotions' own GSR Metrics export (GSRPeakMetrics.csv and
#      GSRPeakTable.csv, from Export > R PROCESSING > GSR Metrics Export)
#
# Run it from the folder that holds 001_PILOT01.csv:  Rscript pilot_analysis.R
# Needs only base R (version 4.0 or later).
# =============================================================================

# ---- 0. Settings -------------------------------------------------------------------------
data_file <- "001_PILOT01.csv"
out_dir   <- "."                      # output files are written next to the data file
# Artifact rule from Kleckner et al. (2018), Table I. Skin conductance outside 0.05 to 60
# microsiemens, or changing faster than 10 microsiemens per second, is invalid, and so is the
# data within 5 seconds of it. (The paper's fourth rule, skin temperature, does not apply:
# the Shimmer3 GSR+ records no temperature channel.)
rule_min_uS     <- 0.05
rule_max_uS     <- 60
rule_max_rate   <- 10     # microsiemens per second
rule_margin_ms  <- 5000   # 5 seconds on either side

options(width = 140)
cat("pilot_analysis.R  run on", format(Sys.time(), "%Y-%m-%d %H:%M"), "\n")
cat("REAL DATA (lab director's own pilot, shared with permission). One respondent: description only.\n\n")

# ---- 1. Read the header ------------------------------------------------------------------
# The file starts with a byte order mark and the #INFO and #METADATA blocks. The data table
# starts at the line whose first field is "Row".
con <- file(data_file, encoding = "UTF-8-BOM")   # "UTF-8-BOM" drops the byte order mark
lines <- readLines(con, n = 200, warn = FALSE)
close(con)
first_field <- sub(",.*$", "", lines)
header_line <- which(first_field == "Row")[1]
cat("STEP 1  The data table starts on line", header_line, "of the file.\n")

split_csv <- function(x) scan(text = x, what = "", sep = ",", quote = "\"", quiet = TRUE,
                              na.strings = character(0))
info_rows <- lines[1:(header_line - 1)]
info <- lapply(info_rows, split_csv)
info_names <- vapply(info, function(v) if (length(v)) v[1] else "", "")
show_info <- c("#Study name", "#Respondent Name", "#iMotions version", "#Recording time", "#Export time")
for (nm in show_info) {
  v <- info[[which(info_names == nm)[1]]]
  v <- v[nzchar(v)]
  cat(sprintf("        %-18s %s\n", nm, paste(v[-1], collapse = " | ")))
}
sensor_rows <- info[info_names == "#Sensor info"]
cat("        '#Sensor info' rows (one per device or notebook):\n")
for (v in sensor_rows) cat("          -", v[2], "\n")

# ---- 2. Codebook from the #METADATA block ------------------------------------------------
header <- split_csv(lines[header_line])
meta_names <- c("#Device", "#Provided By", "#Category", "#Description", "#Unit", "#Short unit",
                "#Instance", "#Group", "#Display name", "#Data format", "#Channel identifier")
get_meta <- function(nm) {
  v <- info[[which(info_names == nm)[1]]]
  length(v) <- length(header)
  v[is.na(v)] <- ""
  v
}
codebook <- data.frame(Column = header, stringsAsFactors = FALSE)
for (nm in meta_names) codebook[[sub("^#", "", nm)]] <- get_meta(nm)
codebook <- codebook[-1, ]   # the first column, Row, has no metadata
cat("\nSTEP 2  Codebook built:", nrow(codebook), "columns described by the #METADATA block.\n")
cat("        Units and data formats of the Shimmer channels used below:\n")
show_cols <- c("GSR Conductance CAL", "GSR Resistance CAL", "Internal ADC A13 PPG CAL",
               "Heart Rate PPG ALG", "IBI PPG ALG", "Wide Range Accelerometer Z CAL",
               "Packet reception rate RAW")
for (cn in show_cols) {
  r <- codebook[codebook$Column == cn, ][1, ]
  cat(sprintf("          %-32s unit: %-24s format: %s\n", cn, r$Unit, r$`Data format`))
}

# ---- 3. Read the data rows ---------------------------------------------------------------
con <- file(data_file, encoding = "UTF-8-BOM")
all_lines <- readLines(con, warn = FALSE)
close(con)
dat <- read.csv(text = all_lines[header_line:length(all_lines)], header = TRUE,
                check.names = FALSE, stringsAsFactors = FALSE)
# Four columns are all called EventSource, one per device. Name them by device.
es <- which(names(dat) == "EventSource")
names(dat)[es] <- paste0("EventSource_", c("SlideEvents", "SystemLoad", "PeakDetection", "Shimmer"))
cat("\nSTEP 3  Data rows:", nrow(dat), " Columns:", ncol(dat), "\n")

# ---- 4. Split the rows by device ---------------------------------------------------------
is_on <- function(x) !is.na(x) & x == 1
shim  <- dat[is_on(dat$EventSource_Shimmer), ]
pdet  <- dat[is_on(dat$EventSource_PeakDetection), ]
sysld <- dat[is_on(dat$EventSource_SystemLoad), ]
slide <- dat[is_on(dat$EventSource_SlideEvents), ]
both  <- sum(is_on(dat$EventSource_Shimmer) & is_on(dat$EventSource_PeakDetection))
cat("STEP 4  Shimmer rows:", nrow(shim), "| Peak Detection rows:", nrow(pdet),
    "(", both, "shared with the Shimmer,", nrow(pdet) - both, "of its own )",
    "| System Load rows:", nrow(sysld), "| slide event rows:", nrow(slide), "\n")
cat("        Rows from different devices are interleaved: always filter on EventSource first.\n")

# ---- 5. Stimulus windows -----------------------------------------------------------------
stims <- unique(slide$SourceStimuliName)
win <- do.call(rbind, lapply(stims, function(s) {
  g <- slide[slide$SourceStimuliName == s, ]
  data.frame(Stimulus = s,
             Start_ms = g$Timestamp[g$SlideEvent == "StartMedia"][1],
             End_ms   = g$Timestamp[g$SlideEvent == "EndMedia"][1],
             stringsAsFactors = FALSE)
}))
win$Length_s <- (win$End_ms - win$Start_ms) / 1000
in_win <- function(t, i) t >= win$Start_ms[i] & t <= win$End_ms[i]
cat("\nSTEP 5  Stimuli (StartMedia to EndMedia, seconds since the recording started):\n")
for (i in seq_len(nrow(win)))
  cat(sprintf("          %-42s %6.2f to %6.2f s  (%.2f s)\n", win$Stimulus[i],
              win$Start_ms[i] / 1000, win$End_ms[i] / 1000, win$Length_s[i]))

# ---- 6. Clean heart rate and IBI ---------------------------------------------------------
hr  <- shim$`Heart Rate PPG ALG`
ibi <- shim$`IBI PPG ALG`
cat("\nSTEP 6  Heart Rate PPG ALG: rows equal to -1:", sum(hr == -1, na.rm = TRUE),
    "| equal to 0:", sum(hr == 0, na.rm = TRUE), "\n")
neg <- shim$Timestamp[hr == -1]
if (length(neg)) cat(sprintf("        The -1 rows run from %.1f to %.1f s: no heart rate was available there.\n",
                             min(neg) / 1000, max(neg) / 1000))
cat("        IBI PPG ALG: rows with a value:", sum(ibi > 0, na.rm = TRUE), "of", length(ibi),
    "(one per detected beat; -1 everywhere else).\n")
hr_clean  <- ifelse(hr > 0, hr, NA)
ibi_clean <- ifelse(ibi > 0, ibi, NA)

# ---- 7. Artifact rule --------------------------------------------------------------------
t_ms <- shim$Timestamp
sc   <- shim$`GSR Conductance CAL`
rate <- c(NA, diff(sc) / (diff(t_ms) / 1000))           # microsiemens per second
bad  <- (sc < rule_min_uS | sc > rule_max_uS) | (!is.na(rate) & abs(rate) > rule_max_rate)
# The step into a sample is flagged on both of its samples
bad  <- bad | c(bad[-1], FALSE)
bad_t <- t_ms[bad]
flag <- vapply(t_ms, function(t) any(abs(bad_t - t) <= rule_margin_ms), logical(1))
cat("\nSTEP 7  Artifact rule (Kleckner et al., 2018): samples breaking the range or rate rule:",
    sum(bad), "\n")
if (length(bad_t)) {
  runs <- split(bad_t, cumsum(c(1, diff(bad_t) > 1000)))
  for (r in runs) cat(sprintf("        flagged stretch %.2f to %.2f s; fastest change %.1f microsiemens per second\n",
                              min(r) / 1000, max(r) / 1000,
                              max(abs(rate[t_ms >= min(r) & t_ms <= max(r)]), na.rm = TRUE)))
}
cat("        Samples inside the 5 s margin (set aside):", sum(flag), "of", length(flag), "\n")

# ---- 8. Peak table -----------------------------------------------------------------------
pd_t   <- pdet$Timestamp
inpk   <- pdet$`Peak Detected` == 1
starts <- which(inpk & !c(FALSE, head(inpk, -1)))
ends   <- which(inpk & !c(tail(inpk, -1), FALSE))
peaks <- do.call(rbind, lapply(seq_along(starts), function(k) {
  idx <- starts[k]:ends[k]
  top <- idx[pdet$`Peak Occurrences`[idx] == 1][1]
  data.frame(Peak = k,
             Onset_s    = pd_t[starts[k]] / 1000,
             Peak_s     = pd_t[top] / 1000,
             Offset_s   = pd_t[ends[k]] / 1000,
             Duration_ms = pd_t[ends[k]] - pd_t[starts[k]],
             Amplitude_uS = pdet$`Peak Amplitude`[top],
             Rise_GSR_Interpolated_uS = max(pdet$`GSR Interpolated`[idx]) - pdet$`GSR Interpolated`[starts[k]],
             Rise_Phasic_uS = max(pdet$`Phasic Signal`[idx]) - pdet$`Phasic Signal`[starts[k]],
             stringsAsFactors = FALSE)
}))
peaks$Stimulus <- vapply(peaks$Peak_s * 1000, function(t) {
  i <- which(t >= win$Start_ms & t <= win$End_ms); if (length(i)) win$Stimulus[i[1]] else "between stimuli"
}, "")
peaks$Inside_artifact_margin <- vapply(peaks$Peak_s * 1000, function(t) any(abs(bad_t - t) <= rule_margin_ms), logical(1))
cat("\nSTEP 8  Peaks found by GSR Peak Detection:", nrow(peaks), "\n")
cat("        Peak Amplitude equals the rise of GSR Interpolated (the manuals' definition) in",
    sum(abs(peaks$Amplitude_uS - peaks$Rise_GSR_Interpolated_uS) < 1e-6), "of", nrow(peaks),
    "peaks, and the rise of the Phasic Signal in",
    sum(abs(peaks$Amplitude_uS - peaks$Rise_Phasic_uS) < 1e-6), "\n")
cat("        Peaks inside the artifact margin:", sum(peaks$Inside_artifact_margin), "\n")
print(format(peaks[, c("Peak", "Onset_s", "Peak_s", "Duration_ms", "Amplitude_uS", "Stimulus",
                       "Inside_artifact_margin")], digits = 4), row.names = FALSE)

# ---- 9. Summary by stimulus --------------------------------------------------------------
summ <- do.call(rbind, lapply(seq_len(nrow(win)), function(i) {
  s_in <- in_win(shim$Timestamp, i)
  p_in <- in_win(pdet$Timestamp, i)
  k    <- peaks$Stimulus == win$Stimulus[i]
  k_ok <- k & !peaks$Inside_artifact_margin
  len_min <- win$Length_s[i] / 60
  data.frame(
    Stimulus = win$Stimulus[i],
    Length_s = round(win$Length_s[i], 2),
    Shimmer_rows = sum(s_in),
    Samples_per_s = round(sum(s_in) / win$Length_s[i], 1),
    Interpolated_samples = sum(pdet$`Signal Interpolation`[p_in] == 1, na.rm = TRUE),
    Packet_reception_pct = round(mean(shim$`Packet reception rate RAW`[s_in], na.rm = TRUE), 1),
    Peaks = sum(k),
    Peaks_per_min = round(sum(k) / len_min, 2),
    Peaks_after_rule = sum(k_ok),
    Peaks_per_min_after_rule = round(sum(k_ok) / len_min, 2),
    Mean_peak_amplitude_uS = round(mean(peaks$Amplitude_uS[k]), 3),
    Mean_tonic_uS = round(mean(pdet$`Tonic Signal`[p_in], na.rm = TRUE), 3),
    Mean_HR_bpm_clean = round(mean(hr_clean[s_in], na.rm = TRUE), 1),
    Mean_HR_bpm_WRONG_with_codes = round(mean(hr[s_in], na.rm = TRUE), 1),
    HR_from_mean_IBI_bpm = round(60000 / mean(ibi_clean[s_in], na.rm = TRUE), 1),
    stringsAsFactors = FALSE)
}))
cat("\nSTEP 9  Summary by stimulus (description only, one respondent):\n")
print(summ, row.names = FALSE)
cat("        Mean_HR_bpm_WRONG_with_codes shows what happens if the -1 rows are left in.\n")

# ---- 10. Peak Detection settings from the header -----------------------------------------
pd_info <- sensor_rows[[which(vapply(sensor_rows, function(v) grepl("Peak Detection", v[2]), logical(1)))[1]]]
kv <- pd_info[grepl("=", pd_info)]
settings <- data.frame(Setting = gsub('"', "", sub("=.*$", "", kv)),
                       Value = gsub('"', "", sub("^[^=]*=", "", kv)), stringsAsFactors = FALSE)
cat("\nSTEP 10 GSR Peak Detection settings written in the export header (copy into the methods):\n")
print(settings, row.names = FALSE)

# ---- 11. Figure --------------------------------------------------------------------------
png(file.path(out_dir, "pilot_figure_R.png"), width = 1600, height = 1000, res = 150)
par(mfrow = c(2, 1), mar = c(4, 4.5, 2.5, 1))
plot(t_ms / 1000, sc, type = "l", col = "#2a78d6", xlab = "", ylab = "Skin conductance (microsiemens)",
     main = "GSR Conductance CAL with detected peaks (open circle: inside the artifact margin)",
     cex.main = 0.9, ylim = c(0, 2.6))
abline(v = c(win$Start_ms, win$End_ms) / 1000, col = "grey70", lty = 3)
rect(min(bad_t) / 1000, -1, max(bad_t) / 1000, 3, col = adjustcolor("grey50", 0.25), border = NA)
pk_y <- approx(t_ms / 1000, sc, xout = peaks$Peak_s)$y
points(peaks$Peak_s, pk_y + 0.1, pch = ifelse(peaks$Inside_artifact_margin, 1, 25), bg = "black", cex = 0.9)
plot(t_ms / 1000, hr_clean, type = "s", col = "#eb6834", xlab = "Time since the recording started (s)",
     ylab = "Beats per minute", main = "Heart Rate PPG ALG after removing the -1 codes", cex.main = 0.9)
abline(v = c(win$Start_ms, win$End_ms) / 1000, col = "grey70", lty = 3)
invisible(dev.off())
cat("\nSTEP 11 Figure written: pilot_figure_R.png\n")

# ---- 12. Output files --------------------------------------------------------------------
write.csv(codebook, file.path(out_dir, "pilot_codebook.csv"), row.names = FALSE)
write.csv(peaks,    file.path(out_dir, "pilot_peak_table.csv"), row.names = FALSE)
write.csv(summ,     file.path(out_dir, "pilot_summary_by_stimulus.csv"), row.names = FALSE)
write.csv(settings, file.path(out_dir, "pilot_peak_detection_settings.csv"), row.names = FALSE)
cat("STEP 12 Written: pilot_codebook.csv, pilot_peak_table.csv, pilot_summary_by_stimulus.csv,",
    "pilot_peak_detection_settings.csv\n")
# ---- 13. Check against iMotions' own GSR Metrics export ---------------------------------
# The two files start with a #METADATA block; the table starts at the line "Study Name,...".
read_imotions_metrics <- function(f) {
  con <- file(f, encoding = "UTF-8-BOM"); x <- readLines(con, warn = FALSE); close(con)
  start <- which(sub(",.*$", "", x) == "Study Name")[1]
  read.csv(text = x[start:length(x)], check.names = FALSE, stringsAsFactors = FALSE)
}
if (file.exists("GSRPeakMetrics.csv") && file.exists("GSRPeakTable.csv")) {
  im <- read_imotions_metrics("GSRPeakMetrics.csv")
  it <- read_imotions_metrics("GSRPeakTable.csv")
  cmp <- data.frame(Stimulus = summ$Stimulus,
                    Peaks_script = summ$Peaks,
                    Peaks_iMotions = im$`Peak Count`[match(summ$Stimulus, im$Label)],
                    PPM_script = summ$Peaks_per_min,
                    PPM_iMotions = im$`Peaks Per Minute`[match(summ$Stimulus, im$Label)],
                    Amp_script = summ$Mean_peak_amplitude_uS,
                    Amp_iMotions = im$`Average Peak Amplitude`[match(summ$Stimulus, im$Label)],
                    stringsAsFactors = FALSE)
  cat("\nSTEP 13 The script against iMotions' GSRPeakMetrics.csv (same peaks, same rates):\n")
  print(cmp, row.names = FALSE)
  same_times <- nrow(it) == nrow(peaks) &&
    all(abs(it$`Peak Time` - round(peaks$Peak_s * 1000)) <= 1) &&
    all(abs(it$`Peak Onset Time` - round(peaks$Onset_s * 1000)) <= 1)
  cat("        GSRPeakTable.csv lists", nrow(it), "peaks; onset and peak times match the script to 1 ms:",
      same_times, "\n")
  early <- it[it$`Peak Onset Time` < it$`Stimulus Start`, ]
  if (nrow(early)) for (k in seq_len(nrow(early)))
    cat(sprintf("        Note: peak %d is counted in \"%s\" (its maximum came at %.2f s), but its onset at %.2f s came\n              before that stimulus started at %.2f s, so it began during the previous stimulus.\n",
                early$`Peak Index`[k], early$`Stimulus Label`[k], early$`Peak Time`[k] / 1000,
                early$`Peak Onset Time`[k] / 1000, early$`Stimulus Start`[k] / 1000))
} else {
  cat("\nSTEP 13 GSRPeakMetrics.csv and GSRPeakTable.csv not found next to the data file; check skipped.\n")
}

cat("\nReminder: one respondent. These numbers describe a recording; they test nothing.\n")
