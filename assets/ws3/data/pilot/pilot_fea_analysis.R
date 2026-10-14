# pilot_fea_analysis.R
# Workshop 3, section 7.5: practice on the real AFFDEX pilot file.
# Base R only; no packages to install. Run it from the folder that holds the files:
#   001_PILOT01.csv (Sensor Data export), FEAMetrics.csv, FEAExpressionTable.csv
# The respondent is the lab director (PILOT01), who shared his own recording as practice data.
# One respondent: the numbers describe a recording; they test nothing.
options(width = 150)
f <- "001_PILOT01.csv"
lines <- readLines(file(f, encoding = "UTF-8-BOM"), warn = FALSE)

# STEP 1. Read the #INFO header ---------------------------------------------
cat("STEP 1. Header\n")
info <- lines[1:(grep("^#METADATA", lines)[1] - 1)]
clean <- function(x) { v <- strsplit(gsub("\"", "", x), ",")[[1]]; v[v != ""] }
for (k in c("#Study name", "#iMotions version", "#Recording time")) cat("  ", paste(clean(info[startsWith(info, k)][1]), collapse = " | "), "\n")
aff <- clean(info[grepl("^#Sensor info,Affectiva", info)])
cat("  ", paste(aff, collapse = " | "), "\n")
fea <- clean(info[grepl("FEA thresholding", info)])
cat("  FEA thresholding settings (first ten):", paste(head(fea[-(1:3)], 10), collapse = "; "), "\n")
thr <- fea[grepl("_Threshold=", fea)]
cat("  Thresholds:", length(thr), "settings;", sum(grepl("=50$", thr)), "of them equal 50\n\n")

# STEP 2. Read the data and the codebook -------------------------------------------
cat("STEP 2. Codebook for the AFFDEX columns\n")
data_line <- grep("^#DATA", lines)[1] + 1
meta <- read.csv(text = lines[(grep("^#METADATA", lines)[1] + 1):(data_line - 3)], header = FALSE,
                 check.names = FALSE, stringsAsFactors = FALSE, na.strings = "")
rownames(meta) <- meta[, 1]
dat <- read.csv(f, skip = data_line - 1, check.names = FALSE, stringsAsFactors = FALSE,
                na.strings = c("", "NA"), fileEncoding = "UTF-8-BOM")
device <- c(NA, unname(unlist(meta["#Device", -1])))[seq_len(ncol(dat))]   # the Row column has no metadata
unit   <- c(NA, unname(unlist(meta["#Unit", -1])))[seq_len(ncol(dat))]
names(dat) <- make.unique(names(dat))          # expression names appear twice: raw score and thresholded copy
codebook <- data.frame(column = names(dat), device = device, unit = unit, stringsAsFactors = FALSE)
write.csv(codebook, "pilot_fea_codebook.csv", row.names = FALSE)
print(codebook[codebook$column %in% c("Joy", "Brow Furrow", "Joy.1", "Brow Furrow.1", "Attention", "Pitch", "Interocular Distance"), ], row.names = FALSE)
cat("\n")

# STEP 3. Filter by device ------------------------------------------------------------
cat("STEP 3. Rows per device\n")
es <- which(startsWith(names(dat), "EventSource"))
for (j in es) cat(sprintf("  %-40s %6d rows\n", device[j], sum(dat[[j]] == 1, na.rm = TRUE)))
af <- dat[!is.na(dat[[es[device[es] == "Affectiva AFFDEX"]]]) & dat[[es[device[es] == "Affectiva AFFDEX"]]] == 1, ]
cat("\n")

# STEP 4. Stimuli -----------------------------------------------------------------------
ev <- dat[!is.na(dat$SlideEvent) & dat$SlideEvent %in% c("StartMedia", "EndMedia"), ]
stim <- data.frame(stimulus = ev$SourceStimuliName[ev$SlideEvent == "StartMedia"],
                   start_ms = ev$Timestamp[ev$SlideEvent == "StartMedia"],
                   end_ms = ev$Timestamp[ev$SlideEvent == "EndMedia"], stringsAsFactors = FALSE)
cat("STEP 4. Stimuli (StartMedia to EndMedia)\n")
print(data.frame(stimulus = stim$stimulus, start_s = round(stim$start_ms / 1000, 2), end_s = round(stim$end_ms / 1000, 2),
                 duration_s = round((stim$end_ms - stim$start_ms) / 1000, 2)), row.names = FALSE)
cat("\n")
in_stim <- function(t, i) t >= stim$start_ms[i] & t <= stim$end_ms[i]

# STEP 5. Frame rate and face found ---------------------------------------------------------
cat("STEP 5. Frame rate and face found per stimulus\n")
gap <- diff(af$Timestamp)
cat(sprintf("  Frames per second over the recording: %.1f (%d frames in %.1f s)\n", nrow(af) / (diff(range(af$Timestamp)) / 1000), nrow(af), diff(range(af$Timestamp)) / 1000))
cat(sprintf("  Gaps between frames are uneven: %.0f%% are under 20 ms and %.0f%% are 40 to 60 ms (median %.1f ms),\n  so compute the rate from frames and time, not from the median gap.\n",
            100 * mean(gap < 20), 100 * mean(gap >= 40 & gap < 60), median(gap)))
face <- !is.na(af$Joy)                          # an empty score means no face was found in that frame
cat(sprintf("  Frames with a face over the whole recording: %d of %d (%.1f%%)\n", sum(face), length(face), 100 * mean(face)))
q <- do.call(rbind, lapply(seq_len(nrow(stim)), function(i) {
  w <- in_stim(af$Timestamp, i)
  data.frame(stimulus = stim$stimulus[i], frames = sum(w), fps = round(sum(w) / ((stim$end_ms[i] - stim$start_ms[i]) / 1000), 1), face_pct = round(100 * mean(face[w]), 1),
             attention_median = round(median(af$Attention[w], na.rm = TRUE), 1),
             interocular_median_px = round(median(af$`Interocular Distance`[w], na.rm = TRUE), 1), stringsAsFactors = FALSE)
}))
print(q, row.names = FALSE)
cat("\n")

# STEP 6. Time above threshold per stimulus ---------------------------------------------------
cat("STEP 6. Percent of frames at or above the threshold (thresholded columns)\n")
expr <- c("Joy", "Brow Furrow", "Anger", "Disgust", "Surprise", "Speaking")
th_col <- function(e) if (paste0(e, ".1") %in% names(dat)) paste0(e, ".1") else e   # the FEA thresholding copy (0 or 1)
tp <- do.call(rbind, lapply(seq_len(nrow(stim)), function(i) {
  w <- in_stim(af$Timestamp, i)
  v <- sapply(expr, function(e) round(100 * mean(af[w, th_col(e)] == 1, na.rm = TRUE), 2))
  data.frame(stimulus = stim$stimulus[i], t(v), check.names = FALSE, stringsAsFactors = FALSE)
}))
print(tp, row.names = FALSE)
fm_lines <- readLines(file("FEAMetrics.csv", encoding = "UTF-8-BOM"), warn = FALSE)
fm <- read.csv("FEAMetrics.csv", skip = grep("^Study Name", fm_lines) - 1, check.names = FALSE,
               stringsAsFactors = FALSE, na.strings = c("NA", ""), fileEncoding = "UTF-8-BOM")
cmp <- data.frame(stimulus = fm$Label, frames_iMotions = fm$`Count Frames`)
for (e in expr) {
  cmp[[paste(e, "iMotions")]] <- fm[[paste(e, "Time Percentage")]]
}
cat("\n  The same columns in FEAMetrics.csv (Count Frames and Time Percentage):\n")
print(cmp, row.names = FALSE)
diffs <- sapply(expr, function(e) max(abs(tp[[e]] - cmp[[paste(e, "iMotions")]][match(tp$stimulus, cmp$stimulus)])))
cat("  Largest difference between the script and FEAMetrics.csv:", round(max(diffs), 2), "percentage points\n\n")
write.csv(merge(q, tp, by = "stimulus", sort = FALSE), "pilot_fea_summary_by_stimulus.csv", row.names = FALSE)

# STEP 7. The raw score behind the threshold ---------------------------------------------
cat("STEP 7. Raw score and threshold for Brow Furrow\n")
bf_raw <- af$`Brow Furrow`; bf_th <- af$`Brow Furrow.1`
agree <- mean((bf_raw >= 50) == (bf_th == 1), na.rm = TRUE)
cat(sprintf("  Frames where 'raw score >= 50' and the thresholded column agree: %.1f%%\n", 100 * agree))
for (i in which(grepl("^0[2-4]", stim$stimulus))) {
  w <- in_stim(af$Timestamp, i)
  cat(sprintf("  %-38s median %5.1f  90th percentile %5.1f  max %5.1f\n", stim$stimulus[i],
              median(bf_raw[w], na.rm = TRUE), quantile(bf_raw[w], 0.9, na.rm = TRUE), max(bf_raw[w], na.rm = TRUE)))
}
cat("\n")

# STEP 8. Episodes: FEAExpressionTable.csv ---------------------------------------------------
cat("STEP 8. Brow Furrow episodes (FEAExpressionTable.csv)\n")
et_lines <- readLines(file("FEAExpressionTable.csv", encoding = "UTF-8-BOM"), warn = FALSE)
xt <- read.csv("FEAExpressionTable.csv", skip = grep("^Study Name", et_lines) - 1, check.names = FALSE,
               stringsAsFactors = FALSE, na.strings = c("NA", ""), fileEncoding = "UTF-8-BOM")
bf <- xt[xt$`Expression Type` == "Brow Furrow", c("Stimulus Label", "Expression Onset", "Expression Duration", "Truncation Type")]
bf$onset_s <- round(bf$`Expression Onset` / 1000, 2)
print(bf[, c("Stimulus Label", "onset_s", "Expression Duration", "Truncation Type")], row.names = FALSE)
cat("  Episodes of every expression in the file:", nrow(xt), "\n")
cat("  Segments (rows of the Expression table) per stimulus, Joy and Brow Furrow:\n")
seg <- table(factor(xt$`Stimulus Label`, levels = stim$stimulus), xt$`Expression Type`)[, c("Joy", "Brow Furrow")]
print(seg)
cat("\n")

# STEP 9. Try it: the same raw scores at other thresholds ------------------------------------
cat("STEP 9. Brow Furrow at thresholds of 25, 50, and 75 (from the raw score)\n")
sens <- do.call(rbind, lapply(which(grepl("^0[26]", stim$stimulus)), function(i) {
  x <- bf_raw[in_stim(af$Timestamp, i)]
  do.call(rbind, lapply(c(25, 50, 75), function(k) {
    r <- rle(x >= k)
    data.frame(stimulus = stim$stimulus[i], threshold = k, time_pct = round(100 * mean(x >= k, na.rm = TRUE), 2),
               stretches = sum(r$values, na.rm = TRUE), stringsAsFactors = FALSE)
  }))
}))
print(sens, row.names = FALSE)
write.csv(sens, "pilot_fea_threshold_sensitivity.csv", row.names = FALSE)
cat("\n")

# STEP 10. Figure -----------------------------------------------------------------------------
cat("STEP 10. Figure written to pilot_fea_figure_R.png\n")
png("pilot_fea_figure_R.png", width = 1600, height = 900, res = 150)
par(mfrow = c(2, 1), mar = c(3.6, 4.2, 2.6, 1), mgp = c(2.4, 0.7, 0))
tt <- af$Timestamp / 1000
for (e in c("Brow Furrow", "Joy")) {
  plot(range(tt), c(0, 100), type = "n", axes = FALSE, xlab = "Time from recording start (s)", ylab = paste(e, "score"))
  axis(1, at = seq(0, 250, 50)); axis(2, at = seq(0, 100, 25), las = 1); box()
  for (i in seq_len(nrow(stim))) rect(stim$start_ms[i] / 1000, 0, stim$end_ms[i] / 1000, 100,
                                      col = if (grepl("^0[26]", stim$stimulus[i])) "#F3E3E5" else NA, border = "grey80")
  lines(tt, af[[e]], col = if (e == "Joy") "#2F6690" else "#B03A48", lwd = 0.8)
  abline(h = 50, lty = 2, col = "grey30")
  text((stim$start_ms + stim$end_ms) / 2000, 93, substr(stim$stimulus, 1, 2), cex = 0.75)
  title(main = paste0(e, " score frame by frame; dashed line = threshold 50; shaded = the two creatives"), cex.main = 0.9, line = 0.8)
}
dev.off()
cat("Done.\n")
