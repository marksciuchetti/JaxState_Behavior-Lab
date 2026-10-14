options(width = 150)
# pilot_et_analysis.R
# Workshop 2, section 7.7: practice on the real Aurora pilot file.
# Base R only; no packages to install. Run it from the folder that holds the files:
#   001_PILOT01.csv (Sensor Data export), FixationTable.csv, IndividualAOImetrics.csv
# The respondent is the lab director (PILOT01), who shared his own recording as practice data.
# One respondent: the numbers describe a recording; they test nothing.

f <- "001_PILOT01.csv"
lines <- readLines(file(f, encoding = "UTF-8-BOM"), warn = FALSE)   # UTF-8-BOM drops the byte order mark

# STEP 1. Read the #INFO header ---------------------------------------------
cat("STEP 1. Header\n")
info <- lines[1:(grep("^#METADATA", lines)[1] - 1)]
first_two <- function(x) { v <- strsplit(x, ",")[[1]]; v <- v[v != ""]; paste(v, collapse = " | ") }
for (k in c("#Study name", "#Respondent Name", "#Gaze Calibration", "#iMotions version", "#Recording time"))
  cat("  ", first_two(info[startsWith(info, k)][1]), "\n")
si <- info[startsWith(info, "#Sensor info")]
for (s in si) cat("  ", substr(first_two(s), 1, 110), "\n")
ivt <- si[grepl("I-VT", si)]
ivt_par <- gsub("\"", "", regmatches(ivt, gregexpr("[A-Za-z]+=[^,]+", ivt))[[1]])
cat("  I-VT settings:", paste(ivt_par, collapse = "; "), "\n\n")

# STEP 2. Build the codebook from #METADATA ------------------------------------
cat("STEP 2. Codebook\n")
data_line <- grep("^#DATA", lines)[1] + 1                     # the column header row
meta <- read.csv(text = lines[(grep("^#METADATA", lines)[1] + 1):(data_line - 3)],
                 header = FALSE, check.names = FALSE, stringsAsFactors = FALSE, na.strings = "")
rownames(meta) <- meta[, 1]
dat <- read.csv(f, skip = data_line - 1, check.names = FALSE, stringsAsFactors = FALSE,
                na.strings = c("", "NA"), fileEncoding = "UTF-8-BOM")
device <- c(NA, unname(unlist(meta["#Device", -1])))[seq_len(ncol(dat))]   # the Row column has no metadata
unit   <- c(NA, unname(unlist(meta["#Unit", -1])))[seq_len(ncol(dat))]
names(dat) <- make.unique(names(dat))                          # the EventSource and expression names repeat
codebook <- data.frame(column = names(dat), device = device, unit = unit, stringsAsFactors = FALSE)
write.csv(codebook, "pilot_et_codebook.csv", row.names = FALSE)
show <- c("ET_GazeLeftx", "ET_PupilRight", "ET_DistanceLeft", "Gaze X", "Fixation Duration", "Saccade Amplitude")
print(codebook[codebook$column %in% show, ], row.names = FALSE)
cat("  Columns in the file:", ncol(dat), "\n\n")

# STEP 3. Filter by device --------------------------------------------------------
cat("STEP 3. Rows per device\n")
es_cols <- which(startsWith(names(dat), "EventSource"))
for (j in es_cols) cat(sprintf("  %-40s %6d rows\n", device[j], sum(dat[[j]] == 1, na.rm = TRUE)))
cat(sprintf("  %-40s %6d rows\n\n", "All rows", nrow(dat)))
et_col <- es_cols[device[es_cols] == "Eyetracker Smart Eye Aurora"]
et <- dat[!is.na(dat[[et_col]]) & dat[[et_col]] == 1, ]

# STEP 4. Find the stimuli --------------------------------------------------------
cat("STEP 4. Stimuli (StartMedia to EndMedia)\n")
ev <- dat[!is.na(dat$SlideEvent) & dat$SlideEvent %in% c("StartMedia", "EndMedia"), c("Timestamp", "SlideEvent", "SourceStimuliName")]
stim <- data.frame(stimulus = ev$SourceStimuliName[ev$SlideEvent == "StartMedia"],
                   start_ms = ev$Timestamp[ev$SlideEvent == "StartMedia"],
                   end_ms   = ev$Timestamp[ev$SlideEvent == "EndMedia"], stringsAsFactors = FALSE)
stim$duration_s <- round((stim$end_ms - stim$start_ms) / 1000, 2)
print(transform(stim, start_s = round(start_ms / 1000, 2), end_s = round(end_ms / 1000, 2))[, c("stimulus", "start_s", "end_s", "duration_s")], row.names = FALSE)
cat("\n")

# STEP 5. Sampling rate and gaze data quality -------------------------------------------
cat("STEP 5. Sampling rate and gaze data per stimulus\n")
dt <- diff(et$Timestamp)
cat(sprintf("  Median time between eye tracker rows: %.2f ms (%.1f samples per second)\n", median(dt), 1000 / median(dt)))
valid <- (et$ET_GazeLeftx != -1) | (et$ET_GazeRightx != -1)     # -1 means the eye was not found
cat(sprintf("  Rows with gaze over the whole recording: %.1f%%\n", 100 * mean(valid)))
in_stim <- function(t, s) t >= s$start_ms & t <= s$end_ms
summ <- do.call(rbind, lapply(seq_len(nrow(stim)), function(i) {
  s <- stim[i, ]; w <- in_stim(et$Timestamp, s)
  data.frame(stimulus = s$stimulus, rows = sum(w), gaze_pct = round(100 * mean(valid[w]), 1),
             mean_distance_mm = round(mean(et$ET_DistanceLeft[w], na.rm = TRUE), 0),
             stringsAsFactors = FALSE)
}))
print(summ, row.names = FALSE)
cat("\n")

# STEP 6. Fixations per stimulus (I-VT columns in the Sensor Data file) -------------------
cat("STEP 6. Fixations per stimulus (I-VT filter output)\n")
fix_rows <- et[!is.na(et$`Fixation Index`), ]
fx <- unique(fix_rows[, c("Fixation Index", "Fixation Start", "Fixation Duration")])
fx$stimulus <- NA
for (i in seq_len(nrow(stim))) fx$stimulus[in_stim(fx$`Fixation Start`, stim[i, ])] <- stim$stimulus[i]
ft <- read.csv("FixationTable.csv", skip = grep("^Study Name", readLines(file("FixationTable.csv", encoding = "UTF-8-BOM"), warn = FALSE)) - 1,
               check.names = FALSE, stringsAsFactors = FALSE, na.strings = c("NA", ""), fileEncoding = "UTF-8-BOM")
fs <- do.call(rbind, lapply(stim$stimulus, function(s) {
  a <- fx[fx$stimulus %in% s, ]
  b <- unique(ft[ft$`Stimulus Label` == s, c("Fixation Index", "Fixation Duration", "Truncation Type")])
  data.frame(stimulus = s, starting_inside = nrow(a), FixationTable = nrow(b),
             partial = sum(b$`Truncation Type` == "Partial"),
             mean_fixation_ms = round(mean(b$`Fixation Duration`), 0), stringsAsFactors = FALSE)
}))
print(fs, row.names = FALSE)
summ <- merge(summ, fs, by = "stimulus", sort = FALSE)
write.csv(summ, "pilot_et_summary_by_stimulus.csv", row.names = FALSE)
cat("\n")

# STEP 7. Rebuild the AOI metrics from the FixationTable ---------------------------------
cat("STEP 7. AOI metrics rebuilt from FixationTable.csv\n")
aoi <- ft[!is.na(ft$`AOI Label`) & ft$`AOI Type` == "Static AOI", ]   # the AOIs drawn on the two creatives
rb <- do.call(rbind, lapply(split(aoi, list(aoi$`Stimulus Label`, aoi$`AOI Label`), drop = TRUE), function(a) {
  a <- a[order(a$`Fixation Start`), ]
  data.frame(stimulus = a$`Stimulus Label`[1], aoi = a$`AOI Label`[1],
             ttff_ms = round(a$`Fixation Start`[1] - a$`Stimulus Start`[1], 2),
             fixation_count = nrow(a),
             dwell_ms = round(sum(a$`Fixation Duration`), 2),
             dwell_pct = round(100 * sum(a$`Fixation Duration`) / a$`Stimulus Duration`[1], 2),
             revisits = length(unique(a$`Dwell Index (on AOI)`)) - 1, stringsAsFactors = FALSE)
}))
rownames(rb) <- NULL
im_lines <- readLines(file("IndividualAOImetrics.csv", encoding = "UTF-8-BOM"), warn = FALSE)
im <- read.csv("IndividualAOImetrics.csv", skip = grep("^Study Name", im_lines) - 1, check.names = FALSE,
               stringsAsFactors = FALSE, na.strings = c("NA", ""), fileEncoding = "UTF-8-BOM")
im <- im[im$Type == "Static AOI", ]
key <- function(s, a) paste(s, a)
m <- match(key(rb$stimulus, rb$aoi), key(im$`Parent Label`, im$Label))
rb$iMotions_ttff_ms   <- im$`TTFF (AOI)`[m]
rb$iMotions_fixations <- im$`Fixation count`[m]
rb$iMotions_dwell_ms  <- im$`Dwell time (fixation, ms)`[m]
rb$iMotions_revisits  <- im$`Revisit count (fixation dwells)`[m]
rb$match <- with(rb, abs(ttff_ms - iMotions_ttff_ms) < 0.01 & fixation_count == iMotions_fixations & abs(dwell_ms - iMotions_dwell_ms) < 0.01 & revisits == iMotions_revisits)
print(rb, row.names = FALSE)
write.csv(rb, "pilot_aoi_metrics_rebuilt.csv", row.names = FALSE)
cat("\n")

# STEP 8. Survey pages get AOIs automatically ---------------------------------------------
cat("STEP 8. Automatic AOIs on the questionnaire pages\n")
dyn <- read.csv("IndividualAOImetrics.csv", skip = grep("^Study Name", im_lines) - 1, check.names = FALSE,
                stringsAsFactors = FALSE, na.strings = c("NA", ""), fileEncoding = "UTF-8-BOM")
dyn <- dyn[dyn$Type == "Dynamic AOI", ]
print(table(dyn$`Parent Label`))
cat("  Labels:", paste(unique(dyn$Label), collapse = ", "), "\n\n")

# STEP 9. Figure ---------------------------------------------------------------------------
cat("STEP 9. Figure written to pilot_et_figure_R.png\n")
png("pilot_et_figure_R.png", width = 1600, height = 1150, res = 150)
layout(matrix(c(1, 2, 3, 3), 2, byrow = TRUE), heights = c(1.15, 1))
par(mar = c(3, 3, 2.5, 1), mgp = c(1.8, 0.6, 0))
cols <- c(Face = "#B03A48", Product = "#2F6690")
for (s in stim$stimulus[grepl("^0[26] Creative", stim$stimulus)]) {
  a <- unique(ft[ft$`Stimulus Label` == s, c("Fixation Index", "Fixation X", "Fixation Y", "Fixation Duration")])
  lab2 <- sapply(a$`Fixation Index`, function(k) { v <- ft$`AOI Label`[ft$`Fixation Index` == k & !is.na(ft$`AOI Label`)]; if (length(v)) v[1] else NA })
  plot(NA, xlim = c(0, 2560), ylim = c(1440, 0), xlab = "Screen x (pixels)", ylab = "Screen y (pixels)",
       main = paste0(sub(" \\(15s\\)", "", s), ": ", nrow(a), " fixations"), cex.main = 0.95, asp = 1)
  rect(0, 1440, 2560, 0, border = "grey60")
  lines(a$`Fixation X`, a$`Fixation Y`, col = "grey75")
  symbols(a$`Fixation X`, a$`Fixation Y`, circles = sqrt(a$`Fixation Duration`) * 3, inches = FALSE, add = TRUE,
          bg = ifelse(is.na(lab2), "grey85", adjustcolor(cols[lab2], 0.7)), fg = "grey30")
  legend("topright", c("Face", "Product", "No AOI"), pt.bg = c(cols, "grey85"), pch = 21, cex = 0.8, bty = "n")
}
par(mar = c(4.2, 4.2, 3.2, 1), mgp = c(2.6, 0.7, 0))
tt <- (et$Timestamp) / 1000
plot(range(tt), c(0, 100), type = "n", axes = FALSE, xlab = "Time from recording start (s)",
     ylab = "Rows with gaze (%), 1 s bins")
axis(1, at = seq(0, 250, 50)); axis(2, at = seq(0, 100, 20), las = 1); box()
title(main = "Gaze data over the whole recording", cex.main = 0.95, line = 1.6)
bins <- floor(tt)
pct <- tapply(valid, bins, mean) * 100
for (i in seq_len(nrow(stim))) rect(stim$start_ms[i] / 1000, 0, stim$end_ms[i] / 1000, 100,
                                    col = if (grepl("Creative [AB] -", stim$stimulus[i])) "#F3E3E5" else NA, border = "grey80")
lines(as.numeric(names(pct)), pct, type = "s", col = "#2F2F2F")
text((stim$start_ms + stim$end_ms) / 2000, 8, substr(stim$stimulus, 1, 2), cex = 0.8)
mtext("01 and 05 baseline, 02 Creative A, 06 Creative B (shaded), 03 and 07 recovery, 04 and 08 questionnaire", side = 3, line = 0.1, cex = 0.7)
dev.off()
cat("Done.\n")
