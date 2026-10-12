# =============================================================================
# analysis_full.R  (SIMULATED DATA, teaching only)
# Workshop 8: analyzing and reporting results. The full analysis of the SIMULATED Multimodal Ad Study
# Martin's Family Clothing Consumer Behavior Research Laboratory, JSU
#
# SIMULATED DATA. This script reads the analysis ready files that Workshop 7 built from a SIMULATED
# export bundle (make_simulated_data.py, seed 20271231). No person was recorded, no sensor was switched
# on, and no file from a lab study was used. Participants carry codes P01 to P24 only. The brands
# ALTAVO (Ad_A, a hiking ad) and BRENLOW (Ad_B, a cycling ad) are invented. The numbers must never be
# quoted as findings.
#
# CONTINUITY RULE. Workshop 8 continues the study of Workshop 7. H1 to H3 are run again here with the
# same data and the same code, and STEP 4 STOPS the script if any printed line or any stored value
# differs from Workshop 7 (analysis_output.txt and SIMULATED_test_results.csv of Workshop 7). Workshop 8
# adds analyses; it never changes the Workshop 7 results.
#
# Analysis plan (STEP 2 prints it as a table and writes SIMULATED_analysis_plan.csv). All tests two
# tailed, alpha = .05.
#   Planned in Workshop 7:  H1 Logo fixation and brand recall; H2 Joy and liking; H3 GSR peaks per
#                           minute, each ad versus Baseline (Holm over the two ads).
#   Planned in Workshop 8:  Q4 attention by ad and AOI (linear mixed model, then Holm corrected paired
#                           comparisons); Q5 facial expression by ad (Joy, Engagement, Brow Furrow; Holm
#                           over three); Q6 head to head comparisons that the client asked for (liking,
#                           recall, arousal; Holm over three); O1 order and position checks of the
#                           counterbalancing (Holm over eight).
#   Exploratory:            E1 a cross modality correlation matrix (Spearman rho with 95% CIs, no p
#                           values).
#   Sensitivity power:      P1 what n = 22 could detect, and what the next study needs.
#
# What the script does, in order:
#   STEP 0.   Packages, paths, helper functions
#   STEP 1.   Read the Workshop 7 files and check every column against the codebook
#   STEP 2.   The analysis plan as a table; the analysis sample
#   STEP 3.   Descriptive statistics for every modality by stimulus (SIMULATED_descriptives.csv)
#   STEP 4.   H1 to H3 reproduced exactly as in Workshop 7 (continuity check; stops on a mismatch)
#   STEP 5.   Q4: attention by ad and AOI (mixed model, marginal means, Holm corrected comparisons)
#   STEP 6.   Q5: facial expression by ad (paired tests, Holm, d_z, Wilcoxon checks)
#   STEP 6B.  Q6: head to head comparisons for the client report (liking, recall, arousal)
#   STEP 7.   Order effects (AB versus BA, first versus second): the counterbalancing check
#   STEP 8.   EXPLORATORY: cross modality correlation matrix (SIMULATED_correlations.csv)
#   STEP 9.   Sensitivity power analysis (SIMULATED_power.csv)
#   STEP 10.  Figures (figures/), including the assumption checks
#   STEP 11.  Summary table of all tests (SIMULATED_all_tests.csv)
#
# How to run: Rscript analysis_full.R (or Source in RStudio with the working directory set to this
# folder). The script looks for the Workshop 7 files in ../../ws7/data/ first and then in ws7_inputs/
# (byte identical copies shipped with this folder; STEP 1 compares the checksums when both exist).
# Output: analysis_output.txt (the full console log), the CSV files, and figures/.
# Packages: readr, dplyr, tidyr, ggplot2, lme4, lmerTest, and grid and tools (part of base R). Nothing
# else is needed; estimated marginal means are computed from the model by hand (no emmeans). If a
# package is missing, run once:
#   install.packages(c("readr", "dplyr", "tidyr", "ggplot2", "lme4", "lmerTest"))
# =============================================================================

# ---- STEP 0. Setup ------------------------------------------------------------------------------
suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(tidyr)
  library(ggplot2)
  library(lme4)
  library(lmerTest)
  library(grid)
})
script_dir <- tryCatch({
  args <- commandArgs(trailingOnly = FALSE)
  f <- sub("^--file=", "", args[grep("^--file=", args)])
  if (length(f) == 1) dirname(normalizePath(f)) else getwd()
}, error = function(e) getwd())
setwd(script_dir)
fig_dir <- file.path(script_dir, "figures")
dir.create(fig_dir, showWarnings = FALSE)
options(width = 160, dplyr.summarise.inform = FALSE)
sink(file.path(script_dir, "analysis_output.txt"), split = TRUE)
cat("SIMULATED DATA: SIMULATED Multimodal Ad Study, full analysis for Workshop 8 (no real person was recorded)\n")
cat("Run on", format(Sys.time(), "%Y-%m-%d %H:%M"), "\n")
cat("R ", R.version$major, ".", R.version$minor, "; readr ", as.character(packageVersion("readr")), "; dplyr ",
    as.character(packageVersion("dplyr")), "; tidyr ", as.character(packageVersion("tidyr")), "; ggplot2 ",
    as.character(packageVersion("ggplot2")), "; lme4 ", as.character(packageVersion("lme4")), "; lmerTest ",
    as.character(packageVersion("lmerTest")), "\n\n", sep = "")

# Okabe and Ito palette (safe for the common forms of color blindness)
col_blue <- "#0072B2"; col_verm <- "#D55E00"; col_orange <- "#E69F00"; col_green <- "#009E73"
col_sky <- "#56B4E9"; col_purple <- "#CC79A7"; col_gray <- "#8C8C8C"; col_dark <- "#222222"
ad_cols <- c(Ad_A = col_blue, Ad_B = col_verm)
ad_lab <- c(Ad_A = "Ad_A\nALTAVO", Ad_B = "Ad_B\nBRENLOW")
theme_ws8 <- function(base_size = 12) {
  theme_minimal(base_size = base_size) +
    theme(panel.grid.minor = element_blank(), panel.grid.major = element_line(color = "#E3E3E3", linewidth = 0.4),
          plot.title = element_text(face = "bold", size = base_size + 1), plot.subtitle = element_text(color = "#555555", size = base_size - 1.5),
          plot.caption = element_text(color = "#777777", size = base_size - 3, hjust = 0), legend.position = "top",
          legend.title = element_blank(), plot.title.position = "plot", plot.caption.position = "plot",
          strip.text = element_text(face = "bold", size = base_size), plot.background = element_rect(fill = "white", color = NA))
}
wrap_txt <- function(x, w) paste(strwrap(x, width = w), collapse = "\n")
pf <- function(p) ifelse(p < 0.0001, sprintf("%.2e", p), sprintf("%.5f", p))         # p for the output file
fp <- function(p) ifelse(p < .001, "p < .001", paste0("p = ", sub("^0", "", sprintf("%.3f", p))))   # p for figures
fv <- function(v) ifelse(abs(v - round(v)) < 1e-9, sprintf("%.0f", v), sprintf("%.1f", v))

# 95% limits of d_z from the NONCENTRAL t distribution (Lakens, 2013): the two noncentrality values
# for which the observed t sits at the 97.5th and the 2.5th percentile, divided by sqrt(n). The search
# interval is t +/- 15 (as in Workshop 7) or wider for a very large t. R's pt() is accurate only for
# noncentrality values up to about 37.6, so a limit beyond that is flagged and the approximation is used.
nct_limits <- function(t_obs, df, conf = 0.95) {
  a <- (1 - conf) / 2
  w <- max(15, abs(t_obs))
  f <- function(ncp, target) suppressWarnings(pt(t_obs, df, ncp)) - target
  c(uniroot(f, c(t_obs - w, t_obs + w), target = 1 - a, tol = 1e-12)$root,
    uniroot(f, c(t_obs - w, t_obs + w), target = a, tol = 1e-12)$root)
}
banner <- function(what) sprintf("# SIMULATED DATA for teaching (seed 20271231). %s Made by analysis_full.R (Workshop 8) from the Workshop 7 analysis ready files. No real person was recorded.", what)
write_sim_csv <- function(df, file, what) {
  writeLines(banner(what), file)
  write_csv(df, file, append = TRUE, col_names = TRUE, na = "NA")
}
all_tests <- list()          # one row per test or estimate, for SIMULATED_all_tests.csv (STEP 11)
add_test <- function(...) all_tests[[length(all_tests) + 1]] <<- tibble(...)

# ---- STEP 1. Read the Workshop 7 files and check them against the codebook -----------------------
# Line 1 of each data file is a "# SIMULATED" banner, so skip = 1. NA marks a value that does not exist
# (see the missing_code column of codebook.csv).
cat("---- STEP 1. Files read and checked against the codebook ----\n")
in_files <- c("analysis_ready_long.csv", "analysis_ready_wide.csv", "codebook.csv", "analysis_output.txt", "SIMULATED_test_results.csv")
cand <- c(normalizePath(file.path(script_dir, "..", "..", "ws7", "data"), mustWork = FALSE), file.path(script_dir, "ws7_inputs"))
has_all <- vapply(cand, function(d) all(file.exists(file.path(d, in_files))), logical(1))
if (!any(has_all)) stop("The Workshop 7 files were not found in ../../ws7/data or ws7_inputs/: ", paste(in_files, collapse = ", "))
in_dir <- cand[has_all][1]
cat(sprintf("Workshop 7 files read from: %s\n", if (in_dir == cand[1]) "../../ws7/data (the Workshop 7 folder)" else "ws7_inputs/ (copies shipped with Workshop 8)"))
if (all(has_all)) {
  md5_a <- unname(tools::md5sum(file.path(cand[1], in_files))); md5_b <- unname(tools::md5sum(file.path(cand[2], in_files)))
  if (!identical(md5_a, md5_b)) stop("ws7_inputs/ is not identical to ../../ws7/data: copy the Workshop 7 files again")
  cat(sprintf("ws7_inputs/ holds byte identical copies of the %d Workshop 7 files (MD5 checksums agree)\n", length(in_files)))
}
inp <- function(f) file.path(in_dir, f)
long <- read_csv(inp("analysis_ready_long.csv"), skip = 1, show_col_types = FALSE)
wide <- read_csv(inp("analysis_ready_wide.csv"), skip = 1, show_col_types = FALSE)
cb   <- read_csv(inp("codebook.csv"), skip = 1, show_col_types = FALSE)
seed_txt <- sub(".*\\(seed ([0-9]+)\\).*", "\\1", readLines(inp("analysis_ready_long.csv"), n = 1))
cat(sprintf("Seed of the simulation: %s; long table %d x %d; wide table %d x %d; codebook %d variables\n", seed_txt, nrow(long), ncol(long),
            nrow(wide), ncol(wide), nrow(cb)))
stopifnot(all(names(long) %in% cb$variable), all(names(wide) %in% cb$variable), !any(duplicated(long[c("participant_id", "stimulus")])),
          nrow(long) == 72, nrow(wide) == 24, n_distinct(long$participant_id) == 24)
# Every numeric column must lie inside the allowed range that the codebook states ("a to b")
range_check <- function(d, tab) {
  bind_rows(lapply(intersect(names(d), cb$variable), function(v) {
    rr <- cb$allowed_range[cb$variable == v][1]
    if (!grepl("^-?[0-9.]+ to -?[0-9.]+$", rr) || !is.numeric(d[[v]])) return(NULL)
    lim <- as.numeric(strsplit(rr, " to ")[[1]])
    x <- d[[v]][!is.na(d[[v]])]
    tibble(table = tab, variable = v, n = length(x), min = if (length(x)) min(x) else NA, max = if (length(x)) max(x) else NA,
           allowed = rr, inside = length(x) == 0 || (min(x) >= lim[1] & max(x) <= lim[2]))
  }))
}
rc <- bind_rows(range_check(long, "long"), range_check(wide, "wide"))
cat(sprintf("Range check: %d numeric columns compared with the allowed range in the codebook; outside the range: %d\n", nrow(rc), sum(!rc$inside)))
if (any(!rc$inside)) { print(as.data.frame(rc %>% filter(!inside))); stop("Values outside the codebook range") }
print(as.data.frame(rc %>% filter(table == "long") %>% select(-table) %>% mutate(across(where(is.double), ~ round(.x, 2)))), row.names = FALSE)
cat("Every column is in the codebook; participant x stimulus is unique; every value lies in its allowed range.\n\n")

# ---- STEP 2. The analysis plan and the analysis sample -------------------------------------------------
cat("---- STEP 2. Analysis plan ----\n")
plan <- tribble(
  ~id, ~family, ~question, ~measures, ~test, ~correction, ~status,
  "H1", "Attention and memory", "Is brand recall more likely when the Logo was fixated?", "Logo fixated (AOI export); brand recall (survey)",
  "Fisher exact test on the pooled 2 x 2 table; each ad; mixed logistic regression", "none (own family)", "planned (Workshop 7)",
  "H2", "Emotion and liking", "Do ads that produce more Joy get liked more?", "Joy, percent of time (FEA export); liking (survey)",
  "Linear mixed model, liking ~ Joy/10 + ad + (1 | participant); correlations for each ad", "none (own family)", "planned (Workshop 7)",
  "H3", "Arousal", "Are GSR peaks per minute higher during each ad than during Baseline?", "GSR Peaks Per Minute (GSR export)",
  "Paired t tests, d_z with 95% CI, Wilcoxon checks", "Holm over 2 ads", "planned (Workshop 7)",
  "Q4", "Attention", "Does gaze dwell differ by AOI, and does the AOI pattern differ between the ads?", "Gaze dwell, percent of the ad: Logo, Product, Headline (AOI export)",
  "Linear mixed model, dwell ~ ad x AOI + (1 | participant): Type III F (Satterthwaite) and likelihood ratio tests; check: repeated measures ANOVA, Greenhouse Geisser", "none (omnibus)", "planned (Workshop 8)",
  "Q4a", "Attention", "Which AOIs differ within each ad?", "Gaze dwell, percent of the ad", "6 paired t tests, d_z with 95% CI, Wilcoxon checks", "Holm over 6", "planned (Workshop 8)",
  "Q4b", "Attention", "Does each AOI hold more gaze in one ad than in the other?", "Gaze dwell, percent of the ad", "3 paired t tests, d_z with 95% CI, Wilcoxon checks", "Holm over 3", "planned (Workshop 8)",
  "Q5", "Emotion", "Do Joy, Engagement, and Brow Furrow differ between Ad_A and Ad_B?", "Percent of time at or above threshold 50 (FEA export)",
  "3 paired t tests, d_z with 95% CI, Wilcoxon checks (values pile up at zero)", "Holm over 3", "planned (Workshop 8)",
  "Q6", "Client comparisons", "Do the ads differ on liking, brand recall, and arousal?", "Liking and recall (survey); GSR Peaks Per Minute (GSR export)",
  "Paired t tests (liking, GSR), exact McNemar test (recall)", "Holm over 3", "planned (Workshop 8)",
  "O1", "Counterbalancing check", "Did order (AB or BA) or position (first or second) change liking or dwell?", "Liking; gaze dwell on Logo, Product, Headline",
  "Linear mixed model, y ~ ad x order + (1 | participant): order effect and position effect", "Holm over 8", "planned check (Workshop 8)",
  "E1", "Triangulation", "How do the measures of different modalities go together across participants?", "Logo dwell, Joy, GSR change, liking, recall, purchase intention, attitude",
  "Spearman rho with 95% CI (Fisher z, Bonett and Wright), 21 pairs, no p values", "none (no tests)", "exploratory",
  "P1", "Planning", "What could n = 22 detect, and what n does the next study need?", "(design values)",
  "power.t.test (paired); power.prop.test and a Fisher exact test simulation for H1", "not applicable", "sensitivity analysis")
print(as.data.frame(plan %>% select(id, family, correction, status)), row.names = FALSE, right = FALSE)
cat("\n")
print(as.data.frame(plan %>% select(id, question)), row.names = FALSE, right = FALSE)
cat("\n")
print(as.data.frame(plan %>% select(id, test)), row.names = FALSE, right = FALSE)
write_sim_csv(plan, "SIMULATED_analysis_plan.csv", "The analysis plan of Workshop 8 (planned, check, exploratory).")
cat("wrote SIMULATED_analysis_plan.csv\n\n")

cat("---- STEP 2 (continued). Analysis sample ----\n")
ana <- long %>% filter(include == 1) %>% mutate(stimulus = factor(stimulus, levels = c("Baseline", "Ad_A", "Ad_B")))
ads <- ana %>% filter(stimulus != "Baseline") %>% mutate(stimulus = droplevels(stimulus))
anw <- wide %>% filter(include == 1)
n_part <- n_distinct(ana$participant_id)
cat(sprintf("Participants: %d of %d (excluded by rule R2: %s)\n", n_part, nrow(wide),
            paste(sprintf("%s, %s", wide$participant_id[wide$include == 0], wide$exclusion_reason[wide$include == 0]), collapse = "; ")))
cat(sprintf("Order: AB %d, BA %d; rows: %d participant x stimulus, %d participant x ad; GSR valid rows: %d (missing: %s)\n\n",
            sum(anw$order == "AB"), sum(anw$order == "BA"), nrow(ana), nrow(ads), sum(ana$gsr_valid == 1),
            paste(paste(ana$participant_id, ana$stimulus)[ana$gsr_valid == 0], collapse = ", ")))

# ---- STEP 3. Descriptive statistics for every modality by stimulus ---------------------------------------
# n, mean, SD, median, first and third quartile (R quantile type 7 = Excel QUARTILE.INC), IQR, min, max,
# and, for measures that pile up at zero, the percent of participants at exactly zero.
cat("---- STEP 3. Descriptive statistics (analysis sample, n = 22) ----\n")
gsr_w <- ana %>% select(participant_id, stimulus, gsr_peaks_per_min) %>% pivot_wider(names_from = stimulus, values_from = gsr_peaks_per_min)
gsr_chg <- bind_rows(tibble(participant_id = gsr_w$participant_id, stimulus = "Ad_A", gsr_change = gsr_w$Ad_A - gsr_w$Baseline),
                     tibble(participant_id = gsr_w$participant_id, stimulus = "Ad_B", gsr_change = gsr_w$Ad_B - gsr_w$Baseline))
dstat <- function(d, var, modality, measure, units, zero = FALSE) {
  d %>% filter(!is.na(.data[[var]])) %>% group_by(stimulus) %>%
    summarize(modality = modality, measure = measure, variable = var, units = units, n = n(), mean = mean(.data[[var]]), sd = sd(.data[[var]]),
              median = median(.data[[var]]), q1 = unname(quantile(.data[[var]], 0.25, type = 7)), q3 = unname(quantile(.data[[var]], 0.75, type = 7)),
              min = min(.data[[var]]), max = max(.data[[var]]), pct_zero = if (zero) 100 * mean(.data[[var]] == 0) else NA_real_, .groups = "drop") %>%
    mutate(iqr = q3 - q1, stimulus = as.character(stimulus))
}
study_lvl <- anw %>% mutate(stimulus = "Study (once, after both ads)")
desc <- bind_rows(
  dstat(ads, "logo_gaze_dwell_pct", "Eye tracking", "Gaze dwell on the Logo", "percent of the ad", TRUE),
  dstat(ads, "product_gaze_dwell_pct", "Eye tracking", "Gaze dwell on the Product", "percent of the ad", TRUE),
  dstat(ads, "headline_gaze_dwell_pct", "Eye tracking", "Gaze dwell on the Headline", "percent of the ad", TRUE),
  dstat(ads, "logo_fixated", "Eye tracking", "Logo fixated at least once", "proportion"),
  dstat(ads, "logo_ttff_ms", "Eye tracking", "Logo time to first fixation (fixated only)", "ms"),
  dstat(ads, "logo_fix_count", "Eye tracking", "Fixations on the Logo", "fixations", TRUE),
  dstat(ana, "joy_pct", "Facial expression", "Joy", "percent of time", TRUE),
  dstat(ana, "engagement_pct", "Facial expression", "Engagement", "percent of time", TRUE),
  dstat(ana, "browfurrow_pct", "Facial expression", "Brow Furrow", "percent of time", TRUE),
  dstat(ana, "gsr_peaks_per_min", "GSR", "GSR Peaks Per Minute", "peaks per minute", TRUE),
  dstat(gsr_chg, "gsr_change", "GSR", "GSR Peaks Per Minute, change from Baseline", "peaks per minute", TRUE),
  dstat(ads, "liking", "Survey", "Liking of the ad", "1 to 7"),
  dstat(ads, "recall_correct", "Survey", "Brand recalled correctly", "proportion"),
  dstat(study_lvl, "attitude", "Survey", "Attitude toward reusable water bottles", "1 to 7"),
  dstat(study_lvl, "purchase_intention", "Survey", "Purchase intention (any reusable bottle)", "1 to 7"),
  dstat(ana, "et_quality_pct", "Data quality", "Eye tracking quality", "percent"),
  dstat(ana, "face_quality_pct", "Data quality", "Face found (AFFDEX quality)", "percent")) %>%
  select(modality, measure, variable, units, stimulus, n, mean, sd, median, q1, q3, iqr, min, max, pct_zero)
print(as.data.frame(desc %>% select(-variable, -units, -modality) %>% mutate(measure = substr(measure, 1, 44), across(where(is.double), ~ round(.x, 2)))),
      row.names = FALSE)
write_sim_csv(desc, "SIMULATED_descriptives.csv", "Descriptive statistics of every modality by stimulus, analysis sample (include = 1).")
cat("wrote SIMULATED_descriptives.csv\n")
cat(sprintf("Purchase intention was asked ONCE, about any reusable water bottle (not about ALTAVO or BRENLOW, not per ad): M = %.2f, SD = %.2f, n = %d.\n",
            mean(anw$purchase_intention), sd(anw$purchase_intention), nrow(anw)))
# Top 2 box: the share of participants who chose one of the two most favorable answers (6 or 7 on a
# 1 to 7 scale), the way client reports usually show Likert items (Project doc 02b, section 5).
t2b <- function(x) c(k = sum(x >= 6), n = length(x), pct = 100 * mean(x >= 6))
t2_a <- t2b(ads$liking[ads$stimulus == "Ad_A"]); t2_b <- t2b(ads$liking[ads$stimulus == "Ad_B"]); t2_pi <- t2b(anw$purchase_intention)
cat(sprintf("Top 2 box (6 or 7 on the 1 to 7 scale): liking Ad_A %d of %d (%.1f%%), liking Ad_B %d of %d (%.1f%%); purchase intention %d of %d (%.1f%%)\n\n",
            t2_a[["k"]], t2_a[["n"]], t2_a[["pct"]], t2_b[["k"]], t2_b[["n"]], t2_b[["pct"]], t2_pi[["k"]], t2_pi[["n"]], t2_pi[["pct"]]))

# ---- STEP 4. H1 to H3 reproduced exactly as in Workshop 7 ------------------------------------------------
# The code below is the code of Workshop 7 analysis.R, STEPS 4 to 6, unchanged. Its printed lines are
# captured and compared, one by one, with Workshop 7's analysis_output.txt, and the stored values with
# Workshop 7's SIMULATED_test_results.csv at full precision. Any difference stops the script.
cat("---- STEP 4. H1 to H3 reproduced from Workshop 7 (continuity check) ----\n")
results <- list()
rep_lines <- capture.output({
  h1_table <- function(d, label) {
    tab <- table(factor(d$logo_fixated, levels = c(0, 1), labels = c("Logo not fixated", "Logo fixated")),
                 factor(d$recall_correct, levels = c(0, 1), labels = c("Not recalled", "Recalled")))
    ft <- fisher.test(tab)
    ct <- suppressWarnings(chisq.test(tab, correct = FALSE))
    cy <- suppressWarnings(chisq.test(tab, correct = TRUE))
    a <- tab[2, 2]; b <- tab[2, 1]; c_ <- tab[1, 2]; d_ <- tab[1, 1]
    or_sample <- (a * d_) / (b * c_)
    phi <- (a * d_ - b * c_) / sqrt((a + b) * (c_ + d_) * (a + c_) * (b + d_))
    cat(sprintf("%s (N = %d)\n", label, sum(tab)))
    print(addmargins(tab))
    cat(sprintf("  Recalled when fixated: %d of %d (%.1f%%); when not fixated: %d of %d (%.1f%%)\n", a, a + b, 100 * a / (a + b), c_, c_ + d_,
                100 * c_ / (c_ + d_)))
    cat(sprintf("  Sample odds ratio = (%d x %d) / (%d x %d) = %.2f; Fisher conditional estimate OR = %.2f, 95%% CI [%.2f, %.2f], p = %s\n",
                a, d_, b, c_, or_sample, ft$estimate, ft$conf.int[1], ft$conf.int[2], pf(ft$p.value)))
    cat(sprintf("  Pearson chi square (no continuity correction) X2(1, N = %d) = %.3f, p = %s; phi = %.3f; Yates corrected X2 = %.3f, p = %s\n",
                sum(tab), ct$statistic, pf(ct$p.value), phi, cy$statistic, pf(cy$p.value)))
    cat(sprintf("  Smallest expected count = %.2f\n\n", min(ct$expected)))
    list(tab = tab, a = a, b = b, c = c_, d = d_, or_sample = or_sample, or = unname(ft$estimate), or_lo = ft$conf.int[1], or_hi = ft$conf.int[2],
         p_fisher = ft$p.value, x2 = unname(ct$statistic), p_x2 = ct$p.value, phi = phi, n = sum(tab), min_exp = min(ct$expected))
  }
  h1 <- h1_table(ads, "Pooled over the two ads (two observations per participant)")
  h1_a <- h1_table(ads %>% filter(stimulus == "Ad_A"), "Ad_A only (one observation per participant)")
  h1_b <- h1_table(ads %>% filter(stimulus == "Ad_B"), "Ad_B only (one observation per participant)")
  g1 <- glmer(recall_correct ~ logo_fixated + stimulus + (1 | participant_id), data = ads, family = binomial,
              control = glmerControl(optimizer = "bobyqa"))
  cf1 <- summary(g1)$coefficients
  or_g <- exp(cf1["logo_fixated", "Estimate"])
  ci_g <- exp(cf1["logo_fixated", "Estimate"] + c(-1, 1) * 1.96 * cf1["logo_fixated", "Std. Error"])
  sd_int <- sqrt(as.numeric(VarCorr(g1)$participant_id))
  cat(sprintf("Mixed logistic model, recall ~ logo_fixated + ad + (1 | participant): OR = %.2f, 95%% Wald CI [%.2f, %.2f], z = %.2f, p = %s;\n",
              or_g, ci_g[1], ci_g[2], cf1["logo_fixated", "z value"], pf(cf1["logo_fixated", "Pr(>|z|)"])))
  cat(sprintf("  SD of the participant intercepts = %.3f (logit scale)%s.\n\n", sd_int,
              if (isSingular(g1)) "; the fit is singular: the data show no extra similarity within participants, so the model reduces to ordinary logistic regression" else ""))
  results$H1 <- tibble(id = "H1", test = "Fisher exact test, pooled 2 x 2 (Logo fixated x brand recalled)", n = h1$n, statistic_name = "X2 (no correction)",
                       statistic = h1$x2, df = 1, p = h1$p_fisher, es_type = "odds ratio (Fisher)", es = h1$or, es_ci_low = h1$or_lo, es_ci_high = h1$or_hi,
                       m1 = h1$c / (h1$c + h1$d), m2 = h1$a / (h1$a + h1$b), note = sprintf("phi = %.3f; chi square p = %.5f; sample OR = %.3f", h1$phi, h1$p_x2, h1$or_sample))
  results$H1_A <- tibble(id = "H1_AD_A", test = "Fisher exact test, Ad_A", n = h1_a$n, statistic_name = "X2 (no correction)", statistic = h1_a$x2, df = 1,
                         p = h1_a$p_fisher, es_type = "odds ratio (Fisher)", es = h1_a$or, es_ci_low = h1_a$or_lo, es_ci_high = h1_a$or_hi,
                         m1 = h1_a$c / (h1_a$c + h1_a$d), m2 = h1_a$a / (h1_a$a + h1_a$b), note = sprintf("phi = %.3f", h1_a$phi))
  results$H1_B <- tibble(id = "H1_AD_B", test = "Fisher exact test, Ad_B", n = h1_b$n, statistic_name = "X2 (no correction)", statistic = h1_b$x2, df = 1,
                         p = h1_b$p_fisher, es_type = "odds ratio (Fisher)", es = h1_b$or, es_ci_low = h1_b$or_lo, es_ci_high = h1_b$or_hi,
                         m1 = h1_b$c / (h1_b$c + h1_b$d), m2 = h1_b$a / (h1_b$a + h1_b$b), note = sprintf("phi = %.3f", h1_b$phi))
  results$H1_GLMM <- tibble(id = "H1_GLMM", test = "Mixed logistic regression (random intercept per participant)", n = nrow(ads), statistic_name = "z",
                            statistic = cf1["logo_fixated", "z value"], df = NA, p = cf1["logo_fixated", "Pr(>|z|)"], es_type = "odds ratio (Wald CI)",
                            es = or_g, es_ci_low = ci_g[1], es_ci_high = ci_g[2], m1 = NA, m2 = NA, note = sprintf("SD of intercepts %.3f; singular %s", sd_int, isSingular(g1)))

  ads2 <- ads %>% mutate(joy10 = joy_pct / 10, ad_b = as.integer(stimulus == "Ad_B"))
  m2 <- lmer(liking ~ joy10 + ad_b + (1 | participant_id), data = ads2, REML = TRUE)
  s2 <- summary(m2)$coefficients
  ci2 <- confint(m2, parm = "joy10", method = "Wald")
  vc <- as.data.frame(VarCorr(m2))
  icc <- vc$vcov[1] / sum(vc$vcov)
  cat(sprintf("Model: liking ~ Joy (per 10 points) + ad + (1 | participant); %d observations, %d participants\n", nrow(ads2), n_part))
  print(round(s2, 4))
  cat(sprintf("Joy slope b = %.4f liking points per 10 percentage points of Joy, SE = %.4f, 95%% CI [%.4f, %.4f] (Wald), t(%.2f) = %.4f, p = %s\n",
              s2["joy10", "Estimate"], s2["joy10", "Std. Error"], ci2[1], ci2[2], s2["joy10", "df"], s2["joy10", "t value"], pf(s2["joy10", "Pr(>|t|)"])))
  cat(sprintf("Ad_B minus Ad_A (at equal Joy): b = %.3f, p = %s\n", s2["ad_b", "Estimate"], pf(s2["ad_b", "Pr(>|t|)"])))
  cat(sprintf("Random intercept SD = %.4f, residual SD = %.4f; share of liking variance between participants (ICC) = %.4f\n",
              sqrt(vc$vcov[1]), sqrt(vc$vcov[2]), icc))
  m2z <- lmer(scale(liking) ~ scale(joy_pct) + ad_b + (1 | participant_id), data = ads2)
  beta_z <- fixef(m2z)[2]
  cat(sprintf("Standardized slope (both variables in SD units) = %.4f\n", beta_z))
  cor_ad <- function(ad) {
    d <- ads2 %>% filter(stimulus == ad)
    ct <- cor.test(d$joy_pct, d$liking)
    sp <- suppressWarnings(cor.test(d$joy_pct, d$liking, method = "spearman", exact = FALSE))
    cat(sprintf("  %s: Pearson r(%d) = %.4f, 95%% CI [%.4f, %.4f], p = %s; Spearman rho = %.4f, p = %s\n", ad, ct$parameter, ct$estimate,
                ct$conf.int[1], ct$conf.int[2], pf(ct$p.value), sp$estimate, pf(sp$p.value)))
    list(r = unname(ct$estimate), lo = ct$conf.int[1], hi = ct$conf.int[2], p = ct$p.value, df = unname(ct$parameter), rho = unname(sp$estimate), p_rho = sp$p.value)
  }
  cat("Correlations for each ad (one rating per participant):\n")
  r_a <- cor_ad("Ad_A"); r_b <- cor_ad("Ad_B")
  cat(sprintf("Percent of time with Joy: Ad_A M = %.2f (SD = %.2f), Ad_B M = %.2f (SD = %.2f); zero Joy in %d of %d participant x ad rows\n\n",
              mean(ads2$joy_pct[ads2$stimulus == "Ad_A"]), sd(ads2$joy_pct[ads2$stimulus == "Ad_A"]), mean(ads2$joy_pct[ads2$stimulus == "Ad_B"]),
              sd(ads2$joy_pct[ads2$stimulus == "Ad_B"]), sum(ads2$joy_pct == 0), nrow(ads2)))
  results$H2 <- tibble(id = "H2", test = "Linear mixed model: liking ~ Joy/10 + ad + (1 | participant), Satterthwaite df", n = nrow(ads2), statistic_name = "t",
                       statistic = s2["joy10", "t value"], df = s2["joy10", "df"], p = s2["joy10", "Pr(>|t|)"], es_type = "b per 10 points of Joy",
                       es = s2["joy10", "Estimate"], es_ci_low = ci2[1], es_ci_high = ci2[2], m1 = NA, m2 = NA,
                       note = sprintf("standardized slope %.3f; SE %.4f; ICC %.3f", beta_z, s2["joy10", "Std. Error"], icc))
  results$H2_A <- tibble(id = "H2_R_AD_A", test = "Pearson correlation, Ad_A", n = r_a$df + 2, statistic_name = "r", statistic = r_a$r, df = r_a$df, p = r_a$p,
                         es_type = "r", es = r_a$r, es_ci_low = r_a$lo, es_ci_high = r_a$hi, m1 = NA, m2 = NA, note = sprintf("Spearman rho %.3f, p %.5f", r_a$rho, r_a$p_rho))
  results$H2_B <- tibble(id = "H2_R_AD_B", test = "Pearson correlation, Ad_B", n = r_b$df + 2, statistic_name = "r", statistic = r_b$r, df = r_b$df, p = r_b$p,
                         es_type = "r", es = r_b$r, es_ci_low = r_b$lo, es_ci_high = r_b$hi, m1 = NA, m2 = NA, note = sprintf("Spearman rho %.3f, p %.5f", r_b$rho, r_b$p_rho))

  gw <- ana %>% select(participant_id, stimulus, gsr_peaks_per_min) %>% pivot_wider(names_from = stimulus, values_from = gsr_peaks_per_min)
  h3_test <- function(ad) {
    d <- gw %>% filter(!is.na(.data[[ad]]), !is.na(Baseline))
    x <- d$Baseline; y <- d[[ad]]; dd <- y - x; n <- length(dd)
    tt <- t.test(y, x, paired = TRUE)
    dz <- mean(dd) / sd(dd)
    lim <- nct_limits(unname(tt$statistic), n - 1) / sqrt(n)
    se_dz <- sqrt(1 / n + dz^2 / (2 * n))
    wx <- suppressWarnings(wilcox.test(y, x, paired = TRUE, exact = FALSE))
    sw <- shapiro.test(dd)
    cat(sprintf("%s versus Baseline (n = %d%s)\n", ad, n, if (n < nrow(gw)) sprintf("; missing GSR: %s", paste(setdiff(gw$participant_id, d$participant_id), collapse = ", ")) else ""))
    cat(sprintf("  Baseline: M = %.4f, SD = %.4f, median = %.3f; %s: M = %.4f, SD = %.4f, median = %.3f\n", mean(x), sd(x), median(x), ad, mean(y), sd(y), median(y)))
    cat(sprintf("  Difference: M = %.4f peaks per minute, SD = %.4f, 95%% CI [%.4f, %.4f]\n", mean(dd), sd(dd), tt$conf.int[1], tt$conf.int[2]))
    cat(sprintf("  Paired t(%d) = %.4f, p = %s; d_z = %.4f, 95%% CI [%.4f, %.4f] (noncentral t), approximation [%.4f, %.4f]\n", n - 1, tt$statistic,
                pf(tt$p.value), dz, lim[1], lim[2], dz - 1.96 * se_dz, dz + 1.96 * se_dz))
    cat(sprintf("  For the Excel check: SE of the mean difference = %.4f; critical t (95%%) = %.4f; SE of d_z (approximation) = %.4f\n",
                sd(dd) / sqrt(n), qt(0.975, n - 1), se_dz))
    cat(sprintf("  Wilcoxon signed rank V = %s, p = %s; Shapiro Wilk on the differences W = %.3f, p = %s; increases: %d of %d\n\n",
                fv(wx$statistic), pf(wx$p.value), sw$statistic, pf(sw$p.value), sum(dd > 0), n))
    list(ad = ad, n = n, m1 = mean(x), m2 = mean(y), sd1 = sd(x), sd2 = sd(y), md = mean(dd), sdd = sd(dd), lo = tt$conf.int[1], hi = tt$conf.int[2],
         t = unname(tt$statistic), df = n - 1, p = tt$p.value, dz = dz, dz_lo = lim[1], dz_hi = lim[2], dz_lo_a = dz - 1.96 * se_dz,
         dz_hi_a = dz + 1.96 * se_dz, V = unname(wx$statistic), p_w = wx$p.value, sw_p = sw$p.value, dd = dd, up = sum(dd > 0))
  }
  h3a <- h3_test("Ad_A"); h3b <- h3_test("Ad_B")
  p_holm <- p.adjust(c(h3a$p, h3b$p), method = "holm")
  cat(sprintf("Holm adjusted p: Ad_A %s, Ad_B %s\n\n", pf(p_holm[1]), pf(p_holm[2])))
  for (h in list(list(h3a, p_holm[1], "H3_AD_A"), list(h3b, p_holm[2], "H3_AD_B"))) {
    r <- h[[1]]
    results[[h[[3]]]] <- tibble(id = h[[3]], test = sprintf("Paired t test, GSR Peaks Per Minute, %s minus Baseline", r$ad), n = r$n, statistic_name = "t",
                                statistic = r$t, df = r$df, p = r$p, es_type = "d_z", es = r$dz, es_ci_low = r$dz_lo, es_ci_high = r$dz_hi, m1 = r$m1, m2 = r$m2,
                                sd1 = r$sd1, sd2 = r$sd2, mean_diff = r$md, sd_diff = r$sdd, diff_ci_low = r$lo, diff_ci_high = r$hi,
                                es_ci_low_approx = r$dz_lo_a, es_ci_high_approx = r$dz_hi_a, p_holm = h[[2]], np_stat = r$V, np_p = r$p_w,
                                note = sprintf("Holm p %.6f; mean diff %.4f (SD %.4f), CI [%.4f, %.4f]; d_z approx CI [%.4f, %.4f]; Wilcoxon V %s, p %.6f",
                                               h[[2]], r$md, r$sdd, r$lo, r$hi, r$dz_lo_a, r$dz_hi_a, fv(r$V), r$p_w))
  }
})
cat(rep_lines, sep = "\n")
# Continuity check 1: every printed line must appear, character for character, in Workshop 7's output
ws7_out <- readLines(inp("analysis_output.txt"))
rep_chk <- rep_lines[nzchar(trimws(rep_lines))]
in_ws7 <- rep_chk %in% ws7_out
# Continuity check 2: every stored value must equal Workshop 7's SIMULATED_test_results.csv at full precision
ws7_tr <- read_csv(inp("SIMULATED_test_results.csv"), show_col_types = FALSE)
tr7 <- bind_rows(results)
num_cols <- c("n", "statistic", "df", "p", "es", "es_ci_low", "es_ci_high", "m1", "m2", "sd1", "sd2", "mean_diff", "sd_diff", "diff_ci_low",
              "diff_ci_high", "es_ci_low_approx", "es_ci_high_approx", "p_holm", "np_stat", "np_p")
cmp <- inner_join(tr7 %>% select(id, all_of(num_cols)) %>% pivot_longer(-id, values_to = "ws8"),
                  ws7_tr %>% select(id, all_of(num_cols)) %>% pivot_longer(-id, values_to = "ws7"), by = c("id", "name")) %>%
  filter(!(is.na(ws8) & is.na(ws7))) %>% mutate(rel_diff = abs(ws8 - ws7) / pmax(1, abs(ws7)))
cat("CONTINUITY CHECK against Workshop 7\n")
cat(sprintf("  Printed lines: %d of %d non-empty lines of the H1 to H3 output appear verbatim in Workshop 7's analysis_output.txt\n", sum(in_ws7), length(rep_chk)))
cat(sprintf("  Stored values: %d values for %d tests (%s) compared with Workshop 7's SIMULATED_test_results.csv; largest relative difference = %.1e; missing values that differ: %d\n",
            nrow(cmp), n_distinct(cmp$id), paste(unique(cmp$id), collapse = ", "), max(cmp$rel_diff, na.rm = TRUE), sum(xor(is.na(cmp$ws8), is.na(cmp$ws7)))))
if (!all(in_ws7) || any(xor(is.na(cmp$ws8), is.na(cmp$ws7))) || max(cmp$rel_diff, na.rm = TRUE) > 1e-9 || nrow(tr7) != nrow(ws7_tr)) {
  cat("  Lines not found in Workshop 7:\n"); cat(paste0("    ", rep_chk[!in_ws7]), sep = "\n")
  sink(); stop("CONTINUITY CHECK FAILED: H1 to H3 differ from Workshop 7")
}
cat("  RESULT: H1 to H3 match Workshop 7 exactly (H1 Fisher OR = 4.34, p = .032; H2 b = 0.32 per 10 points, p = .110; H3 d_z = 0.83 and 0.95).\n\n")
for (r in seq_len(nrow(tr7))) {
  x <- tr7[r, ]
  add_test(family = c(H1 = "H1", H1_AD_A = "H1", H1_AD_B = "H1", H1_GLMM = "H1", H2 = "H2", H2_R_AD_A = "H2", H2_R_AD_B = "H2", H3_AD_A = "H3", H3_AD_B = "H3")[[x$id]],
           id = x$id, test = x$test, status = if (x$id %in% c("H1", "H2", "H3_AD_A", "H3_AD_B")) "planned (Workshop 7)" else "planned check (Workshop 7)",
           n = x$n, statistic_name = x$statistic_name, statistic = x$statistic, df1 = x$df, df2 = NA_real_, p = x$p,
           p_holm = if (is.na(x$p_holm)) NA_real_ else x$p_holm, holm_family = if (grepl("^H3", x$id)) "H3: 2 ads" else NA_character_,
           es_type = x$es_type, es = x$es, es_ci_low = x$es_ci_low, es_ci_high = x$es_ci_high, m1 = x$m1, m2 = x$m2,
           sd1 = if (is.na(x$sd1)) NA_real_ else x$sd1, sd2 = if (is.na(x$sd2)) NA_real_ else x$sd2,
           mean_diff = if (is.na(x$mean_diff)) NA_real_ else x$mean_diff, sd_diff = if (is.na(x$sd_diff)) NA_real_ else x$sd_diff,
           diff_ci_low = if (is.na(x$diff_ci_low)) NA_real_ else x$diff_ci_low, diff_ci_high = if (is.na(x$diff_ci_high)) NA_real_ else x$diff_ci_high,
           es_ci_low_approx = if (is.na(x$es_ci_low_approx)) NA_real_ else x$es_ci_low_approx,
           es_ci_high_approx = if (is.na(x$es_ci_high_approx)) NA_real_ else x$es_ci_high_approx,
           np_stat = if (is.na(x$np_stat)) NA_real_ else x$np_stat, np_p = if (is.na(x$np_p)) NA_real_ else x$np_p, note = x$note)
}

# A paired comparison used from here on: a minus b, paired t test, d_z with noncentral t and approximate
# 95% CIs, the 95% CI of the mean difference, a Wilcoxon signed rank check (zero differences are dropped
# by that test, which matters when many values are 0), and the Shapiro Wilk test on the differences.
paired_cmp <- function(a, b) {
  ok <- !is.na(a) & !is.na(b); a <- a[ok]; b <- b[ok]
  d <- a - b; n <- length(d)
  tt <- t.test(a, b, paired = TRUE)
  dz <- mean(d) / sd(d)
  lim <- nct_limits(unname(tt$statistic), n - 1)
  nct_ok <- all(abs(lim) <= 37.62)
  lim <- lim / sqrt(n)
  se_dz <- sqrt(1 / n + dz^2 / (2 * n))
  wx <- suppressWarnings(wilcox.test(a, b, paired = TRUE, exact = FALSE))
  sw <- tryCatch(shapiro.test(d), error = function(e) list(statistic = NA_real_, p.value = NA_real_))
  list(n = n, m1 = mean(a), sd1 = sd(a), m2 = mean(b), sd2 = sd(b), md = mean(d), sdd = sd(d), lo = tt$conf.int[1], hi = tt$conf.int[2],
       t = unname(tt$statistic), df = n - 1, p = tt$p.value, dz = dz, dz_lo = if (nct_ok) lim[1] else NA_real_, dz_hi = if (nct_ok) lim[2] else NA_real_,
       dz_lo_a = dz - 1.96 * se_dz, dz_hi_a = dz + 1.96 * se_dz, se_dz = se_dz, se = sd(d) / sqrt(n), tcrit = qt(0.975, n - 1), nct_ok = nct_ok,
       V = unname(wx$statistic), p_w = wx$p.value, n_nonzero = sum(d != 0), sw_W = unname(sw$statistic), sw_p = sw$p.value,
       n_pos = sum(d > 0), n_neg = sum(d < 0), n_tie = sum(d == 0), n_zero_both = sum(a == 0 & b == 0), d = d)
}
print_cmp <- function(r, la, lb, units, holm = NULL) {
  cat(sprintf("  %s: M = %.4f, SD = %.4f; %s: M = %.4f, SD = %.4f (n = %d)\n", la, r$m1, r$sd1, lb, r$m2, r$sd2, r$n))
  cat(sprintf("  Difference (%s minus %s): M = %.4f %s, SD = %.4f, 95%% CI [%.4f, %.4f]; SE = %.4f, critical t = %.4f\n", la, lb, r$md, units, r$sdd,
              r$lo, r$hi, r$se, r$tcrit))
  cat(sprintf("  Paired t(%d) = %.4f, p = %s%s; d_z = %.4f, 95%% CI %s, approximation [%.4f, %.4f] (SE of d_z = %.4f)\n", r$df, r$t, pf(r$p),
              if (is.null(holm)) "" else sprintf(", Holm p = %s", pf(holm)), r$dz,
              if (r$nct_ok) sprintf("[%.4f, %.4f] (noncentral t)", r$dz_lo, r$dz_hi) else "not computed (noncentrality beyond the accurate range of pt())",
              r$dz_lo_a, r$dz_hi_a, r$se_dz))
  cat(sprintf("  Wilcoxon signed rank V = %s, p = %s (nonzero differences: %d of %d); Shapiro Wilk W = %s, p = %s; %s higher: %d, %s higher: %d, equal: %d\n",
              fv(r$V), pf(r$p_w), r$n_nonzero, r$n, ifelse(is.na(r$sw_W), "NA", sprintf("%.3f", r$sw_W)), ifelse(is.na(r$sw_p), "NA", pf(r$sw_p)),
              la, r$n_pos, lb, r$n_neg, r$n_tie))
}
store_cmp <- function(r, family, id, test, status, holm_p, holm_family, note = "") {
  add_test(family = family, id = id, test = test, status = status, n = r$n, statistic_name = "t", statistic = r$t, df1 = r$df, df2 = NA_real_,
           p = r$p, p_holm = holm_p, holm_family = holm_family, es_type = "d_z", es = r$dz,
           es_ci_low = if (r$nct_ok) r$dz_lo else r$dz_lo_a, es_ci_high = if (r$nct_ok) r$dz_hi else r$dz_hi_a,
           m1 = r$m1, m2 = r$m2, sd1 = r$sd1, sd2 = r$sd2, mean_diff = r$md, sd_diff = r$sdd, diff_ci_low = r$lo, diff_ci_high = r$hi,
           es_ci_low_approx = r$dz_lo_a, es_ci_high_approx = r$dz_hi_a, np_stat = r$V, np_p = r$p_w,
           note = paste0(if (r$nct_ok) "d_z CI noncentral t" else "d_z CI approximation (noncentrality too large for pt())",
                         sprintf("; Wilcoxon on %d nonzero differences; %d higher, %d lower, %d equal", r$n_nonzero, r$n_pos, r$n_neg, r$n_tie),
                         if (nzchar(note)) paste0("; ", note) else ""))
}

# ---- STEP 5. Q4: attention by ad and AOI ---------------------------------------------------------------------
# Each participant has 6 dwell values (2 ads x 3 AOIs). The mixed model gives each participant a random
# intercept; sum to zero contrasts make the Type III tests meaningful. Likelihood ratio tests compare
# nested models fitted by maximum likelihood (ML). Because the three AOIs share the same 30 s, a
# participant who looks a lot at one AOI looks less at another, so participants may not differ in their
# AVERAGE dwell; the random intercept can then be estimated at zero (a singular fit). The repeated
# measures ANOVA (base R aov with error strata, Greenhouse Geisser corrected) is the check: it tests each
# effect against its own participant x effect error term.
cat("---- STEP 5. Q4: gaze dwell by ad and AOI ----\n")
dw <- ads %>% select(participant_id, order, stimulus, Logo = logo_gaze_dwell_pct, Product = product_gaze_dwell_pct, Headline = headline_gaze_dwell_pct) %>%
  pivot_longer(c(Logo, Product, Headline), names_to = "aoi", values_to = "dwell") %>%
  mutate(ad = factor(as.character(stimulus), levels = c("Ad_A", "Ad_B")), aoi = factor(aoi, levels = c("Logo", "Product", "Headline")),
         participant_id = factor(participant_id))
contrasts(dw$ad) <- contr.sum(2); contrasts(dw$aoi) <- contr.sum(3)
cat(sprintf("Data: %d observations (%d participants x 2 ads x 3 AOIs); missing: %d\n", nrow(dw), n_distinct(dw$participant_id), sum(is.na(dw$dwell))))
cell <- dw %>% group_by(ad, aoi) %>% summarize(n = n(), mean = mean(dwell), sd = sd(dwell), median = median(dwell), pct_zero = 100 * mean(dwell == 0),
                                              lo = mean - qt(0.975, n - 1) * sd / sqrt(n), hi = mean + qt(0.975, n - 1) * sd / sqrt(n), .groups = "drop")
cat("Cell means (percent of the 30 s ad), with 95% t intervals from each cell's own SD:\n")
print(as.data.frame(cell %>% mutate(across(where(is.double), ~ round(.x, 3)))), row.names = FALSE)
lz <- dw %>% filter(aoi == "Logo") %>% group_by(participant_id) %>% summarize(a = dwell[ad == "Ad_A"] == 0, b = dwell[ad == "Ad_B"] == 0, .groups = "drop")
cat(sprintf("Participants with no gaze on the Logo: Ad_A %d of %d, Ad_B %d of %d, in both ads %d\n", sum(lz$a), nrow(lz), sum(lz$b), nrow(lz), sum(lz$a & lz$b)))
tot <- dw %>% group_by(participant_id, ad) %>% summarize(total = sum(dwell), .groups = "drop")
cat(sprintf("Sum of the three AOIs per participant and ad: M = %.2f percent, SD = %.2f, range %.2f to %.2f (the rest of the time the gaze was elsewhere or lost)\n",
            mean(tot$total), sd(tot$total), min(tot$total), max(tot$total)))
m4 <- lmer(dwell ~ ad * aoi + (1 | participant_id), data = dw, REML = TRUE)
a4 <- anova(m4, type = 3, ddf = "Satterthwaite")
vc4 <- as.data.frame(VarCorr(m4))
cat(sprintf("\nModel: dwell ~ ad * AOI + (1 | participant), REML; random intercept SD = %.4f, residual SD = %.4f; singular fit: %s\n",
            vc4$sdcor[1], vc4$sdcor[2], isSingular(m4)))
cat("Type III tests (Satterthwaite degrees of freedom):\n")
a4t <- tibble(term = rownames(a4), F = a4$`F value`, df1 = a4$NumDF, df2 = a4$DenDF, p = a4$`Pr(>F)`) %>%
  mutate(eta_p2_approx = F * df1 / (F * df1 + df2))
for (i in seq_len(nrow(a4t))) cat(sprintf("  %-7s F(%d, %.2f) = %.4f, p = %s; partial eta squared (from F) = %.4f\n", a4t$term[i], a4t$df1[i], a4t$df2[i], a4t$F[i],
                                         pf(a4t$p[i]), a4t$eta_p2_approx[i]))
m4_full <- lmer(dwell ~ ad * aoi + (1 | participant_id), data = dw, REML = FALSE)
m4_add <- lmer(dwell ~ ad + aoi + (1 | participant_id), data = dw, REML = FALSE)
m4_noaoi <- lmer(dwell ~ ad + (1 | participant_id), data = dw, REML = FALSE)
m4_noad <- lmer(dwell ~ aoi + (1 | participant_id), data = dw, REML = FALSE)
lrt <- function(m0, m1, lab) {
  a <- anova(m0, m1, refit = FALSE)
  cat(sprintf("  %-40s chi2(%d) = %.4f, p = %s (ML log likelihoods %.3f and %.3f)\n", lab, a$Df[2], a$Chisq[2], pf(a$`Pr(>Chisq)`[2]), a$logLik[1], a$logLik[2]))
  list(chi = a$Chisq[2], df = a$Df[2], p = a$`Pr(>Chisq)`[2])
}
cat("Likelihood ratio tests (models fitted by ML):\n")
l_int <- lrt(m4_add, m4_full, "ad x AOI (additive vs full model)")
l_aoi <- lrt(m4_noaoi, m4_add, "AOI (ad only vs additive model)")
l_ad <- lrt(m4_noad, m4_add, "ad (AOI only vs additive model)")
# Estimated marginal means: each mean is a row L of the fixed effects design; estimate = L b,
# SE = sqrt(L V L'), Satterthwaite df, from lmerTest::contest1D (no extra package).
grid4 <- expand.grid(ad = levels(dw$ad), aoi = levels(dw$aoi))
X4 <- model.matrix(~ ad * aoi, data = grid4, contrasts.arg = list(ad = contr.sum(2), aoi = contr.sum(3)))
stopifnot(identical(colnames(X4), names(fixef(m4))))
emm_row <- function(L, lab_ad, lab_aoi) {
  ct <- contest1D(m4, L, confint = TRUE, level = 0.95)
  tibble(ad = lab_ad, aoi = lab_aoi, emm = ct$Estimate, se = ct$`Std. Error`, df = ct$df, lower = ct$lower, upper = ct$upper)
}
emm <- bind_rows(lapply(seq_len(nrow(grid4)), function(i) emm_row(X4[i, ], as.character(grid4$ad[i]), as.character(grid4$aoi[i]))),
                 lapply(levels(dw$aoi), function(a) emm_row(colMeans(X4[grid4$aoi == a, ]), "both ads", a)),
                 lapply(levels(dw$ad), function(a) emm_row(colMeans(X4[grid4$ad == a, ]), a, "all AOIs")))
cat("Estimated marginal means from the model (equal to the cell means because every participant has every cell; the SE pools the residual variance):\n")
print(as.data.frame(emm %>% mutate(across(where(is.double), ~ round(.x, 3)))), row.names = FALSE)
write_sim_csv(emm, "SIMULATED_q4_marginal_means.csv", "Q4 estimated marginal means of gaze dwell (percent of the ad) by ad and AOI.")
cat("wrote SIMULATED_q4_marginal_means.csv\n")
# The check: repeated measures ANOVA with error strata, Greenhouse Geisser epsilon computed by hand
aov4 <- aov(dwell ~ ad * aoi + Error(participant_id / (ad * aoi)), data = dw)
s4 <- summary(aov4)
rm_row <- function(stratum, term) {
  tab <- s4[[stratum]][[1]]; rn <- trimws(rownames(tab))
  list(F = tab[rn == term, "F value"], df1 = tab[rn == term, "Df"], df2 = tab[rn == "Residuals", "Df"], ss = tab[rn == term, "Sum Sq"],
       ss_err = tab[rn == "Residuals", "Sum Sq"], p = tab[rn == term, "Pr(>F)"])
}
gg_eps <- function(M) {       # M: participants x k matrix of the k repeated values
  k <- ncol(M); C <- contr.poly(k); S <- cov(M); A <- t(C) %*% S %*% C
  sum(diag(A))^2 / ((k - 1) * sum(diag(A %*% A)))
}
wide_aoi <- dw %>% select(participant_id, ad, aoi, dwell) %>% pivot_wider(names_from = c(ad, aoi), values_from = dwell)
M_aoi <- (as.matrix(wide_aoi[, paste0("Ad_A_", levels(dw$aoi))]) + as.matrix(wide_aoi[, paste0("Ad_B_", levels(dw$aoi))])) / 2
M_int <- as.matrix(wide_aoi[, paste0("Ad_A_", levels(dw$aoi))]) - as.matrix(wide_aoi[, paste0("Ad_B_", levels(dw$aoi))])
rm_ad <- rm_row("Error: participant_id:ad", "ad"); rm_aoi <- rm_row("Error: participant_id:aoi", "aoi"); rm_int <- rm_row("Error: participant_id:ad:aoi", "ad:aoi")
eps_aoi <- gg_eps(M_aoi); eps_int <- gg_eps(M_int)
rm_aoi$p_gg <- stats::pf(rm_aoi$F, eps_aoi * rm_aoi$df1, eps_aoi * rm_aoi$df2, lower.tail = FALSE)
rm_int$p_gg <- stats::pf(rm_int$F, eps_int * rm_int$df1, eps_int * rm_int$df2, lower.tail = FALSE)
cat("Check: repeated measures ANOVA (each effect tested against its own participant x effect error term):\n")
cat(sprintf("  ad      F(%d, %d) = %.4f, p = %s; partial eta squared = %.4f\n", rm_ad$df1, rm_ad$df2, rm_ad$F, pf(rm_ad$p), rm_ad$ss / (rm_ad$ss + rm_ad$ss_err)))
cat(sprintf("  AOI     F(%d, %d) = %.4f, p = %s; Greenhouse Geisser epsilon = %.4f, corrected p = %s; partial eta squared = %.4f\n", rm_aoi$df1, rm_aoi$df2,
            rm_aoi$F, pf(rm_aoi$p), eps_aoi, pf(rm_aoi$p_gg), rm_aoi$ss / (rm_aoi$ss + rm_aoi$ss_err)))
cat(sprintf("  ad:AOI  F(%d, %d) = %.4f, p = %s; Greenhouse Geisser epsilon = %.4f, corrected p = %s; partial eta squared = %.4f\n", rm_int$df1, rm_int$df2,
            rm_int$F, pf(rm_int$p), eps_int, pf(rm_int$p_gg), rm_int$ss / (rm_int$ss + rm_int$ss_err)))
for (i in seq_len(nrow(a4t))) {
  tm <- a4t$term[i]
  l <- list(ad = l_ad, aoi = l_aoi, `ad:aoi` = l_int)[[tm]]
  add_test(family = "Q4", id = paste0("Q4_", c(ad = "AD", aoi = "AOI", `ad:aoi` = "AD_X_AOI")[[tm]]), test = sprintf("Linear mixed model dwell ~ ad x AOI + (1 | participant), Type III F, term %s", tm),
           status = "planned (Workshop 8)", n = nrow(dw), statistic_name = "F", statistic = a4t$F[i], df1 = a4t$df1[i], df2 = a4t$df2[i], p = a4t$p[i],
           p_holm = NA_real_, holm_family = NA_character_, es_type = "partial eta squared (from F, approximation)", es = a4t$eta_p2_approx[i],
           es_ci_low = NA_real_, es_ci_high = NA_real_,
           note = sprintf("LRT chi2(%d) = %.3f, p = %.5f; singular fit %s", l$df, l$chi, l$p, isSingular(m4)))
}
for (z in list(list("AD", "ad", rm_ad, NA_real_, NA_real_), list("AOI", "AOI", rm_aoi, eps_aoi, rm_aoi$p_gg), list("AD_X_AOI", "ad x AOI", rm_int, eps_int, rm_int$p_gg))) {
  r <- z[[3]]
  add_test(family = "Q4", id = paste0("Q4_RM_", z[[1]]), test = sprintf("Check: repeated measures ANOVA, term %s", z[[2]]), status = "planned check (Workshop 8)",
           n = nrow(dw), statistic_name = "F", statistic = r$F, df1 = r$df1, df2 = r$df2, p = r$p, p_holm = NA_real_, holm_family = NA_character_,
           es_type = "partial eta squared", es = r$ss / (r$ss + r$ss_err), es_ci_low = NA_real_, es_ci_high = NA_real_,
           note = if (is.na(z[[4]])) "1 df: no sphericity correction needed" else sprintf("Greenhouse Geisser epsilon %.4f, corrected p %.6f", z[[4]], z[[5]]))
}
# Q4a: pairwise AOI comparisons within each ad (paired t tests on each participant's own difference;
# unlike the model contrasts, they do not assume one common residual variance for all AOIs), Holm over 6.
cat("\nQ4a: pairwise comparisons of the AOIs within each ad (paired t tests, Holm over 6)\n")
dwv <- function(ad, aoi) (dw %>% filter(ad == !!ad, aoi == !!aoi) %>% arrange(participant_id))$dwell
pairs4 <- expand.grid(pair = c("Product - Headline", "Product - Logo", "Headline - Logo"), ad = c("Ad_A", "Ad_B"), stringsAsFactors = FALSE)
q4a <- lapply(seq_len(nrow(pairs4)), function(i) {
  pr <- strsplit(pairs4$pair[i], " - ")[[1]]
  paired_cmp(dwv(pairs4$ad[i], pr[1]), dwv(pairs4$ad[i], pr[2]))
})
q4a_holm <- p.adjust(sapply(q4a, `[[`, "p"), method = "holm")
for (i in seq_along(q4a)) {
  pr <- strsplit(pairs4$pair[i], " - ")[[1]]
  cat(sprintf("%s, %s minus %s\n", pairs4$ad[i], pr[1], pr[2]))
  print_cmp(q4a[[i]], pr[1], pr[2], "percentage points", q4a_holm[i])
  store_cmp(q4a[[i]], "Q4a", sprintf("Q4a_%s_%s_%s", pairs4$ad[i], toupper(pr[1]), toupper(pr[2])),
            sprintf("Paired t test, gaze dwell, %s: %s minus %s", pairs4$ad[i], pr[1], pr[2]), "planned (Workshop 8)", q4a_holm[i], "Q4a: 6 AOI pairs")
}
# Q4b: the same AOI in the two ads, Holm over 3
cat("\nQ4b: each AOI, Ad_A versus Ad_B (paired t tests, Holm over 3)\n")
q4b <- lapply(levels(dw$aoi), function(a) paired_cmp(dwv("Ad_A", a), dwv("Ad_B", a)))
names(q4b) <- levels(dw$aoi)
q4b_holm <- p.adjust(sapply(q4b, `[[`, "p"), method = "holm")
for (a in levels(dw$aoi)) {
  cat(sprintf("%s, Ad_A minus Ad_B\n", a))
  print_cmp(q4b[[a]], "Ad_A", "Ad_B", "percentage points", q4b_holm[[a]])
  store_cmp(q4b[[a]], "Q4b", sprintf("Q4b_%s", toupper(a)), sprintf("Paired t test, gaze dwell on the %s, Ad_A minus Ad_B", a), "planned (Workshop 8)",
            q4b_holm[[a]], "Q4b: 3 AOIs")
}
cat("\n")

# ---- STEP 6. Q5: facial expression by stimulus -----------------------------------------------------------------
# Percent of time at or above threshold 50 for Joy, Engagement, and Brow Furrow, Ad_A versus Ad_B, paired.
# Percent of time piles up at zero (a participant who never reached the threshold), so the Wilcoxon
# signed rank test is the check; it drops the participants with no difference (for example 0 and 0).
cat("---- STEP 6. Q5: facial expression, Ad_A versus Ad_B (paired t tests, Holm over 3) ----\n")
expr <- c(Joy = "joy_pct", Engagement = "engagement_pct", `Brow Furrow` = "browfurrow_pct")
adv <- function(ad, v) (ads %>% filter(stimulus == ad) %>% arrange(participant_id))[[v]]
q5 <- lapply(expr, function(v) paired_cmp(adv("Ad_A", v), adv("Ad_B", v)))
q5_holm <- p.adjust(sapply(q5, `[[`, "p"), method = "holm")
for (e in names(expr)) {
  r <- q5[[e]]
  cat(sprintf("%s (percent of time), Ad_A minus Ad_B; at zero: Ad_A %d of %d, Ad_B %d of %d, both ads %d\n", e, sum(adv("Ad_A", expr[[e]]) == 0), r$n,
              sum(adv("Ad_B", expr[[e]]) == 0), r$n, r$n_zero_both))
  print_cmp(r, "Ad_A", "Ad_B", "percentage points", q5_holm[[e]])
  store_cmp(r, "Q5", sprintf("Q5_%s", toupper(gsub(" ", "", e))), sprintf("Paired t test, %s percent of time, Ad_A minus Ad_B", e), "planned (Workshop 8)",
            q5_holm[[e]], "Q5: 3 expressions", sprintf("zero in both ads: %d", r$n_zero_both))
}
cat("\n")

# ---- STEP 6B. Q6: head to head comparisons for the client report ------------------------------------------------
# The ALTAVO brand team asked how its ad (Ad_A) compares with the BRENLOW ad (Ad_B) on liking, brand
# recall, and arousal. Attention (Q4b) and emotion (Q5) are tested above. Recall is yes or no for each
# ad and person, so the exact McNemar test uses the discordant pairs only. Purchase intention was asked
# once, about the category, so the ads CANNOT be compared on it.
cat("---- STEP 6B. Q6: head to head comparisons for the client (Holm over 3) ----\n")
q6_like <- paired_cmp(adv("Ad_A", "liking"), adv("Ad_B", "liking"))
q6_gsr <- paired_cmp(adv("Ad_A", "gsr_peaks_per_min"), adv("Ad_B", "gsr_peaks_per_min"))
ra <- adv("Ad_A", "recall_correct"); rb <- adv("Ad_B", "recall_correct")
b_ <- sum(ra == 1 & rb == 0); c_ <- sum(ra == 0 & rb == 1); nr <- length(ra)
mcn <- binom.test(b_, b_ + c_, 0.5)
rd <- (b_ - c_) / nr
rd_se <- sqrt((b_ + c_) - (b_ - c_)^2 / nr) / nr
q6_p <- c(liking = q6_like$p, recall = mcn$p.value, gsr = q6_gsr$p)
q6_holm <- p.adjust(q6_p, method = "holm")
cat("Liking (1 to 7), Ad_A minus Ad_B\n"); print_cmp(q6_like, "Ad_A", "Ad_B", "scale points", q6_holm[["liking"]])
cat(sprintf("Brand recall: Ad_A %d of %d (%.1f%%), Ad_B %d of %d (%.1f%%); recalled only after Ad_A: %d, only after Ad_B: %d, both: %d, neither: %d\n",
            sum(ra), nr, 100 * mean(ra), sum(rb), nr, 100 * mean(rb), b_, c_, sum(ra == 1 & rb == 1), sum(ra == 0 & rb == 0)))
cat(sprintf("  Exact McNemar test (binomial test on the %d discordant pairs): p = %s, Holm p = %s; difference = %.1f percentage points, 95%% CI [%.1f, %.1f] (Wald)\n",
            b_ + c_, pf(mcn$p.value), pf(q6_holm[["recall"]]), 100 * rd, 100 * (rd - 1.96 * rd_se), 100 * (rd + 1.96 * rd_se)))
cat("GSR Peaks Per Minute, Ad_A minus Ad_B (participants with valid GSR in both ads)\n"); print_cmp(q6_gsr, "Ad_A", "Ad_B", "peaks per minute", q6_holm[["gsr"]])
cat("  (Because Baseline is the same for both ads, this is also the difference between the two changes from Baseline.)\n\n")
store_cmp(q6_like, "Q6", "Q6_LIKING", "Paired t test, liking, Ad_A minus Ad_B", "planned (Workshop 8)", q6_holm[["liking"]], "Q6: 3 client comparisons")
add_test(family = "Q6", id = "Q6_RECALL", test = "Exact McNemar test, brand recalled, Ad_A versus Ad_B", status = "planned (Workshop 8)", n = nr,
         statistic_name = "discordant pairs (A only, B only)", statistic = b_, df1 = c_, df2 = NA_real_, p = mcn$p.value, p_holm = q6_holm[["recall"]],
         holm_family = "Q6: 3 client comparisons", es_type = "difference in proportions (Wald CI)", es = rd, es_ci_low = rd - 1.96 * rd_se,
         es_ci_high = rd + 1.96 * rd_se, m1 = mean(ra), m2 = mean(rb), note = sprintf("df1 column holds the B only count (%d); both %d, neither %d", c_,
                                                                                  sum(ra == 1 & rb == 1), sum(ra == 0 & rb == 0)))
store_cmp(q6_gsr, "Q6", "Q6_GSR", "Paired t test, GSR Peaks Per Minute, Ad_A minus Ad_B", "planned (Workshop 8)", q6_holm[["gsr"]], "Q6: 3 client comparisons")

# ---- STEP 7. Order effects: the counterbalancing check -----------------------------------------------------------
# Half the participants saw Ad_A first (AB), half saw Ad_B first (BA). Model for each outcome:
# y ~ ad * order + (1 | participant) with sum to zero contrasts. Two contrasts, each with its CI:
#   order effect    = mean of the AB group minus mean of the BA group (averaged over the two ads)
#   position effect = an ad seen FIRST minus the same ad seen SECOND (averaged over the two ads)
# With these contrasts, order effect = 2 x the order coefficient and position effect = 2 x the
# interaction coefficient. Holm over the 8 tests (4 outcomes x 2 effects).
cat("---- STEP 7. Order effects (counterbalancing check, Holm over 8) ----\n")
order_out <- c(liking = "Liking (1 to 7)", logo_gaze_dwell_pct = "Gaze dwell on the Logo (%)", product_gaze_dwell_pct = "Gaze dwell on the Product (%)",
               headline_gaze_dwell_pct = "Gaze dwell on the Headline (%)")
ord_res <- list()
for (v in names(order_out)) {
  d <- ads %>% transmute(participant_id, ad = factor(as.character(stimulus), levels = c("Ad_A", "Ad_B")), order = factor(order, levels = c("AB", "BA")), y = .data[[v]])
  contrasts(d$ad) <- contr.sum(2); contrasts(d$order) <- contr.sum(2)
  m <- lmer(y ~ ad * order + (1 | participant_id), data = d, REML = TRUE)
  stopifnot(identical(names(fixef(m)), c("(Intercept)", "ad1", "order1", "ad1:order1")))
  cm <- d %>% group_by(order, ad) %>% summarize(m = mean(y), .groups = "drop")
  cat(sprintf("%s: cell means AB group: Ad_A %.3f, Ad_B %.3f; BA group: Ad_A %.3f, Ad_B %.3f (n = %d and %d); singular fit: %s\n", order_out[[v]],
              cm$m[cm$order == "AB" & cm$ad == "Ad_A"], cm$m[cm$order == "AB" & cm$ad == "Ad_B"], cm$m[cm$order == "BA" & cm$ad == "Ad_A"],
              cm$m[cm$order == "BA" & cm$ad == "Ad_B"], sum(anw$order == "AB"), sum(anw$order == "BA"), isSingular(m)))
  for (eff in c("order", "position")) {
    L <- if (eff == "order") c(0, 0, 2, 0) else c(0, 0, 0, 2)
    ct <- contest1D(m, L, confint = TRUE, level = 0.95)
    ord_res[[paste(v, eff)]] <- tibble(variable = v, label = order_out[[v]], effect = eff, est = ct$Estimate, se = ct$`Std. Error`, df = ct$df,
                                       t = ct$`t value`, p = ct$`Pr(>|t|)`, lo = ct$lower, hi = ct$upper)
  }
}
ord <- bind_rows(ord_res) %>% mutate(p_holm = p.adjust(p, method = "holm"))
for (i in seq_len(nrow(ord))) {
  o <- ord[i, ]
  cat(sprintf("  %-32s %-8s %s = %.4f, 95%% CI [%.4f, %.4f], t(%.2f) = %.4f, p = %s, Holm p = %s\n", o$label, o$effect,
              if (o$effect == "order") "AB minus BA" else "first minus second", o$est, o$lo, o$hi, o$df, o$t, pf(o$p), pf(o$p_holm)))
  add_test(family = "O1", id = sprintf("O1_%s_%s", toupper(sub("_gaze_dwell_pct", "", o$variable)), toupper(o$effect)),
           test = sprintf("Linear mixed model y ~ ad x order + (1 | participant), %s effect, %s", o$effect, o$label), status = "planned check (Workshop 8)",
           n = nrow(ads), statistic_name = "t", statistic = o$t, df1 = o$df, df2 = NA_real_, p = o$p, p_holm = o$p_holm, holm_family = "O1: 8 order checks",
           es_type = if (o$effect == "order") "difference AB minus BA (raw units)" else "difference first minus second (raw units)", es = o$est,
           es_ci_low = o$lo, es_ci_high = o$hi)
}
cat(sprintf("Order checks with Holm p below .05: %d of %d\n\n", sum(ord$p_holm < .05), nrow(ord)))

# ---- STEP 8. EXPLORATORY: cross modality correlation matrix --------------------------------------------------------
# One value per participant (averaged over the two ads where a measure exists for each ad):
#   logo_dwell = mean Logo dwell (%), joy = mean Joy (%), gsr_change = mean of the two ads minus Baseline
#   (peaks per minute; missing for P05, whose Ad_B GSR was invalid), liking = mean liking, recall = number
#   of ads with the correct brand (0, 1, 2), purchase_intention and attitude (asked once).
# Spearman rho (ranks: robust to the zeros and the skew), 95% CI from the Fisher z transformation with
# the Bonett and Wright (2000) variance (1 + rho^2 / 2) / (n - 3). EXPLORATORY: no p values are printed
# on purpose; the matrix generates questions for the next study, it does not test any.
cat("---- STEP 8. EXPLORATORY: cross modality correlation matrix (Spearman, participant level) ----\n")
pl <- ads %>% group_by(participant_id) %>%
  summarize(logo_dwell = mean(logo_gaze_dwell_pct), joy = mean(joy_pct), liking = mean(liking), recall = sum(recall_correct), .groups = "drop") %>%
  left_join(gsr_w %>% transmute(participant_id, gsr_change = (Ad_A + Ad_B) / 2 - Baseline), by = "participant_id") %>%
  left_join(anw %>% select(participant_id, purchase_intention, attitude), by = "participant_id")
cor_vars <- c(logo_dwell = "Logo dwell (%)", joy = "Joy (%)", gsr_change = "GSR change (peaks/min)", liking = "Liking",
              recall = "Recall (0 to 2)", purchase_intention = "Purchase intention", attitude = "Attitude")
write_sim_csv(pl %>% select(participant_id, all_of(names(cor_vars))), "SIMULATED_participant_level.csv",
              "One row per participant (analysis sample): the measures of the exploratory correlation matrix.")
cr <- list()
vv <- names(cor_vars)
for (i in 1:(length(vv) - 1)) for (j in (i + 1):length(vv)) {
  x <- pl[[vv[i]]]; y <- pl[[vv[j]]]; ok <- !is.na(x) & !is.na(y); n <- sum(ok)
  rho <- cor(x[ok], y[ok], method = "spearman")
  se <- sqrt((1 + rho^2 / 2) / (n - 3))
  cr[[length(cr) + 1]] <- tibble(var1 = vv[i], var2 = vv[j], label1 = cor_vars[[vv[i]]], label2 = cor_vars[[vv[j]]], n = n, rho = rho,
                                 lo = tanh(atanh(rho) - qnorm(0.975) * se), hi = tanh(atanh(rho) + qnorm(0.975) * se))
}
cr <- bind_rows(cr) %>% mutate(excludes_zero = lo > 0 | hi < 0)
rmat <- matrix(NA, length(vv), length(vv), dimnames = list(vv, vv)); diag(rmat) <- 1
for (k in seq_len(nrow(cr))) { rmat[cr$var1[k], cr$var2[k]] <- cr$rho[k]; rmat[cr$var2[k], cr$var1[k]] <- cr$rho[k] }
cat("Spearman rho (n = 22; n = 21 for pairs with GSR change):\n")
print(round(rmat, 2))
cat("\nEach pair with its 95% CI:\n")
for (k in seq_len(nrow(cr))) cat(sprintf("  %-24s x %-22s n = %d, rho = %6.3f, 95%% CI [%6.3f, %6.3f]%s\n", cr$label1[k], cr$label2[k], cr$n[k], cr$rho[k],
                                         cr$lo[k], cr$hi[k], if (cr$excludes_zero[k]) "  (interval excludes 0)" else ""))
n_cor <- nrow(cr)
cat(sprintf("\nCorrelations estimated: %d (7 measures, every pair once). Intervals that exclude 0: %d. If no measure were related to any other, about %.2f of %d\n",
            n_cor, sum(cr$excludes_zero), 0.05 * n_cor, n_cor))
cat("95% intervals would still exclude 0 by chance alone. These estimates are EXPLORATORY: they suggest hypotheses for a new study and test none.\n\n")
write_sim_csv(cr, "SIMULATED_correlations.csv", "EXPLORATORY Spearman correlations, participant level, with 95% CIs (Fisher z, Bonett and Wright variance). No p values by design.")
cat("wrote SIMULATED_participant_level.csv and SIMULATED_correlations.csv\n\n")
for (k in seq_len(nrow(cr))) add_test(family = "E1", id = sprintf("E1_%s_%s", toupper(cr$var1[k]), toupper(cr$var2[k])),
                                      test = sprintf("Spearman correlation, %s x %s (participant level)", cr$label1[k], cr$label2[k]), status = "exploratory",
                                      n = cr$n[k], statistic_name = "rho", statistic = cr$rho[k], df1 = NA_real_, df2 = NA_real_, p = NA_real_,
                                      p_holm = NA_real_, holm_family = NA_character_, es_type = "Spearman rho (Fisher z CI)", es = cr$rho[k],
                                      es_ci_low = cr$lo[k], es_ci_high = cr$hi[k], note = "exploratory: no p value by design")

# ---- STEP 9. Sensitivity power analysis ------------------------------------------------------------------------
# A sensitivity analysis asks what effect the design COULD detect; it does not use the observed p value
# ("observed power" adds nothing to the p value). All values: alpha = .05, two tailed, power = .80.
cat("---- STEP 9. Sensitivity power analysis (alpha = .05, two tailed, power = .80) ----\n")
pw22 <- power.t.test(n = 22, sd = 1, sig.level = 0.05, power = 0.80, type = "paired", alternative = "two.sided")
n30 <- power.t.test(delta = 0.30, sd = 1, sig.level = 0.05, power = 0.80, type = "paired")$n
n50 <- power.t.test(delta = 0.50, sd = 1, sig.level = 0.05, power = 0.80, type = "paired")$n
pow30 <- power.t.test(n = 22, delta = 0.30, sd = 1, sig.level = 0.05, type = "paired")$power
pow50 <- power.t.test(n = 22, delta = 0.50, sd = 1, sig.level = 0.05, type = "paired")$power
cat(sprintf("Paired design with n = 22: the smallest d_z detectable with 80%% power = %.4f (base R power.t.test, type = \"paired\")\n", pw22$delta))
cat(sprintf("  Power of the present design: %.4f for d_z = 0.30 and %.4f for d_z = 0.50\n", pow30, pow50))
cat(sprintf("  Participants needed for 80%% power: d_z = 0.30: %.2f, so %d; d_z = 0.50: %.2f, so %d\n", n30, ceiling(n30), n50, ceiling(n50)))
n30h <- power.t.test(delta = 0.30, sd = 1, sig.level = 0.05 / 3, power = 0.80, type = "paired")$n
n50h <- power.t.test(delta = 0.50, sd = 1, sig.level = 0.05 / 3, power = 0.80, type = "paired")$n
cat(sprintf("  At the strictest Holm step of a family of three tests (alpha = .05 / 3 = .0167): d_z = 0.30: %d; d_z = 0.50: %d participants\n",
            ceiling(n30h), ceiling(n50h)))
# H1: the observed effect. Recall when the Logo was NOT fixated = 7 of 20 = .35; the Fisher OR = 4.34
# implies a recall rate when fixated of OR x odds / (1 + OR x odds).
p0 <- h1$c / (h1$c + h1$d); or_obs <- h1$or
p1 <- or_obs * (p0 / (1 - p0)) / (1 + or_obs * (p0 / (1 - p0)))
pp <- power.prop.test(p1 = p0, p2 = p1, sig.level = 0.05, power = 0.80)
cat(sprintf("H1, observed odds ratio %.4f (Fisher), recall when not fixated %.4f, so recall when fixated %.4f:\n", or_obs, p0, p1))
cat(sprintf("  Standard formula (base R power.prop.test, normal approximation, equal groups): %.2f observations per group, so %d per group and %d in total\n",
            pp$n, ceiling(pp$n), 2 * ceiling(pp$n)))
# Simulation with the test actually used (Fisher's exact test) and the observed share of fixated
# observations (24 of 44): for each total N, draw 4,000 studies of N independent observations and count
# how often Fisher's p is below .05. Seed fixed for reproducibility.
set.seed(20271231)
share_fix <- (h1$a + h1$b) / h1$n
sim_power <- function(N, reps = 4000) {
  mean(replicate(reps, {
    fx <- rbinom(N, 1, share_fix); rc <- rbinom(N, 1, ifelse(fx == 1, p1, p0))
    if (length(unique(fx)) < 2 || length(unique(rc)) < 2) 1 else fisher.test(table(fx, rc))$p.value
  }) < 0.05)
}
Ns <- seq(40, 120, by = 4)
sp <- tibble(N = Ns, power = vapply(Ns, sim_power, numeric(1)))
N_sim <- sp$N[which(sp$power >= 0.80)[1]]
cat(sprintf("  Simulation (Fisher's exact test, %.1f%% of observations with the Logo fixated, 4,000 studies per N, N = 40 to 120 in steps of 4):\n", 100 * share_fix))
cat(sprintf("    power at N = 44 (the pooled H1 table) = %.4f; smallest N with power of at least .80 = %d (power %.4f)\n", sp$power[sp$N == 44], N_sim,
            sp$power[sp$N == N_sim]))
cat("    N counts INDEPENDENT observations; with two ads per person, the observations of one person may be related, so the number of people\n")
cat("    should be planned from N (one ad per person) or with a mixed model simulation, not by halving N.\n")
r_h2 <- unname(beta_z)
n_r <- ((qnorm(0.975) + qnorm(0.80)) / atanh(r_h2))^2 + 3
cat(sprintf("H2, for orientation: a correlation the size of the standardized Joy slope (%.4f) needs about %d participants (Fisher z formula, one rating per person)\n\n",
            r_h2, ceiling(n_r)))
pwr <- tibble(
  item = c("Smallest d_z detectable, n = 22 paired", "Power of n = 22 for d_z = 0.30", "Power of n = 22 for d_z = 0.50", "n for d_z = 0.30 (paired)",
           "n for d_z = 0.50 (paired)", "n for d_z = 0.30, alpha .05/3", "n for d_z = 0.50, alpha .05/3", "H1: recall rate when Logo not fixated", "H1: recall rate when fixated implied by OR", "H1: n per group, power.prop.test",
           "H1: total N, power.prop.test (equal groups)", "H1: simulated power at N = 44", "H1: smallest N with simulated power >= .80",
           "H1: simulated power at that N",
           "H2: n for r equal to the standardized slope (Fisher z)"),
  value = c(pw22$delta, pow30, pow50, ceiling(n30), ceiling(n50), ceiling(n30h), ceiling(n50h), p0, p1, ceiling(pp$n), 2 * ceiling(pp$n), sp$power[sp$N == 44], N_sim, sp$power[sp$N == N_sim], ceiling(n_r)),
  exact = c(pw22$delta, pow30, pow50, n30, n50, n30h, n50h, p0, p1, pp$n, 2 * pp$n, sp$power[sp$N == 44], N_sim, sp$power[sp$N == N_sim], n_r),
  method = c(rep("power.t.test, type = paired", 7), "observed (7 of 20)", "OR x odds / (1 + OR x odds)", "power.prop.test", "power.prop.test",
             "simulation, Fisher exact test, 4,000 studies", "simulation, Fisher exact test, 4,000 studies per N",
             "simulation, Fisher exact test, 4,000 studies", "((z.975 + z.80) / atanh(r))^2 + 3"))
write_sim_csv(pwr, "SIMULATED_power.csv", "Sensitivity power analysis (alpha .05, two tailed, power .80).")
write_sim_csv(sp, "SIMULATED_power_h1_simulation.csv", "Simulated power of Fisher's exact test for the H1 odds ratio by total N (4,000 studies per N).")
cat("wrote SIMULATED_power.csv and SIMULATED_power_h1_simulation.csv\n\n")

# ---- STEP 10. Figures --------------------------------------------------------------------------------------------
# Publication quality: PNG, 1400 px wide, 150 dpi, Okabe and Ito colors, every figure titled SIMULATED,
# the raw data (one point per participant, lines join the same person) shown with the means and their
# 95% CIs; no bar of means with an error whisker ("dynamite plot").
cat("---- STEP 10. Figures ----\n")
set.seed(1)
jit <- tibble(participant_id = sort(unique(as.character(ads$participant_id))), jx = runif(n_part, -0.09, 0.09))
mean_ci <- function(x) { x <- x[!is.na(x)]; n <- length(x); m <- mean(x); h <- qt(0.975, n - 1) * sd(x) / sqrt(n); c(m = m, lo = m - h, hi = m + h, n = n) }
paired_panel <- function(d, facet_var, ylab, sub_by_facet, title, ncol = 3) {
  d <- d %>% left_join(jit, by = "participant_id") %>% mutate(x = as.numeric(factor(ad, levels = c("Ad_A", "Ad_B"))) + jx)
  s <- d %>% group_by(.data[[facet_var]], ad) %>% summarize(m = mean_ci(val)[["m"]], lo = mean_ci(val)[["lo"]], hi = mean_ci(val)[["hi"]], .groups = "drop") %>%
    mutate(xs = as.numeric(factor(ad, levels = c("Ad_A", "Ad_B"))) + ifelse(ad == "Ad_A", -0.28, 0.28))
  lab <- tibble(!!facet_var := names(sub_by_facet), txt = unname(sub_by_facet))
  if (is.factor(d[[facet_var]])) lab[[facet_var]] <- factor(lab[[facet_var]], levels = levels(d[[facet_var]]))
  ggplot(d, aes(x, val)) +
    geom_line(aes(group = participant_id), color = "#C8C8C8", linewidth = 0.35) +
    geom_point(aes(color = ad), size = 1.9, alpha = 0.8, show.legend = FALSE) +
    geom_errorbar(data = s, aes(x = xs, ymin = lo, ymax = hi, color = ad), inherit.aes = FALSE, width = 0.08, linewidth = 0.9, show.legend = FALSE) +
    geom_point(data = s, aes(x = xs, y = m, fill = ad), inherit.aes = FALSE, shape = 23, size = 3.6, color = col_dark, show.legend = FALSE) +
    geom_text(data = lab, aes(x = 1.5, y = Inf, label = txt), inherit.aes = FALSE, vjust = 1.3, size = 3.1, color = "#444444", lineheight = 0.95) +
    facet_wrap(as.formula(paste("~", facet_var)), scales = "free_y", ncol = ncol) +
    scale_x_continuous(breaks = 1:2, labels = ad_lab, limits = c(0.55, 2.45)) +
    scale_y_continuous(expand = expansion(mult = c(0.04, 0.22))) +
    scale_color_manual(values = ad_cols) + scale_fill_manual(values = ad_cols) +
    labs(x = NULL, y = ylab) + theme_ws8(11)
}
save_png <- function(p, file, h, title, foot) {
  png(file.path(fig_dir, file), width = 1400, height = h, res = 150)
  grid.newpage()
  pushViewport(viewport(layout = grid.layout(3, 1, heights = unit(c(30, 1, 46), c("points", "null", "points")))))
  grid.text(title, x = 0.02, hjust = 0, gp = gpar(fontface = "bold", fontsize = 14.5), vp = viewport(layout.pos.row = 1, layout.pos.col = 1))
  print(p, vp = viewport(layout.pos.row = 2, layout.pos.col = 1))
  grid.text(wrap_txt(foot, 158), x = 0.02, hjust = 0, gp = gpar(fontsize = 8.3, col = "#666666"), vp = viewport(layout.pos.row = 3, layout.pos.col = 1))
  invisible(dev.off())
  cat(sprintf("saved figure: figures/%s\n", file))
}
# fig1: dwell by ad and AOI
q4b_txt <- setNames(sprintf("Ad_A minus Ad_B = %.2f points\nd_z = %.2f, Holm %s", sapply(q4b, `[[`, "md"), sapply(q4b, `[[`, "dz"), fp(q4b_holm)), levels(dw$aoi))
f1 <- paired_panel(dw %>% transmute(participant_id = as.character(participant_id), ad, aoi, val = dwell), "aoi", "Gaze dwell (percent of the 30 s ad)", q4b_txt)
save_png(f1, "SIMULATED_fig1_dwell_by_ad_aoi.png", 760, "SIMULATED: gaze dwell by ad and AOI (Q4)",
         sprintf("SIMULATED teaching data (seed %s), n = %d participants; points = participants (lines join the same person), diamonds = means with 95%% CIs. Mixed model: AOI F(%d, %.0f) = %.1f, %s; ad x AOI F(%d, %.0f) = %.2f, %s. Each panel has its own y axis: the Product held about 40 percent of the time, the Logo about 1.5 percent.",
                 seed_txt, n_part, a4t$df1[2], a4t$df2[2], a4t$F[2], fp(a4t$p[2]), a4t$df1[3], a4t$df2[3], a4t$F[3], fp(a4t$p[3])))
# fig2: expressions by ad
q5_txt <- setNames(sprintf("Ad_A minus Ad_B = %.2f points\nd_z = %.2f, Holm %s", sapply(q5, `[[`, "md"), sapply(q5, `[[`, "dz"), fp(q5_holm)), names(expr))
d2 <- bind_rows(lapply(names(expr), function(e) ads %>% transmute(participant_id = as.character(participant_id), ad = as.character(stimulus),
                                                                expression = e, val = .data[[expr[[e]]]]))) %>% mutate(expression = factor(expression, levels = names(expr)))
f2 <- paired_panel(d2, "expression", "Percent of time at or above threshold 50", q5_txt)
save_png(f2, "SIMULATED_fig2_expressions_by_ad.png", 760, "SIMULATED: facial expressions during each ad (Q5)",
         sprintf("SIMULATED teaching data (seed %s), n = %d; AFFDEX, threshold 50. Points = participants, lines join the same person, diamonds = means with 95%% CIs. Many values sit at 0 (Brow Furrow: %d of %d participants in Ad_A, %d in Ad_B), so the Wilcoxon test is the check. Holm over the three expressions.",
                 seed_txt, n_part, sum(adv("Ad_A", "browfurrow_pct") == 0), n_part, sum(adv("Ad_B", "browfurrow_pct") == 0)))
# fig3: GSR change from Baseline
d3 <- gsr_chg %>% filter(!is.na(gsr_change)) %>% transmute(participant_id, ad = stimulus, val = gsr_change, panel = "GSR peaks per minute, ad minus Baseline")
f3 <- paired_panel(d3, "panel", "Change from Baseline (peaks per minute)",
                   c(`GSR peaks per minute, ad minus Baseline` = sprintf("H3: Ad_A d_z = %.2f, Ad_B d_z = %.2f (both Holm %s)\nQ6: Ad_A minus Ad_B = %.2f, Holm %s",
                                                                          h3a$dz, h3b$dz, fp(max(p_holm)), q6_gsr$md, fp(q6_holm[["gsr"]]))), ncol = 1) +
  geom_hline(yintercept = 0, linetype = "dashed", color = col_gray) + theme(strip.text = element_blank())
save_png(f3, "SIMULATED_fig3_gsr_change.png", 900, "SIMULATED: GSR peaks per minute, change from Baseline (H3 and Q6)",
         sprintf("SIMULATED teaching data (seed %s). Points = participants (n = %d for Ad_A, %d for Ad_B; P05's Ad_B GSR failed rule R3), lines join the same person, diamonds = means with 95%% CIs; dashed line = no change. Both ads raised the peak rate above Baseline; the two ads did not differ clearly from each other.",
                 seed_txt, h3a$n, h3b$n))
# fig4: correlation matrix (lower triangle)
cm4 <- bind_rows(cr %>% transmute(row = label2, col = label1, rho, lo, hi, n)) %>%
  mutate(row = factor(row, levels = rev(unname(cor_vars)[-1])), col = factor(col, levels = unname(cor_vars)[-length(cor_vars)]),
         txt = sprintf("%.2f\n[%.2f, %.2f]", rho, lo, hi), excl = lo > 0 | hi < 0)
f4 <- ggplot(cm4, aes(col, row, fill = rho)) +
  geom_tile(color = "white", linewidth = 1.2) +
  geom_tile(data = cm4 %>% filter(excl), color = col_dark, linewidth = 0.7, fill = NA) +
  geom_text(aes(label = txt), size = 3.15, lineheight = 0.9, color = "#111111") +
  scale_fill_gradient2(low = col_blue, mid = "white", high = col_verm, midpoint = 0, limits = c(-1, 1), name = "Spearman rho") +
  scale_x_discrete(position = "top") + scale_y_discrete() +
  labs(x = NULL, y = NULL) + theme_ws8(11) +
  theme(panel.grid.major = element_blank(), axis.text.x = element_text(angle = 25, hjust = 0), legend.position = "right", legend.title = element_text(size = 9))
save_png(f4, "SIMULATED_fig4_correlation_matrix.png", 1150, "SIMULATED, EXPLORATORY: how the measures go together across participants (E1)",
         sprintf("SIMULATED teaching data (seed %s), one value per participant (n = 22; n = 21 with GSR change). Cells: Spearman rho and 95%% CI (Fisher z, Bonett and Wright variance); outlined = interval excludes 0 (%d of %d). EXPLORATORY: %d correlations were estimated and about %.1f would exclude 0 by chance alone; no p values, no claims.",
                 seed_txt, sum(cr$excludes_zero), n_cor, n_cor, 0.05 * n_cor))
# fig5: client one page KPI dashboard
jeff <- function(k, n) c(qbeta(0.025, k + 0.5, n - k + 0.5), qbeta(0.975, k + 0.5, n - k + 0.5))
kpi_row <- function(kpi, ad, x) { ci <- mean_ci(x); tibble(kpi = kpi, ad = ad, m = ci[["m"]], lo = ci[["lo"]], hi = ci[["hi"]], n = ci[["n"]]) }
gsr_a <- gsr_chg %>% filter(stimulus == "Ad_A"); gsr_b <- gsr_chg %>% filter(stimulus == "Ad_B")
kp <- bind_rows(kpi_row("att", "Ad_A", adv("Ad_A", "logo_gaze_dwell_pct")), kpi_row("att", "Ad_B", adv("Ad_B", "logo_gaze_dwell_pct")),
                kpi_row("emo", "Ad_A", adv("Ad_A", "joy_pct")), kpi_row("emo", "Ad_B", adv("Ad_B", "joy_pct")),
                kpi_row("aro", "Ad_A", gsr_a$gsr_change), kpi_row("aro", "Ad_B", gsr_b$gsr_change),
                kpi_row("lik", "Ad_A", adv("Ad_A", "liking")), kpi_row("lik", "Ad_B", adv("Ad_B", "liking")),
                tibble(kpi = "rec", ad = c("Ad_A", "Ad_B"), m = 100 * c(mean(ra), mean(rb)), lo = 100 * c(jeff(sum(ra), nr)[1], jeff(sum(rb), nr)[1]),
                       hi = 100 * c(jeff(sum(ra), nr)[2], jeff(sum(rb), nr)[2]), n = nr),
                kpi_row("int", "Both", anw$purchase_intention))
verdict <- function(p_h, md, units, hi_lab = "ALTAVO (Ad_A)", lo_lab = "BRENLOW (Ad_B)") {
  if (p_h < .05) sprintf("%s higher by %s %s (Holm %s)", if (md > 0) hi_lab else lo_lab, sprintf("%.1f", abs(md)), units, fp(p_h))
  else sprintf("We cannot tell which ad is higher\n(difference %.1f %s, Holm %s)", md, units, fp(p_h))
}
tiles <- list(
  att = list("Brand attention", "Gaze dwell on the Logo, % of ad", verdict(q4b_holm[["Logo"]], q4b$Logo$md, "points")),
  emo = list("Emotion", "Joy, % of time (AFFDEX)", verdict(q5_holm[["Joy"]], q5$Joy$md, "points")),
  aro = list("Arousal", "GSR peaks per minute above Baseline", verdict(q6_holm[["gsr"]], q6_gsr$md, "peaks/min")),
  lik = list("Liking", "Survey, 1 to 7", verdict(q6_holm[["liking"]], q6_like$md, "points")),
  rec = list("Brand recall", "Correct brand named, % of viewers", verdict(q6_holm[["recall"]], 100 * rd, "points")),
  int = list("Purchase intention", "Any reusable bottle, 1 to 7 (asked once)", "Asked once, about the category:\nthe ads cannot be compared"))
tile_plot <- function(k) {
  d <- kp %>% filter(kpi == k) %>% mutate(lab = case_when(ad == "Ad_A" ~ "ALTAVO\n(Ad_A)", ad == "Ad_B" ~ "BRENLOW\n(Ad_B)", TRUE ~ "After\nboth ads"),
                                          lab = factor(lab, levels = rev(c("ALTAVO\n(Ad_A)", "BRENLOW\n(Ad_B)", "After\nboth ads"))))
  cols <- c(Ad_A = col_blue, Ad_B = col_verm, Both = col_gray)
  lims <- switch(k, lik = c(1, 7), int = c(1, 7), rec = c(0, 100), c(min(0, d$lo), max(d$hi) * 1.32))
  brks <- switch(k, lik = 1:7, int = 1:7, rec = c(0, 25, 50, 75, 100), waiver())
  ggplot(d, aes(m, lab, color = ad)) +
    geom_vline(xintercept = if (k %in% c("lik", "int")) 1 else 0, color = "#DDDDDD") +
    geom_errorbarh(aes(xmin = lo, xmax = hi), height = 0.18, linewidth = 1.1) +
    geom_point(size = 3.6) +
    geom_text(aes(x = hi, label = sprintf("%.1f", m)), hjust = -0.25, size = 3.6, show.legend = FALSE, color = "#222222") +
    scale_color_manual(values = cols, guide = "none") + scale_x_continuous(limits = lims, breaks = brks) +
    labs(title = tiles[[k]][[1]], subtitle = tiles[[k]][[2]], caption = tiles[[k]][[3]], x = NULL, y = NULL) +
    theme_ws8(10.5) + theme(plot.caption = element_text(size = 9.2, color = if (grepl("cannot", tiles[[k]][[3]])) "#666666" else "#111111",
                                                       face = if (grepl("higher by", tiles[[k]][[3]])) "bold" else "plain", hjust = 0),
                            panel.grid.major.y = element_blank(), plot.background = element_rect(fill = "#FAFAFA", color = "#DDDDDD"))
}
png(file.path(fig_dir, "SIMULATED_fig5_client_dashboard.png"), width = 1400, height = 1050, res = 150)
grid.newpage()
pushViewport(viewport(layout = grid.layout(5, 3, heights = unit(c(0.085, 0.405, 0.02, 0.405, 0.085), "npc"))))
grid.text("SIMULATED client summary: ALTAVO (Ad_A) versus BRENLOW (Ad_B), 22 viewers", x = 0.02, y = 0.62, hjust = 0,
          gp = gpar(fontface = "bold", fontsize = 15), vp = viewport(layout.pos.row = 1, layout.pos.col = 1:3))
grid.text("Dot = average; bar = 95% confidence interval. Each verdict comes from a test on the same viewers, not from whether the bars overlap.",
          x = 0.02, y = 0.18, hjust = 0, gp = gpar(fontsize = 9, col = "#555555"), vp = viewport(layout.pos.row = 1, layout.pos.col = 1:3))
for (i in seq_along(tiles)) {
  rr <- if (i <= 3) 2 else 4; cc <- ((i - 1) %% 3) + 1
  print(tile_plot(names(tiles)[i]), vp = viewport(layout.pos.row = rr, layout.pos.col = cc))
}
grid.text(wrap_txt(sprintf("SIMULATED teaching data (seed %s), not real viewers; ALTAVO and BRENLOW are invented brands. Tests: paired t tests (attention, emotion, arousal, liking) and the exact McNemar test (recall), Holm corrected within each family; recall intervals: Jeffreys. Arousal: averages use every valid recording (22 and 21 viewers), the test the %d viewers with both; both ads raised arousal above Baseline (H3).", seed_txt, q6_gsr$n), 158),
          x = 0.02, hjust = 0, gp = gpar(fontsize = 8.3, col = "#666666"), vp = viewport(layout.pos.row = 5, layout.pos.col = 1:3))
invisible(dev.off())
cat("saved figure: figures/SIMULATED_fig5_client_dashboard.png\n")
# Assumption checks
res4 <- tibble(fitted = fitted(m4), resid = resid(m4), aoi = dw$aoi)
qq_df <- function(x, lab) { x <- sort(x); tibble(theory = qnorm(ppoints(length(x))), sample = (x - mean(x)) / sd(x), panel = lab) }
sd_by_aoi <- res4 %>% group_by(aoi) %>% summarize(s = sd(resid))
pA <- ggplot(res4, aes(fitted, resid, color = aoi, shape = aoi)) + geom_hline(yintercept = 0, linetype = "dashed", color = col_gray) +
  geom_point(position = position_jitter(width = 0.25, height = 0, seed = 2), size = 1.8, alpha = 0.8) + scale_color_manual(values = c(Logo = col_orange, Product = col_green, Headline = col_purple)) +
  labs(title = "A. Q4 model: residuals versus fitted", subtitle = sprintf("Residual SD by AOI: Logo %.2f, Product %.2f, Headline %.2f;\nunequal spread, so the paired tests are the follow ups",
                                                                         sd_by_aoi$s[1], sd_by_aoi$s[2], sd_by_aoi$s[3]), x = "Fitted dwell (%)", y = "Residual") + theme_ws8(10)
pB <- ggplot(qq_df(res4$resid, "B"), aes(theory, sample)) + geom_abline(color = col_gray, linetype = "dashed") + geom_point(color = col_blue, size = 1.6) +
  labs(title = "B. Q4 model: normal Q-Q plot of the residuals", subtitle = "Points near the line = roughly normal residuals", x = "Normal quantiles",
       y = "Standardized residual") + theme_ws8(10)
q5qq <- bind_rows(lapply(names(expr), function(e) qq_df(q5[[e]]$d, e))) %>% mutate(panel = factor(panel, levels = names(expr)))
pC <- ggplot(q5qq, aes(theory, sample)) + geom_abline(color = col_gray, linetype = "dashed") + geom_point(color = col_verm, size = 1.6) +
  facet_wrap(~ panel, nrow = 1) +
  labs(title = "C. Q5: normal Q-Q plots of the differences (Ad_A minus Ad_B)",
       subtitle = sprintf("Shapiro Wilk: Joy %s, Engagement %s, Brow Furrow %s;\nthe flat run = participants with no difference (0 and 0)", fp(q5$Joy$sw_p),
                          fp(q5$Engagement$sw_p), fp(q5$`Brow Furrow`$sw_p)), x = "Normal quantiles", y = "Standardized difference") + theme_ws8(10)
zero_df <- desc %>% filter(!is.na(pct_zero), stimulus %in% c("Ad_A", "Ad_B"), variable %in% c("logo_gaze_dwell_pct", "joy_pct", "engagement_pct", "browfurrow_pct")) %>%
  mutate(measure = factor(c(logo_gaze_dwell_pct = "Logo dwell", joy_pct = "Joy", engagement_pct = "Engagement", browfurrow_pct = "Brow Furrow")[variable],
                          levels = c("Logo dwell", "Joy", "Engagement", "Brow Furrow")))
pD <- ggplot(zero_df, aes(pct_zero, measure, color = stimulus, shape = stimulus)) + geom_point(size = 3) +
  scale_color_manual(values = ad_cols) + scale_x_continuous(limits = c(0, 100)) +
  labs(title = "D. Values at exactly zero", subtitle = "Percent of participants with 0; a mean hides this pile up", x = "Percent of participants at 0", y = NULL) +
  theme_ws8(10)
png(file.path(fig_dir, "SIMULATED_assumption_checks.png"), width = 1400, height = 1250, res = 150)
grid.newpage()
pushViewport(viewport(layout = grid.layout(4, 2, heights = unit(c(0.055, 0.44, 0.44, 0.065), "npc"))))
grid.text("SIMULATED: assumption checks for Q4 and Q5", x = 0.02, hjust = 0, gp = gpar(fontface = "bold", fontsize = 15), vp = viewport(layout.pos.row = 1, layout.pos.col = 1:2))
print(pA, vp = viewport(layout.pos.row = 2, layout.pos.col = 1)); print(pB, vp = viewport(layout.pos.row = 2, layout.pos.col = 2))
print(pC, vp = viewport(layout.pos.row = 3, layout.pos.col = 1)); print(pD, vp = viewport(layout.pos.row = 3, layout.pos.col = 2))
grid.text(wrap_txt(sprintf("SIMULATED teaching data (seed %s), not real participants. The Q4 random intercept was estimated at zero (singular fit) because the three AOIs share the same 30 s; the repeated measures ANOVA and the paired comparisons do not depend on it. Percent of time measures pile up at zero, so every paired t test has a Wilcoxon check. The H2 and H3 checks are in Workshop 7 (figures/SIMULATED_assumption_checks.png there).", seed_txt), 158),
          x = 0.02, hjust = 0, gp = gpar(fontsize = 8.3, col = "#666666"), vp = viewport(layout.pos.row = 4, layout.pos.col = 1:2))
invisible(dev.off())
cat("saved figure: figures/SIMULATED_assumption_checks.png\n\n")

# ---- STEP 11. Summary table of all tests ---------------------------------------------------------------------------
cat("---- STEP 11. Summary of all tests (SIMULATED_all_tests.csv) ----\n")
at <- bind_rows(all_tests) %>% relocate(family, id, status, test, n, statistic_name, statistic, df1, df2, p, p_holm, holm_family, es_type, es, es_ci_low, es_ci_high)
show <- at %>% transmute(id, status = sub(" \\(Workshop ([78])\\)", " WS\\1", status), n, stat = substr(statistic_name, 1, 6), statistic = round(statistic, 3),
                         df1 = round(df1, 2), df2 = round(df2, 2), p = signif(p, 4), p_holm = signif(p_holm, 4), es_type = substr(es_type, 1, 22),
                         es = round(es, 3), ci_low = round(es_ci_low, 3), ci_high = round(es_ci_high, 3))
op <- options(width = 220); print(as.data.frame(show), row.names = FALSE, right = FALSE); options(op)
cat(sprintf("\nRows: %d (planned Workshop 7: %d; planned Workshop 8: %d; planned checks: %d; exploratory: %d). Tests with a p value: %d.\n", nrow(at),
            sum(at$status == "planned (Workshop 7)"), sum(at$status == "planned (Workshop 8)"), sum(grepl("check", at$status)), sum(at$status == "exploratory"),
            sum(!is.na(at$p))))
write_sim_csv(at, "SIMULATED_all_tests.csv", "Every test and estimate of Workshop 8 with its family, status (planned, check, exploratory), Holm p, and effect size with 95% CI.")
cat("wrote SIMULATED_all_tests.csv\n")
cat("References: Bates et al. (2015); Bonett and Wright (2000); Holm (1979); Kuznetsova et al. (2017); Lakens (2013). See results_text.md.\n")
sink()
