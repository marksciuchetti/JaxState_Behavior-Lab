# =============================================================================
# plan_power.R  (PLANNED STUDY; SIMULATED DATA; teaching only)
# Workshop 9: study design, consent, and the IRB. The planning toolkit for the follow up study
# Martin's Family Clothing Consumer Behavior Research Laboratory, JSU
#
# PLANNED AND SIMULATED. Nothing in this script comes from a real participant. The script plans the
# FOLLOW UP study that Workshop 8 called for ("What the next study needs", ws8/data/results_text.md).
# Every power value is computed from SIMULATED data sets drawn from planning assumptions, and the
# assumptions come from the SIMULATED Multimodal Ad Study of Workshops 7 and 8 (seed 20271231) or are
# stated as judgments. ALTAVO and BRENLOW are invented brands. Participant codes (P01, R01, X01) and
# staff codes (S1 to S6) are placeholders. No number here may be quoted as a finding.
#
# The follow up study (PLANNED): a within subjects lab study in which every participant sees several
# 30 s video ads for each brand, one participant at a time at the single Smart Eye Aurora seat, with
# AFFDEX (respondent camera) and one Shimmer3 GSR+. Liking and purchase intention are asked after each
# ad; brand recall for each ad is asked at the end. The Workshop 7 measures are kept: Logo, Product,
# and Headline AOIs; Joy, Engagement, and Brow Furrow; GSR peaks per minute with a 60 s Baseline;
# liking and recall. Because several ads per brand are shown, ads can be treated as a random factor,
# so that a brand difference can be separated from the particular executions (Judd et al., 2012;
# Westfall et al., 2014).
#
# What the script does, in order:
#   STEP 0.  Setup, the Workshop 6 to 8 inputs, and the planning assumptions (each with its source line)
#   STEP 1.  Analytic power for a paired comparison (power.t.test): d_z = 0.30, 0.50, and the Workshop 8
#            estimate for liking, with and without a Holm family of three
#   STEP 2.  Power by SIMULATION for the key planned test: the brand effect on liking in a linear mixed
#            model with crossed random intercepts for participants and ads, over a grid of participants
#            (24 to 72) and ads per brand (2, 3, 4, 6); then checks at the chosen design (Type I error,
#            the ad SD, a participant by brand slope)
#   STEP 3.  Power by SIMULATION for H1 (Logo fixation and brand recall): a mixed logistic model
#   STEP 4.  The chosen number of ads and participants, with the reasoning, the attrition allowance,
#            and the stopping rule
#   STEP 5.  Counterbalancing: a balanced Latin square (Williams design), the order of the brands, and
#            the assignment of participants to rows (SIMULATED_order_schedule.csv)
#   STEP 6.  The session schedule for one Aurora seat (SIMULATED_session_schedule.csv)
#   STEP 7.  Figures (figures/)
#   STEP 8.  Summary tables (SIMULATED_power_grid.csv and the other CSV files)
#
# How to run: Rscript plan_power.R   (or Source in RStudio with the working directory set to this folder)
#   About 9 minutes on one core (the time is printed at the end of plan_output.txt). On macOS or Linux,
#   N_CORES <- 2L in STEP 0 runs two cells at a time with parallel::mclapply (Windows cannot fork).
#   For a quick teaching run: Rscript plan_power.R quick   (100 simulated studies per cell instead of
#   500; the numbers then differ a little from plan_output.txt, and the comment checks in STEP 4 are
#   skipped).
# Output: plan_output.txt (the full console log), the CSV files listed in STEP 8, figures/.
# Packages: lme4 and lmerTest (as in Workshops 7 and 8), ggplot2 for the figures, and parallel, grid,
#   and tools from base R. Nothing else is needed. If a package is missing, run once:
#   install.packages(c("lme4", "lmerTest", "ggplot2"))
# Inputs: the script reads Workshop 6, 7, and 8 files from ../../ws6/data, ../../ws7/data, and
#   ../../ws8/data when those folders exist, and otherwise from prior_inputs/ (byte identical copies
#   shipped with this folder; STEP 0 compares the MD5 checksums when both exist).
# =============================================================================

# ---- STEP 0. Setup ----------------------------------------------------------------------------------
t_start <- Sys.time()
suppressPackageStartupMessages({
  library(lme4)
  library(lmerTest)
  library(ggplot2)
  library(grid)
  library(parallel)
})
script_dir <- tryCatch({
  args <- commandArgs(trailingOnly = FALSE)
  f <- sub("^--file=", "", args[grep("^--file=", args)])
  if (length(f) == 1) dirname(normalizePath(f)) else getwd()
}, error = function(e) getwd())
setwd(script_dir)
fig_dir <- file.path(script_dir, "figures")
dir.create(fig_dir, showWarnings = FALSE)
options(width = 160, warn = 1)
QUICK <- "quick" %in% commandArgs(trailingOnly = TRUE)
N_SIM_GRID  <- if (QUICK) 100 else 500   # simulated studies per cell of the main grid (STEP 2)
N_SIM_CHECK <- if (QUICK) 100 else 500   # per check at the chosen design (STEP 2, part B)
N_SIM_H1    <- if (QUICK) 60 else 200    # per cell of the H1 simulation (STEP 3; a mixed logistic fit takes about 4 times as long)
SEED <- 20270208                         # fixed seed; every simulated study i of cell k uses seed SEED + 10000 * k + i
N_CORES <- 1L   # On macOS or Linux with several free cores, N_CORES <- 2L runs two cells at a time (parallel::mclapply). The
                # results do not change, because every simulated study has its own seed. Windows always runs on one core.
if (.Platform$OS.type == "windows") N_CORES <- 1L

sink(file.path(script_dir, "plan_output.txt"), split = TRUE)
cat("PLANNED STUDY, SIMULATED DATA: power and planning for the follow up ad study (Workshop 9). No real person was recorded.\n")
cat("Run on", format(Sys.time(), "%Y-%m-%d %H:%M"), "\n")
cat("R ", R.version$major, ".", R.version$minor, "; lme4 ", as.character(packageVersion("lme4")), "; lmerTest ",
    as.character(packageVersion("lmerTest")), "; ggplot2 ", as.character(packageVersion("ggplot2")), "\n", sep = "")
cat(sprintf("Simulated studies per cell: grid %d, checks %d, H1 %d%s; cores used: %d; seed %d\n\n", N_SIM_GRID, N_SIM_CHECK, N_SIM_H1,
            if (QUICK) " (QUICK RUN)" else "", N_CORES, SEED))

# Okabe and Ito palette (safe for the common forms of color blindness), as in Workshop 8
col_blue <- "#0072B2"; col_verm <- "#D55E00"; col_orange <- "#E69F00"; col_green <- "#009E73"
col_sky <- "#56B4E9"; col_purple <- "#CC79A7"; col_gray <- "#8C8C8C"; col_dark <- "#222222"
theme_ws9 <- function(base_size = 12) {
  theme_minimal(base_size = base_size) +
    theme(panel.grid.minor = element_blank(), panel.grid.major = element_line(color = "#E3E3E3", linewidth = 0.4),
          plot.title = element_text(face = "bold", size = base_size + 1), plot.subtitle = element_text(color = "#555555", size = base_size - 1.5),
          plot.caption = element_text(color = "#777777", size = base_size - 3, hjust = 0), legend.position = "top",
          plot.title.position = "plot", plot.caption.position = "plot", strip.text = element_text(face = "bold", size = base_size),
          plot.background = element_rect(fill = "white", color = NA))
}
banner <- function(kind, what) sprintf("# %s for teaching (seed %d). %s Made by plan_power.R (Workshop 9). No real person was recorded; ALTAVO and BRENLOW are invented brands.", kind, SEED, what)
write_banner_csv <- function(df, file, kind, what) {
  writeLines(banner(kind, what), file)
  suppressWarnings(write.table(df, file, sep = ",", row.names = FALSE, col.names = TRUE, append = TRUE, na = "NA", qmethod = "double"))
}
f3 <- function(x) sprintf("%.3f", x)
fdate <- function(d, f = "%A, %B %d, %Y") sub(" 0([1-9])(,|$)", " \\1\\2", format(d, f))   # "February 8", not "February 08"
mcse <- function(p, R) sqrt(p * (1 - p) / R)

# ---- STEP 0 (continued). Inputs from Workshops 6, 7, and 8 -------------------------------------------
# Each workshop folder is read where it lives in the hub; a reader who downloaded only Workshop 9 gets
# byte identical copies in prior_inputs/. When both exist, the MD5 checksums must agree.
cat("---- STEP 0. Inputs from Workshops 6 to 8 and the planning assumptions ----\n")
inputs <- data.frame(
  ws   = c("ws6", "ws7", "ws7", "ws8", "ws8", "ws8", "ws8", "ws8"),
  hub  = c("analysis_output.txt", "analysis_ready_long.csv", "exports/08_Session_Log/SIMULATED_AdStudy_SessionLog.csv",
           "SIMULATED_power.csv", "SIMULATED_all_tests.csv", "SIMULATED_descriptives.csv", "analysis_output.txt", "results_text.md"),
  stringsAsFactors = FALSE)
inputs$copy <- file.path("prior_inputs", inputs$ws, basename(inputs$hub))
inputs$hub_path <- file.path(normalizePath(file.path(script_dir, "..", ".."), mustWork = FALSE), inputs$ws, "data", inputs$hub)
use_hub <- all(file.exists(inputs$hub_path))
if (!use_hub && !all(file.exists(inputs$copy))) stop("Workshop 6 to 8 inputs not found in ../../ws6..8/data or prior_inputs/")
inputs$path <- if (use_hub) inputs$hub_path else inputs$copy
inputs$label <- paste0(inputs$ws, "/data/", inputs$hub)
if (use_hub && all(file.exists(inputs$copy))) {
  same <- unname(tools::md5sum(inputs$hub_path)) == unname(tools::md5sum(inputs$copy))
  if (!all(same)) stop("prior_inputs/ differs from the workshop folders: copy again: ", paste(inputs$copy[!same], collapse = ", "))
  cat(sprintf("Read from the Workshop 6 to 8 folders; prior_inputs/ holds byte identical copies of all %d files (MD5 checksums agree)\n", nrow(inputs)))
} else cat(sprintf("Read from %s\n", if (use_hub) "the Workshop 6 to 8 folders" else "prior_inputs/ (copies shipped with Workshop 9)"))
inp <- function(label) inputs$path[inputs$label == label]

# src_line(): the line number and text of the first line of a file that matches a pattern, so that
# every assumption can name the exact line it came from. The script stops if the line is missing.
src_line <- function(label, pattern, fixed = TRUE) {
  x <- readLines(inp(label), warn = FALSE, encoding = "UTF-8")
  k <- if (fixed) grep(pattern, x, fixed = TRUE) else grep(pattern, x)
  if (length(k) == 0) stop("Source line not found in ", label, ": ", pattern)
  list(where = sprintf("%s, line %d", label, k[1]), text = x[k[1]], line = k[1])
}
read_banner <- function(label) read.csv(inp(label), skip = 1, stringsAsFactors = FALSE, check.names = FALSE)

w7  <- read_banner("ws7/data/analysis_ready_long.csv")
log7 <- read_banner("ws7/data/exports/08_Session_Log/SIMULATED_AdStudy_SessionLog.csv")
p8  <- read_banner("ws8/data/SIMULATED_power.csv")
t8  <- read_banner("ws8/data/SIMULATED_all_tests.csv")
d8  <- read_banner("ws8/data/SIMULATED_descriptives.csv")

# Variance components of liking, fitted here on the Workshop 7 analysis sample (include = 1; Ad_A and
# Ad_B rows): liking ~ ad + (1 | participant). The participant SD and the residual SD are the two
# variance assumptions the follow up study can borrow; the ad SD cannot be estimated from two ads of
# two different brands, so it is a judgment (A5 below).
lk <- subset(w7, include == 1 & stimulus %in% c("Ad_A", "Ad_B"))
vc_fit <- lmer(liking ~ stimulus + (1 | participant_id), data = lk)
vc <- as.data.frame(VarCorr(vc_fit))
sd_part <- vc$sdcor[vc$grp == "participant_id"]
sd_res  <- vc$sdcor[vc$grp == "Residual"]
n_ws7 <- length(unique(w7$participant_id)); n_ws7_in <- length(unique(w7$participant_id[w7$include == 1]))
lik <- d8[d8$variable == "liking", ]
fix8 <- d8[d8$variable == "logo_fixated" & d8$stimulus == "Ad_A", ]
q6 <- t8[t8$id == "Q6_LIKING", ]; h1 <- t8[t8$id == "H1", ]; h3 <- t8[t8$id == "H3_AD_A", ]
rec0 <- as.numeric(p8$value[p8$item == "H1: recall rate when Logo not fixated"])
ws6_done <- src_line("ws6/data/analysis_output.txt", "Completed (Status Processed):")
ws6_excl <- src_line("ws6/data/analysis_output.txt", "Completed respondents screened:")
n6 <- as.numeric(regmatches(ws6_done$text, regexpr("[0-9]+ of [0-9]+", ws6_done$text)) |> strsplit(" of ") |> unlist())
n6b <- as.numeric(regmatches(ws6_excl$text, gregexpr("[0-9]+", ws6_excl$text))[[1]][c(1, 2, 5)])   # screened, excluded, kept (the 20 and 2 are the percent)
# Session spacing in Workshop 7: start times within a day, from the session log (first recording of each participant)
lg <- log7[log7$Recording == 1, ]
st <- as.POSIXct(paste(lg$`Session date`, lg$`Start time`), format = "%Y-%m-%d %H:%M", tz = "UTC")
per_day <- table(lg$`Session date`)
gaps <- unlist(tapply(as.numeric(st), lg$`Session date`, function(s) diff(sort(s)) / 60))

A <- list()   # the planning assumptions, one row each (written to SIMULATED_planning_assumptions.csv)
add_A <- function(id, assumption, value, units, source, how) A[[length(A) + 1]] <<- data.frame(id = id, assumption = assumption,
  value = value, units = units, source = source, how = how, stringsAsFactors = FALSE)
s_lik <- src_line("ws8/data/SIMULATED_descriptives.csv", "Survey,Liking of the ad,liking,1 to 7,Ad_A")
s_q6  <- src_line("ws8/data/results_text.md", "**Q6: Head to head comparisons.**")
s_lim <- src_line("ws8/data/results_text.md", "Only one ad per brand was tested")
s_nxt <- src_line("ws8/data/results_text.md", "**What the next study needs.**")
s_fix <- src_line("ws8/data/SIMULATED_descriptives.csv", "Eye tracking,Logo fixated at least once,logo_fixated,proportion,Ad_A")
s_rec <- src_line("ws8/data/SIMULATED_power.csv", "H1: recall rate when Logo not fixated")
s_or  <- src_line("ws8/data/SIMULATED_all_tests.csv", "H1,H1,planned (Workshop 7)")
s_h1z <- src_line("ws8/data/results_text.md", "the participant variance was estimated at zero (a singular fit)")
s_q6t <- src_line("ws8/data/SIMULATED_all_tests.csv", "Q6,Q6_LIKING")
s_h3t <- src_line("ws8/data/SIMULATED_all_tests.csv", "H3,H3_AD_A")
s_smp <- src_line("ws8/data/analysis_output.txt", "Participants: 22 of 24")
s_pw  <- src_line("ws8/data/analysis_output.txt", "---- STEP 9. Sensitivity power analysis")
add_A("A1", "Mean liking of an ad (1 to 7)", round(mean(lik$mean), 2), "scale points",
      paste0(s_lik$where, " and the next line"), sprintf("Average of the Ad_A and Ad_B means (%.2f and %.2f)", lik$mean[1], lik$mean[2]))
add_A("A2", "Participant SD of liking (random intercept)", round(sd_part, 3), "scale points",
      "ws7/data/analysis_ready_long.csv (include = 1; Ad_A and Ad_B rows)", "lmer(liking ~ stimulus + (1 | participant_id)) fitted here")
add_A("A3", "Residual SD of liking (within person)", round(sd_res, 3), "scale points",
      "ws7/data/analysis_ready_long.csv (include = 1; Ad_A and Ad_B rows)", "Same model; the residual includes the rounding to whole scale points")
add_A("A4", "Brand effect on liking to plan for (smallest effect of interest)", 0.50, "scale points (ALTAVO minus BRENLOW)",
      s_q6$where, sprintf("Decision of the (fictional) client: half a scale point. For reference, the Workshop 8 difference between the two ads was %.2f, 95%% CI [%.2f, %.2f], and it mixes brand and execution",
                          q6$mean_diff, q6$diff_ci_low, q6$diff_ci_high))
add_A("A5", "SD of the ads within a brand (ad random intercept)", 0.20, "scale points", s_lim$where,
      "JUDGMENT, no lab estimate exists (one ad per brand). Two ads of one campaign then differ by about 0.23 points on average (1.128 x 0.20), less than half the 0.55 between two ads of different brands")
add_A("A6", "Ad SD for the sensitivity checks", "0.10 and 0.30", "scale points", "STEP 2, part B", "Half and one and a half times A5")
add_A("A7", "SD of a participant's own brand preference (participant by brand slope)", "0 in the grid; 0.50 in a check", "scale points",
      "STEP 2, part B", "Not estimable from Workshop 7 (one ad per brand); the residual of A3 is reduced so that the total within person variance stays the same")
add_A("A8", "Liking is a whole number from 1 to 7", "1 to 7", "scale points", s_lik$where, "Simulated ratings are rounded and kept inside 1 to 7")
add_A("A9", "Probability that the Logo of an ad is fixated", round(fix8$mean, 3), "proportion", s_fix$where,
      "12 of 22 participants in each ad; fixation is simulated as independent across ads (4 of 22 never fixated either Logo; 4.6 expected if independent)")
add_A("A10", "Brand recall when the Logo was not fixated", rec0, "proportion", s_rec$where, "Observed 7 of 20 in Workshop 7")
add_A("A11", "Odds ratio of recall, Logo fixated versus not (Workshop 8 estimate)", round(as.numeric(h1$es), 2), "odds ratio", s_or$where,
      sprintf("Fisher conditional estimate, 95%% CI [%.2f, %.2f]: wide, so a smaller value is also planned (A12)", as.numeric(h1$es_ci_low), as.numeric(h1$es_ci_high)))
add_A("A12", "Smallest odds ratio of interest for H1", 2.0, "odds ratio", "decision", "A doubling of the odds of recall; protects against planning on an overestimate from 44 observations")
add_A("A13", "Participant SD of recall (logit scale)", 0.50, "logit", s_h1z$where, "Workshop 7 estimated zero (a singular fit) from 22 people; a modest positive value is assumed to be cautious")
add_A("A14", "Ad SD of recall (logit scale)", 0.30, "logit", "judgment", "Ads may differ in how memorable their brand is")
add_A("A15", "Exclusion rate in the lab (rule R2)", round(2 / 24, 4), "proportion", s_smp$where,
      sprintf("%d of %d participants kept in Workshop 7 (2 of 24 excluded)", n_ws7_in, n_ws7))
add_A("A16", "Remote study: completed of those who opened the link", round(n6[1] / n6[2], 4), "proportion", ws6_done$where,
      sprintf("%d of %d in Workshop 6; then %d of %d completed respondents were excluded, so %d of %d (%.1f%%) were kept", n6[1], n6[2], n6b[2], n6b[1], n6b[3], n6[2], 100 * n6b[3] / n6[2]))
add_A("A17", "Paired effect sizes from Workshop 8 for STEP 1", sprintf("liking %.3f; GSR %.3f", as.numeric(q6$es), as.numeric(h3$es)), "d_z",
      paste0(s_q6t$where, "; ", s_h3t$where), "Q6 liking (Ad_A minus Ad_B) and H3 GSR (Ad_A minus Baseline)")
add_A("A18", "Session slot length and slots per day (one Aurora seat)", sprintf("%d min; %d per day", as.integer(round(median(gaps) / 15) * 15), as.integer(max(per_day))), "minutes",
      "ws7/data/exports/08_Session_Log/SIMULATED_AdStudy_SessionLog.csv (Start time)",
      sprintf("Median gap between session starts on a day in Workshop 7 = %d min (range %d to %d), rounded to the nearest 15 min; first sessions near 09:00",
              as.integer(round(median(gaps))), as.integer(round(min(gaps))), as.integer(round(max(gaps)))))
add_A("A19", "Alpha, sidedness, and the power target", ".05, two tailed, .80", "", s_pw$where, "As in Workshops 7 and 8; a Holm family of three has .05 / 3 = .0167 at its strictest step")
A <- do.call(rbind, A)
for (i in seq_len(nrow(A))) {
  cat(sprintf("%-4s %s = %s %s\n       source: %s\n       how: %s\n", A$id[i], A$assumption[i], A$value[i], A$units[i], A$source[i], A$how[i]))
}
cat(sprintf("\nWorkshop 8 called for this study (%s):\n  \"%s\"\n", s_nxt$where, paste(strwrap(sub("^\\*\\*What the next study needs\\.\\*\\* ", "", s_nxt$text), 150), collapse = "\n   ")))
MU <- as.numeric(A$value[A$id == "A1"]); SD_P <- sd_part; SD_E <- sd_res
SD_E_LATENT <- sqrt(SD_E^2 - 1 / 12)    # the rounding to whole points adds about 1/12 to the residual variance
BETA <- 0.50; SD_AD <- 0.20
P_FIX <- fix8$mean; P_REC0 <- rec0; OR_WS8 <- as.numeric(h1$es)
cat(sprintf("\nSimulation values: mean %.2f; participant SD %.3f; residual SD %.3f (latent %.3f before rounding); brand effect %.2f; ad SD %.2f\n\n",
            MU, SD_P, SD_E, SD_E_LATENT, BETA, SD_AD))

# ---- STEP 1. Analytic power for a paired comparison (power.t.test) -----------------------------------
# A paired t test compares two conditions that every participant gives one score for, for example each
# participant's mean liking of the ALTAVO ads and of the BRENLOW ads. It treats the ads as FIXED: it
# answers "do people like THESE ALTAVO ads more than THESE BRENLOW ads?" and says nothing about other
# ads of the brands. That is the right tool for the secondary comparisons of the client report (STEP 4)
# and for H3; the brand claim that should hold for ads in general needs STEP 2.
cat("---- STEP 1. Analytic power for a paired comparison (base R power.t.test, type = \"paired\") ----\n")
dz_list <- c(0.30, 0.50, as.numeric(q6$es), as.numeric(h3$es))
dz_name <- c("small (0.30)", "moderate (0.50)", "Workshop 8 liking estimate (Q6)", "Workshop 7 GSR estimate (H3, Ad_A)")
alphas <- c(.05, .05 / 3)
an <- do.call(rbind, lapply(seq_along(dz_list), function(i) do.call(rbind, lapply(alphas, function(a) {
  nn <- power.t.test(delta = dz_list[i], sd = 1, sig.level = a, power = .80, type = "paired")$n
  data.frame(effect = dz_name[i], d_z = dz_list[i], alpha = a, alpha_label = ifelse(a == .05, ".05 (single test)", ".0167 (strictest Holm step of 3)"),
             n_exact = nn, n_needed = ceiling(nn),
             power_n22 = power.t.test(n = 22, delta = dz_list[i], sd = 1, sig.level = a, type = "paired")$power,
             power_n36 = power.t.test(n = 36, delta = dz_list[i], sd = 1, sig.level = a, type = "paired")$power,
             power_n48 = power.t.test(n = 48, delta = dz_list[i], sd = 1, sig.level = a, type = "paired")$power,
             power_n60 = power.t.test(n = 60, delta = dz_list[i], sd = 1, sig.level = a, type = "paired")$power)
}))))
print(data.frame(effect = an$effect, d_z = f3(an$d_z), alpha = an$alpha_label, n_exact = sprintf("%.2f", an$n_exact), n = an$n_needed,
                 `power n=22` = f3(an$power_n22), `power n=36` = f3(an$power_n36), `power n=48` = f3(an$power_n48),
                 `power n=60` = f3(an$power_n60), check.names = FALSE), row.names = FALSE)
# Continuity check: the four sample sizes of Workshop 8 STEP 9 must be reproduced exactly
chk <- data.frame(item = c("n for d_z = 0.30 (paired)", "n for d_z = 0.50 (paired)", "n for d_z = 0.30, alpha .05/3", "n for d_z = 0.50, alpha .05/3"),
                  here = c(an$n_exact[an$d_z == .30 & an$alpha == .05], an$n_exact[an$d_z == .50 & an$alpha == .05],
                           an$n_exact[an$d_z == .30 & an$alpha < .05], an$n_exact[an$d_z == .50 & an$alpha < .05]))
chk$ws8 <- as.numeric(p8$exact[match(chk$item, p8$item)])
if (any(is.na(chk$ws8)) || max(abs(chk$here - chk$ws8)) > 1e-6) stop("STEP 1 does not reproduce Workshop 8 SIMULATED_power.csv")
cat(sprintf("Continuity check: the 4 sample sizes of Workshop 8 STEP 9 are reproduced (largest difference %.1e): %s\n",
            max(abs(chk$here - chk$ws8)), paste(sprintf("%s = %d", chk$item, ceiling(chk$here)), collapse = "; ")))
cat("Read: for a moderate paired difference (d_z = 0.50), 34 participants give 80% power for one test and 45 at the strictest step of a\n",
    "Holm family of three. The H3 arousal effect of Workshop 7 (d_z = 0.83) needs only 14. The Workshop 8 liking estimate (d_z = 0.36) needs 65,\n",
    "but that estimate mixes brand and execution and comes from 22 people, so it is not used to plan the brand test.\n\n", sep = "")

# ---- STEP 2. Power by SIMULATION: the brand effect on liking, ads as a random factor ----------------
# The key planned test (P1 in analysis_plan_prereg.md):
#     liking ~ brand + (1 | participant) + (1 | ad)      REML, lmerTest, Satterthwaite df
# brand is coded +0.5 for ALTAVO and -0.5 for BRENLOW, so its coefficient is the ALTAVO minus BRENLOW
# difference in mean liking. Each simulated study draws new participants AND new ads (ad effects from
# a normal distribution with SD = A5), simulates every participant rating every ad, rounds to 1 to 7,
# fits the model, and records whether p < .05. Power = the share of simulated studies with p < .05.
#
# The singular fit rule (written into the analysis plan in advance): when the ad variance is estimated
# at zero (a singular fit) or near zero (ad SD below 1 percent of the residual SD), the model has in
# effect dropped the ads as a random factor and its df jump from about (number of ads - 2) to several
# hundred, which makes the test too liberal. In that case the brand effect is tested on the ad means: a
# two sample t test of the ALTAVO ad means against the BRENLOW ad means (df = number of ads - 2), the
# "by ad" analysis that Clark (1973) discussed. In a complete, balanced data set this is exactly the
# test that the mixed model gives whenever the ad variance is clearly above zero, which STEP 2 checks.
# The share of singular fits and the power that the naive test would claim are reported as well.
cat("---- STEP 2. Power by simulation: brand effect on liking, crossed random intercepts for participants and ads ----\n")
cat("Model: liking ~ brand + (1 | participant) + (1 | ad), REML (lme4); brand = +0.5 ALTAVO, -0.5 BRENLOW. Test: the brand effect against\n")
cat("the variation between ads (t test on the ad means, df = ads - 2), which is the lmerTest Satterthwaite test of this balanced design\n")
cat("(checked below) and the singular fit rule at the same time. Alpha .05 two tailed; power also at .0167 (strictest Holm step of 3).\n")

sim_liking <- function(n, m, beta, sd_ad, sd_slope = 0) {
  n_ads <- 2 * m
  brand_ad <- rep(c(0.5, -0.5), each = m)                          # ads 1..m ALTAVO, m+1..2m BRENLOW
  a <- rnorm(n_ads, 0, sd_ad)                                      # ad (execution) effects: NEW ads in every simulated study
  u <- rnorm(n, 0, SD_P)                                           # participant intercepts
  w <- rnorm(n, 0, sd_slope)                                       # participant by brand slopes (0 unless a check asks for them)
  sd_e <- sqrt(SD_E_LATENT^2 - sd_slope^2 / 4)                     # keeps the total within person variance equal
  pid <- rep(seq_len(n), each = n_ads); ad <- rep(seq_len(n_ads), times = n)
  latent <- MU + beta * brand_ad[ad] + u[pid] + w[pid] * brand_ad[ad] + a[ad] + rnorm(n * n_ads, 0, sd_e)
  data.frame(participant = factor(pid), ad = factor(ad), brand = brand_ad[ad], liking = pmin(7, pmax(1, round(latent))))
}
# analyze_liking(): the planned analysis of one simulated study.
# The model is fitted with lme4 (REML). The planned p value is the brand test on the ad means (a two
# sample t test of the ALTAVO ad means against the BRENLOW ad means, df = number of ads - 2). In a
# complete, balanced data set this IS the model's Satterthwaite test whenever the ad variance is above
# zero, and it is the singular fit rule when it is zero, so the planned test is always this one.
# Computing it this way takes about 0.03 s per study instead of about 0.07 s with lmerTest; with
# verify = TRUE the full lmerTest analysis is run as well and compared (first 20 studies of every cell).
# For REAL data with missing ratings, run lmerTest directly (analysis_plan_prereg.md, section 5).
analyze_liking <- function(d, verify = FALSE) {
  fit <- suppressMessages(suppressWarnings(lme4::lmer(liking ~ brand + (1 | participant) + (1 | ad), data = d)))
  sing <- isSingular(fit)
  cf <- coef(summary(fit))                                         # lme4: Estimate, Std. Error, t value
  am <- tapply(d$liking, d$ad, mean); bm <- tapply(d$brand, d$ad, mean)
  xa <- am[bm > 0]; xb <- am[bm < 0]; m <- length(xa)
  dif <- mean(xa) - mean(xb)
  sp2 <- (sum((xa - mean(xa))^2) + sum((xb - mean(xb))^2)) / (2 * m - 2)   # pooled variance of the ad means
  se_am <- sqrt(sp2 * 2 / m); df_am <- 2 * m - 2
  t_am <- dif / se_am
  p_am <- if (is.finite(t_am)) 2 * pt(-abs(t_am), df_am) else as.numeric(dif == 0)
  # the naive test: the model's own t; when the fit is singular, lmerTest's df are those of a model
  # without the ad term (observations - participants - 1)
  df_naive <- if (sing) nrow(d) - nlevels(d$participant) - 1 else df_am
  p_naive <- if (sing) 2 * pt(-abs(cf["brand", "t value"]), df_naive) else p_am
  th <- getME(fit, "theta"); theta_ad <- unname(th[grep("^ad\\.", names(th))])   # estimated ad SD / residual SD
  out <- c(est = cf["brand", "Estimate"], se = se_am, t = t_am, df = df_am, p = p_am, p_naive = p_naive, singular = sing,
           theta_ad = theta_ad, d_t = NA, d_df = NA, d_p = NA, same_decision = NA)
  if (verify) {
    ft <- suppressMessages(suppressWarnings(lmerTest::lmer(liking ~ brand + (1 | participant) + (1 | ad), data = d)))
    ct <- suppressWarnings(summary(ft)$coefficients)["brand", ]
    ref_t <- if (sing) cf["brand", "t value"] else t_am; ref_df <- df_naive; ref_p <- p_naive
    out[c("d_t", "d_df", "d_p")] <- c(ct["t value"] - ref_t, ct["df"] - ref_df, ct["Pr(>|t|)"] - ref_p)
    out["same_decision"] <- as.numeric((ct["Pr(>|t|)"] < .05) == (ref_p < .05))
  }
  out
}
N_VERIFY <- 20   # studies per cell that are also analyzed with lmerTest (Satterthwaite) and compared
run_cell <- function(cell) {
  out <- t(vapply(seq_len(cell$R), function(i) {
    set.seed(SEED + 10000L * cell$k + i)
    analyze_liking(sim_liking(cell$n, cell$m, cell$beta, cell$sd_ad, cell$sd_slope), verify = i <= N_VERIFY)
  }, numeric(12)))
  colnames(out) <- c("est", "se", "t", "df", "p", "p_naive", "singular", "theta_ad", "d_t", "d_df", "d_p", "same_decision")
  out
}
# the formula for a complete, balanced design and a continuous outcome: t with df = 2m - 2 and
# noncentrality = beta / sqrt(2 (sd_ad^2 + sd_e^2 / n) / m). Participants' own intercepts cancel.
power_formula <- function(n, m, beta = BETA, sd_ad = SD_AD, sd_e = SD_E, alpha = .05) {
  df <- 2 * m - 2; ncp <- beta / sqrt(2 * (sd_ad^2 + sd_e^2 / n) / m); tc <- qt(1 - alpha / 2, df)
  1 - pt(tc, df, ncp) + pt(-tc, df, ncp)
}
summarize_cell <- function(cell, out) {
  R <- nrow(out); p05 <- mean(out[, "p"] < .05); ph <- mean(out[, "p"] < .05 / 3); pn <- mean(out[, "p_naive"] < .05)
  v <- out[!is.na(out[, "same_decision"]), , drop = FALSE]; vs <- v[, "singular"] == 1
  vi <- !vs & v[, "theta_ad"] >= 0.01; vn <- !vs & !vi     # clearly above zero; near zero (ad SD below 1% of the residual SD)
  ci90 <- out[, "est"] + outer(out[, "se"] * qt(.95, out[, "df"]), c(-1, 1))
  data.frame(n_participants = cell$n, ads_per_brand = cell$m, ads_total = 2 * cell$m, ratings_per_participant = 2 * cell$m,
             observations = cell$n * 2 * cell$m, brand_effect = cell$beta, ad_sd = cell$sd_ad, slope_sd = cell$sd_slope, n_sim = R,
             power = p05, mcse = mcse(p05, R), mc_low = max(0, p05 - 1.96 * mcse(p05, R)), mc_high = min(1, p05 + 1.96 * mcse(p05, R)),
             power_holm3 = ph, mcse_holm3 = mcse(ph, R), power_formula = power_formula(cell$n, cell$m, cell$beta, cell$sd_ad),
             power_formula_holm3 = power_formula(cell$n, cell$m, cell$beta, cell$sd_ad, alpha = .05 / 3),
             power_if_singular_accepted = pn, pct_singular = 100 * mean(out[, "singular"]),
             mean_estimate = mean(out[, "est"]), sd_estimate = sd(out[, "est"]), mean_se = mean(out[, "se"]), df_planned = 2 * cell$m - 2,
             pct_equivalent = 100 * mean(ci90[, 1] > -0.5 & ci90[, 2] < 0.5),
             verify_n = nrow(v), verify_singular = sum(vs), verify_near_zero = sum(vn), verify_interior = sum(vi),
             verify_max_abs_diff_t_interior = if (any(vi)) max(abs(v[vi, "d_t"])) else NA,
             verify_max_abs_diff_df_interior = if (any(vi)) max(abs(v[vi, "d_df"])) else NA,
             verify_max_abs_diff_p_singular = if (any(vs)) max(abs(v[vs, "d_p"])) else NA,
             verify_same_decision = sum(v[, "same_decision"]))
}
verify_report <- function(tab, what) {
  tot <- sum(tab$verify_n); same <- sum(tab$verify_same_decision)
  mt <- max(tab$verify_max_abs_diff_t_interior, na.rm = TRUE); md <- max(tab$verify_max_abs_diff_df_interior, na.rm = TRUE)
  ms <- suppressWarnings(max(tab$verify_max_abs_diff_p_singular, na.rm = TRUE))
  cat(sprintf("Check (%s): the first %d studies of every cell (%d in all) were also analyzed with lmerTest (Satterthwaite df).\n", what, N_VERIFY, tot))
  cat(sprintf("  ad variance clearly above zero (%d studies): lmerTest t and df equal the test on the ad means (largest |difference|: t %.1e, df %.1e)\n",
              sum(tab$verify_interior), mt, md))
  cat(sprintf("  singular fits (%d): the naive lmerTest p equals the computed naive p (largest |difference| %.1e); the planned test uses the ad means\n",
              sum(tab$verify_singular), if (is.finite(ms)) ms else 0))
  cat(sprintf("  ad SD estimated near zero but not flagged singular (%d): lmerTest's numerical df are unstable there, which is why the rule exists\n",
              sum(tab$verify_near_zero)))
  cat(sprintf("  the lmerTest p and the computed p lead to the same decision at .05 in %d of %d studies\n", same, tot))
  if (mt > 0.01 || md > 0.05 || (is.finite(ms) && ms > 1e-6)) stop("The lmerTest analysis and the planned computation disagree; the shortcut is not valid for these data")
}

# Part A: the grid. 5 participant counts x 4 ad counts = 20 cells
grid_cells <- expand.grid(n = c(24, 36, 48, 60, 72), m = c(2, 3, 4, 6))
grid_cells <- grid_cells[order(grid_cells$m, grid_cells$n), ]
grid_cells$k <- seq_len(nrow(grid_cells)); grid_cells$beta <- BETA; grid_cells$sd_ad <- SD_AD; grid_cells$sd_slope <- 0; grid_cells$R <- N_SIM_GRID
t_a <- Sys.time()
res_grid <- mclapply(split(grid_cells, grid_cells$k), function(cl) run_cell(cl), mc.cores = N_CORES, mc.preschedule = FALSE)
grid <- do.call(rbind, Map(summarize_cell, split(grid_cells, grid_cells$k), res_grid))
cat(sprintf("Part A: %d cells x %d simulated studies = %s mixed model fits in %.1f min\n", nrow(grid_cells), N_SIM_GRID,
            format(nrow(grid_cells) * N_SIM_GRID, big.mark = ","), as.numeric(difftime(Sys.time(), t_a, units = "mins"))))
cat(sprintf("Planning values: brand effect %.2f, ad SD %.2f, participant SD %.3f, residual SD %.3f; MC SE = Monte Carlo standard error of the power\n\n",
            BETA, SD_AD, SD_P, SD_E))
show <- with(grid, data.frame(participants = n_participants, `ads per brand` = ads_per_brand, ratings = observations,
                              power = f3(power), `MC SE` = f3(mcse), `MC 95% interval` = sprintf("[%s, %s]", f3(mc_low), f3(mc_high)),
                              `power, Holm step` = f3(power_holm3), formula = f3(power_formula), `naive (singular accepted)` = f3(power_if_singular_accepted),
                              `% singular` = sprintf("%.1f", pct_singular), `df` = df_planned, check.names = FALSE))
print(show, row.names = FALSE)
wide <- reshape(grid[, c("n_participants", "ads_per_brand", "power")], idvar = "n_participants", timevar = "ads_per_brand", direction = "wide")
names(wide) <- c("participants", paste0(sort(unique(grid$ads_per_brand)), " per brand"))
cat("\nPower (p < .05) by participants (rows) and ads per brand (columns):\n")
print(format(wide, digits = 3), row.names = FALSE)
cat("\n")
verify_report(grid, "grid")
zdev <- (grid$power - grid$power_formula) / grid$mcse
cat(sprintf("Simulation versus the formula for a balanced design: mean difference %+.3f; %d of %d cells differ by more than 1.96 MC SE (about %.1f\n",
            mean(grid$power - grid$power_formula), sum(abs(zdev) > 1.96), nrow(grid), 0.05 * nrow(grid)))
cat(sprintf("expected by chance alone); largest difference %.3f (%.1f MC SE). The formula treats liking as continuous; the simulation rounds it to 1 to 7.\n",
            max(abs(grid$power - grid$power_formula)), max(abs(zdev))))
cat("Neighboring cells can therefore be out of order by a few points (for example 60 and 72 participants with 4 ads per brand); the formula\n",
    "lines in figures/SIMULATED_fig1_power_curves.png show the smooth trend.\n", sep = "")
cat("Read: more participants help only up to a point, because the brand difference must be judged against the variation BETWEEN ads, and\n",
    "with 2 or 3 ads per brand that variation is estimated from 2 or 4 degrees of freedom. With 3 ads per brand (six ads in all) power is\n",
    sprintf("at most %.2f even with 72 participants. Accepting singular fits would overstate power (naive column), most of all with few ads.\n\n",
            max(grid$power[grid$ads_per_brand == 3])), sep = "")

# Part B: checks at the design chosen in STEP 4 (6 ads per brand, 48 participants). These are run here
# so that STEP 4 can use them; the design itself is fixed by the rule in STEP 4.
chk_cells <- data.frame(
  label = c("C1 Type I error: no brand difference", "C2 Ad SD 0.10 (ads of a brand very similar)", "C3 Ad SD 0.30 (ads of a brand more different)",
            "C4 Type I error with a participant by brand slope (SD 0.50)", "C5 Power with a participant by brand slope (SD 0.50)"),
  n = 48, m = 6, beta = c(0, BETA, BETA, 0, BETA), sd_ad = c(SD_AD, 0.10, 0.30, SD_AD, SD_AD), sd_slope = c(0, 0, 0, 0.5, 0.5), stringsAsFactors = FALSE)
chk_cells$k <- 100 + seq_len(nrow(chk_cells)); chk_cells$R <- N_SIM_CHECK
t_b <- Sys.time()
res_chk <- mclapply(split(chk_cells, chk_cells$k), function(cl) run_cell(cl), mc.cores = N_CORES, mc.preschedule = FALSE)
checks <- do.call(rbind, Map(summarize_cell, split(chk_cells, chk_cells$k), res_chk))
checks <- cbind(check = chk_cells$label[order(chk_cells$k)], checks)
cat(sprintf("Part B: checks at 6 ads per brand and 48 participants (%d simulated studies each; %.1f min)\n", N_SIM_CHECK,
            as.numeric(difftime(Sys.time(), t_b, units = "mins"))))
verify_report(checks, "checks")
print(with(checks, data.frame(check, `rejects at .05` = f3(power), `MC SE` = f3(mcse), formula = ifelse(slope_sd > 0, "  -  ", f3(power_formula)),
                              `naive (singular accepted)` = f3(power_if_singular_accepted), `% singular` = sprintf("%.1f", pct_singular),
                              `% 90% CI inside +/-0.5` = sprintf("%.1f", pct_equivalent), check.names = FALSE)), row.names = FALSE)
cat("Read: C1 shows that the planned test keeps its 5% false positive rate when the brands do not differ; the column '% 90% CI inside +/-0.5'\n",
    "in C1 is the chance that the equivalence test (two one sided tests at .05 against +/-0.50) concludes that the brands do not differ by a\n",
    "meaningful amount. C2 and C3 show how strongly the plan depends on the unknown ad SD. C4 and C5 show what happens if participants also\n",
    "differ in their own brand preference, which the planned model does not include (the robustness model in the analysis plan does).\n\n", sep = "")

# ---- STEP 3. Power by SIMULATION for H1: Logo fixation and brand recall -----------------------------
# H1 (replication of Workshop 7): brand recall is more likely when the Logo of the ad was fixated.
#     recall ~ logo_fixated + (1 | participant) + (1 | ad), binomial (glmer, Laplace), Wald z test
# Workshop 8 found that H1 needs about 72 INDEPENDENT observations. In the follow up study every
# participant gives one observation per ad (12 with 6 ads per brand), so the observations are nested in
# people and ads, and the mixed logistic model accounts for that. Fixation is simulated per ad with
# probability A9, recall with the rate A10 when not fixated and the odds ratio A11 or A12.
cat("---- STEP 3. Power by simulation for H1: recall ~ logo_fixated + (1 | participant) + (1 | ad), mixed logistic ----\n")
sim_h1 <- function(n, m, or) {
  n_ads <- 2 * m
  u <- rnorm(n, 0, 0.50); a <- rnorm(n_ads, 0, 0.30)
  pid <- rep(seq_len(n), each = n_ads); ad <- rep(seq_len(n_ads), times = n)
  fix <- rbinom(n * n_ads, 1, P_FIX)
  rec <- rbinom(n * n_ads, 1, plogis(qlogis(P_REC0) + log(or) * fix + u[pid] + a[ad]))
  data.frame(participant = factor(pid), ad = factor(ad), logo_fixated = fix, recall = rec)
}
run_h1 <- function(cell) {
  t(vapply(seq_len(cell$R), function(i) {
    set.seed(SEED + 10000L * cell$k + i)
    d <- sim_h1(cell$n, cell$m, cell$or)
    g <- tryCatch(suppressMessages(suppressWarnings(glmer(recall ~ logo_fixated + (1 | participant) + (1 | ad), data = d, family = binomial))),
                  error = function(e) NULL)
    if (is.null(g)) return(c(NA, NA, NA))
    cf <- summary(g)$coefficients
    c(cf["logo_fixated", "Estimate"], cf["logo_fixated", "Pr(>|z|)"], isSingular(g))
  }, numeric(3)))
}
h1_cells <- expand.grid(n = c(24, 36, 48), or = c(2.0, OR_WS8)); h1_cells$m <- 6
h1_cells$k <- 200 + seq_len(nrow(h1_cells)); h1_cells$R <- N_SIM_H1
t_c <- Sys.time()
res_h1 <- mclapply(split(h1_cells, h1_cells$k), run_h1, mc.cores = N_CORES, mc.preschedule = FALSE)
h1_tab <- do.call(rbind, Map(function(cl, out) {
  ok <- !is.na(out[, 2]); pw <- mean(out[ok, 2] < .05); R <- sum(ok)
  data.frame(n_participants = cl$n, ads_per_brand = cl$m, observations = cl$n * 2 * cl$m, odds_ratio = cl$or, n_sim = R, n_failed = sum(!ok),
             power = pw, mcse = mcse(pw, R), mc_low = max(0, pw - 1.96 * mcse(pw, R)), mc_high = min(1, pw + 1.96 * mcse(pw, R)),
             pct_singular = 100 * mean(out[ok, 3]), mean_log_or = mean(out[ok, 1]))
}, split(h1_cells, h1_cells$k), res_h1))
cat(sprintf("%d cells x %d simulated studies (%.1f min); 6 ads per brand, so 12 observations per participant\n", nrow(h1_cells), N_SIM_H1,
            as.numeric(difftime(Sys.time(), t_c, units = "mins"))))
print(with(h1_tab, data.frame(participants = n_participants, observations, `odds ratio` = sprintf("%.2f", odds_ratio), power = f3(power),
                              `MC SE` = f3(mcse), `MC 95% interval` = sprintf("[%s, %s]", f3(mc_low), f3(mc_high)), `% singular` = sprintf("%.1f", pct_singular),
                              `mean estimated OR` = sprintf("%.2f", exp(mean_log_or)), failed = n_failed, check.names = FALSE)), row.names = FALSE)
cat("Read: with 12 ads per person, even 24 participants give several hundred observations, and H1 is well powered for the Workshop 8 odds ratio.\n",
    "For the smaller odds ratio of 2.0 the number of participants matters more. N_SIM_H1 is smaller than in STEP 2 because a mixed logistic fit\n",
    "takes about four times as long as a linear one; the Monte Carlo SE column shows the precision this buys.\n\n", sep = "")

# ---- STEP 4. The chosen design and its reasoning ------------------------------------------------------
# REASONING (written before the simulation was run; STEP 4 checks that the numbers support it):
# 1. Smallest effect of interest. The client decided that a brand difference smaller than half a point
#    on the 7 point liking scale would not change a decision (A4). The plan is powered for 0.50, not for
#    the Workshop 8 estimate (0.55), because that estimate mixes brand with execution and comes from 22
#    people; planning on an estimate from a small study tends to give too few participants.
# 2. Power target. 80% at alpha = .05 (two tailed) for the primary test P1 (brand effect on liking,
#    ads random), as in Workshop 8.
# 3. Ads per brand. The grid shows that the number of ads, not the number of participants, limits the
#    power of a brand claim. With 2, 3, or 4 ads per brand, no number of participants up to 72 reaches
#    80%. The plan therefore asks the client for 6 ads per brand (12 ads), the smallest number in the
#    grid that reaches the target. Twelve 30 s ads with two questions each take about 10 minutes, which
#    keeps the recording near 20 minutes (STEP 6).
# 4. Participants. With 6 ads per brand, the smallest number of participants in the grid that reaches
#    80% and also covers the secondary paired comparisons (45 participants for d_z = 0.50 at the
#    strictest step of a Holm family of three, STEP 1) is 48. Going from 48 to 72 participants adds
#    little power for P1, because the ads limit it. 48 is also a multiple of 12, the number of rows of
#    the Williams design, so every row is used equally often (STEP 5).
# 5. Attrition and exclusions. Workshop 7 lost 2 of 24 participants (8.3%) to rule R2 in the lab. To end
#    with 48 analyzable participants, about 48 / (1 - 0.083) = 53 sessions are needed, so 5 reserve
#    sessions are booked; no shows are rebooked into 2 open slots. Data collection stops when 48
#    participants meet the rules, or at a hard cap of 60 sessions (25% above the target). Workshop 6
#    showed that remote studies lose far more (only 142 of 260 who opened the link were kept, 54.6%):
#    the same 48 would need about 88 people to open a remote link, so this study is not moved online
#    without a new plan.
# 6. Sensitivity. If ads of one brand differ more than assumed (SD 0.30), power at this design falls
#    (check C3); the pilot cannot estimate the ad SD, so the plan names this risk instead of hiding it.
cat("---- STEP 4. The chosen design ----\n")
target <- .80
reach <- grid[grid$power >= target, ]
if (nrow(reach) == 0) stop("No cell of the grid reaches the power target")
m_choice <- min(reach$ads_per_brand)
n_holm <- an$n_needed[an$d_z == .50 & an$alpha < .05]
n_choice <- min(reach$n_participants[reach$ads_per_brand == m_choice & reach$n_participants >= n_holm])
row_c <- grid[grid$ads_per_brand == m_choice & grid$n_participants == n_choice, ]
if (!QUICK && (m_choice != 6 || n_choice != 48)) stop("The simulation no longer supports the comments of STEP 4 (6 ads per brand, 48 participants): revise them")
excl <- 2 / 24
n_sessions <- ceiling(n_choice / (1 - excl)); n_reserve <- n_sessions - n_choice; n_cap <- n_choice * 1.25
n_remote <- ceiling(n_choice / (n6b[3] / n6[2]))
best3 <- max(grid$power[grid$ads_per_brand == 3]); best4 <- max(grid$power[grid$ads_per_brand == 4])
cat(sprintf("Rule: the smallest number of ads per brand with power >= %.2f somewhere in the grid, then the smallest number of participants with\n", target))
cat(sprintf("power >= %.2f that also covers the secondary Holm family (n >= %d, STEP 1).\n", target, n_holm))
cat(sprintf("Ads per brand: best power with 3 per brand = %.3f, with 4 per brand = %.3f; 6 per brand reaches the target -> %d ads per brand (%d ads)\n",
            best3, best4, m_choice, 2 * m_choice))
cat(sprintf("Participants: %d (power %.3f, MC SE %.3f, MC 95%% interval [%.3f, %.3f]; formula %.3f; at the strictest Holm step %.3f)\n",
            n_choice, row_c$power, row_c$mcse, row_c$mc_low, row_c$mc_high, row_c$power_formula, row_c$power_holm3))
cat(sprintf("  for comparison: 72 participants with 6 per brand = %.3f; 48 participants with 3 per brand = %.3f\n",
            grid$power[grid$ads_per_brand == 6 & grid$n_participants == 72], grid$power[grid$ads_per_brand == 3 & grid$n_participants == 48]))
cat(sprintf("Type I error of the planned test at this design (C1): %.3f (MC SE %.3f); naive test accepting singular fits: %.3f\n",
            checks$power[1], checks$mcse[1], checks$power_if_singular_accepted[1]))
cat(sprintf("Sensitivity to the ad SD: 0.10 -> %.3f; 0.20 -> %.3f; 0.30 -> %.3f\n", checks$power[2], row_c$power, checks$power[3]))
cat(sprintf("With a participant by brand slope (SD 0.50): Type I error %.3f, power %.3f\n", checks$power[4], checks$power[5]))
cat(sprintf("H1 at %d participants: power %.3f for OR = 2.00 and %.3f for OR = %.2f\n", n_choice, h1_tab$power[h1_tab$n_participants == n_choice & h1_tab$odds_ratio == 2],
            h1_tab$power[h1_tab$n_participants == n_choice & h1_tab$odds_ratio != 2], OR_WS8))
cat(sprintf("Attrition: Workshop 7 exclusion rate 2/24 = %.3f -> %d / (1 - %.3f) = %.2f -> %d sessions: %d main + %d reserve; hard cap %d sessions\n",
            excl, n_choice, excl, n_choice / (1 - excl), n_sessions, n_choice, n_reserve, n_cap))
cat(sprintf("Remote comparison (Workshop 6: %d of %d kept, %.1f%%): %d analyzable participants would need about %d people to open the link\n",
            n6b[3], n6[2], 100 * n6b[3] / n6[2], n_choice, n_remote))
cat("Stopping rule: stop when 48 participants meet rules R1 to R2, or after 60 sessions, whichever comes first; an excluded participant is\n",
    "replaced by the next reserve code, who takes the same Latin square row. Outcome data are not analyzed before the stopping point.\n\n", sep = "")

# ---- STEP 5. Counterbalancing: a Williams design -------------------------------------------------------
# A balanced Latin square for an even number of ads t (Williams, 1949): row 1 is 0, 1, t-1, 2, t-2, 3, ...
# and row r adds r - 1 to every entry (mod t). Every ad appears once in every row and once in every
# position, and every ad is IMMEDIATELY followed by every other ad exactly once, so first order carry
# over (what the previous ad does to the next) is balanced. One square of t rows suffices for even t.
# Which symbols belong to which brand is a free choice that does not change these properties, but it
# decides how the BRANDS are ordered. A poor choice shows all ads of one brand in a block in some rows.
# The script tries every way to split the t symbols into two brands and keeps the one with the shortest
# runs of one brand; labeling by parity (even symbols ALTAVO, odd symbols BRENLOW) is such a split: no
# row shows more than 2 ads of one brand in a row, and the brand pattern is ABBAABBA... or BAABBAAB...
cat("---- STEP 5. Counterbalancing: Williams design, brand order, assignment of participants to rows ----\n")
williams <- function(t) {
  if (t %% 2 != 0) stop("one Williams square needs an even number of ads")
  first <- vapply(seq_len(t), function(j) if (j == 1) 0 else if (j %% 2 == 0) j / 2 else t - (j - 1) / 2, numeric(1))
  t(vapply(seq_len(t), function(r) (first + r - 1) %% t, numeric(t)))
}
ad_label <- function(v, t) ifelse(v %% 2 == 0, paste0("ALTAVO_", v / 2 + 1), paste0("BRENLOW_", (v + 1) / 2))   # parity labeling
check_williams <- function(W) {
  t <- ncol(W)
  rows_ok <- all(apply(W, 1, function(x) length(unique(x)) == t)); cols_ok <- all(apply(W, 2, function(x) length(unique(x)) == t))
  pairs <- table(factor(W[, -t], levels = 0:(t - 1)), factor(W[, -1], levels = 0:(t - 1)))
  c(rows = rows_ok, cols = cols_ok, carry = all(pairs[row(pairs) != col(pairs)] == 1) && all(diag(pairs) == 0))
}
for (tt in c(4, 6, 8, 12)) {
  ck <- check_williams(williams(tt))
  cat(sprintf("t = %2d ads: each ad once per row %s, once per position %s, each ordered pair adjacent exactly once %s\n", tt, ck[1], ck[2], ck[3]))
}
T_ADS <- 2 * m_choice
W <- williams(T_ADS)
if (!all(check_williams(W))) stop("Williams square failed its checks")
# every split of the T_ADS symbols into two brands of equal size (symbol 0 always ALTAVO, so each split counts once)
splits <- combn(seq_len(T_ADS - 1), T_ADS / 2 - 1, function(x) c(0, x))
max_run <- apply(splits, 2, function(a) max(apply(W, 1, function(x) max(rle(x %in% a)$lengths))))
par_run <- max(apply(W, 1, function(x) max(rle(x %% 2 == 0)$lengths)))
cat(sprintf("\nBrand labeling: %d ways to split %d ads into two brands; longest run of one brand within a row ranges from %d to %d;\n",
            ncol(splits), T_ADS, min(max_run), max(max_run)))
cat(sprintf("%d of them reach the minimum of %d; the parity labeling (even symbols ALTAVO, odd BRENLOW) gives %d\n", sum(max_run == min(max_run)),
            min(max_run), par_run))
if (par_run != min(max_run)) stop("The parity labeling is not optimal for this number of ads")
WL <- matrix(ad_label(W, T_ADS), nrow = T_ADS)
brand_of <- function(lbl) ifelse(startsWith(lbl, "ALTAVO"), "A", "B")
cat(sprintf("\nThe square for the chosen design (%d ads; A = ALTAVO, B = BRENLOW):\n", T_ADS))
for (r in seq_len(T_ADS)) cat(sprintf("  row %2d: %s   brands %s\n", r, paste(sprintf("%-9s", WL[r, ]), collapse = " "), paste(brand_of(WL[r, ]), collapse = "")))
per_pos <- colSums(matrix(brand_of(WL) == "A", nrow = T_ADS))
runs <- apply(WL, 1, function(x) max(rle(brand_of(x))$lengths))
cat(sprintf("Brand order: ALTAVO ads per position across the %d rows = %s (always %d of %d); rows that start with ALTAVO = %d, with BRENLOW = %d;\n",
            T_ADS, paste(per_pos, collapse = " "), T_ADS / 2, T_ADS, sum(brand_of(WL[, 1]) == "A"), sum(brand_of(WL[, 1]) == "B")))
cat(sprintf("longest run of one brand within a row = %d ads (rows: %s)\n", max(runs), paste(runs, collapse = " ")))
# Assignment of participants to rows: blocks of 12 consecutive participants; within each block, the 12
# rows in a random order (fixed seed). If data collection stopped after any complete block, every row
# would still be used equally often. Reserve participants take the row of the participant they replace.
set.seed(SEED + 5)
n_blocks <- n_choice / T_ADS
rows_main <- unlist(lapply(seq_len(n_blocks), function(b) sample(seq_len(T_ADS))))
codes_main <- sprintf("P%02d", seq_len(n_choice))
ord <- data.frame(participant_code = codes_main, phase = "main", block = rep(seq_len(n_blocks), each = T_ADS), latin_square_row = rows_main, stringsAsFactors = FALSE)
pil <- data.frame(participant_code = c("X01", "X02"), phase = "pilot", block = NA, latin_square_row = c(1, 2), stringsAsFactors = FALSE)
res <- data.frame(participant_code = sprintf("R%02d", seq_len(n_reserve)), phase = "reserve", block = NA, latin_square_row = NA, stringsAsFactors = FALSE)
ord <- rbind(pil, ord, res)
pos <- t(vapply(ord$latin_square_row, function(r) if (is.na(r)) rep(NA_character_, T_ADS) else WL[r, ], character(T_ADS)))
colnames(pos) <- sprintf("position_%02d", seq_len(T_ADS))
ord <- cbind(ord, pos)
ord$brand_sequence <- apply(pos, 1, function(x) if (all(is.na(x))) NA else paste(brand_of(x), collapse = ""))
ord$first_brand <- ifelse(is.na(ord$brand_sequence), NA, ifelse(substr(ord$brand_sequence, 1, 1) == "A", "ALTAVO", "BRENLOW"))
ord$note <- ifelse(ord$phase == "pilot", "pilot: data checked, not analyzed", ifelse(ord$phase == "reserve", "takes the row of the participant it replaces", ""))
mm <- ord[ord$phase == "main", ]
cnt_pos <- vapply(seq_len(T_ADS), function(j) all(table(mm[[sprintf("position_%02d", j)]]) == n_blocks), logical(1))
adj <- table(factor(as.vector(as.matrix(mm[, sprintf("position_%02d", 1:(T_ADS - 1))]))), factor(as.vector(as.matrix(mm[, sprintf("position_%02d", 2:T_ADS)]))))
cat(sprintf("Main participants: %d in %d blocks of %d; every row used %d times; every ad in every position %d times (%s);\n", n_choice, n_blocks, T_ADS,
            n_blocks, n_blocks, all(cnt_pos)))
cat(sprintf("every ordered pair of different ads adjacent %d times (%s); first brand ALTAVO for %d and BRENLOW for %d participants\n",
            n_blocks, all(adj[row(adj) != col(adj)] == n_blocks), sum(mm$first_brand == "ALTAVO"), sum(mm$first_brand == "BRENLOW")))
cat("Rows by participant (main):", paste(sprintf("%s=%d", mm$participant_code, mm$latin_square_row), collapse = " "), "\n")
cat("In iMotions, the sequences can be entered as study flows or as a test plan (Using a Test Plan for Study Design: a CSV file with the\n",
    "respondents and their stimulus sequences). The test plan article says respondents cannot be added after recording starts, so the reserve\n",
    "codes must be in the file from the beginning (Items to verify in README.md).\n", sep = "")
sq <- data.frame(row = seq_len(T_ADS), WL, stringsAsFactors = FALSE); names(sq)[-1] <- sprintf("position_%02d", seq_len(T_ADS))
sq$brand_sequence <- apply(WL, 1, function(x) paste(brand_of(x), collapse = ""))
write_banner_csv(sq, "SIMULATED_latin_square.csv", "PLANNED DESIGN", sprintf("Williams design for %d ads (one balanced Latin square): row, ad in each position, brand sequence (A = ALTAVO, B = BRENLOW).", T_ADS))
write_banner_csv(ord, "SIMULATED_order_schedule.csv", "PLANNED DESIGN",
                 sprintf("Order of the %d ads for each participant code: Latin square row, the ad in each position, the brand sequence. Codes only; pilot X, main P, reserve R.", T_ADS))
cat("wrote SIMULATED_latin_square.csv and SIMULATED_order_schedule.csv\n\n")

# ---- STEP 6. Session schedule for one Aurora seat ----------------------------------------------------
# The lab has ONE Smart Eye Aurora, so screen based eye tracking runs one participant at a time. The
# session length comes from the Workshop 7 design (Instructions, a 60 s Baseline, 30 s ads, survey
# questions) extended to the chosen number of ads, plus the lab procedure manual's times for consent,
# setup, completion, and the post session tasks. Two people staff every session: a lead operator
# (iMotions, calibration, monitoring) and an assistant (greeting, consent, sensors, cleaning). Staff
# carry codes S1 to S6 and work half day shifts in three fixed pairs.
cat("---- STEP 6. Session schedule (one Aurora seat, two staff per session) ----\n")
rec <- data.frame(part = c("Instructions (self paced)", "Baseline, gray screen with a fixation cross", sprintf("Ads, 30 s each (Workshop 7: 2; follow up: %d)", T_ADS),
                           "Liking and purchase intention after each ad (about 20 s per ad)", "End questions about each ad (Workshop 7: liking and recall of 2 ads; follow up: recall of each ad), about 10 s per item",
                           "Attitude items and demographics"),
                  ws7_seconds = c(60, 60, 60, 0, 40, 100), seconds = c(60, 60, 30 * T_ADS, 20 * T_ADS, 10 * T_ADS, 100),
                  source = c("Workshop 7 design (ws7/data/README.md, section 1)", "Workshop 7 design", "Workshop 7 design", "PLANNED; time to verify in the pilot",
                             "PLANNED; time to verify in the pilot", "Workshop 7 survey (ATT_1 to ATT_3R) plus age and gender"))
print(rec[, c("part", "ws7_seconds", "seconds", "source")], row.names = FALSE)
rec_min <- sum(rec$seconds) / 60
cat(sprintf("Recording: %.1f min for the follow up design (Workshop 7 design: %.1f min); planned as 20 min with time for a recalibration\n", rec_min, sum(rec$ws7_seconds) / 60))
sess <- data.frame(part = c("Arrival, consent, screening", "Sensor setup: Shimmer3 GSR+ on two fingers, Aurora positioning and calibration, camera check",
                            "Recording (above, rounded up)", "Sensors off, debrief, compensation", "Backup, session log, cleaning, room empty for ventilation"),
                   minutes = c(15, 10, 20, 15, 30),
                   source = c("lab procedure manual 7.1.3 (15 minutes)", "PLANNED: the manual's 20 to 30 minutes (7.1.4) include the EEG cap, which this study does not use; verify in the pilot",
                              "STEP 6 above", "lab procedure manual 7.1.6 (15 minutes)", "lab procedure manual 7.1.7 (30 minutes); 3.4 (at least 10 minutes with the room empty); 7.2 (20 to 30 minutes between participants)"))
print(sess, row.names = FALSE)
part_min <- sum(sess$minutes[1:4]); slot_min <- sum(sess$minutes)
cat(sprintf("Participant time %d min (the consent form says about 60 minutes); slot %d min, the same as the Workshop 7 slots (A18); setup 45 min before\n",
            part_min, slot_min))
cat("the first session of a day (lab procedure manual 7.1.2).\n")
SLOTS <- c("09:00", "10:30", "12:00", "13:30", "15:00")
start_main <- as.Date("2027-02-08"); pilot_day <- as.Date("2027-02-01")
n_days <- ceiling((n_choice + n_reserve) / length(SLOTS))
days <- seq(start_main, by = "day", length.out = n_days * 2)
days <- days[!format(days, "%u") %in% c("6", "7")][seq_len(n_days)]
pairs <- list(c("S1", "S2"), c("S3", "S4"), c("S5", "S6"))
tm <- function(hhmm, add) format(as.POSIXct(paste("2027-01-01", hhmm), tz = "UTC") + add * 60, "%H:%M")
slots <- expand.grid(slot = seq_along(SLOTS), day = seq_len(n_days))
slots$date <- days[slots$day]
slots$shift <- 2 * (slots$day - 1) + (slots$slot >= 4)
slots$pair <- slots$shift %% 3 + 1
k <- seq_len(nrow(slots))
slots$phase <- ifelse(k <= n_choice, "main", ifelse(k <= n_choice + n_reserve, "reserve", "open"))
slots$participant_code <- ifelse(slots$phase == "main", sprintf("P%02d", k), ifelse(slots$phase == "reserve", sprintf("R%02d", k - n_choice), NA))
pilot <- data.frame(slot = 1:2, day = 0, date = pilot_day, shift = c(-2, -2), pair = c(1, 2), phase = "pilot", participant_code = c("X01", "X02"))
S <- rbind(pilot, slots)
S$session <- c("PILOT1", "PILOT2", sprintf("S%02d", seq_len(nrow(slots))))
S$weekday <- format(S$date, "%A")
S$arrive <- SLOTS[S$slot]; S$recording_start <- tm(S$arrive, 25); S$recording_end <- tm(S$arrive, 45)
S$participant_leaves <- tm(S$arrive, part_min); S$room_ready <- tm(S$arrive, slot_min)
S$latin_square_row <- ord$latin_square_row[match(S$participant_code, ord$participant_code)]
S$lead_operator <- vapply(S$pair, function(p) pairs[[p]][1], ""); S$assistant <- vapply(S$pair, function(p) pairs[[p]][2], "")
S$notes <- ifelse(S$phase == "pilot", "pilot; all staff observe; go or no go meeting after the second pilot",
           ifelse(S$phase == "reserve", "replaces an excluded participant (same row); cancel if not needed",
           ifelse(S$phase == "open", "open slot for rebooking a no show", "")))
first_of_day <- !duplicated(S$date)
S$notes[first_of_day] <- ifelse(S$notes[first_of_day] == "", "setup from 08:15", paste0(S$notes[first_of_day], "; setup from 08:15"))
S <- S[, c("session", "phase", "date", "weekday", "arrive", "recording_start", "recording_end", "participant_leaves", "room_ready",
           "participant_code", "latin_square_row", "lead_operator", "assistant", "notes")]
S$date <- format(S$date, "%Y-%m-%d")
write_banner_csv(S, "SIMULATED_session_schedule.csv", "PLANNED SCHEDULE",
                 "One row per session slot for one Aurora seat (90 min slots, 5 per weekday). Placeholder dates; participant and staff codes only.")
cat(sprintf("Main collection: %d slots on %d weekdays (%s to %s): %d main, %d reserve, %d open; pilot on %s (2 sessions)\n",
            nrow(slots), n_days, fdate(min(days)), fdate(max(days)), n_choice, n_reserve, sum(slots$phase == "open"), fdate(pilot_day)))
cat(sprintf("Staff time: %d sessions x 2 people x %d min = %.0f person hours, plus 45 min setup per day for the first pair\n",
            n_choice + n_reserve, slot_min, (n_choice + n_reserve) * 2 * slot_min / 60))
cat("PLACEHOLDER DATES: set them after IRB approval and check the JSU academic calendar. wrote SIMULATED_session_schedule.csv\n\n")

# ---- STEP 7. Figures ---------------------------------------------------------------------------------
cat("---- STEP 7. Figures ----\n")
save_fig <- function(p, file, h_px) {
  ggsave(file.path(fig_dir, file), p, width = 1400 / 150, height = h_px / 150, dpi = 150, bg = "white")
  cat("saved figure: figures/", file, "\n", sep = "")
}
m_cols <- c("2" = col_orange, "3" = col_sky, "4" = col_green, "6" = col_blue)
m_shapes <- c("2" = 15, "3" = 17, "4" = 18, "6" = 16)
fg <- grid; fg$m <- factor(fg$ads_per_brand)
ff <- expand.grid(n = seq(20, 76, 1), m = c(2, 3, 4, 6)); ff$power <- mapply(power_formula, ff$n, ff$m); ff$m <- factor(ff$m)
p1 <- ggplot(fg, aes(n_participants, power, color = m, shape = m)) +
  geom_hline(yintercept = .80, linetype = "dashed", color = col_gray) +
  annotate("text", x = 76, y = .785, label = "80% target", hjust = 1, vjust = 1, color = "#555555", size = 3.4) +
  geom_line(data = ff, aes(n, power, group = m, color = m), linetype = "22", linewidth = 0.5, alpha = 0.8, inherit.aes = FALSE) +
  geom_line(linewidth = 0.9) + geom_errorbar(aes(ymin = mc_low, ymax = mc_high), width = 1.2, linewidth = 0.5) + geom_point(size = 3) +
  annotate("point", x = n_choice, y = row_c$power, size = 7, shape = 21, color = col_dark, fill = NA, stroke = 1.1) +
  annotate("text", x = n_choice - 1, y = 1.035, label = sprintf("chosen design: %d ads per brand, %d participants, power %.2f", m_choice, n_choice, row_c$power),
           hjust = 0.5, size = 3.5, color = col_dark, fontface = "bold") +
  annotate("segment", x = n_choice, xend = n_choice, y = 1.015, yend = row_c$power + 0.035, color = col_dark, linewidth = 0.4) +
  scale_color_manual(values = m_cols, name = "Ads per brand") + scale_shape_manual(values = m_shapes, name = "Ads per brand") +
  scale_x_continuous(breaks = c(24, 36, 48, 60, 72), limits = c(20, 76)) + scale_y_continuous(limits = c(0, 1.05), breaks = seq(0, 1, .2)) +
  labs(title = "SIMULATED / PLANNED: power of the brand test on liking, ads as a random factor",
       subtitle = sprintf("liking ~ brand + (1 | participant) + (1 | ad); brand difference 0.50 points; ad SD %.2f; %d simulated studies per point\nBars: Monte Carlo 95%% intervals. Dashed lines: formula for a complete, balanced design.", SD_AD, N_SIM_GRID),
       x = "Participants (each rates every ad)", y = "Power (share of simulated studies with p < .05)",
       caption = "SIMULATED planning data (plan_power.R, Workshop 9). Fictional brands ALTAVO and BRENLOW. More participants cannot make up for too few ads.") +
  theme_ws9() + theme(legend.title = element_text(size = 11))
save_fig(p1, "SIMULATED_fig1_power_curves.png", 900)
fb <- expand.grid(m = 2:10, sd_ad = c(0.10, 0.20, 0.30)); fb$power <- mapply(function(m, s) power_formula(n_choice, m, sd_ad = s), fb$m, fb$sd_ad)
fb$sd_lab <- factor(sprintf("ad SD %.2f", fb$sd_ad))
pts <- rbind(data.frame(m = grid$ads_per_brand[grid$n_participants == n_choice], sd_ad = SD_AD, power = grid$power[grid$n_participants == n_choice],
                        lo = grid$mc_low[grid$n_participants == n_choice], hi = grid$mc_high[grid$n_participants == n_choice]),
             data.frame(m = 6, sd_ad = c(.10, .30), power = checks$power[2:3], lo = checks$mc_low[2:3], hi = checks$mc_high[2:3]))
pts$sd_lab <- factor(sprintf("ad SD %.2f", pts$sd_ad))
sd_cols <- c("ad SD 0.10" = col_green, "ad SD 0.20" = col_blue, "ad SD 0.30" = col_verm)
p2 <- ggplot(fb, aes(m, power, color = sd_lab)) + geom_hline(yintercept = .80, linetype = "dashed", color = col_gray) +
  geom_vline(xintercept = m_choice, linetype = "dotted", color = col_dark) +
  annotate("text", x = m_choice + 0.1, y = 0.05, label = sprintf("chosen: %d per brand", m_choice), hjust = 0, size = 3.4, color = col_dark) +
  geom_line(linewidth = 0.9) + geom_errorbar(data = pts, aes(ymin = lo, ymax = hi), width = 0.15) + geom_point(data = pts, aes(shape = sd_lab), size = 3) +
  scale_color_manual(values = sd_cols, name = NULL) + scale_shape_manual(values = c(16, 17, 15), name = NULL) +
  scale_x_continuous(breaks = 2:10) + scale_y_continuous(limits = c(0, 1), breaks = seq(0, 1, .2)) +
  labs(title = sprintf("SIMULATED / PLANNED: power with %d participants as the number of ads per brand grows", n_choice),
       subtitle = "Lines: formula for a complete, balanced design. Points: simulation with Monte Carlo 95% intervals. The ad SD is a planning judgment.",
       x = "Ads per brand", y = "Power of the brand test (p < .05)",
       caption = "SIMULATED planning data (plan_power.R, Workshop 9). How different the ads of one brand are decides how many ads are needed.") + theme_ws9()
save_fig(p2, "SIMULATED_fig2_power_by_ads.png", 800)
h1p <- h1_tab; h1p$or_lab <- factor(sprintf("odds ratio %.2f", h1p$odds_ratio))
p3 <- ggplot(h1p, aes(n_participants, power, color = or_lab, shape = or_lab)) + geom_hline(yintercept = .80, linetype = "dashed", color = col_gray) +
  geom_line(linewidth = 0.9) + geom_errorbar(aes(ymin = mc_low, ymax = mc_high), width = 1) + geom_point(size = 3) +
  scale_color_manual(values = c(col_verm, col_blue), name = NULL) + scale_shape_manual(values = c(17, 16), name = NULL) +
  scale_x_continuous(breaks = c(24, 36, 48)) + scale_y_continuous(limits = c(0, 1), breaks = seq(0, 1, .2)) +
  labs(title = "SIMULATED / PLANNED: power for H1, Logo fixation and brand recall",
       subtitle = sprintf("recall ~ logo_fixated + (1 | participant) + (1 | ad), mixed logistic; %d ads per participant; %d simulated studies per point", T_ADS, N_SIM_H1),
       x = "Participants", y = "Power (p < .05)", caption = "SIMULATED planning data (plan_power.R, Workshop 9). Odds ratio 4.34 = Workshop 8 estimate; 2.00 = smallest odds ratio of interest.") +
  theme_ws9()
save_fig(p3, "SIMULATED_fig3_power_h1.png", 750)
sqd <- data.frame(row = rep(seq_len(T_ADS), times = T_ADS), pos = rep(seq_len(T_ADS), each = T_ADS), lab = as.vector(WL))
sqd$brand <- ifelse(startsWith(sqd$lab, "ALTAVO"), "ALTAVO", "BRENLOW"); sqd$short <- sub("ALTAVO_", "A", sub("BRENLOW_", "B", sqd$lab))
p4 <- ggplot(sqd, aes(pos, row, fill = brand)) + geom_tile(color = "white", linewidth = 1.2) +
  geom_text(aes(label = short), color = "white", fontface = "bold", size = 4) +
  scale_fill_manual(values = c(ALTAVO = col_blue, BRENLOW = col_verm), name = NULL) +
  scale_x_continuous(breaks = seq_len(T_ADS), expand = c(0, 0), position = "top") +
  scale_y_reverse(breaks = seq_len(T_ADS), labels = sprintf("row %d", seq_len(T_ADS)), expand = c(0, 0)) +
  labs(title = sprintf("PLANNED: Williams design for %d ads (A1 to A%d = ALTAVO, B1 to B%d = BRENLOW)", T_ADS, T_ADS / 2, T_ADS / 2),
       subtitle = sprintf("Each ad once in every row and every position; each ad follows every other ad exactly once. %d participants: each row used %d times.", n_choice, n_blocks),
       x = "Position in the sequence", y = NULL, caption = "PLANNED design (plan_power.R, Workshop 9). Fictional brands. The row of each participant code is in SIMULATED_order_schedule.csv.") +
  coord_fixed(ratio = 0.62) + theme_ws9() + theme(panel.grid.major = element_blank(), axis.text = element_text(size = 10))
save_fig(p4, "SIMULATED_fig4_williams_square.png", 1000)
cat("\n")

# ---- STEP 8. Summary tables -------------------------------------------------------------------------
cat("---- STEP 8. Summary tables ----\n")
write_banner_csv(A, "SIMULATED_planning_assumptions.csv", "PLANNING ASSUMPTIONS", "Every value the simulations use, with its source line in the Workshop 6 to 8 files or the label JUDGMENT or decision.")
write_banner_csv(an, "SIMULATED_power_analytic.csv", "SIMULATED POWER", "STEP 1: analytic power of a paired t test (base R power.t.test, type = paired) for four d_z values, alpha .05 and .05/3.")
gout <- grid; gout$status <- ifelse(gout$ads_per_brand == m_choice & gout$n_participants == n_choice, "CHOSEN DESIGN", "")
write_banner_csv(gout, "SIMULATED_power_grid.csv", "SIMULATED POWER",
                 sprintf("STEP 2: power of the brand test on liking by participants and ads per brand (%d simulated studies per cell), with Monte Carlo SE, the formula for a balanced design, and the share of singular fits.", N_SIM_GRID))
write_banner_csv(checks, "SIMULATED_power_checks.csv", "SIMULATED POWER", sprintf("STEP 2, part B: checks at %d ads per brand and %d participants (Type I error, ad SD, participant by brand slope).", m_choice, n_choice))
write_banner_csv(h1_tab, "SIMULATED_power_h1.csv", "SIMULATED POWER", "STEP 3: power for H1 (Logo fixation and brand recall), mixed logistic model, by participants and odds ratio.")
cat("wrote SIMULATED_planning_assumptions.csv, SIMULATED_power_analytic.csv, SIMULATED_power_grid.csv, SIMULATED_power_checks.csv, SIMULATED_power_h1.csv\n")
cat("(and in STEP 5 and 6: SIMULATED_latin_square.csv, SIMULATED_order_schedule.csv, SIMULATED_session_schedule.csv)\n\n")
cat("SUMMARY (PLANNED):\n")
cat(sprintf("  Design: %d ads (%d per brand), within subjects, Williams order; %d participants analyzed; %d sessions booked (%d main + %d reserve) + %d open; cap %d\n",
            T_ADS, m_choice, n_choice, n_sessions, n_choice, n_reserve, sum(slots$phase == "open"), n_cap))
cat(sprintf("  P1 brand test on liking: power %.3f (MC SE %.3f) for 0.50 points with ad SD %.2f; Type I error %.3f\n", row_c$power, row_c$mcse, SD_AD, checks$power[1]))
cat(sprintf("  Six ads in all (3 per brand) would give %.3f at 48 participants and %.3f at 72\n", grid$power[grid$ads_per_brand == 3 & grid$n_participants == 48],
            grid$power[grid$ads_per_brand == 3 & grid$n_participants == 72]))
cat(sprintf("  Sessions: %d min participant time, %d min slots, 5 per weekday, %d weekdays (%s to %s) after a pilot day (%s)\n", part_min, slot_min, n_days,
            fdate(min(days), "%B %d"), fdate(max(days), "%B %d, %Y"), fdate(pilot_day, "%B %d")))
cat(sprintf("Total runtime: %.1f min\n", as.numeric(difftime(Sys.time(), t_start, units = "mins"))))
sink()
