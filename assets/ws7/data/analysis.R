# =============================================================================
# analysis.R
# Workshop 7: three simple tests that the MERGED data make possible
# Martin's Family Clothing Consumer Behavior Research Laboratory, JSU
#
# SIMULATED DATA. This script reads analysis_ready_long.csv and analysis_ready_wide.csv, which
# prepare_data.R built from the SIMULATED export bundle (make_simulated_data.py). No person was
# recorded. Each hypothesis needs measures from at least two export families, which is the point
# of Workshop 7: none of the three tests is possible with one export alone.
#
# Analysis sample: include = 1 (rule R2 of the protocol, applied in prepare_data.R). GSR values of a
# stimulus are missing when rule R3 marked them (gsr_valid = 0); such a participant drops out of the
# GSR test for that ad only.
#
# Hypotheses (all tests two tailed, alpha = .05), written before collection:
#   H1 (eye tracking x survey): brand recall is more likely when the Logo was fixated. Each participant
#      saw both ads, so the 44 participant x ad observations are NOT independent (two per person).
#      Primary: Fisher's exact test and the odds ratio on the pooled 2 x 2 table, reported with that
#      caveat; checks: Fisher's test for each ad separately (independent observations), and a mixed
#      effects logistic regression with a random intercept per participant (lme4; Bates et al., 2015).
#   H2 (facial expression x survey): ads that made a participant show more Joy were liked more. Mixed
#      model: liking ~ Joy (per 10 percentage points of time) + ad + (1 | participant) (lmerTest,
#      Satterthwaite df). In plain words: within the same people, does more Joy go with more liking?
#      Checks: Pearson and Spearman correlations for each ad.
#   H3 (GSR): GSR Peaks Per Minute are higher during each ad than during Baseline. Paired t test on the
#      differences, Cohen's d_z with its 95% CI (Lakens, 2013), Wilcoxon signed rank check, Holm
#      adjustment over the two ads.
#
# What the script does, in order:
#   STEP 0.  Packages, paths, helper functions
#   STEP 1.  Read the analysis ready files and check them against the codebook
#   STEP 2.  The analysis sample
#   STEP 3.  Descriptive statistics by stimulus (SIMULATED_descriptives.csv)
#   STEP 4.  H1: Logo fixated and brand recall
#   STEP 5.  H2: Joy and liking (mixed model)
#   STEP 6.  H3: GSR Peaks Per Minute, each ad versus Baseline
#   STEP 7.  Assumption checks (figure)
#   STEP 8.  Overview of the merged measures (figure)
#   STEP 9.  Summary table (SIMULATED_test_results.csv)
#
# How to run: after prepare_data.R, in the same folder:  Rscript analysis.R  (or Source in RStudio
# with the working directory set to this folder). Output: analysis_output.txt.
# Packages: readr, dplyr, tidyr, ggplot2, lme4, lmerTest, grid. If one is missing, run once:
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
options(width = 160)
sink(file.path(script_dir, "analysis_output.txt"), split = TRUE)
cat("SIMULATED Multimodal Ad Study: hypothesis tests on the merged data, run on", format(Sys.time(), "%Y-%m-%d %H:%M"), "\n")
cat("R ", R.version$major, ".", R.version$minor, "; lme4 ", as.character(packageVersion("lme4")), "; lmerTest ",
    as.character(packageVersion("lmerTest")), "\n\n", sep = "")

col_blue <- "#0072B2"; col_verm <- "#D55E00"; col_orange <- "#E69F00"; col_green <- "#009E73"
col_gray <- "#8C8C8C"; col_dark <- "#222222"; col_purple <- "#CC79A7"
stim_cols <- c(Baseline = col_gray, Ad_A = col_blue, Ad_B = col_verm)
theme_ws7 <- function(base_size = 12) {
  theme_minimal(base_size = base_size) +
    theme(panel.grid.minor = element_blank(), panel.grid.major = element_line(color = "#E3E3E3", linewidth = 0.4),
          plot.title = element_text(face = "bold", size = base_size + 1), plot.subtitle = element_text(color = "#555555", size = base_size - 1.5),
          plot.caption = element_text(color = "#777777", size = base_size - 3, hjust = 0), legend.position = "top",
          legend.title = element_blank(), plot.title.position = "plot", plot.caption.position = "plot",
          plot.background = element_rect(fill = "white", color = NA))
}
wrap_txt <- function(x, w) paste(strwrap(x, width = w), collapse = "\n")
pf <- function(p) ifelse(p < 0.0001, sprintf("%.2e", p), sprintf("%.5f", p))         # p for the output file
fp <- function(p) ifelse(p < .001, "p < .001", paste0("p = ", sub("^0", "", sprintf("%.3f", p))))   # p for figures
fv <- function(v) ifelse(abs(v - round(v)) < 1e-9, sprintf("%.0f", v), sprintf("%.1f", v))

# 95% limits of a standardized effect from the NONCENTRAL t distribution (Lakens, 2013): the two
# noncentrality values for which the observed t sits at the 97.5th and the 2.5th percentile,
# divided by sqrt(n) for d_z.
nct_limits <- function(t_obs, df, conf = 0.95) {
  a <- (1 - conf) / 2
  f <- function(ncp, target) suppressWarnings(pt(t_obs, df, ncp)) - target
  c(uniroot(f, c(t_obs - 15, t_obs + 15), target = 1 - a, tol = 1e-12)$root,
    uniroot(f, c(t_obs - 15, t_obs + 15), target = a, tol = 1e-12)$root)
}
results <- list()

# ---- STEP 1. Read the analysis ready files --------------------------------------------------------
# Line 1 of each file is a "# SIMULATED" banner, so skip = 1. NA marks a value that does not exist
# (see the missing_code column of codebook.csv).
cat("---- STEP 1. Files read ----\n")
long <- read_csv("analysis_ready_long.csv", skip = 1, show_col_types = FALSE)
wide <- read_csv("analysis_ready_wide.csv", skip = 1, show_col_types = FALSE)
cb   <- read_csv("codebook.csv", skip = 1, show_col_types = FALSE)
seed_txt <- sub(".*\\(seed ([0-9]+)\\).*", "\\1", readLines("analysis_ready_long.csv", n = 1))
cat(sprintf("Seed of the simulation: %s; long table %d x %d; wide table %d x %d; codebook %d variables\n", seed_txt, nrow(long), ncol(long),
            nrow(wide), ncol(wide), nrow(cb)))
stopifnot(all(names(long) %in% cb$variable), all(names(wide) %in% cb$variable), !any(duplicated(long[c("participant_id", "stimulus")])))
cat("Every column is in the codebook; participant x stimulus is unique.\n\n")

# ---- STEP 2. The analysis sample ----------------------------------------------------------------------
cat("---- STEP 2. Analysis sample ----\n")
ana <- long %>% filter(include == 1) %>% mutate(stimulus = factor(stimulus, levels = c("Baseline", "Ad_A", "Ad_B")))
ads <- ana %>% filter(stimulus != "Baseline") %>% mutate(stimulus = droplevels(stimulus))
anw <- wide %>% filter(include == 1)
n_part <- n_distinct(ana$participant_id)
cat(sprintf("Participants: %d of %d (excluded by rule R2: %s)\n", n_part, nrow(wide),
            paste(sprintf("%s, %s", wide$participant_id[wide$include == 0], wide$exclusion_reason[wide$include == 0]), collapse = "; ")))
cat(sprintf("Order: AB %d, BA %d; gender: %s; age M = %.2f, SD = %.2f (n = %d; one age not entered)\n", sum(anw$order == "AB"),
            sum(anw$order == "BA"), paste(names(table(anw$gender, useNA = "ifany")), table(anw$gender, useNA = "ifany"), sep = " = ", collapse = ", "),
            mean(anw$age, na.rm = TRUE), sd(anw$age, na.rm = TRUE), sum(!is.na(anw$age))))
cat(sprintf("Rows: %d participant x stimulus rows; %d participant x ad rows; GSR valid rows: %d (missing: %s)\n\n", nrow(ana), nrow(ads),
            sum(ana$gsr_valid == 1), paste(paste(ana$participant_id, ana$stimulus)[ana$gsr_valid == 0], collapse = ", ")))

# ---- STEP 3. Descriptive statistics by stimulus ------------------------------------------------------------
cat("---- STEP 3. Descriptive statistics (analysis sample) ----\n")
desc_var <- function(d, v, lab) {
  d %>% group_by(stimulus) %>% summarize(measure = lab, n = sum(!is.na(.data[[v]])), mean = mean(.data[[v]], na.rm = TRUE),
                                         sd = sd(.data[[v]], na.rm = TRUE), median = median(.data[[v]], na.rm = TRUE),
                                         min = min(.data[[v]], na.rm = TRUE), max = max(.data[[v]], na.rm = TRUE), .groups = "drop")
}
desc <- bind_rows(desc_var(ana, "gsr_peaks_per_min", "GSR Peaks Per Minute"), desc_var(ana, "joy_pct", "Joy, percent of time"),
                  desc_var(ana, "engagement_pct", "Engagement, percent of time"), desc_var(ana, "et_quality_pct", "Eye tracking quality (%)"),
                  desc_var(ads, "liking", "Liking (1 to 7)"), desc_var(ads, "logo_fixated", "Logo fixated (proportion)"),
                  desc_var(ads, "recall_correct", "Brand recalled (proportion)"), desc_var(ads, "logo_gaze_dwell_pct", "Logo gaze dwell (% of ad)")) %>%
  relocate(measure)
print(as.data.frame(desc %>% mutate(across(where(is.double), ~ round(.x, 3)))), row.names = FALSE)
cat(sprintf("\nAttitude (mean of ATT_1, ATT_2, recoded ATT_3R): M = %.2f, SD = %.2f; purchase intention: M = %.2f, SD = %.2f (n = %d)\n\n",
            mean(anw$attitude), sd(anw$attitude), mean(anw$purchase_intention), sd(anw$purchase_intention), nrow(anw)))
write_csv(desc, "SIMULATED_descriptives.csv")

# ---- STEP 4. H1: Logo fixated and brand recall ----------------------------------------------------------------
# A 2 x 2 table: rows = Logo fixated (no, yes), columns = brand recalled (no, yes). The odds ratio (OR)
# compares the odds of recall when the Logo was fixated with the odds when it was not; OR = 1 means no
# association. Fisher's exact test is used because some expected counts are small. phi is the
# correlation of the two yes or no variables. Excel's CHISQ.TEST gives the chi square WITHOUT the
# continuity correction, which is printed too.
cat("---- STEP 4. H1: Logo fixated (eye tracking) and brand recall (survey) ----\n")
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
# Clustering check: the same 22 people appear twice. A mixed logistic model gives each participant a
# baseline tendency to recall brands (random intercept), so the Logo effect is estimated within people.
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

# ---- STEP 5. H2: Joy and liking -----------------------------------------------------------------------------------
# Each participant rated both ads, and some people like everything more than others. The mixed model
# gives each participant their own starting level (random intercept), so the Joy slope says: for the
# same kind of person, how much higher is liking when the ad produced 10 more percentage points of
# Joy? The ad is in the model so that the difference between the ads is not mistaken for a Joy effect.
cat("---- STEP 5. H2: percent of time with Joy (AFFDEX) and liking (survey) ----\n")
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
# Standardized slope: the same model on z scores (SD units), comparable with a correlation
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

# ---- STEP 6. H3: GSR Peaks Per Minute, each ad versus Baseline ---------------------------------------------------------
# Paired design: every participant has a Baseline and an ad value, so the test works on the differences
# (ad minus Baseline). d_z = mean difference / SD of the differences. Holm's adjustment keeps the chance
# of at least one false positive over the two tests at .05.
cat("---- STEP 6. H3: GSR Peaks Per Minute, each ad versus Baseline ----\n")
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

# ---- STEP 7. Assumption checks (figure) --------------------------------------------------------------------------------
cat("---- STEP 7. Assumption checks ----\n")
res2 <- tibble(fitted = fitted(m2), resid = resid(m2))
re2 <- tibble(re = ranef(m2)$participant_id[, 1])
qq_df <- function(x, lab) { x <- sort(x); tibble(theory = qnorm(ppoints(length(x))), sample = (x - mean(x)) / sd(x), panel = lab) }
qa <- bind_rows(qq_df(h3a$dd, sprintf("A. H3 differences, Ad_A minus Baseline (n = %d)", h3a$n)),
                qq_df(h3b$dd, sprintf("B. H3 differences, Ad_B minus Baseline (n = %d)", h3b$n)),
                qq_df(res2$resid, "D. H2 model residuals (normal Q-Q)"))
pA <- ggplot(qa %>% filter(!grepl("^D", panel)), aes(theory, sample)) + geom_abline(color = col_gray, linetype = "dashed") +
  geom_point(color = col_blue, size = 2) + facet_wrap(~ panel, ncol = 1) +
  labs(title = "H3: normal Q-Q plots of the paired differences", subtitle = sprintf("Shapiro Wilk: Ad_A %s, Ad_B %s; the Wilcoxon test is the check",
                                                                                   fp(h3a$sw_p), fp(h3b$sw_p)),
       x = "Normal quantiles", y = "Standardized difference") + theme_ws7(10)
pC <- ggplot(res2, aes(fitted, resid)) + geom_hline(yintercept = 0, color = col_gray, linetype = "dashed") +
  geom_point(color = col_verm, size = 2) +
  labs(title = "C. H2: residuals versus fitted liking", subtitle = "Look for a funnel or a curve; neither is visible",
       x = "Fitted liking", y = "Residual") + theme_ws7(10)
pD <- ggplot(qa %>% filter(grepl("^D", panel)), aes(theory, sample)) + geom_abline(color = col_gray, linetype = "dashed") +
  geom_point(color = col_verm, size = 2) + labs(title = "D. H2: normal Q-Q plot of the residuals", x = "Normal quantiles", y = "Standardized residual") +
  theme_ws7(10)
png(file.path(fig_dir, "SIMULATED_assumption_checks.png"), width = 1400, height = 1050, res = 150)
grid.newpage()
pushViewport(viewport(layout = grid.layout(4, 2, heights = unit(c(0.06, 0.47, 0.41, 0.06), "npc"))))
grid.text("SIMULATED: assumption checks for H2 and H3", x = 0.02, hjust = 0, gp = gpar(fontface = "bold", fontsize = 15),
          vp = viewport(layout.pos.row = 1, layout.pos.col = 1:2))
print(pA, vp = viewport(layout.pos.row = 2:3, layout.pos.col = 1))
print(pC, vp = viewport(layout.pos.row = 2, layout.pos.col = 2))
print(pD, vp = viewport(layout.pos.row = 3, layout.pos.col = 2))
grid.text(wrap_txt(sprintf("SIMULATED teaching data (seed %s), not real participants. H1 needs no normal distribution; its smallest expected count is %.2f in the pooled table and %.2f in each per ad table (below 5), so Fisher's exact test is reported throughout. H3 Ad_A differences are skewed (one strong responder), so the Wilcoxon test is the check.", seed_txt, h1$min_exp, h1_a$min_exp), 160),
          x = 0.02, hjust = 0, gp = gpar(fontsize = 8.5, col = "#777777"), vp = viewport(layout.pos.row = 4, layout.pos.col = 1:2))
invisible(dev.off())
cat("saved figure: figures/SIMULATED_assumption_checks.png\n\n")

# ---- STEP 8. Overview of the merged measures (figure) ----------------------------------------------------------------------
cat("---- STEP 8. Overview of the merged measures ----\n")
gl <- ana %>% filter(!is.na(gsr_peaks_per_min))
p4a <- ggplot(gl, aes(stimulus, gsr_peaks_per_min)) +
  geom_line(aes(group = participant_id), color = "#CCCCCC", linewidth = 0.4) +
  geom_point(aes(color = stimulus), size = 1.8, alpha = 0.8, show.legend = FALSE) +
  stat_summary(fun = mean, geom = "point", shape = 23, size = 3.6, fill = "white", color = col_dark) +
  scale_color_manual(values = stim_cols) +
  labs(title = "A. GSR Peaks Per Minute", subtitle = sprintf("GSR export. H3: d_z = %.2f (Ad_A),\n%.2f (Ad_B); diamonds = means", h3a$dz, h3b$dz),
       x = NULL, y = "Peaks per minute") + theme_ws7(10)
p4b <- ggplot(ana, aes(stimulus, joy_pct)) +
  geom_boxplot(aes(fill = stimulus), width = 0.5, outlier.shape = NA, alpha = 0.35, show.legend = FALSE) +
  geom_jitter(aes(color = stimulus), width = 0.12, height = 0, size = 1.6, show.legend = FALSE) +
  scale_fill_manual(values = stim_cols) + scale_color_manual(values = stim_cols) +
  labs(title = "B. Joy, percent of time", subtitle = "FEA export, threshold 50; many zeros", x = NULL, y = "Percent of time") + theme_ws7(10)
rec <- ads %>% group_by(stimulus, logo_fixated) %>% summarize(n = n(), k = sum(recall_correct), .groups = "drop") %>%
  mutate(p = k / n, fix = factor(logo_fixated, levels = c(0, 1), labels = c("Not fixated", "Fixated")),
         lo = qbeta(0.025, k + 0.5, n - k + 0.5), hi = qbeta(0.975, k + 0.5, n - k + 0.5))
p4c <- ggplot(rec, aes(fix, 100 * p, fill = stimulus)) +
  geom_col(position = position_dodge(width = 0.75), width = 0.7) +
  geom_errorbar(aes(ymin = 100 * lo, ymax = 100 * hi), position = position_dodge(width = 0.75), width = 0.15, color = col_dark) +
  geom_text(aes(y = 4, label = sprintf("%d/%d", k, n)), position = position_dodge(width = 0.75), color = "white", size = 3, vjust = 0) +
  scale_fill_manual(values = stim_cols[c("Ad_A", "Ad_B")]) + scale_y_continuous(limits = c(0, 100)) +
  labs(title = "C. Recall by Logo fixation", subtitle = sprintf("AOI x survey. H1 pooled: OR = %.2f, %s", h1$or, fp(h1$p_fisher)),
       x = "Logo AOI", y = "Recalled (%)") + theme_ws7(10) + theme(legend.position = c(0.2, 0.92), legend.direction = "horizontal")
p4d <- ggplot(ads2, aes(joy_pct, liking, color = stimulus, shape = stimulus)) +
  geom_jitter(width = 0, height = 0.12, size = 2) +
  geom_abline(intercept = fixef(m2)[1], slope = fixef(m2)[2] / 10, color = col_blue, linewidth = 0.8) +
  geom_abline(intercept = fixef(m2)[1] + fixef(m2)[3], slope = fixef(m2)[2] / 10, color = col_verm, linewidth = 0.8, linetype = "dashed") +
  scale_color_manual(values = stim_cols[c("Ad_A", "Ad_B")]) + scale_y_continuous(breaks = 1:7, limits = c(0.7, 7.3)) +
  labs(title = "D. Liking by Joy (FEA x survey)", subtitle = sprintf("H2: b = %.2f per 10 points, %s;\nlines = mixed model (solid Ad_A)", s2["joy10", "Estimate"],
                                                                     fp(s2["joy10", "Pr(>|t|)"])),
       x = "Joy, percent of time", y = "Liking (1 to 7)") + theme_ws7(10) + theme(legend.position = "none")
qual <- long %>% filter(stimulus != "Baseline") %>% group_by(participant_id, include) %>%
  summarize(et = min(et_quality_pct), face = min(face_quality_pct), .groups = "drop") %>%
  mutate(status = if_else(include == 1, "Included", "Excluded by rule R2"))
p4e <- ggplot(qual, aes(et, face, color = status, shape = status)) +
  annotate("rect", xmin = -Inf, xmax = 90, ymin = -Inf, ymax = Inf, fill = col_orange, alpha = 0.10) +
  annotate("rect", xmin = -Inf, xmax = Inf, ymin = -Inf, ymax = 80, fill = col_orange, alpha = 0.10) +
  geom_vline(xintercept = 90, linetype = "dashed", color = col_gray) + geom_hline(yintercept = 80, linetype = "dashed", color = col_gray) +
  geom_point(size = 2.4) + geom_text(data = qual %>% filter(include == 0), aes(label = participant_id), nudge_x = 1.6, size = 3, show.legend = FALSE) +
  scale_color_manual(values = c("Included" = col_green, "Excluded by rule R2" = col_purple)) +
  labs(title = "E. Data quality, all 24", subtitle = "_DataQuality.txt, lower of the two ads;\nshaded = fails rule R2",
       x = "Eye tracking quality (%)", y = "Face found (%)") + theme_ws7(10) + theme(legend.position = c(0.3, 0.15))
p4f <- ggplot(anw, aes(attitude, purchase_intention)) +
  geom_jitter(width = 0.03, height = 0.12, color = col_purple, size = 2) +
  labs(title = "F. Attitude and intention", subtitle = sprintf("Survey; attitude = mean of 3 items,\nATT_3R recoded; r = %.2f",
                                                                              cor(anw$attitude, anw$purchase_intention)),
       x = "Attitude (1 to 7)", y = "Purchase intention (1 to 7)") + theme_ws7(10)
png(file.path(fig_dir, "SIMULATED_fig4_merged_overview.png"), width = 1400, height = 1300, res = 150)
grid.newpage()
pushViewport(viewport(layout = grid.layout(5, 3, heights = unit(c(0.05, 0.43, 0.04, 0.43, 0.05), "npc"))))
grid.text("SIMULATED: the merged data set at a glance (analysis sample unless noted)", x = 0.02, hjust = 0, gp = gpar(fontface = "bold", fontsize = 15),
          vp = viewport(layout.pos.row = 1, layout.pos.col = 1:3))
print(p4a, vp = viewport(layout.pos.row = 2, layout.pos.col = 1)); print(p4b, vp = viewport(layout.pos.row = 2, layout.pos.col = 2))
print(p4c, vp = viewport(layout.pos.row = 2, layout.pos.col = 3)); print(p4d, vp = viewport(layout.pos.row = 4, layout.pos.col = 1))
print(p4e, vp = viewport(layout.pos.row = 4, layout.pos.col = 2)); print(p4f, vp = viewport(layout.pos.row = 4, layout.pos.col = 3))
grid.text(wrap_txt(sprintf("SIMULATED teaching data (seed %s), not real participants. n = %d participants (panel E: all 24). Panels C and D join two export families (AOI or FEA with the survey) through participant_id and stimulus; panel C bars: 95%% intervals (Jeffreys).", seed_txt, n_part), 160),
          x = 0.02, hjust = 0, gp = gpar(fontsize = 8.5, col = "#777777"), vp = viewport(layout.pos.row = 5, layout.pos.col = 1:3))
invisible(dev.off())
cat("saved figure: figures/SIMULATED_fig4_merged_overview.png\n\n")

# ---- STEP 9. Summary table ------------------------------------------------------------------------------------------------
cat("---- STEP 9. Summary of the tests (SIMULATED_test_results.csv) ----\n")
tr <- bind_rows(results)
print(as.data.frame(tr %>% mutate(across(where(is.double), ~ signif(.x, 4)))), row.names = FALSE)
write_csv(tr, "SIMULATED_test_results.csv")
cat("\nwrote SIMULATED_test_results.csv and SIMULATED_descriptives.csv\n")
cat("References: Bates et al. (2015); Lakens (2013). See refs/references.json.\n")
sink()
