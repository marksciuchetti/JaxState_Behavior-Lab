# Analysis plan (preregistration style): the follow up ad study for ALTAVO and BRENLOW

**PLANNED STUDY, SIMULATED PLANNING DATA.** This is a teaching model of a preregistration for the follow up study that Workshop 8 called for. No data exist; every power value comes from simulated data sets (`plan_power.R`, seed 20270208; full log in `plan_output.txt`), and every planning value comes from the SIMULATED Multimodal Ad Study of Workshops 7 and 8 or is labeled as a judgment. ALTAVO and BRENLOW are invented brands; the client is fictional; participant codes are placeholders. The plan follows the general structure of common preregistration templates (study information, design, sampling, variables, analysis plan). Before a real study, the plan must be completed (names, dates, the approved IRB number), dated, and frozen before the first participant is recorded, and it must agree with the protocol the JSU IRB approves.

**Version:** [version number, date] (frozen before data collection). **Authors:** [names]. **IRB protocol number:** [number].

---

## 1. Study information

### 1.1 Title

Brand and execution effects of video advertising on attention, facial expression, arousal, and self report: a within subjects lab study with ads as a random factor [working title].

### 1.2 Background and continuity

Workshop 7 recorded 24 participants (22 analyzed) who saw one 30 s video ad for each of two fictional water bottle brands, ALTAVO and BRENLOW, with the Smart Eye Aurora (60 Hz), AFFDEX, and a Shimmer3 GSR+. Workshop 8 found that the two ads did not differ detectably on any measure and that the design could not tell whether they differ: with one ad per brand, any difference mixes the brand with the particular execution, and 22 paired observations detect only d_z of 0.63 or more. Workshop 8 asked for a next study with several ads per brand, purchase intention asked for each ad, a smallest effect of interest named in advance, and a preregistration (ws8/data/results_text.md, "What the next study needs"). This plan is that study.

### 1.3 Research questions and hypotheses

All tests are two tailed. The direction stated is the direction expected from Workshop 8, which is weak evidence (an interval that includes zero), so the tests do not assume it.

**Primary hypothesis (confirmatory).**

- **P1.** Liking of the ads differs between the brands when ads are treated as a sample of each brand's possible ads (expected: ALTAVO higher; Workshop 8 estimate 0.55 points, 95% CI [-0.13, 1.23], which mixes brand and execution). The smallest effect of interest is 0.50 points on the 1 to 7 scale.

**Secondary hypotheses (confirmatory, one family, Holm correction over three).** These compare the brands on the twelve ads that were shown (ads treated as fixed), which is what the client report needs:

- **S1.** Purchase intention, asked after each ad, differs between the ALTAVO ads and the BRENLOW ads.
- **S2.** The percent of time with Joy (AFFDEX, threshold [50, to be confirmed in the pilot]) differs between the ALTAVO ads and the BRENLOW ads.
- **S3.** GSR peaks per minute differ between the ALTAVO ads and the BRENLOW ads.

**Replication hypotheses (confirmatory; each its own family, as in Workshop 7).**

- **R1 (Workshop 7 H1).** Brand recall for an ad is more likely when the Logo of that ad was fixated.
- **R2 (Workshop 7 H2).** Ads that produce more Joy in a participant are liked more by that participant.
- **R3 (Workshop 7 H3).** GSR peaks per minute are higher during the ads than during the Baseline.

## 2. Design plan

### 2.1 Study type

Experiment in the lab. Every participant sees every ad (within subjects). Brand (ALTAVO, BRENLOW) is a within subjects factor; ad is a random factor nested in brand (6 ads per brand, 12 ads in all); position (1 to 12) is counterbalanced.

### 2.2 Blinding

Participants are told that the study is about reactions to water bottle ads; they are not told the hypotheses. Brand names are visible in the ads, as in real advertising. The operators know the brands; [decide whether the operators can be kept unaware of the expected direction, and if not, say so in the report]. The analyst runs the confirmatory analyses with the code written before data collection (section 5.8).

### 2.3 Stimuli

Twelve 30 s video ads supplied by the client: six for ALTAVO and six for BRENLOW, each brand's six from one campaign, with the same layout so that three areas of interest (AOIs) sit in the same place in every ad: Logo (top left), Product (center), Headline (bottom), as in Workshop 7. Stimuli are equalized for length and resolution. [Client to supply; the pilot checks that each AOI is visible for the whole ad.]

### 2.4 Procedure

One participant at a time at the single Smart Eye Aurora seat (lab procedure manual, section 7.1): consent and screening (about 15 min), Shimmer3 GSR+ on two fingers and Aurora calibration (about 10 min), recording (about 20 min: instructions, a 60 s Baseline with a fixation cross, the 12 ads in the participant's order, two questions after each ad, end questions), sensors off and debriefing (about 15 min). Recording uses iMotions Lab 11.1 on [the computer and license key used]. The session schedule is in `SIMULATED_session_schedule.csv`.

### 2.5 Randomization and counterbalancing

The order of the 12 ads follows a Williams design (Williams, 1949): a 12 x 12 Latin square in which every ad appears once in every position and every ad immediately follows every other ad exactly once (`SIMULATED_latin_square.csv`). Ads are labeled so that no row shows more than two ads of one brand in a row (brand pattern ABBA... or BAAB...). Participants P01 to P48 are assigned to rows in 4 blocks of 12, each block using every row once in an order drawn with a fixed seed (`SIMULATED_order_schedule.csv`). A participant excluded under rule R2 is replaced by the next reserve code, who receives the same row.

## 3. Sampling plan

### 3.1 Existing data

None. This plan is written and dated before any data are collected. The two pilot sessions (codes X01 and X02) test the procedure and the exports; their outcome data are not analyzed for any hypothesis and are not part of the sample.

### 3.2 Population and recruitment

[Population and recruitment channel as approved by the JSU IRB]. Inclusion: adults ([18 or older]; the age the IRB approves), normal or corrected to normal vision. Not eligible after screening (lab procedure manual, section 2.2): a skin allergy to electrode gel, adhesive, or alcohol wipes; a history of seizures or sensitivity to flashing light; a skin condition on the fingers. [Compensation as approved.]

### 3.3 Sample size

**48 participants who meet the inclusion and quality rules**, each seeing all 12 ads (576 ratings of liking).

### 3.4 Sample size rationale

The rationale is in `plan_output.txt` (STEPS 1 to 4) and summarized here; the planning values and their sources are in `SIMULATED_planning_assumptions.csv`.

- **Effect to detect.** The smallest effect of interest for P1 is a brand difference of 0.50 points in liking. Planning on the Workshop 8 estimate was rejected because that estimate mixes brand and execution and comes from 22 people (Lakens, 2022).
- **Planning values.** Participant SD 0.80 and residual SD 1.09 (fitted on the Workshop 7 data); ad SD within a brand 0.20 (a judgment: no lab estimate exists).
- **Method.** Power by simulation of the planned model, 500 simulated studies per cell, over 24 to 72 participants and 2, 3, 4, or 6 ads per brand (Westfall et al., 2014, explain why both numbers matter). The number of ads limited power more than the number of participants: with 3 ads per brand power was .47 at 48 participants and .51 at 72; with 4 ads per brand it never reached .80; with 6 ads per brand it was .88 at 48 participants (Monte Carlo SE .015). The simulated Type I error of the planned test at this design was .052.
- **Choice.** Six ads per brand is the smallest number that reaches 80 percent power. Forty eight is the smallest number of participants in the grid that reaches 80 percent and also gives 80 percent power for the secondary family (45 participants for d_z = 0.50 at the strictest Holm step of three; `power.t.test`, as in Workshop 8), and it is a multiple of the 12 rows of the Williams design.
- **R1.** With 12 observations per participant, simulated power for R1 at 48 participants was .975 for an odds ratio of 2.0 and 1.000 for the Workshop 8 estimate of 4.34 (200 simulated studies each).
- **R3.** The Workshop 7 effect (d_z = 0.83) needs 14 participants for 80 percent power.
- **R2.** Power was not simulated. Workshop 8 estimated that a between person correlation of .24 needs about 131 people; the within person test here uses 12 observations per person and should be more sensitive, but a nonsignificant R2 will be reported as inconclusive, not as evidence of no relation.
- **Risk named in advance.** If the ads of one brand differ more than assumed (ad SD 0.30), power for P1 falls to about .64; if they differ less (0.10), it rises to about .98.

### 3.5 Stopping rule

Data collection stops when 48 participants meet rules R1 and R2 (section 6), or after 60 sessions (25 percent above the target), whichever comes first. Based on the Workshop 7 exclusion rate (2 of 24), 53 sessions are expected; 48 main and 5 reserve sessions are booked, and 2 open slots absorb no shows. Outcome variables are not analyzed before the stopping rule is met; only the quality checks of rules R1 to R3 are run after each session. If more than 3 participants of a block of 12 fail rule R2, collection pauses to look for a procedural cause.

## 4. Variables

### 4.1 Manipulated variables

- **Brand:** ALTAVO, BRENLOW (coded +0.5 and -0.5 in the models, so a coefficient is the ALTAVO minus BRENLOW difference).
- **Ad:** 12 ads, 6 nested in each brand (a random factor).
- **Position:** 1 to 12 (counterbalanced; used only in checks).

### 4.2 Measured variables

| Variable | Definition | Source in iMotions | Level |
|---|---|---|---|
| Liking | "How much did you like this ad?" 1 (not at all) to 7 (very much), after each ad | Survey slide after each ad | Participant x ad |
| Purchase intention | "How likely would you be to buy a water bottle from this brand?" 1 to 7, after each ad [wording to finalize before the IRB submission] | Survey slide after each ad | Participant x ad |
| Brand recall | For each ad, "Which brand was advertised in the ad with [scene]?" ALTAVO, BRENLOW, two distractor brands, "I do not remember"; scored correct or not | End questions | Participant x ad |
| Logo fixated | At least one fixation in the Logo AOI during the ad | AOI Metrics export (Workshop 7 layout, modeled) | Participant x ad |
| Dwell on Logo, Product, Headline | Gaze dwell time as percent of the 30 s ad | AOI Metrics export | Participant x ad x AOI |
| Joy, Engagement, Brow Furrow | Percent of time at or above the threshold | FEA Metrics export (R notebook) | Participant x ad |
| GSR peaks per minute | Peak count divided by valid signal minutes | GSR Peak Metrics export (R notebook) | Participant x stimulus (Baseline and each ad) |
| Attitude toward reusable water bottles | Three items, 1 to 7 (as in Workshop 7) | End questions | Participant |
| Age, gender | As approved | End questions | Participant |
| Data quality | Eye tracking quality and face quality percent; GSR Signal Duration | Data quality export; GSR export | Participant x stimulus |

### 4.3 Indices

For S1 to S3: each participant's mean over the six ALTAVO ads minus the mean over the six BRENLOW ads (only valid ads count; at least three valid ads per brand are required). For R3: each participant's mean GSR peaks per minute over the valid ads minus the Baseline value. Joy is entered in R2 per 10 percentage points, as in Workshop 7.

## 5. Analysis plan

All analyses use R (version recorded at analysis) with lme4 (Bates et al., 2015) and lmerTest (Kuznetsova et al., 2017). Alpha is .05, two tailed.

### 5.1 P1 (primary): brand effect on liking, ads as a random factor

**Model:** `liking ~ brand + (1 | participant) + (1 | ad)`, REML. The brand coefficient is the ALTAVO minus BRENLOW difference.

**Test and decision rule.** Because ads are a random factor, the brand difference is judged against the variation between ads (Clark, 1973; Judd et al., 2012). The coefficient is tested with Satterthwaite degrees of freedom (lmerTest), which perform acceptably for mixed models of this kind (Luke, 2017) and which, for this complete, balanced design, equal the number of ads minus 2 (10); the result is then identical to a two sample t test of the six ALTAVO ad means against the six BRENLOW ad means (verified on simulated studies in `plan_output.txt`, STEP 2). When the ad variance is estimated at or near zero (a singular fit, or an estimated ad SD below 1 percent of the residual SD), the Satterthwaite df become unreliable and can jump to several hundred, which makes the test too liberal; **the decision is then based on the t test of the ad means (df = 10)**. Both outputs are reported. If ratings are missing and the design is no longer balanced, the Satterthwaite test of the model is used, with the same rule for a singular fit. Report: the estimate, its 95% CI (t with the df used), t, df, p, the variance components, and the twelve ad means.

**Equivalence.** If P1 is not significant, the two one sided tests procedure against plus or minus 0.50 points is applied (Lakens et al., 2018): the brands are declared equivalent within the smallest effect of interest if the 90% CI of the brand difference lies inside [-0.50, 0.50]. Simulated chance of concluding equivalence when the brands truly do not differ: 88 percent (check C1 in `plan_output.txt`).

**Robustness (planned, not decisive).** (a) Adding a participant by brand random slope, `(1 + brand | participant)`, the maximal model for this design (Barr et al., 2013), reported if it converges without a singular fit. In the simulation, ignoring such a slope (SD 0.50) raised the Type I error of the planned test only to .060 (check C4). (b) Adding position (1 to 12, centered) as a fixed covariate.

### 5.2 S1 to S3 (secondary family): brand comparisons on these twelve ads

Paired t tests of the participant level indices (section 4.3): purchase intention (S1), percent of time with Joy (S2), GSR peaks per minute (S3). Holm's (1979) correction over the three tests. Report the mean difference with its 95% CI and d_z with its 95% CI from the noncentral t distribution (Lakens, 2013), and a Wilcoxon signed rank test as a check, because percent of time measures pile up at zero (Workshop 8). These tests treat the ads as fixed: they describe the twelve ads shown and do not generalize to other ads of the brands.

### 5.3 R1: Logo fixation and brand recall

`recall ~ logo_fixated + (1 | participant) + (1 | ad)`, binomial, Laplace approximation (glmer). Wald z test of the logo_fixated coefficient; report the odds ratio with its 95% CI. A singular fit (participant or ad variance at zero, as in Workshop 7) is reported and the model is kept. Check: the pooled Fisher exact test, as in Workshop 7, labeled as ignoring the nesting.

### 5.4 R2: Joy and liking

`liking ~ joy10 + brand + (1 | participant) + (1 | ad)`, REML, Satterthwaite df; joy10 = percent of time with Joy divided by 10. Report the slope per 10 points of Joy with its 95% CI. Check: the same model with Joy split into each participant's mean (between) and the deviation from it (within), to show which part carries the relation.

### 5.5 R3: GSR during the ads versus Baseline

Paired t test of the participant level index (section 4.3), d_z with its 95% CI, Wilcoxon check.

### 5.6 Inference criteria

P1, R1, R2, and R3 are each tested at .05. S1 to S3 are tested with Holm's correction at a family wise .05. No other correction is applied across families; the number of families is stated in the report. Every planned test is reported whatever its outcome, with an effect size and a 95% CI. A nonsignificant result is described as "not detected" unless the equivalence test supports equivalence (P1 only).

### 5.7 Missing data

No imputation. A missing rating or a missing sensor value for one ad drops that observation from the analyses that use it (rules R2b and R3); the mixed models use all remaining observations. For S1 to S3 and R3, a participant needs at least three valid ads per brand (S) or six valid ads and a valid Baseline (R3).

### 5.8 Code written in advance

The confirmatory code is written and run on simulated data before data collection (`plan_power.R` contains the model and the decision rule for P1) and is not changed afterward except to read the real exports (Workshop 7 pipeline). Any change is reported as a deviation (section 7).

## 6. Data exclusion rules (written in advance)

The rules adapt rules R1 to R3 of Workshop 7 (ws7/data/README.md) to twelve ads and use the lab's data quality thresholds (lab procedure manual, Appendix D: eye tracking above 90 percent valid gaze samples, GSR artifacts below 5 percent) and the Workshop 3 face rule.

| Rule | Level | Rule | Applies to |
|---|---|---|---|
| R1 | Recording | Use only recordings that the session log marks Use recording = Yes. An aborted recording is rerun under the same participant code with a new recording number. | All analyses |
| R2 | Participant | Exclude a participant from ALL analyses when the session ended before the last ad, or eye tracking quality is below 90 percent on 3 or more of the 12 ads, or AFFDEX face quality is below 80 percent on 3 or more of the 12 ads. The excluded participant is replaced (section 3.5). | All analyses (one sample for every hypothesis, as in Workshop 7) |
| R2b | Observation | For a participant kept under R2, an ad with eye tracking quality below 90 percent is missing for the eye tracking measures, and an ad with face quality below 80 percent is missing for the facial expression measures. | R1, R2, S2, exploratory AOI analyses |
| R3 | Observation | GSR for a stimulus is missing when Signal Duration (ms) is below 80 percent of Duration (ms). A participant without valid Baseline GSR, or with valid GSR on fewer than 6 of the 12 ads, is left out of R3; S3 needs at least 3 valid ads per brand. | S3, R3 |

Every exclusion is logged with its rule in the cleaning log (Workshop 7 format), and the flow of participants (recorded, excluded by rule, replaced, analyzed) is reported.

## 7. Deviations

Any deviation from this plan (a changed rule, model, or measure) is reported in a table with the reason and the result under the planned analysis, and any change to the procedure goes to the JSU IRB as an amendment before it is used (lab procedure manual, section 2.1).

## 8. Exploratory analyses (labeled exploratory in advance; no confirmatory claims)

- E1. Dwell on Logo, Product, and Headline by brand and AOI (mixed model with participants and ads random), following Workshop 8 Q4.
- E2. Engagement and Brow Furrow by brand.
- E3. Position and carryover: liking and GSR by position (1 to 12) and by the brand of the preceding ad.
- E4. Ad level description for the client: the mean and 95% CI of liking, purchase intention, Joy, and recall for each of the twelve ads, without tests.
- E5. Cross modality correlations at the participant level with 95% CIs and without p values (Workshop 8 E1).
- E6. Sensitivity of S2 to the AFFDEX threshold (for example 25 and 75 in addition to the planned value).

## 9. Items to verify before the plan is frozen

1. The JSU IRB template, review level, and timelines; the IRB number for the plan header.
2. The final wording of the liking, purchase intention, and recall questions, and the two distractor brand names.
3. The AFFDEX threshold and the GSR peak detection notebook parameters (Workshop 8, items to verify).
4. The fixation filter and its parameters for the Aurora in iMotions 11.1 (Project doc 02a).
5. How the Williams order is implemented in iMotions (study flows or a test plan) and that the exported stimulus names match the ad labels.
6. That the client can supply six comparable ads per brand; with fewer, the brand claim of P1 is underpowered (`plan_output.txt`, STEP 2), and the plan must be revised before the IRB submission.

## Sources

- Barr, D. J., Levy, R., Scheepers, C., & Tily, H. J. (2013). Random effects structure for confirmatory hypothesis testing: Keep it maximal. *Journal of Memory and Language, 68*(3), 255-278. https://doi.org/10.1016/j.jml.2012.11.001
- Bates, D., Mächler, M., Bolker, B., & Walker, S. (2015). Fitting linear mixed-effects models using lme4. *Journal of Statistical Software, 67*(1), 1-48. https://doi.org/10.18637/jss.v067.i01
- Clark, H. H. (1973). The language-as-fixed-effect fallacy: A critique of language statistics in psychological research. *Journal of Verbal Learning and Verbal Behavior, 12*(4), 335-359. https://doi.org/10.1016/S0022-5371(73)80014-3
- Holm, S. (1979). A simple sequentially rejective multiple test procedure. *Scandinavian Journal of Statistics, 6*(2), 65-70.
- Judd, C. M., Westfall, J., & Kenny, D. A. (2012). Treating stimuli as a random factor in social psychology: A new and comprehensive solution to a pervasive but largely ignored problem. *Journal of Personality and Social Psychology, 103*(1), 54-69. https://doi.org/10.1037/a0028347
- Kuznetsova, A., Brockhoff, P. B., & Christensen, R. H. B. (2017). lmerTest package: Tests in linear mixed effects models. *Journal of Statistical Software, 82*(13), 1-26. https://doi.org/10.18637/jss.v082.i13
- Lakens, D. (2013). Calculating and reporting effect sizes to facilitate cumulative science: A practical primer for t-tests and ANOVAs. *Frontiers in Psychology, 4*, Article 863. https://doi.org/10.3389/fpsyg.2013.00863
- Lakens, D. (2022). Sample size justification. *Collabra: Psychology, 8*(1), Article 33267. https://doi.org/10.1525/collabra.33267
- Lakens, D., Scheel, A. M., & Isager, P. M. (2018). Equivalence testing for psychological research: A tutorial. *Advances in Methods and Practices in Psychological Science, 1*(2), 259-269. https://doi.org/10.1177/2515245918770963
- Luke, S. G. (2017). Evaluating significance in linear mixed-effects models in R. *Behavior Research Methods, 49*(4), 1494-1502. https://doi.org/10.3758/s13428-016-0809-y
- Nosek, B. A., Ebersole, C. R., DeHaven, A. C., & Mellor, D. T. (2018). The preregistration revolution. *Proceedings of the National Academy of Sciences, 115*(11), 2600-2606. https://doi.org/10.1073/pnas.1708274114
- Westfall, J., Kenny, D. A., & Judd, C. M. (2014). Statistical power and optimal design in experiments in which samples of participants respond to samples of stimuli. *Journal of Experimental Psychology: General, 143*(5), 2020-2045. https://doi.org/10.1037/xge0000014
- Williams, E. J. (1949). Experimental designs balanced for the estimation of residual effects of treatments. *Australian Journal of Scientific Research, Series A, 2*(2), 149-168. https://doi.org/10.1071/CH9490149
