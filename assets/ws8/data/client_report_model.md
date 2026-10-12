# SIMULATED Ad Test Report: the ALTAVO "Hiking" Ad versus the BRENLOW "Cycling" Ad

**SIMULATED MODEL REPORT FOR TEACHING.** This report was written from SIMULATED data (seed 20271231) for Workshop 8. No real person was recorded, ALTAVO and BRENLOW are invented brands, and the "client" is fictional. Every number below was copied from `analysis_output.txt` (produced by `analysis_full.R`); none may be quoted as a finding about real advertising. The structure follows the client reports that iMotions publishes as sample reports (project summary, executive summary, objectives, method, results, conclusions and recommendations); those samples show format only and are not evidence about advertising effects.

| | |
|---|---|
| **Prepared for** | ALTAVO brand team (fictional client) |
| **Prepared by** | Martin's Family Clothing Consumer Behavior Research Laboratory, Jacksonville State University (teaching model) |
| **Study dates** | November 16 to 20, 2026 (simulated) |
| **Test material** | Two 30-second video ads: ALTAVO "Hiking" (Ad_A) and BRENLOW "Cycling" (Ad_B, the comparison ad) |
| **Viewers** | 24 recruited, 22 analyzed (2 removed for poor eye tracking or face video quality) |
| **Measures** | Eye tracking (where viewers looked), facial expression analysis, skin conductance (arousal), short survey |

---

## Executive summary

1. **Both ads got viewers' bodies to respond.** Skin conductance responses, a standard sign of arousal, rose from about 2 per minute at rest to 7.3 per minute during the ALTAVO ad and 6.0 during the BRENLOW ad; 19 of 22 viewers responded more to the ALTAVO ad than at rest.
2. **The ALTAVO logo is almost never seen.** Viewers spent 1.8 percent of the ad looking at the logo, and 10 of 22 never looked at it at all. This matters because viewers who did look at the logo named the right brand afterward far more often (71 percent) than viewers who did not (35 percent).
3. **We cannot tell which ad performs better.** ALTAVO's ad scored somewhat higher on smiling (Joy) and liking, but the differences are small enough that chance alone could produce them with 22 viewers. Brand recall was identical: 12 of 22 viewers (55 percent) named the right brand for each ad.

**Recommendation.** Do not choose between the two ads on the basis of this study. Instead, make the ALTAVO logo easier to see (larger, earlier, and on screen longer), and test the revised ad against the current one with at least 34 viewers.

---

## Background and objectives

The ALTAVO brand team wanted to know how its new hiking ad performs against a competitor's cycling ad on five questions: Do viewers look at the brand? Does the ad produce positive emotion? Does it get a physical reaction? Do viewers like it? Do they remember the brand? The team also asked about purchase intention.

## What we did

Twenty-four university students came to the lab one at a time. Each sat in front of a computer screen and watched a one minute rest screen, then both ads, half of them ALTAVO first and half BRENLOW first, so that the order could not favor either ad. While they watched, a screen based eye tracker recorded where they looked, a webcam with facial expression software recorded smiles and frowns, and two finger sensors recorded skin conductance. Afterward, each viewer answered a short survey about the ads. Two viewers were removed before analysis because their eye tracking or face video was incomplete, which left 22 viewers (21 for skin conductance during the BRENLOW ad, where one sensor recording failed). Each ad was compared with itself at rest and with the other ad, and we checked every difference with a statistical test that accounts for the same people seeing both ads. Details are in the appendix.

---

## Key findings

### Finding 1. At a glance: no clear winner (Figure: `figures/SIMULATED_fig5_client_dashboard.png`)

The one page scorecard compares the ads on six measures. The dot is the average; the bar shows the range in which the average for similar viewers plausibly lies.

| Measure | ALTAVO (Ad_A) | BRENLOW (Ad_B) | Can we tell them apart? |
|---|---|---|---|
| Time looking at the logo | 1.8% of the ad | 1.4% | No (difference 0.4 points) |
| Smiling (Joy), share of time | 12.1% | 8.3% | No (difference 3.8 points) |
| Arousal above rest (responses per minute) | 5.2 | 4.1 | No |
| Liking (1 to 7) | 4.6 | 4.0 | No (difference 0.5 points) |
| Liked it a lot (6 or 7) | 5 of 22 (22.7%) | 3 of 22 (13.6%) | Not tested |
| Named the right brand | 12 of 22 (54.5%) | 12 of 22 (54.5%) | No (identical) |
| Purchase intention (1 to 7) | 3.7, asked once about reusable bottles in general | | Cannot be compared (see below) |

ALTAVO's averages are higher on smiling and liking. A difference of this size could easily appear by chance with 22 viewers, so we cannot say that ALTAVO's ad is better, and we also cannot say that the two ads are equal. A larger study is needed to settle it.

### Finding 2. Viewers watch the product, not the brand (Figure: `figures/SIMULATED_fig1_dwell_by_ad_aoi.png`)

In both ads, the product held about 41 to 43 percent of viewing time and the headline 18 to 20 percent. The logo held under 2 percent in both ads. Ten of 22 viewers never looked at the ALTAVO logo, and those who did found it late: on average 11.8 seconds (11,849 milliseconds) into the 30-second ad. The pattern of attention was the same for both ads; neither ad pulled more attention to its logo, product, or headline than the other.

### Finding 3. Seeing the logo goes with remembering the brand (Figure: `../../ws7/data/figures/SIMULATED_fig4_merged_overview.png`, panel C, from Workshop 7)

Across both ads, viewers who looked at the logo at least once named the right brand 71 percent of the time (17 of 24 cases), against 35 percent (7 of 20 cases) when they had not looked at it. This difference is larger than chance would produce. It shows that logo attention and brand memory go together; it does not prove that a bigger logo would raise recall, which is why we recommend testing a revised ad rather than assuming the effect.

### Finding 4. Both ads produce a physical response; neither clearly more than the other (Figure: `figures/SIMULATED_fig3_gsr_change.png`)

At rest, viewers showed about 2.1 skin conductance responses per minute. During the ALTAVO ad this rose to 7.3 per minute (19 of 22 viewers increased) and during the BRENLOW ad to 6.0 per minute (16 of 21 increased). Both increases are clear. Comparing the two ads directly, ALTAVO was ahead by about 1.0 response per minute, which is within what chance can produce. Skin conductance shows how strongly the body reacts, not whether the reaction is positive or negative.

### Finding 5. Smiling and frowning were similar for both ads (Figure: `figures/SIMULATED_fig2_expressions_by_ad.png`)

Viewers smiled during 12.1 percent of the ALTAVO ad and 8.3 percent of the BRENLOW ad; 12 viewers smiled more during ALTAVO's ad and 7 during BRENLOW's. Frowning (brow furrow) was almost absent: 17 of 22 viewers did not frown during either ad. None of these differences can be told apart from chance.

---

## What the results do not show

- **They do not show that ALTAVO's ad is better or worse than BRENLOW's.** Every head to head difference was within the range that chance can produce with 22 viewers.
- **They do not show an effect on purchase intention.** The survey asked once, after both ads, how likely viewers were to buy any reusable water bottle in the next three months (average 3.7 of 7; only 1 of 22 chose 6 or 7). The question did not name a brand and was not asked after each ad, so it cannot separate the two ads.
- **They do not show what viewers felt.** The software detects facial movements such as a smile; a smile is not proof of joy, and skin conductance measures intensity, not whether a feeling was good or bad.
- **They do not describe the market.** The viewers were 22 university students, and each brand was represented by one ad.
- **They do not prove that the logo causes recall.** Finding 3 is an association; a test of a revised ad can show cause and effect.

## Recommendations

1. **Revise the ALTAVO ad so that the logo is easy to find:** larger, on screen earlier than about 12 seconds, and held longer, for example on the end card.
2. **Test the revised ad against the current one** in the same lab setup with at least 34 viewers, which gives a good chance (80 percent) of detecting a moderate difference on one measure; 45 viewers if three measures are judged together, and about 90 to detect a small difference.
3. **Ask purchase intention for each ad and name the brand** ("How likely are you to buy an ALTAVO bottle?"), so that the ads can be compared on the outcome that matters most.
4. **Agree on the smallest difference that would change a decision before the next study** (for example, half a point of liking), so that the next report can say either "different" or "equal for practical purposes".

---

## Appendix: methods

**Sample.** 24 university students (simulated), tested individually; 22 analyzed after a data quality rule set before the study (eye tracking at least 90 percent valid and the face found in at least 80 percent of frames during both ads). For skin conductance, a recording counted only if at least 80 percent of the signal was present, which removed one viewer's BRENLOW recording.

**Procedure.** Instructions, a 60-second rest screen with a fixation cross, the two 30-second ads in counterbalanced order (11 viewers saw ALTAVO first, 11 BRENLOW first), then a survey. A check found no effect of order or position on liking or viewing time once the eight checks were corrected for multiple testing (one check, headline viewing time by order group, looked different before that correction).

**Equipment and software.** iMotions Lab (Version 11.1.47136.14; iMotions A/S, Copenhagen, Denmark); Smart Eye Aurora screen based eye tracker (60 Hz) with three areas of interest per ad (logo, product, headline); Affectiva AFFDEX facial expression analysis (an expression counted as present when its score was 50 or higher); Shimmer3 GSR+ skin conductance sensor with the iMotions peak detection notebook.

**Measures.** Time looking at an area: gaze dwell time as a percentage of the ad. Smiling and frowning: percentage of time with Joy or Brow Furrow at or above 50. Arousal: skin conductance responses (peaks) per minute, compared with the rest screen. Survey: liking of each ad (1 to 7), brand named for each ad (multiple choice), attitude toward reusable bottles, and purchase intention for reusable bottles (1 to 7, asked once).

**Analysis.** Paired comparisons (the same viewers saw both ads), with a correction for testing several measures at once (Holm); exact tests for brand recall; a confidence interval for every difference. Full statistics: `results_text.md` and `SIMULATED_all_tests.csv`.

**Items to verify before a real report.** The fixation filter settings, the AFFDEX threshold, and the skin conductance notebook settings actually used on the lab computer; the iMotions version on the data collection computer.
