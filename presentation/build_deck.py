import os
import pptx_theme as th
from slide_builders import (title_slide, agenda_slide, headline_slide, divider_slide, data_slide,
                             data_slide_wide, recommendation_slide, summary_slide, thankyou_slide,
                             appendix_divider, appendix_text_slide)
from pptx.util import Emu

HERE = os.path.dirname(os.path.abspath(__file__))
CH = lambda name: os.path.join(HERE, "charts", name + ".png")

prs = th.new_deck()
TOTAL = 4

# ---- 1: title -------------------------------------------------------------
title_slide(
    prs,
    "DATACO SUPPLY CHAIN",
    "Signal vs Noise",
    "Three years of order data, six questions — and three data problems worth catching before trusting any of the six answers",
    "Order-level data review  |  180,519 order lines, Jan 2015 – Jan 2018  |  September 2026",
)

# ---- 2: agenda --------------------------------------------------------------
agenda_slide(prs, [
    ("01", "Data trust", "Can we trust the growth number this data seems to show?"),
    ("02", "Fraud & delivery risk", "Are fraud and late deliveries actually costing us profit?"),
    ("03", "Region rotation", "Which regions are actually growing or declining?"),
    ("04", "Discount program", "Is the discount program actually working?"),
], page=2)

# ---- 3: headline numbers ----------------------------------------------------
headline_slide(
    prs,
    "Total sales and profit by year — a healthy, steady business",
    CH("headline_chart"),
    [
        ("180,519", "order lines, full raw dataset (Jan 2015 – Jan 2018)"),
        ("12,330", "verified real customers, clean analysis window"),
        ("+4.7%", "revenue growth, 2017 vs. 2016 (like-for-like, Jan–Sep)"),
        ("3", "data-generation artifacts found and corrected for"),
    ],
    "2015–2017 shown as full calendar years; 2018 excluded (one partial month only). "
    "2017's dip is explained in Insight 1 — its last quarter falls inside a data-recording "
    "change, not a real sales decline: the like-for-like Jan–Sep read is +4.7%.",
    page=3,
)

# ---- STORY 1: Data trust ----------------------------------------------------
divider_slide(prs, 1, TOTAL, "Data Trust",
              "Can we trust the growth number this data seems to show?", page=4)

data_slide(prs, 1, TOTAL, "Data Trust",
           "Order composition collapses in November 2017",
           CH("story1_chart_a"),
           "1.00", "Line items per order, Nov 2017 onward (was ~3.0)",
           "Average order value falls from $620 to $156 across the same months, while order counts "
           "hold steady or grow. That combination only makes sense as a change in how orders are "
           "recorded — not a change in demand.",
           page=5)

data_slide(prs, 1, TOTAL, "Data Trust",
           "Margin doesn't move — even across that same boundary",
           CH("story1_chart_b"),
           "10.8–12.7%", "Quarterly margin band, every quarter Q1 2015 – Q1 2018",
           "Order counts even jump the quarter this happens, from ~5,200 to 6,280 — more orders, "
           "just recorded smaller. Margin, and the underlying business, are unaffected; only "
           "revenue-per-order and anything built on it needs correcting.",
           page=6)

recommendation_slide(
    prs, 1, TOTAL, "Data Trust",
    "Exclude or separately model October 2017 onward in any growth, trend, or cohort analysis on "
    "this dataset — this is a data-engineering fix, not a business problem.",
    "Flag the order-composition change to whoever owns this pipeline, alongside two related issues "
    "found the same way — a region-rotation pattern and a batch of customer IDs that only exist "
    "after this date.",
    "The underlying business is healthy: +4.7% YoY growth and a stable margin band on the window "
    "that can be trusted.",
    page=7,
)

# ---- STORY 2: Fraud & delivery risk -----------------------------------------
divider_slide(prs, 2, TOTAL, "Fraud & Delivery Risk",
              "Are fraud and late deliveries actually costing us profit?", page=8)

data_slide(prs, 2, TOTAL, "Fraud & Delivery Risk",
           "Fraud, late, and canceled orders book normal profit",
           CH("story3_chart_a"),
           "12.04%", "Company average margin — every order-status category lands within ~1pt",
           "Late-delivery, suspected-fraud, and canceled orders all book profit close to the company "
           "average, both company-wide and inside Cardio Equipment specifically. The profit field "
           "simply isn't written down differently for these statuses — there's no recoverable "
           "dollar pool here.",
           page=9)

data_slide_wide(prs, 2, TOTAL, "Fraud & Delivery Risk",
                 "But the rate of these events is a real regional signal",
                 CH("story3_chart_b"), Emu(3383282),
                 "A permutation test confirms region-level spread in fraud/late-delivery rates is real "
                 "(p<0.01); category-level spread is just noise. Western Europe's 2.61% fraud rate is "
                 "modest but, on 25,000+ orders, the strongest statistical signal in the dataset. Canada "
                 "pairs the lowest late-delivery rate with an elevated (small-sample) fraud rate; Central "
                 "Africa and South Asia run the highest late-delivery rates, ~56–58%.",
                 page=10)

recommendation_slide(
    prs, 2, TOTAL, "Fraud & Delivery Risk",
    "Treat this as an operational review, not a profit-recovery project — the rate gap is real "
    "even though there's no profit sitting behind it to recover.",
    "Western Europe gets a fraud-controls review; Central Africa and South Asia get a "
    "delivery/logistics review.",
    "Neither traces back to shipping-method mix — already ruled out as an explanation.",
    page=11,
)

# ---- STORY 3: Region rotation -----------------------------------------------
divider_slide(prs, 3, TOTAL, "Region Rotation",
              "Which regions are actually growing or declining?", page=12)

data_slide_wide(prs, 3, TOTAL, "Region Rotation",
                 "Most regions are only active 5–6 of 36 months",
                 CH("story4_chart_a"), Emu(3383282),
                 "19 of 23 regions go permanently dark before the dataset ends. Regions arrive and "
                 "disappear in clean rotating cohorts, not organic activity. Company-wide order volume "
                 "stays flat throughout — including through what first looked like a 19-month "
                 "regional dead zone. That's a data-generation rotation, not a market signal.",
                 page=13)

data_slide(prs, 3, TOTAL, "Region Rotation",
           "Regions move together in cohorts, not independently",
           CH("story4_chart_b"),
           "0 of 7", "Regions that diverged from their cohort-mates' direction",
           "Only 7 of 23 regions even have two comparable time windows to test. All 3 Latin America "
           "regions decline together; all 3 Europe regions grow together — every time, same "
           "direction as their cohort. Once corrected, Central America overtakes Western Europe as "
           "the #1 region by profit dollars.",
           page=14, stat_color=th.GOOD)

recommendation_slide(
    prs, 3, TOTAL, "Region Rotation",
    "Stop reading region trend or ranking charts on this dataset at face value.",
    "Use the corrected Jan 2015 – Sep 2017 window and check each region's active-month count "
    "before quoting any regional ranking or trend.",
    "This dataset can't honestly answer “which region should we invest in” — what it can "
    "confirm is Central America is the real #1 region by profit.",
    page=15,
)

# ---- STORY 4: Discount program -----------------------------------------------
divider_slide(prs, 4, TOTAL, "Discount Program",
              "Is the discount program actually working?", page=16)

data_slide(prs, 4, TOTAL, "Discount Program",
           "The discount “sweet spot” survives reshuffling — that's the problem",
           CH("story6_chart"),
           "≈ identical", "Real order pattern vs. a randomly shuffled null test",
           "Orders averaging a 6–15% discount carry ~3x the line items of orders at 0% or "
           "21–25% — looked like a real threshold. Reshuffle every discount value at random and "
           "rerun the same aggregation: the identical curve comes back. Order size creates this "
           "pattern by chance, not discount level.",
           page=17, stat_color=th.INK_FAINT)

recommendation_slide(
    prs, 4, TOTAL, "Discount Program",
    "Run an actual randomized experiment before trusting anything discount-related in this data.",
    "Four arms — 0/10/20/35% — stratified by segment and category, conversion rate as the "
    "primary metric with margin as a guardrail; ~2,100 orders per arm detects a 5% lift in order "
    "value.",
    "The real blocker isn't sample size — there's currently no record of “offered a discount, "
    "didn't buy.” Build that tracking before or during the test.",
    page=18,
)

# ---- 19: summary --------------------------------------------------------------
summary_slide(
    prs,
    "Four findings, ranked by how sure we are — not by size",
    [
        ("Growth number is real, once corrected",
         "Exclude/model Oct 2017+ separately in all trend and cohort work; flag to data engineering",
         "High"),
        ("Region rankings need the same correction",
         "Adopt the corrected window + active-month check as the region-reporting standard",
         "High"),
        ("Fraud/late-delivery rate gaps by region",
         "Operational review: Western Europe (fraud), Central Africa & South Asia (delivery)",
         "High"),
        ("Discount rate's effect on behavior is unproven",
         "Run the randomized 4-arm test; build “offered, didn't buy” tracking first",
         "Low"),
    ],
    page=19,
)

# ---- 20: thank you --------------------------------------------------------------
thankyou_slide(prs, "Which one or two should we move on first?")

# ---- 21-23: appendix --------------------------------------------------------------
appendix_divider(prs, "Data window, definitions & methodology", page=21)

appendix_text_slide(
    prs,
    "Data window, definitions & known gaps",
    [
        ("Data window:",
         "180,519 order lines, Jan 2015–Jan 2018 (raw). Clean analysis window used for growth, "
         "trend, cohort, and region work: Jan 2015–Sep 2017 (171,962 order lines, 12,330 verified "
         "customers)."),
        ("Known artifact:",
         "Order composition changes structurally from Nov 2017 — every order becomes 1 line item, "
         "average order value falls from ~$620 to ~$156 by Jan 2018, while order counts hold or rise. "
         "Treated as a recording change, not a demand change."),
        ("Margin definition:",
         "Profit ÷ Order Item Total (post-discount revenue) — not ÷ Sales (pre-discount "
         "list price), which understates margin. Verified against the dataset's own profit-ratio "
         "field."),
        ("Region-rotation caveat:",
         "19 of 23 regions are active only 5–6 of 36 months, arriving/disappearing in cohorts. Any "
         "region-level ranking or trend must check months-active before it's trusted."),
    ],
    page=22,
)

appendix_text_slide(
    prs,
    "Methodology behind the statistical claims",
    [
        ("Region-level fraud/late-delivery signal:",
         "Permutation test on rate spread by region (p=0.007 fraud, p=0.002 late-delivery) vs. by "
         "category (p=0.156, p=0.224) — region spread is real, category spread is noise."),
        ("Discount-rate mirage:",
         "Every discount value was reshuffled at random across all rows and the same order-level "
         "aggregation rerun; the reshuffled data reproduced the same inverted-U curve as the real "
         "data, confirming the pattern is a property of order size, not of discount level."),
    ],
    page=23,
)

out_path = os.path.join(HERE, "DataCo_Supply_Chain_Findings.pptx")
prs.save(out_path)
print("Saved:", out_path, "slides:", len(prs.slides._sldIdLst))
