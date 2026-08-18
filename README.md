# Corporate AI Disclosure in 10-K Filings

## Overview

This repository contains the quantitative analysis for an exploratory study of how corporate AI-related disclosures change following legal or regulatory scrutiny.

The project uses MATLAB to retrieve and analyze SEC Form 10-K filings and measure changes in AI-related disclosure over time. Companies subject to different forms of scrutiny are compared with matched control companies over similar periods.

## Research Question

How do public companies change their AI-related disclosures following legal or regulatory scrutiny, and how do those changes compare with similar firms not subject to the same event?

## Study Design

The analysis uses three matched case-control pairs:

| Case Company | Matched Control | Event Type |
|---|---|---|
| Presto Automation | PAR Technology | SEC Enforcement |
| GitLab | JFrog | Securities Litigation |
| Apple | Alphabet | Consumer Litigation |

Presto is the formal SEC-enforcement case. GitLab and Apple are included as additional forms of legal scrutiny and should not be interpreted as SEC-enforcement cases.

The primary analysis covers fiscal years 2020–2025 where filings are available.

## MATLAB Analysis

`sec_ai_10k_FINAL_RESEARCH.m` automatically:

- Retrieves Form 10-K filing information from SEC EDGAR
- Extracts filing text
- Identifies paragraphs containing AI-related terminology
- Counts AI-related paragraphs and term mentions
- Normalizes AI mentions by filing length
- Identifies risk, caution, specificity, and thematic disclosure language
- Measures year-to-year textual similarity
- Estimates new versus repeated AI-related language
- Produces a researcher-defined exploratory AI Disclosure Index
- Conducts pre/post-event comparisons
- Compares changes with matched control firms

The AI Disclosure Index and automated textual categories are exploratory research measures and should not be interpreted as established SEC or industry metrics.

## Main Findings

The analysis does not identify a uniform disclosure response following legal or regulatory scrutiny.

### GitLab vs. JFrog

GitLab increased its AI-related disclosure following the litigation event:

- AI mentions: 58 → 81
- AI paragraphs: 23 → 31
- AI Disclosure Index: 51.13 → 46.44

However, matched control JFrog experienced a larger increase over the comparable period. The matched difference in the AI Disclosure Index was **-13.75 points**.

### Apple vs. Alphabet

Apple also increased its AI-related disclosure:

- AI mentions: 6 → 11
- AI paragraphs: 4 → 7
- AI Disclosure Index: 46.40 → 50.07

Apple's index increased modestly more than Alphabet's, producing a matched difference of **+3.14 points**.

### Presto Automation

Presto serves as the study's formal SEC-enforcement case. However, no post-event Form 10-K was available, preventing a direct before-and-after estimate.

## Interpretation

The results suggest that changes in corporate AI disclosure following scrutiny are heterogeneous rather than uniform.

Importantly, increases in AI-related language cannot automatically be attributed to legal scrutiny. AI discussion was increasing broadly across corporate filings during the study period, making matched controls useful for distinguishing company-specific changes from broader reporting trends.

The results are descriptive and exploratory and do not establish a causal relationship between legal or regulatory events and subsequent disclosure behavior.

## Repository Contents

- `sec_ai_10k_FINAL_RESEARCH.m` — MATLAB SEC filing retrieval and textual-analysis pipeline
- `AI_Disclosure_Clean_Analysis.xlsx` — cleaned case and matched-control analysis
- `AI_Disclosure_Paper_Tables_Figures.xlsx` — tables and figures prepared for the research paper

## Reproducibility

The MATLAB script requires users to provide their own contact email for SEC EDGAR requests.

Before running the program, replace:

`YOUR_EMAIL@example.com`

with an appropriate contact email.

The program then retrieves available filing data directly from SEC EDGAR and generates the study outputs.

## Limitations

This exploratory study uses a small case-study sample. Only Presto represents formal SEC enforcement, and no post-event 10-K was available for that company. The other events represent different forms of private litigation.

Keyword-based textual classification can produce false positives and false negatives. Automated classifications were supplemented with manual review, and the researcher-defined AI Disclosure Index should be interpreted as an exploratory measure.

Matched controls reduce, but do not eliminate, differences between firms or broader time trends. The findings therefore describe observed disclosure patterns rather than causal effects.

## Status

Research in progress.
