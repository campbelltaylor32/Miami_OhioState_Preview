# CFP Matchup Analysis: Miami vs Ohio State

A comprehensive, data-driven College Football Playoff matchup analysis pipeline using R for statistical computation and Python for PDF report generation.

## Overview

This project generates an in-depth statistical comparison between two CFP teams by:
1. **R Script** → Pulls play-by-play and game data from the CFBD API, computes advanced metrics, and exports CSV files
2. **Python Script** → Reads the CSV files and generates a professional PDF report with percentile-based analysis

## Output

**[Miami_vs_OhioState_Comprehensive_Report.pdf](Miami_vs_OhioState_Comprehensive_Report.pdf)** — A 24-page report featuring:
- Executive summary with statistical edge counts
- Full percentile-based metric comparison (40+ metrics)
- Head-to-head matchup matrices
- Strength vs. vulnerability analysis
- Momentum/trend analysis (last 3 games vs season)
- Quarter-by-quarter and situational breakdowns
- Line of scrimmage deep-dive
- Keys to the game and final verdict

---

## Files

| File | Language | Purpose |
|------|----------|---------|
| `Miami_OSU_Enhanced.R` | R | Data collection, feature engineering, statistical analysis |
| `Generate_PDF_Enhanced.py` | Python | PDF report generation from CSV outputs |
| `Miami_vs_OhioState_Comprehensive_Report.pdf` | — | Final output report |

---

## Pipeline

```
┌─────────────────────────────────────────────────────────────┐
│                    Miami_OSU_Enhanced.R                     │
│                                                             │
│  ┌─────────────┐    ┌──────────────────┐    ┌───────────┐  │
│  │  CFBD API   │───▶│ Feature Engineer │───▶│ CSV Files │  │
│  │  (cfbfastR) │    │  & Percentiles   │    │ (~20 files)│  │
│  └─────────────┘    └──────────────────┘    └───────────┘  │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                 Generate_PDF_Enhanced.py                    │
│                                                             │
│  ┌───────────┐    ┌──────────────────┐    ┌─────────────┐  │
│  │ CSV Files │───▶│ ReportLab Tables │───▶│   PDF Report │  │
│  │           │    │ & Formatting     │    │   (24 pages) │  │
│  └───────────┘    └──────────────────┘    └─────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

---

## R Script: `Miami_OSU_Enhanced.R`

### Dependencies
```r
library(cfbfastR)
library(tidyverse)
library(data.table)
library(RcppRoll)
library(gt)
library(scales)
```

### Data Sources (via CFBD API)
- Game info and scores
- Betting lines (spread, over/under)
- Team game stats
- Play-by-play with EPA/WPA
- Advanced season stats
- SP+ ratings

### Key Metrics Computed

| Category | Metrics |
|----------|---------|
| **Efficiency** | EPA/play, EPA/rush, EPA/pass, success rates |
| **Situational** | 3rd down, red zone, standard/passing downs |
| **Explosiveness** | Explosive play rates (10+ rush, 20+ pass) |
| **Pressure** | Sack rate, pressure rate, disruption rate |
| **Trends** | Last 3 games vs season averages |
| **Game Script** | Performance by score differential |
| **Line of Scrimmage** | Stuff rate, TFL rate, YPC progression by quarter |

### Output CSV Files
```
Miami_vs_OhioState_Comparison.csv
Miami_vs_OhioState_Vulnerabilities.csv
Miami_vs_OhioState_Strengths.csv
Miami_vs_OhioState_Edges.csv
Miami_vs_OhioState_Trends.csv
Miami_vs_OhioState_GameScript.csv
Miami_vs_OhioState_Quarter_Analysis.csv
Miami_vs_OhioState_Situation_Analysis.csv
Miami_vs_OhioState_Field_Position.csv
Miami_vs_OhioState_Drive_Efficiency.csv
Miami_vs_OhioState_Pressure_Analysis.csv
Miami_vs_OhioState_Late_Close.csv
Miami_vs_OhioState_Tempo.csv
Miami_vs_OhioState_Run_Progression_Quarters.csv
Miami_vs_OhioState_Run_Progression_Trend.csv
Miami_vs_OhioState_Obvious_Run.csv
Miami_vs_OhioState_Run_Defense_Quarters.csv
Miami_vs_OhioState_Run_Defense_Trend.csv
Miami_vs_OhioState_Run_Defense_Obvious.csv
Miami_vs_OhioState_Protection_Obvious_Pass.csv
Miami_vs_OhioState_Pass_Rush_Obvious.csv
Miami_vs_OhioState_LOS_Summary.csv
Miami_vs_OhioState_Full_Season_Data.csv
```

---

## Python Script: `Generate_PDF_Enhanced.py`

### Dependencies
```python
pandas
numpy
reportlab
```

### Features
- **Percentile visualization**: Color-coded cells (green = elite, red = vulnerable)
- **Quality-adjusted percentiles**: Defense metrics inverted so higher = better
- **Symbols**: ★ for elite (90th+), ■ for vulnerability (20th or below)
- **Edge detection**: Automatic advantage assignment based on percentile gaps

### Report Sections
1. Title & Matchup Summary
2. Executive Summary
3. Full Statistical Comparison
4. Head-to-Head Matchup Matrix
5. Key Matchup Edges
6. Team Strengths & Vulnerabilities
7. Offensive Analysis
8. Defensive Analysis
9. Momentum/Trend Analysis
10. Quarter-by-Quarter Performance
11. Down & Distance Situational Analysis
12. Drive Efficiency & Pressure Analysis
13. Game Script Performance
14. Line of Scrimmage Summary
15. Run Game Progression
16. Run Defense Progression
17. Pass Protection & Rush Analysis
18. LOS Battle Commentary
19. Narrative Analysis ("Is Miami Better Up Front?")
20. Keys to the Game
21. Final Verdict

---

## Usage

### Step 1: Run the R Script
```bash
Rscript Miami_OSU_Enhanced.R
```
> Requires a valid CFBD API key set via `Sys.setenv(CFBD_API_KEY = "your_key")`

### Step 2: Generate the PDF
```bash
python Generate_PDF_Enhanced.py
```

### Configuration
To analyze a different matchup, modify these variables:

**In R script:**
```r
teams_to_compare <- c("Team1", "Team2")
week_update <- 15  # Latest week of data
```

**In Python script:**
```python
TEAM1 = "Team1"
TEAM2 = "Team2"
```

---

## Key Insights from Sample Report

| Advantage | Miami | Ohio State |
|-----------|-------|------------|
| **Statistical Edges** | 8 | 38 |
| **Offensive EPA/Play** | 69th %ile | 91st %ile ★ |
| **Defensive EPA/Play** | 90th %ile ★ | 94th %ile ★ |
| **Red Zone TD Rate Allowed** | 7th %ile ■ | 70th %ile |
| **3rd & Long Sack Rate** | 3.8% | 16.4% |

---

## Requirements

### R
- R ≥ 4.0
- cfbfastR (+ CFBD API key)
- tidyverse, data.table, RcppRoll, gt, scales

### Python
- Python ≥ 3.8
- pandas, numpy, reportlab

---

## License

This project is for educational and analytical purposes. Data is sourced from the College Football Data API (collegefootballdata.com).
