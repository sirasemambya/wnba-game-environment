# WNBA Game Environment

A daily R script that builds a color-coded Excel "game environment" sheet for every WNBA game on a given date. For each team in each game it pulls recent form, pace, and opponent defensive rankings, flags back-to-backs, and (optionally) adds the implied team total from market lines, so you can see at a glance how fast, how high-scoring, and how favorable a matchup is.

It is a data-wrangling and presentation tool, not a predictive model. The value is in putting the relevant context for a whole slate in one sheet.

## What is in the sheet

One row per team per game.

| Section | Columns | Notes |
|---|---|---|
| Overview | `team`, `matchup`, `b2b`, `tip_off` | `b2b` is flagged when the team played the previous night |
| Scoring | `implied`, `ppg_avg`, `ppg_l10`, `plus_minus`, `plus_minus_l10` | `plus_minus` is the implied team total minus the team's average, so positive means the game is expected to be higher scoring than the team's norm |
| Pace | `opp_pace_rank`, `pace_rank`, `proj_pace` | `proj_pace` is the average of both teams' last-10 pace |
| Points defense | `opp_papg_rank`, `opp_drat_rank` | Opponent points allowed and defensive rating ranks (1 = toughest defense) |
| Rebounding | `opp_reb_rank`, `opp_reb_pct_rank` | Opponent rebounding ranks |

Rank cells are colored by tier across the 15-team league, with opponent-defense ranks inverted so green always means an easier scoring environment. `proj_pace` is colored on an absolute scale, and a pace key is written under the table.

## Sample output

From a run on September 19, 2026 (columns trimmed):

| team | matchup | tip_off | implied | ppg_avg | plus_minus | proj_pace | opp_papg_rank |
|---|---|---|---|---|---|---|---|
| Dallas Wings | vs Phoenix Mercury | 1:00 pm ET | 91.8 | 89.2 | +2.6 | 96.5 | 9 |
| Phoenix Mercury | at Dallas Wings | 1:00 pm ET | 80.8 | 84.4 | -3.6 | 96.5 | 5 |
| Atlanta Dream | vs Chicago Sky | 7:00 pm ET | 95.5 | 91.1 | +4.4 | 98.7 | 11 |
| Chicago Sky | at Atlanta Dream | 7:00 pm ET | 80.5 | 86.5 | -6.0 | 98.7 | 2 |
| Seattle Storm | at Golden State Valkyries | 9:00 pm ET | 72.5 | 82.7 | -10.2 | 95.4 | 1 |

## Data sources

| Source | Provides | Access |
|---|---|---|
| stats.wnba.com | League-wide team stats (base, advanced, opponent) over the full season and last 10 games | Public but unofficial endpoints, no key |
| cdn.wnba.com | Daily scoreboard and static season schedule | Public JSON |
| [The Odds API](https://the-odds-api.com) (optional) | Market totals and spreads, used to derive implied team totals as `(total -/+ spread) / 2` | Free API key |

The schedule lookup tries three sources in order (stats scoreboard, CDN scoreboard, static season schedule) and uses the first that returns games for the date.

## Running it

```r
install.packages(c("tidyverse", "httr", "jsonlite", "openxlsx"))

# Optional: implied team totals. Put this in your .Renviron, never in code.
# ODDS_API_KEY=your_key_here

source("wnba_game_env.R")
df <- run_game_env()                 # today
df <- run_game_env("09/19/2026")     # a specific date (MM/DD/YYYY)
```

The workbook is written to `output/wnba_env_<timestamp>.xlsx`. Without an API key everything still works and the implied-total columns are simply blank.

## Limitations

- **Unofficial, unauthenticated endpoints.** The stats.wnba.com and stats.nba.com endpoints are not a supported public API. They can rate limit, time out, or change without notice, and you should check their terms of use before relying on them.
- **Hand-maintained team map.** `WNBA_TEAM_MAP` contains team IDs that have drifted from the IDs in the live schedule for the newest franchises. The script falls back to matching on team name when an ID lookup fails, which works but is a patch, not a fix.
- **Fixed Eastern time offset.** Tip-off conversion assumes UTC-4 (daylight time), which is fine during the season but not year-round.
- **Hard-coded season and thresholds.** The season string, the 15-team rank tiers, and the pace color thresholds are constants in the script.
- **macOS only for auto-open.** The script opens the finished workbook with the macOS `open` command.
- **No tests and no history.** Each run is a snapshot. Nothing is stored between days and there are no automated tests.

## What I'd build next

- Pull the team map and schedule from a single source of truth so IDs cannot drift.
- Store each day's table so pace and defense trends can be charted across a season.
- Add a schedule-strength and rest-days view across the whole league, not just the day's slate.
- Publish the sheet as a small web page instead of an Excel file.
