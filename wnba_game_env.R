# wnba_game_env.R
# Daily WNBA game environment table
# Pulls L10 team stats + opponent defensive profile + implied totals
# Outputs a color-coded Excel file for game-context analysis

library(tidyverse)
library(httr)
library(jsonlite)
library(openxlsx)

# Optional. Set in your environment or .Renviron, never in code:
#   ODDS_API_KEY=your_key_here
ODDS_API_KEY <- Sys.getenv("ODDS_API_KEY")
SEASON       <- "2026"
LEAGUE_ID    <- "10"  # WNBA

# ── WNBA Stats API ─────────────────────────────────────────────────────────────

wnba_headers <- function(host = "stats.wnba.com") {
  c(
    "Host"               = host,
    "User-Agent"         = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept"             = "application/json, text/plain, */*",
    "Accept-Language"    = "en-US,en;q=0.9",
    "Accept-Encoding"    = "gzip, deflate, br",
    "x-nba-stats-origin" = "stats",
    "x-nba-stats-token"  = "true",
    "Origin"             = "https://www.wnba.com",
    "Referer"            = "https://www.wnba.com/",
    "Sec-Fetch-Dest"     = "empty",
    "Sec-Fetch-Mode"     = "cors",
    "Sec-Fetch-Site"     = "same-site",
    "Connection"         = "keep-alive"
  )
}

wnba_get <- function(endpoint, params = list(), max_tries = 3) {
  url <- paste0("https://stats.wnba.com/stats/", endpoint)
  for (i in seq_len(max_tries)) {
    resp <- tryCatch(
      GET(url, add_headers(.headers = wnba_headers()), query = params,
          config = httr::timeout(60)),
      error = function(e) { message("  Attempt ", i, " failed: ", e$message); NULL }
    )
    if (!is.null(resp) && status_code(resp) == 200)
      return(content(resp, "parsed", encoding = "UTF-8"))
    if (!is.null(resp))
      message("  HTTP ", status_code(resp), " on attempt ", i)
    if (i < max_tries) Sys.sleep(3 * i)
  }
  warning("WNBA stats API failed for: ", endpoint)
  NULL
}

parse_wnba <- function(resp, idx = 1, name = NULL) {
  if (is.null(resp)) return(tibble())
  rs_list <- resp$resultSets
  if (!is.null(name)) {
    nm <- which(sapply(rs_list, function(x) x$name) == name)
    idx <- if (length(nm) > 0) nm[1] else idx
  }
  rs <- rs_list[[idx]]
  if (is.null(rs) || length(rs$rowSet) == 0) return(tibble())
  hdrs <- unlist(rs$headers)
  rows <- lapply(rs$rowSet, function(r) lapply(r, function(x) if (is.null(x)) NA else x))
  df   <- as.data.frame(do.call(rbind, rows), stringsAsFactors = FALSE)
  names(df) <- hdrs
  as_tibble(df) %>%
    mutate(across(where(is.list), ~ unlist(.))) %>%
    mutate(across(where(~ all(!is.na(.) & grepl("^-?[0-9.]+$", as.character(.)))), as.numeric))
}

# ── 1. Today's schedule ────────────────────────────────────────────────────────

pull_schedule_cdn <- function(date_mmddyyyy) {
  iso <- format(as.Date(date_mmddyyyy, "%m/%d/%Y"), "%Y-%m-%d")
  resp <- tryCatch(
    GET("https://cdn.wnba.com/static/json/liveData/scoreboard/todaysScoreboard_10.json",
        add_headers(.headers = wnba_headers("cdn.wnba.com")), config = httr::timeout(15)),
    error = function(e) { message("  CDN HTTP error: ", e$message); NULL }
  )
  message("  CDN HTTP status: ", if (is.null(resp)) "NULL" else status_code(resp))
  if (is.null(resp) || status_code(resp) != 200) return(tibble())
  data <- tryCatch(fromJSON(content(resp, "text", encoding = "UTF-8"), simplifyVector = FALSE),
                   error = function(e) NULL)
  if (is.null(data)) return(tibble())
  sb <- data$scoreboard
  if (is.null(sb)) return(tibble())
  cdn_date <- if (!is.null(sb$gameDate)) sb$gameDate else ""
  message("  CDN gameDate=", cdn_date, " (looking for ", iso, ")")
  if (!startsWith(cdn_date, iso)) {
    message("  CDN date mismatch — skipping")
    return(tibble())
  }
  games <- sb$games
  if (length(games) == 0) { message("  CDN: no games for ", cdn_date); return(tibble()) }
  bind_rows(lapply(games, function(g) tibble(
    game_id   = g$gameId,
    tip_off   = g$gameStatusText,
    home_id   = as.numeric(g$homeTeam$teamId),
    away_id   = as.numeric(g$awayTeam$teamId),
    home_name = paste(g$homeTeam$teamCity, g$homeTeam$teamName),
    away_name = paste(g$awayTeam$teamCity, g$awayTeam$teamName)
  )))
}

utc_to_et <- function(utc_str) {
  if (is.null(utc_str) || nchar(trimws(utc_str)) == 0) return("")
  # Extract HH:MM from any format: "1900-01-01T23:30:00Z", "2026-05-08T23:30:00Z", "23:30:00Z"
  m <- regmatches(utc_str, regexpr("(\\d{2}):(\\d{2}):\\d{2}Z?$", utc_str, perl = TRUE))
  if (length(m) == 0) return(utc_str)
  parts <- strsplit(m, ":")[[1]]
  h_utc <- as.integer(parts[1])
  mn    <- as.integer(parts[2])
  # ET = UTC - 4 (EDT)
  h_et  <- (h_utc - 4) %% 24
  ampm  <- if (h_et < 12) "AM" else "PM"
  h_12  <- h_et %% 12; if (h_12 == 0) h_12 <- 12
  sprintf("%d:%02d %s ET", h_12, mn, ampm)
}

pull_schedule_static <- function(date_mmddyyyy) {
  resp <- tryCatch(
    GET("https://cdn.wnba.com/static/json/staticData/scheduleLeagueV2_10.json",
        add_headers(.headers = wnba_headers("cdn.wnba.com")), config = httr::timeout(20)),
    error = function(e) { message("  Static HTTP error: ", e$message); NULL }
  )
  message("  Static HTTP status: ", if (is.null(resp)) "NULL" else status_code(resp))
  if (is.null(resp) || status_code(resp) != 200) return(tibble())
  data <- tryCatch(
    fromJSON(content(resp, "text", encoding = "UTF-8"), simplifyVector = FALSE),
    error = function(e) NULL
  )
  if (is.null(data)) return(tibble())
  game_dates <- data$leagueSchedule$gameDates
  if (is.null(game_dates)) return(tibble())

  # Match on either MM/DD/YYYY or ISO YYYY-MM-DD prefix
  d          <- as.Date(date_mmddyyyy, "%m/%d/%Y")
  prefix_mdy <- format(d, "%m/%d/%Y")
  prefix_iso <- format(d, "%Y-%m-%d")

  message("  Static schedule: searching for date '", prefix_mdy, "' or '", prefix_iso, "'")
  message("  Static schedule: total game-date entries = ", length(game_dates))
  # Show a sample of actual date strings to reveal format
  sample_dates <- head(sapply(game_dates, function(gd) if (!is.null(gd$gameDate)) gd$gameDate else ""), 5)
  message("  Static schedule sample dates: ", paste(sample_dates, collapse = " | "))

  day <- NULL
  for (gd in game_dates) {
    gd_date <- if (!is.null(gd$gameDate)) gd$gameDate else ""
    if (startsWith(gd_date, prefix_mdy) || startsWith(gd_date, prefix_iso)) { day <- gd; break }
  }
  if (is.null(day) || length(day$games) == 0) {
    # Try looser match in case year format differs (e.g. 2-digit year)
    for (gd in game_dates) {
      gd_date <- if (!is.null(gd$gameDate)) gd$gameDate else ""
      if (grepl(format(d, "%m/%d"), gd_date, fixed = TRUE)) { day <- gd; break }
    }
  }
  if (is.null(day) || length(day$games) == 0) return(tibble())
  for (g in day$games) {
    message("  Static schedule game: ", g$homeTeam$teamCity, " ", g$homeTeam$teamName,
            " (", g$homeTeam$teamId, ") vs ",
            g$awayTeam$teamCity, " ", g$awayTeam$teamName,
            " (", g$awayTeam$teamId, ")")
  }
  bind_rows(lapply(day$games, function(g) tibble(
    game_id   = g$gameId,
    tip_off   = utc_to_et(g$gameTimeUTC),
    home_id   = as.numeric(g$homeTeam$teamId),
    away_id   = as.numeric(g$awayTeam$teamId),
    home_name = paste(g$homeTeam$teamCity, g$homeTeam$teamName),
    away_name = paste(g$awayTeam$teamCity, g$awayTeam$teamName)
  )))
}

pull_schedule <- function(date = format(Sys.Date(), "%m/%d/%Y")) {
  message("Pulling WNBA schedule for ", date, "...")

  # 1. scoreboardv2 with WNBA league ID
  resp <- tryCatch(
    GET("https://stats.nba.com/stats/scoreboardv2",
        add_headers(.headers = wnba_headers()),
        query = list(DayOffset = 0, LeagueID = LEAGUE_ID, gameDate = date),
        config = httr::timeout(30)),
    error = function(e) NULL
  )
  raw <- parse_wnba(resp, name = "GameHeader")
  if (!is.null(raw) && nrow(raw) > 0) {
    return(raw %>%
      select(GAME_ID, GAME_STATUS_TEXT, HOME_TEAM_ID, VISITOR_TEAM_ID) %>%
      rename(game_id = GAME_ID, tip_off = GAME_STATUS_TEXT,
             home_id = HOME_TEAM_ID, away_id = VISITOR_TEAM_ID) %>%
      mutate(across(c(home_id, away_id), as.numeric),
             home_name = NA_character_, away_name = NA_character_))
  }

  # 2. WNBA CDN live scoreboard
  message("  scoreboardv2 empty — trying WNBA CDN...")
  games <- pull_schedule_cdn(date)
  if (!is.null(games) && nrow(games) > 0) return(games)

  # 3. WNBA static season schedule
  message("  CDN empty/stale — trying WNBA static schedule...")
  games <- pull_schedule_static(date)
  if (!is.null(games) && nrow(games) > 0) return(games)

  message("No games found")
  tibble()
}

# ── Check back-to-back ─────────────────────────────────────────────────────────

get_b2b_teams <- function(date = format(Sys.Date(), "%m/%d/%Y")) {
  yesterday <- format(as.Date(date, "%m/%d/%Y") - 1, "%m/%d/%Y")
  games <- pull_schedule(yesterday)
  if (is.null(games) || nrow(games) == 0) return(integer(0))
  c(as.numeric(games$home_id), as.numeric(games$away_id))
}

# ── 2. League-wide team stats (L10) ───────────────────────────────────────────

pull_team_stats <- function(measure, last_n = 10) {
  Sys.sleep(1)
  message("  Pulling ", measure, " stats (L", last_n, ")...")
  resp <- wnba_get("leaguedashteamstats", list(
    Conference       = "",
    DateFrom         = "",
    DateTo           = "",
    Division         = "",
    GameScope        = "",
    GameSegment      = "",
    Height           = "",
    LastNGames       = last_n,
    LeagueID         = LEAGUE_ID,
    Location         = "",
    MeasureType      = measure,
    Month            = 0,
    OpponentTeamID   = 0,
    Outcome          = "",
    PORound          = 0,
    PaceAdjust       = "N",
    PerMode          = "PerGame",
    Period           = 0,
    PlayerExperience = "",
    PlayerPosition   = "",
    PlusMinus        = "N",
    Rank             = "Y",
    Season           = SEASON,
    SeasonSegment    = "",
    SeasonType       = "Regular Season",
    ShotClockRange   = "",
    StarterBench     = "",
    TeamID           = 0,
    TwoWay           = 0,
    VsConference     = "",
    VsDivision       = ""
  ))
  parse_wnba(resp, 1)
}

build_team_table <- function() {
  message("Pulling WNBA team stats...")
  base    <- pull_team_stats("Base",     last_n = 0)   # full season PPG
  base_l10 <- pull_team_stats("Base",   last_n = 10)  # L10 PPG
  adv     <- pull_team_stats("Advanced", last_n = 10)  # L10 pace/rating
  opp     <- pull_team_stats("Opponent", last_n = 10)  # L10 defensive

  # Base: season PPG
  b <- if (!is.null(base) && nrow(base) > 0) {
    base %>%
      select(TEAM_ID, TEAM_NAME, PTS) %>%
      rename(team_id = TEAM_ID, team_name = TEAM_NAME, ppg = PTS) %>%
      mutate(team_id = as.numeric(team_id))
  } else tibble(team_id = numeric(0), team_name = character(0), ppg = numeric(0))

  # L10 PPG
  b_l10 <- if (!is.null(base_l10) && nrow(base_l10) > 0) {
    base_l10 %>%
      select(TEAM_ID, PTS) %>%
      rename(team_id = TEAM_ID, ppg_l10 = PTS) %>%
      mutate(team_id = as.numeric(team_id))
  } else tibble(team_id = numeric(0), ppg_l10 = numeric(0))

  # Advanced
  a <- if (!is.null(adv) && nrow(adv) > 0) {
    adv %>%
      select(TEAM_ID,
             OFF_RATING, OFF_RATING_RANK,
             DEF_RATING, DEF_RATING_RANK,
             PACE,       PACE_RANK,
             REB_PCT,    REB_PCT_RANK) %>%
      rename(team_id = TEAM_ID,
             o_rat = OFF_RATING, o_rat_rank = OFF_RATING_RANK,
             d_rat = DEF_RATING, d_rat_rank = DEF_RATING_RANK,
             pace  = PACE,       pace_rank  = PACE_RANK,
             reb_pct = REB_PCT,  reb_pct_rank = REB_PCT_RANK) %>%
      mutate(team_id = as.numeric(team_id))
  } else tibble(team_id = numeric(0))

  # Opponent
  o <- if (!is.null(opp) && nrow(opp) > 0) {
    opp %>%
      select(TEAM_ID,
             OPP_PTS,     OPP_PTS_RANK,
             OPP_REB,     OPP_REB_RANK,
             OPP_FG3A,    OPP_FG3A_RANK,
             OPP_FG3_PCT, OPP_FG3_PCT_RANK) %>%
      rename(team_id = TEAM_ID,
             papg       = OPP_PTS,     papg_rank       = OPP_PTS_RANK,
             opp_reb    = OPP_REB,     opp_reb_rank    = OPP_REB_RANK,
             opp_3pa    = OPP_FG3A,    opp_3pa_rank    = OPP_FG3A_RANK,
             opp_3p_pct = OPP_FG3_PCT, opp_3p_pct_rank = OPP_FG3_PCT_RANK) %>%
      mutate(team_id = as.numeric(team_id))
  } else tibble(team_id = numeric(0))

  # If no games played yet, seed from team list so schedule still builds
  if (nrow(b) == 0) {
    message("  No stats yet (season not started) — table will show N/A")
    return(WNBA_TEAM_MAP %>% rename(team_id = team_id_local, team_name = odds_name) %>%
             mutate(ppg = NA_real_, ppg_l10 = NA_real_,
                    o_rat = NA_real_, o_rat_rank = NA_real_,
                    d_rat = NA_real_, d_rat_rank = NA_real_,
                    pace  = NA_real_, pace_rank  = NA_real_,
                    reb_pct = NA_real_, reb_pct_rank = NA_real_,
                    papg = NA_real_,    papg_rank = NA_real_,
                    opp_reb = NA_real_, opp_reb_rank = NA_real_,
                    opp_3pa = NA_real_, opp_3pa_rank = NA_real_,
                    opp_3p_pct = NA_real_, opp_3p_pct_rank = NA_real_))
  }

  b %>%
    left_join(b_l10, by = "team_id") %>%
    left_join(a,     by = "team_id") %>%
    left_join(o,     by = "team_id")
}

# ── 3. Implied totals from The Odds API ───────────────────────────────────────

pull_implied_totals <- function(date = format(Sys.Date(), "%m/%d/%Y")) {
  if (ODDS_API_KEY == "") {
    message("No ODDS_API_KEY — implied totals will be blank")
    return(tibble())
  }

  # Preference order: first source with a total line wins
  LINE_SOURCES <- c("pinnacle", "draftkings", "fanduel", "betmgm", "caesars")

  message("Pulling WNBA totals + spreads (with fallback sources)...")
  tot_resp <- GET(
    "https://api.the-odds-api.com/v4/sports/basketball_wnba/odds/",
    query = list(apiKey = ODDS_API_KEY, regions = "us",
                 markets = "totals,spreads",
                 bookmakers = paste(LINE_SOURCES, collapse = ","),
                 oddsFormat = "american")
  )
  if (status_code(tot_resp) != 200) {
    message("Odds API: HTTP ", status_code(tot_resp))
    return(tibble())
  }
  data <- fromJSON(content(tot_resp, "text", encoding = "UTF-8"), flatten = TRUE)
  if (length(data) == 0 || nrow(data) == 0) return(tibble())

  # Filter to today's games using commence_time (UTC → ET date)
  target_date <- as.Date(date, "%m/%d/%Y")
  if ("commence_time" %in% names(data)) {
    game_dates <- as.Date(substr(data$commence_time, 1, 10)) - 4/24  # rough UTC→ET
    # keep games whose ET date matches target, allowing games up to 6am next day UTC
    data <- data[as.Date(data$commence_time) == target_date |
                 (as.Date(data$commence_time) == target_date + 1 &
                  as.integer(substr(data$commence_time, 12, 13)) < 6), , drop = FALSE]
    if (nrow(data) == 0) { message("  No odds games for ", date); return(tibble()) }
  }

  extract_line <- function(mkts, home_team) {
    total_line <- NA_real_; spread_home <- NA_real_
    for (j in seq_len(nrow(mkts))) {
      outcomes <- mkts$outcomes[[j]]
      if (is.null(outcomes)) next
      if (mkts$key[j] == "totals")  total_line  <- as.numeric(outcomes$point[1])
      if (mkts$key[j] == "spreads") {
        home_row <- outcomes[outcomes$name == home_team, ]
        if (nrow(home_row) > 0) spread_home <- as.numeric(home_row$point[1])
      }
    }
    list(total_line = total_line, spread_home = spread_home)
  }

  results <- list()
  for (i in seq_len(nrow(data))) {
    home_team <- data$home_team[i]
    away_team <- data$away_team[i]
    srcs      <- data$bookmakers[[i]]
    if (is.null(srcs) || length(srcs) == 0) next

    # Walk priority order — use first source that has a total line
    total_line <- NA_real_; spread_home <- NA_real_; source_used <- NA_character_
    for (src_key in LINE_SOURCES) {
      idx <- which(srcs$key == src_key)
      if (length(idx) == 0) next
      mkts <- srcs$markets[[idx[1]]]
      if (is.null(mkts)) next
      line <- extract_line(mkts, home_team)
      if (!is.na(line$total_line)) {
        total_line  <- line$total_line
        spread_home <- line$spread_home
        source_used <- src_key
        break
      }
    }

    if (!is.na(total_line)) {
      sp <- coalesce(spread_home, 0)
      message("  ", home_team, " vs ", away_team, " — using ", source_used)
      results[[i]] <- tibble(
        home_team    = home_team,
        away_team    = away_team,
        home_implied = round((total_line - sp) / 2, 1),
        away_implied = round((total_line + sp) / 2, 1)
      )
    } else {
      message("  ", home_team, " vs ", away_team, " — no line from any source")
    }
  }
  out <- bind_rows(results)
  message("=== Odds API teams returned ===")
  if (nrow(out) > 0) {
    for (k in seq_len(nrow(out)))
      message("  HOME: '", out$home_team[k], "'  AWAY: '", out$away_team[k], "'")
  } else {
    message("  (none)")
  }
  out
}

# ── WNBA team map: Odds API names → WNBA stats team IDs ──────────────────────

WNBA_TEAM_MAP <- tribble(
  ~odds_name,                  ~team_id_local,
  "Atlanta Dream",              1611661320,
  "Chicago Sky",                1611661321,
  "Connecticut Sun",            1611661323,
  "Dallas Wings",               1611661325,
  "Golden State Valkyries",     1611661329,
  "Indiana Fever",              1611661324,
  "Las Vegas Aces",             1611661319,
  "Los Angeles Sparks",         1611661318,
  "Minnesota Lynx",             1611661322,
  "New York Liberty",           1611661313,
  "Phoenix Mercury",            1611661317,
  "Portland Fire",              1611661330,
  "Seattle Storm",              1611661328,
  "Toronto Tempo",              1611661331,
  "Washington Mystics",         1611661327
)

# ── 4. Assemble game environment table ────────────────────────────────────────

build_game_env <- function(date = format(Sys.Date(), "%m/%d/%Y")) {
  schedule   <- pull_schedule(date)
  if (nrow(schedule) == 0) return(NULL)

  team_stats <- build_team_table()
  totals     <- pull_implied_totals(date)

  b2b_ids    <- tryCatch(get_b2b_teams(date), error = function(e) integer(0))

  rows <- list()

  for (i in seq_len(nrow(schedule))) {
    home_id <- schedule$home_id[i]
    away_id <- schedule$away_id[i]
    tip     <- schedule$tip_off[i]

    home <- team_stats %>% filter(team_id == home_id)
    away <- team_stats %>% filter(team_id == away_id)

    # If ID lookup fails, fall back to name matching using schedule names
    if ((nrow(home) == 0 || nrow(away) == 0) && "home_name" %in% names(schedule)) {
      sched_hn <- schedule$home_name[i]; sched_an <- schedule$away_name[i]
      name_match <- function(ts, nm) ts %>% filter(
        !is.na(nm) & (str_detect(team_name, fixed(nm, ignore_case = TRUE)) |
                      str_detect(nm, fixed(team_name, ignore_case = TRUE))))
      if (nrow(home) == 0) home <- name_match(team_stats, sched_hn)
      if (nrow(away) == 0) away <- name_match(team_stats, sched_an)
    }

    if (nrow(home) == 0 || nrow(away) == 0) {
      message("  Skipping game — unrecognized team ID(s): home=", home_id, " away=", away_id)
      next
    }

    # Match implied totals by team name directly (avoids stale WNBA_TEAM_MAP IDs)
    home_impl <- NA_real_; away_impl <- NA_real_
    if (nrow(totals) > 0) {
      match <- totals %>% filter(
        str_detect(home_team, fixed(home$team_name, ignore_case = TRUE)) |
        str_detect(home$team_name, fixed(home_team, ignore_case = TRUE))
      )
      if (nrow(match) > 0) { home_impl <- match$home_implied[1]; away_impl <- match$away_implied[1] }
    }

    proj_pace <- if (!is.na(home$pace) && !is.na(away$pace)) round((home$pace + away$pace) / 2, 1) else NA_real_

    make_row <- function(team_s, opp_s, impl, matchup_str) {
      tibble(
        # ── Section 1: Overview ──
        team           = team_s$team_name,
        matchup        = matchup_str,
        b2b            = team_s$team_id %in% b2b_ids,
        tip_off        = tip,
        implied        = impl,
        ppg_avg        = team_s$ppg,
        ppg_l10        = team_s$ppg_l10,
        plus_minus     = if (!is.na(impl) && !is.na(team_s$ppg))     round(impl - team_s$ppg,     1) else NA_real_,
        plus_minus_l10 = if (!is.na(impl) && !is.na(team_s$ppg_l10)) round(impl - team_s$ppg_l10, 1) else NA_real_,
        # ── Section 2: Pace ──
        opp_pace_rank  = opp_s$pace_rank,
        pace_rank      = team_s$pace_rank,
        proj_pace      = proj_pace,
        # ── Section 3: Points ──
        opp_papg_rank  = opp_s$papg_rank,
        opp_drat_rank  = opp_s$d_rat_rank,
        # ── Section 4: Rebounding ──
        opp_reb_rank     = opp_s$opp_reb_rank,
        opp_reb_pct_rank = opp_s$reb_pct_rank
      )
    }

    rows[[length(rows) + 1]] <- make_row(home, away, home_impl, paste0("vs ", away$team_name))
    rows[[length(rows) + 1]] <- make_row(away, home, away_impl, paste0("at ", home$team_name))
  }

  bind_rows(rows)
}

# ── 5. Color-coded Excel export ───────────────────────────────────────────────
#
# Rank color logic (15 WNBA teams):
#
#   OWN metrics (pace_rank, reb_pct_rank):
#     Rank 1-4   = dark green  (elite)
#     Rank 5-7   = light green
#     Rank 8-11  = light red
#     Rank 12-15 = dark red    (bad)
#
#   OPP defensive metrics (papg, drat, opp_reb, opp_pace):
#     Rank 1-4   = dark red    (tough matchup)
#     Rank 5-7   = light red
#     Rank 8-11  = light green (soft matchup)
#     Rank 12-15 = dark green  (great spot)

export_game_env <- function(df, path = NULL) {
  if (is.null(path))
    path <- paste0("output/wnba_env_", format(Sys.time(), "%Y-%m-%d_%H%M%S"), ".xlsx")

  dir.create(dirname(path), showWarnings = FALSE, recursive = TRUE)

  wb <- createWorkbook()
  addWorksheet(wb, "Game Environment")

  ci <- function(nm) which(names(df) == nm)

  # ── Section header styles ──
  hdr_overview   <- createStyle(fontName = "Calibri", fontSize = 10, textDecoration = "bold",
                                 fontColour = "#FFFFFF", halign = "center", wrapText = TRUE,
                                 fgFill = "#1D3557")
  hdr_pts_avg    <- createStyle(fontName = "Calibri", fontSize = 10, textDecoration = "bold",
                                 fontColour = "#FFFFFF", halign = "center", wrapText = TRUE,
                                 fgFill = "#2E75B6")
  hdr_pace       <- createStyle(fontName = "Calibri", fontSize = 10, textDecoration = "bold",
                                 fontColour = "#FFFFFF", halign = "center", wrapText = TRUE,
                                 fgFill = "#7030A0")
  hdr_points     <- createStyle(fontName = "Calibri", fontSize = 10, textDecoration = "bold",
                                 fontColour = "#FFFFFF", halign = "center", wrapText = TRUE,
                                 fgFill = "#C00000")
  hdr_rebounding <- createStyle(fontName = "Calibri", fontSize = 10, textDecoration = "bold",
                                 fontColour = "#FFFFFF", halign = "center", wrapText = TRUE,
                                 fgFill = "#538135")

  cell_base <- createStyle(fontName = "Calibri", fontSize = 10, halign = "center")
  alt_row   <- createStyle(fgFill = "#F0F4FA")
  num_1dp   <- createStyle(numFmt = "0.0", halign = "center")

  # Rank colors (shared for all 15-team coloring)
  dk_green <- createStyle(fgFill = "#1A7C3E", fontColour = "#FFFFFF", halign = "center", textDecoration = "bold", fontName = "Calibri", fontSize = 10)
  lt_green <- createStyle(fgFill = "#A9D18E", fontColour = "#1D4A1D", halign = "center", fontName = "Calibri", fontSize = 10)
  lt_red   <- createStyle(fgFill = "#F4B8B8", fontColour = "#7B0000", halign = "center", fontName = "Calibri", fontSize = 10)
  dk_red   <- createStyle(fgFill = "#C00000", fontColour = "#FFFFFF", halign = "center", textDecoration = "bold", fontName = "Calibri", fontSize = 10)
  neutral  <- createStyle(fgFill = "#FFFFCC", fontColour = "#7F6000", halign = "center", fontName = "Calibri", fontSize = 10)

  pos_diff <- createStyle(fgFill = "#C6EFCE", fontColour = "#276221", halign = "center", numFmt = "+0.0;-0.0", fontName = "Calibri", fontSize = 10)
  neg_diff <- createStyle(fgFill = "#FFC7CE", fontColour = "#9C0006", halign = "center", numFmt = "+0.0;-0.0", fontName = "Calibri", fontSize = 10)
  neu_diff <- createStyle(halign = "center", numFmt = "+0.0;-0.0", fontName = "Calibri", fontSize = 10)

  writeData(wb, "Game Environment", df)

  nrows     <- nrow(df)
  data_rows <- 2:(nrows + 1)
  even_rows <- seq(3, nrows + 1, by = 2)
  ncols     <- ncol(df)

  addStyle(wb, "Game Environment", cell_base, rows = 1:(nrows+1), cols = 1:ncols, gridExpand = TRUE)
  if (length(even_rows) > 0)
    addStyle(wb, "Game Environment", alt_row, rows = even_rows, cols = 1:ncols, gridExpand = TRUE, stack = TRUE)

  # ── Section headers ──
  sec_overview   <- intersect(c("team","matchup","b2b","tip_off"), names(df))
  sec_pts_avg    <- intersect(c("implied","ppg_avg","ppg_l10","plus_minus","plus_minus_l10"), names(df))
  sec_pace       <- intersect(c("opp_pace_rank","pace_rank","proj_pace"), names(df))
  sec_points     <- intersect(c("opp_papg_rank","opp_drat_rank"), names(df))
  sec_rebounding <- intersect(c("opp_reb_rank","opp_reb_pct_rank"), names(df))

  for (col in sec_overview)   addStyle(wb, "Game Environment", hdr_overview,   rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_pts_avg)    addStyle(wb, "Game Environment", hdr_pts_avg,    rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_pace)       addStyle(wb, "Game Environment", hdr_pace,       rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_points)     addStyle(wb, "Game Environment", hdr_points,     rows = 1, cols = ci(col), stack = TRUE)
  for (col in sec_rebounding) addStyle(wb, "Game Environment", hdr_rebounding, rows = 1, cols = ci(col), stack = TRUE)

  # ── Rank coloring helper (15-team: 1-4 / 5-7 / 8-11 / 12-15) ──
  color_rank_col <- function(col_name, direction) {
    if (!col_name %in% names(df)) return(invisible(NULL))
    col_i <- ci(col_name)
    vals  <- as.integer(df[[col_name]])
    if (direction == "low_good") {
      tiers <- list(
        list(range = 1:4,   style = dk_green),
        list(range = 5:7,   style = lt_green),
        list(range = 8:11,  style = lt_red),
        list(range = 12:15, style = dk_red)
      )
    } else {
      tiers <- list(
        list(range = 1:4,   style = dk_red),
        list(range = 5:7,   style = lt_red),
        list(range = 8:11,  style = lt_green),
        list(range = 12:15, style = dk_green)
      )
    }
    for (tier in tiers) {
      r <- which(vals %in% tier$range) + 1
      if (length(r) > 0)
        addStyle(wb, "Game Environment", tier$style, rows = r, cols = col_i, stack = TRUE)
    }
  }

  # Apply rank colors
  color_rank_col("pace_rank",      "low_good")
  color_rank_col("opp_pace_rank",  "low_good")
  for (col in c("opp_papg_rank","opp_drat_rank","opp_reb_rank","opp_reb_pct_rank"))
    color_rank_col(col, "high_good")

  # Proj pace: color by absolute value (WNBA avg ~95-97)
  pace_color_tiers <- function(pp) list(
    list(test = pp <= 93,              style = dk_red),
    list(test = pp > 93  & pp <= 95,   style = lt_red),
    list(test = pp > 95  & pp <= 97,   style = neutral),
    list(test = pp > 97  & pp <= 99,   style = lt_green),
    list(test = pp > 99,               style = dk_green)
  )
  if ("proj_pace" %in% names(df)) {
    pp_i <- ci("proj_pace")
    for (tier in pace_color_tiers(as.numeric(df$proj_pace))) {
      r <- which(tier$test) + 1
      if (length(r) > 0) addStyle(wb, "Game Environment", tier$style, rows = r, cols = pp_i, stack = TRUE)
    }
  }

  # +/- coloring
  for (pm_col in c("plus_minus", "plus_minus_l10")) {
    pm_i <- ci(pm_col)
    if (length(pm_i) == 0) next
    pm    <- df[[pm_col]]
    pos_r <- which(!is.na(pm) & pm > 0)  + 1
    neg_r <- which(!is.na(pm) & pm < 0)  + 1
    neu_r <- which(!is.na(pm) & pm == 0) + 1
    if (length(pos_r) > 0) addStyle(wb, "Game Environment", pos_diff, rows = pos_r, cols = pm_i, stack = TRUE)
    if (length(neg_r) > 0) addStyle(wb, "Game Environment", neg_diff, rows = neg_r, cols = pm_i, stack = TRUE)
    if (length(neu_r) > 0) addStyle(wb, "Game Environment", neu_diff, rows = neu_r, cols = pm_i, stack = TRUE)
  }

  # B2B
  if ("b2b" %in% names(df)) {
    b2b_warn  <- createStyle(fgFill = "#FF6B00", fontColour = "#FFFFFF", fontName = "Calibri",
                              fontSize = 10, textDecoration = "bold", halign = "center")
    b2b_clear <- createStyle(fgFill = "#FFFFFF", fontName = "Calibri", fontSize = 10, halign = "center")
    b2b_col   <- ci("b2b")
    b2b_rows  <- which(df$b2b == TRUE)  + 1
    clr_rows  <- which(df$b2b == FALSE) + 1
    if (length(b2b_rows) > 0) addStyle(wb, "Game Environment", b2b_warn,  rows = b2b_rows, cols = b2b_col, stack = TRUE)
    if (length(clr_rows) > 0) addStyle(wb, "Game Environment", b2b_clear, rows = clr_rows, cols = b2b_col, stack = TRUE)
    df$b2b <- ifelse(df$b2b, "B2B", "")
    writeData(wb, "Game Environment", df["b2b"], startRow = 2, startCol = b2b_col, colNames = FALSE)
  }

  # Number formatting
  for (col in intersect(c("implied","ppg_avg","ppg_l10","proj_pace"), names(df)))
    addStyle(wb, "Game Environment", num_1dp, rows = data_rows, cols = ci(col), stack = TRUE)

  setColWidths(wb, "Game Environment", cols = 1:ncols, widths = "auto")
  freezePane(wb, "Game Environment", firstRow = TRUE, firstCol = FALSE)

  # ── Pace key ──
  key_row <- nrows + 3
  key <- list(
    list(label = "PROJ PACE KEY (WNBA)", fill = "#1D3557", font = "#FFFFFF", bold = TRUE),
    list(label = "≤ 93  — Very Slow",    fill = "#C00000", font = "#FFFFFF", bold = FALSE),
    list(label = "94-95 — Slow",         fill = "#F4B8B8", font = "#7B0000", bold = FALSE),
    list(label = "96-97 — Average",      fill = "#FFFFCC", font = "#7F6000", bold = FALSE),
    list(label = "98-99 — Fast",         fill = "#A9D18E", font = "#1D4A1D", bold = FALSE),
    list(label = "100+  — Very Fast",    fill = "#1A7C3E", font = "#FFFFFF", bold = FALSE)
  )
  for (k in seq_along(key)) {
    writeData(wb, "Game Environment", key[[k]]$label, startRow = key_row + k - 1, startCol = 1)
    sty <- createStyle(fgFill = key[[k]]$fill, fontColour = key[[k]]$font,
                       fontName = "Calibri", fontSize = 10,
                       textDecoration = if (key[[k]]$bold) "bold" else NULL,
                       halign = "left")
    addStyle(wb, "Game Environment", sty, rows = key_row + k - 1, cols = 1, stack = FALSE)
  }

  saveWorkbook(wb, path, overwrite = TRUE)
  message("Saved: ", path)
  system(paste("open", shQuote(path)))
  invisible(df)
}

# ── Master runner ──────────────────────────────────────────────────────────────

run_game_env <- function(date = format(Sys.Date(), "%m/%d/%Y")) {
  message("\n=== WNBA Game Environment — ", format(as.Date(date, "%m/%d/%Y"), "%B %d, %Y"), " ===\n")

  df <- build_game_env(date)

  if (is.null(df) || nrow(df) == 0) {
    message("No games found for ", date)
    return(invisible(NULL))
  }

  n_games <- nrow(df) / 2
  message("\n", n_games, " games on the slate — ", nrow(df), " team rows\n")
  print(df %>% select(team, matchup, implied, ppg_avg, plus_minus,
                      opp_papg_rank, opp_drat_rank, proj_pace, tip_off), n = 30)

  export_game_env(df)
  invisible(df)
}

# ── To run ────────────────────────────────────────────────────────────────────
# From the project root:
# source("wnba_game_env.R")
# df <- run_game_env()                   # today
# df <- run_game_env("05/17/2026")       # specific date
