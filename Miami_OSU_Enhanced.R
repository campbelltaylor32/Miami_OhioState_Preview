################################################################################
#                    MIAMI vs OHIO STATE - CFP MATCHUP ANALYSIS                #
#                    Statistical Edge & Vulnerability Finder           #
################################################################################

library(cfbfastR)
library(tidyverse)
library(data.table)
library(RcppRoll)
library(gt)
library(scales)

### Load in Key
Sys.setenv(CFBD_API_KEY = "blank")

### Configuration ###
week_update <- 15
years_download <- 2025
weeks <- seq(1, week_update, 1)
teams_to_compare <- c("Miami", "Ohio State")

### Helper Functions ###
cummean_na <- function(x) {
  n <- cumsum(!is.na(x))
  s <- cumsum(replace(x, is.na(x), 0))
  out <- s / ifelse(n == 0, NA_real_, n)
  out
}

pct_rank <- function(x) {
  percent_rank(x) * 100
}

################################################################################
#                       SECTION 1: DATA COLLECTION                              #
################################################################################

cat("\n========== COLLECTING DATA ==========\n")

### 1A: Get Full Games ###
tot_games <- data.frame()
for (j in seq_along(weeks)) {
  print(paste0("Week ", weeks[j], "..."))
  tmp_games <- cfbd_game_info(year = years_download, week = weeks[j], season_type = "both")
  tot_games <- rbind(tot_games, tmp_games)
}

tot_games <- tot_games %>%
  filter(completed == TRUE) %>%
  select(
    game_id, season, week, season_type, start_date,
    neutral_site, conference_game, venue_id,
    home_id, home_team, home_conference,
    home_points, away_id, away_team, away_conference,
    away_points
  ) %>%
  arrange(season, week)

### 1B: Get Betting Data ###
tot_betting <- data.frame()
for (j in seq_along(weeks)) {
  print(paste0("Betting Week ", weeks[j], "..."))
  tmp_betting <- cfbd_betting_lines(year = years_download, week = weeks[j])
  tmp_betting$week <- weeks[j]
  tot_betting <- rbind(tot_betting, tmp_betting)
}

consensus_only <- tot_betting %>%
  filter(provider == "consensus") %>%
  select(game_id, home_team, away_team, spread, over_under, formatted_spread, week)

fallback <- tot_betting %>%
  group_by(game_id, home_team, away_team, week) %>%
  summarise(
    spread = mean(spread, na.rm = TRUE),
    over_under = mean(over_under, na.rm = TRUE),
    formatted_spread = first(formatted_spread),
    .groups = "drop"
  )

tot_betting_final <- consensus_only %>%
  bind_rows(
    fallback %>% anti_join(consensus_only, by = c("game_id", "home_team", "away_team"))
  )

### 1C: Get Basic Game Stats ###
tot_stats <- data.frame()
# Regular season by week
for (j in seq_along(weeks)) {
  print(paste0("Stats Week ", weeks[j], "..."))
  tmp_stats <- cfbd_game_team_stats(year = years_download, week = weeks[j], season_type = "regular")
  tmp_stats$week <- weeks[j]
  tmp_stats$year <- years_download
  tot_stats <- rbind(tot_stats, tmp_stats)
}
# Postseason separately (no week parameter)
print("Stats Postseason...")
tmp_stats_post <- tryCatch({
  cfbd_game_team_stats(year = years_download, season_type = "postseason", team = "Miami")
}, error = function(e) data.frame())
if (nrow(tmp_stats_post) > 0) {
  tmp_stats_post$week <- 99  # Placeholder for postseason
  tmp_stats_post$year <- years_download
  tot_stats <- rbind(tot_stats, tmp_stats_post)
}

tot_stats <- tot_stats %>%
  select(-c(conference, opponent, opponent_conference)) %>%
  select(game_id, team = school, week, year, everything())

tot_stats <- as.data.frame(tot_stats)
tot_stats[is.na(tot_stats)] <- 0

tot_stats <- tot_stats %>%
  mutate(
    week = as.character(week),
    year = as.character(year),
    game_id = as.character(game_id)
  )

### Parse efficiency columns
tot_stats <- tot_stats %>%
  separate(third_down_eff, into = c("third_down_conversion", "third_down_attempts"), sep = "-") %>%
  separate(fourth_down_eff, into = c("fourth_down_conversion", "fourth_down_attempts"), sep = "-") %>%
  separate(total_penalties_yards, into = c("total_penalties", "penalty_yards"), sep = "-") %>%
  separate(completion_attempts, into = c("completions", "attempted_passes"), sep = "-") %>%
  separate(completion_attempts_allowed, into = c("completions_against", "completion_attempts_against"), sep = "-") %>%
  separate(third_down_eff_allowed, into = c("third_down_conversion_allowed", "third_down_attempts_allowed"), sep = "-") %>%
  separate(fourth_down_eff_allowed, into = c("fourth_down_conversion_allowed", "fourth_down_attempts_allowed"), sep = "-") %>%
  separate(total_penalties_yards_allowed, into = c("penalties_allowed", "penalty_yards_allowed"), sep = "-") %>%
  mutate(
    possession_time = as.numeric(sub("^(\\d+):(\\d+)$", "\\1", possession_time)) +
      as.numeric(sub("^(\\d+):(\\d+)$", "\\2", possession_time)) / 60,
    possession_time_allowed = as.numeric(sub("^(\\d+):(\\d+)$", "\\1", possession_time_allowed)) +
      as.numeric(sub("^(\\d+):(\\d+)$", "\\2", possession_time_allowed)) / 60
  )

tot_stats[, 6:ncol(tot_stats)] <- lapply(tot_stats[, 6:ncol(tot_stats)], as.numeric)

### 1D: Get Play-by-Play EPA Data ###
cat("\nCollecting Play-by-Play EPA Data...\n")
tot_epa <- data.frame()
# Regular season by week
for (j in seq_along(weeks)) {
  print(paste0("PBP Week ", weeks[j], "..."))
  tmp_epa <- tryCatch({
    cfbd_pbp_data(year = years_download, week = weeks[j], epa_wpa = TRUE, season_type = "regular")
  }, error = function(e) data.frame())
  if (nrow(tmp_epa) > 0) {
    tmp_epa$week <- weeks[j]
    tmp_epa$year <- years_download
    tot_epa <- rbind(tot_epa, tmp_epa)
  }
}
# Postseason separately (no week parameter)
print("PBP Postseason...")
tmp_epa_post <- tryCatch({
  cfbd_pbp_data(year = years_download, epa_wpa = TRUE, season_type = "postseason")
}, error = function(e) data.frame())
if (nrow(tmp_epa_post) > 0) {
  tmp_epa_post$week <- 99  # Placeholder for postseason
  tmp_epa_post$year <- years_download
  tot_epa <- rbind(tot_epa, tmp_epa_post)
}

### 1E: Get Advanced Team Stats (filter by team to avoid API limits) ###
cat("\nCollecting Advanced Stats...\n")
advanced_stats_miami <- tryCatch({
  cfbd_stats_season_advanced(year = years_download, team = "Miami")
}, error = function(e) data.frame())
advanced_stats_osu <- tryCatch({
  cfbd_stats_season_advanced(year = years_download, team = "Ohio State")
}, error = function(e) data.frame())
advanced_stats <- rbind(advanced_stats_miami, advanced_stats_osu)

### 1F: Get Team Records ###
cat("\nCollecting Team Records...\n")
team_records <- cfbd_game_records(year = years_download)

### 1G: Get SP+ Ratings ###
cat("\nCollecting Ratings...\n")
sp_ratings <- tryCatch(
  cfbd_ratings_sp(year = years_download),
  error = function(e) NULL
)

################################################################################
#                    SECTION 2: FEATURE ENGINEERING                             #
################################################################################

cat("\n========== ENGINEERING FEATURES ==========\n")

### 2A: Basic Stats Feature Engineering ###
tot_stats <- tot_stats %>%
  mutate(
    third_down_pct_offense = third_down_conversion / third_down_attempts,
    fourth_down_pct_offense = fourth_down_conversion / fourth_down_attempts,
    third_down_pct_defense = third_down_conversion_allowed / third_down_attempts_allowed,
    fourth_down_pct_defense = fourth_down_conversion_allowed / fourth_down_attempts_allowed,
    pressure_pct = qb_hurries / completion_attempts_against,
    sack_pct = sacks / completion_attempts_against,
    pressure_pct_allowed = qb_hurries_allowed / attempted_passes,
    sack_pct_allowed = sacks_allowed / attempted_passes,
    int_rate_offense = passes_intercepted / attempted_passes,
    int_rate_defense = interceptions / completion_attempts_against,
    fumble_rate = fumbles_lost / (rushing_attempts + completions),
    #turnover_margin = turnovers - turnovers_allowed,
    point_differential = points - points_allowed,
    possession_difference = possession_time - possession_time_allowed,
    penalty_yard_margin = penalty_yards - penalty_yards_allowed,
    total_plays = rushing_attempts + attempted_passes,
    rush_percentage = rushing_attempts / total_plays,
    yards_per_play = total_yards / total_plays,
    yards_per_rush = rushing_yards / rushing_attempts,
    yards_per_pass = net_passing_yards / attempted_passes,
    completion_pct = completions / attempted_passes,
    total_plays_against = rushing_attempts_allowed + completion_attempts_against,
    rush_pct_against = rushing_attempts_allowed / total_plays_against,
    yards_per_play_allowed = total_yards_allowed / total_plays_against,
    yards_per_rush_allowed = rushing_yards_allowed / rushing_attempts_allowed,
    yards_per_pass_allowed = net_passing_yards_allowed / completion_attempts_against,
    scoring_efficiency = points / pmax(total_yards / 75, 1),
    big_play_yds_per_completion = net_passing_yards / completions,
    big_play_yds_per_rush = rushing_yards / rushing_attempts,
    points_per_play = points / total_plays,
    points_per_play_allowed = points_allowed / total_plays_against
  )

tot_stats[is.na(tot_stats)] <- 0
tot_stats[sapply(tot_stats, is.infinite)] <- 0

### 2B: EPA Feature Engineering ###
play_types_pass <- c("Sack", "Passing Touchdown", "Interception Return",
                     "Interception Return Touchdown", "Pass Incompletion", "Pass Reception")
play_types_rush <- c("Rush", "Rushing Touchdown")
play_types_all <- c(play_types_pass, play_types_rush)

### Offense EPA Metrics
offense_epa <- tot_epa %>%
  filter(play_type %in% play_types_all) %>%
  group_by(team = pos_team, week, year) %>%
  summarise(
    Total_Offense_Drives = n_distinct(drive_id),
    Total_Offense_Plays = n_distinct(id_play),
    Offense_Run_Plays = n_distinct(id_play[play_type %in% play_types_rush]),
    Offense_Pass_Plays = n_distinct(id_play[play_type %in% play_types_pass]),
    Offense_Pass_Rate = Offense_Pass_Plays / Total_Offense_Plays,
    Total_Offense_EPA = sum(EPA, na.rm = TRUE),
    Offense_EPA_per_Play = Total_Offense_EPA / Total_Offense_Plays,
    Offense_EPA_Rush = sum(EPA[play_type %in% play_types_rush], na.rm = TRUE),
    Offense_EPA_per_Rush = Offense_EPA_Rush / Offense_Run_Plays,
    Offense_EPA_Pass = sum(EPA[play_type %in% play_types_pass], na.rm = TRUE),
    Offense_EPA_per_Pass = Offense_EPA_Pass / Offense_Pass_Plays,
    Offense_Success_Rate = sum(success, na.rm = TRUE) / Total_Offense_Plays,
    Offense_Rush_Success_Rate = sum(success[play_type %in% play_types_rush], na.rm = TRUE) / Offense_Run_Plays,
    Offense_Pass_Success_Rate = sum(success[play_type %in% play_types_pass], na.rm = TRUE) / Offense_Pass_Plays,
    Offense_Explosive_Plays = sum(yards_gained >= 20, na.rm = TRUE),
    Offense_Explosive_Rate = Offense_Explosive_Plays / Total_Offense_Plays,
    Offense_Rush_Explosive = sum(yards_gained >= 20 & play_type %in% play_types_rush, na.rm = TRUE),
    Offense_Pass_Explosive = sum(yards_gained >= 20 & play_type %in% play_types_pass, na.rm = TRUE),
    Offense_Negative_Plays = sum(yards_gained < 0, na.rm = TRUE),
    Offense_Stuff_Rate = Offense_Negative_Plays / Total_Offense_Plays,
    Offense_1st_Down_Success = sum(success[down == 1], na.rm = TRUE) / n_distinct(id_play[down == 1]),
    Offense_2nd_Down_Success = sum(success[down == 2], na.rm = TRUE) / n_distinct(id_play[down == 2]),
    Offense_3rd_Down_Success = sum(success[down == 3], na.rm = TRUE) / n_distinct(id_play[down == 3]),
    Offense_3rd_Down_EPA = sum(EPA[down == 3], na.rm = TRUE) / n_distinct(id_play[down == 3]),
    Offense_Standard_Down_EPA = sum(EPA[(down == 1) | (down == 2 & distance < 8)], na.rm = TRUE) / 
      n_distinct(id_play[(down == 1) | (down == 2 & distance < 8)]),
    Offense_Passing_Down_EPA = sum(EPA[(down == 2 & distance >= 8) | (down == 3 & distance >= 5)], na.rm = TRUE) /
      n_distinct(id_play[(down == 2 & distance >= 8) | (down == 3 & distance >= 5)]),
    Offense_RedZone_Drives = n_distinct(drive_id[yard_line <= 20 & yard_line > 0]),
    Offense_RedZone_EPA = sum(EPA[yard_line <= 20 & yard_line > 0], na.rm = TRUE) / pmax(Offense_RedZone_Drives, 1),
    Offense_RedZone_TD_Rate = n_distinct(drive_id[yard_line <= 20 & play_type %in% 
                                                    c("Rushing Touchdown", "Passing Touchdown")]) / pmax(Offense_RedZone_Drives, 1),
    Offense_Scoring_Drive_Pct = n_distinct(drive_id[drive_scoring == 1]) / Total_Offense_Drives,
    Offense_TD_Drive_Pct = n_distinct(drive_id[drive_pts == 7]) / Total_Offense_Drives,
    Offense_Sacks_Taken = n_distinct(id_play[play_type == "Sack"]),
    Offense_INTs_Thrown = n_distinct(id_play[play_type %in% c("Interception Return", "Interception Return Touchdown")]),
    Offense_Yards_Per_Drive = sum(yards_gained, na.rm = TRUE) / Total_Offense_Drives,
    Offense_Plays_Per_Drive = Total_Offense_Plays / Total_Offense_Drives,
    .groups = "drop"
  )

### Defense EPA Metrics
defense_epa <- tot_epa %>%
  filter(play_type %in% play_types_all) %>%
  group_by(team = def_pos_team, week, year) %>%
  summarise(
    Total_Defense_Drives = n_distinct(drive_id),
    Total_Defense_Plays = n_distinct(id_play),
    Defense_Run_Plays = n_distinct(id_play[play_type %in% play_types_rush]),
    Defense_Pass_Plays = n_distinct(id_play[play_type %in% play_types_pass]),
    Defense_Pass_Rate_Faced = Defense_Pass_Plays / Total_Defense_Plays,
    Total_Defense_EPA = sum(EPA, na.rm = TRUE),
    Defense_EPA_per_Play = Total_Defense_EPA / Total_Defense_Plays,
    Defense_EPA_Rush = sum(EPA[play_type %in% play_types_rush], na.rm = TRUE),
    Defense_EPA_per_Rush = Defense_EPA_Rush / Defense_Run_Plays,
    Defense_EPA_Pass = sum(EPA[play_type %in% play_types_pass], na.rm = TRUE),
    Defense_EPA_per_Pass = Defense_EPA_Pass / Defense_Pass_Plays,
    Defense_Success_Rate = sum(success, na.rm = TRUE) / Total_Defense_Plays,
    Defense_Rush_Success_Rate = sum(success[play_type %in% play_types_rush], na.rm = TRUE) / Defense_Run_Plays,
    Defense_Pass_Success_Rate = sum(success[play_type %in% play_types_pass], na.rm = TRUE) / Defense_Pass_Plays,
    Defense_Explosive_Allowed = sum(yards_gained >= 20, na.rm = TRUE),
    Defense_Explosive_Rate = Defense_Explosive_Allowed / Total_Defense_Plays,
    Defense_Rush_Explosive = sum(yards_gained >= 20 & play_type %in% play_types_rush, na.rm = TRUE),
    Defense_Pass_Explosive = sum(yards_gained >= 20 & play_type %in% play_types_pass, na.rm = TRUE),
    Defense_Negative_Plays = sum(yards_gained < 0, na.rm = TRUE),
    Defense_Stuff_Rate = Defense_Negative_Plays / Total_Defense_Plays,
    Defense_1st_Down_Success = sum(success[down == 1], na.rm = TRUE) / n_distinct(id_play[down == 1]),
    Defense_2nd_Down_Success = sum(success[down == 2], na.rm = TRUE) / n_distinct(id_play[down == 2]),
    Defense_3rd_Down_Success = sum(success[down == 3], na.rm = TRUE) / n_distinct(id_play[down == 3]),
    Defense_3rd_Down_EPA = sum(EPA[down == 3], na.rm = TRUE) / n_distinct(id_play[down == 3]),
    Defense_Standard_Down_EPA = sum(EPA[(down == 1) | (down == 2 & distance < 8)], na.rm = TRUE) / 
      n_distinct(id_play[(down == 1) | (down == 2 & distance < 8)]),
    Defense_Passing_Down_EPA = sum(EPA[(down == 2 & distance >= 8) | (down == 3 & distance >= 5)], na.rm = TRUE) /
      n_distinct(id_play[(down == 2 & distance >= 8) | (down == 3 & distance >= 5)]),
    Defense_RedZone_Drives = n_distinct(drive_id[yard_line <= 20 & yard_line > 0]),
    Defense_RedZone_EPA = sum(EPA[yard_line <= 20 & yard_line > 0], na.rm = TRUE) / pmax(Defense_RedZone_Drives, 1),
    Defense_RedZone_TD_Rate = n_distinct(drive_id[yard_line <= 20 & play_type %in% 
                                                    c("Rushing Touchdown", "Passing Touchdown")]) / pmax(Defense_RedZone_Drives, 1),
    Defense_Scoring_Drive_Pct = n_distinct(drive_id[drive_scoring == 1]) / Total_Defense_Drives,
    Defense_TD_Drive_Pct = n_distinct(drive_id[drive_pts == 7]) / Total_Defense_Drives,
    Defense_Sacks = n_distinct(id_play[play_type == "Sack"]),
    Defense_INTs = n_distinct(id_play[play_type %in% c("Interception Return", "Interception Return Touchdown")]),
    Defense_Yards_Per_Drive_Allowed = sum(yards_gained, na.rm = TRUE) / Total_Defense_Drives,
    Defense_Plays_Per_Drive_Allowed = Total_Defense_Plays / Total_Defense_Drives,
    .groups = "drop"
  )

### 2C: Merge Stats ###
team_epa <- merge(offense_epa, defense_epa, by = c("team", "week", "year"), all = TRUE)
team_epa[is.na(team_epa)] <- 0
all_stats <- merge(tot_stats, team_epa, by = c("team", "week", "year"), all = TRUE)
all_stats[is.na(all_stats)] <- 0

################################################################################
#                    SECTION 3: CALCULATE ROLLING AVERAGES                      #
################################################################################

cat("\n========== CALCULATING ROLLING AVERAGES ==========\n")

numeric_cols <- names(all_stats)[sapply(all_stats, is.numeric)]
numeric_cols <- setdiff(numeric_cols, c("year", "week", "game_id"))

all_stats_rolling <- all_stats %>%
  group_by(year, team) %>%
  arrange(as.numeric(week), .by_group = TRUE) %>%
  mutate(
    across(
      .cols = all_of(numeric_cols),
      .fns = list(
        season_avg = ~ cummean_na(.x),
        last3_avg = ~ roll_mean(.x, n = 3, align = "right", fill = NA, na.rm = TRUE),
        last5_avg = ~ roll_mean(.x, n = 5, align = "right", fill = NA, na.rm = TRUE)
      ),
      .names = "{.col}_{.fn}"
    )
  ) %>%
  ungroup()

### Calculate trends
season_avg_cols <- names(all_stats_rolling)[grepl("_season_avg$", names(all_stats_rolling))]
last3_avg_cols <- names(all_stats_rolling)[grepl("_last3_avg$", names(all_stats_rolling))]
season_bases <- sub("_season_avg$", "", season_avg_cols)
last3_bases <- sub("_last3_avg$", "", last3_avg_cols)
common_bases <- intersect(season_bases, last3_bases)

all_stats_trend <- all_stats_rolling
for (base_metric in common_bases) {
  season_col <- paste0(base_metric, "_season_avg")
  last3_col <- paste0(base_metric, "_last3_avg")
  trend_col <- paste0(base_metric, "_trend")
  all_stats_trend[[trend_col]] <- all_stats_trend[[last3_col]] - all_stats_trend[[season_col]]
}

################################################################################
#                  SECTION 4: MIAMI vs OHIO STATE MATCHUP                       #
################################################################################

cat("\n========== GENERATING MATCHUP ANALYSIS ==========\n")

latest_week <- max(as.numeric(all_stats_rolling$week), na.rm = TRUE) - 1
miami_stats <- all_stats_rolling %>% filter(team == "Miami", week == as.character(latest_week + 1))
ohio_state_stats <- all_stats_rolling %>% filter(team == "Ohio State", week == as.character(15))
all_teams_latest <- all_stats_rolling %>% 
  group_by(team) %>%
  slice_max(week) %>% 
  ungroup()

get_percentile <- function(value, distribution) {
  round(sum(distribution <= value, na.rm = TRUE) / sum(!is.na(distribution)) * 100, 1)
}

################################################################################
#               SECTION 5: ENHANCED COMPARISON REPORT                           #
################################################################################

build_matchup_comparison <- function(miami_data, osu_data, all_teams_data) {
  
  key_metrics <- c(
    # Overall Efficiency
    "Offense_EPA_per_Play_season_avg",
    "Defense_EPA_per_Play_season_avg",
    # Rush Metrics
    "Offense_EPA_per_Rush_season_avg",
    "Defense_EPA_per_Rush_season_avg",
    "Offense_Rush_Success_Rate_season_avg",
    "Defense_Rush_Success_Rate_season_avg",
    # Pass Metrics
    "Offense_EPA_per_Pass_season_avg",
    "Defense_EPA_per_Pass_season_avg",
    "Offense_Pass_Success_Rate_season_avg",
    "Defense_Pass_Success_Rate_season_avg",
    # Explosives
    "Offense_Explosive_Rate_season_avg",
    "Defense_Explosive_Rate_season_avg",
    # Situational
    "Offense_3rd_Down_Success_season_avg",
    "Defense_3rd_Down_Success_season_avg",
    "Offense_RedZone_EPA_season_avg",
    "Defense_RedZone_EPA_season_avg",
    "Offense_RedZone_TD_Rate_season_avg",
    "Defense_RedZone_TD_Rate_season_avg",
    # Pressure
    "sack_pct_season_avg",
    "sack_pct_allowed_season_avg",
    "pressure_pct_season_avg",
    "pressure_pct_allowed_season_avg",
    # Turnovers
    "turnover_margin_season_avg",
    "int_rate_offense_season_avg",
    "int_rate_defense_season_avg",
    # Standard vs Passing Downs
    "Offense_Standard_Down_EPA_season_avg",
    "Offense_Passing_Down_EPA_season_avg",
    "Defense_Standard_Down_EPA_season_avg",
    "Defense_Passing_Down_EPA_season_avg",
    # Scoring
    "Offense_Scoring_Drive_Pct_season_avg",
    "Defense_Scoring_Drive_Pct_season_avg",
    "Offense_TD_Drive_Pct_season_avg",
    "Defense_TD_Drive_Pct_season_avg",
    # Negative Plays
    "Offense_Stuff_Rate_season_avg",
    "Defense_Stuff_Rate_season_avg",
    # Drive Efficiency
    "Offense_Yards_Per_Drive_season_avg",
    "Offense_Plays_Per_Drive_season_avg",
    "Defense_Yards_Per_Drive_Allowed_season_avg",
    # Traditional Stats
    "yards_per_play_season_avg",
    "yards_per_play_allowed_season_avg",
    "completion_pct_season_avg",
    "points_per_play_season_avg",
    "points_per_play_allowed_season_avg",
    # First Down Success
    "Offense_1st_Down_Success_season_avg",
    "Defense_1st_Down_Success_season_avg",
    "Offense_2nd_Down_Success_season_avg",
    "Defense_2nd_Down_Success_season_avg"
  )
  
  comparison_df <- data.frame(
    Metric = key_metrics,
    Miami = NA_real_,
    Ohio_State = NA_real_,
    Miami_Pctl = NA_real_,
    OSU_Pctl = NA_real_,
    Advantage = NA_character_,
    stringsAsFactors = FALSE
  )
  
  for (i in seq_along(key_metrics)) {
    metric <- key_metrics[i]
    if (metric %in% names(miami_data) & metric %in% names(osu_data)) {
      miami_val <- as.numeric(miami_data[[metric]])
      osu_val <- as.numeric(osu_data[[metric]])
      all_vals <- as.numeric(all_teams_data[[metric]])
      
      comparison_df$Miami[i] <- round(miami_val, 4)
      comparison_df$Ohio_State[i] <- round(osu_val, 4)
      comparison_df$Miami_Pctl[i] <- get_percentile(miami_val, all_vals)
      comparison_df$OSU_Pctl[i] <- get_percentile(osu_val, all_vals)
      
      is_defense <- grepl("Defense", metric)
      is_allowed <- grepl("allowed|Against|Stuff_Rate", metric, ignore.case = TRUE)
      
      if (is_defense | is_allowed) {
        if (miami_val < osu_val) {
          comparison_df$Advantage[i] <- "MIAMI"
        } else if (osu_val < miami_val) {
          comparison_df$Advantage[i] <- "OHIO STATE"
        } else {
          comparison_df$Advantage[i] <- "EVEN"
        }
      } else {
        if (miami_val > osu_val) {
          comparison_df$Advantage[i] <- "MIAMI"
        } else if (osu_val > miami_val) {
          comparison_df$Advantage[i] <- "OHIO STATE"
        } else {
          comparison_df$Advantage[i] <- "EVEN"
        }
      }
    }
  }
  
  return(comparison_df)
}

comparison_report <- build_matchup_comparison(miami_stats, ohio_state_stats, all_teams_latest)

################################################################################
#                    SECTION 6: TREND ANALYSIS (MOMENTUM)                       #
################################################################################

build_trend_report <- function(team_data, team_name) {
  
  trend_metrics <- c(
    "Offense_EPA_per_Play", "Defense_EPA_per_Play",
    "Offense_EPA_per_Rush", "Defense_EPA_per_Rush",
    "Offense_EPA_per_Pass", "Defense_EPA_per_Pass",
    "Offense_Success_Rate", "Defense_Success_Rate",
    "Offense_Explosive_Rate", "Defense_Explosive_Rate",
    "Offense_3rd_Down_Success", "Defense_3rd_Down_Success",
    "turnover_margin", "sack_pct", "sack_pct_allowed",
    "Offense_Scoring_Drive_Pct", "Defense_Scoring_Drive_Pct",
    "Offense_RedZone_TD_Rate", "Defense_RedZone_TD_Rate"
  )
  
  trend_df <- data.frame(
    Metric = trend_metrics,
    Season_Avg = NA_real_,
    Last_3_Avg = NA_real_,
    Last_5_Avg = NA_real_,
    Trend_3 = NA_real_,
    Trend_5 = NA_real_,
    Direction = NA_character_,
    stringsAsFactors = FALSE
  )
  
  for (i in seq_along(trend_metrics)) {
    metric <- trend_metrics[i]
    season_col <- paste0(metric, "_season_avg")
    last3_col <- paste0(metric, "_last3_avg")
    last5_col <- paste0(metric, "_last5_avg")
    
    if (season_col %in% names(team_data) & last3_col %in% names(team_data)) {
      season_val <- as.numeric(team_data[[season_col]])
      last3_val <- as.numeric(team_data[[last3_col]])
      last5_val <- if (last5_col %in% names(team_data)) as.numeric(team_data[[last5_col]]) else NA
      
      trend_df$Season_Avg[i] <- round(season_val, 4)
      trend_df$Last_3_Avg[i] <- round(last3_val, 4)
      trend_df$Last_5_Avg[i] <- round(last5_val, 4)
      trend_df$Trend_3[i] <- round(last3_val - season_val, 4)
      trend_df$Trend_5[i] <- if (!is.na(last5_val)) round(last5_val - season_val, 4) else NA
      
      is_defense <- grepl("Defense", metric)
      trend_val <- trend_df$Trend_3[i]
      
      if (is.na(trend_val)) {
        trend_df$Direction[i] <- "N/A"
      } else if (is_defense) {
        trend_df$Direction[i] <- ifelse(trend_val < -0.02, "IMPROVING",
                                        ifelse(trend_val > 0.02, "DECLINING", "STABLE"))
      } else {
        trend_df$Direction[i] <- ifelse(trend_val > 0.02, "IMPROVING",
                                        ifelse(trend_val < -0.02, "DECLINING", "STABLE"))
      }
    }
  }
  
  trend_df$Team <- team_name
  return(trend_df)
}

miami_trends <- build_trend_report(miami_stats, "Miami")
osu_trends <- build_trend_report(ohio_state_stats, "Ohio State")
all_trends <- rbind(miami_trends, osu_trends)

################################################################################
#                 SECTION 7: VULNERABILITY & STRENGTH ANALYSIS                  #
################################################################################

identify_vulnerabilities <- function(team_data, all_teams_data, team_name) {
  
  metrics_to_check <- c(
    "Offense_EPA_per_Play_season_avg", "Offense_EPA_per_Rush_season_avg",
    "Offense_EPA_per_Pass_season_avg", "Offense_Success_Rate_season_avg",
    "Offense_Rush_Success_Rate_season_avg", "Offense_Pass_Success_Rate_season_avg",
    "Offense_Explosive_Rate_season_avg", "Offense_3rd_Down_Success_season_avg",
    "Offense_RedZone_TD_Rate_season_avg", "Offense_Standard_Down_EPA_season_avg",
    "Offense_Passing_Down_EPA_season_avg", "Offense_1st_Down_Success_season_avg",
    "Offense_2nd_Down_Success_season_avg", "Offense_Scoring_Drive_Pct_season_avg",
    "Defense_EPA_per_Play_season_avg", "Defense_EPA_per_Rush_season_avg",
    "Defense_EPA_per_Pass_season_avg", "Defense_Success_Rate_season_avg",
    "Defense_Rush_Success_Rate_season_avg", "Defense_Pass_Success_Rate_season_avg",
    "Defense_Explosive_Rate_season_avg", "Defense_3rd_Down_Success_season_avg",
    "Defense_RedZone_TD_Rate_season_avg", "Defense_Stuff_Rate_season_avg",
    "sack_pct_season_avg", "sack_pct_allowed_season_avg", "turnover_margin_season_avg"
  )
  
  vulnerabilities <- data.frame(
    Metric = character(), Value = numeric(), Percentile = numeric(),
    Category = character(), stringsAsFactors = FALSE
  )
  
  for (metric in metrics_to_check) {
    if (metric %in% names(team_data)) {
      val <- as.numeric(team_data[[metric]])
      all_vals <- as.numeric(all_teams_data[[metric]])
      pctl <- get_percentile(val, all_vals)
      
      is_defense <- grepl("Defense", metric)
      is_allowed <- grepl("allowed", metric, ignore.case = TRUE)
      
      is_vulnerability <- FALSE
      if (is_defense | is_allowed) {
        is_vulnerability <- pctl >= 60
      } else {
        is_vulnerability <- pctl <= 40
      }
      
      if (is_vulnerability) {
        category <- ifelse(grepl("Offense", metric), "Offense",
                           ifelse(grepl("Defense", metric), "Defense", "Special"))
        vulnerabilities <- rbind(vulnerabilities, data.frame(
          Metric = metric, Value = round(val, 4), Percentile = pctl,
          Category = category, stringsAsFactors = FALSE
        ))
      }
    }
  }
  
  vulnerabilities$Team <- team_name
  return(vulnerabilities)
}

identify_strengths <- function(team_data, all_teams_data, team_name) {
  
  metrics_to_check <- c(
    "Offense_EPA_per_Play_season_avg", "Offense_EPA_per_Rush_season_avg",
    "Offense_EPA_per_Pass_season_avg", "Offense_Success_Rate_season_avg",
    "Offense_Rush_Success_Rate_season_avg", "Offense_Pass_Success_Rate_season_avg",
    "Offense_Explosive_Rate_season_avg", "Offense_3rd_Down_Success_season_avg",
    "Offense_RedZone_TD_Rate_season_avg", "Offense_Standard_Down_EPA_season_avg",
    "Offense_Passing_Down_EPA_season_avg", "Offense_1st_Down_Success_season_avg",
    "Offense_2nd_Down_Success_season_avg", "Offense_Scoring_Drive_Pct_season_avg",
    "Defense_EPA_per_Play_season_avg", "Defense_EPA_per_Rush_season_avg",
    "Defense_EPA_per_Pass_season_avg", "Defense_Success_Rate_season_avg",
    "Defense_Rush_Success_Rate_season_avg", "Defense_Pass_Success_Rate_season_avg",
    "Defense_Explosive_Rate_season_avg", "Defense_3rd_Down_Success_season_avg",
    "Defense_RedZone_TD_Rate_season_avg", "Defense_Stuff_Rate_season_avg",
    "sack_pct_season_avg", "sack_pct_allowed_season_avg", "turnover_margin_season_avg"
  )
  
  strengths <- data.frame(
    Metric = character(), Value = numeric(), Percentile = numeric(),
    Category = character(), stringsAsFactors = FALSE
  )
  
  for (metric in metrics_to_check) {
    if (metric %in% names(team_data)) {
      val <- as.numeric(team_data[[metric]])
      all_vals <- as.numeric(all_teams_data[[metric]])
      pctl <- get_percentile(val, all_vals)
      
      is_defense <- grepl("Defense", metric)
      is_allowed <- grepl("allowed", metric, ignore.case = TRUE)
      
      is_strength <- FALSE
      if (is_defense | is_allowed) {
        is_strength <- pctl <= 30
      } else {
        is_strength <- pctl >= 70
      }
      
      if (is_strength) {
        category <- ifelse(grepl("Offense", metric), "Offense",
                           ifelse(grepl("Defense", metric), "Defense", "Special"))
        strengths <- rbind(strengths, data.frame(
          Metric = metric, Value = round(val, 4), Percentile = pctl,
          Category = category, stringsAsFactors = FALSE
        ))
      }
    }
  }
  
  strengths$Team <- team_name
  return(strengths)
}

miami_vulnerabilities <- identify_vulnerabilities(miami_stats, all_teams_latest, "Miami")
osu_vulnerabilities <- identify_vulnerabilities(ohio_state_stats, all_teams_latest, "Ohio State")
miami_strengths <- identify_strengths(miami_stats, all_teams_latest, "Miami")
osu_strengths <- identify_strengths(ohio_state_stats, all_teams_latest, "Ohio State")

################################################################################
#               SECTION 8: KEY MATCHUP EDGES                                   #
################################################################################

find_edges <- function(team1_strengths, team2_vulnerabilities, team1_name, team2_name) {
  
  edges <- data.frame(
    Matchup = character(), Team1_Strength = character(), Team1_Value = numeric(),
    Team1_Pctl = numeric(), Team2_Weakness = character(), Team2_Value = numeric(),
    Team2_Pctl = numeric(), Edge_For = character(), stringsAsFactors = FALSE
  )
  
  matchup_map <- list(
    "Offense_EPA_per_Play_season_avg" = "Defense_EPA_per_Play_season_avg",
    "Offense_EPA_per_Rush_season_avg" = "Defense_EPA_per_Rush_season_avg",
    "Offense_EPA_per_Pass_season_avg" = "Defense_EPA_per_Pass_season_avg",
    "Offense_Success_Rate_season_avg" = "Defense_Success_Rate_season_avg",
    "Offense_Rush_Success_Rate_season_avg" = "Defense_Rush_Success_Rate_season_avg",
    "Offense_Pass_Success_Rate_season_avg" = "Defense_Pass_Success_Rate_season_avg",
    "Offense_Explosive_Rate_season_avg" = "Defense_Explosive_Rate_season_avg",
    "Offense_3rd_Down_Success_season_avg" = "Defense_3rd_Down_Success_season_avg",
    "Offense_RedZone_TD_Rate_season_avg" = "Defense_RedZone_TD_Rate_season_avg",
    "sack_pct_season_avg" = "sack_pct_allowed_season_avg"
  )
  
  for (i in seq_len(nrow(team1_strengths))) {
    strength_metric <- team1_strengths$Metric[i]
    
    if (strength_metric %in% names(matchup_map)) {
      defense_metric <- matchup_map[[strength_metric]]
    } else if (strength_metric %in% unlist(matchup_map)) {
      defense_metric <- names(matchup_map)[which(unlist(matchup_map) == strength_metric)]
    } else {
      next
    }
    
    if (defense_metric %in% team2_vulnerabilities$Metric) {
      vuln_row <- team2_vulnerabilities[team2_vulnerabilities$Metric == defense_metric, ]
      edges <- rbind(edges, data.frame(
        Matchup = gsub("_season_avg", "", strength_metric),
        Team1_Strength = team1_strengths$Metric[i],
        Team1_Value = team1_strengths$Value[i],
        Team1_Pctl = team1_strengths$Percentile[i],
        Team2_Weakness = vuln_row$Metric,
        Team2_Value = vuln_row$Value,
        Team2_Pctl = vuln_row$Percentile,
        Edge_For = team1_name,
        stringsAsFactors = FALSE
      ))
    }
  }
  
  return(edges)
}

miami_edges <- find_edges(miami_strengths, osu_vulnerabilities, "Miami", "Ohio State")
osu_edges <- find_edges(osu_strengths, miami_vulnerabilities, "Ohio State", "Miami")
all_edges <- rbind(miami_edges, osu_edges)

################################################################################
#                    SECTION 9: GAME SCRIPT ANALYSIS                           #
################################################################################

game_script_analysis <- function(pbp_data, team_name) {
  
  team_plays <- pbp_data %>% filter(pos_team == team_name)
  
  script_stats <- team_plays %>%
    mutate(
      game_script = case_when(
        score_diff >= 14 ~ "Blowout_Lead",
        score_diff >= 7 ~ "Comfortable_Lead",
        score_diff >= 1 ~ "Close_Lead",
        score_diff == 0 ~ "Tied",
        score_diff >= -6 ~ "Close_Deficit",
        score_diff >= -13 ~ "Moderate_Deficit",
        TRUE ~ "Large_Deficit"
      )
    ) %>%
    group_by(game_script) %>%
    summarise(
      Plays = n(),
      EPA_per_Play = mean(EPA, na.rm = TRUE),
      Success_Rate = mean(success, na.rm = TRUE),
      Pass_Rate = sum(play_type %in% play_types_pass) / n(),
      Rush_Rate = sum(play_type %in% play_types_rush) / n(),
      Explosive_Rate = mean(yards_gained >= 20, na.rm = TRUE),
      .groups = "drop"
    )
  
  script_stats$Team <- team_name
  return(script_stats)
}

miami_game_script <- game_script_analysis(tot_epa, "Miami")
osu_game_script <- game_script_analysis(tot_epa, "Ohio State")
game_script_comparison <- rbind(miami_game_script, osu_game_script)

################################################################################
#                 SECTION 10: QUARTER-BY-QUARTER PERFORMANCE                    #
################################################################################

calculate_quarter_performance <- function(pbp_data, team_name) {
  
  team_plays <- pbp_data %>% filter(pos_team == team_name, !is.na(EPA))
  
  quarter_stats <- team_plays %>%
    group_by(game_id, period) %>%
    summarise(
      plays = n(),
      total_epa = sum(EPA, na.rm = TRUE),
      epa_per_play = mean(EPA, na.rm = TRUE),
      success_rate = mean(success, na.rm = TRUE),
      explosive_plays = sum(yards_gained >= 20, na.rm = TRUE),
      negative_plays = sum(yards_gained < 0, na.rm = TRUE),
      pass_rate = mean(play_type %in% play_types_pass, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    group_by(period) %>%
    summarise(
      games = n_distinct(game_id),
      avg_plays = mean(plays, na.rm = TRUE),
      avg_epa_per_play = mean(epa_per_play, na.rm = TRUE),
      avg_success_rate = mean(success_rate, na.rm = TRUE),
      avg_explosives = mean(explosive_plays, na.rm = TRUE),
      avg_pass_rate = mean(pass_rate, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(team = team_name)
  
  return(quarter_stats)
}

miami_quarters <- calculate_quarter_performance(tot_epa, "Miami")
osu_quarters <- calculate_quarter_performance(tot_epa, "Ohio State")
quarter_analysis <- rbind(miami_quarters, osu_quarters)

################################################################################
#                 SECTION 11: FIELD POSITION ANALYSIS                           #
################################################################################

calculate_field_position_performance <- function(pbp_data, team_name) {
  
  team_plays <- pbp_data %>%
    filter(pos_team == team_name, !is.na(EPA), !is.na(yard_line))
  
  field_zones <- team_plays %>%
    mutate(
      field_zone = case_when(
        yard_line <= 20 ~ "Own 0-20 (Backed Up)",
        yard_line <= 40 ~ "Own 21-40",
        yard_line <= 60 ~ "Midfield (41-60)",
        yard_line <= 80 ~ "Opponent 21-40",
        yard_line > 80 ~ "Red Zone (Opp 0-20)",
        TRUE ~ "Unknown"
      )
    ) %>%
    group_by(field_zone) %>%
    summarise(
      plays = n(),
      epa_per_play = mean(EPA, na.rm = TRUE),
      success_rate = mean(success, na.rm = TRUE),
      explosive_rate = mean(yards_gained >= 20, na.rm = TRUE),
      pass_rate = mean(play_type %in% play_types_pass, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(team = team_name)
  
  return(field_zones)
}

miami_field_position <- calculate_field_position_performance(tot_epa, "Miami")
osu_field_position <- calculate_field_position_performance(tot_epa, "Ohio State")
field_position_analysis <- rbind(miami_field_position, osu_field_position)

################################################################################
#                 SECTION 12: DOWN & DISTANCE ANALYSIS                          #
################################################################################

calculate_down_distance_analysis <- function(pbp_data, team_name) {
  
  team_plays <- pbp_data %>%
    filter(pos_team == team_name, !is.na(EPA), down %in% c(1, 2, 3, 4))
  
  situations <- team_plays %>%
    mutate(
      situation = case_when(
        down == 1 ~ "1st Down",
        down == 2 & distance <= 3 ~ "2nd & Short (1-3)",
        down == 2 & distance <= 6 ~ "2nd & Medium (4-6)",
        down == 2 & distance > 6 ~ "2nd & Long (7+)",
        down == 3 & distance <= 3 ~ "3rd & Short (1-3)",
        down == 3 & distance <= 6 ~ "3rd & Medium (4-6)",
        down == 3 & distance > 6 ~ "3rd & Long (7+)",
        down == 4 ~ "4th Down",
        TRUE ~ "Other"
      )
    ) %>%
    group_by(situation) %>%
    summarise(
      plays = n(),
      epa_per_play = mean(EPA, na.rm = TRUE),
      success_rate = mean(success, na.rm = TRUE),
      pass_rate = mean(play_type %in% play_types_pass, na.rm = TRUE),
      rush_rate = mean(play_type %in% play_types_rush, na.rm = TRUE),
      explosive_rate = mean(yards_gained >= 20, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(team = team_name)
  
  return(situations)
}

miami_situations <- calculate_down_distance_analysis(tot_epa, "Miami")
osu_situations <- calculate_down_distance_analysis(tot_epa, "Ohio State")
situation_analysis <- rbind(miami_situations, osu_situations)

################################################################################
#                 SECTION 13: DRIVE EFFICIENCY                                  #
################################################################################

calculate_drive_efficiency <- function(pbp_data, team_name) {
  
  team_drives <- pbp_data %>%
    filter(pos_team == team_name, !is.na(EPA)) %>%
    group_by(game_id, drive_id) %>%
    summarise(
      plays = n(),
      total_yards = sum(yards_gained, na.rm = TRUE),
      total_epa = sum(EPA, na.rm = TRUE),
      scoring = max(drive_scoring, na.rm = TRUE),
      points = max(drive_pts, na.rm = TRUE),
      start_yard = first(yard_line),
      .groups = "drop"
    )
  
  drive_summary <- team_drives %>%
    summarise(
      total_drives = n(),
      avg_plays_per_drive = mean(plays, na.rm = TRUE),
      avg_yards_per_drive = mean(total_yards, na.rm = TRUE),
      avg_epa_per_drive = mean(total_epa, na.rm = TRUE),
      scoring_pct = mean(scoring, na.rm = TRUE),
      points_per_drive = mean(points, na.rm = TRUE),
      td_pct = mean(points >= 6, na.rm = TRUE),
      fg_pct = mean(points == 3, na.rm = TRUE),
      avg_start_position = mean(start_yard, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  return(drive_summary)
}

miami_drives <- calculate_drive_efficiency(tot_epa, "Miami")
osu_drives <- calculate_drive_efficiency(tot_epa, "Ohio State")
drive_efficiency <- rbind(miami_drives, osu_drives)

################################################################################
#                 SECTION 14: PRESSURE ANALYSIS                                 #
################################################################################

calculate_pressure_metrics <- function(pbp_data, team_name) {
  
  # Offense protection
  pass_plays_off <- pbp_data %>%
    filter(pos_team == team_name, play_type %in% play_types_pass, !is.na(EPA))
  
  protection <- pass_plays_off %>%
    summarise(
      total_dropbacks = n(),
      sacks_allowed = sum(sack == 1, na.rm = TRUE),
      sack_rate_allowed = sacks_allowed / total_dropbacks,
      epa_sack_plays = mean(EPA[sack == 1], na.rm = TRUE),
      epa_clean_plays = mean(EPA[sack == 0 | is.na(sack)], na.rm = TRUE),
      negative_pass_plays = sum(EPA < 0, na.rm = TRUE),
      negative_pass_rate = negative_pass_plays / total_dropbacks
    ) %>%
    mutate(team = team_name, side = "Protection")
  
  # Defense pass rush
  pass_plays_def <- pbp_data %>%
    filter(def_pos_team == team_name, play_type %in% play_types_pass, !is.na(EPA))
  
  pass_rush <- pass_plays_def %>%
    summarise(
      total_dropbacks_faced = n(),
      sacks_generated = sum(sack == 1, na.rm = TRUE),
      sack_rate_generated = sacks_generated / total_dropbacks_faced,
      negative_plays_forced = sum(EPA < 0, na.rm = TRUE),
      disruption_rate = negative_plays_forced / total_dropbacks_faced,
      epa_allowed_per_dropback = mean(EPA, na.rm = TRUE)
    ) %>%
    mutate(team = team_name, side = "Pass Rush")
  
  return(list(protection = protection, pass_rush = pass_rush))
}

miami_pressure <- calculate_pressure_metrics(tot_epa, "Miami")
osu_pressure <- calculate_pressure_metrics(tot_epa, "Ohio State")

require(plyr)
pressure_analysis <- rbind.fill(
  miami_pressure$protection, miami_pressure$pass_rush,
  osu_pressure$protection, osu_pressure$pass_rush
)


################################################################################
#                 SECTION 15: LATE/CLOSE GAME ANALYSIS                          #
################################################################################
detach("package:plyr", unload = TRUE)

calculate_late_close <- function(pbp_data, team_name) {
  
  # Late = 4th quarter, Close = within 8 points
  late_close <- pbp_data %>%
    filter(pos_team == team_name, !is.na(EPA)) %>%
    mutate(
      is_late = period >= 4,
      is_close = abs(score_diff) <= 8,
      is_clutch = is_late & is_close
    )
  
  clutch_stats <- late_close %>%
    group_by(is_clutch) %>%
    summarise(
      plays = n(),
      epa_per_play = mean(EPA, na.rm = TRUE),
      success_rate = mean(success, na.rm = TRUE),
      pass_rate = mean(play_type %in% play_types_pass, na.rm = TRUE),
      explosive_rate = mean(yards_gained >= 20, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      situation = ifelse(is_clutch, "Clutch (Q4 & Close)", "Non-Clutch"),
      team = team_name
    ) %>%
    select(-is_clutch)
  
  return(clutch_stats)
}

miami_late_close <- calculate_late_close(tot_epa, "Miami")
osu_late_close <- calculate_late_close(tot_epa, "Ohio State")
late_close_analysis <- rbind(miami_late_close, osu_late_close)

################################################################################
#                 SECTION 16: TEMPO ANALYSIS                                    #
################################################################################

calculate_tempo <- function(pbp_data, team_name) {
  
  team_plays <- pbp_data %>% filter(pos_team == team_name, !is.na(EPA))
  
  tempo_stats <- team_plays %>%
    group_by(game_id) %>%
    summarise(total_plays = n(), .groups = "drop") %>%
    summarise(
      avg_plays_per_game = mean(total_plays, na.rm = TRUE),
      min_plays = min(total_plays, na.rm = TRUE),
      max_plays = max(total_plays, na.rm = TRUE),
      sd_plays = sd(total_plays, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  return(tempo_stats)
}

miami_tempo <- calculate_tempo(tot_epa, "Miami")
osu_tempo <- calculate_tempo(tot_epa, "Ohio State")
tempo_analysis <- rbind(miami_tempo, osu_tempo)

################################################################################
#                 SECTION 17: LINE OF SCRIMMAGE ANALYSIS                        #
################################################################################
# Deep dive on O-Line vs D-Line matchups

cat("\n========== LINE OF SCRIMMAGE ANALYSIS ==========\n")

### 17A: Run Game Progression Through Quarters ###
# Does the O-line take control as game progresses?
calculate_run_progression <- function(pbp_data, team_name) {
  
  rush_plays <- pbp_data %>%
    filter(pos_team == team_name, play_type %in% play_types_rush, !is.na(EPA))
  
  quarter_progression <- rush_plays %>%
    group_by(game_id, period) %>%
    summarise(
      rush_plays = n(),
      rush_yards = sum(yards_gained, na.rm = TRUE),
      rush_epa = sum(EPA, na.rm = TRUE),
      rush_epa_per_play = mean(EPA, na.rm = TRUE),
      rush_success_rate = mean(success, na.rm = TRUE),
      yards_per_rush = mean(yards_gained, na.rm = TRUE),
      stuff_rate = mean(yards_gained <= 0, na.rm = TRUE),
      explosive_rush_rate = mean(yards_gained >= 10, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    group_by(period) %>%
    summarise(
      games = n_distinct(game_id),
      avg_rush_plays = mean(rush_plays, na.rm = TRUE),
      avg_rush_epa = mean(rush_epa_per_play, na.rm = TRUE),
      avg_rush_success = mean(rush_success_rate, na.rm = TRUE),
      avg_ypc = mean(yards_per_rush, na.rm = TRUE),
      avg_stuff_rate = mean(stuff_rate, na.rm = TRUE),
      avg_explosive_rate = mean(explosive_rush_rate, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(team = team_name)
  
  # Calculate half-by-half and trend
  first_half <- quarter_progression %>% filter(period <= 2)
  second_half <- quarter_progression %>% filter(period > 2 & period <= 4)
  
  progression_trend <- data.frame(
    team = team_name,
    first_half_rush_success = mean(first_half$avg_rush_success, na.rm = TRUE),
    second_half_rush_success = mean(second_half$avg_rush_success, na.rm = TRUE),
    first_half_ypc = mean(first_half$avg_ypc, na.rm = TRUE),
    second_half_ypc = mean(second_half$avg_ypc, na.rm = TRUE),
    first_half_stuff_rate = mean(first_half$avg_stuff_rate, na.rm = TRUE),
    second_half_stuff_rate = mean(second_half$avg_stuff_rate, na.rm = TRUE),
    rush_success_trend = mean(second_half$avg_rush_success, na.rm = TRUE) - mean(first_half$avg_rush_success, na.rm = TRUE),
    ypc_trend = mean(second_half$avg_ypc, na.rm = TRUE) - mean(first_half$avg_ypc, na.rm = TRUE),
    stuff_rate_trend = mean(second_half$avg_stuff_rate, na.rm = TRUE) - mean(first_half$avg_stuff_rate, na.rm = TRUE)
  )
  
  return(list(by_quarter = quarter_progression, trend = progression_trend))
}

miami_run_prog <- calculate_run_progression(tot_epa, "Miami")
osu_run_prog <- calculate_run_progression(tot_epa, "Ohio State")

run_progression_quarters <- rbind(miami_run_prog$by_quarter, osu_run_prog$by_quarter)
run_progression_trend <- rbind(miami_run_prog$trend, osu_run_prog$trend)

### 17B: Run Game in Obvious Run Situations ###
# Short yardage, winning late, clock management
calculate_obvious_run_situations <- function(pbp_data, team_name) {
  
  team_plays <- pbp_data %>%
    filter(pos_team == team_name, !is.na(EPA))
  
  # Define situations
  obvious_run <- team_plays %>%
    mutate(
      # Short yardage situations
      short_yardage = (down %in% c(3, 4) & distance <= 2) | (down == 1 & distance <= 3),
      
      # Winning late (4th quarter, up by 1-14 points)
      winning_late = period >= 4 & score_diff >= 1 & score_diff <= 14,
      
      # Clock killing (4th quarter, leading, under 5 min - approximated)
      clock_killing = period >= 4 & score_diff >= 7,
      
      # Goal line (inside 5 yard line)
      goal_line = yard_line >= 95,
      
      # 2nd and short
      second_short = down == 2 & distance <= 3,
      
      is_rush = play_type %in% play_types_rush
    )
  
  # Short yardage performance
  short_yardage_stats <- obvious_run %>%
    filter(short_yardage) %>%
    summarise(
      situation = "Short Yardage (3rd/4th & ≤2)",
      total_plays = n(),
      rush_plays = sum(is_rush, na.rm = TRUE),
      rush_rate = rush_plays / total_plays,
      rush_success = mean(success[is_rush], na.rm = TRUE),
      rush_epa = mean(EPA[is_rush], na.rm = TRUE),
      conversion_rate = mean(yards_gained >= distance, na.rm = TRUE),
      avg_yards = mean(yards_gained[is_rush], na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # Winning late performance
  winning_late_stats <- obvious_run %>%
    filter(winning_late) %>%
    summarise(
      situation = "Winning Late (Q4, +1-14)",
      total_plays = n(),
      rush_plays = sum(is_rush, na.rm = TRUE),
      rush_rate = rush_plays / total_plays,
      rush_success = mean(success[is_rush], na.rm = TRUE),
      rush_epa = mean(EPA[is_rush], na.rm = TRUE),
      conversion_rate = mean(success, na.rm = TRUE),
      avg_yards = mean(yards_gained[is_rush], na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # Clock killing performance
  clock_kill_stats <- obvious_run %>%
    filter(clock_killing) %>%
    summarise(
      situation = "Clock Killing (Q4, +7+)",
      total_plays = n(),
      rush_plays = sum(is_rush, na.rm = TRUE),
      rush_rate = rush_plays / total_plays,
      rush_success = mean(success[is_rush], na.rm = TRUE),
      rush_epa = mean(EPA[is_rush], na.rm = TRUE),
      conversion_rate = mean(success, na.rm = TRUE),
      avg_yards = mean(yards_gained[is_rush], na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # Goal line performance
  goal_line_stats <- obvious_run %>%
    filter(goal_line) %>%
    summarise(
      situation = "Goal Line (Inside 5)",
      total_plays = n(),
      rush_plays = sum(is_rush, na.rm = TRUE),
      rush_rate = rush_plays / total_plays,
      rush_success = mean(success[is_rush], na.rm = TRUE),
      rush_epa = mean(EPA[is_rush], na.rm = TRUE),
      conversion_rate = mean(play_type %in% c("Rushing Touchdown", "Passing Touchdown"), na.rm = TRUE),
      avg_yards = mean(yards_gained[is_rush], na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # Second and short
  second_short_stats <- obvious_run %>%
    filter(second_short) %>%
    summarise(
      situation = "2nd & Short (≤3)",
      total_plays = n(),
      rush_plays = sum(is_rush, na.rm = TRUE),
      rush_rate = rush_plays / total_plays,
      rush_success = mean(success[is_rush], na.rm = TRUE),
      rush_epa = mean(EPA[is_rush], na.rm = TRUE),
      conversion_rate = mean(yards_gained >= distance, na.rm = TRUE),
      avg_yards = mean(yards_gained[is_rush], na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  all_situations <- rbind(short_yardage_stats, winning_late_stats, clock_kill_stats, 
                          goal_line_stats, second_short_stats)
  
  return(all_situations)
}

miami_obvious_run <- calculate_obvious_run_situations(tot_epa, "Miami")
osu_obvious_run <- calculate_obvious_run_situations(tot_epa, "Ohio State")
obvious_run_analysis <- rbind(miami_obvious_run, osu_obvious_run)

### 17C: Defensive Line - Run Defense Progression ###
# Does stuff rate increase as game goes on? Does run D hold up?
calculate_run_defense_progression <- function(pbp_data, team_name) {
  
  rush_plays_against <- pbp_data %>%
    filter(def_pos_team == team_name, play_type %in% play_types_rush, !is.na(EPA))
  
  quarter_progression <- rush_plays_against %>%
    group_by(game_id, period) %>%
    summarise(
      rush_plays_faced = n(),
      rush_yards_allowed = sum(yards_gained, na.rm = TRUE),
      rush_epa_allowed = sum(EPA, na.rm = TRUE),
      rush_epa_per_play_allowed = mean(EPA, na.rm = TRUE),
      rush_success_allowed = mean(success, na.rm = TRUE),
      yards_per_rush_allowed = mean(yards_gained, na.rm = TRUE),
      stuff_rate = mean(yards_gained <= 0, na.rm = TRUE),
      tfl_rate = mean(yards_gained < 0, na.rm = TRUE),
      explosive_rush_allowed = mean(yards_gained >= 10, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    group_by(period) %>%
    summarise(
      games = n_distinct(game_id),
      avg_rush_plays_faced = mean(rush_plays_faced, na.rm = TRUE),
      avg_rush_epa_allowed = mean(rush_epa_per_play_allowed, na.rm = TRUE),
      avg_rush_success_allowed = mean(rush_success_allowed, na.rm = TRUE),
      avg_ypc_allowed = mean(yards_per_rush_allowed, na.rm = TRUE),
      avg_stuff_rate = mean(stuff_rate, na.rm = TRUE),
      avg_tfl_rate = mean(tfl_rate, na.rm = TRUE),
      avg_explosive_allowed = mean(explosive_rush_allowed, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(team = team_name)
  
  # Calculate trend
  first_half <- quarter_progression %>% filter(period <= 2)
  second_half <- quarter_progression %>% filter(period > 2 & period <= 4)
  
  defense_trend <- data.frame(
    team = team_name,
    first_half_stuff_rate = mean(first_half$avg_stuff_rate, na.rm = TRUE),
    second_half_stuff_rate = mean(second_half$avg_stuff_rate, na.rm = TRUE),
    first_half_ypc_allowed = mean(first_half$avg_ypc_allowed, na.rm = TRUE),
    second_half_ypc_allowed = mean(second_half$avg_ypc_allowed, na.rm = TRUE),
    first_half_success_allowed = mean(first_half$avg_rush_success_allowed, na.rm = TRUE),
    second_half_success_allowed = mean(second_half$avg_rush_success_allowed, na.rm = TRUE),
    stuff_rate_trend = mean(second_half$avg_stuff_rate, na.rm = TRUE) - mean(first_half$avg_stuff_rate, na.rm = TRUE),
    ypc_allowed_trend = mean(second_half$avg_ypc_allowed, na.rm = TRUE) - mean(first_half$avg_ypc_allowed, na.rm = TRUE),
    run_defense_trend = mean(first_half$avg_rush_epa_allowed, na.rm = TRUE) - mean(second_half$avg_rush_epa_allowed, na.rm = TRUE)  # Positive = improving
  )
  
  return(list(by_quarter = quarter_progression, trend = defense_trend))
}

miami_run_def_prog <- calculate_run_defense_progression(tot_epa, "Miami")
osu_run_def_prog <- calculate_run_defense_progression(tot_epa, "Ohio State")

run_defense_quarters <- rbind(miami_run_def_prog$by_quarter, osu_run_def_prog$by_quarter)
run_defense_trend <- rbind(miami_run_def_prog$trend, osu_run_def_prog$trend)

### 17D: Run Defense in Obvious Run Situations ###
calculate_run_defense_obvious <- function(pbp_data, team_name) {
  
  team_plays <- pbp_data %>%
    filter(def_pos_team == team_name, !is.na(EPA))
  
  obvious_run_def <- team_plays %>%
    mutate(
      short_yardage = (down %in% c(3, 4) & distance <= 2) | (down == 1 & distance <= 3),
      opponent_winning_late = period >= 4 & score_diff <= -1 & score_diff >= -14,
      opponent_clock_killing = period >= 4 & score_diff <= -7,
      goal_line = yard_line >= 95,
      is_rush = play_type %in% play_types_rush
    )
  
  # Short yardage defense
  short_yardage_def <- obvious_run_def %>%
    filter(short_yardage, is_rush) %>%
    summarise(
      situation = "Short Yardage Defense",
      plays = n(),
      stuff_rate = mean(yards_gained <= 0, na.rm = TRUE),
      success_allowed = mean(success, na.rm = TRUE),
      epa_allowed = mean(EPA, na.rm = TRUE),
      conversion_allowed = mean(yards_gained >= distance, na.rm = TRUE),
      avg_yards_allowed = mean(yards_gained, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # Opponent clock killing defense
  clock_kill_def <- obvious_run_def %>%
    filter(opponent_clock_killing, is_rush) %>%
    summarise(
      situation = "vs Clock Killing Rush",
      plays = n(),
      stuff_rate = mean(yards_gained <= 0, na.rm = TRUE),
      success_allowed = mean(success, na.rm = TRUE),
      epa_allowed = mean(EPA, na.rm = TRUE),
      conversion_allowed = mean(success, na.rm = TRUE),
      avg_yards_allowed = mean(yards_gained, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # Goal line defense
  goal_line_def <- obvious_run_def %>%
    filter(goal_line, is_rush) %>%
    summarise(
      situation = "Goal Line Rush Defense",
      plays = n(),
      stuff_rate = mean(yards_gained <= 0, na.rm = TRUE),
      success_allowed = mean(success, na.rm = TRUE),
      epa_allowed = mean(EPA, na.rm = TRUE),
      conversion_allowed = mean(play_type == "Rushing Touchdown", na.rm = TRUE),
      avg_yards_allowed = mean(yards_gained, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  all_def_situations <- rbind(short_yardage_def, clock_kill_def, goal_line_def)
  return(all_def_situations)
}

miami_run_def_obvious <- calculate_run_defense_obvious(tot_epa, "Miami")
osu_run_def_obvious <- calculate_run_defense_obvious(tot_epa, "Ohio State")
run_defense_obvious <- rbind(miami_run_def_obvious, osu_run_def_obvious)

### 17E: Pass Protection in Obvious Pass Situations ###
calculate_protection_obvious_pass <- function(pbp_data, team_name) {
  
  pass_plays <- pbp_data %>%
    filter(pos_team == team_name, play_type %in% play_types_pass, !is.na(EPA))
  
  obvious_pass <- pass_plays %>%
    mutate(
      # Obvious pass situations
      third_long = down == 3 & distance >= 7,
      second_long = down == 2 & distance >= 8,
      trailing_late = period >= 4 & score_diff <= -7,
      two_minute = period %in% c(2, 4) & score_diff <= 7,  # Approximation
      passing_down = (down == 2 & distance >= 8) | (down == 3 & distance >= 5)
    )
  
  # 3rd and long protection
  third_long_prot <- obvious_pass %>%
    filter(third_long) %>%
    summarise(
      situation = "3rd & Long (7+)",
      plays = n(),
      sack_rate = mean(sack == 1, na.rm = TRUE),
      success_rate = mean(success, na.rm = TRUE),
      epa_per_play = mean(EPA, na.rm = TRUE),
      conversion_rate = mean(yards_gained >= distance, na.rm = TRUE),
      negative_play_rate = mean(yards_gained < 0, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # Trailing late protection
  trailing_late_prot <- obvious_pass %>%
    filter(trailing_late) %>%
    summarise(
      situation = "Trailing Late (Q4, -7+)",
      plays = n(),
      sack_rate = mean(sack == 1, na.rm = TRUE),
      success_rate = mean(success, na.rm = TRUE),
      epa_per_play = mean(EPA, na.rm = TRUE),
      conversion_rate = mean(success, na.rm = TRUE),
      negative_play_rate = mean(yards_gained < 0, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # All passing downs
  passing_down_prot <- obvious_pass %>%
    filter(passing_down) %>%
    summarise(
      situation = "All Passing Downs",
      plays = n(),
      sack_rate = mean(sack == 1, na.rm = TRUE),
      success_rate = mean(success, na.rm = TRUE),
      epa_per_play = mean(EPA, na.rm = TRUE),
      conversion_rate = mean(yards_gained >= distance, na.rm = TRUE),
      negative_play_rate = mean(yards_gained < 0, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # 2nd and long
  second_long_prot <- obvious_pass %>%
    filter(second_long) %>%
    summarise(
      situation = "2nd & Long (8+)",
      plays = n(),
      sack_rate = mean(sack == 1, na.rm = TRUE),
      success_rate = mean(success, na.rm = TRUE),
      epa_per_play = mean(EPA, na.rm = TRUE),
      conversion_rate = mean(yards_gained >= distance/2, na.rm = TRUE),  # Staying on schedule
      negative_play_rate = mean(yards_gained < 0, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  all_protection <- rbind(third_long_prot, trailing_late_prot, passing_down_prot, second_long_prot)
  return(all_protection)
}

miami_protection_obvious <- calculate_protection_obvious_pass(tot_epa, "Miami")
osu_protection_obvious <- calculate_protection_obvious_pass(tot_epa, "Ohio State")
protection_obvious_pass <- rbind(miami_protection_obvious, osu_protection_obvious)

### 17F: Pass Rush in Obvious Pass Situations ###
calculate_pass_rush_obvious <- function(pbp_data, team_name) {
  
  pass_plays <- pbp_data %>%
    filter(def_pos_team == team_name, play_type %in% play_types_pass, !is.na(EPA))
  
  obvious_pass <- pass_plays %>%
    mutate(
      third_long = down == 3 & distance >= 7,
      second_long = down == 2 & distance >= 8,
      opponent_trailing_late = period >= 4 & score_diff >= 7,
      passing_down = (down == 2 & distance >= 8) | (down == 3 & distance >= 5)
    )
  
  # 3rd and long pass rush
  third_long_rush <- obvious_pass %>%
    filter(third_long) %>%
    summarise(
      situation = "3rd & Long Pass Rush",
      plays = n(),
      sack_rate = mean(sack == 1, na.rm = TRUE),
      success_allowed = mean(success, na.rm = TRUE),
      epa_allowed = mean(EPA, na.rm = TRUE),
      conversion_allowed = mean(yards_gained >= distance, na.rm = TRUE),
      negative_play_rate = mean(yards_gained < 0, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # Opponent trailing late
  opponent_trailing_rush <- obvious_pass %>%
    filter(opponent_trailing_late) %>%
    summarise(
      situation = "vs Trailing Late Pass",
      plays = n(),
      sack_rate = mean(sack == 1, na.rm = TRUE),
      success_allowed = mean(success, na.rm = TRUE),
      epa_allowed = mean(EPA, na.rm = TRUE),
      conversion_allowed = mean(success, na.rm = TRUE),
      negative_play_rate = mean(yards_gained < 0, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  # All passing downs
  passing_down_rush <- obvious_pass %>%
    filter(passing_down) %>%
    summarise(
      situation = "Passing Down Pass Rush",
      plays = n(),
      sack_rate = mean(sack == 1, na.rm = TRUE),
      success_allowed = mean(success, na.rm = TRUE),
      epa_allowed = mean(EPA, na.rm = TRUE),
      conversion_allowed = mean(yards_gained >= distance, na.rm = TRUE),
      negative_play_rate = mean(yards_gained < 0, na.rm = TRUE)
    ) %>%
    mutate(team = team_name)
  
  all_pass_rush <- rbind(third_long_rush, opponent_trailing_rush, passing_down_rush)
  return(all_pass_rush)
}

miami_pass_rush_obvious <- calculate_pass_rush_obvious(tot_epa, "Miami")
osu_pass_rush_obvious <- calculate_pass_rush_obvious(tot_epa, "Ohio State")
pass_rush_obvious <- rbind(miami_pass_rush_obvious, osu_pass_rush_obvious)

### 17G: Line of Scrimmage Summary Metrics ###
calculate_los_summary <- function(pbp_data, team_name) {
  
  # Offensive line summary
  off_plays <- pbp_data %>%
    filter(pos_team == team_name, !is.na(EPA))
  
  rush_plays <- off_plays %>% filter(play_type %in% play_types_rush)
  pass_plays <- off_plays %>% filter(play_type %in% play_types_pass)
  
  oline_summary <- data.frame(
    team = team_name,
    side = "O-Line",
    # Run blocking
    rush_ypc = mean(rush_plays$yards_gained, na.rm = TRUE),
    rush_success_rate = mean(rush_plays$success, na.rm = TRUE),
    rush_epa_per_play = mean(rush_plays$EPA, na.rm = TRUE),
    stuff_rate_faced = mean(rush_plays$yards_gained <= 0, na.rm = TRUE),
    explosive_rush_rate = mean(rush_plays$yards_gained >= 10, na.rm = TRUE),
    # Pass protection
    sack_rate = mean(pass_plays$sack == 1, na.rm = TRUE),
    negative_pass_rate = mean(pass_plays$yards_gained < 0, na.rm = TRUE),
    pass_epa_per_play = mean(pass_plays$EPA, na.rm = TRUE)
  )
  
  # Defensive line summary
  def_plays <- pbp_data %>%
    filter(def_pos_team == team_name, !is.na(EPA))
  
  def_rush <- def_plays %>% filter(play_type %in% play_types_rush)
  def_pass <- def_plays %>% filter(play_type %in% play_types_pass)
  
  dline_summary <- data.frame(
    team = team_name,
    side = "D-Line",
    # Run defense
    rush_ypc_allowed = mean(def_rush$yards_gained, na.rm = TRUE),
    rush_success_allowed = mean(def_rush$success, na.rm = TRUE),
    rush_epa_allowed = mean(def_rush$EPA, na.rm = TRUE),
    stuff_rate = mean(def_rush$yards_gained <= 0, na.rm = TRUE),
    tfl_rate = mean(def_rush$yards_gained < 0, na.rm = TRUE),
    # Pass rush
    sack_rate = mean(def_pass$sack == 1, na.rm = TRUE),
    disruption_rate = mean(def_pass$yards_gained < 0, na.rm = TRUE),
    pass_epa_allowed = mean(def_pass$EPA, na.rm = TRUE)
  )
  
  return(list(oline = oline_summary, dline = dline_summary))
}

miami_los <- calculate_los_summary(tot_epa, "Miami")
osu_los <- calculate_los_summary(tot_epa, "Ohio State")

los_summary <- bind_rows(
  miami_los$oline, miami_los$dline,
  osu_los$oline, osu_los$dline
)

################################################################################
#                    SECTION 18: SAVE ALL OUTPUT FILES                          #
################################################################################

cat("\n\n=== SAVING OUTPUT FILES ===\n")

write.csv(comparison_report, "Miami_vs_OhioState_Comparison.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Comparison.csv\n")

all_vulnerabilities <- rbind(miami_vulnerabilities, osu_vulnerabilities)
write.csv(all_vulnerabilities, "Miami_vs_OhioState_Vulnerabilities.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Vulnerabilities.csv\n")

all_strengths <- rbind(miami_strengths, osu_strengths)
write.csv(all_strengths, "Miami_vs_OhioState_Strengths.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Strengths.csv\n")

write.csv(all_edges, "Miami_vs_OhioState_Edges.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Edges.csv\n")

write.csv(all_trends, "Miami_vs_OhioState_Trends.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Trends.csv\n")

write.csv(game_script_comparison, "Miami_vs_OhioState_GameScript.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_GameScript.csv\n")

write.csv(quarter_analysis, "Miami_vs_OhioState_Quarter_Analysis.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Quarter_Analysis.csv\n")

write.csv(situation_analysis, "Miami_vs_OhioState_Situation_Analysis.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Situation_Analysis.csv\n")

write.csv(field_position_analysis, "Miami_vs_OhioState_Field_Position.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Field_Position.csv\n")

write.csv(drive_efficiency, "Miami_vs_OhioState_Drive_Efficiency.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Drive_Efficiency.csv\n")

write.csv(pressure_analysis, "Miami_vs_OhioState_Pressure_Analysis.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Pressure_Analysis.csv\n")

write.csv(late_close_analysis, "Miami_vs_OhioState_Late_Close.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Late_Close.csv\n")

write.csv(tempo_analysis, "Miami_vs_OhioState_Tempo.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Tempo.csv\n")

# LINE OF SCRIMMAGE ANALYSIS FILES
write.csv(run_progression_quarters, "Miami_vs_OhioState_Run_Progression_Quarters.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Run_Progression_Quarters.csv\n")

write.csv(run_progression_trend, "Miami_vs_OhioState_Run_Progression_Trend.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Run_Progression_Trend.csv\n")

write.csv(obvious_run_analysis, "Miami_vs_OhioState_Obvious_Run.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Obvious_Run.csv\n")

write.csv(run_defense_quarters, "Miami_vs_OhioState_Run_Defense_Quarters.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Run_Defense_Quarters.csv\n")

write.csv(run_defense_trend, "Miami_vs_OhioState_Run_Defense_Trend.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Run_Defense_Trend.csv\n")

write.csv(run_defense_obvious, "Miami_vs_OhioState_Run_Defense_Obvious.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Run_Defense_Obvious.csv\n")

write.csv(protection_obvious_pass, "Miami_vs_OhioState_Protection_Obvious_Pass.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Protection_Obvious_Pass.csv\n")

write.csv(pass_rush_obvious, "Miami_vs_OhioState_Pass_Rush_Obvious.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Pass_Rush_Obvious.csv\n")

write.csv(los_summary, "Miami_vs_OhioState_LOS_Summary.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_LOS_Summary.csv\n")

teams_full_data <- all_stats_rolling %>% filter(team %in% teams_to_compare)
write.csv(teams_full_data, "Miami_vs_OhioState_Full_Season_Data.csv", row.names = FALSE)
cat("Saved: Miami_vs_OhioState_Full_Season_Data.csv\n")

