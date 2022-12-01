library(dplyr); options(dplyr.summarise.inform = F)

##################################################


DP_1_2_1 <- function(data,
                     ResearchPeriod, unit,
                     var_facet, var_group)
{
  begin_date <- min(data$screen_date, na.rm = T)
  end_date <- max(data$screen_date, na.rm = T)
  
  suppressMessages({
    df <- data %>%
      rename(Date = screen_date) %>%
      select_at(vars(Date, var_facet, var_group)) %>% na.omit %>%
      mutate(Date = factor(as.character(Date), levels = as.character(seq(min(ResearchPeriod[1], begin_date), max(ResearchPeriod[2], end_date), by = 'day')))) %>%
      group_by_at(vars(Date, var_facet, var_group), .drop = F) %>% summarise(N_Screen_Act = n()) %>% ungroup %>%
      mutate(Date = as.Date(fast_strptime(as.character(Date), format = '%Y-%m-%d'))) %>%
      mutate(Week = sprintf('Week %d', as.integer(ceiling(difftime(Date, ResearchPeriod[1]-1, units = 'weeks')))),
             Month = sprintf('Month %d', 12*(year(Date) - year(ResearchPeriod[1])) + (month(Date) - month(ResearchPeriod[1])) + 1),
             Quarter = sprintf('Quarter %d', 4*(year(Date) - year(ResearchPeriod[1])) + (quarter(Date) - quarter(ResearchPeriod[1])) + 1),
             Year = sprintf('Year %d', (year(Date) - year(ResearchPeriod[1])) + 1),
             .after = Date) %>%
      select(-setdiff(c("Week", "Month", "Quarter", "Year"), unit)) %>%
      group_by_at(unit) %>% mutate(Date_Start = first(Date), Date_End = last(Date), .after = Date) %>% ungroup %>%
      group_by_at(vars(Date_Start, Date_End, unit, var_facet, var_group)) %>% summarise(N_Screen_Act = sum(N_Screen_Act)) %>% ungroup %>%
      group_by_at(vars(var_facet, var_group)) %>% mutate(CumN_Screen_Act = cumsum(N_Screen_Act)) %>% ungroup
  })
  
  return(df)
}


##################################################


DP_1_2_2 <- function(data, data_Tar,
                     ResearchPeriod, unit,
                     var_facet, var_group)
{
  begin_date <- min(data$enrolled_date, na.rm = T)
  end_date <- max(data$enrolled_date, na.rm = T)
  
  N_facet <- ifelse(is.null(var_facet), 1, nlevels(pull(data, var_facet)))
  N_group <- ifelse(is.null(var_group), 1, nlevels(pull(data, var_group)))
  
  suppressMessages({
    df <- data %>%
      rename(Date = enrolled_date) %>%
      select_at(vars(Date, var_facet, var_group)) %>% na.omit %>%
      mutate(Date = factor(as.character(Date), levels = as.character(seq(min(ResearchPeriod[1], begin_date), max(ResearchPeriod[2], end_date), by = 'day')))) %>%
      group_by_at(vars(Date, var_facet, var_group), .drop = F) %>% summarise(N_Enrolled_Act = n()) %>% ungroup %>%
      mutate(Date = as.Date(fast_strptime(as.character(Date), format = '%Y-%m-%d'))) %>%
      mutate(Week = sprintf('Week %d', as.integer(ceiling(difftime(Date, ResearchPeriod[1]-1, units = 'weeks')))),
             Month = sprintf('Month %d', 12*(year(Date) - year(ResearchPeriod[1])) + (month(Date) - month(ResearchPeriod[1])) + 1),
             Quarter = sprintf('Quarter %d', 4*(year(Date) - year(ResearchPeriod[1])) + (quarter(Date) - quarter(ResearchPeriod[1])) + 1),
             Year = sprintf('Year %d', (year(Date) - year(ResearchPeriod[1])) + 1),
             .after = Date) %>%
      select(-setdiff(c("Week", "Month", "Quarter", "Year"), unit)) %>%
      group_by_at(unit) %>% mutate(Date_Start = first(Date), Date_End = last(Date), .after = Date) %>% ungroup %>%
      group_by_at(vars(Date_Start, Date_End, unit, var_facet, var_group)) %>% summarise(N_Enrolled_Act = sum(N_Enrolled_Act)) %>% ungroup %>%
      group_by_at(vars(var_facet, var_group)) %>% mutate(CumN_Enrolled_Act = cumsum(N_Enrolled_Act)) %>% ungroup %>%
      left_join(data_Tar %>% select(Date, CumN_Enrolled_Tar) %>% rename(Date_End = Date)) %>% mutate(CumN_Enrolled_Tar = round(CumN_Enrolled_Tar/N_facet/N_group, 1)) %>% relocate(CumN_Enrolled_Tar, .before = CumN_Enrolled_Act) %>%
      group_by_at(vars(var_facet, var_group)) %>% mutate(N_Enrolled_Tar = round(rev(CumN_Enrolled_Tar)[1]/n(), 1), .before = N_Enrolled_Act) %>% ungroup %>%
      group_by_at(vars(var_facet, var_group)) %>% mutate(CumN_Enrolled_Tar = cumsum(N_Enrolled_Tar)) %>% ungroup %>%
      mutate(N_Enrolled_Achieve = round(N_Enrolled_Act/N_Enrolled_Tar, 2), .after = N_Enrolled_Act) %>%
      mutate(CumN_Enrolled_Achieve = round(CumN_Enrolled_Act/CumN_Enrolled_Tar, 2), .after = CumN_Enrolled_Act)
  })
  
  return(df)
}


##################################################


DP_1_2_3 <- function(data,
                     ResearchPeriod, unit,
                     var_facet, var_group)
{
  begin_date <- min(data$screen_date, na.rm = T)
  end_date <- max(data$screen_date, na.rm = T)
  
  
  suppressMessages({
    df <- data %>%
      filter(!is.na(is_enrolled)) %>%
      rename(Date = screen_date) %>%
      select_at(vars(Date, var_facet, var_group)) %>% na.omit %>%
      mutate(Date = factor(as.character(Date), levels = as.character(seq(min(ResearchPeriod[1], begin_date), max(ResearchPeriod[2], end_date), by = 'day')))) %>%
      group_by_at(vars(Date, var_facet, var_group), .drop = F) %>% summarise(N_Eligible_Act = n()) %>% ungroup %>%
      mutate(Date = as.Date(fast_strptime(as.character(Date), format = '%Y-%m-%d'))) %>%
      mutate(Week = sprintf('Week %d', as.integer(ceiling(difftime(Date, ResearchPeriod[1]-1, units = 'weeks')))),
             Month = sprintf('Month %d', 12*(year(Date) - year(ResearchPeriod[1])) + (month(Date) - month(ResearchPeriod[1])) + 1),
             Quarter = sprintf('Quarter %d', 4*(year(Date) - year(ResearchPeriod[1])) + (quarter(Date) - quarter(ResearchPeriod[1])) + 1),
             Year = sprintf('Year %d', (year(Date) - year(ResearchPeriod[1])) + 1),
             .after = Date) %>%
      select(-setdiff(c("Week", "Month", "Quarter", "Year"), unit)) %>%
      group_by_at(unit) %>% mutate(Date_Start = first(Date), Date_End = last(Date), .after = Date) %>% ungroup %>%
      group_by_at(vars(Date_Start, Date_End, unit, var_facet, var_group)) %>% summarise(N_Eligible_Act = sum(N_Eligible_Act)) %>% ungroup %>%
      group_by_at(vars(var_facet, var_group)) %>% mutate(CumN_Eligible_Act = cumsum(N_Eligible_Act)) %>% ungroup
  })
  
  return(df)
}


##################################################



