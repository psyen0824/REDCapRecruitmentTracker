library(ggplot2)
library(plotly)
library(dplyr); options(dplyr.summarise.inform = F)
library(tidyr)
##################################################
# Plot Function for Tracking Module: Screened Data 
Plot_1_2_1 <- function(data,
                       ResearchPeriod, ObservedPeriod, unit,
                       var_facet, var_group)
{
  df_plot <- data %>%
    filter(ObservedPeriod[1] <= Date_End & ObservedPeriod[2] >= Date_Start) %>%
    select(-c(Date_Start, N_Screen_Act))
  
  p <- ggplot(df_plot, aes(text = str_c(get(unit), tryCatch(str_c(', ', get(var_group)), error = function(x) ''), ': ', CumN_Screen_Act))) +
    geom_line(aes(x = Date_End, y = CumN_Screen_Act, color = tryCatch(get(var_group), error = function(x) NULL), group = interaction(tryCatch(get(var_facet), error = function(x) 1), tryCatch(get(var_group), error = function(x) 1))), alpha = 0.7) +
    scale_x_date(breaks = pretty(data$Date_End)-1, limits = ResearchPeriod, date_labels = '%Y-%m') +
    facet_wrap(var_facet) +
    labs(x = NULL, y = 'Cumulative Number of Screened Participants') +
    theme_bw() +
    theme(legend.title = element_blank(),
          panel.spacing = unit(1, "lines"),
          strip.background = element_rect(fill = 'black'),
          strip.text = element_text(color = 'white', size = 12, face = 'bold'),
          axis.text.x = element_text(angle = 45, hjust = 1))
  
  gp <- ggplotly(p, tooltip = c('text')) %>%
    layout(margin = list(l = 80, b = 60),
           legend = list(x = 1.02, y = 0.5),
           hovermode = 'x unified') %>%
    config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F)

  if (unit %in% c('Quarter', 'Year')) {
    for (i in 1:length(gp$x$data)) {
      gp$x$data[[i]]$mode <- 'markers+lines'
    }
  }

  for (i in 1:length(gp$x$data)) {
    gp$x$data[[i]]$x <- as.Date(gp$x$data[[i]]$x, origin = '1970-01-01')
  }
  for (i in str_which(names(gp$x$layout), pattern = 'xaxis')) {
    gp$x$layout[[i]]$range <- as.Date(gp$x$layout[[i]]$range, origin = '1970-01-01')
    gp$x$layout[[i]]$tickvals <- as.Date(gp$x$layout[[i]]$tickvals, origin = '1970-01-01')
    gp$x$layout[[i]]$type <- 'date'
    gp$x$layout[[i]]$tickformat <- '%Y-%m'
    gp$x$layout[[i]]$hoverformat <- '%Y-%m-%d'
  }
  
  gp$x$layout$annotations[[1]]$x <- -0.06
  gp <- style(gp, hoverinfo = 'x+text')

  return(gp)
}

##################################################
# Plot Function for Tracking Module: Randomized Data 
Plot_1_2_2 <- function(data,
                       ResearchPeriod, ObservedPeriod, unit,
                       var_facet, var_group)
{
  df_plot <- data %>%
    filter(ObservedPeriod[1] <= Date_End & ObservedPeriod[2] >= Date_Start) %>%
    select(Date_End, unit, var_facet, var_group, CumN_Enrolled_Tar, CumN_Enrolled_Act) %>%
    gather(key = 'Type', value = 'CumN_Enrolled', CumN_Enrolled_Tar, CumN_Enrolled_Act) %>%
    mutate(Type = recode(Type, CumN_Enrolled_Tar = 'Target', CumN_Enrolled_Act = 'Actual'))
  
  p <- ggplot(df_plot, aes(text = str_c(get(unit), ', ', tryCatch(str_c(get(var_group), Type, sep = ', '), error = function(x) Type), ': ', CumN_Enrolled))) +
    geom_line(aes(x = Date_End, y = CumN_Enrolled, color = tryCatch(get(var_group), error = function(x) NULL), linetype = Type, group = interaction(Type, tryCatch(get(var_facet), error = function(x) 1), tryCatch(get(var_group), error = function(x) 1))), alpha = 0.7) +
    scale_x_date(breaks = pretty(data$Date_End)-1, limits = ResearchPeriod, date_labels = '%Y-%m') +
    labs(x = NULL, y = 'Cumulative Number of Randomized Participants') +
    facet_wrap(var_facet) +
    theme_bw() +
    theme(legend.title = element_blank(),
          panel.spacing = unit(1, 'lines'),
          strip.background = element_rect(fill = 'black'),
          strip.text = element_text(color = 'white', size = 12, face = 'bold'),
          axis.text.x = element_text(angle = 45, hjust = 1))
  
  gp <- ggplotly(p, tooltip = c('text')) %>%
    layout(margin = list(l = 80, b = 60),
           legend = list(x = 1.02, y = 0.5),
           hovermode = 'x unified') %>%
    config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F)

  if (unit %in% c('Quarter', 'Year')) {
    for (i in str_which(sapply(gp$x$data, function(x) x$name), pattern = 'Actual')) {
      gp$x$data[[i]]$mode <- 'markers+lines'
    }
  }

  for (i in 1:length(gp$x$data)) {
    gp$x$data[[i]]$x <- as.Date(gp$x$data[[i]]$x, origin = '1970-01-01')
  }
  for (i in str_which(names(gp$x$layout), pattern = 'xaxis')) {
    gp$x$layout[[i]]$range <- as.Date(gp$x$layout[[i]]$range, origin = '1970-01-01')
    gp$x$layout[[i]]$tickvals <- as.Date(gp$x$layout[[i]]$tickvals, origin = '1970-01-01')
    gp$x$layout[[i]]$type <- 'date'
    gp$x$layout[[i]]$tickformat <- '%Y-%m'
    gp$x$layout[[i]]$hoverformat <- '%Y-%m-%d'
  }

  gp$x$layout$annotations[[1]]$x <- -0.06
  gp <- style(gp, hoverinfo = 'x+text')

  return(gp)
}

##################################################
# Plot Function for Tracking Module: Eligible Data 
Plot_1_2_3 <- function(data,
                       ResearchPeriod, ObservedPeriod, unit,
                       var_facet, var_group)
{
  df_plot <- data %>%
    filter(ObservedPeriod[1] <= Date_End & ObservedPeriod[2] >= Date_Start) %>%
    select(-c(Date_Start, N_Eligible_Act))
  
  p <- ggplot(df_plot, aes(text = str_c(get(unit), tryCatch(str_c(', ', get(var_group)), error = function(x) ''), ': ', CumN_Eligible_Act))) +
    geom_line(aes(x = Date_End, y = CumN_Eligible_Act, color = tryCatch(get(var_group), error = function(x) NULL), group = interaction(tryCatch(get(var_facet), error = function(x) 1), tryCatch(get(var_group), error = function(x) 1))), alpha = 0.7) +
    scale_x_date(breaks = pretty(data$Date_End)-1, limits = ResearchPeriod, date_labels = '%Y-%m') +
    facet_wrap(var_facet) +
    labs(x = NULL, y = 'Cumulative Number of Eligible Participants') +
    theme_bw() +
    theme(legend.title = element_blank(),
          panel.spacing = unit(1, "lines"),
          strip.background = element_rect(fill = 'black'),
          strip.text = element_text(color = 'white', size = 12, face = 'bold'),
          axis.text.x = element_text(angle = 45, hjust = 1))
  
  gp <- ggplotly(p, tooltip = c('text')) %>%
    layout(margin = list(l = 80, b = 60),
           legend = list(x = 1.02, y = 0.5),
           hovermode = 'x unified') %>%
    config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F)
  
  if (unit %in% c('Quarter', 'Year')) {
    for (i in 1:length(gp$x$data)) {
      gp$x$data[[i]]$mode <- 'markers+lines'
    }
  }
  
  for (i in 1:length(gp$x$data)) {
    gp$x$data[[i]]$x <- as.Date(gp$x$data[[i]]$x, origin = '1970-01-01')
  }
  for (i in str_which(names(gp$x$layout), pattern = 'xaxis')) {
    gp$x$layout[[i]]$range <- as.Date(gp$x$layout[[i]]$range, origin = '1970-01-01')
    gp$x$layout[[i]]$tickvals <- as.Date(gp$x$layout[[i]]$tickvals, origin = '1970-01-01')
    gp$x$layout[[i]]$type <- 'date'
    gp$x$layout[[i]]$tickformat <- '%Y-%m'
    gp$x$layout[[i]]$hoverformat <- '%Y-%m-%d'
  }
  
  gp$x$layout$annotations[[1]]$x <- -0.06
  gp <- style(gp, hoverinfo = 'x+text')
  
  return(gp)
}

##################################################
# Plot Function for Descriptive Statistics Module: Univariate Analysis 
Plot_1_3_1 <- function(data,
                       var, varType, Vars,
                       nbins = NULL)
{
  if (varType == 'cate') {
    
    df_cat <- data %>%
      group_by_at(vars(var)) %>% summarise(Count = n()) %>% ungroup
    
    nTotal <- sum(df_cat$Count)
    nMiss <- sum(df_cat$Count[is.na(pull(df_cat, var))])
    nValid <- nTotal - nMiss
    
    df_cat <- df_cat %>%
      na.omit() %>%
      mutate(Percent = sprintf('%.1f%%', 100*Count/sum(Count)))
    
    p <- ggplot(df_cat, aes(label = Percent)) +
      geom_bar(aes_string(x = var, y = 'Count'), stat = 'identity', col = 'blue', fill = 'cadetblue2') +
      scale_x_discrete(label = sprintf('%s\n(Valid N = %d)', pull(df_cat, var), pull(df_cat, Count))) +
      labs(x = NULL, y = NULL, title = sprintf('%s\n(N Total %d = Valid %d + Missing %d)', names(Vars)[Vars == var], nTotal, nValid, nMiss)) +
      theme_bw() +
      theme(plot.title = element_text(size = 12, hjust = 0.5))
    
  } else if (varType == 'num') {
    
    df_num <- data %>%
      select(var)
    
    nTotal <- nrow(df_num)
    nMiss <- sum(is.na(pull(df_num, var)))
    nValid <- nTotal - nMiss
    
    df_num <- df_num %>%
      na.omit()
    
    x <- pull(df_num, var)
    breaks <- seq(from = floor(min(x)), by = ceiling((ceiling(max(x)) - floor(min(x))) / nbins), length = nbins + 1)
    
    p_pre <- ggplot(df_num) +
      geom_histogram(aes_string(x = var), breaks = breaks, closed = 'left')
    p <- ggplot(ggplot_build(p_pre)$data[[1]], aes(text = sprintf('range: [%d, %d)', xmin, xmax))) +
      geom_bar(aes(x = factor(xmin), y = count), stat = 'identity', width = 1, position = position_nudge(x = 0.5), col = "red", fill = "bisque") +
      labs(x = NULL, y = NULL, title = sprintf('%s\n(N Total %d = Valid %d + Missing %d, Mean = %.2f, SD = %.2f, Median = %.2f, IQR = %.2f)', names(Vars)[Vars == var], nTotal, nValid, nMiss, mean(x), sd(x), median(x), IQR(x))) +
      theme_bw() +
      theme(plot.title = element_text(size = 12, hjust = 0.5))
    
  }
  
  return(p)
}

##################################################
# Plot Function for Descriptive Statistics Module: Bivariate Analysis 
Plot_1_3_2 <- function(data,
                       var1, var2, var1Type, var2Type, Vars)
{
  if (var1Type == 'cate') {
    if (var2Type == 'cate') {
      df_catcat <- data %>%
        group_by_at(vars(var1, var2)) %>% summarise(Count = n()) %>% ungroup
      
      nTotal <- sum(df_catcat$Count)
      nMiss <- sum(df_catcat$Count[is.na(pull(df_catcat, var1)) | is.na(pull(df_catcat, var2))])
      nValid <- nTotal - nMiss
      
      df_catcat <- df_catcat %>%
        na.omit() %>%
        group_by_at(vars(var1)) %>% mutate(Percent = sprintf('%.1f%%', 100*Count/sum(Count))) %>% ungroup()
      
      p <- ggplot(df_catcat, aes(label = Percent)) +
        geom_bar(aes_string(x = var1, y = 'Count', fill = var2), stat = 'identity', position = 'dodge', col = 'blue') +
        coord_flip() +
        scale_x_discrete(label = sprintf('%s\n(Valid N = %d)', levels(pull(df_catcat, var1)), tapply(df_catcat$Count, INDEX = pull(df_catcat, var1), sum))) +
        scale_fill_brewer() +
        labs(x = NULL, y = NULL, title = sprintf('%s by %s\n(N Total %d = Valid %d + Missing %d)', names(Vars)[Vars == var1], names(Vars)[Vars == var2], nTotal, nValid, nMiss)) +
        theme_bw() +
        theme(plot.title = element_text(hjust = 0.5),
              legend.title = element_blank())
    } else {
      df_catnum <- data %>%
        select(var1, var2)
      
      nTotal <- nrow(df_catnum)
      nMiss <- sum(is.na(pull(df_catnum, var1)) | is.na(pull(df_catnum, var2)))
      nValid <- nTotal - nMiss
      
      df_catnum <- df_catnum %>%
        na.omit()
      
      p <- ggplot(df_catnum, aes_string(x = var1, y = var2, fill = var1)) +
        geom_boxplot() +
        scale_fill_brewer(palette = 'Dark2') +
        scale_x_discrete(label = sprintf('%s\n(Valid N = %d)', levels(pull(df_catnum, var1)), table(pull(df_catnum, var1)))) +
        labs(x = NULL, y = names(Vars)[Vars == var2], title = sprintf('%s by %s\n(N Total %d = Valid %d + Missing %d)', names(Vars)[Vars == var1], names(Vars)[Vars == var2], nTotal, nValid, nMiss)) +
        theme_bw() +
        theme(plot.title = element_text(hjust = 0.5),
              legend.title = element_blank())
    }
  } else {
    if (var2Type == 'cate') {
      df_numcat <- data %>%
        select(var1, var2)
      
      nTotal <- nrow(df_numcat)
      nMiss <- sum(is.na(pull(df_numcat, var1)) | is.na(pull(df_numcat, var2)))
      nValid <- nTotal - nMiss
      
      df_numcat <- df_numcat %>%
        na.omit()
      
      p <- ggplot(df_numcat, aes_string(x = var2, y = var1, fill = var2)) +
        geom_boxplot() +
        scale_fill_brewer(palette = 'Dark2') +
        scale_x_discrete(label = sprintf('%s\n(Valid N = %d)', levels(pull(df_numcat, var2)), table(pull(df_numcat, var2)))) +
        labs(x = NULL, y = names(Vars)[Vars == var1], title = sprintf('%s by %s\n(N Total %d = Valid %d + Missing %d)', names(Vars)[Vars == var1], names(Vars)[Vars == var2], nTotal, nValid, nMiss)) +
        theme_bw() +
        theme(plot.title = element_text(hjust = 0.5),
              legend.title = element_blank())
    } else {
      df_numnum <- data %>%
        select(var1, var2)
      
      nTotal <- nrow(df_numnum)
      nMiss <- sum(is.na(pull(df_numnum, var1)) | is.na(pull(df_numnum, var2)))
      nValid <- nTotal - nMiss
      
      df_numnum <- df_numnum %>%
        na.omit()
      
      p <- ggplot(df_numnum, aes_string(x = var1, y = var2)) +
        geom_point() +
        geom_smooth() +
        labs(x = names(Vars)[Vars == var1], y = names(Vars)[Vars == var2], title = sprintf('%s by %s\n(N Total %d = Valid %d + Missing %d)', names(Vars)[Vars == var1], names(Vars)[Vars == var2], nTotal, nValid, nMiss)) +
        theme_bw() +
        theme(plot.title = element_text(hjust = 0.5))
    }
  }
  
  return(p)
}