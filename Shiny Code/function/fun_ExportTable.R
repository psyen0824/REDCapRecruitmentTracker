library(openxlsx)
library(dplyr)
library(tidyr)
options('openxlsx.minWidth' = 11)
##################################################
# Excel Function for Tracking Module: Screened Data 
ET_1_2_1 <- function(file, data, Vars,
                     ResearchPeriod, ObservedPeriod, unit,
                     var_facet, var_group)
{
  N_facet <- ifelse(is.null(var_facet), 1, nlevels(pull(data, var_facet)))
  N_group <- ifelse(is.null(var_group), 1, nlevels(pull(data, var_group)))
  
  df_ET <- data %>%
    mutate(N_Screen_Act = replace(N_Screen_Act, Date_End < ObservedPeriod[1] | Date_Start > ObservedPeriod[2], NA),
           CumN_Screen_Act = replace(CumN_Screen_Act, Date_End < ObservedPeriod[1] | Date_Start > ObservedPeriod[2], NA))
  
  wb <- createWorkbook()
  
  # Create summary worksheet
  addWorksheet(wb, sheetName = 'Summary', gridLines = F)
  
  if (N_facet > 1 & N_group > 1) {
    df_summary <- df_ET %>%
      na.omit() %>%
      tail(N_facet*N_group) %>%
      select_at(vars(var_facet, var_group, 'CumN_Screen_Act')) %>%
      spread(key = var_group, value = 'CumN_Screen_Act')
    colnames(df_summary)[1] <- 'Summary'
    
    df_summary$Summary <- as.character(df_summary$Summary)
    df_summary$Total <- rowSums(df_summary[,-1])
    df_summary <- rbind(df_summary, c(list('Total'), unname(as.list(colSums(df_summary[,-1])))))
    
    addStyle(wb, sheet = 'Summary', rows = 1:(2+N_facet+1), cols = 1:(2+N_group+1), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    writeData(wb, sheet = 'Summary', startRow = 2, startCol = 2, x = df_summary)
    
    mergeCells(wb, sheet = 'Summary', rows = 1:2, cols = 1:2)
    mergeCells(wb, sheet = 'Summary', rows = 3:(2+N_facet+1), cols = 1)
    mergeCells(wb, sheet = 'Summary', rows = 1, cols = 3:(2+N_group+1))
    writeData(wb, sheet = 'Summary', xy = c(1,1), x = 'Summary')
    writeData(wb, sheet = 'Summary', xy = c(3,1), x = names(Vars)[Vars == var_group])
    writeData(wb, sheet = 'Summary', xy = c(1,3), x = names(Vars)[Vars == var_facet])
    addStyle(wb, sheet = 'Summary', rows = c(1,3,1), cols = c(1,1,3), stack = T, style = createStyle(halign = 'center', valign = 'center'))
    addStyle(wb, sheet = 'Summary', rows = 2, cols = 3:(2+N_group+1), gridExpand = T, stack = T, style = createStyle(halign = 'center', valign = 'center'))
  } else if (N_facet > 1 & N_group == 1) {
    df_summary <- df_ET %>%
      na.omit() %>%
      tail(N_facet) %>%
      select_at(vars(var_facet, 'CumN_Screen_Act'))
    colnames(df_summary) <- c('Summary', names(Vars)[Vars == var_facet])
    
    df_summary$Summary <- as.character(df_summary$Summary)
    df_summary <- rbind(df_summary, c(list('Total'), unname(as.list(colSums(df_summary[,-1])))))
    
    addStyle(wb, sheet = 'Summary', rows = 1:(1+N_facet+1), cols = 1:2, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    addStyle(wb, sheet = 'Summary', rows = 1, cols = 1:2, gridExpand = T, stack = T, style = createStyle(halign = 'center', valign = 'center'))
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 1, x = df_summary)
  } else if (N_facet == 1 & N_group > 1) {
    df_summary <- df_ET %>%
      na.omit() %>%
      tail(N_group) %>%
      select_at(vars(var_group, 'CumN_Screen_Act'))
    colnames(df_summary) <- c('Summary', names(Vars)[Vars == var_group])
    
    df_summary$Summary <- as.character(df_summary$Summary)
    df_summary <- rbind(df_summary, c(list('Total'), unname(as.list(colSums(df_summary[,-1])))))
    
    addStyle(wb, sheet = 'Summary', rows = 1:(1+N_group+1), cols = 1:2, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    addStyle(wb, sheet = 'Summary', rows = 1, cols = 1:2, gridExpand = T, stack = T, style = createStyle(halign = 'center', valign = 'center'))
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 1, x = df_summary)
  } else {
    x_summary <- df_ET %>%
      na.omit() %>%
      tail(1) %>%
      pull(CumN_Screen_Act)
    
    addStyle(wb, sheet = 'Summary', rows = 1, cols = 1:2, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 1, x = 'Total')
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 2, x = x_summary)
  }
  
  # Create report worksheets
  addWorksheet(wb, sheetName = 'Sheet1', gridLines = F)
  
  mergeCells(wb, sheet = 'Sheet1', rows = 1, cols = 1:2)
  mergeCells(wb, sheet = 'Sheet1', rows = 1, cols = 3:4)
  mergeCells(wb, sheet = 'Sheet1', rows = 2, cols = 1:2)
  mergeCells(wb, sheet = 'Sheet1', rows = 2, cols = 3:4)
  addStyle(wb, sheet = 'Sheet1', rows = 1:2, cols = 1:4, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
  addStyle(wb, sheet = 'Sheet1', rows = 2, cols = 1:4, stack = T, style = createStyle(fgFill = '#D9E1F2'))
  writeData(wb, sheet = 'Sheet1', startRow = 1, startCol = 1, x = 'Research Period')
  writeData(wb, sheet = 'Sheet1', startRow = 1, startCol = 3, x = paste(format(ResearchPeriod, '%Y-%m-%d'), collapse = ' to '))
  writeData(wb, sheet = 'Sheet1', startRow = 2, startCol = 1, x = 'Observation Period')
  writeData(wb, sheet = 'Sheet1', startRow = 2, startCol = 3, x = paste(format(ObservedPeriod, '%Y-%m-%d'), collapse = ' to '))
  
  mergeCells(wb, sheet = 'Sheet1', rows = 4, cols = 1:(3 + 2*N_group))
  addStyle(wb, sheet = 'Sheet1', rows = 4, cols = 1:(3 + 2*N_group), stack = T, style = createStyle(fontColour = 'white', fgFill = 'black'))
  writeData(wb, sheet = 'Sheet1', startRow = 4, startCol = 1, x = ifelse(is.null(var_facet), 'Participants Screened Overtime', sprintf('Participants Screened Overtime by %s', names(Vars)[Vars == var_facet])))
  
  addStyle(wb, sheet = 'Sheet1', rows = 5:(4 + 2 + nrow(df_ET)/(N_facet*N_group)), cols = 1:(3 + 2*N_group), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center', valign = 'center'))
  
  mergeCells(wb, sheet = 'Sheet1', rows = 5:6, cols = 1)
  mergeCells(wb, sheet = 'Sheet1', rows = 5, cols = 2:3)
  writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 1, x = unit)
  writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 2, x = 'Period')
  writeData(wb, sheet = 'Sheet1', startRow = 6, startCol = 2, x = 'Start Date')
  writeData(wb, sheet = 'Sheet1', startRow = 6, startCol = 3, x = 'End Date')
  writeData(wb, sheet = 'Sheet1', startRow = 7, startCol = 1, x = df_ET %>% pull(unit) %>% unique)
  writeData(wb, sheet = 'Sheet1', startRow = 7, startCol = 2, x = df_ET %>% pull(Date_Start) %>% unique %>% format('%Y-%m-%d'))
  writeData(wb, sheet = 'Sheet1', startRow = 7, startCol = 3, x = df_ET %>% pull(Date_End) %>% unique %>% format('%Y-%m-%d'))
  
  if (N_group == 1) {
    mergeCells(wb, sheet = 'Sheet1', rows = 5, cols = 4:5)
    writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 4, x = 'Overall')
    setColWidths(wb, sheet = 'Sheet1', cols = 3 + c(2), widths = 'auto')
  } else {
    for (j in 1:N_group) {
      group <- levels(pull(df_ET, var_group))[j]
      
      mergeCells(wb, sheet = 'Sheet1', rows = 5, cols = (3 + 2*(j-1) + 1):(3 + 2*(j-1) + 2))
      writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 3 + 2*(j-1) + 1, x = sprintf('%s: %s', names(Vars)[Vars == var_group], group))
      setColWidths(wb, sheet = 'Sheet1', cols = 3 + 2*(j-1) + c(2), widths = 'auto')
    }
  }
  
  if (N_facet == 1) {
    renameWorksheet(wb, sheet = 'Sheet1', newName = 'Stratified')
    
    if (N_group == 1) {
      df_ET %>%
        select(N_Screen_Act, CumN_Screen_Act) %>%
        rename(Actual = N_Screen_Act, `Cumulative Actual` = CumN_Screen_Act) %>%
        writeData(wb, sheet = 'Stratified', startRow = 6, startCol = 4, x = .)
    } else {
      for (j in 1:N_group) {
        group <- levels(pull(df_ET, var_group))[j]
        df_ET %>%
          filter(get(var_group) == group) %>%
          select(N_Screen_Act, CumN_Screen_Act) %>%
          rename(Actual = N_Screen_Act, `Cumulative Actual` = CumN_Screen_Act) %>%
          writeData(wb, sheet = 'Stratified', startRow = 6, startCol = 3 + 2*(j-1) + 1, x = .)
      }
    }
  } else {
    for (i in 1:N_facet) {
      facet <- levels(pull(df_ET, var_facet))[i]
      cloneWorksheet(wb, sheetName = facet, clonedSheet = 'Sheet1')
      
      if (N_group == 1) {
        df_ET %>%
          filter(get(var_facet) == facet) %>%
          select(N_Screen_Act, CumN_Screen_Act) %>%
          rename(Actual = N_Screen_Act, `Cumulative Actual` = CumN_Screen_Act) %>%
          writeData(wb, sheet = facet, startRow = 6, startCol = 4, x = .)
      } else {
        for (j in 1:N_group) {
          group <- levels(pull(df_ET, var_group))[j]
          df_ET %>%
            filter(get(var_facet) == facet & get(var_group) == group) %>%
            select(N_Screen_Act, CumN_Screen_Act) %>%
            rename(Actual = N_Screen_Act, `Cumulative Actual` = CumN_Screen_Act) %>%
            writeData(wb, sheet = facet, startRow = 6, startCol = 3 + 2*(j-1) + 1, x = .)
        }
      }
    }
    removeWorksheet(wb, 'Sheet1')
  }
  
  saveWorkbook(wb, file, overwrite = T)
}

##################################################
# Excel Function for Tracking Module: Randomized Data 
ET_1_2_2 <- function(file, data, Vars,
                     ResearchPeriod, ObservedPeriod, unit,
                     var_facet, var_group)
{
  N_facet <- ifelse(is.null(var_facet), 1, nlevels(pull(data, var_facet)))
  N_group <- ifelse(is.null(var_group), 1, nlevels(pull(data, var_group)))
  
  df_ET <- data %>%
    mutate(N_Enrolled_Tar = replace(N_Enrolled_Tar, Date_End < ObservedPeriod[1] | Date_Start > ObservedPeriod[2], NA),
           N_Enrolled_Act = replace(N_Enrolled_Act, Date_End < ObservedPeriod[1] | Date_Start > ObservedPeriod[2], NA),
           N_Enrolled_Achieve = replace(N_Enrolled_Achieve, Date_End < ObservedPeriod[1] | Date_Start > ObservedPeriod[2], NA),
           CumN_Enrolled_Tar = replace(CumN_Enrolled_Tar, Date_End < ObservedPeriod[1] | Date_Start > ObservedPeriod[2], NA),
           CumN_Enrolled_Act = replace(CumN_Enrolled_Act, Date_End < ObservedPeriod[1] | Date_Start > ObservedPeriod[2], NA),
           CumN_Enrolled_Achieve = replace(CumN_Enrolled_Achieve, Date_End < ObservedPeriod[1] | Date_Start > ObservedPeriod[2], NA))
  
  wb <- createWorkbook()
  
  # Create summary worksheet
  addWorksheet(wb, sheetName = 'Summary', gridLines = F)
  
  if (N_facet > 1 & N_group > 1) {
    df_summary <- df_ET %>%
      na.omit() %>%
      tail(N_facet*N_group) %>%
      select_at(vars(var_facet, var_group, 'CumN_Enrolled_Act')) %>%
      spread(key = var_group, value = 'CumN_Enrolled_Act')
    colnames(df_summary)[1] <- 'Summary'
    
    df_summary$Summary <- as.character(df_summary$Summary)
    df_summary$Total <- rowSums(df_summary[,-1])
    df_summary <- rbind(df_summary, c(list('Total'), unname(as.list(colSums(df_summary[,-1])))))
    
    addStyle(wb, sheet = 'Summary', rows = 1:(2+N_facet+1), cols = 1:(2+N_group+1), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    writeData(wb, sheet = 'Summary', startRow = 2, startCol = 2, x = df_summary)
    
    mergeCells(wb, sheet = 'Summary', rows = 1:2, cols = 1:2)
    mergeCells(wb, sheet = 'Summary', rows = 3:(2+N_facet+1), cols = 1)
    mergeCells(wb, sheet = 'Summary', rows = 1, cols = 3:(2+N_group+1))
    writeData(wb, sheet = 'Summary', xy = c(1,1), x = 'Summary')
    writeData(wb, sheet = 'Summary', xy = c(3,1), x = names(Vars)[Vars == var_group])
    writeData(wb, sheet = 'Summary', xy = c(1,3), x = names(Vars)[Vars == var_facet])
    addStyle(wb, sheet = 'Summary', rows = c(1,3,1), cols = c(1,1,3), stack = T, style = createStyle(halign = 'center', valign = 'center'))
    addStyle(wb, sheet = 'Summary', rows = 2, cols = 3:(2+N_group+1), gridExpand = T, stack = T, style = createStyle(halign = 'center', valign = 'center'))
  } else if (N_facet > 1 & N_group == 1) {
    df_summary <- df_ET %>%
      na.omit() %>%
      tail(N_facet) %>%
      select_at(vars(var_facet, 'CumN_Enrolled_Act'))
    colnames(df_summary) <- c('Summary', names(Vars)[Vars == var_facet])
    
    df_summary$Summary <- as.character(df_summary$Summary)
    df_summary <- rbind(df_summary, c(list('Total'), unname(as.list(colSums(df_summary[,-1])))))
    
    addStyle(wb, sheet = 'Summary', rows = 1:(1+N_facet+1), cols = 1:2, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    addStyle(wb, sheet = 'Summary', rows = 1, cols = 1:2, gridExpand = T, stack = T, style = createStyle(halign = 'center', valign = 'center'))
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 1, x = df_summary)
  } else if (N_facet == 1 & N_group > 1) {
    df_summary <- df_ET %>%
      na.omit() %>%
      tail(N_group) %>%
      select_at(vars(var_group, 'CumN_Enrolled_Act'))
    colnames(df_summary) <- c('Summary', names(Vars)[Vars == var_group])
    
    df_summary$Summary <- as.character(df_summary$Summary)
    df_summary <- rbind(df_summary, c(list('Total'), unname(as.list(colSums(df_summary[,-1])))))
    
    addStyle(wb, sheet = 'Summary', rows = 1:(1+N_group+1), cols = 1:2, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    addStyle(wb, sheet = 'Summary', rows = 1, cols = 1:2, gridExpand = T, stack = T, style = createStyle(halign = 'center', valign = 'center'))
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 1, x = df_summary)
  } else {
    x_summary <- df_ET %>%
      na.omit() %>%
      tail(1) %>%
      pull(CumN_Enrolled_Act)
    
    addStyle(wb, sheet = 'Summary', rows = 1, cols = 1:2, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 1, x = 'Total')
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 2, x = x_summary)
  }
  
  # Create report worksheets
  addWorksheet(wb, sheetName = 'Sheet1', gridLines = F)
  
  mergeCells(wb, sheet = 'Sheet1', rows = 1, cols = 1:2)
  mergeCells(wb, sheet = 'Sheet1', rows = 1, cols = 3:4)
  mergeCells(wb, sheet = 'Sheet1', rows = 2, cols = 1:2)
  mergeCells(wb, sheet = 'Sheet1', rows = 2, cols = 3:4)
  addStyle(wb, sheet = 'Sheet1', rows = 1:2, cols = 1:4, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
  addStyle(wb, sheet = 'Sheet1', rows = 2, cols = 1:4, stack = T, style = createStyle(fgFill = '#D9E1F2'))
  writeData(wb, sheet = 'Sheet1', startRow = 1, startCol = 1, x = 'Research Period')
  writeData(wb, sheet = 'Sheet1', startRow = 1, startCol = 3, x = paste(format(ResearchPeriod, '%Y-%m-%d'), collapse = ' to '))
  writeData(wb, sheet = 'Sheet1', startRow = 2, startCol = 1, x = 'Observation Period')
  writeData(wb, sheet = 'Sheet1', startRow = 2, startCol = 3, x = paste(format(ObservedPeriod, '%Y-%m-%d'), collapse = ' to '))
  
  mergeCells(wb, sheet = 'Sheet1', rows = 4, cols = 1:(3 + 6*N_group))
  addStyle(wb, sheet = 'Sheet1', rows = 4, cols = 1:(3 + 6*N_group), stack = T, style = createStyle(fontColour = 'white', fgFill = 'black'))
  writeData(wb, sheet = 'Sheet1', startRow = 4, startCol = 1, x = ifelse(is.null(var_facet), 'Participants Randomized Overtime', sprintf('Participants Randomized Overtime by %s', names(Vars)[Vars == var_facet])))
  
  addStyle(wb, sheet = 'Sheet1', rows = 5:(4 + 2 + nrow(df_ET)/(N_facet*N_group)), cols = 1:(3 + 6*N_group), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center', valign = 'center'))
  
  mergeCells(wb, sheet = 'Sheet1', rows = 5:6, cols = 1)
  mergeCells(wb, sheet = 'Sheet1', rows = 5, cols = 2:3)
  writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 1, x = unit)
  writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 2, x = 'Period')
  writeData(wb, sheet = 'Sheet1', startRow = 6, startCol = 2, x = 'Start Date')
  writeData(wb, sheet = 'Sheet1', startRow = 6, startCol = 3, x = 'End Date')
  writeData(wb, sheet = 'Sheet1', startRow = 7, startCol = 1, x = df_ET %>% pull(unit) %>% unique)
  writeData(wb, sheet = 'Sheet1', startRow = 7, startCol = 2, x = df_ET %>% pull(Date_Start) %>% unique %>% format('%Y-%m-%d'))
  writeData(wb, sheet = 'Sheet1', startRow = 7, startCol = 3, x = df_ET %>% pull(Date_End) %>% unique %>% format('%Y-%m-%d'))
  
  if (N_group == 1) {
    mergeCells(wb, sheet = 'Sheet1', rows = 5, cols = 4:9)
    writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 4, x = 'Overall')
    setColWidths(wb, sheet = 'Sheet1', cols = 3 + c(3,4,5,6), widths = 'auto')
  } else {
    for (j in 1:N_group) {
      group <- levels(pull(df_ET, var_group))[j]
      
      mergeCells(wb, sheet = 'Sheet1', rows = 5, cols = (3 + 6*(j-1) + 1):(3 + 6*(j-1) + 6))
      writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 3 + 6*(j-1) + 1, x = sprintf('%s: %s', names(Vars)[Vars == var_group], group))
      setColWidths(wb, sheet = 'Sheet1', cols = 3 + 6*(j-1) + c(3,4,5,6), widths = 'auto')
    }
  }
  
  if (N_facet == 1) {
    renameWorksheet(wb, sheet = 'Sheet1', newName = 'Stratified')
    
    if (N_group == 1) {
      df_ET %>%
        select(N_Enrolled_Tar, N_Enrolled_Act, N_Enrolled_Achieve, CumN_Enrolled_Tar, CumN_Enrolled_Act, CumN_Enrolled_Achieve) %>%
        rename(Target = N_Enrolled_Tar, Actual = N_Enrolled_Act, `Accrual Rate` = N_Enrolled_Achieve, `Cumulative Target` = CumN_Enrolled_Tar, `Cumulative Actual` = CumN_Enrolled_Act, `Cumulative Accrual Rate` = CumN_Enrolled_Achieve) %>%
        writeData(wb, sheet = 'Stratified', startRow = 6, startCol = 4, x = .)
      addStyle(wb, sheet = 'Stratified', rows = 7:(6 + nrow(data)), cols = 3 + c(1,4), gridExpand = T, stack = T, style = createStyle(numFmt = '0.0'))
      addStyle(wb, sheet = 'Stratified', rows = 6:(6 + nrow(data)), cols = 3 + c(3,6), gridExpand = T, stack = T, style = createStyle(fgFill = '#D9E1F2'))
      addStyle(wb, sheet = 'Stratified', rows = 7:(6 + nrow(data)), cols = 3 + c(3,6), gridExpand = T, stack = T, style = createStyle(numFmt = '0%'))
    } else {
      for (j in 1:N_group) {
        group <- levels(pull(df_ET, var_group))[j]
        df_ET %>%
          filter(get(var_group) == group) %>%
          select(N_Enrolled_Tar, N_Enrolled_Act, N_Enrolled_Achieve, CumN_Enrolled_Tar, CumN_Enrolled_Act, CumN_Enrolled_Achieve) %>%
          rename(Target = N_Enrolled_Tar, Actual = N_Enrolled_Act, `Accrual Rate` = N_Enrolled_Achieve, `Cumulative Target` = CumN_Enrolled_Tar, `Cumulative Actual` = CumN_Enrolled_Act, `Cumulative Accrual Rate` = CumN_Enrolled_Achieve) %>%
          writeData(wb, sheet = 'Stratified', startRow = 6, startCol = 3 + 6*(j-1) + 1, x = .)
        addStyle(wb, sheet = 'Stratified', rows = 7:(6 + nrow(data)/N_group), cols = 3 + 6*(j-1) + c(1,4), gridExpand = T, stack = T, style = createStyle(numFmt = '0.0'))
        addStyle(wb, sheet = 'Stratified', rows = 6:(6 + nrow(data)/N_group), cols = 3 + 6*(j-1) + c(3,6), gridExpand = T, stack = T, style = createStyle(fgFill = '#D9E1F2'))
        addStyle(wb, sheet = 'Stratified', rows = 7:(6 + nrow(data)/N_group), cols = 3 + 6*(j-1) + c(3,6), gridExpand = T, stack = T, style = createStyle(numFmt = '0%'))
      }
    }
  } else {
    for (i in 1:N_facet) {
      facet <- levels(pull(df_ET, var_facet))[i]
      cloneWorksheet(wb, sheetName = facet, clonedSheet = 'Sheet1')
      
      if (N_group == 1) {
        df_ET %>%
          filter(get(var_facet) == facet) %>%
          select(N_Enrolled_Tar, N_Enrolled_Act, N_Enrolled_Achieve, CumN_Enrolled_Tar, CumN_Enrolled_Act, CumN_Enrolled_Achieve) %>%
          rename(Target = N_Enrolled_Tar, Actual = N_Enrolled_Act, `Accrual Rate` = N_Enrolled_Achieve, `Cumulative Target` = CumN_Enrolled_Tar, `Cumulative Actual` = CumN_Enrolled_Act, `Cumulative Accrual Rate` = CumN_Enrolled_Achieve) %>%
          writeData(wb, sheet = facet, startRow = 6, startCol = 4, x = .)
        addStyle(wb, sheet = facet, rows = 7:(6 + nrow(data)/N_facet), cols = 3 + c(1,4), gridExpand = T, stack = T, style = createStyle(numFmt = '0.0'))
        addStyle(wb, sheet = facet, rows = 6:(6 + nrow(data)/N_facet), cols = 3 + c(3,6), gridExpand = T, stack = T, style = createStyle(fgFill = '#D9E1F2'))
        addStyle(wb, sheet = facet, rows = 7:(6 + nrow(data)/N_facet), cols = 3 + c(3,6), gridExpand = T, stack = T, style = createStyle(numFmt = '0%'))
      } else {
        for (j in 1:N_group) {
          group <- levels(pull(df_ET, var_group))[j]
          df_ET %>%
            filter(get(var_facet) == facet & get(var_group) == group) %>%
            select(N_Enrolled_Tar, N_Enrolled_Act, N_Enrolled_Achieve, CumN_Enrolled_Tar, CumN_Enrolled_Act, CumN_Enrolled_Achieve) %>%
            rename(Target = N_Enrolled_Tar, Actual = N_Enrolled_Act, `Accrual Rate` = N_Enrolled_Achieve, `Cumulative Target` = CumN_Enrolled_Tar, `Cumulative Actual` = CumN_Enrolled_Act, `Cumulative Accrual Rate` = CumN_Enrolled_Achieve) %>%
            writeData(wb, sheet = facet, startRow = 6, startCol = 3 + 6*(j-1) + 1, x = .)
          addStyle(wb, sheet = facet, rows = 7:(6 + nrow(data)/(N_facet*N_group)), cols = 3 + 6*(j-1) + c(1,4), gridExpand = T, stack = T, style = createStyle(numFmt = '0.0'))
          addStyle(wb, sheet = facet, rows = 6:(6 + nrow(data)/(N_facet*N_group)), cols = 3 + 6*(j-1) + c(3,6), gridExpand = T, stack = T, style = createStyle(fgFill = '#D9E1F2'))
          addStyle(wb, sheet = facet, rows = 7:(6 + nrow(data)/(N_facet*N_group)), cols = 3 + 6*(j-1) + c(3,6), gridExpand = T, stack = T, style = createStyle(numFmt = '0%'))
        }
      }
    }
    removeWorksheet(wb, 'Sheet1')
  }
  
  saveWorkbook(wb, file, overwrite = T)
}


##################################################
# Excel Function for Tracking Module: Eligible Data 
ET_1_2_3 <- function(file, data, Vars,
                     ResearchPeriod, ObservedPeriod, unit,
                     var_facet, var_group)
{
  N_facet <- ifelse(is.null(var_facet), 1, nlevels(pull(data, var_facet)))
  N_group <- ifelse(is.null(var_group), 1, nlevels(pull(data, var_group)))
  
  df_ET <- data %>%
    mutate(N_Eligible_Act = replace(N_Eligible_Act, Date_End < ObservedPeriod[1] | Date_Start > ObservedPeriod[2], NA),
           CumN_Eligible_Act = replace(CumN_Eligible_Act, Date_End < ObservedPeriod[1] | Date_Start > ObservedPeriod[2], NA))
  
  wb <- createWorkbook()
  
  # Create summary worksheet
  addWorksheet(wb, sheetName = 'Summary', gridLines = F)
  
  if (N_facet > 1 & N_group > 1) {
    df_summary <- df_ET %>%
      na.omit() %>%
      tail(N_facet*N_group) %>%
      select_at(vars(var_facet, var_group, 'CumN_Eligible_Act')) %>%
      spread(key = var_group, value = 'CumN_Eligible_Act')
    colnames(df_summary)[1] <- 'Summary'
    
    df_summary$Summary <- as.character(df_summary$Summary)
    df_summary$Total <- rowSums(df_summary[,-1])
    df_summary <- rbind(df_summary, c(list('Total'), unname(as.list(colSums(df_summary[,-1])))))
    
    addStyle(wb, sheet = 'Summary', rows = 1:(2+N_facet+1), cols = 1:(2+N_group+1), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    writeData(wb, sheet = 'Summary', startRow = 2, startCol = 2, x = df_summary)
    
    mergeCells(wb, sheet = 'Summary', rows = 1:2, cols = 1:2)
    mergeCells(wb, sheet = 'Summary', rows = 3:(2+N_facet+1), cols = 1)
    mergeCells(wb, sheet = 'Summary', rows = 1, cols = 3:(2+N_group+1))
    writeData(wb, sheet = 'Summary', xy = c(1,1), x = 'Summary')
    writeData(wb, sheet = 'Summary', xy = c(3,1), x = names(Vars)[Vars == var_group])
    writeData(wb, sheet = 'Summary', xy = c(1,3), x = names(Vars)[Vars == var_facet])
    addStyle(wb, sheet = 'Summary', rows = c(1,3,1), cols = c(1,1,3), stack = T, style = createStyle(halign = 'center', valign = 'center'))
    addStyle(wb, sheet = 'Summary', rows = 2, cols = 3:(2+N_group+1), gridExpand = T, stack = T, style = createStyle(halign = 'center', valign = 'center'))
  } else if (N_facet > 1 & N_group == 1) {
    df_summary <- df_ET %>%
      na.omit() %>%
      tail(N_facet) %>%
      select_at(vars(var_facet, 'CumN_Eligible_Act'))
    colnames(df_summary) <- c('Summary', names(Vars)[Vars == var_facet])
    
    df_summary$Summary <- as.character(df_summary$Summary)
    df_summary <- rbind(df_summary, c(list('Total'), unname(as.list(colSums(df_summary[,-1])))))
    
    addStyle(wb, sheet = 'Summary', rows = 1:(1+N_facet+1), cols = 1:2, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    addStyle(wb, sheet = 'Summary', rows = 1, cols = 1:2, gridExpand = T, stack = T, style = createStyle(halign = 'center', valign = 'center'))
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 1, x = df_summary)
  } else if (N_facet == 1 & N_group > 1) {
    df_summary <- df_ET %>%
      na.omit() %>%
      tail(N_group) %>%
      select_at(vars(var_group, 'CumN_Eligible_Act'))
    colnames(df_summary) <- c('Summary', names(Vars)[Vars == var_group])
    
    df_summary$Summary <- as.character(df_summary$Summary)
    df_summary <- rbind(df_summary, c(list('Total'), unname(as.list(colSums(df_summary[,-1])))))
    
    addStyle(wb, sheet = 'Summary', rows = 1:(1+N_group+1), cols = 1:2, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    addStyle(wb, sheet = 'Summary', rows = 1, cols = 1:2, gridExpand = T, stack = T, style = createStyle(halign = 'center', valign = 'center'))
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 1, x = df_summary)
  } else {
    x_summary <- df_ET %>%
      na.omit() %>%
      tail(1) %>%
      pull(CumN_Eligible_Act)
    
    addStyle(wb, sheet = 'Summary', rows = 1, cols = 1:2, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight'))
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 1, x = 'Total')
    writeData(wb, sheet = 'Summary', startRow = 1, startCol = 2, x = x_summary)
  }
  
  # Create report worksheets
  addWorksheet(wb, sheetName = 'Sheet1', gridLines = F)
  
  mergeCells(wb, sheet = 'Sheet1', rows = 1, cols = 1:2)
  mergeCells(wb, sheet = 'Sheet1', rows = 1, cols = 3:4)
  mergeCells(wb, sheet = 'Sheet1', rows = 2, cols = 1:2)
  mergeCells(wb, sheet = 'Sheet1', rows = 2, cols = 3:4)
  addStyle(wb, sheet = 'Sheet1', rows = 1:2, cols = 1:4, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
  addStyle(wb, sheet = 'Sheet1', rows = 2, cols = 1:4, stack = T, style = createStyle(fgFill = '#D9E1F2'))
  writeData(wb, sheet = 'Sheet1', startRow = 1, startCol = 1, x = 'Research Period')
  writeData(wb, sheet = 'Sheet1', startRow = 1, startCol = 3, x = paste(format(ResearchPeriod, '%Y-%m-%d'), collapse = ' to '))
  writeData(wb, sheet = 'Sheet1', startRow = 2, startCol = 1, x = 'Observation Period')
  writeData(wb, sheet = 'Sheet1', startRow = 2, startCol = 3, x = paste(format(ObservedPeriod, '%Y-%m-%d'), collapse = ' to '))
  
  mergeCells(wb, sheet = 'Sheet1', rows = 4, cols = 1:(3 + 2*N_group))
  addStyle(wb, sheet = 'Sheet1', rows = 4, cols = 1:(3 + 2*N_group), stack = T, style = createStyle(fontColour = 'white', fgFill = 'black'))
  writeData(wb, sheet = 'Sheet1', startRow = 4, startCol = 1, x = ifelse(is.null(var_facet), 'Participants Eligible Overtime', sprintf('Participants Eligible Overtime by %s', names(Vars)[Vars == var_facet])))
  
  addStyle(wb, sheet = 'Sheet1', rows = 5:(4 + 2 + nrow(df_ET)/(N_facet*N_group)), cols = 1:(3 + 2*N_group), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center', valign = 'center'))
  
  mergeCells(wb, sheet = 'Sheet1', rows = 5:6, cols = 1)
  mergeCells(wb, sheet = 'Sheet1', rows = 5, cols = 2:3)
  writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 1, x = unit)
  writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 2, x = 'Period')
  writeData(wb, sheet = 'Sheet1', startRow = 6, startCol = 2, x = 'Start Date')
  writeData(wb, sheet = 'Sheet1', startRow = 6, startCol = 3, x = 'End Date')
  writeData(wb, sheet = 'Sheet1', startRow = 7, startCol = 1, x = df_ET %>% pull(unit) %>% unique)
  writeData(wb, sheet = 'Sheet1', startRow = 7, startCol = 2, x = df_ET %>% pull(Date_Start) %>% unique %>% format('%Y-%m-%d'))
  writeData(wb, sheet = 'Sheet1', startRow = 7, startCol = 3, x = df_ET %>% pull(Date_End) %>% unique %>% format('%Y-%m-%d'))
  
  if (N_group == 1) {
    mergeCells(wb, sheet = 'Sheet1', rows = 5, cols = 4:5)
    writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 4, x = 'Overall')
    setColWidths(wb, sheet = 'Sheet1', cols = 3 + c(2), widths = 'auto')
  } else {
    for (j in 1:N_group) {
      group <- levels(pull(df_ET, var_group))[j]
      
      mergeCells(wb, sheet = 'Sheet1', rows = 5, cols = (3 + 2*(j-1) + 1):(3 + 2*(j-1) + 2))
      writeData(wb, sheet = 'Sheet1', startRow = 5, startCol = 3 + 2*(j-1) + 1, x = sprintf('%s: %s', names(Vars)[Vars == var_group], group))
      setColWidths(wb, sheet = 'Sheet1', cols = 3 + 2*(j-1) + c(2), widths = 'auto')
    }
  }
  
  if (N_facet == 1) {
    renameWorksheet(wb, sheet = 'Sheet1', newName = 'Stratified')
    
    if (N_group == 1) {
      df_ET %>%
        select(N_Eligible_Act, CumN_Eligible_Act) %>%
        rename(Actual = N_Eligible_Act, `Cumulative Actual` = CumN_Eligible_Act) %>%
        writeData(wb, sheet = 'Stratified', startRow = 6, startCol = 4, x = .)
    } else {
      for (j in 1:N_group) {
        group <- levels(pull(df_ET, var_group))[j]
        df_ET %>%
          filter(get(var_group) == group) %>%
          select(N_Eligible_Act, CumN_Eligible_Act) %>%
          rename(Actual = N_Eligible_Act, `Cumulative Actual` = CumN_Eligible_Act) %>%
          writeData(wb, sheet = 'Stratified', startRow = 6, startCol = 3 + 2*(j-1) + 1, x = .)
      }
    }
  } else {
    for (i in 1:N_facet) {
      facet <- levels(pull(df_ET, var_facet))[i]
      cloneWorksheet(wb, sheetName = facet, clonedSheet = 'Sheet1')
      
      if (N_group == 1) {
        df_ET %>%
          filter(get(var_facet) == facet) %>%
          select(N_Eligible_Act, CumN_Eligible_Act) %>%
          rename(Actual = N_Eligible_Act, `Cumulative Actual` = CumN_Eligible_Act) %>%
          writeData(wb, sheet = facet, startRow = 6, startCol = 4, x = .)
      } else {
        for (j in 1:N_group) {
          group <- levels(pull(df_ET, var_group))[j]
          df_ET %>%
            filter(get(var_facet) == facet & get(var_group) == group) %>%
            select(N_Eligible_Act, CumN_Eligible_Act) %>%
            rename(Actual = N_Eligible_Act, `Cumulative Actual` = CumN_Eligible_Act) %>%
            writeData(wb, sheet = facet, startRow = 6, startCol = 3 + 2*(j-1) + 1, x = .)
        }
      }
    }
    removeWorksheet(wb, 'Sheet1')
  }
  
  saveWorkbook(wb, file, overwrite = T)
}

##################################################
# Excel Function for Descriptive Statistics Module: Univariate Analysis 
ET_1_3_1 <- function(file, data, Vars,
                     ResearchPeriod, ObservedPeriod,
                     vars_cate, vars_num)
{
  wb <- createWorkbook()
  
  addWorksheet(wb, sheetName = 'Sheet1', gridLines = F)
  setColWidths(wb, sheet = 'Sheet1', cols = 1:1000, widths = 20)
  
  mergeCells(wb, sheet = 'Sheet1', rows = 1, cols = 1:2)
  mergeCells(wb, sheet = 'Sheet1', rows = 1, cols = 3:4)
  mergeCells(wb, sheet = 'Sheet1', rows = 2, cols = 1:2)
  mergeCells(wb, sheet = 'Sheet1', rows = 2, cols = 3:4)
  addStyle(wb, sheet = 'Sheet1', rows = 1:2, cols = 1:4, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
  addStyle(wb, sheet = 'Sheet1', rows = 2, cols = 1:4, stack = T, style = createStyle(fgFill = '#D9E1F2'))
  writeData(wb, sheet = 'Sheet1', startRow = 1, startCol = 1, x = 'Research Period')
  writeData(wb, sheet = 'Sheet1', startRow = 1, startCol = 3, x = paste(format(ResearchPeriod, '%Y-%m-%d'), collapse = ' to '))
  writeData(wb, sheet = 'Sheet1', startRow = 2, startCol = 1, x = 'Observation Period')
  writeData(wb, sheet = 'Sheet1', startRow = 2, startCol = 3, x = paste(format(ObservedPeriod, '%Y-%m-%d'), collapse = ' to '))
  
  # Sheet 1: Categorical Variable
  if (length(vars_cate) > 0) {
    cloneWorksheet(wb, sheetName = 'cate', clonedSheet = 'Sheet1')
    i <- 4
    
    for (var in vars_cate) {
      nMiss <- sum(is.na(pull(data, var)))
      df <- data %>%
        group_by_at(var) %>% summarise(Frequency = n()) %>%
        na.omit() %>%
        mutate(Percent = Frequency/sum(Frequency)) %>%
        mutate(`Cumulative Frequency` = cumsum(Frequency),
               `Cumulative Percent` = cumsum(Percent)) %>%
        mutate(Percent = round(Percent*100, 2),
               `Cumulative Percent` = round(`Cumulative Percent`*100, 2))
      colnames(df)[1] <- names(Vars)[Vars == var]
      
      addStyle(wb, sheet = 'cate', rows = i:(i + nrow(df)), cols = 1:ncol(df), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
      addStyle(wb, sheet = 'cate', rows = (i+1):(i + nrow(df)), cols = c(3,5), gridExpand = T, stack = T, style = createStyle(numFmt = '0.00'))
      writeData(wb, sheet = 'cate', startRow = i, startCol = 1, x = df)
      
      if (nMiss > 0) {
        mergeCells(wb, sheet = 'cate', rows = (i + nrow(df) + 1), cols = 1:ncol(df))
        addStyle(wb, sheet = 'cate', rows = (i + nrow(df) + 1), cols = 1:ncol(df), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
        writeData(wb, sheet = 'cate', startRow = (i + nrow(df) + 1), startCol = 1, x = sprintf('Frequency Missing = %d', nMiss))
        i <- i + 1
      }
      
      i <- i + 1 + nrow(df) + 1
    }
  }
  
  # Sheet 2: Numerical Variable
  if (length(vars_num) > 0) {
    cloneWorksheet(wb, sheetName = 'num', clonedSheet = 'Sheet1')
    i <- 4
    
    df <- data.frame(Variable = sapply(vars_num, function(var) names(Vars)[Vars == var]),
                     N = apply(select(data, vars_num), MARGIN = 2, function(x) sum(!is.na(x))),
                     `N Miss` = apply(select(data, vars_num), MARGIN = 2, function(x) sum(is.na(x))),
                     Mean = apply(select(data, vars_num), MARGIN = 2, mean, na.rm = T),
                     `Std Dev` = apply(select(data, vars_num), MARGIN = 2, sd, na.rm = T),
                     IQR = apply(select(data, vars_num), MARGIN = 2, IQR, na.rm = T),
                     Minimum = apply(select(data, vars_num), MARGIN = 2, min, na.rm = T),
                     Q1 = apply(select(data, vars_num), MARGIN = 2, quantile, probs = 0.25, na.rm = T),
                     Median = apply(select(data, vars_num), MARGIN = 2, median, na.rm = T),
                     Q3 = apply(select(data, vars_num), MARGIN = 2, quantile, probs = 0.75, na.rm = T),
                     Maximum = apply(select(data, vars_num), MARGIN = 2, max, na.rm = T)
                     )
    
    addStyle(wb, sheet = 'num', rows = i:(i + nrow(df)), cols = 1:ncol(df), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
    addStyle(wb, sheet = 'num', rows = (i+1):(i + nrow(df)), cols = 4:ncol(df), gridExpand = T, stack = T, style = createStyle(numFmt = '0.00'))
    
    writeData(wb, sheet = 'num', startRow = i, startCol = 1, x = df)
  }
  
  removeWorksheet(wb, 'Sheet1')
  saveWorkbook(wb, file, overwrite = T)
}

##################################################
# Excel Function for Descriptive Statistics Module: Bivariate Analysis 
ET_1_3_2 <- function(file, data,
                     ResearchPeriod, ObservedPeriod,
                     var_main, vars_minor,
                     vars_cate, vars_num)
{
  vars_minor_cate <- vars_minor[vars_minor %in% vars_cate]
  vars_minor_num <- vars_minor[vars_minor %in% vars_num]
  
  {
    wb <- createWorkbook()
    
    addWorksheet(wb, sheetName = 'Sheet1', gridLines = F)
    setColWidths(wb, sheet = 'Sheet1', cols = 1:1000, widths = 20)
    
    mergeCells(wb, sheet = 'Sheet1', rows = 1, cols = 1:2)
    mergeCells(wb, sheet = 'Sheet1', rows = 1, cols = 3:4)
    mergeCells(wb, sheet = 'Sheet1', rows = 2, cols = 1:2)
    mergeCells(wb, sheet = 'Sheet1', rows = 2, cols = 3:4)
    addStyle(wb, sheet = 'Sheet1', rows = 1:2, cols = 1:4, gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
    addStyle(wb, sheet = 'Sheet1', rows = 2, cols = 1:4, stack = T, style = createStyle(fgFill = '#D9E1F2'))
    writeData(wb, sheet = 'Sheet1', startRow = 1, startCol = 1, x = 'Research Period')
    writeData(wb, sheet = 'Sheet1', startRow = 1, startCol = 3, x = paste(format(ResearchPeriod, '%Y-%m-%d'), collapse = ' to '))
    writeData(wb, sheet = 'Sheet1', startRow = 2, startCol = 1, x = 'Observation Period')
    writeData(wb, sheet = 'Sheet1', startRow = 2, startCol = 3, x = paste(format(ObservedPeriod, '%Y-%m-%d'), collapse = ' to '))
    
    if (var_main %in% vars_cate) {
      
      # Sheet 1: Category (Main) vs Category (Minor)
      if (length(vars_minor_cate) > 0) {
        cloneWorksheet(wb, sheetName = 'cate vs cate', clonedSheet = 'Sheet1')
        i <- 4
        
        for (var_minor in vars_minor_cate) {
          df <- data %>% select(var_main, var_minor)
          
          nMiss <- sum(!complete.cases(df))
          
          df <- na.omit(df)
          tb <- table(df) %>% t()
          
          addStyle(wb, sheet = 'cate vs cate', rows = i:(i+3*(2+nrow(tb)+1)-1), cols = 1:(2+ncol(tb)+1), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center', valign = 'center'))
          
          # Frequency
          tb1 <- addmargins(tb)
          rownames(tb1)[nrow(tb1)] <- colnames(tb1)[ncol(tb1)] <- 'Total'
          
          mergeCells(wb, sheet = 'cate vs cate', rows = i:(i+1), cols = 1:2)
          mergeCells(wb, sheet = 'cate vs cate', rows = (i+2):(i+1+nrow(tb1)), cols = 1)
          mergeCells(wb, sheet = 'cate vs cate', rows = i, cols = 3:(2+ncol(tb1)))
          
          writeData(wb, sheet = 'cate vs cate', startRow = i, startCol = 1, x = 'Frequency')
          writeData(wb, sheet = 'cate vs cate', startRow = i, startCol = 3, x = names(vars_cate)[vars_cate == var_main])
          writeData(wb, sheet = 'cate vs cate', startRow = i+2, startCol = 1, x = names(vars_cate)[vars_cate == var_minor])
          writeData(wb, sheet = 'cate vs cate', startRow = i+1, startCol = 2, x = tb1)

          i <- i + 2 + nrow(tb1)
          
          # Column Percent
          tb2 <- prop.table(tb1[-nrow(tb1),], margin = 2)
          tb2 <- addmargins(tb2, 1)*100
          rownames(tb2)[nrow(tb2)] <- 'Total'

          mergeCells(wb, sheet = 'cate vs cate', rows = i:(i+1), cols = 1:2)
          mergeCells(wb, sheet = 'cate vs cate', rows = (i+2):(i+1+nrow(tb2)), cols = 1)
          mergeCells(wb, sheet = 'cate vs cate', rows = i, cols = 3:(2+ncol(tb2)))
          
          addStyle(wb, sheet = 'cate vs cate', rows = (i+2):(i+1+nrow(tb2)), cols = 3:(2+ncol(tb2)), gridExpand = T, stack = T, style = createStyle(numFmt = '0.00'))
          
          writeData(wb, sheet = 'cate vs cate', startRow = i, startCol = 1, x = 'Column Percent')
          writeData(wb, sheet = 'cate vs cate', startRow = i, startCol = 3, x = names(vars_cate)[vars_cate == var_main])
          writeData(wb, sheet = 'cate vs cate', startRow = i+2, startCol = 1, x = names(vars_cate)[vars_cate == var_minor])
          writeData(wb, sheet = 'cate vs cate', startRow = i+1, startCol = 2, x = tb2)
          
          i <- i + 2 + nrow(tb2)
          
          # Row Percent
          tb3 <- prop.table(tb1[,-ncol(tb1)], margin = 1)
          tb3 <- addmargins(tb3, 2)*100
          colnames(tb3)[ncol(tb3)] <- 'Total'
          
          mergeCells(wb, sheet = 'cate vs cate', rows = i:(i+1), cols = 1:2)
          mergeCells(wb, sheet = 'cate vs cate', rows = (i+2):(i+1+nrow(tb3)), cols = 1)
          mergeCells(wb, sheet = 'cate vs cate', rows = i, cols = 3:(2+ncol(tb3)))
          
          addStyle(wb, sheet = 'cate vs cate', rows = (i+2):(i+1+nrow(tb3)), cols = 3:(2+ncol(tb3)), gridExpand = T, stack = T, style = createStyle(numFmt = '0.00'))
          
          writeData(wb, sheet = 'cate vs cate', startRow = i, startCol = 1, x = 'Row Percent')
          writeData(wb, sheet = 'cate vs cate', startRow = i, startCol = 3, x = names(vars_cate)[vars_cate == var_main])
          writeData(wb, sheet = 'cate vs cate', startRow = i+2, startCol = 1, x = names(vars_cate)[vars_cate == var_minor])
          writeData(wb, sheet = 'cate vs cate', startRow = i+1, startCol = 2, x = tb3)
          
          i <- i + 2 + nrow(tb3)
          
          # Frequency Missing
          if (nMiss > 0) {
            mergeCells(wb, sheet = 'cate vs cate', rows = i, cols = 1:(2+ncol(tb3)))
            addStyle(wb, sheet = 'cate vs cate', rows = i, cols = 1:(2+ncol(tb3)), stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
            writeData(wb, sheet = 'cate vs cate', startRow = i, startCol = 1, x = sprintf('Frequency Missing = %d', nMiss))
            i <- i + 1
          }

          i <- i + 1
        }
      }
      
      # Sheet 2: Category (Main) vs Numeric (Minor)
      if (length(vars_minor_num) > 0) {
        cloneWorksheet(wb, sheetName = 'cate vs num', clonedSheet = 'Sheet1')
        i <- 4
        
        df <- select(data, var_main, vars_minor_num)
        df_Miss <- filter(df, is.na(pull(df, var_main)))
        df <- filter(df, !is.na(pull(df, var_main)))
        
        for (level in levels(pull(df, var_main))) {
          df.sub <- df[pull(df, var_main) == level,]
          df.sub.sub <- data.frame(Variable = sapply(vars_minor_num, function(var) names(vars_num)[vars_num == var]),
                                   N = apply(select(df.sub, vars_minor_num), MARGIN = 2, function(x) sum(!is.na(x))),
                                   `N Miss` = apply(select(df.sub, vars_minor_num), MARGIN = 2, function(x) sum(is.na(x))),
                                   Mean = apply(select(df.sub, vars_minor_num), MARGIN = 2, mean, na.rm = T),
                                   `Std Dev` = apply(select(df.sub, vars_minor_num), MARGIN = 2, sd, na.rm = T),
                                   IQR = apply(select(df.sub, vars_minor_num), MARGIN = 2, IQR, na.rm = T),
                                   Minimum = apply(select(df.sub, vars_minor_num), MARGIN = 2, min, na.rm = T),
                                   Q1 = apply(select(df.sub, vars_minor_num), MARGIN = 2, quantile, probs = 0.25, na.rm = T),
                                   Median = apply(select(df.sub, vars_minor_num), MARGIN = 2, median, na.rm = T),
                                   Q3 = apply(select(df.sub, vars_minor_num), MARGIN = 2, quantile, probs = 0.75, na.rm = T),
                                   Maximum = apply(select(df.sub, vars_minor_num), MARGIN = 2, max, na.rm = T)
                                   )
          
          mergeCells(wb, sheet = 'cate vs num', rows = i, cols = 1:ncol(df.sub.sub))
          
          addStyle(wb, sheet = 'cate vs num', rows = i:(i+1+nrow(df.sub.sub)), cols = 1:ncol(df.sub.sub), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
          addStyle(wb, sheet = 'cate vs num', rows = (i+2):(i+1+nrow(df.sub.sub)), cols = 4:ncol(df.sub.sub), stack = T, gridExpand = T, style = createStyle(numFmt = '0.00'))
          
          writeData(wb, sheet = 'cate vs num', startRow = i, startCol = 1, x = sprintf('%s = %s', names(vars_cate)[vars_cate == var_main], level))
          writeData(wb, sheet = 'cate vs num', startRow = i+1, startCol = 1, keepNA = T, x = df.sub.sub)
          
          i <- i + 2 + nrow(df.sub.sub) + 1
        }
        
        if (nrow(df_Miss) > 0) {
          df.sub <- df_Miss
          df.sub.sub <- data.frame(Variable = sapply(vars_minor_num, function(var) names(vars_num)[vars_num == var]),
                                   N = apply(select(df.sub, vars_minor_num), MARGIN = 2, function(x) sum(!is.na(x))),
                                   `N Miss` = apply(select(df.sub, vars_minor_num), MARGIN = 2, function(x) sum(is.na(x))),
                                   Mean = apply(select(df.sub, vars_minor_num), MARGIN = 2, mean, na.rm = T),
                                   `Std Dev` = apply(select(df.sub, vars_minor_num), MARGIN = 2, sd, na.rm = T),
                                   IQR = apply(select(df.sub, vars_minor_num), MARGIN = 2, IQR, na.rm = T),
                                   Minimum = apply(select(df.sub, vars_minor_num), MARGIN = 2, min, na.rm = T),
                                   Q1 = apply(select(df.sub, vars_minor_num), MARGIN = 2, quantile, probs = 0.25, na.rm = T),
                                   Median = apply(select(df.sub, vars_minor_num), MARGIN = 2, median, na.rm = T),
                                   Q3 = apply(select(df.sub, vars_minor_num), MARGIN = 2, quantile, probs = 0.75, na.rm = T),
                                   Maximum = apply(select(df.sub, vars_minor_num), MARGIN = 2, max, na.rm = T)
                                   )
          
          mergeCells(wb, sheet = 'cate vs num', rows = i, cols = 1:ncol(df.sub.sub))
          
          addStyle(wb, sheet = 'cate vs num', rows = i:(i+1+nrow(df.sub.sub)), cols = 1:ncol(df.sub.sub), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
          addStyle(wb, sheet = 'cate vs num', rows = (i+2):(i+1+nrow(df.sub.sub)), cols = 4:ncol(df.sub.sub), stack = T, gridExpand = T, style = createStyle(numFmt = '0.00'))
          
          writeData(wb, sheet = 'cate vs num', startRow = i, startCol = 1, x = sprintf('%s = %s', names(vars_cate)[vars_cate == var_main], 'Missing'))
          writeData(wb, sheet = 'cate vs num', startRow = i+1, startCol = 1, keepNA = T, x = df.sub.sub)
          
          i <- i + 2 + nrow(df.sub.sub) + 1
        }
      }
      
    } else {
      
      # Sheet 1: Numeric (Main) vs Category (Minor)
      if (length(vars_minor_cate) > 0) {
        cloneWorksheet(wb, sheetName = 'num vs cate', clonedSheet = 'Sheet1')
        i <- 4
        
        df <- data.frame(Variable = 'Overall',
                         Level = NA,
                         N = nrow(data),
                         `N Miss` = sum(is.na(pull(data, var_main))),
                         Mean = mean(pull(data, var_main), na.rm = T),
                         `Std Dev` = sd(pull(data, var_main), na.rm = T),
                         IQR = IQR(pull(data, var_main), na.rm = T),
                         Minimum = min(pull(data, var_main), na.rm = T),
                         Q1 = quantile(pull(data, var_main), probs = 0.25, na.rm = T),
                         Median = median(pull(data, var_main), na.rm = T),
                         Q3 = quantile(pull(data, var_main), probs = 0.75, na.rm = T),
                         Maximum = max(pull(data, var_main), na.rm = T)
                         )
        colnames(df) <- gsub(pattern = '\\.', replacement = ' ', x = colnames(df))
        
        for (var_minor in vars_minor_cate) {
          df.sub <- select(data, var_main, var_minor)
          df.sub <- df.sub[!is.na(pull(df.sub, var_minor)),]
          
          df.sub.sub <- df.sub %>%
            group_by_at(var_minor) %>%
            summarise(N = n(),
                      `N Miss` = sum(is.na(get(var_main))),
                      Mean = mean(get(var_main), na.rm = T),
                      `Std Dev` = sd(get(var_main), na.rm = T),
                      IQR = IQR(get(var_main), na.rm = T),
                      Minimum = min(get(var_main), na.rm = T),
                      Q1 = quantile(get(var_main), probs = 0.25, na.rm = T),
                      Median = median(get(var_main), na.rm = T),
                      Q3 = quantile(get(var_main), probs = 0.75, na.rm = T),
                      Maximum = max(get(var_main), na.rm = T)
                      )
          colnames(df.sub.sub)[1] <- 'Level'
          df.sub.sub <- cbind(data.frame(Variable = names(vars_cate)[vars_cate == var_minor]), df.sub.sub)
          
          df <- rbind(df, df.sub.sub)
        }
        
        mergeCells(wb, sheet = 'num vs cate', rows = i, cols = 1:ncol(df))
        mergeCells(wb, sheet = 'num vs cate', rows = i+2, cols = 1:2)
        addStyle(wb, sheet = 'num vs cate', rows = i+2, cols = 1:2, stack = T, style = createStyle(halign = 'center'))
        ibegin <- 0; iend <- 0;
        for (k in 2:nrow(df)) {
          if (k == 2) {
            ibegin <- 2
          } else {
            if (df$Variable[k] != df$Variable[k-1]) {
              iend <- k-1
              if (ibegin < iend) {
                mergeCells(wb, sheet = 'num vs cate', rows = (i+1+ibegin):(i+1+iend), cols = 1)
                addStyle(wb, sheet = 'num vs cate', rows = i+1+ibegin, cols = 1, stack = T, style = createStyle(valign = 'center'))
              }
              ibegin <- k
            }
          }
        }
        iend <- k
        if (ibegin < iend) {
          mergeCells(wb, sheet = 'num vs cate', rows = (i+1+ibegin):(i+1+iend), cols = 1)
          addStyle(wb, sheet = 'num vs cate', rows = i+1+ibegin, cols = 1, stack = T, style = createStyle(valign = 'center'))
        }
        
        addStyle(wb, sheet = 'num vs cate', rows = i:(i+1+nrow(df)), cols = 1:ncol(df), gridExpand = T, stack = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
        addStyle(wb, sheet = 'num vs cate', rows = (i+2):(i+1+nrow(df)), cols = 5:ncol(df), gridExpand = T, stack = T, style = createStyle(numFmt = '0.00'))
        
        writeData(wb, sheet = 'num vs cate', startRow = i, startCol = 1, x = names(vars_num)[vars_num == var_main])
        writeData(wb, sheet = 'num vs cate', startRow = i+1, startCol = 1, keepNA = T, x = df)
        
        i <- i + 2 + nrow(df) + 1
      }
      
      
      # Sheet 2: Numeric (Main) vs Numeric (Minor)
      if (length(vars_minor_num) > 0) {
        cloneWorksheet(wb, sheetName = 'num vs num', clonedSheet = 'Sheet1')
        i <- 4
        
        df <- select(data, var_main, vars_minor_num)
        df <- na.omit(df)
        colnames(df) <- c(names(vars_num)[vars_num == var_main], sapply(vars_minor_num, function(var) names(vars_num)[vars_num == var]))
        
        # Pearson Correlation
        tb <- cor(df, method = 'pearson')
        
        mergeCells(wb, sheet = 'num vs num', rows = i, cols = 1:(1+ncol(tb)))
        
        addStyle(wb, sheet = 'num vs num', rows = i:(i+1+nrow(tb)), cols = 1:(1+ncol(tb)), stack = T, gridExpand = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
        addStyle(wb, sheet = 'num vs num', rows = (i+2):(i+1+nrow(tb)), cols = 2:(1+ncol(tb)), stack = T, gridExpand = T, style = createStyle(numFmt = '0.0000'))
        
        writeData(wb, sheet = 'num vs num', startRow = i, startCol = 1, x = sprintf('Pearson Correlation Coefficients (N = %d)', nrow(df)))
        writeData(wb, sheet = 'num vs num', startRow = i+1, startCol = 1, rowNames = T, x = tb)
        
        i <- i + 2 + nrow(tb) + 1
        
        # Spearman Correlation
        tb <- cor(df, method = 'spearman')
        
        mergeCells(wb, sheet = 'num vs num', rows = i, cols = 1:(1+ncol(tb)))
        
        addStyle(wb, sheet = 'num vs num', rows = i:(i+1+nrow(tb)), cols = 1:(1+ncol(tb)), stack = T, gridExpand = T, style = createStyle(border = 'TopBottomLeftRight', halign = 'center'))
        addStyle(wb, sheet = 'num vs num', rows = (i+2):(i+1+nrow(tb)), cols = 2:(1+ncol(tb)), stack = T, gridExpand = T, style = createStyle(numFmt = '0.0000'))
        
        writeData(wb, sheet = 'num vs num', startRow = i, startCol = 1, x = sprintf('Spearman Correlation Coefficients (N = %d)', nrow(df)))
        writeData(wb, sheet = 'num vs num', startRow = i+1, startCol = 1, rowNames = T, x = tb)
      }
      
    }
    
    removeWorksheet(wb, 'Sheet1')
    saveWorkbook(wb, file, overwrite = T)
  }
  
}



