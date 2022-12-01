updatePasswordInput <- function(session, inputId, label = NULL, value = NULL) {
  dropNulls <- function(x) {x[!vapply(x, is.null, FUN.VALUE = logical(1))]}
  message <- dropNulls(list(label = label, value = value))
  session$sendInputMessage(inputId, message)
}

function(input, output, session) {
  
  tab0 <- reactiveValues(df = NULL, df_label = NULL, df_Tar = NULL, df_label_Enrolled = NULL, Vars = NULL,
                         DateRange_Tar = NULL, Total_Screen_Tar = NULL)
  tab1_1 <- reactiveValues(df = NULL)
  tab1_2 <- reactiveValues(vars_1_2_1 = NULL, vars_1_2_2 = NULL, vars_1_2_3 = NULL)
  tab1_3 <- reactiveValues(vars_1_3_cate = NULL, vars_1_3_num = NULL)
  
  # tab 0
  # Read Variables List
  output$result_fileinput <- renderText('Uploaded file:')
  is_vars <- F
  vars_list_filename <- NULL
  vars_list <- NULL
  observeEvent(input$file_import_vars, {
    is_vars <<- T
    vars_list_filename <<- input$file_import_vars$name
    vars_list <<- read.csv(input$file_import_vars$datapath, header = F)[,1]
    output$result_fileinput <- renderText(paste('Uploaded file:', vars_list_filename))
    output$result_search <- NULL
  })
  
  # Fileinput Reset
  observeEvent(input$reset, {
    is_vars <<- F
    vars_list_filename <<- NULL
    vars_list <<- NULL
    output$result_fileinput <- renderText('Uploaded file:')
    output$result_search <- NULL
  })
  
  
  observeEvent(input$search, {
    # Initialize contents
    output$tb_1_1 <- NULL
    ggplot_1_3_1 <<- list()
    
    # Import Data
    url <- 'https://www.redcap.ihrp.uic.edu/api/'
    
    # Read CodeBook
    formData <- list(token = input$token, content = 'metadata', format = 'csv', returnFormat = 'json')
    metadata <- POST(url, body = formData, encode = 'form') %>% content() %>% as.data.frame()
    
    # Read Data
    formData <- list(token = input$token, content = 'record', format = 'csv', type = 'flat')
    tab0$df <- POST(url, body = formData, encode = 'form') %>% content() %>% as.data.frame()
    
    # Remove Identifier & only read the variables in CodeBook
    tab0$df <- tab0$df[,colnames(tab0$df) %in% metadata$field_name[is.na(metadata$identifier)]]
    
    # Check necessary variables
    if (all(c(input$var_screen_date, input$var_is_eligible, input$var_exclude_reason, input$var_is_enrolled, input$var_enrolled_date) %in% colnames(tab0$df))) {
      # Rename Variable to fixed name
      colnames(tab0$df)[colnames(tab0$df) == input$var_screen_date] <- 'screen_date'
      colnames(tab0$df)[colnames(tab0$df) == input$var_is_eligible] <- 'is_eligible'
      colnames(tab0$df)[colnames(tab0$df) == input$var_exclude_reason] <- 'exclude_reason'
      colnames(tab0$df)[colnames(tab0$df) == input$var_is_enrolled] <- 'is_enrolled'
      colnames(tab0$df)[colnames(tab0$df) == input$var_enrolled_date] <- 'enrolled_date'
      metadata$field_name[metadata$field_name == input$var_screen_date] <- 'screen_date'
      metadata$field_name[metadata$field_name == input$var_is_eligible] <- 'is_eligible'
      metadata$field_name[metadata$field_name == input$var_exclude_reason] <- 'exclude_reason'
      metadata$field_name[metadata$field_name == input$var_is_enrolled] <- 'is_enrolled'
      metadata$field_name[metadata$field_name == input$var_enrolled_date] <- 'enrolled_date'
      
      # Remove unnecessary rows
      tab0$df <- tab0$df %>%
        filter(!is.na(screen_date))
      
      # Filter Variables in Variables List
      if (is_vars) {
        tab0$df <- tab0$df %>%
          select(any_of(c('screen_date', 'is_eligible', 'exclude_reason', 'is_enrolled', 'enrolled_date', vars_list)))
      }
      
      # Check weather data (token & variable list) is successfully imported
      if (nrow(tab0$df) > 0 & ncol(tab0$df) > 0) {
        output$result_search <- renderText('Dataset Successfully Imported!')
        
        # Data Processing
        # (1) Convert categorical variable to class 'factor'
        # (2) Create another data object with labels for categorical variable
        # Please simplify the labels (as less than two strings)
        # tab0$df$study_id <- factor(tab0$df$study_id)
        tab0$df_label <- tab0$df
        metadata_factor <- metadata[!is.na(metadata$select_choices_or_calculations), c('field_name', 'select_choices_or_calculations')]
        for (i in 1:nrow(metadata_factor)) {
          idx <- which(colnames(tab0$df) == metadata_factor$field_name[i])
          if (length(idx) == 0) next
          tmp <- str_split(str_split(metadata_factor[i,2], pattern = '\\|')[[1]], pattern = ', ')
          tab0$df[,idx] <- factor(tab0$df[,idx], levels = sapply(tmp, function(x) str_trim(x[1])), labels = sapply(tmp, function(x) str_trim(x[2])))
          tab0$df_label[,idx] <- factor(tab0$df_label[,idx], levels = sapply(tmp, function(x) str_trim(x[1])), labels = sapply(tmp, function(x) str_c(str_trim(x[1]), str_trim(x[2]), sep = '. ')))
        }
        rm(i, idx, tmp)
        
        
        # Variable Label
        tab0$Vars <- colnames(tab0$df)
        names(tab0$Vars) <- metadata$field_label[sapply(tab0$Vars, function(x) which(metadata$field_name == x))]
        
        # 
        tab1_2$vars_1_2_1 <- tab0$Vars[names(tab0$Vars) %in% metadata$field_label[!is.na(metadata$select_choices_or_calculations)]]
        tab1_2$vars_1_2_3 <- tab1_2$vars_1_2_1[tab1_2$vars_1_2_1 != 'is_eligible']
        tab1_2$vars_1_2_2 <- tab1_2$vars_1_2_3[!(tab1_2$vars_1_2_3 %in% c('is_enrolled', 'exclude_reason'))]
        tab1_3$vars_1_3_cate <- tab1_2$vars_1_2_2
        tab1_3$vars_1_3_num <- tab0$Vars[(names(tab0$Vars) %in% metadata$field_label[metadata$text_validation_type_or_show_slider_number %in% c('integer', 'number')]) & !(tab0$Vars %in% c('study_id', 'outcome'))]
        
        
        # Target Information
        tab0$DateRange_Tar <- input$DateRange_Tar
        tab0$Total_Screen_Tar <- input$Total_Screen_Tar
        tab0$df_Tar <- data.frame(Date = seq(tab0$DateRange_Tar[1], tab0$DateRange_Tar[2], by = 'day')) %>%
          mutate(CumN_Screen_Tar = (seq_len(n())/n()) * tab0$Total_Screen_Tar) %>%
          mutate(CumN_Enrolled_Tar = CumN_Screen_Tar * input$rate_eligible * input$rate_enrolled)
        
        
        # tab 1-1: Project Summary
        tab1_1$df <- data.frame(Index = c('Research Period',
                                          'Observation Period',
                                          'The Total Number of Screened Participants',
                                          'The Total Number of Eligible Participants',
                                          'The Total Number of Excluded Participants',
                                          sprintf('- %s', str_remove(levels(tab0$df_label$exclude_reason), pattern = '^[0-9]+\\.')),
                                          'The Total Number of Targeted Randomized Participants',
                                          'The Total Number of Actual Randomized Participants',
                                          sprintf('- Allocated to %s', str_remove(levels(tab0$df_label$treatment), pattern = '^[0-9]+\\.')),
                                          'Randomization Accrual Rate',
                                          'The Total number of Sites',
                                          'The Average Number of Randomized Participants per Site',
                                          'The Average Number Randomized Participants per Site per Month',
                                          'The Mean Number of Days from Screening to Randomization',
                                          'The Median Number of Days from Screening to Randomization'),
                                Content = c(str_c(format(tab0$DateRange_Tar[1], '%Y-%m-%d'), format(tab0$DateRange_Tar[2], '%Y-%m-%d'), sep = ' to '),
                                            str_c(format(tab0$DateRange_Tar[1], '%Y-%m-%d'), format(min(Sys.Date(), tab0$DateRange_Tar[2]), '%Y-%m-%d'), sep = ' to '),
                                            sum(!is.na(tab0$df$is_eligible)),
                                            sum(str_detect(tab0$df_label$is_eligible, pattern = '^1'), na.rm = T),
                                            sum(!is.na(tab0$df$exclude_reason)),
                                            as.numeric(table(tab0$df_label$exclude_reason)),
                                            round(tab0$Total_Screen_Tar*input$rate_eligible*input$rate_enrolled),
                                            sum(str_detect(tab0$df_label$is_enrolled, pattern = '^1'), na.rm = T),
                                            as.numeric(table(tab0$df_label$treatment)),
                                            sprintf('%.2f%%', 100*sum(str_detect(tab0$df_label$is_enrolled, pattern = '^1'), na.rm = T)/(tab0$Total_Screen_Tar*input$rate_eligible*input$rate_enrolled)),
                                            ifelse('site' %in% colnames(tab0$df), length(levels(tab0$df$site)), 1),
                                            (sum(str_detect(tab0$df_label$is_enrolled, pattern = '^1'), na.rm = T)/ifelse('site' %in% colnames(tab0$df), length(levels(tab0$df$site)), 1)) %>% round(2),
                                            (sum(str_detect(tab0$df_label$is_enrolled, pattern = '^1'), na.rm = T)/ifelse('site' %in% colnames(tab0$df), length(levels(tab0$df$site)), 1)/(interval(tab0$DateRange_Tar[1], min(as.Date('2020-01-31') + 1, tab0$DateRange_Tar[2])) %/% months(1))) %>% round(2),
                                            as.numeric(mean(difftime(tab0$df$enrolled_date, tab0$df$screen_date, units = 'days'), na.rm = T)) %>% round(2),
                                            as.numeric(median(difftime(tab0$df$enrolled_date, tab0$df$screen_date, units = 'days'), na.rm = T)))
        )
        output$tb_1_1 <- renderUI(
          tab1_1$df %>%
            addHtmlTableStyle(align.header = 'lr', align = 'lr',
                              css.header = 'background-color: gray; color: white; border: 2px solid black;',
                              css.cell = c('width: 800; border: 2px solid black;', 'width: 300; border: 2px solid black;')) %>%
            htmlTable(rnames = F)
        )
        
        
        # tab 1-2-1
        updateDateRangeInput(session, 'daterange_1_2_1', start = min(tab0$DateRange_Tar[1], tab0$df$screen_date, na.rm = T), end = max(tab0$DateRange_Tar[2], tab0$df$screen_date, na.rm = T), min = min(tab0$DateRange_Tar[1], tab0$df$screen_date, na.rm = T), max = max(tab0$DateRange_Tar[2], tab0$df$screen_date, na.rm = T))
        updateSelectInput(session, 'var_facet_1_2_1', choices = c('Overall' = 'None', tab1_2$vars_1_2_1))
        updateSelectInput(session, 'var_group_1_2_1', choices = c('Overall' = 'None', tab1_2$vars_1_2_1))
        
        # tab 1-2-2
        updateDateRangeInput(session, 'daterange_1_2_2', start = min(tab0$DateRange_Tar[1], tab0$df$enrolled_date, na.rm = T), end = max(tab0$DateRange_Tar[2], tab0$df$enrolled_date, na.rm = T), min = min(tab0$DateRange_Tar[1], tab0$df$enrolled_date, na.rm = T), max = max(tab0$DateRange_Tar[2], tab0$df$enrolled_date, na.rm = T))
        updateSelectInput(session, 'var_facet_1_2_2', choices = c('Overall' = 'None', tab1_2$vars_1_2_2))
        updateSelectInput(session, 'var_group_1_2_2', choices = c('Overall' = 'None', tab1_2$vars_1_2_2))
        
        # tab 1-2-3
        updateDateRangeInput(session, 'daterange_1_2_3', start = min(tab0$DateRange_Tar[1], tab0$df$screen_date, na.rm = T), end = max(tab0$DateRange_Tar[2], tab0$df$screen_date, na.rm = T), min = min(tab0$DateRange_Tar[1], tab0$df$screen_date, na.rm = T), max = max(tab0$DateRange_Tar[2], tab0$df$screen_date, na.rm = T))
        updateSelectInput(session, 'var_facet_1_2_3', choices = c('Overall' = 'None', tab1_2$vars_1_2_3))
        updateSelectInput(session, 'var_group_1_2_3', choices = c('Overall' = 'None', tab1_2$vars_1_2_3))
        
        
        # tab 1-3
        tab0$df_label_Enrolled <- tab0$df_label %>%
          filter(!is.na(enrolled_date))
        
        # tab 1-3-1
        updateSelectizeInput(session, 'vars_cate_1_3_1', choices = tab1_3$vars_1_3_cate)
        updateSelectizeInput(session, 'vars_num_1_3_1', choices = tab1_3$vars_1_3_num)
        
        # tab 1-3-2
        updateSelectizeInput(session, 'var_main_1_3_2', choices = c(tab1_3$vars_1_3_cate, tab1_3$vars_1_3_num))
        updateSelectizeInput(session, 'vars_minor_1_3_2', choices = c(tab1_3$vars_1_3_cate, tab1_3$vars_1_3_num))
        
      } else {
        output$result_search <- renderText('Fail to find the REDCap Data. (Please check your token or the variable list file.)')
      }
      
    } else {
      output$result_search <- renderText('Fail to find the REDCap Data. (Can not find your variable names in REDCap database.)')
    }
    
  })
  
  # Import Setting
  observeEvent(input$file_import_setting, {
    df <- fromJSON(read_json(input$file_import_setting$datapath)[[1]])
    # updateTextInput(session, 'token', value = df$Content[1])
    updatePasswordInput(session, 'token', value = df$Content[1])
    updateDateRangeInput(session, 'DateRange_Tar', start = as.Date(df$Content[2]), end = as.Date(df$Content[3]))
    updateNumericInput(session, 'Total_Screen_Tar', value = as.numeric(df$Content[4]))
    updateNumericInput(session, 'rate_eligible', value = as.numeric(df$Content[5]))
    updateNumericInput(session, 'rate_enrolled', value = as.numeric(df$Content[6]))
    updateTextInput(session, 'var_screen_date', value = df$Content[7])
    updateTextInput(session, 'var_is_eligible', value = df$Content[8])
    updateTextInput(session, 'var_exclude_reason', value = df$Content[9])
    updateTextInput(session, 'var_is_enrolled', value = df$Content[10])
    updateTextInput(session, 'var_enrolled_date', value = df$Content[11])
    
    is_vars <<- as.logical(df$Content[12])
    if (is_vars) {
      vars_list_filename <<- df$Content[13]
      vars_list <<- str_split(df$Content[14], pattern = '\\|')[[1]]
      output$result_fileinput <- renderText(paste('Uploaded file:', vars_list_filename))
    } else {
      vars_list_filename <<- NULL
      vars_list <<- NULL
      output$result_fileinput <- renderText('Uploaded file:')
    }
    output$result_search <- NULL
  })
  
  # Export Setting
  output$file_export_setting <- downloadHandler(
    filename = 'setting.json',
    content = function(file) {
      df <- data.frame(InputName = c('Token', 'Research_Period_begin', 'Research_Period_end', 'Total_Screen_Tar', 'Rate_Eligible', 'Rate_Enrolled', 'Var_Screen_Date', 'Var_Eligible', 'Var_Exclude_Reason', 'Var_Enrolled', 'Var_Enrolled_Date', 'Is_Vars', 'Vars_List_Filename', 'Vars_List'),
                       Content = c(input$token, format(input$DateRange_Tar[1]), format(input$DateRange_Tar[2]), input$Total_Screen_Tar, input$rate_eligible, input$rate_enrolled, input$var_screen_date, input$var_is_eligible, input$var_exclude_reason, input$var_is_enrolled, input$var_enrolled_date, is_vars, ifelse(is.null(vars_list_filename), '', vars_list_filename), str_c(vars_list, collapse = '|')))
      x <- toJSON(df)
      write_json(x, path = file)
    }
  )
  
  
  ##################################################
  
  
  # tab 1-2
  
  # facet variable -> Stratified variable
  # group variable -> Substratified variable
  
  
  ##################################################
  
  
  # tab 1-2-1
  
  observe({
    updateSelectInput(session, 'var_facet_1_2_1', choices = c('Overall' = 'None', tab1_2$vars_1_2_1) %>% .[. != ifelse(input$var_group_1_2_1 == 'None', '', input$var_group_1_2_1)], selected = input$var_facet_1_2_1)
    updateSelectInput(session, 'var_group_1_2_1', choices = c('Overall' = 'None', tab1_2$vars_1_2_1) %>% .[. != ifelse(input$var_facet_1_2_1 == 'None', '', input$var_facet_1_2_1)], selected = input$var_group_1_2_1)
  })
  
  tab_1_2_1 <- reactiveValues(var_facet = NULL, var_group = NULL,
                              df_export = NULL, gp_export = NULL)
  
  # (error)
  # tab_1_2_1$var_facet <- reactive(ifelse(input$var_facet_1_2_1 == 'None', NULL, input$var_facet_1_2_1))
  # tab_1_2_1$var_group <- reactive(ifelse(input$var_group_1_2_1 == 'None', NULL, input$var_group_1_2_1))
  tab_1_2_1$var_facet <- reactive(if (input$var_facet_1_2_1 == 'None') NULL else {input$var_facet_1_2_1})
  tab_1_2_1$var_group <- reactive(if (input$var_group_1_2_1 == 'None') NULL else {input$var_group_1_2_1})
  
  
  # Data Processing (DP)
  tab_1_2_1$df_export <- reactive({
    if (is.null(tab0$df_label)) {
      NULL
    } else {
      DP_1_2_1(data = tab0$df_label,
               ResearchPeriod = tab0$DateRange_Tar, unit = input$unit_1_2_1,
               var_facet = tab_1_2_1$var_facet(), var_group = tab_1_2_1$var_group())
    }
  })
  
  
  # Plotting (Plot)
  tab_1_2_1$gp_export <- reactive({
    if (is.null(tab_1_2_1$df_export())) {
      NULL
    } else {
      Plot_1_2_1(data = tab_1_2_1$df_export(),
                 ResearchPeriod = tab0$DateRange_Tar, ObservedPeriod = input$daterange_1_2_1, unit = input$unit_1_2_1,
                 var_facet = tab_1_2_1$var_facet(), var_group = tab_1_2_1$var_group())
    }
  })

  output$plot_1_2_1 <- renderPlotly({
    tab_1_2_1$gp_export()
  })
  
  
  # Export Html
  output$file_1_2_1_html <- downloadHandler(
    filename = 'Screened_Data.html',
    content = function(file) {
      saveWidget(tab_1_2_1$gp_export(), file, selfcontained = T)
    }
  )
  
  # Export Plot
  output$file_1_2_1_plot <- downloadHandler(
    filename = 'Screened_Data.png',
    content = function(file) {
      owd <- setwd(tempdir())
      on.exit(setwd(owd))
      saveWidget(tab_1_2_1$gp_export(), 'temp.html', selfcontained = F)
      webshot('temp.html', file = file)
    }
  )
  
  # Export Table (ET)
  output$file_1_2_1_table <- downloadHandler(
    filename = 'Screened_Data.xlsx',
    content = function(file) {
      if (is.null(tab_1_2_1$df_export())) {
        NULL
      } else {
        ET_1_2_1(file = file, data = tab_1_2_1$df_export(), Vars = tab0$Vars,
                 ResearchPeriod = tab0$DateRange_Tar, ObservedPeriod = input$daterange_1_2_1, unit = input$unit_1_2_1,
                 var_facet = tab_1_2_1$var_facet(), var_group = tab_1_2_1$var_group())
      }
    }
  )
  
  
  ##################################################
  
  
  # tab 1-2-2
  
  observe({
    updateSelectInput(session, 'var_group_1_2_2', choices = c('Overall' = 'None', tab1_2$vars_1_2_2) %>% .[. != ifelse(input$var_facet_1_2_2 == 'None', '', input$var_facet_1_2_2)], selected = input$var_group_1_2_2)
    updateSelectInput(session, 'var_facet_1_2_2', choices = c('Overall' = 'None', tab1_2$vars_1_2_2) %>% .[. != ifelse(input$var_group_1_2_2 == 'None', '', input$var_group_1_2_2)], selected = input$var_facet_1_2_2)
  })
  
  tab_1_2_2 <- reactiveValues(var_facet = NULL, var_group = NULL,
                              df_export = NULL, gp_export = NULL)
  
  # (error)
  # tab_1_2_2$var_facet <- reactive(ifelse(input$var_facet_1_2_2 == 'None', NULL, input$var_facet_1_2_2))
  # tab_1_2_2$var_group <- reactive(ifelse(input$var_group_1_2_2 == 'None', NULL, input$var_group_1_2_2))
  tab_1_2_2$var_facet <- reactive(if (input$var_facet_1_2_2 == 'None') NULL else {input$var_facet_1_2_2})
  tab_1_2_2$var_group <- reactive(if (input$var_group_1_2_2 == 'None') NULL else {input$var_group_1_2_2})
  
  
  # Data Processing (DP)
  tab_1_2_2$df_export <- reactive({
    if (is.null(tab0$df_label)) {
      NULL
    } else {
      DP_1_2_2(data = tab0$df_label_Enrolled, data_Tar = tab0$df_Tar,
               ResearchPeriod = tab0$DateRange_Tar, unit = input$unit_1_2_2,
               var_facet = tab_1_2_2$var_facet(), var_group = tab_1_2_2$var_group())
    }
  })
  
  
  # Plotting (Plot)
  tab_1_2_2$gp_export <- reactive({
    if (is.null(tab_1_2_1$df_export())) {
      NULL
    } else {
      Plot_1_2_2(data = tab_1_2_2$df_export(),
                 ResearchPeriod = tab0$DateRange_Tar, ObservedPeriod = input$daterange_1_2_2, unit = input$unit_1_2_2,
                 var_facet = tab_1_2_2$var_facet(), var_group = tab_1_2_2$var_group())
    }
  })
  
  output$plot_1_2_2 <- renderPlotly({
    tab_1_2_2$gp_export()
  })
  
  
  # Export Html
  output$file_1_2_2_html <- downloadHandler(
    filename = 'Randomized_Data.html',
    content = function(file) {
      saveWidget(tab_1_2_2$gp_export(), file, selfcontained = T)
    }
  )
  
  # Export Plot
  output$file_1_2_2_plot <- downloadHandler(
    filename = 'Randomized_Data.png',
    content = function(file) {
      owd <- setwd(tempdir())
      on.exit(setwd(owd))
      saveWidget(tab_1_2_2$gp_export(), 'temp.html', selfcontained = F)
      webshot('temp.html', file = file)
    }
  )
  
  # Export Table (ET)
  output$file_1_2_2_table <- downloadHandler(
    filename = 'Randomized_Data.xlsx',
    content = function(file) {
      if (is.null(tab_1_2_2$df_export())) {
        NULL
      } else {
        ET_1_2_2(file = file, data = tab_1_2_2$df_export(), Vars = tab0$Vars,
                 ResearchPeriod = tab0$DateRange_Tar, ObservedPeriod = input$daterange_1_2_2, unit = input$unit_1_2_2,
                 var_facet = tab_1_2_2$var_facet(), var_group = tab_1_2_2$var_group())
      }
    }
  )
  
  
  ##################################################
  
  
  # tab 1-2-3
  
  observe({
    updateSelectInput(session, 'var_facet_1_2_3', choices = c('Overall' = 'None', tab1_2$vars_1_2_3) %>% .[. != ifelse(input$var_group_1_2_3 == 'None', '', input$var_group_1_2_3)], selected = input$var_facet_1_2_3)
    updateSelectInput(session, 'var_group_1_2_3', choices = c('Overall' = 'None', tab1_2$vars_1_2_3) %>% .[. != ifelse(input$var_facet_1_2_3 == 'None', '', input$var_facet_1_2_3)], selected = input$var_group_1_2_3)
  })
  
  tab_1_2_3 <- reactiveValues(var_facet = NULL, var_group = NULL,
                              df_export = NULL, gp_export = NULL)
  
  tab_1_2_3$var_facet <- reactive(if (input$var_facet_1_2_3 == 'None') NULL else {input$var_facet_1_2_3})
  tab_1_2_3$var_group <- reactive(if (input$var_group_1_2_3 == 'None') NULL else {input$var_group_1_2_3})
  
  
  # Data Processing (DP)
  tab_1_2_3$df_export <- reactive({
    if (is.null(tab0$df_label)) {
      NULL
    } else {
      DP_1_2_3(data = tab0$df_label,
               ResearchPeriod = tab0$DateRange_Tar, unit = input$unit_1_2_3,
               var_facet = tab_1_2_3$var_facet(), var_group = tab_1_2_3$var_group())
    }
  })
  
  
  # Plotting (Plot)
  tab_1_2_3$gp_export <- reactive({
    if (is.null(tab_1_2_3$df_export())) {
      NULL
    } else {
      Plot_1_2_3(data = tab_1_2_3$df_export(),
                 ResearchPeriod = tab0$DateRange_Tar, ObservedPeriod = input$daterange_1_2_3, unit = input$unit_1_2_3,
                 var_facet = tab_1_2_3$var_facet(), var_group = tab_1_2_3$var_group())
    }
  })
  
  output$plot_1_2_3 <- renderPlotly({
    tab_1_2_3$gp_export()
  })
  
  
  # Export Html
  output$file_1_2_3_html <- downloadHandler(
    filename = 'Eligible_Data.html',
    content = function(file) {
      saveWidget(tab_1_2_3$gp_export(), file, selfcontained = T)
    }
  )
  
  # Export Plot
  output$file_1_2_3_plot <- downloadHandler(
    filename = 'Eligible_Data.png',
    content = function(file) {
      owd <- setwd(tempdir())
      on.exit(setwd(owd))
      saveWidget(tab_1_2_3$gp_export(), 'temp.html', selfcontained = F)
      webshot('temp.html', file = file)
    }
  )
  
  
  # Export Table (ET)
  output$file_1_2_3_table <- downloadHandler(
    filename = 'Eligible_Data.xlsx',
    content = function(file) {
      if (is.null(tab_1_2_3$df_export())) {
        NULL
      } else {
        ET_1_2_3(file = file, data = tab_1_2_3$df_export(), Vars = tab0$Vars,
                 ResearchPeriod = tab0$DateRange_Tar, ObservedPeriod = input$daterange_1_2_3, unit = input$unit_1_2_3,
                 var_facet = tab_1_2_3$var_facet(), var_group = tab_1_2_3$var_group())
      }
    }
  )
  
  
  ##################################################
  
  
  # tab 1-3-1
  
  # Plotting (Plot)
  output$plots_1_3_1 <- renderUI({
    plotoutput_list <- lapply(c(input$vars_cate_1_3_1, input$vars_num_1_3_1), function(var) {
      plotname <- paste('plot_1_3_1', var, sep = '_')
      plotlyOutput(plotname, width = '100%', height = 400)
    })
    tagList(plotoutput_list)
  })
  
  
  ggplot_1_3_1 <- list()
  
  # Category
  observeEvent(input$vars_cate_1_3_1, {
    if (!is.null(tab0$df_label_Enrolled)) {
      for (var in input$vars_cate_1_3_1) {
        local({
          plotname <- var
          outputname <- paste('plot_1_3_1', var, sep = '_')

          if (is.null(ggplot_1_3_1[[plotname]])) {
            ggplot_1_3_1[[plotname]] <<- Plot_1_3_1(data = tab0$df_label_Enrolled,
                                                    var = var, varType = 'cate', Vars = tab0$Vars)
            output[[outputname]] <<- renderPlotly({
              ggplotly(ggplot_1_3_1[[plotname]], tooltip = c('Count', 'label')) %>%
                layout(margin = list(t = 70)) %>%
                config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F)
            })
          }
        })
      }
    }
  })
  
  # Numeric
  observeEvent(input$vars_num_1_3_1, {
    if (!is.null(tab0$df_label_Enrolled)) {
      for (var in input$vars_num_1_3_1) {
        local({
          plotname <- var
          outputname <- paste('plot_1_3_1', var, sep = '_')

          if (is.null(ggplot_1_3_1[[plotname]])) {
            ggplot_1_3_1[[plotname]] <<- Plot_1_3_1(data = tab0$df_label_Enrolled,
                                                    var = var, varType = 'num', Vars = tab0$Vars,
                                                    nbins = as.numeric(input$nbins_1_3_1))
            output[[outputname]] <<- renderPlotly({
              ggplotly(ggplot_1_3_1[[plotname]], tooltip = c('count', 'text')) %>%
                layout(margin = list(l = 0, t = 70)) %>%
                config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F)
            })
          }
        })
      }
    }
  })
  
  observeEvent(input$nbins_1_3_1, {
    if (!is.null(tab0$df_label_Enrolled)) {
      for (var in input$vars_num_1_3_1) {
        local({
          plotname <- var
          outputname <- paste('plot_1_3_1', var, sep = '_')
          
          ggplot_1_3_1[[plotname]] <<- Plot_1_3_1(data = tab0$df_label_Enrolled,
                                                  var = var, varType = 'num', Vars = tab0$Vars,
                                                  nbins = as.numeric(input$nbins_1_3_1))
          output[[outputname]] <<- renderPlotly({
            ggplotly(ggplot_1_3_1[[plotname]], tooltip = c('count', 'text')) %>%
              layout(margin = list(l = 0, t = 70)) %>%
              config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F)
          })
        })
      }
    }
  })
  
  
  # Export Html
  output$file_1_3_1_html <- downloadHandler(
    filename = 'Univariate_Statistics.zip',
    content = function(file) {
      owd <- setwd(tempdir())
      on.exit(setwd(owd))
      
      files <- NULL
      for (var in input$vars_cate_1_3_1) {
        plotname <- var
        filename <- paste('plot_1_3_1_', var, '.html', sep = '')
        saveWidget(ggplotly(ggplot_1_3_1[[plotname]], tooltip = c('Count', 'label')) %>%
                     layout(margin = list(t = 70)) %>%
                     config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F), filename, selfcontained = T)
        files <- c(files, filename)
      }
      for (var in input$vars_num_1_3_1) {
        plotname <- var
        filename <- paste('plot_1_3_1_', var, '.html', sep = '')
        saveWidget(ggplotly(ggplot_1_3_1[[plotname]], tooltip = c('count', 'text')) %>%
                     layout(margin = list(l = 0, t = 70)) %>%
                     config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F), filename, selfcontained = T)
        files <- c(files, filename)
      }
      
      zip(file, files)
    }
  )
  
  # Export Plot
  output$file_1_3_1_plot <- downloadHandler(
    filename = 'Univariate_Statistics.zip',
    content = function(file) {
      owd <- setwd(tempdir())
      on.exit(setwd(owd))
      
      files <- NULL
      for (var in c(input$vars_cate_1_3_1, input$vars_num_1_3_1)) {
        filename <- paste('plot_1_3_1_', var, '.png', sep = '')
        ggsave(file = filename, plot = ggplot_1_3_1[[var]])
        files <- c(files, filename)
      }
      
      zip(file, files)
    }
  )
  
  # Export Table (ET)
  output$file_1_3_1_table <- downloadHandler(
    filename = 'Univariate_Statistics.xlsx',
    content = function(file) {
      if (is.null(tab0$df_label)) {
        NULL
      } else {
        ET_1_3_1(file = file, data = tab0$df_label_Enrolled, Vars = tab0$Vars,
                 ResearchPeriod = tab0$DateRange_Tar, ObservedPeriod = c(min(tab0$DateRange_Tar[1], Sys.Date()), min(tab0$DateRange_Tar[2], Sys.Date())),
                 vars_cate = input$vars_cate_1_3_1, vars_num = input$vars_num_1_3_1)
      }
    }
  )
  
  
  ##################################################
  
  
  # tab 1-3-2
  
  observeEvent(input$var_main_1_3_2, {
    updateSelectInput(session, 'vars_minor_1_3_2', choices = c(tab1_3$vars_1_3_cate, tab1_3$vars_1_3_num) %>% .[. != input$var_main_1_3_2], selected = NULL)
  })
  
  # Plotting (Plot)
  output$plots_1_3_2 <- renderUI({
    if (!is.null(input$vars_minor_1_3_2)) {
      plotoutput_list <- mapply(var1 = input$var_main_1_3_2,
                                var2 = input$vars_minor_1_3_2,
                                FUN = function(var1, var2) {
                                  plotname <- paste('plot_1_3_2', var1, var2, sep = '_')
                                  plotlyOutput(plotname, width = '100%', height = 400)
                                })
      tagList(plotoutput_list)
    }
  })
  
  
  ggplot_1_3_2 <- list()
  
  observeEvent(c(input$var_main_1_3_2, input$vars_minor_1_3_2), {
    if (!is.null(tab0$df_label_Enrolled)) {
      var1 <- input$var_main_1_3_2
      
      if (var1 %in% tab1_3$vars_1_3_cate) {
        
        # Category (Main) vs Category (Minor)
        for (var2 in intersect(tab1_3$vars_1_3_cate, input$vars_minor_1_3_2)) {
          local({
            plotname <- paste(var1, var2, sep = '_')
            outputname <- paste('plot_1_3_2', var1, var2, sep = '_')
            
            if (is.null(ggplot_1_3_2[[plotname]])) {
              ggplot_1_3_2[[plotname]] <<- Plot_1_3_2(data = tab0$df_label_Enrolled,
                                                      var1 = var1, var2 = var2, var1Type = 'cate', var2Type = 'cate', Vars = tab0$Vars)
              output[[outputname]] <<- renderPlotly({
                ggplotly(ggplot_1_3_2[[plotname]], tooltip = c('Count', 'label')) %>%
                  layout(margin = list(t = 70)) %>%
                  config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F)
              })
            }
          })
        }
        
        # Category (Main) vs Numeric (Minor)
        for (var2 in intersect(tab1_3$vars_1_3_num, input$vars_minor_1_3_2)) {
          local({
            plotname <- paste(var1, var2, sep = '_')
            outputname <- paste('plot_1_3_2', var1, var2, sep = '_')
            
            if (is.null(ggplot_1_3_2[[plotname]])) {
              ggplot_1_3_2[[plotname]] <<- Plot_1_3_2(data = tab0$df_label_Enrolled,
                                                      var1 = var1, var2 = var2, var1Type = 'cate', var2Type = 'num', Vars = tab0$Vars)
              output[[outputname]] <<- renderPlotly({
                ggplotly(ggplot_1_3_2[[plotname]]) %>%
                  layout(margin = list(t = 70)) %>%
                  config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F)
              })
            }
          })
        }
        
      } else {
        
        # Numeric (Main) vs Category (Minor)
        for (var2 in intersect(tab1_3$vars_1_3_cate, input$vars_minor_1_3_2)) {
          local({
            plotname <- paste(var1, var2, sep = '_')
            outputname <- paste('plot_1_3_2', var1, var2, sep = '_')
            
            if (is.null(ggplot_1_3_2[[plotname]])) {
              ggplot_1_3_2[[plotname]] <<- Plot_1_3_2(data = tab0$df_label_Enrolled,
                                                      var1 = var1, var2 = var2, var1Type = 'num', var2Type = 'cate', Vars = tab0$Vars)
              output[[outputname]] <<- renderPlotly({
                ggplotly(ggplot_1_3_2[[plotname]]) %>%
                  layout(margin = list(t = 70)) %>%
                  config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F)
              })
            }
          })
        }
        
        # Numeric (Main) vs Numeric (Minor)
        for (var2 in intersect(tab1_3$vars_1_3_num, input$vars_minor_1_3_2)) {
          local({
            plotname <- paste(var1, var2, sep = '_')
            outputname <- paste('plot_1_3_2', var1, var2, sep = '_')
            
            if (is.null(ggplot_1_3_2[[plotname]])) {
              ggplot_1_3_2[[plotname]] <<- Plot_1_3_2(data = tab0$df_label_Enrolled,
                                                      var1 = var1, var2 = var2, var1Type = 'num', var2Type = 'num', Vars = tab0$Vars)
              output[[outputname]] <- renderPlotly({
                ggplotly(ggplot_1_3_2[[plotname]]) %>%
                  layout(margin = list(t = 70)) %>%
                  config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F)
              })
            }
          })
        }
        
      }
    }
  })
  
  
  # Export Html
  output$file_1_3_2_html <- downloadHandler(
    filename = 'Bivariate_Statistics.zip',
    content = function(file) {
      owd <- setwd(tempdir())
      on.exit(setwd(owd))
      
      files <- NULL
      var1 <- input$var_main_1_3_2
      for (var2 in input$vars_minor_1_3_2) {
        plotname <- paste(var1, var2, sep = '_')
        filename <- paste('plot_1_3_2_', plotname, '.html', sep = '')
        
        if (var1 %in% tab1_3$Vars_1_3_cate & var2 %in% tab1_3$Vars_1_3_cate) {
          saveWidget(ggplotly(ggplot_1_3_2[[plotname]], tooltip = c('Count', 'label')) %>%
                       layout(margin = list(t = 70)) %>%
                       config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F), filename, selfcontained = T)
        } else {
          saveWidget(ggplotly(ggplot_1_3_2[[plotname]]) %>%
                       layout(margin = list(t = 70)) %>%
                       config(modeBarButtonsToRemove = c('toImage', 'select2d', 'lasso2d', 'autoScale2d', 'toggleSpikelines', 'hoverClosestCartesian', 'hoverCompareCartesian'), displaylogo = F), filename, selfcontained = T)
        }
        
        files <- c(files, filename)
      }
      
      zip(file, files)
    }
  )
  
  # Export Plot
  output$file_1_3_2_plot <- downloadHandler(
    filename = 'Bivariate_Statistics.zip',
    content = function(file) {
      owd <- setwd(tempdir())
      on.exit(setwd(owd))
      
      files <- NULL
      var1 <- input$var_main_1_3_2
      for (var2 in input$vars_minor_1_3_2) {
        plotname <- paste(var1, var2, sep = '_')
        filename <- paste('plot_1_3_2_', plotname, '.png', sep = '')
        ggsave(file = filename, plot = ggplot_1_3_2[[plotname]])
        files <- c(files, filename)
      }
      
      zip(file, files)
    }
  )
  
  # Export Table (ET)
  output$file_1_3_2_table <- downloadHandler(
    filename = 'Bivariate_Statistics.xlsx',
    content = function(file) {
      if (is.null(tab0$df_label)) {
        NULL
      } else {
        ET_1_3_2(file = file, data = tab0$df_label_Enrolled,
                 ResearchPeriod = tab0$DateRange_Tar, ObservedPeriod = c(min(tab0$DateRange_Tar[1], Sys.Date()), min(tab0$DateRange_Tar[2], Sys.Date())),
                 var_main = input$var_main_1_3_2, vars_minor = input$vars_minor_1_3_2,
                 vars_cate = tab1_3$vars_1_3_cate, vars_num = tab1_3$vars_1_3_num)
      }
    }
  )
  
  
  
  
  
  ##################################################
  
  
  
  
}