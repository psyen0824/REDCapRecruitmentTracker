
fileInputOnlyButton <- function(..., label = '') {
  temp <- fileInput(..., label = label)
  # Cut away the label
  temp$children[[1]] <- NULL
  # Cut away the input field (after label is cut, this is position 1 now)
  temp$children[[1]]$children[[2]] <- NULL
  # Remove input group classes (makes button flat on one side)
  temp$children[[1]]$attribs$class <- NULL
  temp$children[[1]]$children[[1]]$attribs$class <- NULL
  temp
}

fileInputOnlyButton2 <- function(..., label = '') {
  temp <- fileInput(..., label = label)
  # Cut away the input field (after label is cut, this is position 1 now)
  temp$children[[2]]$children[[2]] <- NULL
  # Remove input group classes (makes button flat on one side)
  temp$children[[2]]$attribs$class <- NULL
  temp$children[[2]]$children[[1]]$attribs$class <- NULL
  temp
}

# header
header <- dashboardHeader(
  title = span(span('Recruitment Tracker', style = 'font-size: 20px; font-weight: bold;'),
               img(src = 'www/UIC_header.jpg', height = '100%', width = '85%', align = 'right')),
  titleWidth = '100%',
  
  # style setting
  tags$li(class = 'dropdown', tags$style('.skin-blue .main-header .logo {padding: 0px 0px; background-color: #007FA5}'))
)


# sidebar
sidebar <- dashboardSidebar(
  sidebarMenu(
    menuItem('Verification & Setting', icon = icon('th'), tabName = 'tab_0'),
    menuItem('Project Summary', icon = icon('th'), tabName = 'tab_1_1'),
    menuItem('Recruitment Tracking', icon = icon('th'), tabName = 'tab_1_2'),
    menuItem('Descriptive Statistics', icon = icon('th'), tabName = 'tab_1_3'),
    # menuItem('User Manual', icon = icon('th'), tabName = 'tab_1_4'),
    # menuItem('Clinical Trial', icon = icon('th'),
    #          menuSubItem('Project Summary', tabName = 'tab_1_1'),
    #          menuSubItem('Recruitment Tracking', tabName = 'tab_1_2'),
    #          menuSubItem('Descriptive Statistics', tabName = 'tab_1_3'),
    #          menuSubItem('User Manual', tabName = 'tab_1_4')),
    # menuItem('Longitudinal Study', icon = icon('th'),
    #          menuSubItem('Project Summary', tabName = 'tab_2_1'),
    #          menuSubItem('Descriptive Statistics', tabName = 'tab_2_2'),
    #          menuSubItem('User Manual', tabName = 'tab_2_3')),
    # menuItem('Cross-Sectional Study', icon = icon('th'),
    #          menuSubItem('Project Summary', tabName = 'tab_3_1'),
    #          menuSubItem('Descriptive Statistics', tabName = 'tab_3_2'),
    #          menuSubItem('User Manual', tabName = 'tab_3_3')),
    br(),
    # span('Project Title', style = 'margin-left: 10px; font-size: 16px; font-weight: bold; text-decoration: underline;'),
    # br(),
    # span('Clinical Trial Recruitment', style = 'margin-left: 10px; font-size: 14px; font-weight: lighter;'),
    # br(),
    # span('Tracking Shiny App', style = 'margin-left: 10px; font-size: 14px; font-weight: lighter;'),
    # br(),
    # br(),
    # br(),
    a('User Manual', href = 'https://github.com/psyen0824/REDCapTracking/blob/main/RShiny%20User%20Manual_20221126.pdf', style = 'margin-left: 10px; font-size: 16px; font-weight: bold; text-decoration: underline; color: white;'),
    br(),
    br(),
    a('Contact us', href = 'https://ccts.uic.edu/services/biostatistics/personnel/', style = 'margin-left: 10px; font-size: 16px; font-weight: bold; text-decoration: underline; color: white;')
  ),
  
  # style setting
  tags$style(HTML('.skin-blue .main-sidebar {background-color: #007FA5; font-size: 16px; font-weight: bold;}')),
  tags$style(HTML('.skin-blue .treeview-menu > li > a {font-size: 16px; font-weight: bold}')),
  tags$script(JS("document.getElementsByClassName('sidebar-toggle')[0].style.visibility = 'hidden';"))
)


# body
body <- dashboardBody(
  tabItems(
    
    # tab 0
    tabItem(tabName = 'tab_0',
            splitLayout(
              cellWidths = c('150px', '100px'),
              fileInputOnlyButton2('file_import_setting', label = 'Setting (Optional)', buttonLabel = 'Import', accept = c('.json')),
              downloadButton('file_export_setting', label = 'Export', style = 'width: 80px; margin-top: 25px; margin-right: 20px; background-color: forestgreen; color: white;')
            ),
            hr(style = 'border-color: black; margin-top: 5px;'),
            splitLayout(
              cellWidths = c('300px', '300px'),
              # textInput('token', label = 'REDCap Project Token'),
              passwordInput('token', label = 'REDCap Project Token'),
              dateRangeInput('DateRange_Tar', label = 'Research Period of this Clinical Trial Project')
            ),
            splitLayout(
              cellWidths = c('300px', '300px', '300px'),
              numericInput('Total_Screen_Tar', label = 'Total Targeted Number of Screened Participants', value = 1000, min = 0),
              numericInput('rate_eligible', label = 'Expected Rate of Eligibility', value = 0.1, min = 0, max = 1),
              numericInput('rate_enrolled', label = 'Expected Rate of Randomization', value = 0.9, min = 0, max = 1)
            ),
            splitLayout(
              cellWidths = c('270px', '270px', '270px', '270px', '270px'),
              textInput('var_screen_date', label = 'Variable for Screening Date'),
              textInput('var_is_eligible', label = 'Variable for Participant Eligibility'),
              textInput('var_exclude_reason', label = 'Variable for Exclusion Reason'),
              textInput('var_is_enrolled', label = 'Variable for Receipt of Randomization'),
              textInput('var_enrolled_date', label = 'Variable for Randomization Date')
            ),
            splitLayout(
              cellWidths = c('250px', '125px'),
              fileInputOnlyButton2('file_import_vars', label = 'Variable List (CSV file only) (Optional)', buttonLabel = 'Import', accept = c('.csv')),
              actionButton('reset', label = 'Reset', style = 'background-color: blueviolet; color: white; margin-top: 25px; margin-left: 10px')
            ),
            verbatimTextOutput('result_fileinput'),
            hr(style = 'border-color: black'),
            actionButton('search', label = 'Search', style = 'background-color: blue; color: white;'),
            verbatimTextOutput('result_search')
    ),
    
    # tab 1-1
    tabItem(tabName = 'tab_1_1',
            h2('Project Summary'),
            br(),
            htmlOutput('tb_1_1')
    ),
    
    # tab 1-2
    tabItem(tabName = 'tab_1_2',
            tabsetPanel(
              tabPanel(title = 'Screened Data',
                       br(),
                       column(width = 3,
                              box(title = 'Input', status = 'primary', solidHeader = T, width = 12, height = 550,
                                  dateRangeInput('daterange_1_2_1', label = 'Observation Period'),
                                  selectInput('unit_1_2_1', label = 'Unit of Time', choices = c('Week', 'Month', 'Quarter', 'Year'), selected = 'Month'),
                                  selectInput('var_facet_1_2_1', label = 'Stratified Variable', choices = c('Overall' = 'None'), selected = 'None'),
                                  selectInput('var_group_1_2_1', label = 'Substratified Variable', choices = c('Overall' = 'None'), selected = 'None'))),
                       column(width = 9,
                              box(title = 'Output', status = 'warning', solidHeader = T, width = 12, height = 550,
                                  downloadButton('file_1_2_1_table', label = 'EXCEL', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  downloadButton('file_1_2_1_plot', label = 'PNG', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  downloadButton('file_1_2_1_html', label = 'HTML', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  br(),
                                  br(),
                                  plotlyOutput('plot_1_2_1', width = '100%', height = 450)))
              ),
              tabPanel(title = 'Eligible Data',
                       br(),
                       column(width = 3,
                              box(title = 'Input', status = 'primary', solidHeader = T, width = 12, height = 550,
                                  dateRangeInput('daterange_1_2_3', label = 'Observation Period'),
                                  selectInput('unit_1_2_3', label = 'Unit of Time', choices = c('Week', 'Month', 'Quarter', 'Year'), selected = 'Month'),
                                  selectInput('var_facet_1_2_3', label = 'Stratified Variable', choices = c('Overall' = 'None'), selected = 'None'),
                                  selectInput('var_group_1_2_3', label = 'Substratified Variable', choices = c('Overall' = 'None'), selected = 'None'))),
                       column(width = 9,
                              box(title = 'Output', status = 'warning', solidHeader = T, width = 12, height = 550,
                                  downloadButton('file_1_2_3_table', label = 'EXCEL', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  downloadButton('file_1_2_3_plot', label = 'PNG', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  downloadButton('file_1_2_3_html', label = 'HTML', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  br(),
                                  br(),
                                  plotlyOutput('plot_1_2_3', width = '100%', height = 450)))
              ),
              tabPanel(title = 'Randomized Data',
                       br(),
                       column(width = 3,
                              box(title = 'Input', status = 'primary', solidHeader = T, width = 12, height = 550,
                                  dateRangeInput('daterange_1_2_2', label = 'Observation Period'),
                                  selectInput('unit_1_2_2', label = 'Unit of Time', choices = c('Week', 'Month', 'Quarter', 'Year'), selected = 'Month'),
                                  selectInput('var_facet_1_2_2', label = 'Stratified Variable', choices = c('Overall' = 'None'), selected = 'None'),
                                  selectInput('var_group_1_2_2', label = 'Substratified Variable', choices = c('Overall' = 'None'), selected = 'None'))),
                       column(width = 9,
                              box(title = 'Output', status = 'warning', solidHeader = T, width = 12, height = 550,
                                  downloadButton('file_1_2_2_table', label = 'EXCEL', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  downloadButton('file_1_2_2_plot', label = 'PNG', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  downloadButton('file_1_2_2_html', label = 'HTML', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  br(),
                                  br(),
                                  plotlyOutput('plot_1_2_2', width = '100%', height = 450)))
              )
            )
    ),
    
    # tab 1-3
    tabItem(tabName = 'tab_1_3',
            tabsetPanel(
              tabPanel(title = 'Univariate Analysis',
                       br(),
                       column(width = 3,
                              box(title = 'Input (Randomized Data)', status = 'primary', solidHeader = T, width = 12, height = 550,
                                  selectizeInput(inputId = 'vars_cate_1_3_1', label = 'Categorical Variable', choices = 'NULL', selected = NULL, multiple = T),
                                  selectizeInput(inputId = 'vars_num_1_3_1', label = 'Numerical Variable', choices = 'NULL', selected = NULL, multiple = T),
                                  numericInput(inputId = 'nbins_1_3_1', label = 'Number of bins of histogram', value = 10, min = 1, step = 1))),
                       column(width = 9,
                              box(title = 'Output', status = 'warning', solidHeader = T, width = 12, height = 550,
                                  downloadButton('file_1_3_1_table', label = 'EXCEL', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  downloadButton('file_1_3_1_plot', label = 'PNG', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  downloadButton('file_1_3_1_html', label = 'HTML', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  br(),
                                  br(),
                                  uiOutput('plots_1_3_1')))
              ),
              tabPanel(title = 'Bivariate Analysis',
                       br(),
                       column(width = 3,
                              box(title = 'Input (Randomized Data)', status = 'primary', solidHeader = T, width = 12, height = 550,
                                  selectizeInput(inputId = 'var_main_1_3_2', label = HTML('Primary Variable<br/>(fixed option)'), choices = 'NULL'),
                                  selectizeInput(inputId = 'vars_minor_1_3_2', label = HTML('Stratified by<br/>(one or more options)'), choices = 'NULL', selected = NULL, multiple = T))),
                       column(width = 9,
                              box(title = 'Output', status = 'warning', solidHeader = T, width = 12, height = 550,
                                  downloadButton('file_1_3_2_table', label = 'EXCEL', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  downloadButton('file_1_3_2_plot', label = 'PNG', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  downloadButton('file_1_3_2_html', label = 'HTML', style = 'width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;'),
                                  br(),
                                  br(),
                                  uiOutput('plots_1_3_2')))
              )
            )
    )
  ),
  
  # style setting
  tags$style(".shiny-file-input-progress {display: none}"),
  tags$style(HTML('.btn.btn-default.btn-file {width: 80px; background-color: forestgreen; color: white;}'))
  # tags$style(HTML('.btn.btn-default.shiny-download-link {width: 80px; margin-right: 20px; float: right; background-color: forestgreen; color: white;}'))
)

# Create Dashboard page
dashboardPage(title = 'REDCap Tracking', header, sidebar, body)


