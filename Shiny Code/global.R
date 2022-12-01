library(shiny)
library(shinydashboard)

library(htmlwidgets)
library(webshot)
if (is.null(suppressMessages(webshot:::find_phantom()))) {webshot::install_phantomjs()}
addResourcePath('www', 'www')

library(httr)

library(dplyr); options(dplyr.summarise.inform = F)
library(stringr)
library(lubridate)
library(jsonlite)
library(readr)

library(htmlTable)
library(plotly)

source('function/fun_DataProcessing.R')
source('function/fun_Plotting.R')
source('function/fun_ExportTable.R')
