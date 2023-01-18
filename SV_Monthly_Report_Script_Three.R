#setwd("D:\\R_Data")
#setwd("E:\\R_Data")
rm(list=ls(all=TRUE))

options(stringsAsFactors=F, warn = -1)
options(xtable.comment = FALSE)
if (!require(gt)) {
  install.packages("gt", repos = "https://cloud.r-project.org/")
  library(gt)
}

if (!require(tidyverse)) {
  install.packages("tidyverse")
  library(tidyverse)
}

if (!require(Hmisc)) {
  install.packages("Hmisc")
  library(Hmisc)
}

if (!require(gtools)) {
  install.packages("gtools")
  library(gtools)
}

if (!require(openxlsx)) {
  install.packages("openxlsx")
  library(openxlsx)
}

if (!require(dplyr)) {
  install.packages("dplyr")
  library(dplyr)
}

#Have R alarm you when it's done running
#beep() is used at the ends of the three data pull sections of this script to alert you when longer processes are finished running

if (!require(beepr)) {
  install.packages("beepr")
  library(beepr)
}
#beep(8) #Long sound
#beep(5) #Medium sound
beep(11) #Short sound

if (!require(xtable)) {
  install.packages("xtable", repos = "https://cloud.r-project.org/")
  library(xtable)
}

if (!require(rlang)) {
  install.packages("rlang", repos = "https://cloud.r-project.org/")
  library(rlang)
}

if (!require(tidystringdist)) {
  install.packages("tidystringdist", repos = "https://cloud.r-project.org/")
  library(tidystringdist)
}

if (!require(stringdist)) {
  install.packages("stringdist", repos = "https://cloud.r-project.org/")
  library(stringdist)
}

if (!require(stringi)) {
  install.packages("stringi", repos = "https://cloud.r-project.org/")
  library(stringi)
}

if (!require(readxl)) {
  install.packages("readxl", repos = "https://cloud.r-project.org/")
  library(readxl)
}

if (!require(lubridate)) {
  install.packages("lubridate", repos = "https://cloud.r-project.org/")
  library(lubridate)
}

if (!require(data.table)) {
  install.packages("data.table", repos = "https://cloud.r-project.org/")
  library(data.table)
}

if (!require(flextable)) {
  install.packages("flextable", repos = "https://cloud.r-project.org/")
  library(flextable)
}

if (!require(kableExtra)) {
  install.packages("kableExtra", repos = "https://cloud.r-project.org/")
  library(kableExtra)
}

if (!require(plyr)) {
  install.packages("plyr", repos = "https://cloud.r-project.org/")
  library(plyr)
}

if (!require(officer)) {
  install.packages("officer")
  library(officer)
}

### Data loading step

#Import the raw data
raw_prot <- read.csv("SQL17_SV_MONTHLY_Jan2018.csv", na.strings = "")

#Transform the sample dates into the correct format
raw_prot <- raw_prot %>% 
  mutate(., COLDATE = mdy(COLDATE))

raw_prot$COLDATE <- format(as.Date(raw_prot$COLDATE), "%d-%b-%Y")

raw_prot <- raw_prot %>%
  #mutate(., Location = str_replace_all(LOCCODE, pattern = "_", replacement = "")) %>% 
  arrange(COLDATE, LOCCODE)


######################################################
######################################################
######################################################

### Process the raw data into three datasets: Total Coliform, E. coli, and Enterococcus ###

prot_split <- split(raw_prot, raw_prot$LOCCODE)

for (i in 1:length(prot_split)) {
  
  Colname <- paste(as.character(prot_split[[i]][1,2])) #This line assigns the Location as the Result column name to match the existing report format - join next
  
  names(prot_split[[i]])[9] <- Colname
  
  #Split out the location/date pairs and create columns to join
  prot_split[[i]] <- prot_split[[i]] %>% 
    select(ANALYTE, COLDATE, MDL, Colname)  
}

#Join processed columns
prot_join <- plyr::join_all(prot_split, type="full" ) 

#Rearrange the data
prot_join <- prot_join %>% 
  mutate_each(funs(replace(., which(is.na(.)),"---"))) %>%
  mutate(ANALYTE = case_when(ANALYTE == '2-Methylisoborneol' ~ 'MIB',
                             str_detect(ANALYTE, 'Color') ~ 'COLOR',
                             ANALYTE == 'Geosmin' ~ 'GEOSMIN',
                             str_detect(ANALYTE, 'Odor') ~ 'ODOR',
                             str_detect(ANALYTE, 'TOC') ~ 'TOC',
                             ANALYTE == 'Transparency' ~ 'TRANSPARENCY',
                             str_detect(ANALYTE, 'Turbidity') ~ 'TURBIDITY')) %>%

  arrange(ANALYTE)

#print(prot_join %>% select(ANALYTE))

col_list <- names(prot_join) #Create a list of column names (locations) by substituting "_" for " " in SYS locations
col_list <- gsub("_", " ", col_list, fixed = TRUE)
names(prot_join) <- col_list #Assign the new column names to the working dataset

#Split by collection date
prot_join_split <- split(prot_join, prot_join$COLDATE)

date_headers <- as.vector(loc_dates[['COLDATE']])

date_headers <- as.vector(date_headers)

#Count the number of columns which need center aligning
col_count <- length(unique(raw_prot$LOCCODE)) + 1

rep_border <- fp_border(color="black", width = 2)

#Output file construction below
#############################

dates <- unique(raw_prot$COLDATE)
dates <- lapply(dates, function(x) x[!x %in% ""])
dates <- as.vector(unlist(dates))
report.date <- format(max(mdy(dates)), format = '%b-%Y') #This date is not used during processing, only to set the filename when exporting the saved report

csv_string <- paste('SV_Monthly_Report_',report.date,'.csv')
report_string <- paste('SV_Monthly_Report_',report.date,'.xlsx')
csv_path <- gsub(" ", "", csv_string, fixed = TRUE)
report_path <- gsub(" ", "", report_string, fixed = TRUE)


# Load the report into the docx template
doc <- read_docx("San_Vicente_Chemistry_Report_Template.docx")
#doc <- doc %>%
       #body_add_par("Water Quality Chemistry Services" , pos = "after") %>%
       #body_add_par("SAN VICENTE MONTHLY REPORT", pos = "after") %>%
       #body_add_par("January 2018", pos = "after")


for (i in 1:length(prot_join_split)) {

  prot_join_split_iter <- prot_join_split[[i]] %>% select(-c('COLDATE'))
 
  prot_join_split_iter_flex <- flextable(prot_join_split_iter)


  prot_join_split_iter_flex <- prot_join_split_iter_flex %>%
                               border_remove(.) %>%
                               border_outer(., part = "all", border = rep_border) %>% 
                               hline(., part = "body", border = rep_border) %>% 
                               vline(., part = "all", border = rep_border) %>% 
                               bold(., bold = TRUE, part = "header") %>%
                               align(., align = "center", part = "all") %>%
                               bg(., bg = 'lightskyblue', part = 'header') %>%
                               fontsize(., size = 19, part = 'body') %>%
                               bold(., j = 1, bold = TRUE) %>% 
                               font(., fontname = "Open Sans", part = "all") %>% 
                               fontsize(., size = 8, part = "all") %>%
                               align(., j = c(2:col_count), align = 'center', part = 'all') %>% 
                               padding(., j = 1, padding.left = 5, part = "header") %>% 
                               padding(., padding.bottom = 5, part = "all") %>% 
                               line_spacing(., space = 0.75, part = "header") %>% 
                               autofit(.)

  doc <- doc %>% 
         body_add_par(paste("Sample Date:", unique(dates)[i])) %>%
         body_add_flextable(value = prot_join_split_iter_flex, split = TRUE, pos = "after")
         

}

doc <- doc %>% 
    body_add_par("") %>%
    body_add_par("") %>%
    body_add_par("        Notes:", style = "Normal",pos = "after") %>%
    body_add_par("        LA = Lab Accident", style = "Normal",pos = "after") %>%
    body_add_par("        NR = Not Reportable", style = "Normal",pos = "after") %>%
    body_add_par("        NS = Not Sampled", style = "Normal",pos = "after") %>%
    body_add_par("        NA = Not Analyzed", style = "Normal",pos = "after") %>%
    body_add_par("        ND = Analytes not detected with the following MDLs:") %>%
    body_add_par("        NH3-N: 0.0361 mg/L", style = "Normal",pos = "after") %>%
    body_add_par("        NO2: 0.0156 mg/L", style = "Normal",pos = "after") %>%
    body_add_par("        Total chlorine: 0.1 mg/L", style = "Normal",pos = "after") %>%
    body_add_par("        Free chlorine: 0.2 mg/L", style = "Normal",pos = "after") %>%
    body_add_par("") %>%
    body_add_par("Reviewed and approved:", style = "footer",pos = "after") %>%
    body_add_par("") %>%
    body_add_par("________________________________________   Date:______________", style = "Normal") %>%
    body_add_par("Dan Silvaggio, Senior Biologist", style = "footer") %>%
    body_add_par("EMTS Division, Microbiology Section", style = "footer") %>% 
    body_add_par("") %>%
    body_add_par("________________________________________   Date:______________", style = "Normal") %>%
    body_add_par("Daniel Daft, Water Production Superintendent", style = "footer")


path <- paste('SV_Monthly_Report_',report.date,'.docx')  
target <- gsub(" ", "", path, fixed = TRUE)
  
print(doc, target = target)








