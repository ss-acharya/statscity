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
raw_prot <- read.csv("SQL20_TASTING_WATER_Jan2018.csv", na.strings = " ")

#Transform the sample dates into the correct format
raw_prot <- raw_prot %>% 
  mutate(., COLDATE = mdy(COLDATE))

raw_prot$COLDATE <- format(as.Date(raw_prot$COLDATE), "%d-%b-%Y")

raw_prot <- raw_prot %>%
  #mutate(., Location = str_replace_all(LOCCODE, pattern = "_", replacement = "")) %>% 
  arrange(COLDATE, LOCCODE)

# Helper function to get full results
paste_qualifiers <- function(x, y){
   if(x %in% c('<', '>', 'E')){
      return (paste(x, y))
   }

   return(y)

}



######################################################
######################################################
######################################################

### Process the raw data into three datasets: Total Coliform, E. coli, and Heterotrophic Bacteria (HPC) ###

prot_split <- split(raw_prot, raw_prot$ANALYTE)

for (i in 1:length(prot_split)) {
  
  Colname <- paste(as.character(prot_split[[i]][1,8])) #This line assigns the Location as the Result column name to match the existing report format - join next
  print(Colname)  

  #Split out the location/date pairs and create columns to join
  prot_split[[i]] <- prot_split[[i]] %>%
    mutate(Colname =  paste_qualifiers(QUALIFIER, RESULT)) %>% 
    select(LOCCODE, COLDATE, Colname)

  print(prot_split[[i]])


  colnames(prot_split[[i]])[3] <- paste('RESULT_', toString(i))
  print(prot_split[[i]])


        
}

#Join processed columns
prot_join <- plyr::join_all(prot_split, type="full" )

prot_join <- prot_join %>%
             select(LOCCODE, COLDATE, 'RESULT_ 3', 'RESULT_ 1', 'RESULT_ 2')

colnames(prot_join)[3:5] <- c('T_COLIFORM', 'E_COLI', 'HETERO_B')

#Count the number of columns which need center aligning
col_count <- length(unique(raw_prot$LOCCODE)) + 1

rep_border <- fp_border(color="black", width = 2)

#Output file construction below
#############################

dates <- unique(raw_prot$COLDATE)
locations <- unique(raw_prot$LOCCODE)

#Subset out the unique Location - Date pairs
loc_dates <- raw_prot %>% 
  distinct(LOCCODE, COLDATE, .keep_all = FALSE)

dates <- lapply(dates, function(x) x[!x %in% ""])
dates <- as.vector(unlist(dates))
report.date <- format(max(mdy(dates)), format = '%b-%Y') #This date is not used during processing, only to set the filename when exporting the saved report

csv_string <- paste('TASTING_WATER_Monthly_Report_',report.date,'.csv')
report_string <- paste('TASTING_WATER_Monthly_Report_',report.date,'.xlsx')
csv_path <- gsub(" ", "", csv_string, fixed = TRUE)
report_path <- gsub(" ", "", report_string, fixed = TRUE)


# Load the report into the docx template
doc <- read_docx("TASTING_WATER_Monthly_Report_Template.docx")
#doc <- doc %>%
       #body_add_par("Water Quality Chemistry Services" , pos = "after") %>%
       #body_add_par("SAN VICENTE MONTHLY REPORT", pos = "after") %>%
       #body_add_par("January 2018", pos = "after")


#Split by location
prot_join_split <- split(prot_join, prot_join$LOCCODE)

Loc_Date_pair_count = 1

for (i in 1:length(prot_join_split)) {
   
  #Split by collection date
  prot_join_split_two <- split(prot_join_split[[i]], prot_join$COLDATE)

  doc <- doc %>%
         body_add_par(paste("Sample Source:", loc_dates[Loc_Date_pair_count, 1]))

  for (j in 1:length(prot_join_split_two)){
    
    prot_join_split_iter <- prot_join_split_two[[j]] %>% select(-c('LOCCODE', 'COLDATE'))
 
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
                               #bold(., j = 1, bold = TRUE) %>% 
                               font(., fontname = "Open Sans", part = "all") %>% 
                               fontsize(., size = 8, part = "all") %>%
                               align(., j = c(2:col_count), align = 'center', part = 'all') %>% 
                               padding(., j = 1, padding.left = 5, part = "header") %>% 
                               padding(., padding.bottom = 5, part = "all") %>% 
                               line_spacing(., space = 0.75, part = "header") %>% 
                               autofit(.)

    doc <- doc %>%
         body_add_par(paste("Sample Date:", loc_dates[Loc_Date_pair_count, 2])) %>% 
         body_add_flextable(value = prot_join_split_iter_flex, split = TRUE, pos = "after")


   Loc_Date_pair_count = Loc_Date_pair_count + 1
  }


}


doc <- doc %>% 
    body_add_par("") %>%
    body_add_par("") %>%
    body_add_par("A = Absent") %>%
    body_add_par("P = Present") %>%
    body_add_par("< = Less Than") %>%
    body_add_par("> = Greater Than") %>%
    body_add_par("E = Overlimit/Estimate")

path <- paste('TASTING_WATER_Monthly_Report_',report.date,'.docx')  
target <- gsub(" ", "", path, fixed = TRUE)
  
print(doc, target = target)

