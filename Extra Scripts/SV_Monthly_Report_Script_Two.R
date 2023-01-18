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

### Process the raw data into 5 datasets (one for each week)

prot_split <- split(raw_prot, raw_prot$COLDATE)

for (i in 1:length(prot_split)) {
  
}
