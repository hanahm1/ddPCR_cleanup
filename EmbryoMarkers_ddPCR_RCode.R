## ddPCR Embryo Script ## 
install.packages("rstatix")
install.packages("ggpubr")
install.packages("tidyverse")

library(tidyr)
library(dplyr)
library(stringi)
library(openxlsx)
library(rstatix)
library(ggpubr)
library(tidyverse)

######## EDIT !!!!!!! ##################################################
Plate1 <- read.csv("", header = FALSE)
Plate2 <- read.csv("", header = FALSE)
Plate3 <- read.csv("", header = FALSE)
Plate4 <- read.csv("", header = FALSE)
Plate5 <- read.csv("", header = FALSE)

folder_path <- "/Users/hanah/Google Drive/MolMark_Embryos/Finished_Files/Aim1_Round1"
  
output_name <- "102420_Round1_Embryo_ddPCR_Analysis.xlsx"
#####################################################################
# View(Plate1)
# View(Plate2)
# View(Plate3)

########### Do NOT Edit Below this line ###################


### function to cleanup data into workable format ##

workableFun <- function(plateFileVariable) {
  colnames(plateFileVariable) <- plateFileVariable [1,]
  cleanPlate <- plateFileVariable [,c("Well","Sample", "Target", "Concentration", "CopiesPer20uLWell")]
  cleanPlate <- cleanPlate [-1,]
  cleanPlate$Split <- cleanPlate$Sample
  cleanPlate2 <- cleanPlate [,c("Well", "Sample", "Split", "Target", "Concentration", "CopiesPer20uLWell")]
  for(i in 1:nrow(cleanPlate2)){
    SplitID <- cleanPlate2$Split [i]
    if(SplitID == "NoTemp"){
      cleanPlate2$Split [i] <- "NT.NT.NT"
    } else if(SplitID == "Sperm"){
      cleanPlate2$Split [i] <- "Sp.Sp.Sp"
    } else(print("OK"))
  }
  workableDf <- separate(cleanPlate2, Split, c("Round", "Day", "ID"))
  for(i in 1:nrow(workableDf)){
    IDChange <- workableDf$ID [i]
    if(workableDf$Day [i] == 0){
      workableDf$ID [i] <- paste0("G", IDChange)
    } else(print('OK'))
  }
  stri_sub(workableDf$ID, 2, 1) <- "."
  workableDf2 <- separate(workableDf, ID, c("Grade", "ID"))
  for(i in 1:nrow(workableDf2)){
    if(workableDf2$Concentration [i] == "No Call"){
      workableDf2$Concentration [i] <- 0
    } else (print("OK"))
  }
  for(i in 1:nrow(workableDf2)){
    if(workableDf2$Target [i] == "Target1"){
      workableDf2$Target [i] <- "Target1"
    } else if(workableDf2$Target [i] == "Target2"){
      workableDf2$Target [i] <- "Target2"
    }
  }
  return(workableDf2)
} 
########## End Function ###############

Plate1Clean <- workableFun(Plate1)
Plate2Clean <- workableFun(Plate2)
Plate3Clean <- workableFun(Plate3)
Plate4Clean <- workableFun(Plate4)
Plate5Clean <- workableFun(Plate5)

#View(Plate2Clean)

#View(AllPlates)

# Plate1Clean$PlateRep <- 1
# Plate2Clean$PlateRep <- 2

## make a master df with all plate infor ##
AllPlates <- Plate1Clean
AllPlates <- rbind(AllPlates, Plate2Clean)
AllPlates <- rbind(AllPlates, Plate3Clean)
AllPlates <- rbind(AllPlates, Plate4Clean)
AllPlates <- rbind(AllPlates, Plate5Clean)
AllPlates$Concentration <- as.numeric(AllPlates$Concentration)
View(AllPlates)
class(AllPlates$Concentration)
AllPlates$Concentration <- as.numeric(AllPlates$Concentration)
## Get Means sorted by Day, Grade, and Target ##
MeanDf <- aggregate(Concentration ~ Day + Grade + Target, FUN=mean, data = AllPlates)
View(MeanDf)
MeanDf[43,]

## Add appropriate house keeper means to AllPlates row (Sorted by Grade and Day) ##
# double check YWHAZ Rows #

for(i in 1:nrow(AllPlates)){
  if(AllPlates$Grade [i] == "G"){
    if(AllPlates$Day [i] == 0){
      AllPlates$HKMean [i] <- MeanDf$Concentration [42]
    } else if(AllPlates$Day [i] == 1) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [43]
    } else if(AllPlates$Day [i] == 2) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [44]
    } else if(AllPlates$Day [i] == 3) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [45]
    } else if(AllPlates$Day [i] == 4) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [46]
    } else if(AllPlates$Day [i] == 5) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [47]
    } else if(AllPlates$Day [i] == 6) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [48]
    } else if(AllPlates$Day [i] == 7) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [49]
    }
  } else if(AllPlates$Grade [i] == "B"){
      if(AllPlates$Day [i] == 1) {
        AllPlates$HKMean [i] <- MeanDf$Concentration [35]
      } else if(AllPlates$Day [i] == 2) {
        AllPlates$HKMean [i] <- MeanDf$Concentration [36]
      } else if(AllPlates$Day [i] == 3) {
        AllPlates$HKMean [i] <- MeanDf$Concentration [37]
      } else if(AllPlates$Day [i] == 4) {
        AllPlates$HKMean [i] <- MeanDf$Concentration [38]
      } else if(AllPlates$Day [i] == 5) {
        AllPlates$HKMean [i] <- MeanDf$Concentration [39]
      } else if(AllPlates$Day [i] == 6) {
        AllPlates$HKMean [i] <- MeanDf$Concentration [40]
      } else if(AllPlates$Day [i] == 7) {
        AllPlates$HKMean [i] <- MeanDf$Concentration [41]
      } else {
    AllPlates$HKMean [i] <- 0
      }
  }
}
## For Round1 embryos ##
for(i in 1:nrow(AllPlates)){
  if(AllPlates$Grade [i] == "G"){
    if(AllPlates$Day [i] == 0){
      AllPlates$HKMean [i] <- MeanDf$Concentration [40]
    } else if(AllPlates$Day [i] == 1) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [41]
    } else if(AllPlates$Day [i] == 2) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [42]
    } else if(AllPlates$Day [i] == 3) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [43]
    } else if(AllPlates$Day [i] == 4) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [44]
    } else if(AllPlates$Day [i] == 5) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [45]
    } else if(AllPlates$Day [i] == 6) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [46]
    } else if(AllPlates$Day [i] == 7) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [47]
    }
  } else if(AllPlates$Grade [i] == "B"){
    if(AllPlates$Day [i] == 1) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [33]
    } else if(AllPlates$Day [i] == 2) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [34]
    } else if(AllPlates$Day [i] == 3) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [35]
    } else if(AllPlates$Day [i] == 4) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [36]
    } else if(AllPlates$Day [i] == 5) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [37]
    } else if(AllPlates$Day [i] == 6) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [38]
    } else if(AllPlates$Day [i] == 7) {
      AllPlates$HKMean [i] <- MeanDf$Concentration [39]
    } else {
      AllPlates$HKMean [i] <- 0
    }
  }
}
 # column class should be numeric in order to do calculations #
class(AllPlates$HKMean)
AllPlates$HKMean <- as.numeric(AllPlates$HKMean)

# Take Sample Concentration and subtract HKMean (gotten from MeanDf) #
AllPlates <- AllPlates %>% mutate(dCt = Concentration - HKMean)
AllPlates <- AllPlates %>% mutate(Two_dCt = 2^-dCt)

# Get means of dCt sorted by day, grade, target #
dCtMeans <- aggregate(dCt ~ Day + Grade + Target, FUN=mean, data = AllPlates)
View(dCtMeans)

## For loop to add control (Good embryos) dCt means to appropriate row in AllPlates df ##
for(i in 1:nrow(AllPlates)){
  AllDay <- AllPlates$Day [i]
  AllTar <- AllPlates$Target [i]
  for(d in 1:nrow(dCtMeans)){
    dDay <- dCtMeans$Day [d]
    dTar <- dCtMeans$Target [d]
    dGrade <- dCtMeans$Grade [d]
    if(dDay == AllDay){
      if(dTar == AllTar){
        if(dGrade == "G"){
          AllPlates$Good_dCt_Mean [i] <- dCtMeans$dCt [d]
        }
      }
    }
  }
}

## Calculate ddCt and 2-ddCt ##
AllPlates <- AllPlates %>% mutate(ddCt = dCt - Good_dCt_Mean)
AllPlates <- AllPlates %>% mutate(Two_dd_Ct = 2^-ddCt)
AllPlates <- AllPlates %>% mutate(Inv_dCt = HKMean - Concentration)
AllPlates$StatGroup <- paste(AllPlates$Day, AllPlates$Grade, AllPlates$Target, sep = ".")

## Save File ##

XL <- openxlsx::createWorkbook()

Merged_sheet <- addWorksheet(XL, sheetName = "ddPCR_Analysis")
freezePane(XL, sheet = 1, firstRow = TRUE, firstCol = TRUE)
writeDataTable(XL, sheet = 1, AllPlates)

openxlsx::saveWorkbook(XL, file = file.path(folder_path, output_name))
