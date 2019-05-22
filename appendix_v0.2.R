
options(java.parameters = "-Xmx8048m")
memory.limit(size=10000000024)
rm(list=ls())
setwd("D:/Appendix/CJ_DFM_102/3_hee2_CJ_DFM_102_1025_20190515_043700/")

# Install packages 
#install.packages("icesTAF")
#install.packages('dplyr')
#install.packages('lubridate')
#install.packages('reshape')
#install.packages('reshape2')
#install.packages('chron')
#install.packages('magrittr')
#install.packages('rvg', dependencies = T)
#install.packages('R.utils')
#install.packages('knitr')
#install.packages('png')
#install.packages('shiny')
#install.packages('maditr')
#install.packages('rJava')
#install.packages('ReporteRs', dependencies=TRUE)  
#install.packages('ReporteRsjars') https://www.java.com/en/download/manual.jsp
#install.packages("xlsx")
#install.packages("sas7bdat")
#install.packages("haven)
library(sas7bdat)
library(xlsx)
library(maditr)
library(shiny)
library(png)
library(knitr)
library(rJava)
library(R.utils)
library(rvg)
library(ReporteRsjars)
library(ReporteRs)
library(dplyr)
library(lubridate)
library(reshape)
library(reshape2)
library(chron)
library(magrittr)
library(stringr)
library(haven)
library(icesTAF)

# Document Setting Value
options('ReporteRs-fontsize'=9, 'ReporteRs-default-font'='맑은 고딕')
headerprop <- textProperties(font.weight = "bold", vertical.align="middle")
textprop   <- textProperties(vertical.align="middle")
cellprop   <- cellProperties(vertical.align="middle")

# Save CSV Files
mainDir <- getwd()
subDir  <- "SAS2CSV"
if(file.exists(subDir)){
    print("The SAS2CSV files directory already exists")
}else{
    dir.create(file.path(mainDir,subDir))
    sas.files <- list.files()[grepl(".sas7bdat", list.files())]
    for(files_ in sas.files){
        name_ <- toupper(strsplit(files_, ".", fixed = TRUE)[[1]][1])
        outfile <- write.csv(as.data.frame(read_sas(files_), NULL), 
                             file = paste0(file.path(mainDir, subDir), "/", name_, ".csv"), 
                             row.names = FALSE)
    }
}

# Load CSV Files
data.files <- list.files(getwd())

SUBJECT <- as.data.frame(read_sas(data_file=data.files[grepl("\\<subject\\>", data.files)], NULL))
DS0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<ds3\\>", data.files)], NULL))
DM0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<dm\\>", data.files)], NULL))
SU0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<su\\>", data.files)], NULL))
EX0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<ex\\>", data.files)], NULL))
PC0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pc1\\>", data.files)], NULL))
AE0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<ae\\>", data.files)], NULL))
LAB0    <- as.data.frame(read_sas(data_file=data.files[grepl("\\<lab\\>", data.files)], NULL))
LBBC0   <- as.data.frame(read_sas(data_file=data.files[grepl("\\<lbbc\\>", data.files)], NULL))
MH0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<mh\\>", data.files)], NULL))
IE0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<ie\\>", data.files)], NULL))
PE0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<pe\\>", data.files)], NULL))
VS0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<vs\\>", data.files)], NULL))
EG0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<eg\\>", data.files)], NULL))
AEYN0   <- as.data.frame(read_sas(data_file=data.files[grepl("\\<aeyn\\>", data.files)], NULL))



enrollment.fun <- function(){
    # Extract the data from SUBJECT csv file that the code ID does not start with "S"
    #
    # Returns:
    #	Subdata with ID code of A and B from SUBJECT.csv
    
    enrollment <- data.frame(SUBJECT$Subject)
    enrollment <- data.frame(enrollment[-grep("^S", enrollment$SUBJECT.Subject),])
    colnames(enrollment) <- c("Subject")
    enrollment <- enrollment %>%
        arrange(Subject)
    return(enrollment)
}


flex.table.fun <- function(data, head.font=9, body.font=9, calc_=FALSE){
    # Convert data frame into FlexTable format
    #
    # Args:
    #	data: Merged data table
    #	head.font: column header font size (default is 9)
    #	body.font: body content font size  (default is 9)
    #   calc_: whether to calculate the count, mean, sd, etc..
    #
    # Returns:
    #	The flextable format of dataset
    
    headerprop <- textProperties(font.weight = "bold", vertical.align="middle", font.size=head.font)
    bodyprop <- textProperties(vertical.align="middle", font.size=body.font)
    MyFTable <- FlexTable(data, header.columns=TRUE, header.text.props = headerprop, body.text.props=bodyprop)
    MyFTable[to = "header"] = parCenter()
    MyFTable[1, to = "header", side = "bottom"] = borderProperties(style = "solid", width = 3)
    MyFTable[1:nrow(data),1:ncol(data)] = parProperties(text.align = "center")
    
    if(calc_==TRUE){
        count.mean.sd.list <- count.mean.sd.fun(data)
        for(i in 1:length(count.mean.sd.list$list_)){
            MyFTable <- addFooterRow(MyFTable, 
                                     value = c(count.mean.sd.list$names_[i], unlist(count.mean.sd.list$list_[i])),
                                     text.properties=headerprop,
                                     colspan=rep(1,ncol(data)),
                                     parProperties(text.align="center"),
                                     cell.properties=cellProperties(background.color="gray"))
        }
    }
    return (MyFTable)
}

remove.dup.fun <- function(data){
    # Remove the duplicated column by Subject and FolderSeq
    
    dup.idx <- c()
    for(i in 2:nrow(data)){
        if(data[i,]$Subject==data[i-1,]$Subject && data[i,]$FolderSeq==data[i-1,]$FolderSeq){
            dup.idx <- c(dup.idx, i)
        }
    }
    if(!length(dup.idx)==0){
        data <- data[-(dup.idx-1), ]  
    }else{
        return(data)
    }
}

create.table <- function(csvfile, code_id=NULL, fullname=NULL, filter_var=NULL, value_var, period_=NULL){
    # Create the pivot table for infile dataset
    #
    # Args:
    #	infile: the numeric dataset to be used to create the pivot table
    #	code_id: specific code id that is required to filter out from the infile dataset
    #	fullname: the full name of code_id
    #	filter_var: the filter column name
    #	value_var: the variable needed to be put as value
    #	period_: whether period 1 or period 2
    #
    # Returns:
    #	Pivot table for the dataset
    
    # Load the Enrollment file
    enrollment <- enrollment.fun()
    
    if(!is.null(code_id) && !is.null(filter_var)){
        if(!is.null(period_)){
            if(period_=="P1"){
                data <- csvfile %>%
                    mutate(Folder2=substr(Folder, 1, 2)) %>%
                    filter((Folder2==as.character(period_) | Folder2=="SC") & 
                               (csvfile[filter_var]==as.character(code_id))) %>%
                    select(Subject, InstanceRepeatNumber, FolderSeq, value_var) %>%
                    arrange(Subject, InstanceRepeatNumber)
            }
            else if(period_=="P2"){
                data <- csvfile %>%
                    mutate(Folder2=substr(Folder, 1, 2)) %>%
                    filter(Folder2==as.character(period_) &
                               csvfile[filter_var]==as.character(code_id)) %>%
                    select(Subject, InstanceRepeatNumber, FolderSeq, value_var) %>%
                    arrange(Subject, InstanceRepeatNumber)
            }
        }else{
            data <- csvfile %>%
                filter(csvfile[filter_var]==as.character(code_id))%>%
                select(Subject, InstanceRepeatNumber, FolderSeq, value_var) %>%
                arrange(Subject, InstanceRepeatNumber)
        }
    }else{
        data <- csvfile %>%
            select(Subject, InstanceRepeatNumber, FolderSeq, value_var) %>%
            arrange(Subject, InstanceRepeatNumber)
    }
    
    
    # Remove duplicated row by Subject and FolderSeq
    data <- remove.dup.fun(data)
    
    # Create Pivot table by value_var 
    data.pivot <- dcast(data, Subject~FolderSeq, value.var=as.character(value_var))
    
    # Change Column name 
    data.pivot <- change.column.name(data.pivot)
    
    # Map to full name
    if(!is.null(fullname)){
        colnames(data.pivot)[-1] <- paste0(as.character(fullname), "_", colnames(data.pivot)[-1])
    }
    # Merge with enrollment
    data.table <- merge(enrollment, data.pivot, by=c("Subject"))
    
    # data post-processing (Can be consistently added)
    data.table[is.na(data.table)] <- NaN
    data.table[data.table==""]    <- NaN
    data.table[data.table=="01월 03일"] <- "1-3"
    data.table[data.table=="04월 09일"] <- "4-9"
    data.table[data.table=="10월 19일"] <- "10-19"
    
    return (data.table)
}


count.mean.sd.fun <- function(data){
    # Calculate the count, median, mean, min, max ,sd, cv for the given data
    #
    # Args:
    #	data: Numeric data table
    #
    # Returns:
    #	Calculated values will be appended at the end of data table
    
    # doc.table <- flex.table.fun(data)
    
    n    <- c()	# count
    med  <- c()	# median
    min  <- c()	# min
    max  <- c()	# max
    mean <- c()	# mean
    sd   <- c()	# standard deviation
    cv   <- c()	# coefficient variation
    
    for (i in 1:(ncol(data)-1)){
        n[i]    <- length(data[data[,1+i]!="NaN", 2])
        med[i]  <- (median(data[,1+i], na.rm=TRUE))
        min[i]  <- (min(data[,1+i], na.rm=TRUE))
        max[i]  <- (max(data[,1+i], na.rm=TRUE))
        mean[i] <- round(sum(data[,1+i], na.rm=TRUE)/length(data[data[,1+i]!="NaN",2]),2)
        sd[i]   <- round(sd(data[,1+i], na.rm=TRUE),2)
        cv[i]   <- round(sd(data[,1+i], na.rm=TRUE)/
                             (sum(data[,1+i], na.rm=TRUE)/
                                  length(data[data[,1+i]!="NaN",2]))*100,2)
    }
    
    all.list <- list(n, med, min, max, mean, sd, cv)
    all.names <- c("N", "Median", "Min", "Max", "Mean", "SD", "CV(%)")
    
    return(list(list_=all.list, names_=all.names))
}



change.column.name <- function(data){
    # Change the column names according to the given condition 
    # extracted from data dictionary xlsx
    
    dic.file <- list.files("./..")
    data.dic.name <- dic.file[grepl("Data Dictionary", dic.file)]
    if(identical(data.dic.name, character(0))){
        stop("Missing Data Dictionary File")
    }
    data.dic <- read.xlsx(paste("./../", data.dic.name, sep=""), sheetName="Folders")
    Ordinal <- as.character(data.dic$Ordinal)
    FolderName <- as.character(data.dic$FolderName)
    
    FolderName[FolderName=="Adverse Events"] <- "AE"
    FolderName[FolderName=="Screening"] <- "Scr"
    FolderName[FolderName=="Post-Study Visit"] <- "PSV"
    FolderName[FolderName=="Unscheduled Visit"] <- "U/V"
    FolderName[FolderName=="Case Conclusion"] <- "CC"
    FolderName[FolderName=="Concomitant Medications"] <- "CM"
    FolderName[FolderName=="Violation of Protocol"] <- "VP"
    FolderName <- str_replace(FolderName, "Period 1 " ,"")
    FolderName <- str_replace(FolderName, "Period 2 " ,"")
    
    colname_map <- data.frame(Ordinal, FolderName)
    
    colnames(data)[1] <- c("Subject")
    for(i in 1:ncol(data)){
        for(j in 1:nrow(colname_map)){  
            if(colnames(data)[i] == colname_map$Ordinal[j]){
                colnames(data)[i] <- as.character(colname_map$FolderName[j])
            }
        }
    }
    return(data)	
}


# ======== Starting From Here ===========
enrollment <- enrollment.fun()

#11.2.1

DS <- DS0 %>%
    filter(Subject!='' & DSDECOD_STD!='') %>%
    mutate(Treatment ='',
           Final_visit = DSSTDTC_RAW,
           Reason_for_discontinuation = DSDECOD) %>%
    select(Subject, Treatment, Reason_for_discontinuation, Final_visit) %>%
    arrange(Subject)

DM <- DM0 %>%
    mutate(Sex=SEX,
           Age=AGE) %>%
    select(Subject, Sex, Age)

data <- merge(DS, DM, by=c("Subject"="Subject"))
data <- data[,c("Subject", 
             "Treatment", 
             "Sex", 
             "Age", 
             "Final_visit", 
             "Reason_for_discontinuation")]

MyFTable_11.2.1 <- flex.table.fun(data)


# 11.2.4.1

DM <- DM0 %>%
    filter(Subject!='' & substr(Subject, 1, 1) != "S") %>%
    select(Subject, DMDTC_RAW, BRTHDTC_RAW, HEIGHT, WEIGHT, SEX ) %>%
    arrange(Subject)


SU <- SU0 %>%
    filter(Subject!='' & substr(Subject, 1, 1) != "S") %>%
    select(Subject, SUYN1, SUYN2, SUYN3) %>%
    arrange(Subject)


data <- merge(DM, SU, by=c("Subject"))
data <- data[c("Subject", 
               "DMDTC_RAW", 
               "BRTHDTC_RAW", 
               "SUYN1",
               "SUYN2",
               "SUYN3",
               "HEIGHT",
               "WEIGHT",
               "SEX")]

colnames(data)=c("Subject", 
                 "Screening Date",
                 "Birth Date", 
                 "Smoking", 
                 "Alcohol",
                 "Caffeine",
                 "Height", 
                 "Weight", 
                 "Sex")

MyFTable_11.2.4.1 <- flex.table.fun(data)


# 11.2.5.1_1 & 11.2.5.1_2

period_ <- c("P1", "P2")
MyFTable_11.2.5.1_1To11.2.5.1_2 <- list()

for (i in 1:length(period_)){
    
    EX <- EX0 %>%
        mutate(Folder=substr(Folder, 1, 2)) %>%
        filter(Folder==as.character(period_[i])) %>%
        select(Subject, EXSTTIM, FolderName, Folder) %>%
        arrange(Subject)
    
    data <- dcast(EX, Subject~FolderName, value.var="EXSTTIM")
    MyFTable_11.2.5.1_1To11.2.5.1_2[[i]] <- flex.table.fun(data)
}


#11.2.6.4_1  &  11.2.6.4_2

MyFTable_11.2.6.4_1To11.2.6.4_2 <- list()
for (i in 1:length(period_)){
    
    PC <- PC0 %>%
        mutate(Folder=substr(Folder, 1, 2)) %>%
        filter(Folder==as.character(period_[i])) %>%
        mutate(FolderName1=substr(FolderName,10,14)) %>%
        select(Subject, PCTIM, FolderName1, PCTPT_STD) 
    
    data <- dcast(PC, Subject~FolderName1+PCTPT_STD, value.var="PCTIM")
    
    colnames(data)[-1] <- paste0(colnames(data)[-1],"h")
    colnames(data) <- gsub("_-30h", "_투약 전(0 h)", colnames(data))
    
    MyFTable_11.2.6.4_1To11.2.6.4_2[[i]] <- flex.table.fun(data, 7)
}


# 11.2.7

AE <- AE0 %>%
    filter(Subject!='') %>%
    select(Subject, 
           AETERM,
           AESTDTC_RAW,
           AESTTIM,
           AEENDTC_RAW,
           AEENDTIM,
           AESEV,
           AESER,
           AEOUT,
           AEREL,
           AEACNOTH,
           AEACN) %>%
    arrange(Subject)


colnames(AE) <- c("Enrollment No.",
               "이상반응명",
               "시작일",
               "시작시간",
               "종료일",
               "종료시간",
               "중증도",
               "중대성",
               "결과",
               "의약품과의인과관계",
               "생물학적동등성시험용의약품에 대한 조치",
               "이상반응에 대한 조치")

MyFTable_11.2.7 <- flex.table.fun(AE, head.font = 7)

#11.4.2_1.1 ~ 11.4.2_7.4
# Main
name_map <- data.frame(rbind(c("WBC", "WBC"), 
                             c("RBC", "RBC"),
                             c("HGB", "Hemoglobin"),
                             c("HCT", "Hematocrit"),
                             c("PLAT", "Platelets"),
                             c("NEUT", "Seg.neutrophils"),
                             c("LYM", "Lympho."),
                             c("MONO", "Mono."),
                             c("EOS", "Eosino."),
                             c("BASO", "Baso."),
                             c("MCV", "MCV"),
                             c("MCH", "MCH"),
                             c("MCHC", "MCHC"),
                             c("GLUC_C", "Glucose"),
                             c("BUN", "BUN"),
                             c("URATE", "Uric acid"),
                             c("PROT_C", "Total protein"),
                             c("ALB", "Albumin"),
                             c("BILI_C", "Total bilirubin"),
                             c("ALP", "Alk.phosphatase"),
                             c("AST", "AST"),
                             c("ALT", "ALT"),
                             c("GGT", "GGT"),
                             c("LDH", "LDH"),
                             c("CREAT", "Creatinine"),
                             c("SODIUM", "Sodium"),
                             c("K", "Potassium"),
                             c("CL", "Chloride"),
                             c("CA", "Calcium"),
                             c("PH", "Phosphorus"),
                             c("CHOL", "Total cholesterol"),
                             c("EGFR", "MDRD-eGFR"))
)
colnames(name_map) <- c("AnalyteName", "Fullname")

MyFTable_11.4.2_1.1To11.4.2_7.4 <- list()
filter_var <- "AnalyteName"
value_var <- "NumericValue"
subject_code <- c("A", "B")

for(i in 1:length(name_map$AnalyteName)){
    for(k in 1:length(subject_code)){
        MyFTable_11.4.2_1.1To11.4.2_7.4[[2*(i-1)+k]] <- flex.table.fun(
            data = create.table(csvfile    = LAB0[substr(LAB0$Subject,1,1)==subject_code[k],],
                                code_id    = name_map$AnalyteName[i], 
                                fullname   = name_map$Fullname[i],
                                filter_var = filter_var,
                                value_var  = value_var),
            calc_ = TRUE)
        
        print(paste("Creating Table for", name_map$Fullname[i], "on subject code", subject_code[k]))
    }
}


#11.4.2.3_1.1 ~ 11.4.2.3_5.2
# Main 
name_map <- data.frame(rbind(c("GLUC_U", "Glucose"), 
                             c("COLOR", "Color"),
                             c("LEUKASE", "Leukocyte"),
                             c("BILI_U", "Bilirubin"),
                             c("KETONEBD", "Ketone"),
                             c("SPGRAV", "Specific gravity"),
                             c("OCCBLD", "Occult blood"),
                             c("PH", "PH"),
                             c("PROT_U", "Protein"),
                             c("UROBIL", "Urobilinogen"),
                             c("NITRITE", "Nitrite"),
                             c("WBC count(UF)", "Microscopy WBC"),
                             c("RBC count(UF)", "Microscopy RBC"))
)
colnames(name_map) <- c("AnalyteName", "Fullname")

MyFTable_11.4.2.3_1.1To11.4.2.3_5.2 <- list()
filter_var <- "AnalyteName"
value_var <- "AnalyteValue"
for(i in 1:length(name_map$AnalyteName)){
    MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[i]] <- flex.table.fun(
        data = create.table(csvfile    = LAB0, 
                            code_id    = name_map$AnalyteName[i], 
                            fullname   = name_map$Fullname[i],
                            filter_var = filter_var,
                            value_var  = value_var),
        calc_ = FALSE)
    
    print(paste("Creating Table for", name_map$Fullname[i]))
}


#11.4.2.2_4

MyFTable_11.4.2.4 <- list()
for(j in 1:length(subject_code)){
    LBBC <- LBBC0 %>%
        filter(substr(LBBC0$Subject,1,1)==subject_code[j]) %>%
        select(Subject, BCPT, BCAPTT) %>%
        arrange(Subject)
    
    colnames(LBBC)=c("Subject", "PT(INR)", "aPTT")
    
    MyFTable_11.4.2.4[[j]] <- flex.table.fun(LBBC, calc_ = TRUE)
}

# 11.4.2.5

LAB <- LAB0 %>%
    filter(AnalyteName == "HBSAG" | 
               AnalyteName == "HCVCLD" | 
               AnalyteName == "HIVAB" | 
               AnalyteName == "Syphilis reagin test") %>%
    select(Subject, AnalyteName, AnalyteValue) %>%
    arrange(Subject)

AnalyteName.pivot <- dcast(LAB, Subject ~ AnalyteName, value.var = "AnalyteValue")

colnames(AnalyteName.pivot)[2] <- "HBsAg"
colnames(AnalyteName.pivot)[3] <- "Anti-HCVAb"
colnames(AnalyteName.pivot)[4] <- "HIV Ag/Ab"

AnalyteName.pivot <- merge(enrollment, AnalyteName.pivot, by="Subject")
MyFTable_11.4.2.5 <- flex.table.fun(AnalyteName.pivot)


#11.4.3

MH <- MH0 %>%
    filter(Folder=="SCRN") %>%
    select(Subject, MHBODYSYS_STD) %>%
    arrange(Subject)

MH[is.na(MH)] <- 0
MH <- MH[!duplicated(MH[c("Subject","MHBODYSYS_STD")]),]


index.fun <- function(data){
    cnt <- 0
    fix.idx <- 0
    index.data <- data.frame(matrix("No", ncol = 15, nrow = nrow(data)))
    index.data[, ] <- lapply(index.data[, ], as.character)
    
    for(i in 1:nrow(data)){
        fix.idx <- i - cnt
        for(k in 1:15){
            if(data[i,2]==k && data[i,1]==data[i+1,1]){
                cnt <- cnt + 1
                index.data[fix.idx, k] <- "Yes"
                
            }else if(data[i,2]==k && data[i,1]!=data[i+1,1]){
                index.data[fix.idx,k] <- 'Yes'
            }
        }
    }
    index.data <- index.data[1:fix.idx,]
    MH.subject <- data[!duplicated(data[c("Subject")]),]$Subject
    return(cbind(MH.subject,index.data))
}

MH.index <- index.fun(MH)
colnames(MH.index)=c("Enrollment No.",
                     "Cardiovascular", 
                     "Perivascular", 
                     "Skin / mucosa", 
                     "Eye", 
                     "Ear / Nose & Throat", 
                     "Respiratory", 
                     "Musculoskeletal",
                     "Infectious diseases", 
                     "Gastrointestinal", 
                     "Endocrine", 
                     "Renal / Genitourinary", 
                     "Neurologic / Psychiatric", 
                     "Oncology", 
                     "Fracture", 
                     "Operation Hx.")

MH.index <- left_join(enrollment, MH.index, by=c("Subject"="Enrollment No."))
MyFTable_11.4.3 <- flex.table.fun(MH.index,head.font = 7, body.font = 7)


# 11.4.4

IETEST_STD.pivot <- dcast(IE0, Subject ~ IETEST_STD, value.var = "IEORRES")
IETEST_STD.pivot[IETEST_STD.pivot=="예"] <- "Yes"
IETEST_STD.pivot[IETEST_STD.pivot=="아니오"] <- "No"
IETEST_STD.pivot <- merge(enrollment, IETEST_STD.pivot, by=c("Subject"))

IN.idx <- which(grepl("^IN", colnames(IETEST_STD.pivot)))
EX.idx <- c(2:(IN.idx[1]-1))

IETEST_STD.pivot <- IETEST_STD.pivot[,c(1, IN.idx, EX.idx, length(IETEST_STD.pivot))]

MyFTable_11.4.4 <- flex.table.fun(IETEST_STD.pivot, 7, 6) 



# 11.4.5_1

PE <- PE0 %>%
    filter(substr(PE0$Folder, 1, 2)=="P1" | substr(PE0$Folder, 1, 2)=="SC") %>%
    select(Subject, FolderSeq, PEYN_STD) %>%
    arrange(Subject)

FolderSeq.pivot <- dcast(PE, Subject ~ FolderSeq, value.var="PEYN_STD")
FolderSeq.pivot <- change.column.name(FolderSeq.pivot)

result <- merge(enrollment, FolderSeq.pivot, by=c("Subject"="Subject"))
result[is.na(result)] <- "N/A"
result[result==0] <- "N/A"
result[result==1] <- "Normal"
result[result==2] <- "Abnormal"

MyFTable_11.4.5_1 <- flex.table.fun(result, 8)


#11.4.5_2

PE <- PE0 %>%
    filter(substr(PE0$Folder, 1, 2)=="P2") %>%
    select(Subject, FolderSeq, PEYN_STD) %>%
    arrange(Subject)

FolderSeq.pivot <- dcast(PE, Subject ~ FolderSeq, value.var="PEYN_STD")
FolderSeq.pivot <- change.column.name(FolderSeq.pivot)

result <- merge(enrollment, FolderSeq.pivot, by=c("Subject"="Subject"))
result[is.na(result)] <- "N/A"
result[result==0] <- "N/A"
result[result==1] <- "Normal"
result[result==2] <- "Abnormal"

MyFTable_11.4.5_2 <- flex.table.fun(result, 8)


# 11.4.6.1_1 ~ 11.4.6.4_2
# Main
vstest_std <- unique(VS0$VSTEST_STD[VS0$VSTEST_STD!=""])
period_ <- c("P1", "P2")
filter_var <- "VSTEST_STD"
value_var <- "VSORRES"
MyFTable_11.4.6.1_1To11.4.6.4_2 <- list()

for (i in 1:length(vstest_std)){
    for(j in 1:length(period_)){
        for(k in 1:length(subject_code)){
            MyFTable_11.4.6.1_1To11.4.6.4_2[[4*(i-1)+2*(j-1)+k]] <- flex.table.fun(
                data = create.table(csvfile    = VS0[substr(VS0$Subject,1,1)==subject_code[k],], 
                                    code_id    = vstest_std[i],
                                    filter_var = filter_var,
                                    value_var  = value_var,
                                    period_    = period_[j]),
                calc_ = TRUE)
            
            print(paste("Creating Table for",vstest_std[i], "on subject code", subject_code[k], "at period", period_[j]))
        }
    }
}

# 11.4.7_1 ~ 11.4.7_5
# Main
Fullname <- unique(EG0$EGTEST[EG0$EGTEST!=""])
AnalyteName <- unique(EG0$EGTEST_STD[EG0$EGTEST_STD!=""])
name_map <- data.frame(AnalyteName,Fullname)

filter_var <- "EGTEST_STD"
value_var <- "EGORRES"
MyFTable_11.4.7_1To11.4.7_5 <- list()
for (i in 1:length(name_map$AnalyteName)){
    for(k in 1:length(subject_code)){
        MyFTable_11.4.7_1To11.4.7_5[[2*(i-1)+k]] <- flex.table.fun(
            data = create.table(csvfile    = EG0[substr(EG0$Subject,1,1)==subject_code[k],], 
                                code_id    = name_map$AnalyteName[i],
                                fullname   = name_map$Fullname[i],
                                filter_var = filter_var,
                                value_var  = value_var),
            calc_ = TRUE) 
        
        print(paste("Creating Table for",name_map$Fullname[i], "on subject code", subject_code[k]))
    }
}

# 11.4.7_6

EG <- EG0 %>%
    filter(EGTEST_STD=="VR") %>%
    select(Subject, FolderSeq, EGCLSIG) %>%
    arrange(Subject)

EG.pivot <- dcast(EG, Subject ~ FolderSeq, value.var="EGCLSIG")
EG <- change.column.name(EG.pivot)

data <- merge(enrollment, EG, by=c("Subject"="Subject"))
colnames(data)[-1] <- paste0("Overall_", colnames(data)[-1])

data[is.na(data)] <- "NaN"
data[data=="비정상/NCS"] <- "NCS"
data[data =="정상"] <- "Normal"

MyFTable_11.4.7_6 <- flex.table.fun(data)


# BTS (Extra)
if(file.exists(list.files()[which(startsWith(list.files(), "bst"))])){
    BST0     <- as.data.frame(read_sas(data_file=data.files[grepl("\\<bst\\>", data.files)], NULL))
    MyFTable_11.4.8 <- flex.table.fun(
        create.table(csvfile    = BST0,
                     code_id    = NULL,
                     fullname   = NULL,
                     filter_var = NULL,
                     value_var  = "BSTORRES_RAW",
                     period_    = NULL),
        calc_ = FALSE
    )
}


##############################
###       Doc & Table      ###
##############################

doc <- docx(template="D:/Appendix/Template.docx")

doc <- addTitle(doc, "Format cell text values", level=1) #1
doc <- addTitle(doc, "Format cell text values", level=1) #2
doc <- addTitle(doc, "Format cell text values", level=1) #3
doc <- addTitle(doc, "Format cell text values", level=1) #4
doc <- addTitle(doc, "Format cell text values", level=1) #5
doc <- addTitle(doc, "Format cell text values", level=1) #6
doc <- addTitle(doc, "Format cell text values", level=1) #7
doc <- addTitle(doc, "Format cell text values", level=1) #8
doc <- addTitle(doc, "Format cell text values", level=1) #9
doc <- addTitle(doc, "Format cell text values", level=1) #10
doc <- addTitle(doc, "Format cell text values", level=1) #11
doc <- addTitle(doc, "STUDY INFORMATION", level=2)	   #11.1
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Protocol and Protocol Amendments", level=3) # 11.1.1
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Sample Case Report Form", level=3)	#11.1.2
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "List of IRBs - Representative written information for patient and sample consent forms", level=3)
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "List and Description of Investigators and Other Important Participants in the Study, Including Brief (one page) CV's or equivalent summaries of training and experience relevant to the performance of the clinical study", level=3)
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Signatures of Principal or Coordinating Investigator(s) or Sponsor’s Responsible medical officer, Depending on the Regulatory Authority's Requirement", level=3)
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Listing of Patients Receiving Test Drug(s)/Investigational Product from Specific Batches, Where more than one batch was used", level=3)
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Randomization Scheme and Codes (Patient Identification and Treatment Assigned)", level=3)
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Audit Certificates", level=3)
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Documentation of Statistical Methods", level=3)
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Documentation of Inter-laboratory Standardization Methods and Quality Assurance Procedures If Used", level=3)
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Publications Based On the Study", level=3)
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Important Publications Referenced In the Report", level=3)	 # 11.1.12
doc <- addParagraph(doc, "Not Applicable", stylename="rTableLegend",offx=100,offy=100)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "SUBJECT DATA LISTINGS", level=2)		# 11.2


doc <- addTitle(doc, "Discontinued patients", level=3)	#11.2.1
doc <- addFlexTable(doc, MyFTable_11.2.1)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Protocol deviations", level=3)
doc <- addPageBreak(doc)
doc <- addTitle(doc, "Patients excluded from the efficacy analysis", level=3)
doc <- addPageBreak(doc)

doc <- addTitle(doc, "Demographic data", level=3)
doc <- addParagraph(doc, "11.2.4.1 DEMOGRAPHIC CHARACTERISTICS", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.2.4.1)
doc <- addPageBreak(doc)

#doc <- addSection(doc, landscape = TRUE) # 가로 시작
doc <- addTitle(doc, "Compliance and/or drug concentration data (if available)", level=3)
doc <- addParagraph(doc, "11.2.5.1 DRUG ADMINISTRATION; PERIOD 1", stylename="rTableLegend",offx=100,offy=100)
doc <- addFlexTable(doc, MyFTable_11.2.5.1_1To11.2.5.1_2[[1]])
doc <- addPageBreak(doc)

doc <- addParagraph(doc, "11.2.5.2 DRUG ADMINISTRATION; PERIOD 2", stylename="rTableLegend",offx=100,offy=100)
doc <- addFlexTable(doc, MyFTable_11.2.5.1_1To11.2.5.1_2[[2]])
doc <- addPageBreak(doc)

doc <- addSection(doc, landscape = TRUE) # 가로 시작
doc <- addTitle(doc, "Individual Pharmacokinetic data", level=3) 	#11.2.6
doc <- addParagraph(doc, "11.2.6.4 Actual PK Sampling Time (1)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.2.6.4_1To11.2.6.4_2[[1]])
doc <- addPageBreak(doc)


doc <- addParagraph(doc, "11.2.6.4 Actual PK Sampling Time (2)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.2.6.4_1To11.2.6.4_2[[2]])
doc <- addPageBreak(doc)


doc <- addParagraph(doc, "11.2.7 ADVERSE EVENT LISTINGS (EACH PATIENT)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.2.7)
doc <- addPageBreak(doc)

doc <- addTitle(doc, "CASE REPORT FORMS", level = 2)
doc <- addTitle(doc, "CRFs for deaths, other serious adverse events and withdrawals for AE", level=3)
doc <- addTitle(doc, "OTHER CRFs submitted", level=3)
doc <- addTitle(doc, "INDIVIDUAL SUBJECT DATA LISTINGS", level=2) # 11.4
doc <- addTitle(doc, "Preceded or Concomitant medication", level=3) # 11.4.1
doc <- addPageBreak(doc)
doc <- addSection(doc) #끝


#doc <- addSection(doc, landscape = FALSE) #시작 세로
doc <- addTitle(doc, "Clinical laboratory results", level=3) 	#11.4.2
doc <- addParagraph(doc, "11.4.2.1 CLINICAL LABORATORY-HEMATOLOGY;", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[1]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[2]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[3]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[4]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[5]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[6]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[7]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[8]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.1 CLINICAL LABORATORY-HEMATOLOGY;(CONTINUED)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[9]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[10]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[11]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[12]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[13]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[14]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[15]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[16]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.1 CLINICAL LABORATORY-HEMATOLOGY;(CONTINUED)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[17]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[18]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[19]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[20]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[21]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[22]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[23]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[24]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[25]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.2 CLINICAL LABORATORY-CHEMISTRY;", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[26]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[27]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[28]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[29]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[30]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[31]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[32]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[33]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[34]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[35]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[36]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.2 CLINICAL LABORATORY-CHEMISTRY;(CONTINUED)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[37]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[38]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[39]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[40]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[41]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[42]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[43]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[44]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[45]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[46]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.2 CLINICAL LABORATORY-CHEMISTRY; (CONTINUED)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[47]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[48]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[49]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[50]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[51]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[52]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[53]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[54]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[55]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[56]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.2 CLINICAL LABORATORY-CHEMISTRY; (CONTINUED)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[57]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[58]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[59]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[60]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[61]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[62]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[63]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2_1.1To11.4.2_7.4[[64]])
doc <- addPageBreak(doc)

doc <- addParagraph(doc, "11.4.2.3 CLINICAL LABORATORY-URINALAYSIS",stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[1]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[2]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[3]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.3 CLINICAL LABORATORY-URINALYSIS; (CONTINUED)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[4]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[5]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[6]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.3 CLINICAL LABORATORY-URINALYSIS; (CONTINUED)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[7]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[8]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[9]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.3 CLINICAL LABORATORY-URINALYSIS; (CONTINUED)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[10]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[11]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.3 CLINICAL LABORATORY-URINALYSIS; (CONTINUED)", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[12]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2.3_1.1To11.4.2.3_5.2[[13]])
doc <- addPageBreak(doc)

doc <- addParagraph(doc, "11.4.2.4 CLINICAL LABORATORY-COAGULATION AT SCREENING;", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2.4[[1]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.2.4[[2]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.2.5 CLINICAL LABORATORY-SEROLOGY AT SCREENING;", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.2.5)
doc <- addPageBreak(doc)


doc <- addSection(doc, landscape = TRUE) # 시작
doc <- addTitle(doc, "Medical history", level=3) 	#11.4.3
doc <- addFlexTable(doc, MyFTable_11.4.3)
doc <- addPageBreak(doc)

doc <- addTitle(doc, "Inclusion/Exclusion Criteria", level=3) 	#11.4.4
doc <- addFlexTable(doc, MyFTable_11.4.4)
doc <- addPageBreak(doc)
doc <- addSection(doc) #끝

doc <- addTitle(doc, "Physical examination", level=3) 	#11.4.5
doc <- addParagraph(doc, "11.4.5.1 PHYSICAL EXAMINATION; PERIOD 1", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.5_1)
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.5.2 PHYSICAL EXAMINATION; PERIOD 2", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.5_2)
doc <- addPageBreak(doc)


doc <- addTitle(doc, "Vital Signs", level=3)
doc <- addParagraph(doc, "11.4.6.1 VITAL SIGNS[SBP]; PERIOD 1", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[1]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[2]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.6.1 VITAL SIGNS[SBP]; PERIOD 2", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[3]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[4]])
doc <- addPageBreak(doc)

doc <- addParagraph(doc, "11.4.6.2 VITAL SIGNS[DBP]; PERIOD 1", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[5]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[6]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.6.2 VITAL SIGNS[DBP]; PERIOD 2", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[7]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[8]])
doc <- addPageBreak(doc)

doc <- addParagraph(doc, "11.4.6.3 VITAL SIGNS[PULSE]; PERIOD 1", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[9]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[10]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.6.3 VITAL SIGNS[PULSE]; PERIOD 2", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[11]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[12]])
doc <- addPageBreak(doc)

doc <- addParagraph(doc, "11.4.6.4 VITAL SIGNS[BODY TEMPERATURE]; PERIOD 1", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[13]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[14]])
doc <- addPageBreak(doc)
doc <- addParagraph(doc, "11.4.6.4 VITAL SIGNS[BODY TEMPERATURE]; PERIOD 2", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[15]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.6.1_1To11.4.6.4_2[[16]])
doc <- addPageBreak(doc)


doc <- addTitle(doc, "12-Lead Electrocardiography", level=3)
doc <- addParagraph(doc, "11.4.7 12-LEAD ELECTROCARDIOGRAPHY", stylename="rTableLegend")
doc <- addFlexTable(doc, MyFTable_11.4.7_1To11.4.7_5[[1]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.7_1To11.4.7_5[[2]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.7_1To11.4.7_5[[3]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.7_1To11.4.7_5[[4]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.7_1To11.4.7_5[[5]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.7_1To11.4.7_5[[6]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.7_1To11.4.7_5[[7]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.7_1To11.4.7_5[[8]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.7_1To11.4.7_5[[9]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.7_1To11.4.7_5[[10]])
doc <- addPageBreak(doc)
doc <- addFlexTable(doc, MyFTable_11.4.7_6)
doc <- addPageBreak(doc)


is.defined <- function(obj_){
    # Check if the object exists or not
    # Return: TRUE if exists, FALSE if not
    
    obj_ <- deparse(substitute(obj_))
    env <- parent.frame()
    exists(obj_, env)
}

if(is.defined(MyFTable_11.4.8)){
    doc <- addTitle(doc, "BST", level=3) # 11.4.9
    doc <- addFlexTable(doc, MyFTable_11.4.8)
    doc <- addPageBreak(doc)
    doc <- addTitle(doc, "Visit Dates", level=3) # 11.4.9
}else{
    doc <- addTitle(doc, "Visit Dates", level=3) # 11.4.9
}


appendix.name <- strsplit(getwd(), "/")[[1]][3]
writeDoc(doc, file = paste0(dirname(getwd()), "/", appendix.name, "_APPENDIX.docx"))

