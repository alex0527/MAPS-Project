 ##the script is used for processing the plasma data from Metablone, it will check new analyte, check EDTA zscore and stability, find outlier, generate table 2 and table 3, and even upload files
library("xlsx")

library("dplyr",warn.conflicts = FALSE)
ptm <- proc.time()
##extract all CSV filenames into one list
filenames <- list.files("H:/Biochem/MAPS Project/plasma/CLIA raw file/CLIA_RAW_FILES")

##extract the patient numbers from file names into a new list
patientnumberlist = lapply(filenames,function(x) {m <- gregexpr('[0-9]+',x)
regmatches(x,m)
})

##convert the list to vector
patientnumber <- unlist(patientnumberlist)

##use a to make sure check new analyte function just run once
a <- 1


##for loop will go through all data file and copy zscore, raw, anchor to master lists
for(i in patientnumber) {
  
##use Patient.number to process the data
Patient.number = toString(i)

## open the patient CLIA raw file from Metabolon and raed the CSV file into pfdata
pfName <- paste("H:/Biochem/MAPS Project/plasma/CLIA raw file/CLIA_RAW_FILES/ID_",Patient.number," Z Score CLIA.CSV",sep="")
pfdata <- read.csv(file=pfName, header=TRUE,sep=",",stringsAsFactors=FALSE)

## delete top several rows and blank rows
pfdata = subset(pfdata,SUPER_PATHWAY!="")
pfdata = pfdata[-1,]

## if row score is NA, replace zscore as NA
for(j in nrow(pfdata)) {
  if(toString(pfdata[j,3]) == "NA")
    pfdata[j,5] = "NA"
}

## write the data into a new csv file 
PatientXName <- paste("X",Patient.number,sep="")
NewpfName <- paste("H:/Biochem/MAPS Project/plasma/CLIA raw file/Xpatientnumber/",PatientXName,".csv",sep="")
write.csv(pfdata, file=NewpfName, row.names=FALSE) 

##while function to check new analyte, if new analyte is found, new analyte will be added to database of all new analytes and masterlists of raw, zscore, and anchor
while(a==1) {
  ##check new analyte, read all analyte file first
  allanalyte <- read.csv(file="H:/Biochem/MAPS Project/plasma/all analyte/all_analyte.csv",header=TRUE,sep=",",stringsAsFactors=FALSE)
  
  newpfnameonce <- paste("H:/Biochem/MAPS Project/plasma/CLIA raw file/Xpatientnumber/",PatientXName,".csv",sep="")
  pfdataonce <- read.csv(file=newpfnameonce,header=TRUE,sep=",")
  
  
  ##find the analytes in patient file but not in all_analyte file
  diff <- subset(pfdataonce, !CHEMICAL_ID %in% allanalyte$CHEMICAL_ID)
 
  if(nrow(diff)>0) {
  ##append new analytes to all_analyte 
  write.table(diff,"H:/Biochem/MAPS Project/plasma/all analyte/all_analyte.csv",sep=",",col.names=FALSE,append=TRUE,row.names=FALSE)
  
  allanalytecurrentdate <- read.csv(file="H:/Biochem/MAPS Project/plasma/all analyte/all_analyte.csv",header=TRUE,sep=",",stringsAsFactors=FALSE)
  currentDate <- Sys.Date()
  allanalytecurrentdatename <- paste("H:/Biochem/MAPS Project/plasma/all analyte with current date/all analyte_",currentDate,".csv",sep="") 
  write.csv(allanalytecurrentdate, file=allanalytecurrentdatename, row.names=FALSE) 
  
  ##add new analytes to Master lists of anchor, raw score, and zscore
  write.table(diff,"H:/Biochem/MAPS Project/plasma/Patient_Master_list_Anchor.csv",sep=",",col.names=FALSE,append=TRUE,row.names=FALSE)
  write.table(diff,"H:/Biochem/MAPS Project/plasma/Patient_Master_list_raw.csv",sep=",",col.names=FALSE,append=TRUE,row.names=FALSE)
  write.table(diff,"H:/Biochem/MAPS Project/plasma/Patient_Master_list_zscore.csv",sep=",",col.names=FALSE,append=TRUE,row.names=FALSE)
  }
  a=a-1
 print("new analyte checked")
 print(paste(nrow(diff)," new analyte(s) were found",sep=""))
}


##read the master list of zscore
ml1<- read.csv(file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_zscore.csv",header=TRUE,sep=",")

##re-read the Xpatientnumber file
pfdata1 <- read.csv(file=NewpfName, header=TRUE,sep=",")
##identify matching categories between ml1 and pfdata and paste zcore into ml1 sheet
ml1$newpcol1 <- pfdata1$zscore[match(ml1$CHEMICAL_ID,pfdata1$CHEMICAL_ID)]

##rename column after patient number
names(ml1)[names(ml1)=='newpcol1'] <- PatientXName

##two ouputs.  first output, writes over z-score master list
ml1b <- ml1
write.csv(ml1, file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_zscore.csv", row.names=FALSE) 

###2nd output, appends current date to .csv filename and save as new dated master list
currentDate <- Sys.Date() 
csvFileName <- paste("H:/Biochem/MAPS Project/plasma/zscore with current date/Patient_Master_List_zscore",currentDate,".csv",sep="") 
write.csv(ml1b, file=csvFileName, row.names=FALSE) 


####repeat above but make Raw files
ml2<- read.csv(file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_raw.csv",header=TRUE,sep=",")
ml2$newpcol2<-pfdata1$Raw[match(ml2$CHEMICAL_ID, pfdata1$CHEMICAL_ID)]
names(ml2)[names(ml2) == 'newpcol2'] <- PatientXName
ml2b<-ml2
write.csv(ml2, file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_raw.csv", row.names=FALSE) 
currentDate <- Sys.Date() 
csvFileName2 <- paste("H:/Biochem/MAPS Project/plasma/raw score with current date/Patient_Master_List_raw",currentDate,".csv",sep="") 
write.csv(ml2b, file=csvFileName2, row.names=FALSE) 

####repeat above but make Anchor files
ml3<- read.csv(file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_Anchor.csv",header=TRUE,sep=",")
ml3$newpcol3<-pfdata1$Anchor[match(ml3$CHEMICAL_ID, pfdata1$CHEMICAL_ID)]
names(ml3)[names(ml3) == 'newpcol3'] <- PatientXName
ml3b<-ml3
write.csv(ml3, file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_Anchor.csv", row.names=FALSE) 
currentDate <- Sys.Date() 
csvFileName3 <- paste("H:/Biochem/MAPS Project/plasma/anchor with current date/Patient_Master_List_Anchor",currentDate,".csv",sep="") 
write.csv(ml3b, file=csvFileName3, row.names=FALSE)

print(paste(i," data was uploaded to the database",sep=""))
}


##exclude batch Outliers 

## read the master list of zscore
mlzscore <- read.csv(file="H:/Biochem/MAPS Project/plasma/Patient_Master_List_zscore.csv",header=TRUE,sep=",")

##get all column numbers in zscore master list
nl = ncol(mlzscore)

##get all row numbers in zscore master list
nr = nrow(mlzscore)

##get the number of patient files
patientquantity = length(patientnumber)

##just keep the first, second and patient zscore columns
currentbatchcolumns <- mlzscore[,c(1:2,(nl-patientquantity+1):nl)]

##if zscore of heme >= 3, a warning message will be print out
HEMEzscore <- currentbatchcolumns

for(i in nl-patientquantity+1:nl) {
  for(j in 1:nr) {
    if(!is.null(HEMEzscore[j,1]) && !is.null(HEMEzscore[j,i]) && HEMEzscore[j,1]=="heme" && HEMEzscore[j,i] >= 3)
      print(paste("zscore of heme in patient #",patientquantity+i-nl," was >=3, please review the data", sep=""))
  }
}

print("zscore of heme checked")
##copy patient zsore to EDTAzscore and will be used in generating EDTA table
EDTAzscore <- currentbatchcolumns

write.csv(EDTAzscore,file="H:/Biochem/MAPS Project/plasma/EDTA table/patientinput.csv",row.names=FALSE)

patientvector <- character(0)
mishandlingvector <- character(0)
EDTAzscorevector <- character(0)
absentEDTAvector <- character(0)

b<-3

for(i in patientnumber) {
  
  patientXName5 <- paste("X",i,sep="")
  
  analyteincrease <- read.csv(file="H:/Biochem/MAPS Project/plasma/EDTA table/background analytes/analytes increase.csv",header=TRUE,sep=",",stringsAsFactors = FALSE)
  analytedecrease <- read.csv(file="H:/Biochem/MAPS Project/plasma/EDTA table/background analytes/analytes decrease.csv",header=TRUE,sep=",",stringsAsFactors = FALSE)
  
  for(j in 1:nrow(analyteincrease)) {
    analyteincrease$newcol11[j] <- EDTAzscore[match(analyteincrease[j,1],EDTAzscore[,1]),b]
  }
  names(analyteincrease)[names(analyteincrease)=='newcol11'] <- patientXName5
  write.csv(analyteincrease,file="H:/Biochem/MAPS Project/plasma/EDTA table/background analytes/analytes increase.csv",row.names=FALSE)
  
  for(k in 1:nrow(analytedecrease)) {
    analytedecrease$newcol12[k] <- EDTAzscore[match(analytedecrease[k,1],EDTAzscore[,1]),b]
  }
  names(analytedecrease)[names(analytedecrease)=='newcol12'] <- patientXName5
  write.csv(analytedecrease,file="H:/Biochem/MAPS Project/plasma/EDTA table/background analytes/analytes decrease.csv",row.names=FALSE)
  
  patientvector <- c(patientvector,patientXName5)
  EDTAzscorevector <- c(EDTAzscorevector,EDTAzscore[match("EDTA",EDTAzscore[,1]),b])
  
  b <- b+1
}

analyteincrease1 <- read.csv(file="H:/Biochem/MAPS Project/plasma/EDTA table/background analytes/analytes increase.csv",header=TRUE,sep=",")
analytedecrease1 <- read.csv(file="H:/Biochem/MAPS Project/plasma/EDTA table/background analytes/analytes decrease.csv",header=TRUE,sep=",")

c <- 4
for(i in 1:patientquantity) {
  mishandling <- (sum(analyteincrease1[,c]>1.5,na.rm=TRUE)+sum(analytedecrease1[,c]<(-1.5),na.rm=TRUE))/(sum(!is.na(analyteincrease1[,c]))+sum(!is.na(analytedecrease1[,c])))
  mishandlingvector <- c(mishandlingvector,paste(mishandling*100,"%",sep=""))
  c <- c+1
}

output <- data.frame(patientvector,mishandlingvector,EDTAzscorevector)
colnames(output) <- c("Patient ID","% of stability biomarkers that are consistent with mishandling","EDTA z-score")


write.csv(output,file="H:/Biochem/MAPS Project/plasma/EDTA table/output.csv",row.names=FALSE)

##write EDTA table to excel in batchreport file
currentdate <- Sys.Date()
batchreportname <- paste("H:/Biochem/MAPS Project/plasma/batchreport/plasma_batchreport_",currentdate,".xlsx",sep="")
write.xlsx(output,file=batchreportname,sheetName="EDTA",col.names=TRUE, row.names=FALSE)


##find EDTA mishanding
for(i in 1:nrow(output)) {
  if(as.numeric(sub("%","",output[i,2])) > 45 | output[i,3]=="NA" | as.numeric(output[i,3]) < (-5)) {
    print(paste("absent EDTA or mishandling was found in patient#",output[i,1]," this patient data was removed from the database",sep=""))
    patientquantity <- patientquantity-1
    patientNoXnumber <- sub("X","",output[i,1])
    patientnumber <- patientnumber[!patientnumber==patientNoXnumber]
    absentEDTAvector <- c(absentEDTAvector,toString(output[i,1]))
  }
}

if(length(absentEDTAvector)>0) {
  absentEDTAzscore <- read.csv(file="H:/Biochem/MAPS Project/plasma/Patient_Master_List_zscore.csv",header=TRUE,sep=",")
  absentEDTAraw <- read.csv(file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_raw.csv",header=TRUE,sep=",")
  absentEDTAanchor <- read.csv(file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_Anchor.csv",header=TRUE,sep=",")
  
  absentEDTAzscore <- absentEDTAzscore[,-which(names(absentEDTAzscore) %in% absentEDTAvector)]
  absentEDTAraw <- absentEDTAraw[,-which(names(absentEDTAraw) %in% absentEDTAvector)]
  absentEDTAanchor <- absentEDTAanchor[,-which(names(absentEDTAanchor) %in% absentEDTAvector)]
  
  write.csv(absentEDTAzscore,file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_zscore.csv",row.names = FALSE)
  write.csv(absentEDTAraw,file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_raw.csv",row.names = FALSE)
  write.csv(absentEDTAanchor,file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_Anchor.csv",row.names = FALSE)
}

print("EDTA zscore and stability checked")

mlzscore <- read.csv(file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_zscore.csv",header=TRUE,sep=",",stringsAsFactors=FALSE)
nl <- ncol(mlzscore)
##just patient zscore data
justzscore <- mlzscore[,(nl-patientquantity+1):nl]

biggerthan <- paste("bigger","than",sep="")
lessthan <- paste("less","than",sep="")

currentbatchcolumns$newcol4 <- rowSums(justzscore>2.99,na.rm=TRUE)
names(currentbatchcolumns)[names(currentbatchcolumns)=='newcol4'] <- biggerthan

currentbatchcolumns$newcol5 <- rowSums(justzscore<(-2.99),na.rm=TRUE)
names(currentbatchcolumns)[names(currentbatchcolumns)=='newcol5'] <- lessthan


##write the data to outlier file and will be used for outlier analysis
write.csv(currentbatchcolumns,file = "H:/Biochem/MAPS Project/plasma/outlier/currentbatchzscore.csv",row.names=FALSE)

##use criteria to identify analytes that are seen in more than 25% of the samples
criteria <- patientquantity/4

## read outlier raw data
outlierdata <- read.csv(file="H:/Biochem/MAPS Project/plasma/outlier/currentbatchzscore.csv",header=TRUE,sep=",")

##subset the group of count>2.99
positivecriteria <- subset(outlierdata, biggerthan > criteria)

##subset the group of count<-2.99
negativecriteria <- subset(outlierdata, lessthan > criteria)

##copy the two group in one csv file
write.csv(positivecriteria,file="H:/Biochem/MAPS Project/plasma/outlier/outlier.csv",row.names=FALSE)
write.table(negativecriteria,"H:/Biochem/MAPS Project/plasma/outlier/outlier.csv",sep=",",col.names=FALSE,append=TRUE,row.names=FALSE)

##write outlier to excel in batchreport file
batchoutlier <- read.csv(file="H:/Biochem/MAPS Project/plasma/outlier/outlier.csv",header=TRUE,sep=",")

if(nrow(batchoutlier)==0) {
  for(i in 1:ncol(batchoutlier))
    batchoutlier[1,i] = 'NA'
}

currentdate <- Sys.Date()
batchreportname <- paste("H:/Biochem/MAPS Project/plasma/batchreport/plasma_batchreport_",currentdate,".xlsx",sep="")
write.xlsx(batchoutlier,file=batchreportname,sheetName="Outlier",col.names=TRUE, row.names=FALSE,append=TRUE)


##iterate all X files and delete outlier compounds, then save to upload files

## if function to make sure the outlier compound is exist, if outlier compound is not exist, just save Xpatient file to upload file
if(nrow(positivecriteria)+nrow(negativecriteria)>0) {
  
##iterate X files  
  for(i in patientnumber) {
    patientXName1 <- toString(i)
    
##read patient x file
    uploadname <- paste("H:/Biochem/MAPS Project/plasma/CLIA raw file/Xpatientnumber/X",patientXName1,".csv",sep="")
    uploaddata1 <- read.csv(file=uploadname,header=TRUE,sep=",",stringsAsFactors=FALSE)
    
##read outlier compounds file    
    outliercompounds <- read.csv(file="H:/Biochem/MAPS Project/plasma/outlier/outlier.csv",sep=",",header=TRUE,stringsAsFactors=FALSE)
    
##exclude outlier compounds in patient x file
    uploaddata1 = uploaddata1[-(match(outliercompounds$CHEMICAL_ID,uploaddata1$CHEMICAL_ID)),]
    
##remove LIB_ID column    
    uploaddata1$LIB_ID <- NULL
    
##add Table ID and On Report columns
    uploaddata1$newcol6 <- rep(1,nrow(uploaddata1))
    names(uploaddata1)[names(uploaddata1) == 'newcol6'] <- 'Table ID'
    
    uploaddata1$newcol7 <- rep("Y",nrow(uploaddata1))
    names(uploaddata1)[names(uploaddata1) == 'newcol7'] <- 'On Report'
    
##save to upload file    
    uploadfilename <- paste("H:/Biochem/MAPS Project/plasma/upload/",patientXName1,"P.csv",sep="")
    write.csv(uploaddata1,file=uploadfilename,row.names = FALSE)
  }
} else {
  for(i in patientnumber) {
    patientXName1 <- toString(i)
    uploadname <- paste("H:/Biochem/MAPS Project/plasma/CLIA raw file/Xpatientnumber/X",patientXName1,".csv",sep="")
    uploaddata1 <- read.csv(file=uploadname,header=TRUE,sep=",",stringsAsFactors=FALSE)
    
    ##remove LIB_ID column    
    uploaddata1$LIB_ID <- NULL
    
    ##add Table ID and On Report columns
    uploaddata1$newcol6 <- rep(1,nrow(uploaddata1))
    names(uploaddata1)[names(uploaddata1) == 'newcol6'] <- 'Table ID'
    
    uploaddata1$newcol7 <- rep("Y",nrow(uploaddata1))
    names(uploaddata1)[names(uploaddata1) == 'newcol7'] <- 'On Report'
    
    uploadfilename <- paste("H:/Biochem/MAPS Project/plasma/upload/",patientXName1,"P.csv",sep="")
    write.csv(uploaddata1,file=uploadfilename,row.names = FALSE)
  }
  }

print("Outliers checked, if exist, outliers will be removed from upload file ")
##table 2 find rare compounds and table 3 (common compounds missing)

    allrawdata <- read.csv(file="H:/Biochem/MAPS Project/plasma/Patient_Master_list_raw.csv",header=TRUE,sep=",")
    allrawdatafortable3 <- allrawdata
    ncolumn <- ncol(allrawdata)
##count number of NAs in current batch (table 2)    
    allrawdata$newcol8 <- rowSums(!is.na(allrawdata[,(ncolumn-patientquantity+1):ncolumn]))
    names(allrawdata)[names(allrawdata)=='newcol8'] <- 'currentbatch'
##count number of NAs in current batch (table 3)    
    allrawdatafortable3$newcol8 <- rowSums(!is.na(allrawdatafortable3[,(ncolumn-patientquantity+1):ncolumn]))
    names(allrawdatafortable3)[names(allrawdatafortable3)=='newcol8'] <- 'currentbatch'

##count number of NAs in all samples (table 2)     
    allrawdata$newcol9 <- rowSums(!is.na(allrawdata[,12:ncolumn]))
    allrawdata$newcol10 <- allrawdata$newcol9/(ncolumn-12+1)
    
##count number of NAs in all samples (table 3)   
    allrawdatafortable3$newcol9 <- rowSums(!is.na(allrawdatafortable3[,12:ncolumn]))
    allrawdatafortable3$newcol10 <- allrawdatafortable3$newcol9/(ncolumn-12+1)
    
##rename column names in table 2    
    names(allrawdata)[names(allrawdata)=='newcol9'] <- 'allsamples'
    names(allrawdata)[names(allrawdata)=='newcol10'] <- 'percentage'
##rename column names in table 3    
    names(allrawdatafortable3)[names(allrawdatafortable3)=='newcol9'] <- 'allsamples'
    names(allrawdatafortable3)[names(allrawdatafortable3)=='newcol10'] <- 'percentage'    
    
##copy another set of data and process it to generate table 2    
    allrawdata1 <- allrawdata
##copy another set of data and process it to generate table 3    
    allrawdatafortable3b <- allrawdatafortable3    
    
    
##write the data to pretable 2    
    write.csv(allrawdata,file="H:/Biochem/MAPS Project/plasma/table2rare/pretable2.csv",row.names=FALSE)
##write the data to pretable 3    
    write.csv(allrawdatafortable3,file="H:/Biochem/MAPS Project/plasma/table3missingcommon/pretable3.csv",row.names=FALSE)
    
        
##sebset the data with the conbination conditions in table 2   
    allrawdata1 <- subset(allrawdata1, (percentage <= 0.1) & (currentbatch > 0) & (currentbatch <= patientquantity/2) & !(SUPER_PATHWAY=="Peptide") & !(SUB_PATHWAY=="Peptide") )
##only present compound name, chemical ID, patient number, and last three columns in table 2    
    allrawdata1 <- allrawdata1[,c((1:2),(4:5),((ncolumn-patientquantity+1):(ncolumn+3)))] 
    write.csv(allrawdata1,file="H:/Biochem/MAPS Project/plasma/table2rare/table2.csv",row.names=FALSE)

##subset the data with the conbination conditions in table 3   
    allrawdatafortable3b <- subset(allrawdatafortable3b, (percentage >= 0.99) & (currentbatch < patientquantity) & !(SUPER_PATHWAY=="Peptide") & !(SUB_PATHWAY=="Peptide") )
##only present compound name, chemical ID, patient number, and last three columns in table 3    
    allrawdatafortable3b <- allrawdatafortable3b[,c((1:2),(4:5),((ncolumn-patientquantity+1):(ncolumn+3)))] 
    write.csv(allrawdatafortable3b,file="H:/Biochem/MAPS Project/plasma/table3missingcommon/table3.csv",row.names=FALSE)

##clean table 2 and table 3 for report
    ## final table 2
    pretable2 <- read.csv(file="H:/Biochem/MAPS Project/plasma/table2rare/table2.csv",header=TRUE,sep=",")
    
    table2 <- pretable2[,(1:4)]
    
    pretable2column <- ncol(pretable2)
    pretable2row <- nrow(pretable2)
    
    if(pretable2row > 0) {
      
      for(i in 1:pretable2row) {
        table2[i,5] <- ""
        for(j in 5:(pretable2column-3)) {
          if(!is.na(pretable2[i,j])) {
            table2[i,5] <- paste(table2[i,5],colnames(pretable2)[j],sep=" ")
          }
        }
      }
      
      
    }
    
    table2 <- bind_cols(table2,pretable2[,(pretable2column-2):pretable2column])
    
    write.csv(table2,file="H:/Biochem/MAPS Project/plasma/table2rare/table2final.csv",row.names=FALSE)
    
    if(nrow(table2)==0) {
      for(i in 1:ncol(table2))
        table2[1,i] = 'NA'
    }
    
    currentdate <- Sys.Date()
    batchreportname <- paste("H:/Biochem/MAPS Project/plasma/batchreport/plasma_batchreport_",currentdate,".xlsx",sep="")
    write.xlsx(table2,file=batchreportname,sheetName="Table 2",col.names=TRUE, row.names=FALSE,append=TRUE)
    
    ##final table 3
    pretable3 <- read.csv(file="H:/Biochem/MAPS Project/plasma/table3missingcommon/table3.csv",header=TRUE,sep=",")
    
    table3 <- pretable3[,(1:4)]
    
    pretable3column <- ncol(pretable3)
    pretable3row <- nrow(pretable3)
    
    if(pretable3row > 0) {
      
      for(i in 1:pretable3row) {
        table3[i,5] <- ""
        for(j in 5:(pretable3column-3)) {
          if(is.na(pretable3[i,j])) {
            table3[i,5] <- paste(table3[i,5],colnames(pretable3)[j],sep=" ")
          }
        }
      }
      
    }
    
    table3 <- bind_cols(table3,pretable3[,(pretable3column-2):pretable3column])
    
    write.csv(table3,file="H:/Biochem/MAPS Project/plasma/table3missingcommon/table3final.csv",row.names=FALSE)
    
    if(nrow(table3)==0) {
      for(i in 1:ncol(table3))
        table3[1,i] = 'NA'
    }
    
    write.xlsx(table3,file=batchreportname,sheetName="Table 3",col.names=TRUE, row.names=FALSE,append=TRUE)
##implement table 2 in each upload file, change Table ID to 2, change zscore to 0.1
    if(nrow(allrawdata1)>0) {
      for(i in patientnumber) {
        patientXName2 <- toString(i)
        uploadname2 <- paste("H:/Biochem/MAPS Project/plasma/upload/",patientXName2,"P.csv",sep="")
        uploaddata2 <- read.csv(file=uploadname2,header=TRUE,sep=",",check.names = FALSE)
        
        table2data <- read.csv(file="H:/Biochem/MAPS Project/plasma/table2rare/table2.csv",header=TRUE,sep=",")
        columnname <- paste("X",patientXName2,sep="")
##in table 2, find the patient numbers with no NAs, based on chemical ID, change Table ID and zscore in upload files         
        for(j in which(!is.na(table2data[,columnname]))) {
          uploaddata2[match(table2data[j,2],uploaddata2$CHEMICAL_ID),9] <- 2
          uploaddata2[match(table2data[j,2],uploaddata2$CHEMICAL_ID),5] <- 0.1
          
        }
##re-write to upload file        
        write.csv(uploaddata2,file=uploadname2,row.names = FALSE)
      }
    }

##like table 2, implement table 3 in each upload file, change Table ID to 3, change zscore to 99
    if(nrow(allrawdatafortable3b)>0) {
      for(i in patientnumber) {
        patientXName3 <- toString(i)
        uploadname3 <- paste("H:/Biochem/MAPS Project/plasma/upload/",patientXName3,"P.csv",sep="")
        uploaddata3 <- read.csv(file=uploadname3,header=TRUE,sep=",",check.names = FALSE,stringsAsFactors=FALSE)
        
        table3data <- read.csv(file="H:/Biochem/MAPS Project/plasma/table3missingcommon/table3.csv",header=TRUE,sep=",",stringsAsFactors=FALSE)
        columnname <- paste("X",patientXName3,sep="")
##in table 3, find the patient numbers with NAs, based on chemical ID, change Table ID and zscore in upload files         
        for(j in which(is.na(table3data[,columnname]))) {
          uploaddata3[match(table3data[j,2],uploaddata3$CHEMICAL_ID),9] <- 3
          uploaddata3[match(table3data[j,2],uploaddata3$CHEMICAL_ID),5] <- 0.99
          
        }
        ##re-write to upload file        
        write.csv(uploaddata3,file=uploadname3,row.names = FALSE)
      }
    }
 print("table 2 and table 3 were generated, upload files were updated")   
##finalize the upload files
    
##iterate all upload files
for(i in patientnumber) {
    patientXName4 <- toString(i)
    uploadname4 <- paste("H:/Biochem/MAPS Project/plasma/upload/",patientXName4,"P.csv",sep="")
    uploaddata4 <- read.csv(file=uploadname4,header=TRUE,sep=",",check.names = FALSE,stringsAsFactors=FALSE)
  
    ##remove CHEMICAL_ID, RAW, and Anchor columns  
    uploaddata4 <- uploaddata4[,c(-2,-3,-4)]
  
    ##remove zscore with NAs and order it from largest to smallest
    uploaddata4 <- uploaddata4[order(-uploaddata4$zscore,na.last = NA),]
  
    ##in Super_Pathway column, delete all peptides
    uploaddata4 <- uploaddata4[!uploaddata4$SUPER_PATHWAY=="Peptide",]
  
    
    ##change Table ID of xenobiotics to 4
    nrowupload <- nrow(uploaddata4)
    for(j in 1:nrowupload) {
      if(uploaddata4[j,3]=="Xenobiotics" & !uploaddata4[j,2]==0.1 & !uploaddata4[j,2]==0.99) 
        uploaddata4[j,6] = 4
    ##change on report of steroids to N
      if(uploaddata4[j,4]=="Steroids" | uploaddata4[j,4]=="Steroid")
        uploaddata4[j,7] = "N"
      if(uploaddata4[j,2]<2 & uploaddata4[j,2]>-2 & !uploaddata4[j,2]==0.1 & !uploaddata4[j,2]==0.99)
        uploaddata4[j,6] = ""
    }
    write.csv(uploaddata4,file=uploadname4,row.names=FALSE)
}
##generate batchperanalytes table
 analytesperbatch <- read.table(file="H:/Biochem/MAPS Project/plasma/analytesperbatch/analytesperbatch.csv",header=TRUE,sep=",",stringsAsFactors = FALSE)
 batchdate <- as.character(Sys.Date())
 
 firstxfilename <- paste("H:/Biochem/MAPS Project/plasma/CLIA raw file/Xpatientnumber/X",patientnumber[1],".csv",sep="")
 firstxfile <- read.csv(firstxfilename,header=TRUE, sep=",")
 
 analytesnumber <- nrow(firstxfile)
 nonzscore <- sum(is.na(firstxfile$zscore))
 
 z <- data.frame(batchdate,analytesnumber,nonzscore,stringsAsFactors = FALSE)
 
 analytesperbatch <- bind_rows(analytesperbatch,z)
 analytesperbatch <- na.omit(analytesperbatch)
 
 write.csv(analytesperbatch,file="H:/Biochem/MAPS Project/plasma/analytesperbatch/analytesperbatch.csv",row.names=FALSE)
 
 currentdate <- Sys.Date()
 batchreportname <- paste("H:/Biochem/MAPS Project/plasma/batchreport/plasma_batchreport_",currentdate,".xlsx",sep="")
 write.xlsx(analytesperbatch,file=batchreportname,sheetName="analytes per batch",col.names=TRUE, row.names=FALSE,append=TRUE)
 
##format outlier, if value>=2.99 cell color changed to green; if value <=-2.99, cell color change to red
 
 batchoutlier <- read.csv(file="H:/Biochem/MAPS Project/plasma/outlier/outlier.csv",header=TRUE,sep=",")
 
 if(nrow(batchoutlier)>0) {
 currentdate <- Sys.Date()
 batchreportname <- paste("H:/Biochem/MAPS Project/plasma/batchreport/plasma_batchreport_",currentdate,".xlsx",sep="")
 wb <- loadWorkbook(batchreportname)              # load workbook
 fo <- Fill(foregroundColor="green")  # create fill object
 cs <- CellStyle(wb, fill=fo)          # create cell style
 fo1 <- Fill(foregroundColor="red")
 cs1 <- CellStyle(wb,fill=fo1)
 sheets <- getSheets(wb)               # get all sheets
 sheet <- sheets[[2]]          # get specific sheet
 rows <- getRows(sheet, rowIndex=2:(nrow(batchoutlier)+1))
 cells <- getCells(rows, colIndex = 3:(ncol(batchoutlier)-2))
 
 values <- lapply(cells, getCellValue) # extract the values
 
 # find cells meeting conditional criteria 
 highlight <- "test"
 for (i in names(values)) {
   x <- as.numeric(values[i])
   if (x>=2.99 & !is.na(x)) {
     highlight <- c(highlight, i)
   }    
 }
 highlight <- highlight[-1]
 
 lapply(names(cells[highlight]),
        function(ii)setCellStyle(cells[[ii]],cs))
 
 highlight <- "test"
 for (i in names(values)) {
   x <- as.numeric(values[i])
   if (x<=-2.99 & !is.na(x)) {
     highlight <- c(highlight, i)
   }    
 }
 highlight <- highlight[-1]
 
 lapply(names(cells[highlight]),
        function(ii)setCellStyle(cells[[ii]],cs1)) 
 
 saveWorkbook(wb, batchreportname)
 }    
print("Done!")
print(proc.time() - ptm)