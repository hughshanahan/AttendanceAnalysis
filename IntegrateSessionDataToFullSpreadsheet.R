library("readxl")
library("writexl")
library("lubridate")
library("stringr")

# Read in CSV file generated from TurningPoint
readRawCSV <- function(fn){
  currentDF <- read.csv(fn,skip=1)
  n <- grep("X[0-9][0-9].[0-9][0-9].[0-9][0-9][0-9][0-9].[0-9][0-9].[0-9][0-9]",names(currentDF))
  dateStamps <- names(currentDF)[n]
  currentDF
}

# Read in Excel spreadsheet for a specific module

readModuleAttendanceData <- function(fn,moduleName){
  n <- findModuleSheet(fn,moduleName)
  read_xlsx(fn,sheet=n)
}

# Search a sheets object with a module and return 
# the number of the sheet 
findModuleSheet <- function( fn, module ){
  grep(module,excel_sheets(fn))
}

# Search for the year the module is run 
findYear <- function( fn, module ){
  n <- findModuleSheet(fn, module)
  sheetNames <- excel_sheets(fn)
  thisModule <- sheetNames[n]
  season <- unlist(strsplit(thisModule,"[\\(\\)]"))[2]
  academicYear <- unlist(strsplit(sheetNames[1],"20"))[2]
  years <- unlist(strsplit(academicYear,"-"))
  if ( season == "A" ){
    rightYear = years[1]
  }
  else{
    rightYear = years[2]
  }
  rightYear
}

#Go through dates on module attendance sheet and construct 
#names that they would have on the CSV

createCSVColumnDates <- function(sheet, year){
  numWeeks <- 10
  daysOfTheWeek <- list(M=0,TU=1,W=2,TH=3,F=4)
  offset <- 16
  ncols <- dim(sheet)[2]
  TO <- offset + (4*10) 
  if ( offset + (4*10) > ncols ){
    TO <- ncols
  }
  columnDates <- list()
  weekData <- names(sheet)
  startOfWeek <- c()
  for ( i in seq(from=offset,to=TO) ){
    iStr <- toString(i)
    
    if ( grepl("Week", weekData[i])){
      dm <- strsplit(weekData[i],"w/c")[[1]][2]
      dm <- sub(" ","",dm)
      startOfWeek <- dmy(paste(dm,year,sep="/"))
    }
    
    if ( !grepl("Lab",sheet[1,i],ignore.case = TRUE) ){
      columnDates[[iStr]] <-  list()
      columnDates[[iStr]]$date <- startOfWeek + 
        ddays(daysOfTheWeek[[as.character(sheet[2,i])]])
    }

  }
  columnDates
}

findCSVColumns <- function(sheet,attendanceDF, columnDates){
  moduleCols <- list()
    allCols <- names(columnDates)
    for ( myCol in allCols ){
     myDate <- columnDates[[myCol]]$date
     myDay <- twoIntString(day(myDate))
     myMonth <- twoIntString(month(myDate))
     myYear <- twoIntString(year(myDate))
     dateFormat <- paste(myDay,myMonth,myYear,sep=".")
     dateFormat <- paste("X",dateFormat,sep="")
     n <- grep(dateFormat,names(attendanceDF))
     if (length(n) > 0){
       if ( length(n) > 1) { stop(paste("Too many entries found for",dateFormat,"\n",n))}
       moduleCols[[myCol]] <- n
     }
    }
  moduleCols  
}

twoIntString <- function(n){
  str_pad(n,2,pad="0")
}

tryCatch.W.E <- function(expr)
{
  W <- NULL
  w.handler <- function(w){ # warning handler
    W <<- w
    invokeRestart("muffleWarning")
  }
  list(value = withCallingHandlers(tryCatch(expr, error = function(e) e),
                                   warning = w.handler),
       warning = W)
}

updateSheet <- function(sheet, attendanceDF, csvCols){
  sheetIds <- as.data.frame(sheet[4:dim(sheet)[1],1])
  sheetIds <- sheetIds[,1]
  attendanceIds <- attendanceDF$User.ID
  attendanceIds <- attendanceIds[!is.na(attendanceIds)]
  aDF <- attendanceDF[attendanceDF$User.ID %in% attendanceIds,]
# First do some sanity checks in terms of what ID's are missing etc.

  print(paste("IDs in Spreadsheet not listed in CSV are ",
              setdiff(as.character(sheetIds),as.character(attendanceIds))))
  print(paste("IDs in CSV not listed in Spreadsheet are ",
              setdiff(as.character(attendanceIds),as.character(sheetIds))))
  
  for ( id in intersect( sheetIds, attendanceIds) ){
    for ( sheetCol in names(csvCols) ){
      i <- strtoi(sheetCol)
      j <- csvCols[[sheetCol]]
      p <- attendanceDF$User.ID %in% id
      A <- attendanceDF[p,j]
      if ( A != 1 ){
        q <- which(sheet[,1] == id)
        if ( is.na(sheet[q,i]) ) {
          sheet[q,i] = "X"
        }
          
      }
    }
  }  
  sheet
}
runThis <- function(spreadsheetFn, csvFn, newSpreadsheetFn, module ){
  attendanceDF <- readRawCSV(csvFn)
  sheet <- readModuleAttendanceData(spreadsheetFn,module)
  year <- findYear(spreadsheetFn,module)
  allCols <- createCSVColumnDates(sheet,year)
  csvCols <- findCSVColumns(sheet, attendanceDF, allCols)
  updatedSheet <- updateSheet(sheet, attendanceDF, csvCols)
  print(class(updatedSheet))
  write_xlsx(data.frame(updatedSheet),newSpreadsheetFn)
  
}
