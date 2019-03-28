parseTutorialData <- function(fn,master){
  library(xlsx)
  file <- system.file("tests", fn, package = "xlsx")
  wb <- loadWorkbook(fn)
  sheets <- getSheets(wb)
  for ( sheet in sheets ){
    df <- readColumns(sheet,startColumn=1,endColumn=16,startRow=9)
    tutorials <- getTutorialAttendance(df)
    apply(tutorials,1,function(r){
      master$as.character(r$ID)$tutorials = r$attend
    })
  }
    
}

orderAttendanceDFDates <- function(aDF){
  theseNames <- names(aDF)
  n <- grep("X[0-9][0-9].[0-9][0-9].[0-9][0-9][0-9][0-9].[0-9][0-9].[0-9][0-9]",theseNames)
  theseStrings <- substr(theseNames[n],2,11)   
  theseDates <- as.Date(theseStrings,"%d.%m.%Y")
  aDF[,c(1:5, n[sort.int(theseDates,index.return=TRUE)$ix])]
}


# Find all the modules in a master Attendance list and return them as a vector
findModulesFromMaster <- function(master){
  allNames <- unique(unlist(lapply(master,names)))
  setdiff(allNames,c("firstName","lastName"))
}
#This creates a dataframe from a master Attendance list and saves it to a CSV
createSummaryMasterAttendanceList <- function(master,fn){
  dfNames <- c("Student ID", "Last Name", "First Name", findModulesFromMaster(master))
  IDs <- names(master)
  firstNames <- as.vector(sapply(master,function(r){r[["firstName"]]})[IDs])
  lastNames <- as.vector(sapply(master,function(r){r[["lastName"]]})[IDs]) 
  df0 <- data.frame(IDs,
                   lastNames,
                   firstNames)
  
  for (module in findModulesFromMaster(master) ){
    
   t <- sapply(master,function(r){
      if ( module %in% names(r) ){
        round(r[[module]])
      }
      else{
        "-"
      }
    })
   df1 <- cbind(df0,  as.vector(t[IDs]))
   df0 <- df1 
  }

  names(df0) <- dfNames
  avs <- apply(df0,1, function(r){
    theseAverages <- r[findModulesFromMaster(master)]
    round(mean(as.numeric(theseAverages[theseAverages!="-"]))) 
  })
  df1 <- cbind(df0,avs)
  names(df1) <- c(names(df0),"Average")
  write.csv(df1,file=fn)
}

# The following accepts as arguments a set of attendance lists and converts it into 
# one master list, where the top identifier is the student ID and for each ID there is a 
# list of modules with the percentage total attendance for that module along with first and last names.
createMasterAttendanceList <- function(...){
  x <- list(...)
  master <- list()
  for ( d in x ){
    module=d$module
    maxAttendance = d$maxAttendance
    ids <- names(d$students)
    for ( id in ids ){
      idData <- d$students[[id]]
      if (!exists(id, where=master)){
        master[[id]]$firstName = idData$firstName
        master[[id]]$lastName = idData$lastName
      }
      master[[id]][[module]] = ( 100.0 * idData$attended ) / ( 1.0 * maxAttendance ) 
    }
  }
  master
}

# The following merges a set of attendance lists for a specific module 
# The top identifiers for the output are 
# the module name
# the maximum possible attendance for the merged data 
# another list with the student data (first Name, last Name, total attendance, key is id)
# the inputs are a set of lists of the same type 
mergeModuleAttendanceLists <- function(...){
  x <- list(...)
  merged <- list()
  merged$module <- ""
  merged$maxAttendance = 0
  merged$students <- list()
  for ( d in x ){
    if (merged$module != "" & merged$module != d$module){
      stop(paste("Cannot match", module,"and",d$module))
    }
    merged$module = d$module
    merged$maxAttendance = merged$maxAttendance + d$maxAttendance
    
    if ( length(merged$students) == 0 ){
      merged$students <- d$students
    }
    else{
      ids <- names(d$students)
      for ( id in ids ){
        idData <- d$students[[id]]
        if ( exists( id, where=merged$students  )){
          newSum <- merged$students[[id]]$attended + idData$attended
          merged$students[[id]]$attended <- newSum 
        }
        else{
          merged$students[[id]]$attended <- idData$attended
          merged$students[[id]]$firstName <- idData$firstName
          merged$students[[id]]$lastName <- idData$lastName
        }
      }
    }
  }
  merged
}

#The following reads in a data frame of attendance data and 
#converts it into a list of the following form
# the module name
# the maximum possible attendance for the merged data 
# another list with the student data (first Name, last Name, total attendance, key is id)
createAttendanceList <- function(module, fn){
  currentDF <- read.csv(fn,skip=1)
  n <- grep("X[0-9][0-9].[0-9][0-9].[0-9][0-9][0-9][0-9].[0-9][0-9].[0-9][0-9]",names(currentDF))
  attendance <- list()
  attendance$module <- module
  attendance$maxAttendance <- length(n)
  currentDF$User.ID <- as.character(currentDF$User.ID)
  ids <- currentDF$User.ID
  ids <- ids[!is.na(ids)]
  if ( length(ids) != length(unique(ids))){
    stop(paste("IDs are not unique in",fn,"\nIDs are", ids[duplicated(ids)]))
  }
  attendance$students <- list()
  for ( id in ids ){
    attendance$students[[id]] <- list()
    r <- currentDF[currentDF$User.ID %in% id,]
    if ( dim(r)[1] > 1){
      stop(paste("Found multiple entries for",id,r))
    }
    attendance$students[[id]]$firstName <- r$First.Name
    attendance$students[[id]]$lastName <- r$Last.Name
    attendance$students[[id]]$attended <- attendance$maxAttendance - length(which(r[n]=="-"))
  }
  attendance
}

mergeDFs <- function(fileNames){
  finalDF <- data.frame()
  for ( f in fileNames){
    print(paste("Processing", f))
    currentDF <- read.csv(f,skip=1)
    n <- grep("X[0-9][0-9].[0-9][0-9].[0-9][0-9][0-9][0-9].[0-9][0-9].[0-9][0-9]",names(currentDF))
    
    if ( length(finalDF) == 0 ){
      finalDF <- currentDF[,c(1:5,n)]
    }
    else{
      IDs <- currentDF$User.ID[!is.na(currentDF$User.ID)]
# Get the ID's that are the shared 
      sharedIDs <- intersect(IDs,finalDF$User.ID)
      sharedIDs <- sharedIDs[!is.na(sharedIDs)]
      newDF <- cbind(finalDF[finalDF$User.ID %in% sharedIDs,],
                     currentDF[currentDF$User.ID %in% sharedIDs,n])
# Get ID's in the current file but not the sum
      c1 <- setdiff(IDs, 
                    finalDF$User.ID[!is.na(finalDF$User.ID)])
      nx <- length(names(finalDF)) - 5
      ny <- length(c1)
      tm <- matrix(c(rep("-",nx*ny)),nrow=ny,ncol=nx) 
      tDF <- data.frame(currentDF[currentDF$User.ID %in% c1,1:5],
                        tm,
                        currentDF[currentDF$User.ID %in% c1,n])
      names(tDF) <- c(names(finalDF),names(currentDF)[n])
      newDF <- rbind(newDF,tDF)
# Get ID's in the sum but not the current file      
      c2 <- setdiff(finalDF$User.ID[!is.na(finalDF$User.ID)],
                    IDs)
      nx <- length(n)
      ny <- length(c2)
      tm <- matrix(c(rep("-",nx*ny)),nrow=ny,ncol=nx)
      tDF <- data.frame(finalDF[finalDF$User.ID %in% c2,],
                        tm)
      names(tDF) <- c(names(finalDF),names(currentDF)[n])
      finalDF <- rbind(newDF,tDF)
    }
  }
  finalDF
}

getTotalAttendance <- function(merged){
  n <- grep("X[0-9][0-9].[0-9][0-9].[0-9][0-9][0-9][0-9].[0-9][0-9].[0-9][0-9]",
            names(merged))
  fAttend <- apply(merged,1,function(r){ 
                    nT <- length(n) * 1.0
                    nA <- nT - length(grep("-",r[n]))
                    round( (nA * 100.0) / nT ) })
  
  names(fAttend) <- merged$User.ID
  cbind(merged,PerCent=fAttend)
}

