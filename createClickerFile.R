createClickerFile <- function(masterFn, courseListFn, newFn){
 master <- read.csv(masterFn) 
 courseList <- read.csv(courseListFn)
 final <- master[which(master$Current.ID 
                       %in% courseList$ID),]    
 write.table(final,file=newFn, sep=",",quote=FALSE,
           row.names=FALSE,col.names=FALSE)  
}
#This builds a file for Turning Point for an individual module. 
# If you want to create one for a number of modules (i.e. ones that are 
# using the same lecture, then pass a vector of module names)

createClickerModuleFile <- function(masterClickerFile, module){
  system("scp hugh@linux.cim.rhul.ac.uk:/CS/StudentData/CS/CSregistrations.csv .")
  allClickers <- read.csv(masterClickerFile)
  master <- read.table("CSregistrations.csv",sep="\t",quote="\"",stringsAsFactors = FALSE,header=TRUE)
  system("/bin/rm -f CSregistrations.csv")
  thisModule <- master[master$Unit %in% module, ]
  if ( length(module) > 1 ){
    rootFN <- paste(module,collapse="_")
  }
  else{
    rootFN <- module
  }
  clickerModule <- allClickers[allClickers$Current.ID %in% thisModule$BannerID,]
  moduleClickerFN <- paste(rootFN,"StudentData.csv",sep="")
  write.table(clickerModule,file=moduleClickerFN, sep=",",quote=FALSE,
              row.names=FALSE,col.names=FALSE)  
}

setwd("~/OneDrive - Royal Holloway University of London/Attendance/StudentData/")
masterFn <- "ClickerListWeek5.csv"
courseListFn <- "RawData/Week 5 Attendance Spreadsheet CS1803 CS1811.csv"
newFn <- "CS1811_CS1803/StudentData.csv"

createClickerFile(masterFn,courseListFn,newFn )

