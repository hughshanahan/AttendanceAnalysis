setwd("~/OneDrive - Royal Holloway University of London/Attendance/Progressions/Session Data")
soure("Analysis.R")
CS1811 <- createAttendanceList("CS1811", "CS1811_CS1803/CS1811 Week 6_10_sessions.csv")
CS1890 <- createAttendanceList("CS1890", "CS1890/CS1890 Week 6_10_sessions.csv")
CS1860 <- createAttendanceList("CS1860", "CS1860/CS1860 Week 6_10_sessions.csv")
#CS1860 <- mergeModuleAttendanceLists(CS1860_3,CS1860_5)

master <- createMasterAttendanceList(CS1811,CS1860,CS1890)

createSummaryMasterAttendanceList(master,"Summary Weeks 6 to 10.csv")

