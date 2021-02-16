#Going to try a new strategy
#Want all variants of the timepoints
#51, 52, 93, 

#Need 2 inital pieces of information
#First, need the Sequence of Timepoints for each Route (Organize the timepoint output)
#Second, need a good type of export (HASTUS Trip Stop)

#####Data Setup#####

library(readxl)
library(data.table)
library(dplyr)
library(tidyr)
library(lubridate)
library(stringr)
library(openxlsx)

#Pulling in all excel data from this folder (MVS Info)
setwd("G:/OPER_ADM/Service Evaluation/PowerBI/TimeTable Generator/TripStop Source")
file.list <- list.files(pattern = '*.xlsx')
df.list <- lapply(file.list, function(x) read_excel(x, col_types = "text"))
setwd("~/")

#Combining all excel documents into 1 big data frame
attr(df.list, "names") <- file.list
CombinedDF = rbindlist(df.list, idcol="id")

#Cleaning up the file
CombinedDF <- CombinedDF %>%
  filter(!is.na(Route)) %>%
  filter(Route != "Car")

CombinedDF$tstp_pass_time_min <- as.numeric(CombinedDF$tstp_pass_time_min)

#CombinedDF$AdjTime <- ifelse((CombinedDF$tstp_pass_time_min >= 1440), (CombinedDF$tstp_pass_time_min - 1440), CombinedDF$tstp_pass_time_min)
#CombinedDF$AdjTime <- CombinedDF$tstp_pass_time_min
#Extract Routes that need a stop to be displayed on top of timepoints

#Pulling in all excel data from this folder (Marketing TP Info)
setwd("G:/OPER_ADM/Service Evaluation/PowerBI/TimeTable Generator/Modify These To Change Timepoint Order")
file.list <- list.files(pattern = '*.xlsx')
df.list <- lapply(file.list, function(x) read_excel(x, col_types = "text"))
setwd("~/")

#Combining all excel documents into 1 big data frame
attr(df.list, "names") <- file.list
MarketingTP = rbindlist(df.list, idcol="id")

CombinedDF$DepartureSave <- CombinedDF$Departure
CombinedDF <- extract(CombinedDF, Departure, into = c("Time", "Meridian"), "^(\\d+)([a-z]+)$")

CombinedDF$Time <- as.numeric(CombinedDF$Time)

CombinedDF$TimePMLogic <- ifelse(CombinedDF$Meridian == "p" & CombinedDF$Time < 1200, 1200, 0)
CombinedDF$TimeXLogic1 <- ifelse(CombinedDF$Meridian == "x" & CombinedDF$Time > 1200, 1200, 0)
CombinedDF$TimeXLogic2 <- ifelse(CombinedDF$Meridian == "x" & CombinedDF$Time < 1200, 2400, 0)
CombinedDF$DepTime24h <- (CombinedDF$Time + CombinedDF$TimePMLogic + CombinedDF$TimeXLogic1 + CombinedDF$TimeXLogic2)
CombinedDF$DepartureSave <- str_replace(CombinedDF$DepartureSave, "x", paste0("a"))

#Extracting all stops
#Encountering problems, need to create the Route Sequences separately. Cannot use stop_rank.
TPCombinedDF <- CombinedDF %>%
  inner_join(MarketingTP, by = c("Route", "Direction", "Stop")) %>%
  select(trp_number_trim, Route, Direction, Place, Stop, `Op Day`, `Stop Name`, DepartureSave, DepTime24h, `Stop Order`, `Evt, Sta`) %>%
  unite(StopInfo, c(`Stop Name`, Stop), sep = " - ", remove = F)

#TPCombinedDF$stop_rank <- as.numeric(TPCombinedDF$stop_rank)

TPCombinedDF$`Stop Order` <- as.numeric(TPCombinedDF$`Stop Order`)

TPCombinedDF <- TPCombinedDF %>%
  arrange(Route, Direction, `Stop Order`, DepTime24h)

TPCombinedDF$`Evt, Sta` <- ifelse(is.na(TPCombinedDF$`Evt, Sta`), "", TPCombinedDF$`Evt, Sta`)

TPCombinedDF <- TPCombinedDF %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = F) %>%
  distinct()

map <- setNames(c("a"), c("x"))


# TPCombinedDF$StopTime <- parse_time(TPCombinedDF$StopTime, '%H:%M %p')

#write.csv(TPCombinedDF, "TimetableBase.csv", row.names = F)

#Need a Route-Timepoint file in order to move onto next step
#Would probably need the first timepoint as a minimum
#Attempting to sort based on the initial timepoint and using minutes after midnight
#Should get a trp_number_trim order and a timepoint order

#####End Data Setup#####


###I might want to organize this coding mess into a function that can be called by feeding the variable into it

#Check this one#
#####Route 1#####

R1LE <- TPCombinedDF %>%
  filter(Route == "1" | Route == "1L") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R1LEProc <- R1LE %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R1LEProcCt <- R1LEProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R1LEProcCt <- colnames(R1LEProcCt)[max.col(R1LEProcCt,ties.method = "first")]

R1LEProc <- R1LEProc %>%
  arrange_at(R1LEProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R1LEProc2 <- R1LEProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R1LEProc2 <- left_join(R1LEProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R1LEProc2 <- R1LEProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R1LW <- TPCombinedDF %>%
  filter(Route == "1" | Route == "1L") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R1LWProc <- R1LW %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R1LWProcCt <- R1LWProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R1LWProcCt <- colnames(R1LWProcCt)[max.col(R1LWProcCt,ties.method = "first")]

R1LWProc <- R1LWProc %>%
  arrange_at(R1LWProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R1LWProc2 <- R1LWProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R1LWProc2 <- left_join(R1LWProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R1LWProc2 <- R1LWProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R1LEProc2[is.na(R1LEProc2)] <- "....."
R1LWProc2[is.na(R1LWProc2)] <- "....."

#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R1LEProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route1East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R1LWProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route1West.xlsx", overwrite = T)

# write.csv(R1LEProc2, "Route1East.csv", row.names = F)
# write.csv(R1LWProc2, "Route1West.csv", row.names = F)

#####Route 1 End#####


#####Route 2#####


R2E <- TPCombinedDF %>%
  filter(Route == "2") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R2EProc <- R2E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R2EProcCt <- R2EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R2EProcCt <- colnames(R2EProcCt)[max.col(R2EProcCt,ties.method = "first")]

R2EProc <- R2EProc %>%
  arrange_at(R2EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R2EProc2 <- R2EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R2EProc2 <- left_join(R2EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R2EProc2 <- R2EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R2W <- TPCombinedDF %>%
  filter(Route == "2") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R2WProc <- R2W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R2WProcCt <- R2WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R2WProcCt <- colnames(R2WProcCt)[max.col(R2WProcCt,ties.method = "first")]

R2WProc <- R2WProc %>%
  arrange_at(R2WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R2WProc2 <- R2WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R2WProc2 <- left_join(R2WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R2WProc2 <- R2WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R2EProc2[is.na(R2EProc2)] <- "....."
R2WProc2[is.na(R2WProc2)] <- "....."

#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R2EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route2East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R2WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route2West.xlsx", overwrite = T)

# write.csv(R2EProc2, "Route2East.csv", row.names = F)
# write.csv(R2WProc2, "Route2West.csv", row.names = F)

#####Route 2 End#####


#####Route 3#####


R3E <- TPCombinedDF %>%
  filter(Route == "3") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R3EProc <- R3E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R3EProcCt <- R3EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R3EProcCt <- colnames(R3EProcCt)[max.col(R3EProcCt,ties.method = "first")]

R3EProc <- R3EProc %>%
  arrange_at(R3EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R3EProc2 <- R3EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R3EProc2 <- left_join(R3EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R3EProc2 <- R3EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R3W <- TPCombinedDF %>%
  filter(Route == "3") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R3WProc <- R3W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R3WProcCt <- R3WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R3WProcCt <- colnames(R3WProcCt)[max.col(R3WProcCt,ties.method = "first")]

R3WProc <- R3WProc %>%
  arrange_at(R3WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R3WProc2 <- R3WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R3WProc2 <- left_join(R3WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R3WProc2 <- R3WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R3EProc2[is.na(R3EProc2)] <- "....."
R3WProc2[is.na(R3WProc2)] <- "....."

#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R3EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route3East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R3WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route3West.xlsx", overwrite = T)

# write.csv(R3EProc2, "Route3East.csv", row.names = F)
# write.csv(R3WProc2, "Route3West.csv", row.names = F)

#####Route 3 End#####


#Interlining problems on Westbound (Can't generate this one)
#####Route 7#####


R7E <- TPCombinedDF %>%
  filter(Route == "7") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R7EProc <- R7E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R7EProcCt <- R7EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R7EProcCt <- colnames(R7EProcCt)[max.col(R7EProcCt,ties.method = "first")]

R7EProc <- R7EProc %>%
  arrange_at(R7EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R7EProc2 <- R7EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R7EProc2 <- left_join(R7EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R7EProc2 <- R7EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R7W <- TPCombinedDF %>%
  filter(Route == "7") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R7WProc <- R7W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R7WProcCt <- R7WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R7WProcCt <- colnames(R7WProcCt)[max.col(R7WProcCt,ties.method = "first")]

R7WProc <- R7WProc %>%
  arrange_at(R7WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R7WProc2 <- R7WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R7WProc2 <- left_join(R7WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R7WProc2 <- R7WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R7EProc2[is.na(R7EProc2)] <- "....."
R7WProc2[is.na(R7WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R7EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route7East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R7WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route7West.xlsx", overwrite = T)



# write.csv(R7EProc2, "Route7East.csv", row.names = F)
# write.csv(R7WProc2, "Route7West.csv", row.names = F)

#####Route 7 End#####


#####Route 9#####


R9E <- TPCombinedDF %>%
  filter(Route == "9") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R9EProc <- R9E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R9EProcCt <- R9EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R9EProcCt <- colnames(R9EProcCt)[max.col(R9EProcCt,ties.method = "first")]

R9EProc <- R9EProc %>%
  arrange_at(R9EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R9EProc2 <- R9EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R9EProc2 <- left_join(R9EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R9EProc2 <- R9EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R9W <- TPCombinedDF %>%
  filter(Route == "9") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R9WProc <- R9W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R9WProcCt <- R9WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R9WProcCt <- colnames(R9WProcCt)[max.col(R9WProcCt,ties.method = "first")]

R9WProc <- R9WProc %>%
  arrange_at(R9WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R9WProc2 <- R9WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R9WProc2 <- left_join(R9WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R9WProc2 <- R9WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R9EProc2[is.na(R9EProc2)] <- "....."
R9WProc2[is.na(R9WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R9EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route9East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R9WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route9West.xlsx", overwrite = T)



# write.csv(R9EProc2, "Route9East.csv", row.names = F)
# write.csv(R9WProc2, "Route9West.csv", row.names = F)

#####Route 9 End#####


#####Route 10#####


R10E <- TPCombinedDF %>%
  filter(Route == "10") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R10EProc <- R10E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R10EProcCt <- R10EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R10EProcCt <- colnames(R10EProcCt)[max.col(R10EProcCt,ties.method = "first")]

R10EProc <- R10EProc %>%
  arrange_at(R10EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R10EProc2 <- R10EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R10EProc2 <- left_join(R10EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R10EProc2 <- R10EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R10W <- TPCombinedDF %>%
  filter(Route == "10") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R10WProc <- R10W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R10WProcCt <- R10WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R10WProcCt <- colnames(R10WProcCt)[max.col(R10WProcCt,ties.method = "first")]

R10WProc <- R10WProc %>%
  arrange_at(R10WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R10WProc2 <- R10WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R10WProc2 <- left_join(R10WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R10WProc2 <- R10WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R10EProc2[is.na(R10EProc2)] <- "....."
R10WProc2[is.na(R10WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R10EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route10East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R10WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route10West.xlsx", overwrite = T)



# write.csv(R10EProc2, "Route10East.csv", row.names = F)
# write.csv(R10WProc2, "Route10West.csv", row.names = F)

#####Route 10 End#####


#####Route 13#####


R13E <- TPCombinedDF %>%
  filter(Route == "13") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R13EProc <- R13E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R13EProcCt <- R13EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R13EProcCt <- colnames(R13EProcCt)[max.col(R13EProcCt,ties.method = "first")]

R13EProc <- R13EProc %>%
  arrange_at(R13EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R13EProc2 <- R13EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R13EProc2 <- left_join(R13EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R13EProc2 <- R13EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R13W <- TPCombinedDF %>%
  filter(Route == "13") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R13WProc <- R13W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R13WProcCt <- R13WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R13WProcCt <- colnames(R13WProcCt)[max.col(R13WProcCt,ties.method = "first")]

R13WProc <- R13WProc %>%
  arrange_at(R13WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R13WProc2 <- R13WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R13WProc2 <- left_join(R13WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R13WProc2 <- R13WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R13EProc2[is.na(R13EProc2)] <- "....."
R13WProc2[is.na(R13WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R13EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route13East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R13WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route13West.xlsx", overwrite = T)



# write.csv(R13EProc2, "Route13East.csv", row.names = F)
# write.csv(R13WProc2, "Route13West.csv", row.names = F)

#####Route 13 End#####


#####Route 14#####


R14E <- TPCombinedDF %>%
  filter(Route == "14") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R14EProc <- R14E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R14EProcCt <- R14EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R14EProcCt <- colnames(R14EProcCt)[max.col(R14EProcCt,ties.method = "first")]

R14EProc <- R14EProc %>%
  arrange_at(R14EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R14EProc2 <- R14EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R14EProc2 <- left_join(R14EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R14EProc2 <- R14EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R14W <- TPCombinedDF %>%
  filter(Route == "14") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R14WProc <- R14W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R14WProcCt <- R14WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R14WProcCt <- colnames(R14WProcCt)[max.col(R14WProcCt,ties.method = "first")]

R14WProc <- R14WProc %>%
  arrange_at(R14WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R14WProc2 <- R14WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R14WProc2 <- left_join(R14WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R14WProc2 <- R14WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R14EProc2[is.na(R14EProc2)] <- "....."
R14WProc2[is.na(R14WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R14EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route14East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R14WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route14West.xlsx", overwrite = T)



# write.csv(R14EProc2, "Route14East.csv", row.names = F)
# write.csv(R14WProc2, "Route14West.csv", row.names = F)

#####Route 14 End#####


#####Route 15#####


R15E <- TPCombinedDF %>%
  filter(Route == "15") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R15EProc <- R15E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R15EProcCt <- R15EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R15EProcCt <- colnames(R15EProcCt)[max.col(R15EProcCt,ties.method = "first")]

R15EProc <- R15EProc %>%
  arrange_at(R15EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R15EProc2 <- R15EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R15EProc2 <- left_join(R15EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R15EProc2 <- R15EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R15W <- TPCombinedDF %>%
  filter(Route == "15") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R15WProc <- R15W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R15WProcCt <- R15WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R15WProcCt <- colnames(R15WProcCt)[max.col(R15WProcCt,ties.method = "first")]

R15WProc <- R15WProc %>%
  arrange_at(R15WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R15WProc2 <- R15WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R15WProc2 <- left_join(R15WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R15WProc2 <- R15WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R15EProc2[is.na(R15EProc2)] <- "....."
R15WProc2[is.na(R15WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R15EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route15East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R15WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route15West.xlsx", overwrite = T)



# write.csv(R15EProc2, "Route15East.csv", row.names = F)
# write.csv(R15WProc2, "Route15West.csv", row.names = F)

#####Route 15 End#####


#####Route 16#####


R16E <- TPCombinedDF %>%
  filter(Route == "16") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R16EProc <- R16E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R16EProcCt <- R16EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R16EProcCt <- colnames(R16EProcCt)[max.col(R16EProcCt,ties.method = "first")]

R16EProc <- R16EProc %>%
  arrange_at(R16EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R16EProc2 <- R16EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R16EProc2 <- left_join(R16EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R16EProc2 <- R16EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R16W <- TPCombinedDF %>%
  filter(Route == "16") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R16WProc <- R16W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R16WProcCt <- R16WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R16WProcCt <- colnames(R16WProcCt)[max.col(R16WProcCt,ties.method = "first")]

R16WProc <- R16WProc %>%
  arrange_at(R16WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R16WProc2 <- R16WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R16WProc2 <- left_join(R16WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R16WProc2 <- R16WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R16EProc2[is.na(R16EProc2)] <- "....."
R16WProc2[is.na(R16WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R16EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route16East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R16WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route16West.xlsx", overwrite = T)



# write.csv(R16EProc2, "Route16East.csv", row.names = F)
# write.csv(R16WProc2, "Route16West.csv", row.names = F)

#####Route 16 End#####


#####Route 23#####


R23E <- TPCombinedDF %>%
  filter(Route == "23") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R23EProc <- R23E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R23EProcCt <- R23EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R23EProcCt <- colnames(R23EProcCt)[max.col(R23EProcCt,ties.method = "first")]

R23EProc <- R23EProc %>%
  arrange_at(R23EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R23EProc2 <- R23EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R23EProc2 <- left_join(R23EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R23EProc2 <- R23EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R23W <- TPCombinedDF %>%
  filter(Route == "23") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R23WProc <- R23W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R23WProcCt <- R23WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R23WProcCt <- colnames(R23WProcCt)[max.col(R23WProcCt,ties.method = "first")]

R23WProc <- R23WProc %>%
  arrange_at(R23WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R23WProc2 <- R23WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R23WProc2 <- left_join(R23WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R23WProc2 <- R23WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R23EProc2[is.na(R23EProc2)] <- "....."
R23WProc2[is.na(R23WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R23EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route23East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R23WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route23West.xlsx", overwrite = T)



# write.csv(R23EProc2, "Route23East.csv", row.names = F)
# write.csv(R23WProc2, "Route23West.csv", row.names = F)

#####Route 23 End#####


#####Route 24 & 18#####


R24E <- TPCombinedDF %>%
  filter(Route == "24" | Route == "18") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R24EProc <- R24E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R24EProcCt <- R24EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R24EProcCt <- colnames(R24EProcCt)[max.col(R24EProcCt,ties.method = "first")]

R24EProc <- R24EProc %>%
  arrange_at(R24EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R24EProc2 <- R24EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R24EProc2 <- left_join(R24EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R24EProc2 <- R24EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R24W <- TPCombinedDF %>%
  filter(Route == "24" | Route == "18") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R24WProc <- R24W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R24WProcCt <- R24WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R24WProcCt <- colnames(R24WProcCt)[max.col(R24WProcCt,ties.method = "first")]

R24WProc <- R24WProc %>%
  arrange_at(R24WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R24WProc2 <- R24WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R24WProc2 <- left_join(R24WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R24WProc2 <- R24WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R24EProc2[is.na(R24EProc2)] <- "....."
R24WProc2[is.na(R24WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R24EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route24East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R24WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route24West.xlsx", overwrite = T)



# write.csv(R24EProc2, "Route24East.csv", row.names = F)
# write.csv(R24WProc2, "Route24West.csv", row.names = F)

#####Route 24 & 18 End#####


#####Route 31#####


R31E <- TPCombinedDF %>%
  filter(Route == "31") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R31EProc <- R31E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R31EProcCt <- R31EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R31EProcCt <- colnames(R31EProcCt)[max.col(R31EProcCt,ties.method = "first")]

R31EProc <- R31EProc %>%
  arrange_at(R31EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R31EProc2 <- R31EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R31EProc2 <- left_join(R31EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R31EProc2 <- R31EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R31W <- TPCombinedDF %>%
  filter(Route == "31") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R31WProc <- R31W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R31WProcCt <- R31WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R31WProcCt <- colnames(R31WProcCt)[max.col(R31WProcCt,ties.method = "first")]

R31WProc <- R31WProc %>%
  arrange_at(R31WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R31WProc2 <- R31WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R31WProc2 <- left_join(R31WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R31WProc2 <- R31WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R31EProc2[is.na(R31EProc2)] <- "....."
R31WProc2[is.na(R31WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R31EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route31East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R31WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route31West.xlsx", overwrite = T)



# write.csv(R31EProc2, "Route31East.csv", row.names = F)
# write.csv(R31WProc2, "Route31West.csv", row.names = F)

#####Route 31 End#####


#####Route 40#####


R40E <- TPCombinedDF %>%
  filter(Route == "40") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R40EProc <- R40E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R40EProcCt <- R40EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R40EProcCt <- colnames(R40EProcCt)[max.col(R40EProcCt,ties.method = "first")]

R40EProc <- R40EProc %>%
  arrange_at(R40EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R40EProc2 <- R40EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R40EProc2 <- left_join(R40EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R40EProc2 <- R40EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R40W <- TPCombinedDF %>%
  filter(Route == "40") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R40WProc <- R40W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R40WProcCt <- R40WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R40WProcCt <- colnames(R40WProcCt)[max.col(R40WProcCt,ties.method = "first")]

R40WProc <- R40WProc %>%
  arrange_at(R40WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R40WProc2 <- R40WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R40WProc2 <- left_join(R40WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R40WProc2 <- R40WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R40EProc2[is.na(R40EProc2)] <- "....."
R40WProc2[is.na(R40WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R40EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route40East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R40WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route40West.xlsx", overwrite = T)



# write.csv(R40EProc2, "Route40East.csv", row.names = F)
# write.csv(R40WProc2, "Route40West.csv", row.names = F)

#####Route 40 End#####


#####Route 41#####


R41E <- TPCombinedDF %>%
  filter(Route == "41") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R41EProc <- R41E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R41EProcCt <- R41EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R41EProcCt <- colnames(R41EProcCt)[max.col(R41EProcCt,ties.method = "first")]

R41EProc <- R41EProc %>%
  arrange_at(R41EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R41EProc2 <- R41EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R41EProc2 <- left_join(R41EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R41EProc2 <- R41EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R41W <- TPCombinedDF %>%
  filter(Route == "41") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R41WProc <- R41W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R41WProcCt <- R41WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R41WProcCt <- colnames(R41WProcCt)[max.col(R41WProcCt,ties.method = "first")]

R41WProc <- R41WProc %>%
  arrange_at(R41WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R41WProc2 <- R41WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R41WProc2 <- left_join(R41WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R41WProc2 <- R41WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R41EProc2[is.na(R41EProc2)] <- "....."
R41WProc2[is.na(R41WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R41EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route41East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R41WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route41West.xlsx", overwrite = T)



# write.csv(R41EProc2, "Route41East.csv", row.names = F)
# write.csv(R41WProc2, "Route41West.csv", row.names = F)

#####Route 41 End#####


#####Route 42#####


R42E <- TPCombinedDF %>%
  filter(Route == "42") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R42EProc <- R42E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R42EProcCt <- R42EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R42EProcCt <- colnames(R42EProcCt)[max.col(R42EProcCt,ties.method = "first")]

R42EProc <- R42EProc %>%
  arrange_at(R42EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R42EProc2 <- R42EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R42EProc2 <- left_join(R42EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R42EProc2 <- R42EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R42W <- TPCombinedDF %>%
  filter(Route == "42") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R42WProc <- R42W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R42WProcCt <- R42WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R42WProcCt <- colnames(R42WProcCt)[max.col(R42WProcCt,ties.method = "first")]

R42WProc <- R42WProc %>%
  arrange_at(R42WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R42WProc2 <- R42WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R42WProc2 <- left_join(R42WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R42WProc2 <- R42WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R42EProc2[is.na(R42EProc2)] <- "....."
R42WProc2[is.na(R42WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R42EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route42East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R42WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route42West.xlsx", overwrite = T)



# write.csv(R42EProc2, "Route42East.csv", row.names = F)
# write.csv(R42WProc2, "Route42West.csv", row.names = F)

#####Route 42 End#####


#####Route 44#####


R44E <- TPCombinedDF %>%
  filter(Route == "44") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R44EProc <- R44E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R44EProcCt <- R44EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R44EProcCt <- colnames(R44EProcCt)[max.col(R44EProcCt,ties.method = "first")]

R44EProc <- R44EProc %>%
  arrange_at(R44EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R44EProc2 <- R44EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R44EProc2 <- left_join(R44EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R44EProc2 <- R44EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R44W <- TPCombinedDF %>%
  filter(Route == "44") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R44WProc <- R44W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R44WProcCt <- R44WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R44WProcCt <- colnames(R44WProcCt)[max.col(R44WProcCt,ties.method = "first")]

R44WProc <- R44WProc %>%
  arrange_at(R44WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R44WProc2 <- R44WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R44WProc2 <- left_join(R44WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R44WProc2 <- R44WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R44EProc2[is.na(R44EProc2)] <- "....."
R44WProc2[is.na(R44WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R44EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route44East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R44WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route44West.xlsx", overwrite = T)



# write.csv(R44EProc2, "Route44East.csv", row.names = F)
# write.csv(R44WProc2, "Route44West.csv", row.names = F)

#####Route 44 End#####


#####Route 51 & 52#####


R51E <- TPCombinedDF %>%
  filter(Route == "51" | Route == "52") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R51EProc <- R51E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R51EProcCt <- R51EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R51EProcCt <- colnames(R51EProcCt)[max.col(R51EProcCt,ties.method = "first")]

R51EProc <- R51EProc %>%
  arrange_at(R51EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R51EProc2 <- R51EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R51EProc2 <- left_join(R51EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R51EProc2 <- R51EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R51W <- TPCombinedDF %>%
  filter(Route == "51" | Route == "52") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R51WProc <- R51W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R51WProcCt <- R51WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R51WProcCt <- colnames(R51WProcCt)[max.col(R51WProcCt,ties.method = "first")]

R51WProc <- R51WProc %>%
  arrange_at(R51WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R51WProc2 <- R51WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R51WProc2 <- left_join(R51WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R51WProc2 <- R51WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R51EProc2[is.na(R51EProc2)] <- "....."
R51WProc2[is.na(R51WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R51EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route51East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R51WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route51West.xlsx", overwrite = T)



# write.csv(R51EProc2, "Route51East.csv", row.names = F)
# write.csv(R51WProc2, "Route51West.csv", row.names = F)

#####Route 51 & 52 End#####


#####Route 54#####


R54E <- TPCombinedDF %>%
  filter(Route == "54") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R54EProc <- R54E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R54EProcCt <- R54EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R54EProcCt <- colnames(R54EProcCt)[max.col(R54EProcCt,ties.method = "first")]

R54EProc <- R54EProc %>%
  arrange_at(R54EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R54EProc2 <- R54EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R54EProc2 <- left_join(R54EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R54EProc2 <- R54EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R54W <- TPCombinedDF %>%
  filter(Route == "54") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R54WProc <- R54W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R54WProcCt <- R54WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R54WProcCt <- colnames(R54WProcCt)[max.col(R54WProcCt,ties.method = "first")]

R54WProc <- R54WProc %>%
  arrange_at(R54WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R54WProc2 <- R54WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R54WProc2 <- left_join(R54WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R54WProc2 <- R54WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R54EProc2[is.na(R54EProc2)] <- "....."
R54WProc2[is.na(R54WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R54EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route54East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R54WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route54West.xlsx", overwrite = T)



# write.csv(R54EProc2, "Route54East.csv", row.names = F)
# write.csv(R54WProc2, "Route54West.csv", row.names = F)

#####Route 54 End#####


#####Route 60#####


R60E <- TPCombinedDF %>%
  filter(Route == "60") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R60EProc <- R60E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R60EProcCt <- R60EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R60EProcCt <- colnames(R60EProcCt)[max.col(R60EProcCt,ties.method = "first")]

R60EProc <- R60EProc %>%
  arrange_at(R60EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R60EProc2 <- R60EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R60EProc2 <- left_join(R60EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R60EProc2 <- R60EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R60W <- TPCombinedDF %>%
  filter(Route == "60") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R60WProc <- R60W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R60WProcCt <- R60WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R60WProcCt <- colnames(R60WProcCt)[max.col(R60WProcCt,ties.method = "first")]

R60WProc <- R60WProc %>%
  arrange_at(R60WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R60WProc2 <- R60WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R60WProc2 <- left_join(R60WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R60WProc2 <- R60WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R60EProc2[is.na(R60EProc2)] <- "....."
R60WProc2[is.na(R60WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R60EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route60East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R60WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route60West.xlsx", overwrite = T)



# write.csv(R60EProc2, "Route60East.csv", row.names = F)
# write.csv(R60WProc2, "Route60West.csv", row.names = F)

#####Route 60 End#####


#####Route 61#####


R61E <- TPCombinedDF %>%
  filter(Route == "61") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R61EProc <- R61E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R61EProcCt <- R61EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R61EProcCt <- colnames(R61EProcCt)[max.col(R61EProcCt,ties.method = "first")]

R61EProc <- R61EProc %>%
  arrange_at(R61EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R61EProc2 <- R61EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R61EProc2 <- left_join(R61EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R61EProc2 <- R61EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R61W <- TPCombinedDF %>%
  filter(Route == "61") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R61WProc <- R61W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R61WProcCt <- R61WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R61WProcCt <- colnames(R61WProcCt)[max.col(R61WProcCt,ties.method = "first")]

R61WProc <- R61WProc %>%
  arrange_at(R61WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R61WProc2 <- R61WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R61WProc2 <- left_join(R61WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R61WProc2 <- R61WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R61EProc2[is.na(R61EProc2)] <- "....."
R61WProc2[is.na(R61WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R61EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route61East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R61WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route61West.xlsx", overwrite = T)



# write.csv(R61EProc2, "Route61East.csv", row.names = F)
# write.csv(R61WProc2, "Route61West.csv", row.names = F)

#####Route 61 End#####


#####Route 66#####


R66E <- TPCombinedDF %>%
  filter(Route == "66") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R66EProc <- R66E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R66EProcCt <- R66EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R66EProcCt <- colnames(R66EProcCt)[max.col(R66EProcCt,ties.method = "first")]

R66EProc <- R66EProc %>%
  arrange_at(R66EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R66EProc2 <- R66EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R66EProc2 <- left_join(R66EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R66EProc2 <- R66EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R66W <- TPCombinedDF %>%
  filter(Route == "66") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R66WProc <- R66W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R66WProcCt <- R66WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R66WProcCt <- colnames(R66WProcCt)[max.col(R66WProcCt,ties.method = "first")]

R66WProc <- R66WProc %>%
  arrange_at(R66WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R66WProc2 <- R66WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R66WProc2 <- left_join(R66WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R66WProc2 <- R66WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R66EProc2[is.na(R66EProc2)] <- "....."
R66WProc2[is.na(R66WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R66EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route66East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R66WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route66West.xlsx", overwrite = T)



# write.csv(R66EProc2, "Route66East.csv", row.names = F)
# write.csv(R66WProc2, "Route66West.csv", row.names = F)

#####Route 66 End#####


#####Route 67#####


R67E <- TPCombinedDF %>%
  filter(Route == "67") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R67EProc <- R67E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R67EProcCt <- R67EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R67EProcCt <- colnames(R67EProcCt)[max.col(R67EProcCt,ties.method = "first")]

R67EProc <- R67EProc %>%
  arrange_at(R67EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R67EProc2 <- R67EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R67EProc2 <- left_join(R67EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R67EProc2 <- R67EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R67W <- TPCombinedDF %>%
  filter(Route == "67") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R67WProc <- R67W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R67WProcCt <- R67WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R67WProcCt <- colnames(R67WProcCt)[max.col(R67WProcCt,ties.method = "first")]

R67WProc <- R67WProc %>%
  arrange_at(R67WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R67WProc2 <- R67WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R67WProc2 <- left_join(R67WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R67WProc2 <- R67WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R67EProc2[is.na(R67EProc2)] <- "....."
R67WProc2[is.na(R67WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R67EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route67East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R67WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route67West.xlsx", overwrite = T)



# write.csv(R67EProc2, "Route67East.csv", row.names = F)
# write.csv(R67WProc2, "Route67West.csv", row.names = F)

#####Route 67 End#####


#####Route 69#####


R69E <- TPCombinedDF %>%
  filter(Route == "69") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R69EProc <- R69E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R69EProcCt <- R69EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R69EProcCt <- colnames(R69EProcCt)[max.col(R69EProcCt,ties.method = "first")]

R69EProc <- R69EProc %>%
  arrange_at(R69EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R69EProc2 <- R69EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R69EProc2 <- left_join(R69EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R69EProc2 <- R69EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R69W <- TPCombinedDF %>%
  filter(Route == "69") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R69WProc <- R69W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R69WProcCt <- R69WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R69WProcCt <- colnames(R69WProcCt)[max.col(R69WProcCt,ties.method = "first")]

R69WProc <- R69WProc %>%
  arrange_at(R69WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R69WProc2 <- R69WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R69WProc2 <- left_join(R69WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R69WProc2 <- R69WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R69EProc2[is.na(R69EProc2)] <- "....."
R69WProc2[is.na(R69WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R69EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route69East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R69WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route69West.xlsx", overwrite = T)



# write.csv(R69EProc2, "Route69East.csv", row.names = F)
# write.csv(R69WProc2, "Route69West.csv", row.names = F)

#####Route 69 End#####


#####Route 73#####


R73E <- TPCombinedDF %>%
  filter(Route == "73") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R73EProc <- R73E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R73EProcCt <- R73EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R73EProcCt <- colnames(R73EProcCt)[max.col(R73EProcCt,ties.method = "first")]

R73EProc <- R73EProc %>%
  arrange_at(R73EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R73EProc2 <- R73EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R73EProc2 <- left_join(R73EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R73EProc2 <- R73EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R73W <- TPCombinedDF %>%
  filter(Route == "73") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R73WProc <- R73W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R73WProcCt <- R73WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R73WProcCt <- colnames(R73WProcCt)[max.col(R73WProcCt,ties.method = "first")]

R73WProc <- R73WProc %>%
  arrange_at(R73WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R73WProc2 <- R73WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R73WProc2 <- left_join(R73WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R73WProc2 <- R73WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R73EProc2[is.na(R73EProc2)] <- "....."
R73WProc2[is.na(R73WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R73EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route73East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R73WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route73West.xlsx", overwrite = T)



# write.csv(R73EProc2, "Route73East.csv", row.names = F)
# write.csv(R73WProc2, "Route73West.csv", row.names = F)

#####Route 73 End#####


#####Route 80A#####


R80AE <- TPCombinedDF %>%
  filter(Route == "80A") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R80AEProc <- R80AE %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R80AEProcCt <- R80AEProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R80AEProcCt <- colnames(R80AEProcCt)[max.col(R80AEProcCt,ties.method = "first")]

R80AEProc <- R80AEProc %>%
  arrange_at(R80AEProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R80AEProc2 <- R80AEProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R80AEProc2 <- left_join(R80AEProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R80AEProc2 <- R80AEProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R80AW <- TPCombinedDF %>%
  filter(Route == "80A") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R80AWProc <- R80AW %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R80AWProcCt <- R80AWProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R80AWProcCt <- colnames(R80AWProcCt)[max.col(R80AWProcCt,ties.method = "first")]

R80AWProc <- R80AWProc %>%
  arrange_at(R80AWProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R80AWProc2 <- R80AWProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R80AWProc2 <- left_join(R80AWProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R80AWProc2 <- R80AWProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R80AEProc2[is.na(R80AEProc2)] <- "....."
R80AWProc2[is.na(R80AWProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R80AEProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route80AEast.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R80AWProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route80AWest.xlsx", overwrite = T)



# write.csv(R80AEProc2, "Route80AEast.csv", row.names = F)
# write.csv(R80AWProc2, "Route80AWest.csv", row.names = F)

#####Route 80A End#####


#####Route 81#####


R81E <- TPCombinedDF %>%
  filter(Route == "81") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R81EProc <- R81E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R81EProcCt <- R81EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R81EProcCt <- colnames(R81EProcCt)[max.col(R81EProcCt,ties.method = "first")]

R81EProc <- R81EProc %>%
  arrange_at(R81EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R81EProc2 <- R81EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R81EProc2 <- left_join(R81EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R81EProc2 <- R81EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R81W <- TPCombinedDF %>%
  filter(Route == "81") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R81WProc <- R81W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R81WProcCt <- R81WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R81WProcCt <- colnames(R81WProcCt)[max.col(R81WProcCt,ties.method = "first")]

R81WProc <- R81WProc %>%
  arrange_at(R81WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R81WProc2 <- R81WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R81WProc2 <- left_join(R81WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R81WProc2 <- R81WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R81EProc2[is.na(R81EProc2)] <- "....."
R81WProc2[is.na(R81WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R81EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route81East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R81WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route81West.xlsx", overwrite = T)



# write.csv(R81EProc2, "Route81East.csv", row.names = F)
# write.csv(R81WProc2, "Route81West.csv", row.names = F)

#####Route 81 End#####


#####Route 83#####


R83E <- TPCombinedDF %>%
  filter(Route == "83") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R83EProc <- R83E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R83EProcCt <- R83EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R83EProcCt <- colnames(R83EProcCt)[max.col(R83EProcCt,ties.method = "first")]

R83EProc <- R83EProc %>%
  arrange_at(R83EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R83EProc2 <- R83EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R83EProc2 <- left_join(R83EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R83EProc2 <- R83EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R83W <- TPCombinedDF %>%
  filter(Route == "83") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R83WProc <- R83W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R83WProcCt <- R83WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R83WProcCt <- colnames(R83WProcCt)[max.col(R83WProcCt,ties.method = "first")]

R83WProc <- R83WProc %>%
  arrange_at(R83WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R83WProc2 <- R83WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R83WProc2 <- left_join(R83WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R83WProc2 <- R83WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R83EProc2[is.na(R83EProc2)] <- "....."
R83WProc2[is.na(R83WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R83EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route83East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R83WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route83West.xlsx", overwrite = T)



# write.csv(R83EProc2, "Route83East.csv", row.names = F)
# write.csv(R83WProc2, "Route83West.csv", row.names = F)

#####Route 83 End#####


#####Route 84#####


R84E <- TPCombinedDF %>%
  filter(Route == "84") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R84EProc <- R84E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R84EProcCt <- R84EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R84EProcCt <- colnames(R84EProcCt)[max.col(R84EProcCt,ties.method = "first")]

R84EProc <- R84EProc %>%
  arrange_at(R84EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R84EProc2 <- R84EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R84EProc2 <- left_join(R84EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R84EProc2 <- R84EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R84W <- TPCombinedDF %>%
  filter(Route == "84") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R84WProc <- R84W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R84WProcCt <- R84WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R84WProcCt <- colnames(R84WProcCt)[max.col(R84WProcCt,ties.method = "first")]

R84WProc <- R84WProc %>%
  arrange_at(R84WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R84WProc2 <- R84WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R84WProc2 <- left_join(R84WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R84WProc2 <- R84WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R84EProc2[is.na(R84EProc2)] <- "....."
R84WProc2[is.na(R84WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R84EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route84East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R84WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route84West.xlsx", overwrite = T)



# write.csv(R84EProc2, "Route84East.csv", row.names = F)
# write.csv(R84WProc2, "Route84West.csv", row.names = F)

#####Route 84 End#####


#####Route 84A#####


R84AE <- TPCombinedDF %>%
  filter(Route == "84A") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R84AEProc <- R84AE %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R84AEProcCt <- R84AEProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R84AEProcCt <- colnames(R84AEProcCt)[max.col(R84AEProcCt,ties.method = "first")]

R84AEProc <- R84AEProc %>%
  arrange_at(R84AEProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R84AEProc2 <- R84AEProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R84AEProc2 <- left_join(R84AEProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R84AEProc2 <- R84AEProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R84AW <- TPCombinedDF %>%
  filter(Route == "84A") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R84AWProc <- R84AW %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R84AWProcCt <- R84AWProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R84AWProcCt <- colnames(R84AWProcCt)[max.col(R84AWProcCt,ties.method = "first")]

R84AWProc <- R84AWProc %>%
  arrange_at(R84AWProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R84AWProc2 <- R84AWProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R84AWProc2 <- left_join(R84AWProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R84AWProc2 <- R84AWProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R84AEProc2[is.na(R84AEProc2)] <- "....."
R84AWProc2[is.na(R84AWProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R84AEProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route84AEast.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R84AWProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route84AWest.xlsx", overwrite = T)



# write.csv(R84AEProc2, "Route84AEast.csv", row.names = F)
# write.csv(R84AWProc2, "Route84AWest.csv", row.names = F)

#####Route 84A End#####


#####Route 93#####


R93E <- TPCombinedDF %>%
  filter(Route == "93") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R93EProc <- R93E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R93EProcCt <- R93EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R93EProcCt <- colnames(R93EProcCt)[max.col(R93EProcCt,ties.method = "first")]

R93EProc <- R93EProc %>%
  arrange_at(R93EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R93EProc2 <- R93EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R93EProc2 <- left_join(R93EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R93EProc2 <- R93EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R93W <- TPCombinedDF %>%
  filter(Route == "93") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R93WProc <- R93W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R93WProcCt <- R93WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R93WProcCt <- colnames(R93WProcCt)[max.col(R93WProcCt,ties.method = "first")]

R93WProc <- R93WProc %>%
  arrange_at(R93WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R93WProc2 <- R93WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R93WProc2 <- left_join(R93WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R93WProc2 <- R93WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R93EProc2[is.na(R93EProc2)] <- "....."
R93WProc2[is.na(R93WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R93EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route93East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R93WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route93West.xlsx", overwrite = T)



# write.csv(R93EProc2, "Route93East.csv", row.names = F)
# write.csv(R93WProc2, "Route93West.csv", row.names = F)

#####Route 93 End#####


#####Route 94#####


R94E <- TPCombinedDF %>%
  filter(Route == "94") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R94EProc <- R94E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R94EProcCt <- R94EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R94EProcCt <- colnames(R94EProcCt)[max.col(R94EProcCt,ties.method = "first")]

R94EProc <- R94EProc %>%
  arrange_at(R94EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R94EProc2 <- R94EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R94EProc2 <- left_join(R94EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R94EProc2 <- R94EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R94W <- TPCombinedDF %>%
  filter(Route == "94") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R94WProc <- R94W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R94WProcCt <- R94WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R94WProcCt <- colnames(R94WProcCt)[max.col(R94WProcCt,ties.method = "first")]

R94WProc <- R94WProc %>%
  arrange_at(R94WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R94WProc2 <- R94WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R94WProc2 <- left_join(R94WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R94WProc2 <- R94WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R94EProc2[is.na(R94EProc2)] <- "....."
R94WProc2[is.na(R94WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R94EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route94East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R94WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route94West.xlsx", overwrite = T)



# write.csv(R94EProc2, "Route94East.csv", row.names = F)
# write.csv(R94WProc2, "Route94West.csv", row.names = F)

#####Route 94 End#####


#####Route 99#####


R99E <- TPCombinedDF %>%
  filter(Route == "99") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R99EProc <- R99E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R99EProcCt <- R99EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R99EProcCt <- colnames(R99EProcCt)[max.col(R99EProcCt,ties.method = "first")]

R99EProc <- R99EProc %>%
  arrange_at(R99EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R99EProc2 <- R99EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R99EProc2 <- left_join(R99EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R99EProc2 <- R99EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R99W <- TPCombinedDF %>%
  filter(Route == "99") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R99WProc <- R99W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R99WProcCt <- R99WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R99WProcCt <- colnames(R99WProcCt)[max.col(R99WProcCt,ties.method = "first")]

R99WProc <- R99WProc %>%
  arrange_at(R99WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R99WProc2 <- R99WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R99WProc2 <- left_join(R99WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R99WProc2 <- R99WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R99EProc2[is.na(R99EProc2)] <- "....."
R99WProc2[is.na(R99WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R99EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route99East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R99WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route99West.xlsx", overwrite = T)



# write.csv(R99EProc2, "Route99East.csv", row.names = F)
# write.csv(R99WProc2, "Route99West.csv", row.names = F)

#####Route 99 End#####


#####Route 411#####


R411E <- TPCombinedDF %>%
  filter(Route == "411") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R411EProc <- R411E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R411EProcCt <- R411EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R411EProcCt <- colnames(R411EProcCt)[max.col(R411EProcCt,ties.method = "first")]

R411EProc <- R411EProc %>%
  arrange_at(R411EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R411EProc2 <- R411EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R411EProc2 <- left_join(R411EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R411EProc2 <- R411EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R411W <- TPCombinedDF %>%
  filter(Route == "411") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R411WProc <- R411W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R411WProcCt <- R411WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R411WProcCt <- colnames(R411WProcCt)[max.col(R411WProcCt,ties.method = "first")]

R411WProc <- R411WProc %>%
  arrange_at(R411WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R411WProc2 <- R411WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R411WProc2 <- left_join(R411WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R411WProc2 <- R411WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R411EProc2[is.na(R411EProc2)] <- "....."
R411WProc2[is.na(R411WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R411EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route411East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R411WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route411West.xlsx", overwrite = T)



# write.csv(R411EProc2, "Route411East.csv", row.names = F)
# write.csv(R411WProc2, "Route411West.csv", row.names = F)

#####Route 411 End#####


#####Route 413#####


R413E <- TPCombinedDF %>%
  filter(Route == "413") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R413EProc <- R413E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R413EProcCt <- R413EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R413EProcCt <- colnames(R413EProcCt)[max.col(R413EProcCt,ties.method = "first")]

R413EProc <- R413EProc %>%
  arrange_at(R413EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R413EProc2 <- R413EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R413EProc2 <- left_join(R413EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R413EProc2 <- R413EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R413W <- TPCombinedDF %>%
  filter(Route == "413") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R413WProc <- R413W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R413WProcCt <- R413WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R413WProcCt <- colnames(R413WProcCt)[max.col(R413WProcCt,ties.method = "first")]

R413WProc <- R413WProc %>%
  arrange_at(R413WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R413WProc2 <- R413WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R413WProc2 <- left_join(R413WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R413WProc2 <- R413WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R413EProc2[is.na(R413EProc2)] <- "....."
R413WProc2[is.na(R413WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R413EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route413East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R413WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route413West.xlsx", overwrite = T)



# write.csv(R413EProc2, "Route413East.csv", row.names = F)
# write.csv(R413WProc2, "Route413West.csv", row.names = F)

#####Route 413 End#####


#####Route 432#####


R432E <- TPCombinedDF %>%
  filter(Route == "432") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R432EProc <- R432E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R432EProcCt <- R432EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R432EProcCt <- colnames(R432EProcCt)[max.col(R432EProcCt,ties.method = "first")]

R432EProc <- R432EProc %>%
  arrange_at(R432EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R432EProc2 <- R432EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R432EProc2 <- left_join(R432EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R432EProc2 <- R432EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R432W <- TPCombinedDF %>%
  filter(Route == "432") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R432WProc <- R432W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R432WProcCt <- R432WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R432WProcCt <- colnames(R432WProcCt)[max.col(R432WProcCt,ties.method = "first")]

R432WProc <- R432WProc %>%
  arrange_at(R432WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R432WProc2 <- R432WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R432WProc2 <- left_join(R432WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R432WProc2 <- R432WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R432EProc2[is.na(R432EProc2)] <- "....."
R432WProc2[is.na(R432WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R432EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route432East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R432WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route432West.xlsx", overwrite = T)



# write.csv(R432EProc2, "Route432East.csv", row.names = F)
# write.csv(R432WProc2, "Route432West.csv", row.names = F)

#####Route 432 End#####


#####Route 434#####


R434E <- TPCombinedDF %>%
  filter(Route == "434") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R434EProc <- R434E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R434EProcCt <- R434EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R434EProcCt <- colnames(R434EProcCt)[max.col(R434EProcCt,ties.method = "first")]

R434EProc <- R434EProc %>%
  arrange_at(R434EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R434EProc2 <- R434EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R434EProc2 <- left_join(R434EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R434EProc2 <- R434EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R434W <- TPCombinedDF %>%
  filter(Route == "434") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R434WProc <- R434W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R434WProcCt <- R434WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R434WProcCt <- colnames(R434WProcCt)[max.col(R434WProcCt,ties.method = "first")]

R434WProc <- R434WProc %>%
  arrange_at(R434WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R434WProc2 <- R434WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R434WProc2 <- left_join(R434WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R434WProc2 <- R434WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R434EProc2[is.na(R434EProc2)] <- "....."
R434WProc2[is.na(R434WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R434EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route434East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R434WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route434West.xlsx", overwrite = T)



# write.csv(R434EProc2, "Route434East.csv", row.names = F)
# write.csv(R434WProc2, "Route434West.csv", row.names = F)

#####Route 434 End#####


#####Route 501#####


R501E <- TPCombinedDF %>%
  filter(Route == "501") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R501EProc <- R501E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R501EProcCt <- R501EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R501EProcCt <- colnames(R501EProcCt)[max.col(R501EProcCt,ties.method = "first")]

R501EProc <- R501EProc %>%
  arrange_at(R501EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R501EProc2 <- R501EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R501EProc2 <- left_join(R501EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R501EProc2 <- R501EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R501W <- TPCombinedDF %>%
  filter(Route == "501") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R501WProc <- R501W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R501WProcCt <- R501WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R501WProcCt <- colnames(R501WProcCt)[max.col(R501WProcCt,ties.method = "first")]

R501WProc <- R501WProc %>%
  arrange_at(R501WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R501WProc2 <- R501WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R501WProc2 <- left_join(R501WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R501WProc2 <- R501WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R501EProc2[is.na(R501EProc2)] <- "....."
R501WProc2[is.na(R501WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R501EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route501East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R501WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route501West.xlsx", overwrite = T)



# write.csv(R501EProc2, "Route501East.csv", row.names = F)
# write.csv(R501WProc2, "Route501West.csv", row.names = F)

#####Route 501 End#####


#####Route 504#####


R504E <- TPCombinedDF %>%
  filter(Route == "504") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R504EProc <- R504E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R504EProcCt <- R504EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R504EProcCt <- colnames(R504EProcCt)[max.col(R504EProcCt,ties.method = "first")]

R504EProc <- R504EProc %>%
  arrange_at(R504EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R504EProc2 <- R504EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R504EProc2 <- left_join(R504EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R504EProc2 <- R504EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R504W <- TPCombinedDF %>%
  filter(Route == "504") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R504WProc <- R504W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R504WProcCt <- R504WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R504WProcCt <- colnames(R504WProcCt)[max.col(R504WProcCt,ties.method = "first")]

R504WProc <- R504WProc %>%
  arrange_at(R504WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R504WProc2 <- R504WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R504WProc2 <- left_join(R504WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R504WProc2 <- R504WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R504EProc2[is.na(R504EProc2)] <- "....."
R504WProc2[is.na(R504WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R504EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route504East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R504WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route504West.xlsx", overwrite = T)



# write.csv(R504EProc2, "Route504East.csv", row.names = F)
# write.csv(R504WProc2, "Route504West.csv", row.names = F)

#####Route 504 End#####


#####Route 671#####


R671E <- TPCombinedDF %>%
  filter(Route == "671") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R671EProc <- R671E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R671EProcCt <- R671EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R671EProcCt <- colnames(R671EProcCt)[max.col(R671EProcCt,ties.method = "first")]

R671EProc <- R671EProc %>%
  arrange_at(R671EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R671EProc2 <- R671EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R671EProc2 <- left_join(R671EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R671EProc2 <- R671EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R671W <- TPCombinedDF %>%
  filter(Route == "671") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R671WProc <- R671W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R671WProcCt <- R671WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R671WProcCt <- colnames(R671WProcCt)[max.col(R671WProcCt,ties.method = "first")]

R671WProc <- R671WProc %>%
  arrange_at(R671WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R671WProc2 <- R671WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R671WProc2 <- left_join(R671WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R671WProc2 <- R671WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R671EProc2[is.na(R671EProc2)] <- "....."
R671WProc2[is.na(R671WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R671EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route671East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R671WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route671West.xlsx", overwrite = T)



# write.csv(R671EProc2, "Route671East.csv", row.names = F)
# write.csv(R671WProc2, "Route671West.csv", row.names = F)

#####Route 671 End#####


#####Route 672#####


R672E <- TPCombinedDF %>%
  filter(Route == "672") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R672EProc <- R672E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R672EProcCt <- R672EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R672EProcCt <- colnames(R672EProcCt)[max.col(R672EProcCt,ties.method = "first")]

R672EProc <- R672EProc %>%
  arrange_at(R672EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R672EProc2 <- R672EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R672EProc2 <- left_join(R672EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R672EProc2 <- R672EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R672W <- TPCombinedDF %>%
  filter(Route == "672") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R672WProc <- R672W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R672WProcCt <- R672WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R672WProcCt <- colnames(R672WProcCt)[max.col(R672WProcCt,ties.method = "first")]

R672WProc <- R672WProc %>%
  arrange_at(R672WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R672WProc2 <- R672WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R672WProc2 <- left_join(R672WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R672WProc2 <- R672WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R672EProc2[is.na(R672EProc2)] <- "....."
R672WProc2[is.na(R672WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R672EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route672East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R672WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route672West.xlsx", overwrite = T)



# write.csv(R672EProc2, "Route672East.csv", row.names = F)
# write.csv(R672WProc2, "Route672West.csv", row.names = F)

#####Route 672 End#####


#####Route 673#####


R673E <- TPCombinedDF %>%
  filter(Route == "673") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R673EProc <- R673E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R673EProcCt <- R673EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R673EProcCt <- colnames(R673EProcCt)[max.col(R673EProcCt,ties.method = "first")]

R673EProc <- R673EProc %>%
  arrange_at(R673EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R673EProc2 <- R673EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R673EProc2 <- left_join(R673EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R673EProc2 <- R673EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R673W <- TPCombinedDF %>%
  filter(Route == "673") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R673WProc <- R673W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R673WProcCt <- R673WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R673WProcCt <- colnames(R673WProcCt)[max.col(R673WProcCt,ties.method = "first")]

R673WProc <- R673WProc %>%
  arrange_at(R673WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R673WProc2 <- R673WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R673WProc2 <- left_join(R673WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R673WProc2 <- R673WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R673EProc2[is.na(R673EProc2)] <- "....."
R673WProc2[is.na(R673WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R673EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route673East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R673WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route673West.xlsx", overwrite = T)



# write.csv(R673EProc2, "Route673East.csv", row.names = F)
# write.csv(R673WProc2, "Route673West.csv", row.names = F)

#####Route 673 End#####


#####Route 674#####


R674E <- TPCombinedDF %>%
  filter(Route == "674") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R674EProc <- R674E %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R674EProcCt <- R674EProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R674EProcCt <- colnames(R674EProcCt)[max.col(R674EProcCt,ties.method = "first")]

R674EProc <- R674EProc %>%
  arrange_at(R674EProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R674EProc2 <- R674EProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R674EProc2 <- left_join(R674EProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
R674EProc2 <- R674EProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

R674W <- TPCombinedDF %>%
  filter(Route == "674") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

R674WProc <- R674W %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


R674WProcCt <- R674WProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

R674WProcCt <- colnames(R674WProcCt)[max.col(R674WProcCt,ties.method = "first")]

R674WProc <- R674WProc %>%
  arrange_at(R674WProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

R674WProc2 <- R674WProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

R674WProc2 <- left_join(R674WProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
R674WProc2 <- R674WProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

R674EProc2[is.na(R674EProc2)] <- "....."
R674WProc2[is.na(R674WProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R674EProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route674East.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", R674WProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "Route674West.xlsx", overwrite = T)



# write.csv(R674EProc2, "Route674East.csv", row.names = F)
# write.csv(R674WProc2, "Route674West.csv", row.names = F)

#####Route 674 End#####


#####Route A#####


RAE <- TPCombinedDF %>%
  filter(Route == "A") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

RAEProc <- RAE %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


RAEProcCt <- RAEProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

RAEProcCt <- colnames(RAEProcCt)[max.col(RAEProcCt,ties.method = "first")]

RAEProc <- RAEProc %>%
  arrange_at(RAEProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

RAEProc2 <- RAEProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

RAEProc2 <- left_join(RAEProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
RAEProc2 <- RAEProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

RAW <- TPCombinedDF %>%
  filter(Route == "A") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

RAWProc <- RAW %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


RAWProcCt <- RAWProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

RAWProcCt <- colnames(RAWProcCt)[max.col(RAWProcCt,ties.method = "first")]

RAWProc <- RAWProc %>%
  arrange_at(RAWProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

RAWProc2 <- RAWProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

RAWProc2 <- left_join(RAWProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
RAWProc2 <- RAWProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

RAEProc2[is.na(RAEProc2)] <- "....."
RAWProc2[is.na(RAWProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", RAEProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "RouteAEast.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", RAWProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "RouteAWest.xlsx", overwrite = T)



# write.csv(RAEProc2, "RouteAEast.csv", row.names = F)
# write.csv(RAWProc2, "RouteAWest.csv", row.names = F)

#####Route A End#####


#####Route C#####


RCE <- TPCombinedDF %>%
  filter(Route == "C") %>%
  filter(Direction == "East") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

RCEProc <- RCE %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


RCEProcCt <- RCEProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

RCEProcCt <- colnames(RCEProcCt)[max.col(RCEProcCt,ties.method = "first")]

RCEProc <- RCEProc %>%
  arrange_at(RCEProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

RCEProc2 <- RCEProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

RCEProc2 <- left_join(RCEProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave)

#This should convert everything back
RCEProc2 <- RCEProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

#Westbound now

RCW <- TPCombinedDF %>%
  filter(Route == "C") %>%
  filter(Direction == "West") %>%
  select(trp_number_trim, Route, Direction, `Op Day`, StopInfo, DepTime24h, `Evt, Sta`) %>%
  unite(TrpOpDay, c(trp_number_trim, `Op Day`), sep = " - ", remove = T) %>%
  distinct ()

RCWProc <- RCW %>%
  ungroup () %>%
  pivot_wider(names_from = StopInfo, values_from = DepTime24h, values_fill = NA)


RCWProcCt <- RCWProc %>%
  summarize_at(vars(5:ncol(.)), function(x) sum(x > 0, na.rm = T))

RCWProcCt <- colnames(RCWProcCt)[max.col(RCWProcCt,ties.method = "first")]

RCWProc <- RCWProc %>%
  arrange_at(RCWProcCt)

#Rows finally sorted...now to convert it back to real time by joining with the original pared down table

RCWProc2 <- RCWProc %>%
  ungroup () %>%
  pivot_longer(cols = (5:ncol(.)),names_to = "StopInfo", values_to = "DepTime24h")

RCWProc2 <- left_join(RCWProc2, TPCombinedDF) %>%
  select(TrpOpDay, Route, Direction, `Evt, Sta`, StopInfo, DepartureSave) %>%
  distinct()

#This should convert everything back
RCWProc2 <- RCWProc2 %>%
  ungroup () %>%
  pivot_wider(names_from = "StopInfo", values_from = "DepartureSave", values_fill = NA)

RCEProc2[is.na(RCEProc2)] <- "....."
RCWProc2[is.na(RCWProc2)] <- "....."


#Export with format
wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", RCEProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "RouteCEast.xlsx", overwrite = T)

wb <- createWorkbook()
addWorksheet(wb, "sheet1")
writeData(wb, "sheet1", RCWProc2)
hs <- createStyle(halign = 'center')
addStyle(wb, "sheet1", hs, cols = 1:50, rows = 1:300, gridExpand = T)
saveWorkbook(wb, "RouteCWest.xlsx", overwrite = T)



# write.csv(RCEProc2, "RouteCEast.csv", row.names = F)
# write.csv(RCWProc2, "RouteCWest.csv", row.names = F)

#####Route C End#####