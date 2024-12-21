print("starting")
library(data.table)
getwd()
file_loc = commandArgs(trailingOnly = T)[1]
new_wd = dirname(file_loc)
print("new wd will be")
print(new_wd)
setwd(new_wd)

x = fread("02-dataincsv.csv") 
y = x*2

write.table(y, file = "04-modifieddataincsv.csv", sep = ",", col.names = F, row.names = F)
print("output written to csv")
