xl.read.file2 = function (filename, header = TRUE, row.names = NULL, col.names = NULL, 
                          xl.sheet = NULL, top.left.cell = "A1", na = "", password = NULL,
                          write.res.password = NULL,
                          excel.visible = FALSE) 
{
  xl_temp = COMCreate("Excel.Application", existing = FALSE)
  on.exit(xl_temp$quit())
  xl_temp[["Visible"]] = excel.visible
  xl_temp[["DisplayAlerts"]] = FALSE
  if (isTRUE(grepl("^(http|ftp)s?://", filename))) {
    path = filename
  }
  else {
    path = normalizePath(filename, mustWork = TRUE)
  }
  passwords =paste(!is.null(password), !is.null(write.res.password), sep = "_") 
  xl_wb = switch(passwords, 
                 FALSE_FALSE = xl_temp[["Workbooks"]]$Open(path),
                 TRUE_FALSE = xl_temp[["Workbooks"]]$Open(path, 
                                                          password = password
                 ),
                 FALSE_TRUE = xl_temp[["Workbooks"]]$Open(path, 
                                                          writerespassword = write.res.password
                 ),
                 TRUE_TRUE = xl_temp[["Workbooks"]]$Open(path, 
                                                         password = password, 
                                                         writerespassword = write.res.password
                 )
                 
                 
                 
  )
  if (!is.null(xl.sheet)) {
    if (!is.character(xl.sheet) & !is.numeric(xl.sheet)) 
      stop('Argument "xl.sheet" should be character or numeric.')
    sh.count = xl_wb[["Sheets"]][["Count"]]
    sheets = sapply(seq_len(sh.count), function(sh) xl_wb[["Sheets"]][[sh]][["Name"]])
    if (is.numeric(xl.sheet)) {
      if (xl.sheet > length(sheets)) 
        stop("too large sheet number. In workbook only ", 
             length(sheets), " sheet(s).")
      xl_wb[["Sheets"]][[xl.sheet]]$Activate()
    }
    else {
      sheet_num = which(tolower(xl.sheet) == tolower(sheets))
      if (length(sheet_num) == 0) 
        stop("sheet ", xl.sheet, " doesn't exist.")
      xl_wb[["Sheets"]][[sheet_num]]$Activate()
    }
  }
  if (is.null(row.names) && is.null(col.names)) {
    if (header) {
      col.names = TRUE
      temp = excel.link:::xl.read.range(xl_temp[["ActiveSheet"]]$range(top.left.cell), 
                                        na = "")
      row.names = is.na(temp) || all(grepl("^([\\\\s\\\\t]+)$", 
                                           temp, perl = TRUE))
    }
    else {
      row.names = FALSE
      col.names = FALSE
    }
  }
  else {
    if (is.null(row.names)) 
      row.names = FALSE
    if (is.null(col.names)) 
      col.names = FALSE
  }
  top_left_corner = xl_temp$range(top.left.cell)
  xl.rng = top_left_corner[["CurrentRegion"]]
  if (tolower(top.left.cell) != "a1") {
    bottom_row = xl.rng[["row"]] + xl.rng[["rows"]][["count"]] - 
      1
    right_column = xl.rng[["column"]] + xl.rng[["columns"]][["count"]] - 
      1
    xl.rng = xl_temp$range(top_left_corner, xl_temp$cells(bottom_row, 
                                                          right_column))
  }
  excel.link:::xl.read.range(xl.rng, drop = FALSE, na = na, row.names = row.names, 
                             col.names = col.names)
}

#code timer beginning
start.time = Sys.time()

#load packages for stacking
library(tidyr)
library(DBI)
library(excel.link)
library(tibble)
library(stringr)

#location of the files that need to be uploaded
my_wd = "Y:/Accounting/Administration/Pete/R Test Files"

#year_directory 
real_file_directory = "Some File Directory Path"

#password to the files that need to be uploaded
file_password = "apass"
real_file_password = "apl"

#general location of the data on the sheet
data_location = "A1:AL200"

#file type extension
file_pattern = "*.xls" 

#first column name
firstcol = "Account"

#oracle connection
drv = dbDriver("Oracle")
host <- "ahost"
port <- 1521
sid <- "sid"

#oracle username
ora_user = "auser"

#oracle password
ora_password = "apass"

#months
month_vector = c("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC","TOTAL")


setwd(real_file_directory)
data.files = list.files(pattern = file_pattern)
data.files = data.files[-1]
total_data = list()
temp_df = data.frame()

#create file list
for (i in 1:length(data.files)){
  total_data[[i]] = xl.read.file2(data.files[i], header = FALSE, password = real_file_password, 
                                  top.left.cell = data_location, excel.visible = TRUE, write.res.password = real_file_password)
  total_data[[i]][total_data[[i]]==" "] = NA
}
total_data_copy = total_data

#name columns with correct month
count_df = 1
for (item in total_data){
  count = 1
  for (item2 in item){
    item_temp = month_vector %in% toupper(item2)
    true_location = which(item_temp == TRUE)
    if (length(true_location)>0){
      names(total_data_copy[[count_df]])[count]=month_vector[true_location]
    }
    count = count + 1
  }
  count_df = count_df + 1
}

#gather and add filename column
file_count = 1
for (item in total_data_copy){
  
  item = gather(item, Month, StatementValue, -a)
  #add a column for the file name
  item["filename"] = data.files[file_count]
  str(item)
  print(data.files[file_count])
  temp_df = rbind(temp_df, item)
  file_count = file_count + 1
}

#Cleaning Data
names(temp_df)[1] = "Account"
temp_df = subset(temp_df, Account != "NA")
temp_df = subset(temp_df, StatementValue != "NA")
temp_df$Account = trimws(temp_df$Account)
temp_df["asofdate"] = as.character(Sys.time())
rm(item)

#code time ending
end.time = Sys.time()
time.taken = end.time - start.time
time.taken = as.character(time.taken)
currentdate = as.character(Sys.time())
timestamp = c(time.taken, currentdate)

#appending run time to a text file in current working directory
sink("Code Run Times", append = TRUE)
print(timestamp)
sink()

#####################SQL###########################################
library(ROracle)
library(dbConnect)

connect.string <- paste(
  "(DESCRIPTION=",
  "(ADDRESS=(PROTOCOL=tcp)(HOST=", host, ")(PORT=", port, "))",
  "(CONNECT_DATA=(SID=", sid, ")))", sep = "")


con = dbConnect(drv, username = ora_user, password = ora_password,dbname = connect.string)

#create table with only the most recently uploaded data
dbWriteTable(con, "tbl_accounting_current", temp_df, overwrite = TRUE, append = FALSE)

dbDisconnect(con)
con = dbConnect(drv, username = ora_user, password = ora_password,dbname = connect.string)

#create table that archive all data uploaded each time the script is run
dbWriteTable(con, "tbl_accounting_archive", temp_df, overwrite = FALSE, append = TRUE)

dbDisconnect(con)
con = dbConnect(drv, username = ora_user, password = ora_password,dbname = connect.string)

#create table that archive all data uploaded each time the script is run
dbWriteTable(con, "tbl_accounting_historical", temp_df, overwrite = FALSE, append = TRUE)

dbDisconnect(con)
con = dbConnect(drv, username = ora_user, password = ora_password,dbname = connect.string)

#delete repetitive entries from historical
HistVal = dbSendQuery(con, "DELETE FROM \"tbl_accounting_historical\"
                      WHERE \"Month\" IN (SELECT DISTINCT \"Month\" FROM \"tbl_accounting_current\")
                      AND \"filename\" IN (SELECT DISTINCT \"filename\" FROM \"tbl_accounting_current\")")


#add the most recent current directory to historical
HistAdd = dbSendQuery(con, "INSERT INTO \"tbl_accounting_current\" SELECT * FROM \"peter_test\"")
######################################################