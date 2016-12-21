
# ------------------------------- Compare directory listings and send e-mail with differences ---------------------------#

# get the path for the R executable
# file.path(R.home(), "bin", "R")
# in my case it was "C:/PROGRA~1/R/R-31~1.0/bin/R"

# function to check if a package is installed
is_installed <- function(mypkg) is.element(mypkg, installed.packages()[,1]) 
#----

# The "RDCOMClient" is needed for sending e-mail
# Check to see that it's installed and if not, install it.

# However, the RDCOMClient package isn't at the cmu repository
# and so it has to be loaded from the menu

if(!is_installed("RDCOMClient"))  
{  
  install.packages("RDCOMClient",repos="http://lib.stat.cmu.edu/R/CRAN")  
}  
library("RDCOMClient",character.only=TRUE,quietly=TRUE,verbose=FALSE)  

# Set the working directory
# I mapped "\\dcnsbiona01b\EDC_V1_SHR2\Shared\DATA_WCTS_ARP_SABS" to Z:\
# R can't seem to have a network drive as working directory unless
# it's got a drive letternnn 


# Check to see whether Outlook is open and if it is
# Quit

#setwd("Z:")
setwd("C:/Temp") # Change as necessary
#getwd()

# Get username and create an e-mail address

email <- paste(Sys.getenv("USERNAME"),"@dfo-mpo.gc.ca", sep = "")

# -------- function to send e-mail --------#
email_fn <- function(email, subject, attachment) {
  OutApp <- COMCreate("Outlook.Application")
  ## create an email 
  outMail = OutApp$CreateItem(0)
  ## configure  email parameter 
  outMail[["To"]] = email
  outMail[["subject"]] = subject
  #outMail[["body"]] = ""
  outMail[["Attachments"]]$Add(attachment)
  ## send it                     
  outMail$Send()
}
# -------- end function --------#


fileold <- "dirold.txt"
filenew <- "dirnew.txt"
fileall <- "dirfull.txt"



#If the original directory listing doesn't exist, create it
if (!file.exists(fileold)) {
  system("cmd.exe /c dir /b/s > dirold.txt")
}

if (!file.exists(fileall)) {
  system("cmd.exe /c copy dirold.txt dirfull.txt")
}

# Create a new file listing (bare format, check all sub-directories)
system("cmd.exe /c dir /b/s > dirnew.txt")

# The mail function is working even when Outlook isn't running.  However, it is
# possible to test if Outlook is running and start it if it isn't (see tasklist command below)

# start Outlook if it isn't running
#system("cmd.exe /c tasklist /FI \"IMAGENAME eq outlook.exe\" | find /I /N \"outlook.exe\" || \"C:/Program Files (x86)/Microsoft Office/Office14/OUTLOOK.EXE\"")

# Scan the contents of the two directory listings into R
# Compare both directory listings and store
# any differences are stored as 'filesadded' or 'filesremoved'.  
# If either of these has 
# values in them, write them to a new file and send 
# file via e-mail.
# This doesn't error check the e-mail address
dirlistnew <- scan(filenew, what="", sep="\n", quote = "\"")
dirlistold <- scan(fileold, what="", sep="\n", quote = "\"")
dirlistfull <- scan(fileall, what="", sep="\n", quote = "\"")
filesadded <- setdiff(dirlistnew, dirlistold)
filesremoved <- setdiff(dirlistfull, dirlistnew)
if (length(filesadded)>0) {
  # Why does this command write an 'x' to the first line of the file???
  write.csv(filesadded, "NewFiles.csv", row.names=FALSE, col.names = FALSE)
  # Send e-mail with attachment
  email_fn(email, "More new files", "C:/Temp/NewFiles.csv")
  write.table(filesadded, fileall, append = TRUE, row.names=FALSE, col.names = FALSE)
}

if (length(filesremoved)!=0) {
  # Why does this command write an 'x' to the first line of the file???
  write.csv(filesremoved, "RemovedFiles.csv", row.names=FALSE, col.names = FALSE)
  # Send e-mail with attachment
  email_fn(email, "Some files were removed", "C:/Temp/RemovedFiles.csv")
}
# delete the old directory listing and copy new to old
system("cmd.exe /c del dirold.txt")
system("cmd.exe /c rename dirnew.txt dirold.txt")


# ---------Clean up---------------#
# rm(list = ls())