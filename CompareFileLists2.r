
# ------------------------------- Compare directory listings and send e-mail with differences ---------------------------#

# get the path for the R executable
# file.path(R.home(), "bin", "R")
# in my case it was "C:/PROGRA~1/R/R-31~1.0/bin/R"

# function to check if a package is installed
is_installed <- function(mypkg) is.element(mypkg, installed.packages()[,1]) 

# The "RDCOMClient" is needed for sending e-mail
# Check to see that it's installed and if not, install it

if(!is_installed("RDCOMClient"))  
{  
  install.packages("RDCOMClient",repos="http://lib.stat.cmu.edu/R/CRAN")  
}  
library("RDCOMClient",character.only=TRUE,quietly=TRUE,verbose=FALSE)  

# Set the working directory
# I mapped "\\dcnsbiona01b\EDC_V1_SHR2\Shared\DATA_WCTS_ARP_SABS" to Z:\
# R can't seem to have a network drive as working directory unless
# it's got a drive letter
#setwd("Z:")
setwd("C:/Temp") # Change as necessary
#getwd()

# Get username and create an e-mail address

email <- paste(Sys.getenv("USERNAME"),"@dfo-mpo.gc.ca", sep = "")

# -------- function to send e-mail --------#
email_fn <- function(email, subject) {
  OutApp <- COMCreate("Outlook.Application")
  ## create an email 
  outMail = OutApp$CreateItem(0)
  ## configure  email parameter 
  outMail[["To"]] = email
  outMail[["subject"]] = subject
  #outMail[["body"]] = ""
  outMail[["Attachments"]]$Add("C:/Temp/NewFiles.csv")
  ## send it                     
  outMail$Send()
}
# -------- end function --------#


fileold <- "dirold.txt"
filenew <- "dirnew.txt"

#If the original directory listing doesn't exist, create it
if (!file.exists(fileold)) {
  system("cmd.exe /c dir /b/s > dirold.txt")
}

# Create a new file listing (bare format, check all sub-directories)
system("cmd.exe /c dir /b/s > dirnew.txt")


# Compare both directory listings and store
# any differences are stored as 'diff'.  If diff has
# values in it, write them to a new file and send 
# file via e-mail.
# This doesn't error check the e-mail address
dirlistnew <- scan(filenew, what="", sep="\n", quote = "\"")
dirlistold <- scan(fileold, what="", sep="\n", quote = "\"")
diff <- setdiff(dirlistnew, dirlistold)
if (length(diff)>0) {
  # Why does this command write an 'x' to the first line of the file???
  write.csv(diff, "NewFiles.csv", row.names=FALSE, col.names = FALSE)
  # Send e-mail with attachment
  email_fn(email, "More new files")
  write.table(diff, fileold, append = TRUE, row.names=FALSE, col.names = FALSE)
}



# ---------Clean up---------------#
# rm(list = ls())