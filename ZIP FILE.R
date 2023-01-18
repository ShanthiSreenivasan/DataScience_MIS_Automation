rm(list= ls())
#setwd('Z:\\Amrish\\CIS\\Output')
Sys.setenv(TZ = "Pacific/Auckland")
setwd(file.path(Sys.getenv('CIS NEW LENDER BATCH'),'Output'))
today <- Sys.Date()
Sys.setenv('R_ZIPCMD' = 'C:/Rtools/bin/zip.exe')

library(stats) 
library(dplyr,warn.conflicts = FALSE)


Batch <- paste('.\\',today,'\\Batch to lender New cases',sep ='') 
Payments <- paste('.\\',today,'\\Payment file',sep ='')


zip_list <- list(Batch,Payments)


lapply(zip_list, function(x) {
  
 # browser()
  
  list <- list.files(x)
  list <-  list[list != 'Zip_File']
  
  lapply(list, function(y){
    
    #browser()
    if(!dir.exists(paste(x,'\\Zip_File',sep = ''))) dir.create(paste(x,'\\Zip_File',sep = ''))
    
    source <- paste(x,'\\',y,sep = '')
    dest <- stringr::str_replace(paste(x,'\\Zip_File\\',y,sep = ''),'.xlsx','.zip')
    
    pw <- function(y){
      
      password <- case_when(
        grepl('SCB Pay', y, ignore.case = T) ~ paste('creditmantri', format(Sys.Date(), '%m%y'), sep = ''),
        grepl('SCB New', y, ignore.case = T) ~ paste('creditmantri', format(Sys.Date(), '%m%y'), sep = ''),
        T ~ 'cm@1234'
      )
      
      return(password)
      
    }
    
    #zip::zip(dest,source,flags)
    
    zip(dest, files=source, flags = paste("-j -r9Xj -P", pw(y)))
    
    
    
    
  })
  
})

