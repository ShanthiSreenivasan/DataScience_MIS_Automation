

FileCreate <- function(dataset,sheet_name,Filename,sheet_name2 = NULL,
                       sheet_name3 = NULL, dataset2 = NULL, dataset3 = NULL){
  
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, sheet_name)
  hs1 <- createStyle(fgFill = "#4F81BD", 
                     halign = "CENTER", 
                     textDecoration = "Bold",
                     border = "Bottom", 
                     fontColour = "white")
  setColWidths(wb, sheet_name,cols = 1:ncol(dataset),widths = 'auto')
  writeData(wb,sheet_name,dataset,
            headerStyle = hs1,
            borders = 'all' )
  if(!is.null(dataset2)){
    openxlsx::addWorksheet(wb, sheet_name2)
    setColWidths(wb, sheet_name2,cols = 1:ncol(dataset2),widths = 'auto')
    writeData(wb,sheet_name2,dataset2,
              headerStyle = hs1,
              borders = 'all')  
    
  }
  
  if(!is.null(dataset3)){
    openxlsx::addWorksheet(wb, sheet_name3)
    setColWidths(wb, sheet_name3,cols = 1:ncol(dataset3),widths = 'auto')
    writeData(wb,sheet_name3,dataset3,
              headerStyle = hs1,
              borders = 'all')  
    
  }
  
  openxlsx::saveWorkbook(wb = wb,file = Filename,overwrite = T)
  
  
}


#Zip File ----------------------------------------------------------------------
Sys.setenv('R_ZIPCMD' = 'C:/Rtools/bin/zip.exe')

zip.file <- function(source_path){

zip_list1 <- list(source_path)
  
  
  lapply(zip_list1, function(x) {
    
    #browser()
    
    list <- list.files(x)
    list <-  list[list != 'Zip_File']
    
    lapply(list, function(y){
      
      #browser()
      if(!dir.exists(paste(x,'\\Zip_File',sep = ''))) dir.create(paste(x,'\\Zip_File',sep = ''))
      
      source <- paste(x,'\\',y,sep = '')
      dest <- stringr::str_replace(paste(x,'\\Zip_File\\',y,sep = ''),'.xlsx','.zip')
      
      pw <- function(y){
        
        password <- case_when(
          #grepl('SCB Payments', y, ignore.case = T) ~ paste('creditmantri', format(Sys.Date(), '%m%y'), sep = ''),
          grepl('SCB Immediate', y, ignore.case = T) ~ paste('creditmantri', format(Sys.Date(), '%m%y'), sep = ''),
          T ~ 'cm@1234'
        )
        
        return(password)
        
      }
      
      zip(dest, source, flags = paste("-j -r9Xj -P", pw(y)))
      
      
      
      
    })
  })
}
