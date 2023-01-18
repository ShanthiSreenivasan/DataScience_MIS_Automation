#Create File
FileCreation <- function(NewLeads, FilePath){
  
  wb <- createWorkbook()
  addWorksheet(wb,'Sheet1')
  hs1 <- createStyle(fgFill = "#4F81BD", 
                     halign = "CENTER", 
                     textDecoration = "Bold",
                     border = "Bottom", 
                     fontColour = "white")
  setColWidths(wb,'Sheet1',cols = 1:ncol(NewLeads), widths = 'auto')
  setColWidths(wb,'Sheet1',cols = 1, widths = 10)
  writeData(wb, "Sheet1", x = NewLeads ,headerStyle = hs1, borders = "all")
  saveWorkbook(wb,FilePath,overwrite = T)
  
}


#Mail Function ----------------------------------------
mail_function <- function(sender, recipients, cc_recipients = NULL,bcc_recipients = NULL ,message, email_body,filename= NULL){
  
  send.mail(from = sender,
            to = recipients,
            cc =cc_recipients,
            bcc = bcc_recipients,
            subject = message,
            html = TRUE,
            inline = T,
            body = str_interp(email_body),
            smtp = list(host.name = "email-smtp.us-east-1.amazonaws.com", port = 587,
                        user.name = "AKIAXJKXCWJGAXCHHIGY",
                        passwd = "BPBUWKtOcH6emQE3X1uE49fn1fnRm0LqvIlYUoq59wz4" , ssl = TRUE),
            authenticate = TRUE,
            attach.files =filename,
            send = TRUE)
}