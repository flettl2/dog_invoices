#Takes Word documents of dog walker's schedules
#Makes 2 csv files, 1 of combined invoices to send into acounting software QuickBooks
#The other calculates walkers pay.
#By Lucas Flett


library("qdapTools") #uses the read_docx function
library("stringr")
library("dplyr")
library("tidyr")


#Import all the walkers word docs from the week
doc_names <- list.files(pattern = "*.docx")
doc_list <- lapply(doc_names, read_docx)

#Finding out the date of invoice and who submitted it
pay_date <- as.Date(substr(doc_names,1,10)) #Always in YYYY-MM-DD form
walker <- sub("\\.docx", "", substr(doc_names,11, nchar(doc_names)))

#Find where clients are listed in the word doc (Will occur after first "Sunday")
filter_dogs <- function(invoice)
{
  fluff <- which(str_detect(invoice, "Sunday"))[1]
  invoice <- invoice[(fluff + 1):length(invoice)]  
  invoice <- str_subset(invoice, "^[^0123456789 ]") #removing times of day of walks
  invoice <- invoice[c(!str_detect(invoice, "cxl"))] #removing cancelled walks
  
  return(invoice)
}
doc_list <- lapply(doc_list, filter_dogs)

#Sort dogs by what type of walk they had
sort_walks <- function(invoice)
{
  invoice_excel <- as.data.frame(invoice, stringsAsFactors = FALSE)
  colnames(invoice_excel) <- "dummy" #Can't have colname = object name 
  invoice_excel <- separate(invoice_excel, dummy, c("Client", "Walk"), sep = "\\(", fill = "right")
  invoice_excel[is.na(invoice_excel)] <- "group)" #If walk type isn't specified it's a group
  
  #If the Client has 2 dogs, then charge them the double rate
  invoice_excel$Walk[str_detect(invoice_excel$Client, "/")] <- 
    str_c(invoice_excel$Walk[str_detect(invoice_excel$Client, "/")], "_double")
  
  #Make sure Walk types are labelled properly so that Quickbooks will understand them
  invoice_excel$Walk <- str_replace(invoice_excel$Walk, "\\)", "")
  
  #Remove duplicates and add a count
  invoice_excel <- summarise(group_by(invoice_excel, Client, Walk), length(Client)) 
  colnames(invoice_excel) <- c("Client", "Walk", "Count")
  invoice_excel <- as.data.frame(invoice_excel)
  invoice_excel <- spread(invoice_excel, key = Walk, value = Count)
  invoice_excel[is.na(invoice_excel)] <- 0
  
  
  return(invoice_excel)
}
doc_list <- lapply(doc_list, sort_walks)


#Calculates how much walkers are owed based how many walks they did and what the current rates are
walkers_pay <- function(invoice)
{
  #Find out how much each walker should be paid
  pricing <- data.frame(group = 13,
                      group_double = 21,
                      p.v. = 11,
                      "30" = 13.5,
                      "30_double" = 17.5,
                      "45" = 15,
                      "45_double" = 20,
                      Daycare = 22,
                      Daycare_double = 32,
                      check.names = FALSE)
  
  payment <- sum(apply(invoice[,2:ncol(invoice), drop = FALSE], 2, sum)*
               pricing[match(colnames(invoice[,-1, drop = FALSE]), colnames(pricing))])
  
  return(payment)
}

payment <- lapply(doc_list, walkers_pay)

#Organizing and export walkers pay
walker_df <- as.data.frame(walker, stringsAsFactors = FALSE)
colnames(walker_df) <- "Walker"

pay_date_df <- as.data.frame(pay_date)
colnames(pay_date_df) <- "Date"

payment_df <- t(as.data.frame(payment))
colnames(payment_df) <- "Payment"

payment_final <- cbind(pay_date_df, walker_df, payment_df)
rownames(payment_final) <- c()

write.table(payment_final, file = "Payment.csv", sep = ",", row.names = FALSE)

#Combining all invoices and exporting to csv file
invoice <- bind_rows(doc_list)
invoice[is.na(invoice)] <- 0
write.table(invoice, file = "Invoice.csv", sep = ",", row.names = FALSE)



