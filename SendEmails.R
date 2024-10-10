# Sending emails to people who have sold work at an OAS Exhibition
# Requires an Excel spreadsheet with the columns FirstName, Price, Due, and Email
# Price is the sale price, Due is the amount to be sent to the artist
# Other columns can be present but are not used here

library(gmailr)
library(readxl)
library(tidyverse)
library(mailmerge)

# rappdirs::user_data_dir("gmailr")
# gm_auth_configure()
# gm_oauth_client()
# 
# withr::with_envvar(
#   c(GMAILR_EMAIL = "dnbarron@gmail.com"),
#   gm_default_email()
# )

gm_auth()
gm_profile()

# test_email <- 
#   gm_mime() |>
#   gm_to("david.barron@jesus.ox.ac.uk") |>
#   gm_from("dnbarron@gmail.com") |>
#   gm_subject("this is just a gmailr test") |>
#   gm_text_body("Can you hear me now?")
# 
# d <- gm_create_draft(test_email)
# gm_send_draft(d)


setwd("C:/Users/dnbar/Dropbox/OAS") # Excelfile has to be in this directory

dta <- read_xlsx("Sales.xlsx")

msg <- '
---
  subject: "OAS exhibition sale"  
---

Dear {FirstName},

I understand that you made a sale during the recent OAS Open Exhibition. 
Congratulations!  The item sold for £{Price}, so, after commission, you are due
£{Due} from OAS.  So I can pay this to you, could you please let me have your 
bank details.  Thanks very much.

Best wishes,  
David Barron  
OAS Treasurer
'

 dta %>% mail_merge(msg, 
                   to_col = "Email",
                   send = "draft")
