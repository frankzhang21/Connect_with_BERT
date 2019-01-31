library(readxl)

Atrr_table <- read_excel("H:/Project/Connect_with_BERT/Attr.xlsx")

# Parent Excel Binding ----------------------------------------------------


EXCEL_attr <- Atrr_table$Excel_Attr[!is.na(Atrr_table$Excel_Attr)]
Application_attr<- Atrr_table$Application_Attr[!is.na(Atrr_table$Application_Attr)]
Range_attr <- Atrr_table$Range_Attr[!is.na(Atrr_table$Range_Attr)]

EXCEL <- list(name=EXCEL_attr)
a <- purrr::transpose(EXCEL)
EXCEL <- purrr::set_names(a,EXCEL_attr)

EXCEL$Application <- list(name=Application_attr) %>% 
  purrr::transpose() %>% 
  purrr::set_names(Application_attr)

EXCEL$Application$get_Range <- list(name=Range_attr) %>% 
  purrr::transpose() %>% 
  purrr::set_names(Range_attr)



EXCEL$Application$get_Sheets()