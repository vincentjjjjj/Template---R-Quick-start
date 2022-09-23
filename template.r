---
title: ""
author: "Vincent"
date: ""
output: html_document
---

```{r}
# options(encoding = "UTF-8", warn = -1, "openxlsx.dateFormat" = "dd/mm/yyyy", xlsx.date.format = "dd/mm/yyyy", DT.warn.size = FALSE)
options(encoding = "UTF-8", "openxlsx.dateFormat" = "dd/mm/yyyy", xlsx.date.format = "dd/mm/yyyy", DT.warn.size = FALSE)


if (F) {
  input  = basename(rstudioapi::getSourceEditorContext()$path)
  output = gsub(".Rmd", ".R", input)
  knitr::purl(input, output = output)
  writeLines(as.character(parse(output)), output)
  
  # output_bat = gsub(".Rmd", ".bat", input)
  # writeLines(paste0("@echo off\nRScript.exe --encoding=UTF-8 ", output, "\npause\n"), output_bat)
}


if (!"pacman" %in% rownames(installed.packages())) {
  print("Installation du programme en cours...")
  user_libs = Sys.getenv("R_LIBS_USER")
  dir.create(user_libs, showWarnings = F, recursive = T)
  install.packages(c("pacman", "hms"), lib = user_libs, repos = "https://cran.rstudio.com/", quiet = T, verbose = F)
  shell("cls")
  stop("Installation du programme terminé. Relancer le programme.")
}


pacman::p_load(
  stringi,
  fs,
  plyr,
  readxl,
  writexl,
  lubridate,
  tidyverse
)


# write_xlsx <- function(df, filename, sheetName = "Sheet 1") {
#   cols_date   = which(map(df, is.Date) == T)
#   cols_period = which(map(df, is.period) == T)
#   all_rows = 2:nrow(df)
#   all_cols = 1:ncol(df)
#   
#   df = df %>% ungroup()
#   df = df %>% mutate(across(all_of(cols_period), ~ .x %>% hms::hms() %>% str_sub(end = -4)))
#   wb = createWorkbook()
#   wb %>% addWorksheet(sheetName)
#   # time_style = createStyle(numFmt = "TIME")
#   # wb %>% addStyle(1, style = time_style, rows = all_rows, cols = cols_period, gridExpand = T)
#   
#   wb %>% setColWidths(sheet = 1, cols = all_cols, widths = "auto")
#   wb %>% setColWidths(sheet = 1, cols = cols_date, widths = "12")
#   # wb %>% addFilter(1, rows = 1, cols = all_cols)
#   wb %>% writeData(df, sheet = sheetName, withFilter = T)
#   wb %>% saveWorkbook(file = filename, overwrite = T)
# }


write_xlsx <- function(df, filename) {
  filename %>% path_dir() %>% dir_create()
  filename = filename %>% path_ext_set("xlsx")
  df = df %>% ungroup()
  cols_date   = which(map(df, is.Date) == T)
  cols_period = which(map(df, is.period) == T)
  
  df = df %>% mutate(across(all_of(cols_period), ~ .x %>% hms::hms() %>% str_sub(end = -4)))
  df = df %>% mutate(across(everything(), as.character))
  df %>% as.data.frame() %>% writexl::write_xlsx(filename)
}


write_csv <- function(df, filename) {
  filename %>% path_dir() %>% dir_create()
  filename = filename %>% path_ext_set("csv")
  df = df %>% ungroup()
  cols_date   = which(map(df, is.Date) == T)
  cols_period = which(map(df, is.period) == T)
  df = df %>% mutate(across(all_of(cols_period), ~ .x %>% hms::hms() %>% str_sub(end = -4)))
  df = df %>% mutate(across(everything(), as.character))
  df %>% write_csv(filename)
}


clean_colnames <- function(.tbl) {
  .tbl %>% rename_all(~ .x %>% str_to_lower() %>% str_replace_all(" |/", "_") %>% stri_trans_general("Latin-ASCII") %>% str_remove_all("[^a-zA-Z0-9_]"))
}
```


```{r}


```
