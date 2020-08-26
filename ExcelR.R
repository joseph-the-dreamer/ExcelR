#Package: ExcelR
#Type: Package
#Title: Recreates Microsoft Excel/ Google Sheets functions
#Version: 0.1.0
#Author: Marcus Joseph
#Description: This package recreates the functionality and syntax of Microsoft Excel in an R environment
#License: MIT
#Imports: dplyr, stringr, lubridate, stats

install.packages('stringr')
install.packages('dplyr')
install.packages('lubridate')
install.packages('stats')
library(stats)
library(lubridate)
library(stringr)
library(dplyr)


#IF - Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE
excel_if <- function(logical_test, value_if_true, value_if_false){
  if (is.atomic(logical_test)) {
    if (typeof(logical_test) != "logical") 
      storage.mode(logical_test) <- "logical"
    if (length(logical_test) == 1 && is.null(attributes(logical_test))) {
      if (is.na(logical_test)) 
        return(NA)
      else if (logical_test) {
        if (length(value_if_true) == 1) {
          yat <- attributes(value_if_true)
          if (is.null(yat) || (is.function(value_if_true) && identical(names(yat), 
                                                                       "srcref"))) 
            return(value_if_true)
        }
      }
      else if (length(value_if_false) == 1) {
        nat <- attributes(value_if_false)
        if (is.null(nat) || (is.function(value_if_false) && identical(names(nat), 
                                                                      "srcref"))) 
          return(value_if_false)
      }
    }
  }else test <- if (isS4(logical_test)) 
    methods::as(logical_test, "logical")
  else as.logical(logical_test)
  ans <- logical_test
  ok <- !is.na(logical_test)
  if (any(logical_test[ok])) 
    ans[logical_test & ok] <- rep(value_if_true, length.out = length(ans))[logical_test & 
                                                                             ok]
  if (any(!test[ok])) 
    ans[!logical_test & ok] <- rep(value_if_false, length.out = length(ans))[!logical_test & 
                                                                               ok]
  ans
  
} 

#IFERROR - Returns value_if_error if expression is an error and the value of the expression itself otherwise
excel_iferror <- function(value, value_if_error){
  ifelse(class(try(value, silent = T)) == "try-error",value_if_error,value)
}

#LEN - Returns the number of characters in a text string
excel_len <- function (text){
  str_length(text)
}

#AVERAGE - Returns the average (arithmetic mean) of its arguments, which can be numbers or names, arrays, or references that contain numbers
excel_average <- function (number1, ...){
  ifelse(class(try(mean(c(number1,...)), silent = T)) == "try-error","Input must be a number",mean(c(number1,...)))
}

#ROUND UP - Rounds a number up, away from zero
excel_roundup <- function (number){
  ifelse(class(try(ceiling(number), silent = T)) == "try-error","Input must be a number",ceiling(number))
}

#ROUND DOWN - Rounds a number down, toward zero
excel_rounddown <- function (number) {
  ifelse(class(try(floor(number), silent = T)) == "try-error","Input must be a number",floor(number))
}

#LEFT - Returns the specified number of characters from the start of a text string
excel_left <- function(text, num_chars) {
  substr(text, 1, num_chars)
}

#MID - Returns the characters from the middle of a text string, given a starting position and length
excel_mid <- function(text, start_num, num_chars) {
  substr(text, start_num, start_num + num_chars - 1)
}

#RIGHT - Returns the specified number of characters from the end of a text string
excel_right <- function(text, num_chars) {
  substr(text, nchar(text) - (num_chars-1), nchar(text))
}

#COUNT - example count(table$column) - Counts the number of cells in a range that contain numbers
excel_count <- function(table_and_col_name){
  length(table_and_col_name)
}

#VLOOKUP - example vlookup("cell text value referenced", referencedTable$column, returnTable$column) - Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify, By default, the table must be sorted in an ascending order
excel_vlookup <- function(lookup_value, lookupTable_and_col_name, returnTable_and_col_name){
row_num <- min(which(grepl(lookup_value,lookupTable_and_col_name)))
returnTable_and_col_name[row_num]
}

#REMOVE DUPLICATES - complete - Removes duplicate values within a specified range
excel_remove_duplicates <- function(table_and_col_name){
  table = left(table_and_col_name, gregexpr("$", table_and_col_name)[[1]]-1)
  column = right(table_and_col_name,str_length(table_and_col_name)-len(table))
  distinct(table, column, .keep_all = TRUE)
}

#SEARCH - Returns the number of the character at which a specific character or text string is first found, reading left to right (not case-sensitive)
excel_search <- function(find_text,within_text){
  return <- gregexpr(find_text, within_text)[[1]]
  ifelse(return[1]<0,0,return[1])
} 

#FIND - Returns the starting position of one text string within another text string. FIND is case-sensitive
excel_find <- function(find_text,within_text){
  return <- gregexpr(tolower(find_text), tolower(within_text))[[1]]
  ifelse(return[1]<0,0,return[1])
}

#LOWER - Converts all letters in a text string to lowercase 
excel_lower <- function(text){
  tolower(text)
}

#UPPER - Converts a text string to all uppercase letters
excel_upper <- function(text){
  toupper(text)
}

#RANDBETWEEN - Returns a random number between the numbers you specify
excel_randbetween <- function (bottom, top){
  ifelse(is.numeric(bottom)==TRUE&is.numeric(top)==TRUE,sample(bottom:top, 1),"Please only input numeric values")
}

#SUM - Adds all the numbers in a range of cells
excel_sum <- function (number1,...){
  sum(number1,...)
}

#MONTH - Returns the month, a number from 1 (January) to 12 (December)
excel_month <- function(date_value){
  ifelse(class(date_value)=="Date", 
         format(date_value, "%m"),
         "Input for must be a date")
}

#TRIM - Removes all spaces from a text string except for single spaces between words
excel_trim <- function(text){
  trimws(text)
}

#PROPER
excel_proper <- function(value){
  sub("(.)", ("\\U\\1"), tolower(value), pe=TRUE)
}

#ISNUMBER - Checks whether a value is a number, and returns TRUE or FALSE
excel_isnumber <- function(value){
  is.numeric(value)
}

#ISTEXT - Checks whether a value is text, and returns TRUE or FALSE
excel_istext <- function(value){
  is.character(value)
}

#CONCATENATE - Joins several text strings into one text string
excel_concatenate <- function(text1,...){
  input <- c(text1,...)
  paste(input,sep = "",collapse = "")
}

#COUNTIF - Counts the number of cells within a range that meet the given condition
excel_countif <- function(range, criteria){
  sum(range == criteria, na.rm=TRUE)
}
#RAND - Returns a random number greater than or equal to 0 and less than 1
excel_rand <- function(){
  (sample(1:10, 1000, replace = TRUE)/10)[1]
}

#TODAY - Returns the current date formatted as a date
excel_today <- function(){
  lubridate::today()
}

#MAX - Returns the largest value in a set of values. Ignores logical values and text
excel_max <- function(number1, ...){
  max(number1, ...)
}

#MIN - Returns the smallest number in a set of values. Ignores logical values and text
excel_min <- function(number1, ...){
  min(number1, ...)
}

#ABS - Returns the absolute value of a number, a number without its sign
excel_abs <- function(number){
  abs(number)
}

#SQRT - Returns the square root of a number
excel_sqrt <- function(number){
  sqrt(number)
}

#ROUND - Rounds a number to a specified number of digits
excel_round <- function(number, num_digits){
  round(number, num_digits)
}

#AND
excel_and <- function(x,y){
  x&y
}

#OR
excel_or <- function(x,y){
  x|y
}

#REPLACE
  
#ISNA

#MATCH

#STDEV TODO: sort out the excel format for this one, rn user would need to type TRUE or FALSE for the rmNA; find the parameter names in excel
excel_stdev <- function (range, removeNA){
  sd(range,removeNA)
}
excel_stdev(iris$Sepal.Width,TRUE)

#WEEKDAY - find the parameter names in excel; how to ensure its in the correct date format?
excel_weekday <- function(date)
  weekdays(as.Date('05/05/2020'), abbreviate = F)

#SPECIFICDAY - Returns the date for a specific weekday for the current week
excel_specificday <- function (weekday){
wantedvalue <- ifelse(tolower(weekday)=="sunday",0,
              ifelse(tolower(weekday)=="monday",1,
              ifelse(tolower(weekday)=="tuesday",2,
              ifelse(tolower(weekday)=="wednesday",3,
              ifelse(tolower(weekday)=="thursday",4,
              ifelse(tolower(weekday)=="friday",5,
              ifelse(tolower(weekday)=="saturday",6)))))))
todayvalue <- ifelse(weekdays(as.Date(lubridate::today()), abbreviate = F)=="Sunday",0,
              ifelse(weekdays(as.Date(lubridate::today()), abbreviate = F)=="Monday",1,
              ifelse(weekdays(as.Date(lubridate::today()), abbreviate = F)=="Tuesday",2,
              ifelse(weekdays(as.Date(lubridate::today()), abbreviate = F)=="Wednesday",3,
              ifelse(weekdays(as.Date(lubridate::today()), abbreviate = F)=="Thursday",4,
              ifelse(weekdays(as.Date(lubridate::today()), abbreviate = F)=="Friday",5,
              ifelse(weekdays(as.Date(lubridate::today()), abbreviate = F)=="Saturday",6)))))))
lubridate::today()+wantedvalue-todayvalue
}
