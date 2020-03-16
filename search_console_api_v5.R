library(searchConsoleR)
library(openxlsx)
library(jsonlite)

# Rate Limits:
# https://developers.google.com/webmaster-tools/search-console-api-original/v3/limits

scr_auth()

search_console_websites <- list_websites()
print(search_console_websites)

website <- "https://howtostartanllc.com"

start_date <- "2018-03-04"
end_date <- "2018-03-20"

## what to download, choose between date, query, page, device, country
download_dimensions_page_query <- c('page', 'query')

## what type of Google search, choose between 'web', 'video' or 'image'
type <- c('web')

title_sc <- "search_console_"

title_dates <- paste(start_date, "_", end_date, sep = "")

title_file <- paste(title_sc, title_dates, ".xlsx", sep = "")

# print(warnings())

wb <- createWorkbook()
# saveWorkbook(wb, file = title_file, overwrite = TRUE)

worksheet_page_and_queries <- addWorksheet(wb, "Pages and Queries", gridLines = TRUE, visible = TRUE)

worksheet_date_range <- addWorksheet(wb, "Totals, Date Range", gridLines = TRUE, visible = TRUE)

worksheet_each_date <- addWorksheet(wb, "Totals, Each Date", gridLines = TRUE, visible = TRUE)

worksheet_each_page <- addWorksheet(wb, "Totals, Each Page", gridLines = TRUE, visible = TRUE)

get_search_page_and_queries_data <- function(start_date, end_date, download_dimensions_pq){
  search_page_and_queries_data <- search_analytics(siteURL = "https://howtostartanllc.com",
                                                   startDate = start_date, 
                                                   endDate = end_date, 
                                                   dimensions = download_dimensions_pq,
                                                   searchType = 'web', 
                                                   rowLimit = 100000,
                                                   walk_data = "byBatch")
  
  search_page_and_queries_Table <- search_page_and_queries_data
  
  return(unique(search_page_and_queries_Table))
}

search_page_and_queries_DF <- get_search_page_and_queries_data(start_date, end_date, download_dimensions_page_query)

round_df <- function(df, digits) {
  nums <- vapply(df, is.numeric, FUN.VALUE = logical(1))
  
  df[,nums] <- round(df[,nums], digits = digits)
  
  (df)
}

search_page_and_queries_DF <- round_df(search_page_and_queries_DF, digits=3)

n<-dim(search_page_and_queries_DF)[1]
search_page_and_queries_DF<-search_page_and_queries_DF[1:(n-1),]

# print(sum(searchquery$impressions, na.rm=T))

writeDataTable(wb, worksheet_page_and_queries, search_page_and_queries_DF, startCol = 1, startRow = 1, colNames = TRUE, 
               tableStyle = "TableStyleMedium9", withFilter = TRUE, keepNA = TRUE)
setColWidths(wb, worksheet_page_and_queries, cols = c(1,2,3,4,5,6), widths = c(35,25,8,10,8,8))






get_search_date_range_data <- function(start_date, end_date, download_dimensions_date_range){
  search_date_range_data <- search_analytics(siteURL = "https://howtostartanllc.com",
                                             startDate = start_date, 
                                             endDate = end_date
                                             #dimensions = "date"
                                             #searchType = 'web', 
                                             #rowLimit = 100000,
                                             #walk_data = "byBatch")
  )
  
  search_date_range_Table <- search_date_range_data
  
  return(unique(search_date_range_Table))
}

search_date_range_DF <- get_search_date_range_data(start_date, end_date, NULL)

round_df <- function(df, digits) {
  nums <- vapply(df, is.numeric, FUN.VALUE = logical(1))
  
  df[,nums] <- round(df[,nums], digits = digits)
  
  (df)
}

search_date_range_DF <- round_df(search_date_range_DF, digits=3)

print(search_date_range_DF)

search_date_range_DF <- cbind(search_date_range_DF, start_date)

search_date_range_DF <- cbind(search_date_range_DF, end_date)

search_date_range_DF <- search_date_range_DF[c(5, 6, 1, 2, 3, 4)]

writeDataTable(wb, worksheet_date_range, search_date_range_DF, startCol = 1, startRow = 1, colNames = TRUE, 
               tableStyle = "TableStyleMedium9", withFilter = TRUE, keepNA = TRUE)
setColWidths(wb, worksheet_date_range, cols = c(1,2,3,4,5,6), widths = c(10,10,10,10,10,10))









get_search_each_date_data <- function(start_date, end_date, download_dimensions_date_range){
  search_each_date_data <- search_analytics(siteURL = "https://howtostartanllc.com",
                                            startDate = start_date, 
                                            endDate = end_date,
                                            dimensions = "date"
                                            #searchType = 'web', 
                                            #rowLimit = 100000,
                                            #walk_data = "byBatch")
  )
  
  search_each_date_Table <- search_each_date_data
  
  return(unique(search_each_date_Table))
}

search_each_date_DF <- get_search_each_date_data(start_date, end_date, NULL)

round_df <- function(df, digits) {
  nums <- vapply(df, is.numeric, FUN.VALUE = logical(1))
  
  df[,nums] <- round(df[,nums], digits = digits)
  
  (df)
}

search_each_date_DF <- round_df(search_each_date_DF, digits=3)

day_names <- weekdays(as.Date(search_each_date_DF[[1]],'%Y-%m-%d'))

print(day_names)

print(search_each_date_DF[[1]])

search_each_date_DF <- cbind(search_each_date_DF, day_names)

search_each_date_DF <- search_each_date_DF[c(1, 6, 2, 3, 4, 5)]

writeDataTable(wb, worksheet_each_date, search_each_date_DF, startCol = 1, startRow = 1, colNames = TRUE, 
               tableStyle = "TableStyleMedium9", withFilter = TRUE, keepNA = TRUE)
setColWidths(wb, worksheet_each_date, cols = c(1,2,3,4,5,6), widths = c(10,10,10,10,10,10))

#plot_each_date <- plot(search_each_date_DF$date, type = "1", ylim = c(0, max(search_each_date_DF$clicks)))

plot(search_each_date_DF$date, ylim = c(0,max(search_each_date_DF$clicks)),type="l",col="red")
lines(search_each_date_DF$date, ylim = c(0,max(search_each_date_DF$impressions)),col="green")










get_search_each_page_data <- function(start_date, end_date, download_dimensions_date_range){
  search_each_page_data <- search_analytics(siteURL = "https://howtostartanllc.com",
                                            startDate = start_date, 
                                            endDate = end_date,
                                            dimensions = "page"
                                            #searchType = 'web', 
                                            #rowLimit = 100000,
                                            #walk_data = "byBatch")
  )
  
  search_each_page_Table <- search_each_page_data
  
  return(unique(search_each_page_Table))
}

search_each_page_DF <- get_search_each_page_data(start_date, end_date, NULL)

round_df <- function(df, digits) {
  nums <- vapply(df, is.numeric, FUN.VALUE = logical(1))
  
  df[,nums] <- round(df[,nums], digits = digits)
  
  (df)
}

search_each_page_DF <- round_df(search_each_page_DF, digits=3)

#ndr <- dim(search_date_range_DF)[1]
#search_date_range_DF <- search_date_range_DF[1:(ndr-1),]

writeDataTable(wb, worksheet_each_page, search_each_page_DF, startCol = 1, startRow = 1, colNames = TRUE, 
               tableStyle = "TableStyleMedium9", withFilter = TRUE, keepNA = TRUE)
setColWidths(wb, worksheet_each_page, cols = c(1,2,3,4,5), widths = c(35,10,10,10,10))

saveWorkbook(wb, file = title_file, overwrite = TRUE)

print(warnings())
