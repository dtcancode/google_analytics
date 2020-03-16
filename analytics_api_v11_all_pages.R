library(googleAnalyticsR)
library(openxlsx)
library(jsonlite)

#2018-04-10 14:39:32> Default Google Project for googleAnalyticsR is now set.  
#This is shared with all googleAnalyticsR users. 
#If making a lot of API calls, please: 
#1) create your own Google Project at https://console.developers.google.com 
#2) Activate the Google Analytics Reporting API 
#3) set options(googleAuthR.client_id) and options(googleAuthR.client_secret) 
#4) Reload the package.
#2018-04-10 14:39:32> Set API cache
#2018-04-10 14:39:32> No environment argument found, looked in GA_AUTH_FILE
#Warning message:
#  package ‘googleAnalyticsR’ was built under R version 3.4.4 

load_credentials_json_file <- file(description = "C://Users/Daniel/R_Projects/Google_API/client_secret.json", open = "a+")

credentials_json_file <- fromJSON(load_credentials_json_file, flatten = FALSE)

client_id <- credentials_json_file[[1]]["client_id"]
client_secret <- credentials_json_file[[1]]["client_secret"]

options(googleAuthR.client_id = client_id)
options(googleAuthR.client_secret = client_secret)

GA_AUTH_FILE <- ("./client_secret.json")

devtools::reload(pkg = devtools::inst("googleAnalyticsR"))

ga_auth()

#account_list <- ga_account_list()

## account_list will have a column called "viewId"
#print(account_list$viewId)

ga_id <- 31471801

df1 <- dim_filter("pagePath","PARTIAL","Email", not = TRUE)

#df2 <- dim_filter("landingPagePath","PARTIAL","business-ideas", not = TRUE)

df3 <- dim_filter("pagePath","PARTIAL","signin", not = TRUE)

df4 <- dim_filter("pagePath","PARTIAL","translate", not = TRUE)

#df5 <- dim_filter("landingPagePath","PARTIAL","(not set)", not = TRUE)

df6 <- dim_filter("hostname","PARTIAL","google", not = TRUE)

df7 <- dim_filter("hostname","PARTIAL","azure", not = TRUE)

fc2 <- filter_clause_ga4(list(df1, df3, df4, df6, df7), operator = "AND")

order_filter_desc <- order_type("pagePath",
                                  sort_order = "ASCENDING",
                                  orderType = "VALUE")

# maximum of 10 metrics per query
# maximum of 7 dimensions per query

start_date <- "2018-04-10"
end_date <- "2018-04-10"

title_searches <- paste("google_anlytics", "_", sep = "")

title_dates <- paste(start_date, "_", end_date, "_", sep = "")

title_file <- paste(title_searches, title_dates, "all_pages", ".xlsx", sep = "")

wb <- createWorkbook()
# saveWorkbook(wb, file = title_file, overwrite = TRUE)

addWorksheet(wb, "All Pages", gridLines = TRUE, tabColour = NULL,
             zoom = 100, header = NULL, footer = NULL, evenHeader = NULL,
             evenFooter = NULL, firstHeader = NULL, firstFooter = NULL,
             visible = TRUE, paperSize = getOption("openxlsx.paperSize", default = 9),
             orientation = getOption("openxlsx.orientation", default = "portrait"),
             vdpi = getOption("openxlsx.vdpi", default = getOption("openxlsx.dpi", default = 300)), 
             hdpi = getOption("openxlsx.hdpi", default = getOption("openxlsx.dpi", default = 300)))

# Behavior -> Site Content -> All Pages
# Page, Pageviews, Unique Pageviews, Avg. Time on Page, Entrances, Bounce Rate, % Exit

get_analytics_data <- function(start_date, end_date){
  analytics_data <- google_analytics(ga_id, 
                                     date_range = c(start_date, end_date), 
                                     metrics = c(
                                       #"sessions",
                                       #"percentNewSessions",
                                       #"newUsers",
                                       #"bounceRate",
                                       "pageviews",
                                       "uniquePageviews",
                                       "avgTimeOnPage",
                                       "entrances",
                                       "bounceRate",
                                       "exitRate"
                                       #"timeOnPage"
                                       #"pageviews"
                                     ),
                                     #dimensions = c("landingPagePath", "sessionDurationBucket", "sessionCount"),
                                     #dimensions = c("landingPagePath", "sessionCount"),
                                     dimensions = c("hostname", "pagePath"),
                                     dim_filters = fc2,
                                     order = order_filter_desc,
                                     max = -1)
  
  analytics_Table <- analytics_data
  
  return(unique(analytics_Table))
  #return(analytics_Table)
}

analytics_Table <- get_analytics_data(start_date, end_date)

#n<-dim(analytics_Table)[1]
#analytics_Table <- analytics_Table[1:(n-1),]

round_df <- function(df, digits) {
  nums <- vapply(df, is.numeric, FUN.VALUE = logical(1))
  
  df[,nums] <- round(df[,nums], digits = digits)
  
  (df)
}
  
analytics_Table <- round_df(analytics_Table, digits=3)

#analytics_Table$"page" <- paste("https://", analytics_Table$"hostname", analytics_Table$"pagePath", sep = "")

#analytics_Table$"hostname" <- NULL

#analytics_Table$"pagePath" <- NULL

#analytics_Table <- analytics_Table[c("page", "avgSessionDuration", "timeOnPage", "avgTimeOnPage")]

#analytics_Table <- 

print(analytics_Table)

writeDataTable(wb, "All Pages", analytics_Table, startCol = 1, startRow = 1, colNames = TRUE, 
               tableStyle = "TableStyleMedium9", withFilter = TRUE, keepNA = TRUE)
setColWidths(wb, "All Pages", cols = c(1,2,3,4,5,6), widths = c(35,35,15,15,15,15))

saveWorkbook(wb, file = title_file, overwrite = TRUE)
