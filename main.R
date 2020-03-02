# Requirements
#   Open Java Development Kit, or equivalent: https://jdk.java.net/
#   Instructions for installing it on Windows: https://stackoverflow.com/a/52531093
#   Chrome
#     Make sure to set chromever (check your chrome version in the actual browser).
#     Run binman::list_versions("chromedriver") to see currently sourced versions and compare to above.

# Libraries
library(dplyr)
library(RSelenium)
library(xlsx)

# Function to try until an element appears
tryElem <- function(obj, how, what){
  webElem <-NULL
  while(length(webElem) == 0){
    webElem <- tryCatch({obj$findElements(using = how, value = what)}, error = function(e){NULL})
  }
  webElem
}

eCaps <- list(
  chromeOptions = list(
    prefs = list(
      "profile.default_content_settings.popups" = 0L,
      "download.prompt_for_download" = FALSE,
      "download.default_directory" = paste0(normalizePath(getwd()),"\\data")
    )
  )
)

# Set up the server and connect to it
rD <- rsDriver(browser = "chrome", 
               chromever = "80.0.3987.106", 
               verbose = FALSE, 
               extraCapabilities = eCaps)

remDr <- rD$client

# Navigate to the stocks page
remDr$navigate("https://www.oslobors.no/markedsaktivitet/#/list/shares/quotelist/ose/all/all/false")

# Find the links to each ticker
webElem <- tryElem(remDr, "css selector", "td.ITEM_SECTOR > a")

# Create a vector of links
tickLinks <- lapply(webElem, FUN=function(x){x$getElementAttribute("href")}) %>% 
  unlist()

# Create a vector of tick names
tickTexts <- lapply(webElem, FUN=function(x){x$getElementAttribute("text")}) %>% 
  unlist() %>% 
  gsub("[^[:alnum:]]", "", .)

# Create a list to hold the data
stockData <- list()

# Loop through the vector of links and download the data
# Will require a better solution than pausing for making sure page loads/downloads are complete
for (i in seq_along(tickLinks)){
  
  # Go to the page
  remDr$navigate(tickLinks[i])
  
  Sys.sleep(1)
  
  # Find and click the 5-year drop-down
  webElem <- tryElem(remDr, "xpath", "//*/option[@value = 'FIVE_YEARS']")
  webElem[[1]]$clickElement()
  
  # Find the download link and click it
  webElem <- tryElem(remDr, "css selector", ".excelExport-btn.ng-binding")
  webElem[[1]]$clickElement()
  
  Sys.sleep(1)
  
  # Find the latest file
  lastFile <- file.info(list.files(paste0(getwd(),"/data"), full.names = TRUE)) %>% 
    tibble::rownames_to_column() %>% 
    filter(mtime == max(mtime)) %>% 
    .$rowname
  
  # Read the .xlsx file into a list
  stockData[[i]] <- read.xlsx(lastFile, 1, encoding = "UTF-8")
  names(stockData)[i] <- tickTexts[i]
}

# From here you can do whatever you want with the data.frames stored in the list