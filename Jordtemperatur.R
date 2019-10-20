options(java.parameters = "-Xmx8192m")
library(XLConnect)       # Import the files for import/export of Excel data

# The file that forms the basis of the soil temperature
baseFile <- "R:/Prosjekter/15450000 - TOV-vegetasjon/Jordtemperatur - loggere"
# The location to store the output
outFolder <- "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/AnalyserOutput"
# Location of the Excel template file
templateFile <- "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/AnalysisTemplates/JordTemperatur.xlsx"

# allTimeSeries <- readRDS("C:/CloudStorage/Work/TOV/AnalyserOutput/JordTemperatur.rds")
# templateFile <- "C:/CloudStorage/Work/TOV/AnalysisTemplates/JordTemperatur.xlsx"
# outFolder <- "C:/CloudStorage/Work/TOV/AnalyserOutput"

# Set a list of site names with possible alises in the dataset (for example: "Dividal" often appears as "Dividalen")
siteAliases <- list(
  "Børgefjell" = "Børgefjell",
  "Dividal" = c("Dividal", "Dividalen"),
  "Gutulia" = "Gutulia",
  "Lund" = "Lund",
  "Møsvatn" = "Møsvatn",
  "Åmotsdalen" = "Åmotsdalen"
)

# Get a list of directories containing data
allDirs <- list.dirs(baseFile)
dataDirs <- allDirs[sapply(X = allDirs, FUN = function(curDir, possSiteNames) {
  any(sapply(X = possSiteNames, FUN = grepl, x = curDir, perl = TRUE))
}, possSiteNames = paste(baseFile, unlist(siteAliases), "(Vegetasjonsfelter/)*\\d\\d\\d\\d\\-\\d\\d\\d\\d", sep = "/"))]

# Import the data into a set of data frames
allTimeSeries <- do.call(rbind, lapply(X = dataDirs, FUN = function(curDataDir, siteAliases) {
  # Find the site that the current data folder refers to
  curSite <- factor(names(siteAliases)[sapply(X = siteAliases, FUN = function(curAliases, curDataDir) {
    any(sapply(X = curAliases, FUN = grepl, x = curDataDir, fixed = TRUE))
  }, curDataDir = curDataDir)], levels = names(siteAliases))
  # Retrieve the current year range
  curYearRange <- sort(as.integer(strsplit(gsub("^.*/", "", curDataDir, perl = TRUE), "-", fixed = TRUE)[[1]]))
  # Get a list of the files in the current directory
  curFiles <- list.files(curDataDir, recursive = TRUE)
  # Find the file that contains all the logger information information
  curAlleFile <- paste(curDataDir, curFiles[
    grepl("alle", curFiles, fixed = TRUE) & !grepl("working", curFiles, fixed = TRUE) & !grepl("^~\\$", curFiles, perl = TRUE) & !grepl("Copy", curFiles, fixed = TRUE)
    ], sep = "/")
  cat("Processing file", curAlleFile, "...\n", sep = " ")
  # Retrieve a list of sheets in the workbook
  do.call(rbind, lapply(X = getSheets(loadWorkbook(curAlleFile)), function(curSheetName, curSite, curYearRange, curAlleFile) {
    # Format the sheet name so that it can be read correctly
    curFormattedSheetName <- gsub("(_GJENFUNNET|I2016_|_IKKEFULLSTENDIG)", "",
      gsub("\\s+$", "", gsub("^\\s+", "", toupper(curSheetName), perl = TRUE), perl = TRUE), perl = TRUE)
    # Retrieve the plot number that the current sheet refers to (and the logger number if multiple loggers were used at the same plot)
    curPlotNumber <- as.integer(gsub("\\D+$", "", gsub("^\\D+", "", curFormattedSheetName, perl = TRUE), perl = TRUE))
    curLoggerNumber <- gsub("^\\D+\\d+", "", curFormattedSheetName, perl = TRUE)
    if(curLoggerNumber == "") {
      curLoggerNumber <- "A"
    }
    curLoggerNumber <- gsub("\\s+$", "", gsub("^\\s+", "", curLoggerNumber, perl = TRUE), perl = TRUE)
    outFrame <- data.frame()
    if(nrow(readWorksheetFromFile(curAlleFile, curSheetName)) > 3) {
      # Read the logger metadata
      curLoggerMetadata <- readWorksheetFromFile(curAlleFile, curSheetName, startRow = 5, endRow = 23, startCol = 5, endCol = 6, header = FALSE)
      # Get the button ID
      curLoggerID <- as.character(curLoggerMetadata[1, 2])
      # Get the download time
      curDownloadTime <- as.POSIXlt(curLoggerMetadata[11, 2], "%d/%m/%Y %H:%M:%S", tz = "Europe/Oslo")
      # Retrieve the logger records
      rawLoggerRecords <- readWorksheetFromFile(curAlleFile, curSheetName, startRow = 4, startCol = 1, endCol = 3, header = TRUE)
      cat("\tProcessing site ", as.character(curSite), " (plot ", curPlotNumber, ", logger ", curLoggerNumber, " - ID ", curLoggerID, ", years ", curYearRange[1], "-", curYearRange[2], ") ... ", sep = "")
      # Create an output data frame
      outFrame <- data.frame(
        timeStamp = as.POSIXlt(paste(rawLoggerRecords[, 1], rawLoggerRecords[, 2], sep = " "), "%d/%m/%Y %H:%M:%S", tz = "Europe/Oslo"),
        temperature = rawLoggerRecords[, 3],
        sourceFile = rep(curAlleFile, nrow(rawLoggerRecords)),
        site = rep(curSite, nrow(rawLoggerRecords)),
        plotNumber = rep(curPlotNumber, nrow(rawLoggerRecords)),
        sampleDesignation = rep(curLoggerNumber, nrow(rawLoggerRecords)),
        loggerID = rep(curLoggerID, nrow(rawLoggerRecords)),
        timeSeriesID = paste(
          # Create a time series ID for eact plot being monitored
          rep(curSite, nrow(rawLoggerRecords)),
          rep(curPlotNumber, nrow(rawLoggerRecords)),
          rep(curLoggerNumber, nrow(rawLoggerRecords))
        ),
        downloadTimeStamp = rep(curDownloadTime, nrow(rawLoggerRecords)),
        implantationYear = rep(curYearRange[1], nrow(rawLoggerRecords)),
        removalYear = rep(curYearRange[2], nrow(rawLoggerRecords)),
        stringsAsFactors = FALSE,
        row.names = paste(curSite, curPlotNumber, curLoggerNumber, "_", gsub("/", "", rawLoggerRecords[, 1], fixed = TRUE), "_", gsub(":", "", rawLoggerRecords[, 2], fixed = TRUE), sep = "")
      )
    } else {
      cat("\tProcessing site ", as.character(curSite), " (plot ", curPlotNumber, ", *NO LOGGER INFORMATION*, years ", curYearRange[1], "-", curYearRange[2], ") ... ", sep = "")
    }
    cat(nrow(outFrame), "records found\n", sep = " ")
    outFrame
  }, curSite, curYearRange, curAlleFile))
}, siteAliases = siteAliases))
# Store the serialised data as a temporary measure
saveRDS(allTimeSeries, file = paste(outFolder, "JordTemperatur.rds", sep = "/"))

# Start here if you don't want to repeat the data input and processing
allTimeSeries <- readRDS(paste(outFolder, "JordTemperatur.rds", sep = "/"))

# Load the Excel template format file
outWorkbook <- loadWorkbook(templateFile)
# Create a data format for the POSIX-formatted time stamp (because Excel doens't understand them for some stupid reason)
timeStampStyle <- createCellStyle(outWorkbook, name = "POSIXTimeStamp")
setDataFormat(timeStampStyle, "dd.mm.yyyy hh:mm:ss")
# Create a sheet to hold the soil temperature outputs
createSheet(outWorkbook, name = "JordTemperatur")
writeWorksheet(outWorkbook, data = allTimeSeries, sheet = "JordTemperatur")
# Apply the timestamp formatting
setCellStyle(outWorkbook, sheet = "JordTemperatur", cellstyle = timeStampStyle, col = c(rep(1, nrow(allTimeSeries)), rep(9, nrow(allTimeSeries))), row = c(1:nrow(allTimeSeries), 1:nrow(allTimeSeries)) + 1)
# Set the named regions
createName(outWorkbook, "timeValues", formula = paste("JordTemperatur!A2:A", nrow(allTimeSeries) + 1, sep = ""))
createName(outWorkbook, "tempValues", formula = paste("JordTemperatur!B2:B", nrow(allTimeSeries) + 1, sep = ""))
createName(outWorkbook, "siteValues", formula = paste("JordTemperatur!D2:D", nrow(allTimeSeries) + 1, sep = ""))
createName(outWorkbook, "plotID", formula = paste("JordTemperatur!E2:E", nrow(allTimeSeries) + 1, sep = ""))
createName(outWorkbook, "sampleID", formula = paste("JordTemperatur!F2:F", nrow(allTimeSeries) + 1, sep = ""))
createName(outWorkbook, "timeSeriesID", formula = paste("JordTemperatur!G2:G", nrow(allTimeSeries) + 1, sep = ""))

saveWorkbook(outWorkbook, file = paste(outFolder, "JordTemperatur.xlsx", sep = "/"))
