library(openxlsx)       # Import the files for import/export of Excel data
library(ggplot2)

# Files containing extra information taken at plots
extraDataFiles <- c(
  "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/ExtraData/Korrekturlest_HEM_Feltdata Gutulia 2018-PA.xlsx",
  "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/ExtraData/Korrekturlest_HEM_Feltdata Dividal 2018_PA.xlsx"
)
# Place to store the analysis output
outputLocation <- "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/AnalyserOutput"

# Set the sheet names containing the extra data
extraDataSheetNames <- c("Soppskade blåbær", "Soppskader på blåbær")
# Set the plot column
extraDataPlotColumn <- "Analyse_navn"
extraDataDescColumn <- "GjeldendeFloranavnNINA"
extraDataFreqColumn <- "Frekvens"
extraDataCoverColumn <- "Dekn"

# Load in the TOV data
dataLocation <- "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/CleanedData.rds"
TOVData <- readRDS(dataLocation)

# Import and reorder the datasets
extraDataFrame <- lapply(X = extraDataFiles, FUN = function(curFile, extraDataPlotColumn, extraDataDescColumn, extraDataFreqColumn, extraDataCoverColumn, extraDataSheetNames) {
  # Retrieve the relevant sheet name
  curSheet <- extraDataSheetNames[extraDataSheetNames %in% getSheetNames(curFile)][1]
  # Import the relevant data file
  inTable <- read.xlsx(curFile, curSheet)
  # Retrieve the names of the extra data
  extraDataCols <- unique(inTable[, extraDataDescColumn])
  # Retrieve the plot names
  plotNames <- unique(inTable[, extraDataPlotColumn])
  outMatrix <- t(sapply(X = plotNames, FUN = function(curPlot, extraDataCols, plotColName, descColName, freqColName, coverColName, inTable) {
    # Make a temporary test table
    testTable <- inTable[curPlot == inTable[, plotColName], ]
    rownames(testTable) <- testTable[, descColName]
    # Value vector
    valueVec <- c(testTable[extraDataCols, freqColName], testTable[extraDataCols, coverColName])
    setNames(
      ifelse(is.na(valueVec), 0, valueVec),
      c(paste("(Frekvens)", extraDataCols, sep = " "), paste("(Dekning)", extraDataCols, sep = " ")))
  }, extraDataCols = extraDataCols, plotColName = extraDataPlotColumn, descColName = extraDataDescColumn,
    freqColName = extraDataFreqColumn, coverColName = extraDataCoverColumn, inTable = inTable))
  rownames(outMatrix) <- plotNames
  as.data.frame(outMatrix)
}, extraDataPlotColumn = extraDataPlotColumn, extraDataDescColumn = extraDataDescColumn,
  extraDataFreqColumn = extraDataFreqColumn, extraDataCoverColumn = extraDataCoverColumn,
  extraDataSheetNames = extraDataSheetNames)
extraDataFrame <- do.call(rbind, lapply(X = extraDataFrame, FUN = function(curDataFrame, allnames) {
  # Initialise a temporary frame
  outFrame <- as.data.frame(matrix(0, nrow = nrow(curDataFrame), ncol = length(allnames), dimnames = list(rownames(curDataFrame), allnames)))
  outFrame[, colnames(curDataFrame)] <- curDataFrame
  outFrame
}, allnames = sort(unique(unlist(lapply(X = extraDataFrame, FUN = colnames))))))

# Create a new workbook for the output
curWorkbook <- createWorkbook()
# Create box plots of the extra data
boxPlots <- apply(X = TOVData$siteInfo, FUN = function(curRow, extraDataFrame, outputLocation, curWorkbook) {
  # Get the current data frame
  curFrame <- extraDataFrame[grepl(paste("^", curRow[2], sep = ""), rownames(extraDataFrame), perl = TRUE), ]
  # Create a series of box plots
  boxPlots <- lapply(X = c("Frekvens", "Dekning"), FUN = function(curType, curFrame, curSite, outputLocation, curWorkbook) {
    # Retrieve only those columns relevant to the current type of analysis
    tempFrame <- curFrame[, grepl(paste("^\\(", curType, "\\)", sep = ""), colnames(curFrame), perl = TRUE)]
    colnames(tempFrame) <- gsub(paste("^\\(", curType, "\\)", sep = ""), "", colnames(tempFrame), perl = TRUE)
    # Create a summary table
    summaryTable <- apply(X = as.matrix(tempFrame), FUN = function(curCol) {
      setNames(c(
        mean(curCol), sd(curCol), quantile(curCol, c(0.5, 0.25, 0.75))
      ), c(
        "mean", "sd", "median", "lowerQuartile", "upperQuartile"
      ))
    }, MARGIN = 2)
    # Create a worksheet to hold the data
    addWorksheet(curWorkbook, paste(curSite, curType, sep = "_"))
    writeDataTable(curWorkbook, paste(curSite, curType, sep = "_"), rbind(tempFrame, summaryTable), rowNames = TRUE)
    # Rearrange the data frame
    tempFrame <- do.call(rbind, lapply(X = names(tempFrame), FUN = function(curDataType, inFrame) {
      outFrame <- as.data.frame(matrix(as.integer(c()), nrow = 0, ncol = 2, dimnames = list(NULL, c("tempOut", "dataType"))))
      if(any(inFrame[, curDataType] > 0)) {
        outFrame <- data.frame(
          tempOut = inFrame[, curDataType],
          dataType = rep(curDataType, nrow(inFrame))
        )
      }
      outFrame
    }, inFrame = tempFrame))
    # Make the plot of the data
    outPlot <- ggplot(tempFrame, aes(dataType, tempOut)) + geom_boxplot(fill = rgb(238, 174, 238, maxColorValue = 255)) +
      theme_classic() + coord_flip() + ylab(curType) + xlab("")
    ggsave(paste(outputLocation, "/", curSite, "_", curType, "ExtraInformation.svg", sep = ""), width = 15, height = 15, units = "cm")
    outPlot
  }, curFrame = curFrame, curSite = curRow[1], outputLocation = outputLocation, curWorkbook = curWorkbook)
  boxPlots
}, extraDataFrame = extraDataFrame, outputLocation = outputLocation, curWorkbook = curWorkbook, MARGIN = 1)
names(boxPlots) <- rownames(TOVData$siteInfo)
saveWorkbook(curWorkbook, paste(outputLocation, "ExtraInformation.xlsx", sep = "/"), TRUE)
