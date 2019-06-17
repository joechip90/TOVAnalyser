Terrestrisk Naturovervåking: Vegetasjonsanalyser
================
Joseph Chipperfield

Import the Libraries
--------------------

``` r
library(vegan)       # Include the vegetation analysis library
library(ggplot2)     # Import graphics libraries
library(rlang)       # R language object library
library(openxlsx)    # Import the files for import/export of Excel data
library(knitr)       # Allow for markdown-formatted tables and figures
```

Import the Data
---------------

``` r
dataLocation <- "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/CleanedData.rds"
TOVData <- readRDS(dataLocation)

outputLocation <- "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/AnalyserOutput"
if(dir.exists(outputLocation)) {
  unlink(outputLocation, recursive = TRUE)
}
dir.create(outputLocation)
```

Initialise Helper Functions
---------------------------

``` r
# Function to add together columns that are of the same species
compressCommunityMatrix <- function(inputMatrix) {
  # Retrieve the species from the column names
  specNames <- sort(unique(gsub("_[A-Z]$", "", colnames(inputMatrix), perl = TRUE)))
  # Add columns together that belong to the same species
  outputMatrix <- sapply(X = specNames, FUN = function(curSpecies, inputMatrix) {
    speciesMat <- inputMatrix[, curSpecies == gsub("_[A-Z]$", "", colnames(inputMatrix), perl = TRUE), drop = FALSE]
    apply(X = speciesMat, FUN = sum, MARGIN = 1)
  }, inputMatrix = inputMatrix)
  # Make sure the columns and rows are named appropriately
  colnames(outputMatrix) <- specNames
  rownames(outputMatrix) <- rownames(inputMatrix)
  outputMatrix
}

# Function to create trajectory plots when using ordination
createTrajectoryPlot <- function(ordinationFrame, xName, yName) {
  # Retrieve the expressions in the input arguments
  xName <- enquo(xName)
  yName <- enquo(yName)
  # Retrieve the string version of the input arguments
  xNameString <- as_label(xName)
  yNameString <- as_label(yName)
  # Retrieve the different site codes in the ordination frame
  siteCodes <- unique(gsub("\\d+[A-Z]*\\-\\d+$", "", rownames(ordinationFrame), perl = TRUE))
  # Create a trajectory plot for each site
  trajPlots <- lapply(X = siteCodes, FUN = function(curSite, ordinationFrame, xNames, yNames) {
    # Subset only the relevant ordination columns at the current site
    ordinationFrameSubset <- as.data.frame(ordinationFrame[grepl(paste("^", curSite, sep = ""), rownames(ordinationFrame), perl = TRUE), c(xNames$str, yNames$str)])
    # Get a list of sampled years at the current site
    sampledYears <- sort(unique(as.integer(gsub("^[A-Z]\\d+[A-Z]*\\-", "", rownames(ordinationFrameSubset), perl = TRUE))))
    # Initialise a scatter plot of the ordination values
    outPlot <- ggplot(ordinationFrameSubset, aes(!!(xNames$qu), !!(yNames$qu))) + theme_classic() + geom_point()
    colnames(ordinationFrameSubset) <- c("ORDX", "ORDY")
    if(length(sampledYears) > 1) {
      # If there is more than one sampled year then plot arrows between each set of years
      # Create a series of data frames for each year transition
      trajOrdinationFrames <- lapply(X = 1:(length(sampledYears) - 1), FUN = function(yearIter, ordinationFrameSubset, sampledYears) {
        # Get the year the plots are going from and to
        fromYear <- sampledYears[yearIter]
        toYear <- sampledYears[yearIter + 1]
        # Find those plots that have samples in both years
        fromPlots <- gsub("\\-\\d+$", "", rownames(ordinationFrameSubset)[grepl(paste(fromYear, "$", sep = ""), rownames(ordinationFrameSubset), perl = TRUE)], perl = TRUE)
        toPlots <- gsub("\\-\\d+$", "", rownames(ordinationFrameSubset)[grepl(paste(toYear, "$", sep = ""), rownames(ordinationFrameSubset), perl = TRUE)], perl = TRUE)
        plotsBoth <- c(fromPlots, toPlots)[duplicated(c(fromPlots, toPlots))]
        # Create a data.frame with the trajectories included
        outTrajFrame <- as.data.frame(matrix(nrow = 0, ncol = 4, dimnames = list(NULL, c("FROM_ORDX", "TO_ORDX", "FROM_ORDY", "TO_ORDY"))))
        if(length(plotsBoth) > 0) {
          outTrajFrame <- data.frame(
            FROM_ORDX = ordinationFrameSubset[paste(plotsBoth, fromYear, sep = "-"), "ORDX"],
            TO_ORDX = ordinationFrameSubset[paste(plotsBoth, toYear, sep = "-"), "ORDX"],
            FROM_ORDY = ordinationFrameSubset[paste(plotsBoth, fromYear, sep = "-"), "ORDY"],
            TO_ORDY = ordinationFrameSubset[paste(plotsBoth, toYear, sep = "-"), "ORDY"]
          )
          rownames(outTrajFrame) <- plotsBoth
        }
        outTrajFrame
      }, ordinationFrameSubset = ordinationFrameSubset, sampledYears = sampledYears)
      names(trajOrdinationFrames) <- paste(sampledYears[1:(length(sampledYears) - 1)], sampledYears[2:length(sampledYears)], sep = "_")
      # Add trajectory arrows to the ggplot
      outPlot <- outPlot + lapply(X = trajOrdinationFrames, FUN = function(curOrdinationFrame) {
        geom_segment(aes(x = FROM_ORDX, y = FROM_ORDY, xend = TO_ORDX, yend = TO_ORDY), data = curOrdinationFrame, colour = "blue",
          arrow = arrow(type = "closed", length = unit(0.01, "npc")))
      })
    }
    outPlot
  }, ordinationFrame = ordinationFrame, xNames = list(qu = xName, str = xNameString), yNames = list(qu = yName, str = yNameString))
  names(trajPlots) <- siteCodes
  trajPlots
}
```

Ordination Analysis
-------------------

``` r
# Setup community matrices for the frequency values
compressedFreq <- compressCommunityMatrix(t(TOVData$freqMatrix))
# Setup community matrices for the cover values
compressedCover <- compressCommunityMatrix(t(TOVData$coverMatrix))

# Run DCA on each of the different community matrices
compressedFreq_DCA <- decorana(compressedFreq)
compressedCover_DCA <- decorana(compressedCover)

# Run NMDS on each of the different community matrices
compressedFreq_NMDS <- metaMDS(compressedFreq, try = 50, trymax = 10000)
compressedCover_NMDS <- metaMDS(compressedCover, try = 50, trymax = 10000)

# Combine the site scores from the different ordination techniques
compressedFreq_Ordination <- as.data.frame(cbind(scores(compressedFreq_DCA), scores(compressedFreq_NMDS)))
compressedCover_Ordination <- as.data.frame(cbind(scores(compressedCover_DCA), scores(compressedCover_NMDS)))
# Save the ordination scores in an excel files
ordinationBook <- createWorkbook()
addWorksheet(ordinationBook, "Frekvens")
addWorksheet(ordinationBook, "Dekning")
ordHeaderStyle <- createStyle(border = "bottom", borderStyle = "medium", textDecoration = "bold")
writeDataTable(ordinationBook, "Frekvens", compressedFreq_Ordination, headerStyle = ordHeaderStyle, rowNames = TRUE, firstColumn = TRUE)
writeDataTable(ordinationBook, "Dekning", compressedCover_Ordination, headerStyle = ordHeaderStyle, rowNames = TRUE, firstColumn = TRUE)
saveWorkbook(ordinationBook, paste(outputLocation, "OrdinationScores.xlsx", sep = "/"), overwrite = TRUE)

# Create a series of trajectory plots for the DCA scores
freqDCAOrdinationPlotList <- createTrajectoryPlot(compressedFreq_Ordination, DCA1, DCA2)
coverDCAOrdinationPlotList <- createTrajectoryPlot(compressedCover_Ordination, DCA1, DCA2)
# Create a series of trajectory plots for the NMDS scores
freqNMDSOrdinationPlotList <- createTrajectoryPlot(compressedFreq_Ordination, NMDS1, NMDS2)
coverNMDSOrdinationPlotList <- createTrajectoryPlot(compressedCover_Ordination, NMDS1, NMDS2)
```

Yearly-Pairwise Analysis for Each Species
=========================================

``` r
outWorkbook <- createWorkbook()
richnessList <- lapply(X = rownames(TOVData$siteInfo), FUN = function(curSiteCode, communityMatrix, siteInfo, sjiktInfo, outWorkbook) {
  # Current site name
  curSiteName <- as.character(siteInfo[curSiteCode, "SiteName"])
  # Initialise a worksheet for the current site
  addWorksheet(outWorkbook, curSiteName)
  # Subset the matrix with the current site code
  curCommunityMatrix <- communityMatrix[, gsub("\\d+[A-Z]*\\-\\d+$", "", colnames(communityMatrix), perl = TRUE) == curSiteCode]
  # Create a richness matrix
  richnessMatrix <- as.data.frame(sapply(X = unique(gsub(paste("^", curSiteCode, "\\d+[A-Z]*\\-", sep = ""), "", colnames(curCommunityMatrix), perl = TRUE)), FUN = function(curYear, curCommunityMatrix, sjiktInfo) {
    # Filter out the columns that represent samples from the current year
    curYearCommunity <- curCommunityMatrix[, grepl(paste(curYear, "$", sep = ""), colnames(curCommunityMatrix), perl = TRUE)]
    # Count the species richness for each sjikt
    setNames(sapply(X = sjiktInfo$SjiktCode, FUN = function(curSjikt, curYearCommunity) {
      # Assess whether the species row belong to the current sjikt
      isSjikt <- grepl(paste(curSjikt, "$", sep = ""), rownames(curYearCommunity), perl = TRUE)
      outVal <- 0
      if(any(isSjikt)) {
        outVal <- sum(ifelse(apply(X = curYearCommunity[isSjikt, , drop = FALSE], FUN = function(curRow) {
          any(curRow > 0)
        }, MARGIN = 1), 1, 0))
      }
      outVal
    }, curYearCommunity), sjiktInfo$SjiktCode)
  }, curCommunityMatrix = curCommunityMatrix, sjiktInfo = sjiktInfo))
  # Save the table to the workbook
  formattedTable <- cbind(data.frame(Sjikt = rownames(sjiktInfo)), richnessMatrix)
  writeDataTable(outWorkbook, curSiteName, formattedTable, colNames = TRUE, rowNames = FALSE)
  # Display a table
  outTable <- cbind(data.frame(Sjikt = sjiktInfo[, "DescriptionNorsk"]), as.data.frame(richnessMatrix))
  cat("**Species richness at", curSiteName, "**\n")
  print(kable(outTable, format = "markdown", row.names = FALSE, caption = paste("Species richness at", curSiteName, sep = " ")))
  richnessMatrix
}, communityMatrix = TOVData$freqMatrix, siteInfo = TOVData$siteInfo, sjiktInfo = TOVData$sjiktInfo, outWorkbook = outWorkbook)
```

**Species richness at Dividalen **

| Sjikt               |  1993|  1998|  2003|  2008|  2013|  2018|
|:--------------------|-----:|-----:|-----:|-----:|-----:|-----:|
| Trær                |     0|     0|     0|     0|     0|     0|
| Busker              |     0|     0|     0|     0|     0|     1|
| Lyng og dvergbusker |    14|    13|    14|    14|    13|    12|
| Urter               |    47|    49|    49|    52|    48|    49|
| Gras og halvgras    |    17|    15|    16|    16|    17|    18|
| Bladmoser           |    24|    16|    20|    27|    25|    25|
| Levermoser          |    17|    13|    14|    14|    16|    17|
| Busklav             |    25|    24|    20|    23|    23|    23|

**Species richness at Gutuila **

| Sjikt               |  1993|  1998|  2003|  2008|  2013|  2018|
|:--------------------|-----:|-----:|-----:|-----:|-----:|-----:|
| Trær                |     0|     0|     0|     0|     0|     0|
| Busker              |     0|     0|     0|     0|     0|     0|
| Lyng og dvergbusker |    11|    12|    12|    12|    12|    11|
| Urter               |    17|    19|    21|    20|    20|    18|
| Gras og halvgras    |    13|    13|    13|    13|    13|    13|
| Bladmoser           |    18|    18|    19|    18|    23|    21|
| Levermoser          |    11|    11|    11|    11|    13|    10|
| Busklav             |    16|    18|    16|    15|    15|    14|

``` r
names(richnessList) <- rownames(TOVData$siteInfo)
# Save the created workbook
saveWorkbook(outWorkbook, paste(outputLocation, "/SpeciesRichness.xlsx", sep = ""), TRUE)
```

``` r
# Produce an analysis of changes of frequency/cover in different years
yearlyPairwiseWilcoxon <- function(communityMatrix) {
  # Retrieve the site codes
  siteCodes <- sort(unique(gsub("\\d+[A-Z]*\\-\\d+$", "", rownames(communityMatrix), perl = TRUE)))
  outList <- lapply(X = siteCodes, FUN = function(curSite, communityMatrix) {
    # Retrieve just those plots taken at the current site
    communityMatrixSubset <- communityMatrix[grepl(paste("^", curSite, sep = ""), rownames(communityMatrix), perl = TRUE), ]
    # Species list present at site
    speciesList <- colnames(communityMatrixSubset)[apply(X = communityMatrixSubset, FUN = function(curCol) {
      any(curCol > 0)
    }, MARGIN = 2)]
    # Make a list for the analysis of each species
    outList <- lapply(X = speciesList, FUN = function(curSpecies, curSite, communityMatrixSubset) {
      # Find the set of years that the current site has been visited
      yearsVisited <- sort(as.integer(unique(gsub(paste("^", curSite, "\\d+[A-Z]*\\-", sep = ""), "", rownames(communityMatrixSubset), perl = TRUE))))
      # Make a matrix of comparrison years
      yearMatrix <- do.call(rbind, lapply(X = yearsVisited, FUN = function(curYear, yearsVisited) {
        aboveYears <- yearsVisited[curYear < yearsVisited]
        outMat <- matrix(as.integer(c()), ncol = 2, nrow = 0)
        if(length(aboveYears) > 0) {
          outMat <- cbind(rep(curYear, length(aboveYears)), aboveYears)
        }
        colnames(outMat) <- c("fromYear", "toYear")
        outMat
      }, yearsVisited = yearsVisited))
      # Use the matrix of comparrison years to test for differences between species
      cbind(yearMatrix, t(apply(X = yearMatrix, FUN = function(curYears, curSpecies, communityMatrixSubset) {
        # Create two subsets of the community matrix of the 'from' year and the 'to' year
        communityMatrixSubset_fromYear <- communityMatrixSubset[grepl(paste(curYears[1], "$", sep = ""), rownames(communityMatrixSubset), perl = TRUE), ]
        communityMatrixSubset_toYear <- communityMatrixSubset[grepl(paste(curYears[2], "$", sep = ""), rownames(communityMatrixSubset), perl = TRUE), ]
        # Retrieve those plots that are present in both years
        fromPlots <- unique(gsub("\\-\\d+$", "", rownames(communityMatrixSubset_fromYear), perl = TRUE))
        toPlots <- unique(gsub("\\-\\d+$", "", rownames(communityMatrixSubset_toYear), perl = TRUE))
        bothPlots <- sort(c(fromPlots, toPlots)[duplicated(c(fromPlots, toPlots))])
        # Retrieve the values of the community matrix at each year
        values_fromYear <- communityMatrixSubset_fromYear[paste(bothPlots, curYears[1], sep = "-"), curSpecies]
        values_toYear <- communityMatrixSubset_toYear[paste(bothPlots, curYears[2], sep = "-"), curSpecies]
        # Perform the Wilcoxon signed-ranks test
        wilTest <- wilcox.test(values_fromYear, values_toYear, paired = TRUE, exact = FALSE)
        # Produce an output vector with the community data
        setNames(
          c(median(values_fromYear), median(values_toYear), mean(values_fromYear), mean(values_toYear), wilTest$statistic, wilTest$p.value),
          c("fromMedian", "toMedian", "fromMean", "toMean", "Vstat", "pValue"))
      }, curSpecies = curSpecies, communityMatrixSubset = communityMatrixSubset, MARGIN = 1)))
    }, curSite = curSite, communityMatrixSubset = communityMatrixSubset)
    names(outList) <- speciesList
    outList
  }, communityMatrix = communityMatrix)
  names(outList) <- siteCodes
  outList
}
# Run the yearly pairwuse analysis for both the cover and the frequency
compressedFreq_yearlyPairwise <- yearlyPairwiseWilcoxon(compressedFreq)
compressedCover_yearlyPairwise <- yearlyPairwiseWilcoxon(compressedCover)
# Create a set of Excel files containing the pairwise analyses
lapply(X = names(compressedFreq_yearlyPairwise), FUN = function(curSiteCode, freqPairwise, coverPairwise, siteInfo, outputLocation, speciesInfo) {
  # Retrieve the full size name
  fullSiteName <- siteInfo[curSiteCode, "SiteName"]
  # Create a workbook
  siteWorkbook <- createWorkbook()
  # Create sheets for the cover and frequency
  freqSheet <- addWorksheet(siteWorkbook, sheetName = "Frekvens")
  coverSheet <- addWorksheet(siteWorkbook, sheetName = "Dekning")
  # Function to rearrange a species list to a pairwise year data frame
  rearrangePairwiseYears <- function(curSpeciesList) {
    outFrame <- t(sapply(X = curSpeciesList, FUN = function(curYearAnalysis) {
      # Flatten out the input matrix
      setNames(as.double(t(curYearAnalysis[, 3:8])), unlist(lapply(X = paste(curYearAnalysis[, 1], curYearAnalysis[, 2], sep = "-"), FUN = function(curYearText, addColumnHeads) {
        paste(curYearText, addColumnHeads, sep = "_")
      }, addColumnHeads = c("fromMedian", "toMedian", "fromMean", "toMean", "Vstat", "pValue"))))
    }))
    rownames(outFrame) <- names(curSpeciesList)
    outFrame
  }
  # Function to write the pairwise yearly analysis to a worksheet
  writePairwiseYearsSheet <- function(freqTable, siteWorkbook, sheetName, speciesInfo) {
    # Retrieve the vector of comparrison years
    freqYears <- unique(gsub("_.*$", "", colnames(freqTable), perl = TRUE))
    # Initialise a comparrison years vector
    freqYearHeaders <- rep("", 6 * length(freqYears))
    freqYearHeaders[0:(length(freqYearHeaders) - 1) %% 6 == 0] <- paste(freqYears, "Comparison", sep = " ")
    dim(freqYearHeaders) <- c(1, length(freqYearHeaders))
    freqYearHeaders <- cbind("Species", freqYearHeaders)
    # Initialise the species row names
    freqSpecies <- speciesInfo[rownames(freqTable)]
    dim(freqSpecies) <- c(length(freqSpecies), 1)
    # Initialise a column headings vector
    columnHeadings <- c("Median", "Median", "Mean", "Mean", "Wilcoxon's V", "P Value")
    repColumnHeadings <- unlist(lapply(X = freqYears, FUN = function(curYearComp, columnHeadings) {
      # Retrieve the years being compared
      inYears <- c(paste(" (", strsplit(curYearComp, "-", fixed = TRUE)[[1]], ")", sep = ""), "")
      # Add the year information to the relevant columns
      paste(columnHeadings, inYears[c(1, 2, 1, 2, 3, 3)], sep = "")
    }, columnHeadings = columnHeadings))
    dim(repColumnHeadings) <- c(1, length(repColumnHeadings))
    # Write the elements of the worksheet
    writeData(siteWorkbook, sheetName, freqYearHeaders, rowNames = FALSE, colNames = FALSE, startCol = 1, startRow = 1)
    writeData(siteWorkbook, sheetName, repColumnHeadings, rowNames = FALSE, colNames = FALSE, startCol = 2, startRow = 2)
    writeData(siteWorkbook, sheetName, freqSpecies, rowNames = FALSE, colNames = FALSE, startCol = 1, startRow = 3)
    writeData(siteWorkbook, sheetName, freqTable, rowNames = FALSE, colNames = FALSE, startCol = 2, startRow = 3)
    # Add conditional formatting to the Wilcox p-values when significant
    sigIncStyle <- createStyle(bgFill = rgb(193, 255, 193, maxColorValue = 255))
    sigDecStyle <- createStyle(bgFill = rgb(255, 182, 193, maxColorValue = 255))
    lapply(X = 1 + length(columnHeadings) + 0:(length(freqYears)) * length(columnHeadings), FUN = function(curCol, numSpecies, siteWorkbook, sheetName, sigIncStyle, sigDecStyle) {
      conditionalFormatting(siteWorkbook, sheetName, cols = curCol, rows = 1:numSpecies + 2, style = sigIncStyle, rule = paste("AND(", int2col(curCol), "3<=0.05,", int2col(curCol - 3), "3<", int2col(curCol - 2), "3)", sep = ""))
      conditionalFormatting(siteWorkbook, sheetName, cols = curCol, rows = 1:numSpecies + 2, style = sigDecStyle, rule = paste("AND(", int2col(curCol), "3<=0.05,", int2col(curCol - 3), "3>", int2col(curCol - 2), "3)", sep = ""))
    }, numSpecies = length(freqSpecies), siteWorkbook = siteWorkbook, sheetName = sheetName, sigIncStyle = sigIncStyle, sigDecStyle = sigDecStyle)
    # Merge the comparison cells
    lapply(X = 1:length(freqYears), FUN = function(curIter, colLength, siteWorkbook, sheetName) {
      mergeCells(siteWorkbook, sheetName, cols = ((curIter - 1) * colLength):(curIter * colLength - 1) + 2, rows = 1)
    }, colLength = length(columnHeadings), siteWorkbook, sheetName)
    # Style the column headers
    topHeadStyle <- createStyle(textDecoration = "bold")
    lowerHeadStyle <- createStyle(border = "bottom", borderStyle = "medium")
    addStyle(siteWorkbook, sheetName, style = topHeadStyle, rows = 1, cols = 1:ncol(freqYearHeaders))
    addStyle(siteWorkbook, sheetName, style = lowerHeadStyle, rows = 2, cols = 1:ncol(freqYearHeaders))
    # Style the species names
    speciesStyle <- createStyle(textDecoration = "italic")
    addStyle(siteWorkbook, sheetName, style = speciesStyle, rows = 1:nrow(freqSpecies) + 2, cols = 1)
    siteWorkbook
  }
  # Flatten out the cover and frequency analyses
  freqTable <- rearrangePairwiseYears(freqPairwise[[curSiteCode]])
  writePairwiseYearsSheet(freqTable, siteWorkbook, "Frekvens", speciesInfo)
  coverTable <- rearrangePairwiseYears(coverPairwise[[curSiteCode]])
  writePairwiseYearsSheet(coverTable, siteWorkbook, "Dekning", speciesInfo)
  # Save the workbook
  saveWorkbook(siteWorkbook, paste(outputLocation, "/", fullSiteName, "_yearlyPairwise.xlsx", sep = ""), overwrite = TRUE)
}, freqPairwise = compressedFreq_yearlyPairwise, coverPairwise = compressedCover_yearlyPairwise, siteInfo = TOVData$siteInfo, outputLocation = outputLocation,
  # Arrange the species information in an easy lookup vector
  speciesInfo = setNames(sapply(X = colnames(compressedFreq), FUN = function(curSpecCode, speciesInfo) {
    unique(speciesInfo[gsub("_[A-Z]$", "", rownames(speciesInfo), perl = TRUE) == curSpecCode, "SpeciesBinomial"])[1]
  }, speciesInfo = TOVData$speciesInfo), colnames(compressedFreq)))
```
