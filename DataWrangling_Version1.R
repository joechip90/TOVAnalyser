# This file is for the import of data from the series of summary Excel files that are produced from
# from the entries in the TOV database.  This section will eventually be replaced with a version that
# accesses the TOV database directly.

library(openxlsx)       # Import the files for import/export of Excel data

# Folder containing the data summary files
inputFolder <- "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/DataSummaryFiles"
# Find the summary files within the folder
inputFiles <- dir(inputFolder)
# Output file
outputFile <- "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/CleanedData.rds"

# Set aliases for each of the species information columns
speciesInfoAliases <- list(
  SpeciesBinomial = c("GjeldendeFloranavnNINA"),
  NorwegianName = c("PopulaernavnBokmaalADB"),
  NINAName = c("NorskNavnNINA"),
  VegLayer = c("Sjikt", "Row.Labels"),
  FieldCodeName = c("KODENAVN_FELT", "KODENAVN_NINA")
)

# Set aliases for each of the NINA field codes
fieldCodeAliases <- list(
  "Aven fle" = c("Desc fle", "Aven fle"),
  "Cet isla" = c("Ceta isl", "Cet isla"),
  "Clad unc" = c("Clad unc", "Cla unci"),
  "Flav niv" = c("Flav niv", "Fla niva")
)

# Set the information associated with each field site
fieldSiteInfo <- as.data.frame(matrix(dimnames = list(NULL, c(
  "SiteName", "SiteCode")), data = c(
  "Dividalen", "D",
  "Gutuila", "G"
), ncol = 2, byrow = TRUE))
rownames(fieldSiteInfo) <- fieldSiteInfo$SiteCode

# Set the information associated with each vegetation layer
sjiktInfo <- as.data.frame(matrix(dimnames = list(NULL, c(
  "SjiktCode", "DescriptionNorsk")), data = c(
  "A", "Trær",
  "B", "Busker",
  "C", "Lyng og dvergbusker",
  "D", "Urter",
  "E", "Gras og halvgras",
  "F", "Bladmoser",
  "G", "Levermoser",
  "H", "Busklav"
), ncol = 2, byrow = TRUE))
rownames(sjiktInfo) <- sjiktInfo$SjiktCode

# Function that uses an alias list to find the relevant columns of an input data sheet
aliasColumnLookup <- function(inputSheet, aliasList) {
  outputSheet <- inputSheet[, sapply(X = aliasList, FUN = function(curAlias, inputSheet) {
    which(colnames(inputSheet) %in% curAlias)[1]
  }, inputSheet = inputSheet)]
  colnames(outputSheet) <- names(aliasList)
  outputSheet
}

# Define a function to import the community matrices
importCommunityMatrices <- function(curFile, inputFolder) {
  # Load the workbook
  curWorkbookName <- paste(inputFolder, curFile, sep = "/")
  # Retrieve a list of sheet names
  curSheetNames <- getSheetNames(curWorkbookName)
  # Get the names of the sheet
  frekvensSheetName <- curSheetNames[grepl(".*rute.*", tolower(curSheetNames), perl = TRUE) & grepl(".*frekvens.*", tolower(curSheetNames), perl = TRUE)]
  dekningSheetName <- curSheetNames[grepl(".*rute.*", tolower(curSheetNames), perl = TRUE) & grepl(".*deknin*g.*", tolower(curSheetNames), perl = TRUE)]
  # Retrieve the frequency and cover data sheets
  list(
    frekvensSheet = read.xlsx(paste(inputFolder, curFile, sep = "/"), sheet = frekvensSheetName, startRow = 3),
    dekningSheet = read.xlsx(paste(inputFolder, curFile, sep = "/"), sheet = dekningSheetName, startRow = 3))
}
# Import the community matrices from each of the input files
communityMatrixList <- lapply(X = inputFiles, FUN = importCommunityMatrices, inputFolder = inputFolder)
# Get the list of species and their associated information
speciesInfo <- do.call(rbind, lapply(X = communityMatrixList, FUN = function(curMatList, speciesInfoAliases) {
    unique(rbind(
      aliasColumnLookup(curMatList$frekvensSheet, speciesInfoAliases),
      aliasColumnLookup(curMatList$dekningSheet, speciesInfoAliases)))
}, speciesInfoAliases = speciesInfoAliases))
# Replace any transcription errors in the NINA species field code
speciesInfo$FieldCodeName <- sapply(X = as.character(speciesInfo$FieldCodeName), FUN = function(curFieldCode, fieldCodeAliases) {
  # Test to see if the field code is one of the ones with multiple aliases
  isAlias <- sapply(X = fieldCodeAliases, FUN = function(curAlias, curFieldCode) {
    curFieldCode %in% curAlias
  }, curFieldCode = curFieldCode)
  outValue <- curFieldCode
  if(any(isAlias)) {
    # If it is an alias then replace it with the proper field name
    outValue <- names(fieldCodeAliases)[isAlias]
  }
  outValue
}, fieldCodeAliases = fieldCodeAliases)
# Sort the species information by vegetation layer and then species name
speciesInfo <- unique(speciesInfo[order(speciesInfo$VegLayer, speciesInfo$SpeciesBinomial), ])
# Create unique IDs for each species based on the species and the vegetation layer
rownames(speciesInfo) <- paste(gsub("\\s+", "_", speciesInfo$FieldCodeName, perl = TRUE), speciesInfo$VegLayer, sep = "_")
# Replace "(blank)" with NA values
speciesInfo$NorwegianName <- as.factor(ifelse(as.character(speciesInfo$NorwegianName) == "(blank)", NA, as.character(speciesInfo$NorwegianName)))

# Retrieve all the plot codes in the data files
plotInfo <- data.frame(PlotCode = sort(unique(unlist(lapply(X = communityMatrixList, FUN = function(curComMats) {
  # Retrieve the column names from the frequency and cover sheets
  sheetColNames <- unique(c(colnames(curComMats$frekvensSheet), colnames(curComMats$dekningSheet)))
  # Remove those column names that don't correspond to a plot ID code
  sheetColNames[grepl("^[A-Z]\\d+[A-Z]*\\-+\\d+$", sheetColNames, perl = TRUE)]
})))))
# Pattern match the plot codes to get the relevant sample information
regPlotCodes <- regmatches(plotInfo$PlotCode, regexec("^([A-Z])(\\d+[A-Z]*)\\-+(\\d+)$", plotInfo$PlotCode, perl = TRUE))
plotInfo <- cbind(plotInfo, data.frame(
  SiteCode = sapply(X = regPlotCodes, FUN = function(curValue) {curValue[2]}),
  PlotNumber = sapply(X = regPlotCodes, FUN = function(curValue) {curValue[3]}),
  Year = as.numeric(sapply(X = regPlotCodes, FUN = function(curValue) {curValue[4]}))
))
rownames(plotInfo) <- plotInfo$PlotCode

# Function to retrieve the matrix values
getMatrixValues <- function(curElement, sheetName, speciesInfo, plotInfo, fieldCodeAliases) {
  # Retrieve the current frequency/cover sheet
  curSheet <- curElement[[sheetName]]
  sheetSpeciesInfo <- aliasColumnLookup(curSheet, speciesInfoAliases[c("FieldCodeName", "VegLayer")])
  # Replace the field code with its alias if it is in the list of species with a known alias
  sheetSpeciesInfo$FieldCodeName <- sapply(X = sheetSpeciesInfo$FieldCodeName, FUN = function(curFieldCode, fieldCodeAliases) {
    outVal <- curFieldCode
    hasAlias <- sapply(X = fieldCodeAliases, FUN = function(curAlias, curFieldCode) {
      curFieldCode %in% curAlias
    }, curFieldCode = curFieldCode)
    if(any(hasAlias)) {
      outVal <- names(fieldCodeAliases)[hasAlias]
    }
    outVal
  }, fieldCodeAliases = fieldCodeAliases)
  # Set row names of the current frequency/cover sheet
  rownames(curSheet) <- paste(gsub(" ", "_", sheetSpeciesInfo$FieldCodeName, fixed = TRUE), sheetSpeciesInfo$VegLayer, sep = "_")
  # Initialise an output matrix
  outMat <- matrix(data = 0, nrow = nrow(speciesInfo), ncol = nrow(plotInfo), dimnames = list(rownames(speciesInfo), rownames(plotInfo)))
  # Find the species and plot IDs that exist in the current sheet
  inSpecies <- rownames(speciesInfo)[rownames(speciesInfo) %in% rownames(curSheet)]
  inPlots <- rownames(plotInfo)[rownames(plotInfo) %in% colnames(curSheet)]
  # Set the relevant output elements
  outMat[inSpecies, inPlots] <- as.matrix(curSheet[inSpecies, inPlots])
  outMat[is.na(outMat)] <- 0
  outMat
}
# Calculate the frequency matrix
freqMatrix <- do.call(`+`, lapply(X = communityMatrixList, FUN = getMatrixValues, sheetName = "frekvensSheet", speciesInfo = speciesInfo, plotInfo = plotInfo, fieldCodeAliases = fieldCodeAliases))
# Calculate the cover matrix
coverMatrix <- do.call(`+`, lapply(X = communityMatrixList, FUN = getMatrixValues, sheetName = "dekningSheet", speciesInfo = speciesInfo, plotInfo = plotInfo, fieldCodeAliases = fieldCodeAliases))

# Save the community matrices and the plot and species information to the output file
if(file.exists(outputFile)) {
  unlink(outputFile)
}
saveRDS(list(
  # Save the community matrices
  freqMatrix = freqMatrix, coverMatrix = coverMatrix,
  # Save the information frames
  siteInfo = fieldSiteInfo, plotInfo = plotInfo, speciesInfo = speciesInfo, sjiktInfo = sjiktInfo
), file = outputFile)
