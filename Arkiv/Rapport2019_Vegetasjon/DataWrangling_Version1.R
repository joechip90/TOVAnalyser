# This file is for the import of data from the series of summary Excel files that are produced from
# from the entries in the TOV database.  This section will eventually be replaced with a version that
# accesses the TOV database directly.

library(openxlsx)       # Import the files for import/export of Excel data

# Folder containing the data summary files
inputFolder <- "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/DataSummaryFiles"
# File containing the Ellenberg indicator values
ellenbergFiles <- c(
  "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/TraitData/EllenbergValues_PAkomm.xlsx",
  "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/TraitData/engelskeellenbergverdier.xlsx",
  "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/TraitData/Bryoatt_updated_2017.xlsx",
  "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/TraitData/ellenbergkrypt1.xlsx",
  "C:/Users/joseph.chipperfield/OneDrive - NINA/Work/TOV/TraitData/ellenbergkrypt2.xlsx"
)
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
  "Gutulia", "G"
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

# Set the headings of the Ellenberg indicator file
ellenbergTaxCodeCol <- c("species", "Taxon.name", "Genus_g", "Species_s", "SpeciesBinomial")
# Set the Ellenberg indicator values to import
ellenbergCols <- c("L", "F", "R", "N", "S")
# Set the Ellenberg sheet names
ellenbergSheetNames <- c("BRYOATT", "britellenberg for R", "lav", "moser", "Species")
# Ellenberg over-rides - the automatic detection routines give incorrect Ellenberg
# values for some species
ellenbergOverride <- matrix(c(
#  "Arct_alp_C", NA,
#  "Sali_has_C", NA,
#  "Hier/alp_D", NA,
#  "Hier/Hiu_D", NA,
#  "Hieraciz_D", NA,
#  "C_brunne_E", NA,
  "Rhyt/squ_F", "Rhytidiadelphus subpinnatus",
  "Sciu_sta_F", "Brachythecium starkei",
  "Clad/sul_H", "Cladonia sulphurina"
), ncol = 2, byrow = TRUE)



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

# Import the ellenberg values and append them to the species information
ellenbergInput <- do.call(rbind, lapply(X = ellenbergFiles, FUN = function(curFile, ellenbergCols, ellenbergTaxCodeCol, ellenbergSheetNames) {
  # Retrieve the sheet names
  sheetNames <- getSheetNames(curFile)
  # Read the relevant sheet
  curSheet <- read.xlsx(curFile, sheet = sheetNames[sheetNames %in% ellenbergSheetNames][1])
  # Get the taxa names
  curTaxaNames <- do.call(paste, lapply(X = 1:sum(colnames(curSheet) %in% ellenbergTaxCodeCol), FUN = function(curCol, inMatrix) {
    as.character(inMatrix[, curCol])
  }, inMatrix = as.matrix(curSheet[, colnames(curSheet) %in% ellenbergTaxCodeCol, drop = FALSE])))
  cbind(data.frame(taxaName = curTaxaNames, sourceFile = rep(curFile, length(curTaxaNames))), curSheet[, ellenbergCols])
}, ellenbergCols = ellenbergCols, ellenbergTaxCodeCol = ellenbergTaxCodeCol, ellenbergSheetNames = ellenbergSheetNames))
ellenbergInput$taxaName <- gsub("_", " ", ellenbergInput$taxaName, fixed = TRUE)
# Remove any duplicated entries in the Ellenberg data
ellenbergInput <- ellenbergInput[!duplicated(ellenbergInput$taxaName), ]
rownames(ellenbergInput) <- gsub("\\s+", "_", ellenbergInput$taxaName, perl = TRUE)

# Create a matrix of Ellenberg values
ellenbergValMatrix <- t(sapply(X = speciesInfo$SpeciesBinomial, FUN = function(curName, ellenbergInput, warnValue) {
  # Calculate the Levenshtein distance between the taxa names
  diffDist <- adist(
    paste(strsplit(curName, "\\s+", perl = TRUE)[[1]][1:2], collapse = " "),
    ellenbergInput$taxaName
  )
  # Find the shortest distance
  minVal <- min(diffDist)
  bestIndeces <- which(diffDist == minVal)
  # Retrieve the names that refer to Ellenberg values
  ellbergColNames <- colnames(ellenbergInput)[!(colnames(ellenbergInput) %in% c("taxaName", "sourceFile"))]
  outVector <- setNames(rep(as.character(NA), length(ellbergColNames) + 2), c(ellbergColNames, "isMatched", "nearestEllenbergMatch"))
  outVector["nearestEllenbergMatch"] <- paste(ellenbergInput$taxaName[bestIndeces], collapse = ", ")
  outVector["isMatched"] <- "No"
  if(minVal > warnValue) {
    warning(paste("poor match for best matching species name for", curName, "in the Ellenberg files:", paste(ellenbergInput$taxaName[bestIndeces], collapse = ", "), sep = " "))
  } else if(length(bestIndeces) > 1) {
    warning(paste("multiple best matching species names for", curName, "in the Ellenberg files:", paste(ellenbergInput$taxaName[bestIndeces], collapse = ", "), sep = " "))
  } else {
    # Copy the Ellenberg values across
    outVector[ellbergColNames] <- as.character(ellenbergInput[bestIndeces, ellbergColNames])
    outVector["isMatched"] <- "Yes"
  }
  outVector
}, ellenbergInput = ellenbergInput, warnValue = 10))
colnames(ellenbergValMatrix) <- c(ellenbergCols, "isMatched", "nearestEllenbergMatch")
rownames(ellenbergValMatrix) <- rownames(speciesInfo)
# Over-ride Ellenberg values for entries do not have a species-level ID
# ellenbergValMatrix[grepl("\\s+sp\\.", speciesInfo$SpeciesBinomial, perl = TRUE), ellenbergCols] <- NA
# Over-ride Ellenberg values for entries that have specified over-rides
ellenbergValMatrix[ellenbergOverride[, 1], ellenbergCols] <- t(sapply(X = ellenbergOverride[, 2], FUN = function(curSpec, ellenbergInput, ellenbergCols) {
  outVec <- setNames(as.character(rep(NA, length(ellenbergCols))), ellenbergCols)
  if(!is.na(curSpec)) {
    outVec <- as.character(ellenbergInput[ellenbergInput$taxaName == curSpec, ellenbergCols])
  }
  outVec
}, ellenbergInput = ellenbergInput, ellenbergCols = ellenbergCols))
ellenbergValMatrix[ellenbergOverride[, 1], "isMatched"] <- ifelse(is.na(ellenbergOverride[, 2]), "No", "Yes")
ellenbergMatchFrame <- as.data.frame(ellenbergValMatrix[, c("isMatched", "nearestEllenbergMatch")])
# Convert the Ellenber values to integer
ellenbergValMatrix <- apply(X = ellenbergValMatrix[, ellenbergCols], FUN = function(curColumn) {
  # Remove any non-digit character
  outCurColumn <- gsub("[\\D]+", "", curColumn, perl = TRUE)
  as.integer(ifelse(outCurColumn == "", NA, outCurColumn))
}, MARGIN = 2)
colnames(ellenbergValMatrix) <- ellenbergCols
# Append to the species info
speciesInfo <- cbind(speciesInfo, as.data.frame(ellenbergValMatrix), ellenbergMatchFrame)

# Tidy up the Ellenberg values in the input data frame
ellenbergInput[, ellenbergCols] <- apply(X = as.matrix(ellenbergInput[, ellenbergCols]), FUN = function(curColumn) {
  outCurColumn <- tolower(gsub("\\s+", "", curColumn, perl = TRUE))
  outCurColumn <- ifelse(grepl("[^x\\d]+", outCurColumn, perl = TRUE), "", outCurColumn)
  ifelse(outCurColumn == "", NA, outCurColumn)
}, MARGIN = 2)

# Save the community matrices and the plot and species information to the output file
if(file.exists(outputFile)) {
  unlink(outputFile)
}
saveRDS(list(
  # Save the community matrices
  freqMatrix = freqMatrix, coverMatrix = coverMatrix,
  # Save the information frames
  siteInfo = fieldSiteInfo, plotInfo = plotInfo, speciesInfo = speciesInfo, sjiktInfo = sjiktInfo,
  # Save the compiled Ellenberg data (for all species not just those present in the sites)
  ellenbergData = ellenbergInput
), file = outputFile)
