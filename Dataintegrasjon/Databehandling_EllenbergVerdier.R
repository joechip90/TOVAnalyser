library(openxlsx)       # Import the files for import/export of Excel data

# 1. ---- DATA IMPORT SETTINGS ----
# This source file relies on the setting of the "WORKSPACE_TOV" variable in the .Renviron file.  This variables should contain the
# local directory of the TOVAnalyser repository.  This variable can be set easily using the usethis::edit_r_environ() function.
ellenbergFolder <- paste(Sys.getenv("WORKSPACE_TOV"), "EllenbergVerdier", sep = "/")
# File containing the Ellenberg indicator values
ellenbergFiles <- paste(ellenbergFolder, c(
  "EllenbergValues_PAkomm.xlsx",
  "engelskeellenbergverdier.xlsx",
  "Bryoatt_updated_2017.xlsx",
  "ellenbergkrypt1.xlsx",
  "ellenbergkrypt2.xlsx"
), sep = "/")
# Set the output fold to store the processed data
outputFolder <- paste(Sys.getenv("WORKSPACE_TOV"), "Behandlede_datatabeller", sep = "/")
# Set the headings of the Ellenberg indicator files that contain species names
ellenbergTaxCodeCol <- c("species", "Taxon.name", "Genus_g", "Species_s", "SpeciesBinomial")
# Set the Ellenberg indicator values to import
ellenbergCols <- c("L", "F", "R", "N", "S")
# Set the Ellenberg sheet names where the ellenberg information is to be kept
ellenbergSheetNames <- c("BRYOATT", "britellenberg for R", "lav", "moser", "Species")
# Ellenberg over-rides - the automatic detection routines give incorrect Ellenberg
# values for some species so we override some of those detection routines here
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

# 2. ---- IMPORT RAW ELLENBERG VALUES FROM SOURCE FILES ----
# Import the Ellenberg values from the series of files containing them
ellenbergInput <- do.call(rbind, lapply(X = ellenbergFiles, FUN = function(curFile, ellenbergCols, ellenbergTaxCodeCol, ellenbergSheetNames) {
  # Retrieve the sheet names
  sheetNames <- getSheetNames(curFile)
  # Read the relevant sheet
  curSheetName <- sheetNames[sheetNames %in% ellenbergSheetNames][1]
  curSheet <- read.xlsx(curFile, sheet = curSheetName)
  # Tidy up the sheet's missing values
  ellMatrix <- matrix(apply(X = as.matrix(curSheet[, ellenbergCols]), FUN = function(curCol) {
    tidiedCol <- gsub("\\s+", "", curCol, perl = TRUE)
    as.numeric(ifelse(grepl("\\D+", tidiedCol, perl = TRUE), NA, tidiedCol))
  }, MARGIN = 2), ncol = length(ellenbergCols), dimnames = list(NULL, ellenbergCols))
  # Get the taxa names
  curTaxaNames <- do.call(paste, lapply(X = 1:sum(colnames(curSheet) %in% ellenbergTaxCodeCol), FUN = function(curCol, inMatrix) {
    as.character(inMatrix[, curCol])
  }, inMatrix = as.matrix(curSheet[, colnames(curSheet) %in% ellenbergTaxCodeCol, drop = FALSE])))
  cbind(
    data.frame(
      taxaName = curTaxaNames,
      sourceFile = rep(curFile, length(curTaxaNames)),
      sourceSheet = rep(curSheetName, length(curTaxaNames))
    ), as.data.frame(ellMatrix))
}, ellenbergCols = ellenbergCols, ellenbergTaxCodeCol = ellenbergTaxCodeCol, ellenbergSheetNames = ellenbergSheetNames))
ellenbergInput$taxaName <- gsub(".", "", gsub("?", "", gsub("_", " ", ellenbergInput$taxaName, fixed = TRUE), fixed = TRUE), fixed = TRUE)
# Remove any duplicated entries in the Ellenberg data
ellenbergInput <- ellenbergInput[!duplicated(ellenbergInput$taxaName) & !is.na(ellenbergInput$taxaName) & ellenbergInput$taxaName != "NA", ]
rownames(ellenbergInput) <- gsub("\\s+", "_", ellenbergInput$taxaName, perl = TRUE)

# 3. ---- SAVE THE PROCESSED ELLENBERG VALUES ----
write.csv2(ellenbergInput, file = paste(outputFolder, "ellenbergVerdier.csv", sep = "/"))
