library(RODBC)
library(openxlsx)
library(vegan)
library(ggplot2)
library(reshape2)
library(INLA)

# ------ 1. SET REPORT GENERATION CRITERIA ------
# Sites to produce analyses for.  Is a list with names elements, one for each site. Each element is a character vector containing site codes
# used for the site in the TOV database (some sites may have multiple synonyms)
sitesToProcess <- list(
  "Børgefjell" = c("B")
)
# Years to produce analyses for (including years you wish to generate time series of species for)
yearsToProcess <- c(1995, 2000, 2005, 2010, 2015, 2020)
# This source file relies on the setting of the "WORKSPACE_TOV" variable in the .Renviron file.  This variable should contain the
# directory of the TOV project directory.  This variable can be set easily using the usethis::edit_r_environ() function.
tovDirectory <- Sys.getenv("WORKSPACE_TOV")
# Set the output folder for the analysis
outputDirectory <- paste(tovDirectory, "Behandlede_datatabeller", "RapportAnalyser_2020Vegetasjon", sep = "/")
# Ellenberg column names to use in the Ellenberg analysis
ellenbergColNames <- c("L", "F", "R", "N")
ellenbergDescriptionNorsk <- setNames(c("Lys", "Fuktighet", "Reaksjon (baserikhet)", "Nitrogen"), ellenbergColNames)
# Set a series of species to be considered aliases (merges them for the sake of analyses)
speciesAliases <- list(
  "Anthoxanthum_sp" = c("Anthoxanthum_nipponicum", "Anthoxanthum_odoratum_coll"),
  "Cladonia_arbuscula" = c("Cladonia_arbuscula", "Cladonia_arbuscula_ssp_arbuscula"),
  "Sciuro-hypnum_reflexum" = c("Sciuro-hypnum_reflexum", "Brachythecium_reflexum"),
  "Sciuro-hypnum_starkei" = c("Sciuro-hypnum_starkei", "Brachythecium_starkei")
)
# Over-ride the sjikt status for some species (they are the wrong sjikt in the database)
sjiktOverride <- setNames(
  c("G", "G"),
  c("Ptilidium_ciliare", "Calypogeia_sp")
)
# Sjikt summary groupings (how should the sjiktene be grouped for plotting purposes?)
sjiktGroups <- list(
  c("D", "E", "C"),
  c("F", "G", "H")
)
sjiktGroups <- lapply(X = sjiktGroups, FUN = sort)
# Set aliases for the same component in the damage information file
extraAliases <- list(
  "Frostkader på blåbær" = c("Frostskade på blåbær", "Frostskader på blåbær"),
  "Podosphaera/Pucciniastrum" = c("Podosphaera/Pucciniastrum", "Podosphaera")
)

# ------ 2. IMPORT DATA PROCESSING SCRIPTS ------
# This source file relies on the setting of the "WORKSPACE_TOV_GITHUB" variable in the .Renviron file.  This variable should contain the
# local directory of the TOVAnalyser repository.  This variable can be set easily using the usethis::edit_r_environ() function.
source(paste(Sys.getenv("WORKSPACE_TOV_GITHUB"), "Dataintegrasjon", "Datainnsamling.R", sep = "/"))
# Make the directory to store the restructured data
if(dir.exists(outputDirectory)) {
  unlink(outputDirectory, recursive = TRUE)
}
dir.create(outputDirectory)

# ------ 3. RETRIEVE DATA ------
# Import the Ellenberg values
ellenbergInput <- read.csv2(paste(tovDirectory, "Behandlede_datatabeller", "ellenbergVerdier.csv", sep = "/"), row.names = 1)
# Retrieve the vegetation data
rawVegetationData <- retrieveTOVData(c("VEG_Analyse_rutenivaa", "VEG_Analyseinfo", "VEG_Kjemi_data_pr_rute", "VEG_Kjemi_data_pr_felt"))
# Import the blåbær and gnagere dataset
extraInfoLoc <- paste(tovDirectory, "Skade blåbær_smågnager", "Skade blåbær 2012-2020.xlsx", sep = "/")

# ------ 4. RESTRUCTURE DATA FOR REPORT ------
# Produce a site information frame
siteInfo <- data.frame(
  SiteName = names(sitesToProcess),
  SiteCode = sapply(X = sitesToProcess, FUN = paste, collapse = "|"),
  row.names = names(sitesToProcess)
)
# Produce a sjikt information frame
sjiktInfo <- data.frame(
  SjiktCode = c("A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "S", "T", "U"),
  DescriptionNorsk = c("Trær", "Busker", "Lyng og dvergbusker", "Urter", "Gras og halvgras", "Bladmoser", "Levermoser", "Busklav", "Skorpelav", "Alger", "Sopp", "Kransalge", "Crust"),
  DescriptionEnglish = c("Trees", "Shrubs", "Heather and dwarf shrubs", "Forbs", "Grasses", "Mosses", "Liverworts", "Macrolichens", "Microlichens", "Algae", "Fungi", "Charales", "Crust"),
  row.names = c("A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "S", "T", "U")
)
# Retrieve the plot names that are to be used in the current analysis
allPlotCodes <- sort(unique(as.character(rawVegetationData[[2]]$table$Analyse_navn)))
usablePlotCodes <- allPlotCodes[apply(X = sapply(X = as.character(unlist(sitesToProcess)), FUN = function(curSiteCode, allPlotCodes) {
  grepl(paste("^", curSiteCode, "\\d+", sep = ""), allPlotCodes, perl = TRUE)
}, allPlotCodes = allPlotCodes), FUN = any, MARGIN = 1)]
# Retrieve the plot information that corresponds only to those that are at the specified site
rawPlotInfo <- rawVegetationData[[2]]$table[as.character(rawVegetationData[[2]]$table$Analyse_navn) %in% usablePlotCodes, ]
# Filter out those plots not taken in the requested time period
rawPlotInfo <- rawPlotInfo[rawPlotInfo$AAR %in% yearsToProcess, ]
# In some sites there are weird duplicate plots with malformed plot ID codes (Børhefjell has these for sites in 1995). Remove those
# entrie with malformed plot IDs
rawPlotInfo <- rawPlotInfo[grepl("\\d\\d\\d\\d$", as.character(rawPlotInfo$Analyse_navn), perl = TRUE), ]
# Produce a plot information frame
plotInfo <- data.frame(
  PlotCode = as.character(rawPlotInfo$Analyse_navn),
  SiteName = as.character(sapply(X = gsub("\\d+.*$", "", as.character(rawPlotInfo$Analyse_navn), perl = TRUE), FUN = function(curCode, sitesToProcess) {
    names(sitesToProcess)[sapply(X = sitesToProcess, FUN = function(curSiteCodes, curCode) { any(curSiteCodes == curCode) }, curCode = curCode)]
  }, sitesToProcess = sitesToProcess)),
  PlotNumber = rawPlotInfo$REG_SOM_Rute_id,
  DBPlotNumber = as.integer(rawPlotInfo$Analysenr),
  Year = rawPlotInfo$AAR,
  row.names = as.character(rawPlotInfo$Analyse_navn)
)
# Produce a frame of sampling information
samplingInfo <- setNames(lapply(X = sort(unique(as.character(plotInfo$SiteName))), FUN = function(curSite, plotInfo, allYears) {
  t(sapply(X = sort(unique(as.character(plotInfo$PlotNumber))), FUN = function(curPlot, plotInfo, allYears) {
    setNames(sapply(X = allYears, FUN = function(curYear, plotInfo, curPlot) {
      nrow(plotInfo[as.character(plotInfo$PlotNumber) == curPlot & as.integer(plotInfo$Year) == curYear, ])
    }, plotInfo = plotInfo, curPlot = curPlot), allYears)
  }, plotInfo = plotInfo[as.character(plotInfo$SiteName) == curSite, ], allYears = allYears))
}, plotInfo = plotInfo, allYears = sort(unique(as.integer(plotInfo$Year)))), sort(unique(as.character(plotInfo$SiteName))))
# Retrieve the transect data collected at the appropriate field sites
rawQuadratData <- rawVegetationData[[1]]$table[as.integer(rawVegetationData[[1]]$table$REG_SOM_Analysenr) %in% plotInfo$DBPlotNumber, ]
# Retrieve the unique species found in the time period
specNames <- sort(unique(
  # Some of the species entries in the TOV database have trailing or leading whitspace so remove that first when compiling
  # a list of species names present at the sites
  gsub("\\s+$", "", gsub("^\\s+", "", as.character(rawQuadratData$REG_SOM_NINA_Artsnavn), perl = TRUE), perl = TRUE)
))
# Reorganise the species information
speciesInfo <- as.data.frame(t(sapply(X = specNames, FUN = function(curSpecies, rawSpecInfo) {
  # Retrieve the species information in the database
  curSpecInfo <- rawSpecInfo[
    grepl(paste("^\\s*", curSpecies, "\\s*$", sep = ""), as.character(rawSpecInfo$REG_SOM_NINA_Artsnavn), perl = TRUE),
  ]
  curNINAName <- table(as.character(curSpecInfo$REG_SOM_NINA_NorskNavn))
  curSpecCode <- table(as.character(curSpecInfo$REG_SOM_Kodenavn))
  curSjikt <- table(as.character(curSpecInfo$FK_SJIKT_ID))
  curNINAName <- names(curNINAName)[which.max(curNINAName)][1]
  if(is.null(curNINAName) || length(curNINAName) <= 0) {
    curNINAName <- NA
  }
  curSpecCode <- names(curSpecCode)[which.max(curSpecCode)][1]
  if(is.null(curSpecCode) || length(curSpecCode) <= 0) {
    curSpecCode <- NA
  }
  curSjikt <- paste(names(curSjikt)[curSjikt > 0], collapse = "|")
  if(is.null(curSjikt) || length(curSjikt) <= 0) {
    curSjikt <- NA
  }
  setNames(c(curSpecies, curNINAName, curSpecCode, curSjikt),
    c("SpeciesBinomial", "NameNorsk", "NINASpeciesCode", "SjiktPresence"))
}, rawSpecInfo = rawVegetationData[[1]]$table)))
row.names(speciesInfo) <- gsub(".", "", gsub(" ", "_", speciesInfo$SpeciesBinomial, fixed = TRUE), fixed = TRUE)
# Function to match species in the TOV database with the names in the Ellenberg files using fuzzy matching (to account
# for misspleelings in the TOV database)
matchEllenberg <- function(curName, ellenbergNames, numDefs) {
  specRegEx <- "\\s+sp\\.*\\s*$"
  aggRegEx <- "\\s+agg\\.*\\s*$"
  collRegEx <- "\\s+coll\\.*\\s*$"
  sspRegExOne <- "\\s+ssp\\.*$"
  sspRegExTwo <- "\\s+ssp\\.*\\s+.*$"
  outIndex <- NA
  fuzzyTolerance <- 0.1
  if(grepl(specRegEx, curName, perl = TRUE) | grepl(aggRegEx, curName, perl = TRUE) | grepl(collRegEx, curName, perl = TRUE)) {
    # The current species name has a "sp." abbreviation so has only been identified to genus level (also works for "coll." and "ag." suffixes)
    # ... check if there is an exact match in the Ellenberg file
    outIndex <- which(gsub("\\.$", "", curName, perl = TRUE) == gsub("\\.$", "", ellenbergNames, perl = TRUE))
    if(length(outIndex) <= 0) {
      # ... otherwise just take the ellenbergNames that represent genera-level assessment and do fuzzy matching on them
      # Find those ellenbergNames that are also aggegrates
      aggIndeces <- which(grepl(specRegEx, ellenbergNames, perl = TRUE) | grepl(aggRegEx, ellenbergNames, perl = TRUE) | grepl(collRegEx, ellenbergNames, perl = TRUE))
      outIndex <- aggIndeces[agrep(
        gsub(collRegEx, "", gsub(aggRegEx, "", gsub(specRegEx, "", gsub("\\.$", "", curName, perl = TRUE), perl = TRUE), perl = TRUE), perl = TRUE),
        gsub(collRegEx, "", gsub(aggRegEx, "", gsub(specRegEx, "", ellenbergNames[aggIndeces], perl = TRUE), perl = TRUE), perl = TRUE),
        max.distance = fuzzyTolerance
      )]
    }
  } else {
    # The current species seems to be identified to species level
    # ... check if there is an exact match in the Ellenberg file
    if(any(curName == ellenbergNames)) {
      outIndex <- which(curName == ellenbergNames)
    } else {
      # ... otherwise remove any sub-sepecies information and then do fuzzy matching
      trimmedName <- gsub(sspRegExOne, "", gsub(sspRegExTwo, "", curName, perl = TRUE), perl = TRUE)
      # Break up any species names that have a "/" character and then search for combinations of them
      specCombs <- as.matrix(do.call(expand.grid, append(lapply(X = strsplit(trimmedName, "\\s+", perl = TRUE)[[1]], FUN = function(curElement) {
        strsplit(curElement, "/", fixed = TRUE)[[1]]
      }), list(stringsAsFactors = FALSE))))
      # Find the indeces that match
      outIndex <- unlist(apply(X = specCombs, FUN = function(curName, trimmedEllenberg, max.distance) {
        agrep(trimmedName, gsub(sspRegExOne, "", gsub(sspRegExTwo, "", ellenbergNames, perl = TRUE), perl = TRUE), max.distance = max.distance)
      }, trimmedEllenberg = gsub(sspRegExOne, "", gsub(sspRegExTwo, "", ellenbergNames, perl = TRUE), perl = TRUE), max.distance = fuzzyTolerance, MARGIN = 1))
    }
  }
  outIndex <- unique(outIndex)
  if(length(outIndex) <= 0) {
    outIndex <- NA
  } else if(length(outIndex) > 1) {
    # If there were multiple matches then do another fuzzy match against the raw string (minus sub-species info) and pick the closest match
    calcDists <- adist(
      gsub(sspRegExTwo, "", gsub(sspRegExOne, "", gsub("\\.$", "", curName, perl = TRUE), perl = TRUE), perl = TRUE),
      gsub(sspRegExTwo, "", gsub(sspRegExOne, "", ellenbergNames[outIndex], perl = TRUE), perl = TRUE)
    )
    outIndex <- outIndex[which(calcDists == min(calcDists))]
    # If there are *STILL* multiple matches then use the one that has the most Ellenberg information
    if(length(outIndex) > 1) {
      outIndex <- outIndex[which.max(numDefs[outIndex])]
    }
  }
  # Convert the index to the Ellenberg species name
  outValue <- NA
  if(!is.na(outIndex)) {
    outValue <- ellenbergNames[outIndex]
  }
  outValue
}
# Match the species names in the TOV database to their nearest Ellenberg equivalent
matchedSpecies <- sapply(X = specNames, FUN = matchEllenberg, ellenbergNames = as.character(ellenbergInput$taxaName),
  numDefs = apply(X = as.matrix(ellenbergInput[, c("L", "F", "R", "N", "S")]), FUN = function(curRow) {
    sum(!is.na(curRow))
  }, MARGIN = 1))
matchedIndeces <- sapply(X = matchedSpecies, FUN = function(curSpec, ellenbergNames) {
  which(curSpec == ellenbergNames)[1]
}, ellenbergNames = as.character(ellenbergInput$taxaName))
speciesInfo <- cbind(speciesInfo, data.frame(
  EllenbergMatchedTaxa = ellenbergInput$taxaName[matchedIndeces],
  EllenbergSourceFile = ellenbergInput$sourceFile[matchedIndeces],
  EllenbergSourceSheet = ellenbergInput$sourceSheet[matchedIndeces],
  L = ellenbergInput$L[matchedIndeces],
  F = ellenbergInput$F[matchedIndeces],
  R = ellenbergInput$R[matchedIndeces],
  N = ellenbergInput$N[matchedIndeces],
  S = ellenbergInput$S[matchedIndeces],
  SpeciesCode = row.names(speciesInfo)
))
# Create a frequency and cover table
qluTable <- function(val, luTable, matchCol) {
  row.names(luTable)[as.character(luTable[, matchCol]) == val][1]
}
quadratData <- data.frame(
  SpeciesCode = sapply(X = as.character(rawQuadratData$REG_SOM_NINA_Artsnavn), FUN = qluTable, luTable = speciesInfo, matchCol = "SpeciesBinomial"),
  PlotCode = sapply(X = as.character(rawQuadratData$REG_SOM_Analysenr), FUN = qluTable, luTable = plotInfo, matchCol = "DBPlotNumber"),
  Cover = rawQuadratData$Dekn,
  Frequency = rawQuadratData$Frekvens,
  Sjikt = as.character(rawQuadratData$FK_SJIKT_ID)
)
quadratData <- quadratData[!is.na(quadratData$PlotCode), ]
# Correct the sjikt for some species
sjiktCorrect <- function(curSjiktRow, sjiktOverride) {
  outSjikt <- curSjiktRow[1]
  if(curSjiktRow[2] %in% names(sjiktOverride)) {
    outSjikt <- sjiktOverride[curSjiktRow[2]]
  }
  outSjikt
}
speciesInfo$SjiktPresence <- apply(X = cbind(
  as.character(speciesInfo$SjiktPresence),
  row.names(speciesInfo)
), FUN = sjiktCorrect, MARGIN = 1, sjiktOverride = sjiktOverride)
quadratData$Sjikt <- apply(X = cbind(
  as.character(quadratData$Sjikt),
  as.character(quadratData$SpeciesCode)
), FUN = sjiktCorrect, MARGIN = 1, sjiktOverride = sjiktOverride)
# Correct the species code (so aliases are assigned to the same species)
speciesCorrect <- function(curSpecies, speciesAliases) {
  isAlias <- sapply(X = speciesAliases, FUN = function(curAlias, curSpecies) {curSpecies %in% curAlias}, curSpecies = curSpecies)
  outSpecies <- curSpecies
  if(any(isAlias)) {
    outSpecies <- names(speciesAliases)[isAlias]
  }
  outSpecies
}
quadratData$SpeciesCode <- sapply(X = as.character(quadratData$SpeciesCode), FUN = speciesCorrect, speciesAliases = speciesAliases)
speciesInfo$SpeciesCode <- sapply(X = as.character(speciesInfo$SpeciesCode), FUN = speciesCorrect, speciesAliases = speciesAliases)
firstNonNAFind <- function(colName, inFrame) {
  inFrame[which(!is.na(inFrame[, colName]))[1], colName]
}
# Copy across the information on the species that are to be merged
for(curSpecies in names(speciesAliases)) {
  isCurSpecies <- speciesInfo$SpeciesCode == curSpecies
  # Produce a new species binomial names based on the merged species
  curSpecBinom <- paste(speciesInfo[isCurSpecies, "SpeciesBinomial"], collapse = "/")
  levels(speciesInfo$SpeciesBinomial) <- c(levels(speciesInfo[isCurSpecies, "SpeciesBinomial"]), curSpecBinom)
  speciesInfo[isCurSpecies, "SpeciesBinomial"] <- curSpecBinom
  # Copy across the sjikt and species information
  speciesInfo[isCurSpecies, "NameNorsk"] <- as.character(firstNonNAFind("NameNorsk", speciesInfo[isCurSpecies, ]))
  speciesInfo[isCurSpecies, "NINASpeciesCode"] <- as.character(firstNonNAFind("NINASpeciesCode", speciesInfo[isCurSpecies, ]))
  speciesInfo[isCurSpecies, "SjiktPresence"] <- as.character(firstNonNAFind("SjiktPresence", speciesInfo[isCurSpecies, ]))
  speciesInfo[isCurSpecies, "EllenbergMatchedTaxa"] <- as.character(firstNonNAFind("EllenbergMatchedTaxa", speciesInfo[isCurSpecies, ]))
  # Copy across the Ellenberg information
  isEllenbergTaxa <- isCurSpecies & speciesInfo[, "EllenbergMatchedTaxa"] == speciesInfo[isCurSpecies, "EllenbergMatchedTaxa"][1]
  speciesInfo[isCurSpecies, "EllenbergSourceFile"] <- speciesInfo[isEllenbergTaxa, "EllenbergSourceFile"][1]
  speciesInfo[isCurSpecies, "EllenbergSourceSheet"] <- speciesInfo[isEllenbergTaxa, "EllenbergSourceSheet"][1]
  speciesInfo[isCurSpecies, "L"] <- speciesInfo[isEllenbergTaxa, "L"][1]
  speciesInfo[isCurSpecies, "F"] <- speciesInfo[isEllenbergTaxa, "F"][1]
  speciesInfo[isCurSpecies, "R"] <- speciesInfo[isEllenbergTaxa, "R"][1]
  speciesInfo[isCurSpecies, "N"] <- speciesInfo[isEllenbergTaxa, "N"][1]
  speciesInfo[isCurSpecies, "S"] <- speciesInfo[isEllenbergTaxa, "S"][1]
}
# Remove duplicates from the species information table
speciesInfo <- speciesInfo[!duplicated(speciesInfo$SpeciesCode), ]
row.names(speciesInfo) <- sapply(X = row.names(speciesInfo), FUN = speciesCorrect, speciesAliases = speciesAliases)
# Create the community and frequency matrices
makePlotSpecMat <- function(quadratData, columnName) {
  plotNames <- sort(unique(as.character(quadratData$PlotCode)))
  speciesNames <- sort(unique(as.character(quadratData$SpeciesCode)))
  t(sapply(X = plotNames, FUN = function(curPlot, speciesNames, quadratData, columnName) {
    setNames(sapply(X = speciesNames, FUN = function(curSpecies, curPlot, quadratData, columnName) {
      matchingData <- quadratData[quadratData$SpeciesCode == curSpecies & quadratData$PlotCode == curPlot, ]
      outValue <- 0
      if(!is.null(matchingData) && nrow(matchingData) > 0) {
        outValue <- sum(matchingData[, columnName], na.rm = TRUE)
      }
      outValue
    }, curPlot = curPlot, quadratData = quadratData, columnName = columnName), speciesNames)
  }, speciesNames = speciesNames, quadratData = quadratData, columnName = columnName))
}
sjiktNames <- sort(unique(as.character(quadratData$Sjikt)))
communityMatrixFreq <- append(setNames(lapply(X = sjiktNames, FUN = function(curSjikt, quadratData, columnName) {
  makePlotSpecMat(quadratData[as.character(quadratData$Sjikt) == curSjikt, , drop = FALSE], columnName)
}, quadratData = quadratData, columnName = "Frequency"), sjiktNames), list(
  all = makePlotSpecMat(quadratData, "Frequency")
))
communityMatrixCover <- append(setNames(lapply(X = sjiktNames, FUN = function(curSjikt, quadratData, columnName) {
  makePlotSpecMat(quadratData[as.character(quadratData$Sjikt) == curSjikt, , drop = FALSE], columnName)
}, quadratData = quadratData, columnName = "Cover"), sjiktNames), list(
  all = makePlotSpecMat(quadratData, "Cover")
))

# ------ 5. PERFORM PAIRWISE YEARLY WILCOXON TESTS ------

# Find the site in which a plot is in from the plot ID
findSiteName <- function(plotID, siteInfo) {
  sapply(X = plotID, FUN = function(curPlotCode, siteCodes) {
    names(siteCodes)[sapply(X = siteCodes, FUN = function(curSiteCode, curPlotCode) {
      grepl(paste("^", curSiteCode, "\\d+", sep = ""), curPlotCode, perl = TRUE)
    }, curPlotCode = curPlotCode)][1]
  }, siteCodes = setNames(strsplit(as.character(siteInfo$SiteCode), "|", fixed = TRUE), siteInfo$SiteName))
}
# Produce an analysis of changes in frequency/cover/ordination in different years
yearlyPairwiseWilcoxon <- function(communityMatrix, siteInfo) {
  # Firstly retrieve the site information at each of the plots
  setNames(lapply(X = as.character(siteInfo$SiteName), FUN = function(curSiteName, communityMatrix, siteInfo) {
    # Retrieve just the components of the community matrix taken at the current site
    communityMatrixSubset <- communityMatrix[findSiteName(row.names(communityMatrix), siteInfo) == curSiteName, ]
    # Species list present at the site
    speciesList <- colnames(communityMatrixSubset)[apply(X = communityMatrixSubset, FUN = function(curCol) {
      any(curCol > 0)
    }, MARGIN = 2)]
    # Make a list with the analysis of each species
    setNames(lapply(X = speciesList, FUN = function(curSpecies, curSiteName, communityMatrixSubset) {
      # Find the set of years that the current site has been visited
      yearsVisited <- sort(as.integer(unique(gsub("^.*\\-", "", row.names(communityMatrixSubset), perl = TRUE))))
      yearMatrix <- matrix(c(yearsVisited[1:(length(yearsVisited) - 1)], yearsVisited[2:length(yearsVisited)]), ncol = 2, dimnames = list(NULL, c("FromYear", "ToYear")))
      # Use the matrix of comparrison years to test for differences between years for each species
      cbind(as.data.frame(yearMatrix), as.data.frame(t(apply(X = yearMatrix, FUN = function(curYears, curSpecies, communityMatrixSubset) {
        # Create two subsets of the community matrix of the 'from' year and the 'to' year
        communityMatrixSubset_fromYear <- communityMatrixSubset[grepl(paste(curYears[1], "$", sep = ""), row.names(communityMatrixSubset), perl = TRUE), ]
        communityMatrixSubset_toYear <- communityMatrixSubset[grepl(paste(curYears[2], "$", sep = ""), row.names(communityMatrixSubset), perl = TRUE), ]
        # Retrieve those plots that are present in both years
        fromPlots <- unique(gsub("\\-\\d+$", "", row.names(communityMatrixSubset_fromYear), perl = TRUE))
        toPlots <- unique(gsub("\\-\\d+$", "", row.names(communityMatrixSubset_toYear), perl = TRUE))
        bothPlots <- sort(c(fromPlots, toPlots)[duplicated(c(fromPlots, toPlots))])
        # Retrieve the values of the community matrix at each year
        values_fromYear <- communityMatrixSubset_fromYear[paste(bothPlots, curYears[1], sep = "-"), curSpecies]
        values_toYear <- communityMatrixSubset_toYear[paste(bothPlots, curYears[2], sep = "-"), curSpecies]
        # Perform the Wilcoxon signed-ranks test
        wilTest <- wilcox.test(values_fromYear, values_toYear, paired = TRUE, exact = FALSE)
        # Produce an output vector with the community data
        setNames(
          c(sum(ifelse(values_fromYear < values_toYear, 1, 0)), sum(ifelse(values_fromYear > values_toYear, 1, 0)), median(values_fromYear), median(values_toYear), mean(values_fromYear), mean(values_toYear), wilTest$statistic, wilTest$p.value),
          c("NumIncrease", "NumDecrease", "FromMedian", "ToMedian", "FromMean", "ToMean", "VStat", "PValue"))
      }, curSpecies = curSpecies, communityMatrixSubset = communityMatrixSubset, MARGIN = 1))))
    }, curSiteName = curSiteName, communityMatrixSubset = communityMatrixSubset), speciesList)
  }, communityMatrix = communityMatrix, siteInfo = siteInfo), as.character(siteInfo$SiteName))
}
yearlySpeciesFreq <- yearlyPairwiseWilcoxon(communityMatrixFreq$all, siteInfo)
yearlySpeciesCover <- yearlyPairwiseWilcoxon(communityMatrixCover$all, siteInfo)
# Produce summary statistics for the cover and frequency over the different years for each species
yearlySummaryStats <- function(communityMatrix, siteInfo, speciesInfo, inTag, plotLoc = NULL) {
  # Firstly retrieve the site information at each of the plots
  outMat <- do.call(cbind, lapply(X = as.character(siteInfo$SiteName), FUN = function(curSiteName, communityMatrix, speciesInfo, siteInfo, inTag, plotLoc) {
    # Retrieve just the components of the community matrix taken at the current site
    communityMatrixSubset <- communityMatrix[findSiteName(row.names(communityMatrix), siteInfo) == curSiteName, ]
    # Find the set of years that the current site has been visited
    yearsVisited <- sort(as.integer(unique(gsub("^.*\\-", "", row.names(communityMatrixSubset), perl = TRUE))))
    # Calculate summary statistics function
    calcYearStat <- function(curYear, communityMatrixSubset, inFunc) {
      setNames(apply(X = communityMatrixSubset[grepl(paste("\\-", curYear, "$", sep = ""), row.names(communityMatrixSubset), perl = TRUE), ], FUN = inFunc, MARGIN = 2), colnames(communityMatrixSubset))
    }
    meanMat <- sapply(X = yearsVisited, FUN = calcYearStat, communityMatrixSubset = communityMatrixSubset, inFunc = mean)
    sdMat <- sapply(X = yearsVisited, FUN = calcYearStat, communityMatrixSubset = communityMatrixSubset, inFunc = sd)
    medianMat <- sapply(X = yearsVisited, FUN = calcYearStat, communityMatrixSubset = communityMatrixSubset, inFunc = median)
    outMat <- cbind(meanMat, sdMat, medianMat)
    colnames(outMat) <- c(
      paste("Mean", inTag, yearsVisited, curSiteName, sep = ""),
      paste("StdDev", inTag, yearsVisited, curSiteName, sep = ""),
      paste("Median", inTag, yearsVisited, curSiteName, sep = "")
    )
    row.names(outMat) <- colnames(communityMatrixSubset)
    # Calculate the temporal correlation coefficient
    outMat <- cbind(outMat, as.data.frame(t(sapply(X = colnames(communityMatrix), FUN = function(curSpecName, curSiteName, outMat, yearsVisited, inTag) {
      corTestOut <- cor.test(yearsVisited, outMat[curSpecName, paste("Mean", inTag, yearsVisited, curSiteName, sep = "")], method = "spearman")
      setNames(c(corTestOut$estimate, corTestOut$p.value), paste(c("SpearmanRho", "SpearmanRhoPValue"), inTag, curSiteName, sep = ""))
    }, curSiteName = curSiteName, outMat = outMat, yearsVisited = yearsVisited, inTag = inTag))))
    if(!is.null(plotLoc)) {
      # Find those species that have significant changes over the years
      speciesToPlot <- row.names(outMat)[outMat$SpearmanRhoPValue < 0.05]
      # Retrieve those sjikt containing species that contain significantly changing species
      plotSjikt <- sort(unique(unlist(strsplit(as.character(speciesInfo[speciesToPlot, "SjiktPresence"]), "|", fixed = TRUE))))
      allPlotFrame <- do.call(rbind, lapply(X = plotSjikt, FUN = function(curSjikt, speciesToPlot, speciesInfo, outMat, yearsVisited, curSiteName, inTag, plotLoc) {
        # Retrieve just those species in the current sjikt
        curSpeciesToPlot <- speciesToPlot[sapply(X = strsplit(as.character(speciesInfo[speciesToPlot, "SjiktPresence"]), "|", fixed = TRUE), FUN = function(curPosS, curSjikt) {
          curSjikt %in% curPosS
        }, curSjikt = curSjikt)]
        # Create a data frame of species to plot for the current sjikt
        plotFrame <- do.call(rbind, lapply(X = curSpeciesToPlot, FUN = function(curSpecies, outMat, yearsVisited, speciesInfo, inTag, curSiteName, curSjikt) {
          data.frame(
            speciesName = rep(as.character(speciesInfo[curSpecies, "SpeciesBinomial"], length(yearsVisited))),
            norwegianName = rep(
              ifelse(is.na(speciesInfo[curSpecies, "NameNorsk"]), as.character(speciesInfo[curSpecies, "SpeciesBinomial"]), as.character(speciesInfo[curSpecies, "NameNorsk"])),
              length(yearsVisited)),
            sjikt = rep(curSjikt, length(yearsVisited)),
            year = yearsVisited,
            values = as.double(outMat[curSpecies, paste("Mean", inTag, yearsVisited, curSiteName, sep = "")])
          )
        }, outMat = outMat, yearsVisited = yearsVisited, speciesInfo = speciesInfo, inTag = inTag, curSiteName = curSiteName, curSjikt = curSjikt))
        # Create the plot in ggplot
        outPlot <- ggplot(plotFrame, aes(x = year, y = values)) + geom_line(aes(colour = norwegianName)) + labs(x = "År", y = ifelse(inTag == "Freq", "% gjennomsnittlig smårutefrekvens", "% gjennomsnittsdekning")) +
          theme_classic() + theme(legend.title = element_blank())
        # Save it as an editable vector
        ggsave(paste(plotLoc, "_", curSiteName, "_Sjikt", curSjikt, "_", inTag, ".svg", sep = ""), outPlot, width = 7, height = 5)
        # Return the frame
        plotFrame
      }, speciesToPlot = speciesToPlot, speciesInfo = speciesInfo, outMat = outMat, yearsVisited = yearsVisited, curSiteName = curSiteName, inTag = inTag, plotLoc = plotLoc))
    }
    outMat
  }, communityMatrix = communityMatrix, siteInfo = siteInfo, speciesInfo = speciesInfo, inTag = inTag, plotLoc = plotLoc))
  outMat
}
allPopSummary <- cbind(
  yearlySummaryStats(communityMatrixFreq$all, siteInfo, speciesInfo, "Freq", paste(outputDirectory, "populationPlot", sep = "/")),
  yearlySummaryStats(communityMatrixCover$all, siteInfo, speciesInfo, "Cover", paste(outputDirectory, "populationPlot", sep = "/"))
)
speciesInfo <- cbind(speciesInfo, as.data.frame(matrix(NA, nrow = nrow(speciesInfo), ncol = ncol(allPopSummary), dimnames = list(NULL, colnames(allPopSummary)))))
speciesInfo[rownames(allPopSummary), colnames(allPopSummary)] <- allPopSummary
# Produce a population summary file with one sheet for each site
makeSummaryFile <- function(siteInfo, speciesInfo, inTag, yearlySpecies) {
  popSummaryBook <- createWorkbook()
  lapply(X = as.character(siteInfo$SiteName), FUN = function(curSiteName, popSummaryBook, speciesInfo, inTag, yearlySpecies) {
    # Produce a worksheet for the site
    addWorksheet(popSummaryBook, curSiteName)
    # Get the years used at the current site
    popColNames <- colnames(speciesInfo)[grepl(paste("^Mean", inTag, ".*", curSiteName, "$", sep = ""), colnames(speciesInfo), perl = TRUE)]
    yearsUsed <- sort(as.integer(gsub("\\D+$", "", gsub("^\\D+", "", popColNames, perl = TRUE), perl = TRUE)))
    popData <- speciesInfo[, popColNames]
    colnames(popData) <- yearsUsed
    writeData(popSummaryBook, curSiteName, startCol = 1, startRow = 3, data.frame(
      "Species" = speciesInfo$SpeciesBinomial,
      "Sjikt" = speciesInfo$SjiktPresence,
      "SpearmansRho" = speciesInfo[, paste("SpearmanRho", inTag, curSiteName, sep = "")],
      "SpearmansRhoPValue" = speciesInfo[, paste("SpearmanRhoPValue", inTag, curSiteName, sep = "")]
    ), colNames = TRUE, rowNames = FALSE)
    writeData(popSummaryBook, curSiteName, startCol = 5, startRow = 3, popData)
    # Restructure the Wilcoxon yearly analysis
    writeData(popSummaryBook, curSiteName, startCol = 5 + length(yearsUsed), startRow = 3, matrix(rep(c("V Statistic", "p Value"), length(yearsUsed) - 1), nrow = 1), colNames = FALSE, rowNames = FALSE)
    writeData(popSummaryBook, curSiteName, startCol = 5 + length(yearsUsed), startRow = 4, t(sapply(X = rownames(speciesInfo), FUN = function(curSpecies, yearlySpecies) {
      curYearlySpecies <- yearlySpecies[[curSpecies]]
      as.double(t(as.matrix(curYearlySpecies[, c("VStat", "PValue")])))
    }, yearlySpecies = yearlySpecies[[curSiteName]])), colNames = FALSE, rowNames = FALSE)
    # Setup the column headings
    sapply(X = 1:(length(yearsUsed) - 1), FUN = function(curIndex, popSummaryBook, curSiteName, yearsUsed) {
      writeData(popSummaryBook, curSiteName, startCol = 5 + length(yearsUsed) + (curIndex - 1) * 2, startRow = 2, matrix(paste(yearsUsed[curIndex], yearsUsed[curIndex + 1], sep = " - "), nrow = 1, ncol = 1), colNames = FALSE, rowNames = FALSE)
      mergeCells(popSummaryBook, curSiteName, cols = 5:6 + length(yearsUsed) + (curIndex - 1) * 2, rows = 2)
    }, popSummaryBook = popSummaryBook, curSiteName = curSiteName, yearsUsed = yearsUsed)
    writeData(popSummaryBook, curSiteName, startCol = 1, startRow = 1, matrix(c(
      "Species", "Sjikt", "Spearman's Rho", "Rho p Value", ifelse(inTag == "Freq", "Averge percentage of squares present", "Average percent cover"), rep("", length(yearsUsed) - 1), "Wilcoxon yearly change analysis"
    ), nrow = 1), colNames = FALSE, rowNames = FALSE)
    # Merge cells in the column headigns
    mergeCells(popSummaryBook, curSiteName, cols = 1, rows = 1:3)
    mergeCells(popSummaryBook, curSiteName, cols = 2, rows = 1:3)
    mergeCells(popSummaryBook, curSiteName, cols = 3, rows = 1:3)
    mergeCells(popSummaryBook, curSiteName, cols = 4, rows = 1:3)
    mergeCells(popSummaryBook, curSiteName, cols = 5:(4 + length(yearsUsed)), rows = 1:2)
    mergeCells(popSummaryBook, curSiteName, cols = 5 + length(yearsUsed) + 0:((length(yearsUsed) - 1) * 2 - 1), rows = 1)
    # Style the column headers
    numCols <- 5 + length(yearsUsed) + (length(yearsUsed) - 1) * 2 - 1
    topHeadStyle <- createStyle(textDecoration = "bold", halign = "center")
    lowerHeadStyle <- createStyle(border = "bottom", borderStyle = "medium", halign = "center")
    addStyle(popSummaryBook, curSiteName, style = topHeadStyle, rows = 1, cols = 1:numCols)
    addStyle(popSummaryBook, curSiteName, style = topHeadStyle, rows = 2, cols = 1:numCols)
    addStyle(popSummaryBook, curSiteName, style = lowerHeadStyle, rows = 3, cols = 1:numCols)
    # Create conditional formatting to highlight significant cells
    sigIncStyle <- createStyle(bgFill = rgb(193, 255, 193, maxColorValue = 255))
    sigDecStyle <- createStyle(bgFill = rgb(255, 182, 193, maxColorValue = 255))
    lapply(X = 1:(length(yearsUsed) - 1), FUN = function(curIndex, popSummaryBook, curSiteName, sigIncStyle, sigDecStyle, speciesInfo) {
      conditionalFormatting(popSummaryBook, curSiteName, cols = 10 + curIndex * 2, rows = 3 + 1:nrow(speciesInfo), style = sigIncStyle, rule = paste(
        "AND(", int2col(10 + curIndex * 2), "4<=0.05,", int2col(curIndex + 4), "4<", int2col(curIndex + 5), "4)", sep = ""
      ))
      conditionalFormatting(popSummaryBook, curSiteName, cols = 10 + curIndex * 2, rows = 3 + 1:nrow(speciesInfo), style = sigDecStyle, rule = paste(
        "AND(", int2col(10 + curIndex * 2), "4<=0.05,", int2col(curIndex + 4), "4>", int2col(curIndex + 5), "4)", sep = ""
      ))
    }, popSummaryBook = popSummaryBook, curSiteName = curSiteName, sigIncStyle = sigIncStyle, sigDecStyle = sigDecStyle, speciesInfo = speciesInfo)
    conditionalFormatting(popSummaryBook, curSiteName, cols = 4, rows = 3 + 1:nrow(speciesInfo), style = sigIncStyle, rule = paste(
      "AND(", int2col(4), "4<=0.05,", int2col(3), "4>0)", sep = ""
    ))
    conditionalFormatting(popSummaryBook, curSiteName, cols = 4, rows = 3 + 1:nrow(speciesInfo), style = sigDecStyle, rule = paste(
      "AND(", int2col(4), "4<=0.05,", int2col(3), "4<0)", sep = ""
    ))
  }, popSummaryBook = popSummaryBook, speciesInfo = speciesInfo, inTag = inTag, yearlySpecies = yearlySpecies)
  popSummaryBook
}
popSummaryBookFreq <- makeSummaryFile(siteInfo, speciesInfo, "Freq", yearlySpeciesFreq)
popSummaryBookCover <- makeSummaryFile(siteInfo, speciesInfo, "Cover", yearlySpeciesCover)
saveWorkbook(popSummaryBookFreq, paste(outputDirectory, "populationSummary_Freq.xlsx", sep = "/"), overwrite = TRUE)
saveWorkbook(popSummaryBookCover, paste(outputDirectory, "populationSummary_Cover.xlsx", sep = "/"), overwrite = TRUE)

# ------ 6. PERFORM ORDINATION ANALYSIS ------

# Run DCA on the community matrices
DCAFreq <- decorana(communityMatrixFreq$all)
DCACover <- decorana(communityMatrixCover$all)
# Run NMDS on each of the different community matrices
NMDSFreq <- metaMDS(communityMatrixFreq$all, try = 50, trymax = 10000)
NMDSCover <- metaMDS(communityMatrixCover$all, try = 50, trymax = 10000)
# Rearrange the ordination axes (by plot)
tempPlotOrd <- cbind(scores(DCAFreq, display = "sites"), scores(DCACover, display = "sites"), scores(NMDSFreq, display = "sites"), scores(NMDSCover, display = "sites"))
colnames(tempPlotOrd) <- c(
  paste("Freq", colnames(scores(DCAFreq, display = "sites")), sep = ""),
  paste("Cover", colnames(scores(DCACover, display = "sites")), sep = ""),
  paste("Freq", colnames(scores(NMDSFreq, display = "sites")), sep = ""),
  paste("Cover", colnames(scores(NMDSCover, display = "sites")), sep = "")
)
# Merge them into the plot information frame
plotInfo <- cbind(plotInfo, tempPlotOrd)
# Rearrange the ordination axes (by species)
tempSpeciesOrd <- cbind(scores(DCAFreq, display = "species"), scores(DCACover, display = "species"), scores(NMDSFreq, display = "species"), scores(NMDSCover, display = "species"))
colnames(tempSpeciesOrd) <- c(
  paste("Freq", colnames(scores(DCAFreq, display = "species")), sep = ""),
  paste("Cover", colnames(scores(DCACover, display = "species")), sep = ""),
  paste("Freq", colnames(scores(NMDSFreq, display = "species")), sep = ""),
  paste("Cover", colnames(scores(NMDSCover, display = "species")), sep = "")
)
# Merge them into the species information frame
speciesInfo <- cbind(speciesInfo, tempSpeciesOrd)
# Function to create trajectory plots when using ordination
createTrajectoryPlot <- function(sjiktInfo, speciesInfo, plotInfo, siteInfo, inTag, ordType, plotLoc) {
  ordPlots <- setNames(lapply(X = as.character(siteInfo$SiteName), FUN = function(curSiteName, speciesInfo, plotInfo, siteInfo, inTag, ordType, plotLoc) {
    # Use only the data taken at the current site
    curPlotInfo <- plotInfo[as.character(plotInfo$SiteName) == curSiteName, ]
    # Retrieve the years that sampling took place at the current location
    yearsUsed <- sort(unique(curPlotInfo$Year))
    # Retrieve the plot codes at the site
    plotCodes <- sort(unique(as.character(curPlotInfo$PlotNumber)))
    # Restructure the data into a transition data frame
    transFrame <- do.call(rbind, lapply(X = 1:(length(yearsUsed) - 1), FUN = function(curIndex, yearsUsed, plotCodes, curPlotInfo, inTag, ordType) {
      do.call(rbind, lapply(X = plotCodes, FUN = function(curPlotCode, fromYear, toYear, curPlotInfo, inTag, ordType) {
        fromOrdOne <- NA
        fromOrdTwo <- NA
        toOrdOne <- NA
        toOrdTwo <- NA
        isFrom <- curPlotInfo$Year == fromYear & curPlotInfo$PlotNumber == curPlotCode
        isTo <- curPlotInfo$Year == toYear & curPlotInfo$PlotNumber == curPlotCode
        if(any(isFrom)) {
          fromOrdOne <- curPlotInfo[isFrom, paste(inTag, ordType, "1", sep = "")]
          fromOrdTwo <- curPlotInfo[isFrom, paste(inTag, ordType, "2", sep = "")]
        }
        if(any(isTo)) {
          toOrdOne <- curPlotInfo[isTo, paste(inTag, ordType, "1", sep = "")]
          toOrdTwo <- curPlotInfo[isTo, paste(inTag, ordType, "2", sep = "")]
        }
        data.frame(
          SiteName = as.character(curPlotInfo$SiteName[1]),
          PlotNumber = curPlotCode,
          YearTransition = paste(fromYear, toYear, sep = " - "),
          FromOrdOne = fromOrdOne,
          FromOrdTwo = fromOrdTwo,
          ToOrdOne = toOrdOne,
          ToOrdTwo = toOrdTwo
        )
      }, fromYear = yearsUsed[curIndex], toYear = yearsUsed[curIndex + 1], curPlotInfo = curPlotInfo, inTag = inTag, ordType = ordType))
    }, yearsUsed = yearsUsed, plotCodes = plotCodes, curPlotInfo = curPlotInfo, inTag = inTag, ordType = ordType))
    latestPoints <- curPlotInfo[curPlotInfo$Year == yearsUsed[length(yearsUsed)], paste(inTag, ordType, 1:2, sep = "")]
    # Create a trajectory plot for the plot information
    ordFigPlot <- ggplot(transFrame) + geom_segment(aes(x = FromOrdOne, y = FromOrdTwo, xend = ToOrdOne, yend = ToOrdTwo, colour = YearTransition)) +
      geom_point(aes_string(x = paste(inTag, ordType, 1, sep = ""), y = paste(inTag, ordType, 2, sep = "")), data = latestPoints) +
      theme_classic() + xlab(paste(ordType, "1", sep = " ")) + ylab(paste(ordType, "2", sep = " ")) +
      scale_colour_manual(values = colorRampPalette(rgb(c(191, 0), c(239, 178), c(255, 238), maxColorValue = 255))(length(yearsUsed) - 1), name = NULL, labels = paste(yearsUsed[1:(length(yearsUsed) - 1)], yearsUsed[2:length(yearsUsed)], sep = "-"))
    ggsave(paste(plotLoc, "_", curSiteName, ".svg", sep = ""), ordFigPlot, width = 9, height = 7)
    ordFigPlot
  }, speciesInfo = speciesInfo, plotInfo = plotInfo, siteInfo = siteInfo, inTag = inTag, ordType = ordType, plotLoc = plotLoc), as.character(siteInfo$SiteName))
  # Create a species ordination plot
  ordSpecies <- setNames(lapply(X = as.character(sjiktInfo$SjiktCode), FUN = function(curSjikt, speciesInfo) {
    # Retrieve the species within the current Sjikt
    curSpeciesInfo <- speciesInfo[sapply(X = strsplit(as.character(speciesInfo$SjiktPresence), "|", fixed = TRUE), FUN = function(sjiktPos, curSjikt) {
      curSjikt %in% sjiktPos
    }, curSjikt = curSjikt), ]
    curSpeciesInfo <- cbind(curSpeciesInfo, data.frame(
      originX = rep(0.0, nrow(curSpeciesInfo)),
      originY = rep(0.0, nrow(curSpeciesInfo)),
      compositeName = ifelse(is.na(curSpeciesInfo$NameNorsk), as.character(curSpeciesInfo$SpeciesBinomial), as.character(curSpeciesInfo$NameNorsk))
    ))
    ordFigSpecies <- ggplot(curSpeciesInfo, aes_string(x = "originX", y = "originY", xend = paste(inTag, ordType, 1, sep = ""), yend = paste(inTag, ordType, 2, sep = ""), colour = "compositeName")) +
      geom_segment(arrow = arrow(length = unit(0.01, "npc"), type = "closed")) + theme_classic() + xlab(paste(ordType, "1", sep = " ")) + ylab(paste(ordType, "2", sep = " ")) + theme(legend.title = element_blank())
    ggsave(paste(plotLoc, "_Sjikt", curSjikt, ".svg", sep = ""), ordFigSpecies, width = 9, height = 7)
    ordFigSpecies
  }, speciesInfo = speciesInfo), as.character(sjiktInfo$SjiktCode))
  list(plots = ordPlots, species = ordSpecies)
}
createTrajectoryPlot(sjiktInfo, speciesInfo, plotInfo, siteInfo, "Freq", "NMDS", paste(outputDirectory, "ordinationNMDSFreq", sep = "/"))
createTrajectoryPlot(sjiktInfo, speciesInfo, plotInfo, siteInfo, "Cover", "NMDS", paste(outputDirectory, "ordinationNMDSCover", sep = "/"))
createTrajectoryPlot(sjiktInfo, speciesInfo, plotInfo, siteInfo, "Freq", "DCA", paste(outputDirectory, "ordinationDCAFreq", sep = "/"))
createTrajectoryPlot(sjiktInfo, speciesInfo, plotInfo, siteInfo, "Cover", "DCA", paste(outputDirectory, "ordinationDCACover", sep = "/"))
# Create statistics summary for the ordination plots
ordBook <- createWorkbook()
lapply(X = rownames(siteInfo), FUN = function(curSite, plotInfo, ordCols, ordBook) {
  # Add a worksheet to the workbook
  addWorksheet(ordBook, curSite)
  # Do change statistics for each
  statMat <- t(sapply(X = ordCols, FUN = function(curCol, curSite, plotInfo, ordBook) {
    # Retrieve the years that the ordination plots were measured
    allYears <- as.character(sort(unique(plotInfo$Year)))
    # Do a spearman rank correlation
    spearOut <- cor.test(as.numeric(plotInfo$Year), plotInfo[, curCol])
    # Do a series of Wilcoxon signed rank tests
    wilcoxOut <- apply(X = cbind(allYears[1:(length(allYears) - 1)], allYears[2:length(allYears)]), FUN = function(curYears, plotInfo, curCol) {
      # Remove those plots that do not have entries in both years
      plotNames <-sort(unique(as.character(plotInfo$PlotNumber)))
      plotsHaveBoth <- plotNames[sapply(X = plotNames, FUN = function(curPlotName, plotInfo, curYears) {
        curPlotInfo <- plotInfo[as.character(plotInfo$PlotNumber) == curPlotName, ]
        any(as.character(curPlotInfo$Year) == curYears[1]) && any(as.character(curPlotInfo$Year) == curYears[2])
      }, plotInfo = plotInfo, curYears = curYears)]
      curInfo <- plotInfo[as.character(plotInfo$PlotNumber) %in% plotsHaveBoth, ]
      tempFrame <- data.frame(
        fromVal = curInfo[as.character(curInfo$Year) == curYears[1], curCol],
        toVal = curInfo[as.character(curInfo$Year) == curYears[2], curCol]
      )
      # Perform the Wilcoxon's signed ranks test
      wilcox.test(tempFrame$fromVal, tempFrame$toVal, paired = TRUE)
    }, plotInfo = plotInfo, curCol = curCol, MARGIN = 1)
    setNames(c(spearOut$estimate, spearOut$p.value, unlist(lapply(X = wilcoxOut, FUN = function(curWil) {
      c(curWil$statistic, curWil$p.value)
    }))), c(
      "Spearman_Rho",
      "Rho_pValue",
      paste(
        rep(allYears[1:(length(allYears) - 1)], rep(2, length(allYears) - 1)),
        "-",
        rep(allYears[2:length(allYears)], rep(2, length(allYears) - 1)),
        rep(c("_VStat", "_pValue"), length(wilcoxOut)),
      sep = "")
    ))
  }, curSite = curSite, plotInfo = plotInfo[as.character(plotInfo$SiteName) == curSite, ], ordBook = ordBook))
  statFrame <- cbind(data.frame(OrdStat = ordCols), as.data.frame(statMat))
  writeData(ordBook, curSite, statFrame, colNames = FALSE, rowNames = FALSE, startCol = 1, startRow = 4)
  writeData(ordBook, curSite, as.data.frame(matrix(c(
    "Ordination Statistic", "Spearman's Rho", "Rho p Value", "Wilcoxon yearly change analysis"
  ), nrow = 1)), colNames = FALSE, rowNames = FALSE, startCol = 1, startRow = 1)
  yearLab <- gsub("-", " - ", gsub("_.*$", "", colnames(statMat)[3:ncol(statMat)], perl = TRUE), fixed = TRUE)
  writeData(ordBook, curSite, as.data.frame(matrix(c(
    ifelse(duplicated(yearLab), "", yearLab)
  ), nrow = 1)), colNames = FALSE, rowNames = FALSE, startCol = 4, startRow = 2)
  writeData(ordBook, curSite, as.data.frame(matrix(c(
    rep(c("V Stat", "p Value"), (ncol(statMat) - 2) / 2)
  ), nrow = 1)), colNames = FALSE, rowNames = FALSE, startCol = 4, startRow = 3)
  # Merge cells in the column headings
  mergeCells(ordBook, curSite, cols = 1, rows = 1:3)
  mergeCells(ordBook, curSite, cols = 2, rows = 1:3)
  mergeCells(ordBook, curSite, cols = 3, rows = 1:3)
  mergeCells(ordBook, curSite, cols = 4:ncol(statFrame), rows = 1)
  lapply(X = seq(4, ncol(statFrame) - 1, by = 2), FUN = function(curCol, ordBook, curSite) {
    mergeCells(ordBook, curSite, cols = curCol:(curCol + 1), rows = 2)
  }, ordBook = ordBook, curSite = curSite)
  # Style the column headers
  topHeadStyle <- createStyle(textDecoration = "bold", halign = "center")
  lowerHeadStyle <- createStyle(border = "bottom", borderStyle = "medium", halign = "center")
  addStyle(ordBook, curSite, style = topHeadStyle, rows = 1, cols = 1:ncol(statFrame))
  addStyle(ordBook, curSite, style = topHeadStyle, rows = 2, cols = 1:ncol(statFrame))
  addStyle(ordBook, curSite, style = lowerHeadStyle, rows = 3, cols = 1:ncol(statFrame))
  # Create conditional formatting to highlight significant cells
  sigIncStyle <- createStyle(bgFill = rgb(193, 255, 193, maxColorValue = 255))
  lapply(X = c(3, seq(5, ncol(statFrame), by = 2)), FUN = function(curIndex, ordBook, curSite, sigIncStyle, statFrame) {
    conditionalFormatting(ordBook, curSite, cols = curIndex, rows = 3 + 1:nrow(statFrame), style = sigIncStyle, rule = paste(
      int2col(curIndex), "4<=0.05", sep = ""
    ))
  }, ordBook = ordBook, curSite = curSite, sigIncStyle = sigIncStyle, statFrame = statFrame)
}, plotInfo = plotInfo, ordCols = c(
  paste("FreqDCA", 1:4, sep = ""),
  paste("CoverDCA", 1:4, sep = ""),
  paste("FreqNMDS", 1:2, sep = ""),
  paste("CoverNMDS", 1:2, sep = "")
), ordBook = ordBook)
saveWorkbook(ordBook, paste(outputDirectory, "ordinationSummary.xlsx", sep = "/"), overwrite = TRUE)

# ------ 7. PRODUCE ELLENBERG ANALYSIS ------
# Set the maximum and minimum values of the Ellenberg indicator values
attr(ellenbergDescriptionNorsk, "range") <- matrix(c(
  1, 1, 1, 1,
  9, 12, 9, 9
), nrow = 2, byrow = TRUE, dimnames = list(c("min", "max"), ellenbergColNames))
# Set the colour palette to use in the Ellenberg plots
attr(ellenbergDescriptionNorsk, "dispColours") <- rgb(c(255, 176, 255, 180), c(250, 226, 192, 238), c(205, 255, 203, 180), maxColorValue = 255)
# Function to fit a linear regression model to the Ellenberg data
weightedEllenberg <- function(curEllenbergVal, communityMatrix) {
  # Make a data frame to conduct the Ellenberg model analysis on
  weightedValues <- apply(X = communityMatrix, FUN = function(curCol, ellenbergValues) {
    sum(curCol / sum(curCol, na.rm = TRUE) * ellenbergValues, na.rm = TRUE)
  }, MARGIN = 1, ellenbergValues = speciesInfo[, curEllenbergVal])
  weightedValues
}
weightedEllenbergValues <- cbind(
  as.data.frame(sapply(X = names(ellenbergDescriptionNorsk), FUN = weightedEllenberg, communityMatrix = communityMatrixFreq$all)),
  as.data.frame(sapply(X = names(ellenbergDescriptionNorsk), FUN = weightedEllenberg, communityMatrix = communityMatrixCover$all))
)
colnames(weightedEllenbergValues) <- c(
  paste("FreqEllenberg", names(ellenbergDescriptionNorsk), sep = ""),
  paste("CoverEllenberg", names(ellenbergDescriptionNorsk), sep = "")
)
plotInfo <- cbind(plotInfo, weightedEllenbergValues)
# Function to perform the Ellenberg modelling
runEllenbergModel <- function(plotInfo, ellenbergDescriptionNorsk, inTag, plotLoc) {
  # Run the Ellenberg models
  ellenbergModels <- setNames(lapply(X = names(ellenbergDescriptionNorsk), FUN = function(curEllenbergVal, plotInfo, ellenbergDescriptionNorsk, inTag) {
    valRange <- attr(ellenbergDescriptionNorsk, "range")[, curEllenbergVal]
    # Extend the plot informa
    exPlotInfo <- cbind(plotInfo, data.frame(
      scaledValues = (plotInfo[, paste(inTag, "Ellenberg", curEllenbergVal, sep = "")] - valRange[1]) / (valRange[2] - valRange[1])
    ))
    if(length(unique(exPlotInfo$SiteName)) <= 1) {
      modelOut <- inla(scaledValues ~ Year + f(PlotNumber, model = "iid"), data = exPlotInfo, family = "beta", control.predictor = list(compute = TRUE))
    } else {
      # Create a series of linear combinations to monitor the overall effect of Year at each site
      monitorLinCombs <- do.call(c, lapply(X = unique(as.character(exPlotInfo$SiteName)), FUN = function(curSite, modelNames) {
        interactionCoeffName <- paste("Site", curSite, ":Year", sep = "")
        monitorVec <- setNames(1, "Year")
        if(interactionCoeffName %in% modelNames) {
          monitorVec <- c(monitorVec, setNames(1, interactionCoeffName))
        }
        outVal <- do.call(inla.make.lincomb, as.list(monitorVec))
        names(outVal) <- curSite
        outVal
      }, modelNames = colnames(model.matrix(Value ~ SiteName * Year, data = exPlotInfo))))
      modelOut <- inla(scaledValues ~ SiteName * Year + f(PlotNumber, model = "iid"), data = exPlotInfo, family = "beta", control.predictor = list(compute = TRUE), lincomb = monitorLinCombs)
    }
    modelOut
  }, plotInfo = plotInfo, ellenbergDescriptionNorsk = ellenbergDescriptionNorsk, inTag = inTag), names(ellenbergDescriptionNorsk))
  # Reorder the Ellenberg model outputs
  ellenbergFrame <- do.call(rbind, lapply(X = names(ellenbergDescriptionNorsk), FUN = function(curEllenbergVal, plotInfo, ellenbergModels, ellenbergDescriptionNorsk, inTag) {
    valRange <- attr(ellenbergDescriptionNorsk, "range")[, curEllenbergVal]
    data.frame(
      PlotCode = plotInfo$PlotCode,
      SiteName = plotInfo$SiteName,
      Year = plotInfo$Year,
      Value = plotInfo[, paste(inTag, "Ellenberg", curEllenbergVal, sep = "")],
      Prediction = ellenbergModels[[curEllenbergVal]]$summary.fitted.values$mean * (valRange[2] - valRange[1]) + valRange[1],
      EllenbergType = rep(curEllenbergVal, nrow(plotInfo))
    )
  }, plotInfo = plotInfo, ellenbergModels = ellenbergModels, ellenbergDescriptionNorsk = ellenbergDescriptionNorsk, inTag = inTag))
  # Plot a violin plot of the Ellenberg values and the model prediction for each site
  ellenbergPlots <- setNames(lapply(X = sort(unique(as.character(plotInfo$SiteName))), FUN = function(curSiteName, ellenbergModels, ellenbergFrame, ellenbergDescriptionNorsk, inTag, plotLoc) {
    # Make the standard violin plots
    exPlotInfo <- ellenbergFrame[as.character(ellenbergFrame$SiteName) == curSiteName, ]
    exPlotInfo$Year <- as.character(exPlotInfo$Year)
    lowCred <- function(yin) {quantile(yin, probs = 0.025, names = FALSE)}
    uppCred <- function(yin) {quantile(yin, probs = 0.975, names = FALSE)}
    violinPlot <- ggplot(exPlotInfo, aes(x = Year, y = Value, fill = EllenbergType)) +
      geom_violin() + stat_summary(aes(y = Prediction), fun.max = uppCred, fun.min = lowCred, geom = "ribbon", alpha = 0.2, fill = "grey", group = 1) +
      stat_summary(aes(y = Prediction), fun = mean, geom = "line", linetype = "longdash", group = 1) + stat_summary(fun = mean, geom = "point", size = 4, colour = "black") +
      facet_grid(EllenbergType ~ ., scales = "free_y", labeller = labeller(EllenbergType = ellenbergDescriptionNorsk)) + theme_classic() + theme(legend.position = "none", strip.background = element_blank(), strip.text.y = element_text(angle = 90)) +
      xlab("År") + ylab("Ellenberg Indikator Verdi") + scale_fill_manual(values = attr(ellenbergDescriptionNorsk, "dispColours"))
    ggsave(paste(plotLoc, "Violin_", curSiteName, "_", inTag, ".svg", sep = ""), violinPlot, width = 5, height = 10)
    # Produce the standardised regeression density plots
    marginalYearFrame <- do.call(rbind, lapply(X = names(ellenbergDescriptionNorsk), FUN = function(curEllenbergVal, ellenbergModels, curSiteName) {
      frameOut <- NULL
      credInt <- c()
      if(is.null(ellenbergModels[[curEllenbergVal]]$marginals.lincomb.derived)) {
        # Model is simple (only one site being processed)
        marginalFrame <- ellenbergModels[[curEllenbergVal]]$marginals.fixed[["Year"]]
        # Retrieve the effect credible interval
        credInt <- as.double(ellenbergModels[[curEllenbergVal]]$summary.fixed["Year", c("0.025quant", "0.975quant")])
      } else {
        # Otherise the model is complicated (multiple sites being processed)
        # Retrieve the marginal density estimate for the derived yearly effect parameter
        marginalFrame <- ellenbergModels[[curEllenbergVal]]$marginals.lincomb.derived[[curSiteName]]
        # Retrieve the effect credible interval
        credInt <- as.double(ellenbergModels[[curEllenbergVal]]$summary.lincomb.derived[curSiteName, c("0.025quant", "0.975quant")])
      }
      frameOut <- data.frame(
        # Rescale the Ellenberg values
        ellenbergValue = marginalFrame[, "x"],
        # Add the other Ellenberg information
        density = marginalFrame[, "y"],
        ellenbergType = rep(curEllenbergVal, nrow(marginalFrame)),
        site = rep(curSiteName, nrow(marginalFrame))
      )
      # Add a column to help for the display of the Ellenberg credible interval
      frameOut <- cbind(frameOut, data.frame(
        ellenbergCredInt = ifelse(frameOut$ellenbergValue >= credInt[1] & frameOut$ellenbergValue <= credInt[2], frameOut$ellenbergValue, NA)
      ))
      frameOut
    }, ellenbergModels = ellenbergModels, curSiteName = curSiteName))
    # Create a plot to generate the credible interval
    intervalPlot <- ggplot(marginalYearFrame, aes(ellenbergValue, density)) +
      geom_ribbon(aes(x = ellenbergCredInt, ymax = density, fill = ellenbergType), ymin = 0) + geom_line(group = 1) +
      geom_vline(xintercept = 0.0, linetype = "longdash") +
      facet_grid(ellenbergType ~ ., scales = "free_y", labeller = labeller(ellenbergType = ellenbergDescriptionNorsk)) + theme_classic() + theme(legend.position = "none", strip.background = element_blank(), strip.text.y = element_text(angle = 90)) +
      xlab("Skalert årlig regresjonskoeffisient") + ylab("Tetthet") + scale_fill_manual(values = attr(ellenbergDescriptionNorsk, "dispColours")) +
      scale_x_continuous(limits = range(marginalYearFrame$ellenbergCredInt, na.rm = TRUE) + 0.1 * c(-1.0, 1.0) * diff(range(marginalYearFrame$ellenbergCredInt, na.rm = TRUE)))
    ggsave(paste(plotLoc, "Density_", curSiteName, "_", inTag, ".svg", sep = ""), intervalPlot, width = 4, height = 10)
    list(violin = violinPlot, density = intervalPlot)
  }, ellenbergModels = ellenbergModels, ellenbergFrame = ellenbergFrame, ellenbergDescriptionNorsk = ellenbergDescriptionNorsk, inTag = inTag, plotLoc = plotLoc), sort(unique(as.character(plotInfo$SiteName))))
  list(models = ellenbergModels, plots = ellenbergPlots)
}
ellenbergFreq <- runEllenbergModel(plotInfo = plotInfo, ellenbergDescriptionNorsk = ellenbergDescriptionNorsk, inTag = "Freq", plotLoc = paste(outputDirectory, "ellenberg", sep = "/"))
ellenbergCover <- runEllenbergModel(plotInfo = plotInfo, ellenbergDescriptionNorsk = ellenbergDescriptionNorsk, inTag = "Cover", plotLoc = paste(outputDirectory, "ellenberg", sep = "/"))

# ------ 8. PERFORM RICHNESS SUMMARIES ------
richnessBook <- createWorkbook()
lapply(X = as.character(siteInfo$SiteName), FUN = function(curSiteName, communityMatrixFreq, communityMatrixCover, plotInfo, sjiktInfo, richnessBook, sjiktGroups, plotLoc) {
  # Create a data frame of species richness in each sjikt over each of the years
  richnessFrame <- do.call(rbind, lapply(X = names(communityMatrixFreq), FUN = function(curSjikt, curSiteName, communityMatrixFreq, plotInfo, sjiktInfo) {
    curCommunityMatrix <- communityMatrixFreq[[curSjikt]][plotInfo[row.names(communityMatrixFreq[[curSjikt]]), "SiteName"] == curSiteName, ]
    # Make a series of community matrices of the different years
    yearlyCounts <- matrix(sapply(X = sort(unique(plotInfo[row.names(curCommunityMatrix), "Year"])), FUN = function(curYear, curCommunityMatrix) {
      specCountSum <- apply(X = curCommunityMatrix[grepl(paste("\\-", curYear, "$", sep = ""), row.names(curCommunityMatrix), perl = TRUE), ], FUN = sum, MARGIN = 2)
      sum(specCountSum > 0)
    }, curCommunityMatrix = curCommunityMatrix), nrow = 1, dimnames = list(NULL, sort(unique(plotInfo[row.names(curCommunityMatrix), "Year"]))))
    cbind(data.frame(
      Sjikt = ifelse(curSjikt == "all", "Alle sjiktene", as.character(sjiktInfo[curSjikt, "DescriptionNorsk"]))
    ), as.data.frame(yearlyCounts))
  }, curSiteName = curSiteName, communityMatrixFreq = communityMatrixFreq, plotInfo = plotInfo, sjiktInfo = sjiktInfo))
  # Function to create a summary (of cover of frequency) of the different sjikt over the years
  makeSjiktSummary <- function(curSjikt, curSiteName, communityMatrix, plotInfo, sjiktInfo) {
    curCommunityMatrix <- communityMatrix[[curSjikt]][plotInfo[row.names(communityMatrix[[curSjikt]]), "SiteName"] == curSiteName, ]
    # Make a series of community matrices of the different years
    yearlySummary <- matrix(sapply(X = sort(unique(plotInfo[row.names(curCommunityMatrix), "Year"])), FUN = function(curYear, curCommunityMatrix, plotInfo) {
      yearComMat <- curCommunityMatrix[plotInfo[row.names(curCommunityMatrix), "Year"] == curYear, ]
      plotSums <- apply(X = yearComMat, FUN = sum, MARGIN = 1, na.rm = TRUE)
      mean(plotSums, na.rm = TRUE)
    }, curCommunityMatrix = curCommunityMatrix, plotInfo = plotInfo), nrow = 1, dimnames = list(NULL, sort(unique(plotInfo[row.names(curCommunityMatrix), "Year"]))))
    cbind(data.frame(
      Sjikt = ifelse(curSjikt == "all", "Alle sjiktene", as.character(sjiktInfo[curSjikt, "DescriptionNorsk"]))
    ), as.data.frame(yearlySummary))
  }
  sjiktSummaryFrameFreq <- do.call(rbind, lapply(X = names(communityMatrixFreq), FUN = makeSjiktSummary, curSiteName = curSiteName, communityMatrix = communityMatrixFreq, plotInfo = plotInfo, sjiktInfo = sjiktInfo))
  sjiktSummaryFrameCover <- do.call(rbind, lapply(X = names(communityMatrixCover), FUN = makeSjiktSummary, curSiteName = curSiteName, communityMatrix = communityMatrixCover, plotInfo = plotInfo, sjiktInfo = sjiktInfo))
  # Make plots of the sjikt summaries
  makeSjiktPlot <- function(inFrame, sjiktGroups, plotLoc, sjiktInfo, inYLab) {
    outPlots <- setNames(lapply(X = sjiktGroups, FUN = function(curGroup, inFrame, plotLoc, sjiktInfo) {
      curFrame <- as.data.frame(melt(inFrame[as.character(inFrame$Sjikt) %in% sjiktInfo[curGroup, "DescriptionNorsk"], ], "Sjikt"))
      curFrame$variable <- as.numeric(as.character(curFrame$variable))
      outPlot <- ggplot(curFrame, aes(x = variable, y = value, colour = Sjikt)) + geom_line() + theme_classic() + xlab("År") + ylab(inYLab)
      ggsave(paste(plotLoc, paste(sort(curGroup), collapse = ""), ".svg", sep = ""), outPlot, width = 7, height = 5)
      outPlot
    }, inFrame = inFrame, plotLoc = plotLoc, sjiktInfo = sjiktInfo), sapply(X = sjiktGroups, FUN = paste, collapse = "|"))
    outPlots
  }
  richnessPlots <- makeSjiktPlot(richnessFrame, sjiktGroups, paste(plotLoc, "_", curSiteName, "_", "Richness", sep = ""), sjiktInfo, "Antall arter")
  freqPlots <- makeSjiktPlot(sjiktSummaryFrameFreq, sjiktGroups, paste(plotLoc, "_", curSiteName, "_", "Freq", sep = ""), sjiktInfo, "Gjennomsnitt frekvens (%)")
  coverPlots <- makeSjiktPlot(sjiktSummaryFrameCover, sjiktGroups, paste(plotLoc, "_", curSiteName, "_", "Cover", sep = ""), sjiktInfo, "Gjennomsnittsdekning (%)")
  # Add sheets to the summary workbook
  addWorksheet(richnessBook, paste(curSiteName, "Richness", sep = "_"))
  writeDataTable(richnessBook, paste(curSiteName, "Richness", sep = "_"), richnessFrame)
  addWorksheet(richnessBook, paste(curSiteName, "Freq", sep = "_"))
  writeDataTable(richnessBook, paste(curSiteName, "Freq", sep = "_"), sjiktSummaryFrameFreq)
  addWorksheet(richnessBook, paste(curSiteName, "Cover", sep = "_"))
  writeDataTable(richnessBook, paste(curSiteName, "Cover", sep = "_"), sjiktSummaryFrameCover)
}, communityMatrixFreq = communityMatrixFreq, communityMatrixCover = communityMatrixCover, plotInfo = plotInfo, sjiktInfo = sjiktInfo, richnessBook = richnessBook, sjiktGroups = sjiktGroups, plotLoc = paste(outputDirectory, "sjiktPlot", sep = "/"))
saveWorkbook(richnessBook, paste(outputDirectory, "sjiktSummary.xlsx", sep = "/"), overwrite = TRUE)

# ------ 9. PERFORM ANALYSIS OF DAMAGE ------
allExtraInfo <- lapply(X = row.names(siteInfo), FUN = function(curSite, yearsToProcess, extraInfoLoc, extraAliases, plotLoc) {
  # Retrieve the set of sheet names that are relevant to the site
  relevantSheetNames <- getSheetNames(extraInfoLoc)[sapply(X = getSheetNames(extraInfoLoc), FUN = function(curName, yearsToProcess, curSite) {
    curEntry <- strsplit(curName, " ", fixed = TRUE)[[1]]
    curEntry[1] == curSite && curEntry[2] %in% yearsToProcess
  }, yearsToProcess = yearsToProcess, curSite = curSite)]
  # Iterate over each of the relevant sheets
  allFrames <- lapply(X = relevantSheetNames, FUN = function(curSheetName, extraInfoLoc, extraAliases, plotLoc) {
    curSheetData <- readWorkbook(extraInfoLoc, curSheetName)
    # Rename any aliases for the extra information
    curSheetData$GjeldendeFloranavnNINA <- sapply(X = as.character(curSheetData$GjeldendeFloranavnNINA), FUN = function(curVal, extraAliases) {
      outVal <- curVal
      isAlias <- sapply(X = extraAliases, FUN = function(curAlias, curVal) {curVal %in% curAlias}, curVal = curVal)
      if(any(isAlias)) {
        outVal <- names(extraAliases)[isAlias]
      }
      outVal
    }, extraAliases = extraAliases)
    # Rename inconsistently names columns
    colnames(curSheetData)[colnames(curSheetData) == "Mengde"] <- "Dekn"
    # Add a plot code column if one does not exist
    if(!("Analyse_navn" %in% colnames(curSheetData))) {
      curSheetData <- cbind(curSheetData, data.frame(
        Analyse_navn <- paste(as.character(curSheetData[, "RUTE_ID"]), as.character(curSheetData[, "AAR"]), sep = "-")
      ))
    }
    curSheetData$Analyse_navn <- ifelse(is.na(curSheetData$Analyse_navn),
      paste(as.character(curSheetData[, "RUTE_ID"]), as.character(curSheetData[, "AAR"]), sep = "-"),
      as.character(curSheetData$Analyse_navn))
    curEntry <- enc2native(strsplit(curSheetName, " ", fixed = TRUE)[[1]])
    extraColNames <- sort(unique(as.character(curSheetData$GjeldendeFloranavnNINA)))
    allPlotCodes <- sort(unique(as.character(curSheetData$Analyse_navn)))
    createVarMatrix <- function(curValue, curSheetData, typeCol, plotCodes) {
      outVec <- setNames(curSheetData[as.character(curSheetData$GjeldendeFloranavnNINA) == curValue, typeCol], curSheetData$Analyse_navn[as.character(curSheetData$GjeldendeFloranavnNINA) == curValue])
      outVec[plotCodes]
    }
    # Restructure the entries as columns
    varMatFreq <- sapply(X = extraColNames, FUN = createVarMatrix, curSheetData = curSheetData, typeCol = "Frekvens", plotCodes = allPlotCodes)
    colnames(varMatFreq) <- paste(gsub("\\s", "_", extraColNames, perl = TRUE), "Freq", sep = "_")
    varMatCover <- sapply(X = extraColNames, FUN = createVarMatrix, curSheetData = curSheetData, typeCol = "Dekn", plotCodes = allPlotCodes)
    colnames(varMatCover) <- paste(gsub("\\s", "_", extraColNames, perl = TRUE), "Cover", sep = "_")
    # Calcualte the number of zero entries for each plot type
    nullCountFreq <- data.frame(
      colName = extraColNames,
      count = sapply(X = extraColNames, FUN = function(curColName, inFrame) {
        nrow(inFrame[as.character(inFrame$GjeldendeFloranavnNINA) == curColName, ])
      }, inFrame = curSheetData[curSheetData$Frekvens <= 0.0, ])
    )
    nullCountCover <- data.frame(
      colName = extraColNames,
      count = sapply(X = extraColNames, FUN = function(curColName, inFrame) {
        nrow(inFrame[as.character(inFrame$GjeldendeFloranavnNINA) == curColName, ])
      }, inFrame = curSheetData[curSheetData$Dekn <= 0.0, ])
    )
    allVals <- unique(as.character(curSheetData$GjeldendeFloranavnNINA))
    firstLevels <- c("Podosphaera/Pucciniastrum", "Valdensinia heterodoxa")
    curSheetData$GjeldendeFloranavnNINA <- factor(as.character(curSheetData$GjeldendeFloranavnNINA), levels = c(allVals[!(allVals %in% firstLevels)], firstLevels))
    nullCountFreq$colName <- factor(as.character(nullCountFreq$colName), levels = levels(curSheetData$GjeldendeFloranavnNINA))
    nullCountCover$colName <- factor(as.character(nullCountCover$colName), levels = levels(curSheetData$GjeldendeFloranavnNINA))
    nullCountFreq <- nullCountFreq[order(as.integer(nullCountFreq$colName)), ]
    nullCountCover <- nullCountCover[order(as.integer(nullCountCover$colName)), ]
    # Create a set of figures for the current information
    freqPlot <- ggplot(curSheetData[curSheetData$Frekvens > 0, ]) + geom_dotplot(aes(x = GjeldendeFloranavnNINA, y = Frekvens, fill = GjeldendeFloranavnNINA), binaxis = "y", stackdir = "center", stackratio = 1.5, dotsize = 0.4) +
      ylab("Frekvens") + xlab("") + geom_text(aes(x = colName, y = 0.0, label = count, vjust = "top"), data = nullCountFreq) + scale_x_discrete(limits = levels(curSheetData$GjeldendeFloranavnNINA)) +
      theme_classic() + theme(legend.position = "none", axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1))
    coverPlot <- ggplot(curSheetData[curSheetData$Dekn > 0, ]) + geom_dotplot(aes(x = GjeldendeFloranavnNINA, y = Dekn, fill = GjeldendeFloranavnNINA), binaxis = "y", stackdir = "center", stackratio = 1.5, dotsize = 0.4) +
      ylab("Dekning") + xlab("") + geom_text(aes(x = colName, y = 0.0, label = count, vjust = "top"), data = nullCountCover) + scale_x_discrete(limits = levels(curSheetData$GjeldendeFloranavnNINA)) +
      theme_classic() + theme(legend.position = "none", axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1))
    ggsave(paste(plotLoc, "Freq_", curEntry[1], "_", curEntry[2], ".svg", sep = ""), freqPlot, width = 6, height = 5)
    ggsave(paste(plotLoc, "Cover_", curEntry[1], "_", curEntry[2], ".svg", sep = ""), coverPlot, width = 6, height = 5)
    as.data.frame(cbind(varMatFreq, varMatCover))
  }, extraInfoLoc = extraInfoLoc, extraAliases = extraAliases, plotLoc = plotLoc)
  # Retrieve the names for all the extra entries in the information file
  everyColName <- sort(unique(unlist(lapply(X = allFrames, FUN = colnames))))
  # Combine it into a single data frame
  combinedFrame <- do.call(rbind, lapply(X = allFrames, FUN = function(curFrame, everyColName) {
    outFrame <- curFrame
    isInFrame <- sapply(X = everyColName, FUN = function(curColName, frameNames) { curColName %in% frameNames }, frameNames = colnames(curFrame))
    if(any(!isInFrame)) {
      outFrame <- cbind(curFrame, as.data.frame(
        matrix(NA, nrow = nrow(curFrame), ncol = sum(!isInFrame), dimnames = list(NULL, everyColName[!isInFrame]))
      ))
    }
    outFrame[, everyColName]
  }, everyColName = everyColName))
  combinedFrame
}, yearsToProcess = yearsToProcess, extraInfoLoc = extraInfoLoc, extraAliases = extraAliases, plotLoc = paste(outputDirectory, "disturbance", sep = "/"))
extraColNames <- sort(unique(unlist(lapply(X = allExtraInfo, FUN = colnames))))
# Combine all the information into a single data frame
damageInfo <- do.call(rbind, lapply(X = allExtraInfo, FUN = function(curInfo, extraColNames) {
  outFrame <- curInfo
  isInFrame <- sapply(X = extraColNames, FUN = function(curColName, frameNames) { curColName %in% frameNames }, frameNames = colnames(curInfo))
  if(any(!isInFrame)) {
    outFrame <- cbind(curInfo, as.data.frame(
      matrix(NA, nrow = nrow(curInfo), ncol = sum(!isInFrame), dimnames = list(NULL, extraColNames[!isInFrame]))
    ))
  }
  outFrame[, extraColNames]
}, extraColNames = extraColNames))
# Add the damage information to the plot information
plotInfo <- cbind(plotInfo, as.data.frame(matrix(NA, nrow = nrow(plotInfo), ncol = length(extraColNames), dimnames = list(NULL, extraColNames))))
plotInfo[row.names(damageInfo), extraColNames] <- damageInfo
# Create summary statistics for the damage information
damageWorkbook <- createWorkbook()
createSummaryStats <- function(valueColumn, inFrame, damageWorkbook) {
  allSites <- sort(unique(as.character(inFrame$SiteName)))
  setNames(lapply(X = allSites, FUN = function(curSite, inFrame, valueColumn, damageWorkbook) {
    varNameSplit <- strsplit(valueColumn, "[_\\/]", perl = TRUE)[[1]]
    sheetName <- paste(curSite,
      paste(sapply(X = varNameSplit[1:(length(varNameSplit) - 1)], FUN = function(curVar) {
        curValLetters <- strsplit(curVar, "")[[1]]
        paste(toupper(curValLetters[1]), curValLetters[2], sep = "")
      }), collapse = ""),
      varNameSplit[length(varNameSplit)],
      sep = " ")
    addWorksheet(damageWorkbook, sheetName)
    allYears <- sort(unique(as.character(inFrame$Year)))
    summarySheet <- sapply(X = allYears, FUN = function(curYear, curSite, inFrame, valueColumn) {
      # Retrieve a vector of values
      valueVec <- inFrame[as.character(inFrame$SiteName) == curSite & as.character(inFrame$Year) == curYear, valueColumn]
      setNames(
        c(mean(valueVec, na.rm = TRUE), sd(valueVec, na.rm = TRUE), quantile(valueVec, probs = c(0.25, 0.5, 0.75), na.rm = TRUE)),
        c("Mean", "SD", "1st Quartile", "Median", "3rd Quartile")
      )
    }, curSite = curSite, inFrame = inFrame, valueColumn = valueColumn)
    colnames(summarySheet) <- allYears
    writeDataTable(damageWorkbook, sheetName, as.data.frame(summarySheet), rowNames = TRUE)
    summarySheet
  }, inFrame = inFrame, valueColumn = valueColumn, damageWorkbook = damageWorkbook), allSites)
}
damageSummary <- setNames(
  lapply(X = colnames(damageInfo), FUN = createSummaryStats, inFrame = plotInfo[rownames(damageInfo), ], damageWorkbook = damageWorkbook),
  colnames(damageInfo)
)
saveWorkbook(damageWorkbook, paste(outputDirectory, "disturbanceSummary.xlsx", sep = "/"), overwrite = TRUE)

# ------ 10. PRODUCE SOIL CHEMISTY REPORTS ------
# Function to aggregate the chemistry means
retrieveChemistryMeans <- function(inTable, plotCol, yearCol, pHCol, LOICol, NCol, yearsToProcess, siteInfo) {
  do.call(rbind, lapply(X = rownames(siteInfo), FUN = function(curSite, inTable, plotCol, yearCol, pHCol, LOICol, NCol, yearsToProcess, siteInfo) {
    possSiteCodes <- strsplit(as.character(siteInfo[curSite, "SiteCode"]), "|", fixed = TRUE)[[1]]
    siteCodes <- sapply(X = strsplit(as.character(inTable[, plotCol]), ""), FUN = function(curLetters) { curLetters[1] })
    curTable <- inTable[siteCodes %in% possSiteCodes & as.integer(inTable[, yearCol]) %in% yearsToProcess, ]
    curTable <- cbind(curTable, data.frame(
      newPlotCode = paste(
        possSiteCodes[1],
        formatC(as.integer(gsub("^\\D+\\s*", "", curTable[, plotCol], perl = TRUE)), width = 2, flag = "0"),
        "-",
        curTable[, yearCol], sep = "")
    ))
    tempFrame <- aggregate(curTable[, c(pHCol, LOICol, NCol)], by = list(curTable$newPlotCode), FUN = mean, na.rm = TRUE)
    colnames(tempFrame) <- c("plotID", "pH", "LOI", "N")
    rownames(tempFrame) <- tempFrame$plotID
    tempFrame[, 2:4]
  }, inTable = inTable, plotCol = plotCol, yearCol = yearCol, pHCol = pHCol, LOICol = LOICol, NCol = NCol, yearsToProcess = yearsToProcess, siteInfo = siteInfo))
}
# Create one information file with both the earlier and later chemistry samples
soilChemistryInfo <- rbind(
  retrieveChemistryMeans(rawVegetationData[[3]]$table, "REG_SOM_Rute_id", "REG_SOM_AAR", "E3 pH", "Glødetap", "KjN", yearsToProcess, siteInfo),
  retrieveChemistryMeans(rawVegetationData[[4]]$table, "oppdragsgiverens prøve-identifikasjon ", "År", "E3pH ", "gltap % ", "Kj-N mmol/kg (32) ", yearsToProcess, siteInfo)
)
soilChemistryInfo <- cbind(soilChemistryInfo, data.frame(
  LOI_N = soilChemistryInfo$LOI * 100 / soilChemistryInfo$N
))
# Aggregate the chemistry information at the site and year level
aggChemInfo <- aggregate(soilChemistryInfo, by = list(
  site = sapply(X = gsub("\\d+\\-\\d+$", "", rownames(soilChemistryInfo), perl = TRUE), FUN = function(curSiteCode, siteInfo) {
    siteCodes <- strsplit(as.character(siteInfo$SiteCode), "|", fixed = TRUE)
    as.character(siteInfo$SiteName)[sapply(X = siteCodes, FUN = function(curPoss, curSiteCode) { curSiteCode %in% curPoss }, curSiteCode = curSiteCode)]
  }, siteInfo = siteInfo),
  year = as.integer(gsub("^\\D+\\d+\\-", "", rownames(soilChemistryInfo), perl = TRUE))
), FUN = mean, na.rm = TRUE)
lapply(X = unique(as.character(aggChemInfo$site)), FUN = function(curSite, aggChemInfo, plotLoc) {
  pHplot <- ggplot(aggChemInfo[as.character(aggChemInfo$site) == curSite & !is.na(aggChemInfo$pH), ], aes(x = year, y = pH)) + geom_line(colour = "blue") + geom_point(colour = "blue") +
    theme_classic() + xlab("År") + ylab("pH")
  ggsave(paste(plotLoc, "pH_", curSite, ".svg", sep = ""), pHplot, width = 5, height = 4)
  LOI_Nplot <- ggplot(aggChemInfo[as.character(aggChemInfo$site) == curSite & !is.na(aggChemInfo$LOI_N), ], aes(x = year, y = LOI_N)) + geom_line(colour = "red") + geom_point(colour = "red") +
    theme_classic() + xlab("År") + ylab("LOI/Kj-N * 100")
  ggsave(paste(plotLoc, "LOI_N_", curSite, ".svg", sep = ""), LOI_Nplot, width = 5, height = 4)
}, aggChemInfo = aggChemInfo, plotLoc = paste(outputDirectory, "chem_", sep = "/"))
# Append the chemistry information to the plot information
plotInfo <- cbind(plotInfo, as.data.frame(matrix(NA, ncol = ncol(soilChemistryInfo), nrow = nrow(plotInfo), dimnames = list(NULL, colnames(soilChemistryInfo)))))
plotInfo[rownames(soilChemistryInfo), colnames(soilChemistryInfo)] <- soilChemistryInfo

# ------ 11. WRITE RESTRUCTURED DATA TO THE OUTPUT DIRECTORY ------
# Save an RDS file containing all the processed data
save(
  siteInfo,
  sjiktInfo,
  speciesInfo,
  plotInfo,
  samplingInfo,
  quadratData,
  communityMatrixFreq,
  communityMatrixCover,
  yearlySpeciesFreq,
  yearlySpeciesCover,
  ellenbergFreq,
  ellenbergCover,
  DCAFreq,
  DCACover,
  NMDSFreq,
  NMDSCover,
  file = paste(outputDirectory, "behandledeDataTabeller_Vegetasjon.RData", sep = "/"))
# Also creat an Excel workbook with sheets for each of the data tables
ordHeaderStyle <- createStyle(border = "bottom", borderStyle = "medium", textDecoration = "bold")
dataTableBook <- createWorkbook()
# Function to create worksheets and add them to the data table workbook
createExtraWorksheets <- function(inOb, outBook, exPrefix, ordHeaderStyle) {
  if(is.data.frame(inOb) || is.matrix(inOb)) {
    addWorksheet(outBook, exPrefix)
    writeDataTable(outBook, exPrefix, as.data.frame(inOb), headerStyle = ordHeaderStyle, rowNames = TRUE, firstColumn = TRUE)
  } else if(is.list(inOb)) {
    lapply(X = names(inOb), FUN = function(curName, inOb, outBook, exPrefix, ordHeaderStyle) {
      newSheetName <- paste(exPrefix, curName, sep = "")
      addWorksheet(outBook, newSheetName)
      writeDataTable(outBook, newSheetName, as.data.frame(inOb[[curName]]), headerStyle = ordHeaderStyle, rowNames = TRUE, firstColumn = TRUE)
    }, inOb = inOb, outBook = outBook, exPrefix = exPrefix, ordHeaderStyle = ordHeaderStyle)
  }
}
createExtraWorksheets(siteInfo, dataTableBook, "siteInfo", ordHeaderStyle)
createExtraWorksheets(sjiktInfo, dataTableBook, "sjiktInfo", ordHeaderStyle)
createExtraWorksheets(plotInfo, dataTableBook, "plotInfo", ordHeaderStyle)
createExtraWorksheets(samplingInfo, dataTableBook, "samplingInfo_", ordHeaderStyle)
createExtraWorksheets(speciesInfo, dataTableBook, "speciesInfo", ordHeaderStyle)
createExtraWorksheets(quadratData, dataTableBook, "quadratData", ordHeaderStyle)
createExtraWorksheets(communityMatrixFreq, dataTableBook, "comMatFreq_Sjikt", ordHeaderStyle)
createExtraWorksheets(communityMatrixCover, dataTableBook, "comMatCover_Sjikt", ordHeaderStyle)
saveWorkbook(dataTableBook, paste(outputDirectory, "behandledeDataTabeller_Vegetasjon.xlsx", sep = "/"), overwrite = TRUE)
