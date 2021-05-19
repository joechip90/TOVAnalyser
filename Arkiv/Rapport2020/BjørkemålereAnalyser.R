library(RODBC)
library(openxlsx)
library(ggplot2)
library(ggpubr)
library(reshape2)
library(extrafont)
if(!any(names(windowsFonts()) == "Calibri")) {
  # Import the Calibri font if it is not installed on the system
  font_import(prompt = FALSE)
  loadfonts(device = "win")
}

# ------ 1. SET REPORT GENERATION CRITERIA ------
# Sites to produce analyses for.  Is a list with names elements, one for each site. Each element is a character vector containing site codes
# used for the site in the TOV database (some sites may have multiple synonyms)
sitesToProcess <- list(
  "Dividal" = c("D"),
  "Børgefjell" = c("B"),
  "Åmotsdal" = c("Å"),
  "Gutulia" = c("G"),
  "Møsvatn" = c("M")
)
# Names to convert for figures
namesToChange <- setNames(c("Åmotsdal", "Dividal"), c("Åmotsdalen", "Dividalen"))
# Years to produce analyses for (including years you wish to generate time series of species for)
yearsToProcess <- c(2014, 2015, 2016, 2017, 2018, 2019, 2020)
# This source file relies on the setting of the "WORKSPACE_TOV" variable in the .Renviron file.  This variable should contain the
# directory of the TOV project directory.  This variable can be set easily using the usethis::edit_r_environ() function.
tovDirectory <- Sys.getenv("WORKSPACE_TOV")
# Set the output folder for the analysis
outputDirectory <- paste(tovDirectory, "Behandlede_datatabeller", "RapportAnalyser_2020Bjørkemålere", sep = "/")

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
# Retrieve the moth data
rawMothData <- retrieveTOVData(c("Bjorkemaler_Stasjonsinfo_pr_ar", "Bjørkemåler_Stasjonsinfo", "Bjørkemåler_Data"))

# ------ 4. RESTRUCTURE DATA FOR REPORT ------
# Retrieve the station information
stationInfo <- merge(rawMothData[[1]]$table[, c("StasjonsID", "Omradenavn", "Aar")], rawMothData[[2]]$table[, c("StasjonsID", "Tre-høyde")], by = "StasjonsID")
colnames(stationInfo) <- c("StationID", "Site", "Year", "TreeHeight")
rownames(stationInfo) <- paste(stationInfo$StationID, stationInfo$Year, sep = "-")
# Make a series of aggregated counts and grazing damage
aggCountData <- aggregate(rawMothData[[3]]$table[, c("AntallEpirrita", "AntallOperophtera", "AntallUkjent")], by = list(StationID = rawMothData[[3]]$table$StasjonsID, Year = rawMothData[[3]]$table$År), FUN = mean, na.rm = TRUE)
colnames(aggCountData)[3:5] <- c("MeanEpirrita", "MeanOperophtera", "MeanUnknown")
aggDamageData <- aggregate(
  as.double(gsub("\\s*,\\s*", ".", as.character(rawMothData[[3]]$table[, c("%bladBeitet")]), perl = TRUE)),
  by = list(StationID = rawMothData[[3]]$table$StasjonsID, Year = rawMothData[[3]]$table$År), FUN = mean, na.rm = TRUE)
colnames(aggDamageData)[3] <- "GrazingDamagePercentage"
rownames(aggCountData) <- paste(aggCountData$StationID, aggCountData$Year, sep = "-")
rownames(aggDamageData) <- paste(aggDamageData$StationID, aggDamageData$Year, sep = "-")
# Add the counts and grazing damage to the station information
stationInfo <- cbind(stationInfo, as.data.frame(
  matrix(NA, ncol = 4, nrow = nrow(stationInfo), dimnames = list(NULL, c(colnames(aggCountData[3:ncol(aggCountData)]), colnames(aggDamageData[3:ncol(aggDamageData)]))))
))
stationInfo[rownames(aggCountData), colnames(aggCountData)[3:ncol(aggCountData)]] <- aggCountData[, 3:ncol(aggCountData)]
stationInfo[rownames(aggDamageData), colnames(aggDamageData)[3:ncol(aggDamageData)]] <- aggDamageData[, 3:ncol(aggDamageData)]
# Remove those entries that are not in the required analysis
stationInfo <- stationInfo[as.character(stationInfo$Site) %in% names(sitesToProcess) & as.integer(stationInfo$Year) %in% yearsToProcess, ]
# Add in the hunnrakler information
aggHunnraklerData <- aggregate(rawMothData[[3]]$table[, "AntallHunnrakler"], by = list(StationID = rawMothData[[3]]$table$StasjonsID, Year = rawMothData[[3]]$table$År), FUN = mean, na.rm = TRUE)
stationInfo <- cbind(stationInfo, data.frame(
  MeanHunnrakler = rep(NA, nrow(stationInfo))
))
stationInfo[paste(aggHunnraklerData$StationID, aggHunnraklerData$Year, sep = "-"), "TotalHunnrakler"] <- aggHunnraklerData$x

# ------ 5. MAKE FIGURES OF LEAF DAMAGE AND MOTH POPULATIONS ------
popBook <- createWorkbook()
lapply(X = names(sitesToProcess), FUN = function(curSite, stationInfo, plotLoc, outBook) {
  curInfo <- stationInfo[as.character(stationInfo$Site) == curSite, ]
  curAgg <- aggregate(curInfo[, 5:ncol(curInfo)], list(Year = curInfo$Year), FUN = mean, na.rm = TRUE)
  # Melt a frame of population values
  popPlotFrame <- melt(
    curAgg,
    id.vars = c("Year"), measure.vars = c("MeanEpirrita", "MeanOperophtera", "MeanUnknown"), variable.name = "Species")
  popPlotFrame$Species <- factor(as.character(popPlotFrame$Species), levels = c("MeanEpirrita", "MeanOperophtera", "MeanUnknown"), labels = c("Fjellbjørkemåler", "Liten høstmåler", "Ukjent måler"))
  popPlot <- ggplot(popPlotFrame, aes(x = Year, y = value, colour = Species)) + geom_line(size = 0.53) + theme_classic() + xlab("År") + ylab("Antall målerlarver") +
    # Ensure that break only occur at integer values
    scale_y_continuous(breaks = function(x) unique(floor(pretty(seq(0, (max(x) + 1) * 1.1))))) +
    theme(
      legend.title = element_blank(),
      text = element_text(family = "Calibri"),
      axis.text = element_text(size = 9),
      legend.text = element_text(size = 8),
      axis.title = element_text(size = 10),
      axis.line = element_line(size = 0.18, colour = grey(level = 0.35))
    )
  damagePlot <- ggplot(curAgg, aes(x = Year, y = GrazingDamagePercentage)) + geom_line() + theme_classic() + xlab("År") + ylab("Blader med beiteskade (%)") +
    scale_y_continuous(limits = c(0, NA)) +
    theme(
      legend.title = element_blank(),
      text = element_text(family = "Calibri"),
      axis.text = element_text(size = 9),
      legend.text = element_text(size = 8),
      axis.title = element_text(size = 10),
      axis.line = element_line(size = 0.18, colour = grey(level = 0.35))
    )
  combinedPlot <- ggarrange(popPlot, damagePlot, labels = c("A", "B"), common.legend = TRUE, legend = "right")
  ggsave(paste(plotLoc, "/populationPlot_", curSite, ".svg", sep = ""), popPlot, width = 5, height = 3)
  ggsave(paste(plotLoc, "/damagePlot_", curSite, ".svg", sep = ""), damagePlot, width = 4, height = 3)
  ggsave(paste(plotLoc, "/combinedPlot_", curSite, ".svg", sep = ""), combinedPlot, width = 9, height = 3)
  # Create a worksheet for the current site
  addWorksheet(outBook, curSite)
  writeDataTable(outBook, curSite, curAgg)
}, stationInfo = stationInfo, plotLoc = outputDirectory, outBook = popBook)
saveWorkbook(popBook, paste(outputDirectory, "populationSummary.xlsx", sep = "/"), overwrite = TRUE)
renameSites <- function(curVal, nameSite) {
  outVal <- as.character(curVal)
  if(outVal %in% nameSite) {
    outVal <- names(nameSite)[which(outVal == nameSite)[1]]
  }
  outVal
}
stationInfo$Site <- factor(sapply(X = stationInfo$Site, FUN = renameSites, nameSite = namesToChange), levels = sapply(X = names(sitesToProcess), FUN = renameSites, nameSite = namesToChange))

hunnraklerPlot <- ggplot(
  aggregate(stationInfo$TotalHunnrakler, by = list(site = stationInfo$Site, year = as.character(stationInfo$Year)), FUN = mean, na.rm = TRUE),
  aes(x = year, y = x)) + geom_col(aes(fill = site), position = "dodge") + xlab("År") + ylab("Antall hunnrakler") + theme_classic() +
  scale_fill_brewer(palette = "Paired") +
  theme(
    legend.title = element_blank(),
    text = element_text(family = "Calibri"),
    axis.text = element_text(size = 9),
    legend.text = element_text(size = 8),
    axis.title = element_text(size = 10),
    axis.line = element_line(size = 0.18, colour = grey(level = 0.35))
  )
ggsave(paste(outputDirectory, "hunnraklerPlot.svg", sep = "/"), hunnraklerPlot, width = 6, height = 3)

# ------ 6. WRITE RESTRUCTURED DATA TO THE OUTPUT DIRECTORY ------
# Save an RDS file containing all the processed data
save(
  stationInfo,
  file = paste(outputDirectory, "behandledeDataTabeller_Bjørkemålere.RData", sep = "/"))
