## 1. ---- FUNCTION TO RETRIEVE A LIST OF DATA TABLES AND LOCATIONS ----
#' Retrieve data table information associated with the "program for terrestrisk naturovervåking" (TOV)
#'
#' @return A \code{list} with one element for each data element.  Each element is itself a list
#' with the following named elements:
#' \describe{
#'     \item{beskrivelseNorsk}{A description (in Norwegian) of the data table}
#'     \item{descriptionEnglish}{A description (in English) of the data table}
#'     \item{location}{The location of the dataset according to the internal NINA filesystem}
#' }
#' @author Joseph D. Chipperfield, \email{joechip90@@googlemail.com}
#'
retrieveTOVTableInfo <- function() { list(
  # 1.1 ==== General information tables shared by multiple datasets ====
  "Analyse_nummerserier" = setNames(list(
    "Info om områdets analysenerserie",
    "Information on the area's analysis-series",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "ARTSLISTE" = setNames(list(
    "Artsliste, felles",
    "Species list",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "Felt_info" = setNames(list(
    "Info om felt, feltinfo-ID, UTM mm",
    "Field information",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "Felt_pr_aar" = setNames(list(
    "Info om år og gjentak, samt antall trær og gadd for epifytt",
    "Information on the sampling strategy for the epiphytes",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "Omr_info" = setNames(list(
    "Info om område, område-ID",
    "Information on the sampling area",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "Omr_pr_aar" = setNames(list(
    "Hvilke år de ulike områdene er besøkt, gjentak, samt antall trær og gadd for epifytt",
    "Information on the visitation and sampling regimes for the different sites",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "T_SJIKT" = setNames(list(
    "Koder for sjikt, bruke i artsliste og vegetasjonsdata",
    "Layer codes for species lists and vegetation data",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  # 1.2 ==== Vegetation data information tables ====
  "VEG_Analyseinfo" = setNames(list(
    "Info om analysetidspunkt, hvem som har analysert og kommentarer knyttet til ruten (1x1m)",
    "Informtation about the time of the vegetation analysis and comments related to the transects",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "VEG_Analyse_rutenivaa" = setNames(list(
    "Frekvens og dekning for arter på rutenivå (data pr rute, 1x1m)",
    "Frequency and cover of species at the transect-level",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "VEG_Analyse_smaarutenivaa" = setNames(list(
    "Tilstedeværelse av arter på smårutenivå (16 småruter pr rute (1x1m))",
    "Presence of species at the sub-transect level (16 squares per transect)",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "VEG_Totaldekn" = setNames(list(
    "Dekningsdata PR 1x1m rute, total dekning, busksjikt, feltsjikt, bunnsjikt, strø mm.",
    "Coverage data per transect: total coverage, layers, and litter",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "VEG_Kjemi_data_pr_rute" = setNames(list(
    "Kjemidata for alle ruter (1x1m) 1993-till 2012",
    "Transect-based chemical data for transects between 1993 and 2012",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "VEG_Kjemi_data_pr_felt" = setNames(list(
    "Forenklet kjemiserie pr felt fra 2013-2018",
    "Simplified chemistry series per field 2013-2018",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "Skade_Paa_Blaabaer" = setNames(list(
    "Ekstra data ikke inkludert i basen",
    "Extra data pertaining to observations not included in the database",
    "R:/Prosjekter/15450000 - TOV-vegetasjon/Skade blåbær/Skade blåbær 2012-2018.xlsx"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  # 1.3 ==== Epiphyte data information tables ====

  # TODO: include epiphyte data tables here

  # 1.4 ==== Bjøkemålere data information tables ====

  "Bjorkemaler_Stasjonsinfo_pr_ar" = setNames(list(
    "Stasjon område informasjon",
    "Station area information",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "Bjørkemåler_Stasjonsinfo" = setNames(list(
    "Stasjon informasjon",
    "Station information",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location")),

  "Bjørkemåler_Data" = setNames(list(
    "Bjørkemålere antallere",
    "Geometridae moth counts",
    "TidsserieDB"
    ), c("beskrivelseNorsk", "descriptionEnglish", "location"))
) }

## 2. ---- FUNCTION TO RETRIEVE TOV DATA ----
#' Retrieve data tables associated with the "program for terrestrisk naturovervåking" (TOV)
#'
#' @param tableRetrieve A \code{character} vector containing the names of the data tables that
#' you wish to retrieve
#'
#' @return A \code{list} with one element for each data element.  Each element is itself a list
#' with the following named elements:
#' \describe{
#'     \item{beskrivelseNorsk}{A description (in Norwegian) of the data table}
#'     \item{descriptionEnglish}{A description (in English) of the data table}
#'     \item{location}{The location of the dataset according to the internal NINA filesystem}
#'     \item{table}{The data table}
#' }
#' @author Joseph D. Chipperfield, \email{joechip90@@googlemail.com}
#'
retrieveTOVData <- function(tableRetrieve = names(retrieveTOVTableInfo())) {
  # Initialise an output list
  outList <- list()
  # Initialise a meta-data list
  inTOVTableInfo <- retrieveTOVTableInfo()
  names(inTOVTableInfo) <- iconv(names(retrieveTOVTableInfo()), to = localeToCharset(Sys.getlocale("LC_CTYPE")), from = "UTF-8")
  metaDataOut <- inTOVTableInfo[tableRetrieve[tableRetrieve %in% names(inTOVTableInfo)]]
  # 2.1. ==== Retrieve those table elements that are in the TidsserieDB database ====
  dataInTidsserieDB <- metaDataOut[sapply(X = metaDataOut, FUN = function(curTableInfo) { curTableInfo[["location"]] == "TidsserieDB" })]
  # Initialise a connection to the "TidsserieDB" database
  tidsserieDBConnection <- odbcConnect("TidsserieDB")
  if(length(dataInTidsserieDB) > 0) {
    if(tidsserieDBConnection >= 0) {
     outList <- append(outList, lapply(X = names(dataInTidsserieDB), FUN = function(curTableName, tidsserieDBConnection, dataInTidsserieDB) {
       append(
         dataInTidsserieDB[[curTableName]],
         list(
          table = sqlFetch(tidsserieDBConnection, iconv(curTableName, from = localeToCharset(Sys.getlocale("LC_CTYPE")), to = "UTF-8"))
         )
        )
     }, tidsserieDBConnection = tidsserieDBConnection, dataInTidsserieDB = dataInTidsserieDB))
      # Close the database connection
      odbcClose(tidsserieDBConnection)
    } else {
      stop("unable to connect to tidsserieDB database")
    }
  }
  # 2.2. ==== Retrieve the "Skade på blåbær" information sheet ====
  # TODO: include the "Skade på blåbær" information table in the output
  #if("Skade_Paa_Blaabaer" %in% tableRetrieve) {
  #  # Open the relevant workbook containing the information
  #  skadeWorkbook <- loadWorkbook(retrieveTOVTableInfo()[["Skade_Paa_Blaabaer"]][["location"]])
  #  metaDataOut[["Skade_Paa_Blaabaer"]] <- append(metaDataOut[["Skade_Paa_Blaabaer"]], list(
  #    SpeciesList = setNames(readWorksheet(skadeWorkbook, "Artliste_mengde", startRow = 1, startCol = 2, endRow = 7, endCol = 3, header = FALSE), c("Species", "Comments")),
  #    Level = setNames(readWorksheet(skadeWorkbook, "Artliste_mengde", startRow = 13, startCol = 1, endRow = 15, endCol = 2, header = FALSE), c("Level", "PercentageBlueberries"))))
  #}
  outList
}
