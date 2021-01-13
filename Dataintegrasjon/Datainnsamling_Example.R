library(RODBC)
library(XLConnect)

# This source file relies on the setting of the "WORKSPACE_TOV" variable in the .Renviron file.  This variables should contain the
# local directory of the TOVAnalyser repository.  This variable can be set easily using the usethis::edit_r_environ() function.
source(paste(Sys.getenv("WORKSPACE_TOV"), "Dataintegrasjon", "Datainnsamling.R", sep = "/"))
