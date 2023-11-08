#################################################################################
##
## problem: how to download automatically large amount of files?
## solution: double loop for extration
## 
## warning: time consuming process (circa 60 minutes on my workstation)
## warning: large disk space necessary (3 GB)
##
#################################################################################

setwd("d:\\Projects\\Imgw\\")

month <- c("01","02","03","04","05","06","07","08","09","10","11","12")
year <- as.character(c(1951:2022))

### download all the files

for (i in 1:length(year)) {

url <- paste0("https://danepubliczne.imgw.pl/data/dane_pomiarowo_obserwacyjne/dane_hydrologiczne/dobowe/",year[i],"/")
  
  for (j in 1:length(month)) {
    
      destfile <- paste0("codz_",year[i],"_",month[j],".zip")
    
      download.file(paste0(url, destfile), destfile)
                     
  }
}

### unzip all files

for (i in list.files(getwd())) {

  unzip(i)

}

### merge files together

header <- c("Kod stacji", "Nazwa stacji", "Nazwa rzeki/jeziora", "Rok hydrologiczny", "Wskaźnik miesiąca w roku hydrologicznym",
"Dzień", "Stan wody [cm]", "Przepływ [m^3/s]", "Temperatura wody [st. C]", "Miesiąc kalendarzowy")
  
files <- list.files(pattern = "\\.csv$")

DF <-  read.csv(files[1], header = FALSE, col.names = header)

#reading each file within the range and append them to create one file

for (f in files[-1]){
  df <- read.csv(f, header = FALSE, col.names = header)      # read the file
  DF <- rbind(DF, df)    # append the current file
}

#writing the appended file  

write.csv(DF, "Appended.Data.csv", row.names=FALSE, quote=FALSE)
