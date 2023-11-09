## Import RODBC package

if(!require(RODBC)){
  install.packages("RODBC")
  library(RODBC)
}

## Set up driver info and database path
DRIVERINFO <- "Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
MDBPATH <- "C:/Users/rpuak/OneDrive/Dokumenty/nation.accdb"
PATH <- paste0(DRIVERINFO, "DBQ=", MDBPATH)

## Establish connection
channel <- odbcDriverConnect(PATH)

## Get all tables available
tables <- sqlTables(channel)[c("TABLE_NAME","TABLE_TYPE")][sqlTables(channel)$TABLE_TYPE=="TABLE",]
print(tables$TABLE_NAME)

## Available columns

columns <- c()

  for (i in tables$TABLE_NAME) {
  
    column <- sqlColumns(channel, i)[c("TABLE_NAME","ORDINAL","COLUMN_NAME","TYPE_NAME")]
    columns <- rbind(columns, column)
  
  }
  
    print(columns)

## Run exemplary SQL queries
print(sqlQuery(channel, 
"SELECT *
 FROM [Countries]",
stringsAsFactors = FALSE))

### FINE!
print(sqlQuery(channel, "SELECT name, area, region_id
                          FROM Countries
                          WHERE name like '%stan' 
                          ORDER BY name"))

### FINE!
print(sqlQuery(channel, 
"SELECT name, COUNT(*) As Counts
 FROM Countries
 GROUP BY name", 
stringsAsFactors = FALSE))

### FINE!
print(sqlQuery(channel, 
  "SELECT DISTINCT region_id
   from countries", 
stringsAsFactors = FALSE))

### FINE!
print(sqlQuery(channel,
"SELECT 
region_id,
COUNT(region_id) As 'Counts'
from countries
group by region_id", # note difference!
stringsAsFactors = FALSE))

### FINE!
print(sqlQuery(channel, 
               "SELECT 
               COUNT(area) As 'Counts',
               SUM(area) As 'Total Area',
               AVG(area) As 'Average Area'
               from countries",
stringsAsFactors = FALSE))

### FINE!
print(sqlQuery(channel, 
"SELECT 
region_id,
COUNT(area) As 'Counts',
SUM(area) As 'Total Area',
AVG(area) As 'Average Area'
from countries
group by region_id",
stringsAsFactors = FALSE))

### FINE!
print(sqlQuery(channel, "SELECT regions.name AS 'region name',
                         COUNT(country_id) AS Counts,
                         SUM(area) AS Sum
                         FROM Countries INNER JOIN regions
                         ON Countries.region_id=regions.region_id
                         GROUP BY regions.name", 
                        stringsAsFactors = FALSE))

### FINE!
print(sqlQuery(channel,"
SELECT TOP 10 *
FROM Countries INNER JOIN regions
ON Countries.region_id=regions.region_id", 
stringsAsFactors = FALSE))

### FINE!
print(sqlQuery(channel,"
Select *,
(gdp / population) AS 'Gdp per capita'	
from country_stats		
where 
country_id = 221 AND year = 1988 OR
country_id = 221 AND year = 2018", 
stringsAsFactors = FALSE))

## Close and remove channel
close(channel)
rm(channel)
