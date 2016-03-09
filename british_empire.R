# install.packages('xlsx')  # CD: for my benefit
# install.packages('dplyr')


library(XLConnect)
library(dplyr)
library(xlsx)

 
#let's create the worksheet to which we will write our data 
results <- xlsx::createWorkbook()

# we create files for the individuals sheets 
y1952 <- read.xlsx("British_Empire.xlsx",1)
y1953 <- read.xlsx("British_Empire.xlsx",2)
y1954 <- read.xlsx("British_Empire.xlsx",3)
y1955 <- read.xlsx("British_Empire.xlsx",4)
y1956 <- read.xlsx("British_Empire.xlsx",5)
y1957 <- read.xlsx("British_Empire.xlsx",6)
y1958 <- read.xlsx("British_Empire.xlsx",7)
y1959 <- read.xlsx("British_Empire.xlsx",8)
y1960 <- read.xlsx("British_Empire.xlsx",9)
y1961 <- read.xlsx("British_Empire.xlsx",10)
y1962 <- read.xlsx("British_Empire.xlsx",11)
y1963 <- read.xlsx("British_Empire.xlsx",12)
y1964 <- read.xlsx("British_Empire.xlsx",13)

# we note that each year has different no of entries 

colonies_50s <- Reduce(union, list(y1952$colony,y1953$colony,y1954$colony,y1955$colony,y1956$colony,y1957$colony,y1958$colony,y1959$colony))
colonies <- Reduce(union, list(colonies_50s,y1960$colony,y1961$colony,y1962$colony,y1963$colony,y1964$colony))
items <- unique(y1952$X_j)
years <- list(y1952,y1953,y1954,y1955,y1956,y1957,y1958,y1959,y1960,y1961,y1962,y1963,y1964)
solution <- list()
year_no <- seq(1952,1964,1)

for(item in items){ 
    k <- 1
    restructed_data = data.frame(Colonies = colonies, item = item)
    for(year in years){ 
      data_line <- dplyr::filter(year, X_j == item)
      item_data <- data.frame(data_line$colony, data_line$entry)
      colnames(item_data) <- c("Colonies",year_no[k])
      restructed_data <- merge(restructed_data,item_data,by = "Colonies", all.x = TRUE)
      k <- k + 1
    }
    sheet <- createSheet(wb = results, sheetName = item)
    addDataFrame(x = restructed_data,sheet = sheet)
}

saveWorkbook(results,"Results.xlsx")


