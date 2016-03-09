library(dplyr)
library(xlsx)
library(matrixStats)
# we extract all the sheets 
for (i in 1:17){ 
  
  sheets <- read.xlsx2("Results1.xlsx",i)
  colnames(sheets) <- c("Colony", "Item", "y1952", "y1953", "y1954", "y1955", "y1956", "y1957", "y1958", "y1959", "y1960", "y1961", "y1962" , "y1963", "y1964")
  x <- sheets[,3:15]
  x <- sapply(x, as.character)
  x <- matrix(as.numeric(unlist(x)),nrow=nrow(x))
  # we constructed our row means and std dev 
  rmean <- rowMeans(x, na.rm = TRUE)
  row_sd <- rowSds(x, na.rm = TRUE)
  
  print(paste("Data for sheet: ", i))
  for (j in 1:63){ 
    
    y <- x[j,]
    m <- rmean[j]
    st_dev <- row_sd[j]
    if (is.nan(m) | is.nan(st_dev) | is.na(m) | is.na(st_dev)){
      next
    }
    v <- ifelse(y <= m + 2*st_dev & y >= m - 2*st_dev, 0, 1)
    if (sum(v, na.rm = TRUE) > 0){
      print(paste("Error in line: ",j))
     }
    }
}

  
  
  