### MS Data to Heat Map for Zhengxi###

#Step 1: Enter your variables
filename <- "CholRF06142019.xlsx"  #The name of your data file. Must include ".xlsx"!
plate <- 384  #The number of wells in your plate
numplates <- 1  #How many separate plates are in your file
exportgrid <- "CholRF06142019 reformatted.xlsx"  #Choose a file name for the exported excel sheet

#Step 2: Navigate to the directory containing your data file
#Find it in the files tab, then click More and Set As Working Directory

#Step 3: Run the code
#Press control a to select all the code, and then press control enter to run.
#You can find the exported excel sheet in your current directory!


#########################################################################################################


#Load required packages
check.packages <- function(pkg){
    new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
    if (length(new.pkg)) 
        install.packages(new.pkg, dependencies = TRUE)
    sapply(pkg, require, character.only = TRUE)
}
packages <- c("openxlsx", "stringr", "gplots", "grDevices")
check.packages(packages)

#Load file
data <- openxlsx::read.xlsx(filename, startRow = 2)
df <- data.frame("Data.File" = data$Data.File, "Ratio" = data[,9] / data[,7])
   #Ratio of column 9 over column 7

#Organize data
df$rows <- stringr::str_extract(df$Data.File, "[a-zA-Z](?=\\d+.d)")
    #Regex: Looks for a letter that's followed by a number and .d
df$rows <- match(df$rows, LETTERS[1:26])
df$columns <- as.numeric(stringr::str_extract(df$Data.File, "(\\d+)(?!.*\\d)"))
    #Regex: Looks for a number that's not followed by any other numbers
df <- df[(which(!is.na(df$rows))),]

#Define functions
makegrid <- function() {
    if (plate == 384) {
        grid <- matrix(ncol = 24, nrow = 16)
        return(grid)
    } else if (plate == 96) {
        grid <- matrix(ncol = 12, nrow = 8)
        return(grid)
    } else {
        print("Error! Plate size must be 384 or 96.")
    }
}

hm <- function(x) {
    gplots::heatmap.2(x, na.rm = TRUE, dendrogram = "none", Rowv = FALSE, Colv = FALSE, trace = "none",
                     density.info = "none", col = colorRampPalette(c("yellow", "orange", "red")))
}

#Process data
if (numplates == 1) {
    #Make grid
    grid <- makegrid()
    for (i in 1:nrow(df)) {   #Fills in empty grid using values from df
        grid[df[i, 3], df[i, 4]] <- df[i, 2]
    }
    grid[which(is.infinite(grid))] <- NA
    
    #Export file
    wb <- openxlsx::createWorkbook(exportgrid)
    openxlsx::addWorksheet(wb, sheetName = "Sheet 1")
    openxlsx::writeData(wb, "Sheet 1", grid, keepNA = TRUE, colNames = FALSE)
    saveWorkbook(wb, exportgrid, overwrite = TRUE)
    
    #Export heat map
    grDevices::pdf(file = "Heat Map.pdf")
    hm(grid)
    grDevices::dev.off()
} else {
    #Make list of grids
    ls <- list()
    for (p in 1:numplates) {
        grid <- makegrid()
        empty <- plate - nrow(df) / numplates  #Finds the number of empty wells per plate
        if (empty == 0) empty <- plate
        for (i in 1:(nrow(df) / numplates)) {   #Fills in empty grid using values from df
            r <- i*p + (p-1)*(nrow(df)/numplates - i)
            grid[df[r , 3], df[r, 4]] <- df[r, 2]
        }
        grid[which(is.infinite(grid))] <- NA
        ls[[p]] <- grid
    }
    
    #Export excel file
    wb <- openxlsx::createWorkbook(exportgrid)
    for (p in 1:numplates) {
        openxlsx::addWorksheet(wb, sheetName = paste0("Sheet", p))
        openxlsx::writeData(wb, paste0("Sheet", p), ls[[p]], keepNA = TRUE, colNames = FALSE)
    }
    saveWorkbook(wb, exportgrid, overwrite = TRUE)
    
    #Create heat map
    dir.create("Heat Maps/")
    
    for (i in 1:length(ls)) {
        #ls[[i]][1:8,1] <- NA   #Sets control buffer wells to NA
        grDevices::pdf(file = paste0(getwd(), "/", "Heat Maps/Plate ", i, ".pdf"))
        hm(ls[[i]])
        grDevices::dev.off()
    }
}
