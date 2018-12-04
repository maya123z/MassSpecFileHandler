#Step 1: Navigate to the directory containing your data file
#Find it in the files tab, then click More and Set As Working Directory

#Step 2: How many files do you want to process?
#Enter FALSE to process a single data file. Enter TRUE to process all the files in your working directory.

iterate <- FALSE

    #For a single data file, enter the information below:

    filename <- "your data file.xlsx"  #The name of your data file. Don't forget the extension!
    plate <- 384  #The number of wells in your plate (must be 384 or 96)

    #To process all the files in your current folder, enter the information below:
    
    
    
#Step 3: Run the code
#Press control a to select all the code, and then press control enter to run.
#You can find the exported excel sheet and heat map in your current directory!

interact1 <- function() {
    repeat {
        answer <- readline(cat("Welcome to Heat Map Generator!\nPlease navigate to your desired working directory.\nType D when done, or type H if you need help."))
        if (answer == "H") {
            cat("Click on the Files tab in the lower right window of Rstudio.\nNavigate to the folder where your file(s) are contained.\nThen press More and then Set As Working Directory.")
        } else if (answer == "D") {
            print("Great! Moving on.")
            break
        } else {
            print("Error: Your answer must be D or H.")
        }
    }
}

interact2 <- function() {
    repeat {
        answer <- readline(cat("How many total wells are in this plate?\nYour answer must be either 384 or 96."))
        if (answer == "384") {
            print("Got it! You are processing a 384-well plate.")
            break
        } else if (answer == "96") {
            print("Got it! You are processing a 96-well plate.")
            break
        } else {
            print("Error: Your answer must be 394 or 96.")
        }
    }
}

interact3 <- function() {
    repeat {
        answer <- readline(cat("How many files would you like to process?\nType One to process a single file.\nType Many to process all the files in your working directory."))
        if (answer == "One") {
            print("Got it! You are processing one data file.")
            break
        } else if (answer == "Many") {
            print("Got it! You are processing many data files.")
            break
        } else {
            print("Error: Your answer must be One or Many.")
        }
    }
}


#Load required packages
for (package in c("openxlsx", "stringr", "gplots", "grDevices", "tools")) {
    if (!require(package, character.only = T)) {
        install.packages(package)
        library(package, character.only = T)
    } else {
        library(package, character.only = T)
    }
}

#Load file
loadfile <- function(filename) {
    data <- openxlsx::read.xlsx(filename, startRow = 2)
    if ("blank" %in% data$Data.File) {
        data <- data[-grep("blank", data$Data.File), ]
    }
    df <- data.frame("Data.File" = data$Data.File, "Ratio" = data[,9] / data[,7])
    return(df)
}

#Create grid
makegrid <- function(data, platesize = NULL, export = TRUE) {
    df$columns <- as.numeric(stringr::str_extract(data$Data.File, "(?<=Column)\\d+"))
    df$rows <- as.numeric(stringr::str_extract(data$Data.File, "(?<=-r)\\d+(?=.d)"))
    if (!is.null(platesize)) {
        if (platesize == 384) {
            grid <- matrix(ncol = 12, nrow = 8)
        } else if (platesize == 96) {
            grid <- matrix(ncol = 12, nrow = 8)
        } else {
            print("Error! Plate size must be 96 or 384.")
        }
    } else if (max(df$rows) > 8) {
        grid <- matrix(ncol = 24, nrow = 16)
    } else if (max(df$columns) <= 12 & max(df$rows) < 8) {
        grid <- matrix(ncol = 12, nrow = 8)
    } else {
        print("Error! Cannot compute plate size.")
    }
    for (i in 1:nrow(df)) {
        grid[df[i, 4], df[i, 3]] <- df[i, 2]
    }
    if (export == TRUE) {
        file <- paste0(tools::file_path_sans_ext(filename), " Plate grid.xlsx")
        openxlsx::write.xlsx(grid, file = file, keepNA = TRUE, colNames = FALSE)
    } else {
        return(grid)
    }
}

#Create heat map
heat <- function(grid, export = TRUE) {
    map <- gplots::heatmap.2(grid, na.rm = TRUE, dendrogram = "none", Rowv = FALSE, Colv = FALSE,
                             trace = "none", density.info = "none",
                             col = colorRampPalette(c("yellow", "orange", "red")))
    if (export == TRUE) {
        file <- paste0(tools::file_path_sans_ext(filename), " Heat map.pdf")
        pdf(file = file)
        invisible(map)
        dev.off()
    } else {
        invisible(map)
    }
}

#Run the code
heat(makegrid(loadfile(filename), platesize = plate))