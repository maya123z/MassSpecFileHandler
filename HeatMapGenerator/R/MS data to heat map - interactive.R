###Core Functions###

#Load required packages
loadpackagesMS <- function() {
    for (package in c("openxlsx", "stringr", "gplots", "grDevices", "tools")) {
        if (!require(package, character.only = T)) {
            install.packages(package)
            library(package, character.only = T)
        } else {
            library(package, character.only = T)
        }
    }
}

#Load file
loadfileMS <- function(filename) {
    data <- openxlsx::read.xlsx(filename, startRow = 2)
    if ("blank" %in% data$Data.File) {
        data <- data[-grep("blank", data$Data.File), ]
    }
    frame <- data.frame("Data.File" = data$Data.File, "Ratio" = data[,9] / data[,7])
    return(frame)
}

#Create grid
makegrid <- function(data, platesize = NULL, export = TRUE) {
    data$columns <- as.numeric(stringr::str_extract(data$Data.File, "(?<=Column)\\d+"))
    data$rows <- as.numeric(stringr::str_extract(data$Data.File, "(?<=-r)\\d+(?=.d)"))
    if (!is.null(platesize)) {
        if (platesize == 384) {
            grid <- matrix(ncol = 24, nrow = 16)
        } else if (platesize == 96) {
            grid <- matrix(ncol = 12, nrow = 8)
        } else {
            print("Error! Plate size must be 96 or 384.")
        }
    } else if (max(data$rows) > 8) {
        grid <- matrix(ncol = 24, nrow = 16)
    } else if (max(data$columns) <= 12 & max(data$rows) < 8) {
        grid <- matrix(ncol = 12, nrow = 8)
    } else {
        print("Error! Cannot compute plate size.")
    }
    for (i in 1:nrow(data)) {
        grid[data[i, 4], data[i, 3]] <- data[i, 2]
    }
    if (export == TRUE) {
        file <- paste0(tools::file_path_sans_ext(filename), " Plate grid.xlsx")
        openxlsx::write.xlsx(grid, file = file, keepNA = TRUE, colNames = FALSE)
    }
    return(grid)
}

#Create heat map
heat <- function(grid, export = TRUE) {
    map <- gplots::heatmap.2(grid, na.rm = TRUE, dendrogram = "none", Rowv = FALSE, Colv = FALSE,
                             trace = "none", density.info = "none",
                             col = colorRampPalette(c("yellow", "orange", "red")))
    if (isTRUE(export)) {
        file <- paste0(tools::file_path_sans_ext(filename), " Heat map.pdf")
        grDevices::pdf(file = file)
        map
        grDevices::dev.off()
    } else {
        invisible(map)
    }
}


###Interactive Functions###

#Working directory prompt
interact1 <- function() {
    repeat {
        answer <- readline(cat("\nWelcome to Heat Map Generator!\nPlease navigate to your desired working directory.\nType D when done, or type H if you need help.\nYou can also type Exit at any time to leave this program."))
        if (answer == "H") {
            cat("Click on the Files tab in the lower right window of Rstudio.\nNavigate to the folder where your file(s) are contained.\nThen press More and then Set As Working Directory.")
        } else if (answer == "D") {
            print("Great! Moving on.")
            return("Good")
            break
        } else if (answer == "Exit") {
            print("Goodbye!")
            return("Bad")
            break
        } else {
            print("Error: Your answer must be D or H.")
        }
    }
}

#Plate size prompt
interact2 <- function() {
    repeat {
        answer <- readline(cat("\nHow many total wells are in this plate?\nYour answer must be either 384 or 96."))
        if (answer == "384") {
            print("Got it! You are processing a 384-well plate.")
            return(384)
            break
        } else if (answer == "96") {
            print("Got it! You are processing a 96-well plate.")
            return(96)
            break
        } else if (answer == "Exit") {
            print("Goodbye!")
            return("Bad")
            break
        } else {
            cat("Error: Your answer must be 384 or 96.\nYou may also type Exit to leave the program.")
        }
    }
}

#Iterate prompt
interact3 <- function() {
    repeat {
        answer <- readline(cat("\nHow many files would you like to process?\nType One to process a single file.\nType Many to process all the files in your working directory."))
        if (answer == "One") {
            print("Got it! You are processing one data file.")
            return(FALSE)
            break
        } else if (answer == "Many") {
            print("Got it! You are processing many data files.")
            return(TRUE)
            break
        } else if (answer == "Exit") {
            print("Goodbye!")
            return("Bad")
            break
        } else {
            cat("Error: Your answer must be One or Many.\nYou may also type Exit to leave the program.")
        }
    }
}

#File name prompt
interact4 <- function() {
    repeat {
        answer <- readline(cat("\nWhat is the name of the file you'd like to process?\nDon't forget to include the .xlsx file extension!"))
        if (isTRUE(grepl(".xlsx", answer))) {
            allfiles <- list.files(".", ".xlsx")
            if (answer %in% allfiles) {
                return(answer)
                break
            } else {
                print("Error: File does not exist in your current directory.")
            }
        } else if (answer == "Exit") {
            print("Goodbye!")
            return("Bad")
            break
        } else {
            cat("Error: Must include the .xlsx file extension.\nYou may also type Exit to leave the program.")
        }
    }
}

#Exit
testfunc <- function(x) {
    if (x == "Bad") {
        stop()
    }
}

#Master function
heatmapgenerator <- function() {
    test1 <- interact1()
    testfunc(test1)
    plate <- interact2()
    testfunc(plate)
    iterate <- interact3()
    testfunc(iterate)
    if (isFALSE(iterate)) {
        filename <<- interact4()
        testfunc(filename)
    }
    cat("\n Loading required packages...\n")
    loadpackages()
    cat("\n Processing your file(s)...\n")
    if (isFALSE(iterate)) {
        heat(makegrid(loadfile(filename), platesize = plate))
    } else if (isTRUE(iterate)) {
        allfiles <- list.files(".", ".xlsx")
        for (file in allfiles) {
            heat(makegrid(loadfile(file), platesize = plate))
        }
    }
    rm(filename, envir = .GlobalEnv)
    print("Finished!")
}