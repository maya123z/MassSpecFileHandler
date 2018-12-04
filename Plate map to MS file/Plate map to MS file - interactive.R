
#Step 2: Enter your variables
filename <- "your plate map file.xlsx"  #Name of your plate map Excel file
plate <- 384  #Number of wells in your plate (must be 384 or 96)
folder <- "your folder name"  #Name of the folder where your MS data will be stored, usually today's date
outputfile <- "your output file name.xlsx"  #Name for the file that the program will output

interact1 <- function() {
    repeat {
        answer <- readline(cat("\nWelcome to MS File Generator!\nPlease navigate to your desired working directory.\nType D when done, or type H if you need help.\nYou can also type Exit at any time to leave this program."))
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

interact3 <- function() {
    repeat {
        answer <- readline(cat("\nWhat is the name of the plate map file you'd like to process?\nDon't forget to include the .xlsx file extension!"))
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
 

#Load required package
loadpackagesmap <- function() {
    for (package in c("openxlsx")) {
        if (!require(package, character.only = T)) {
            install.packages(package)
            library(package, character.only = T)
        } else {
            library(package, character.only = T)
        }
    }
}

#Load file
loadfilemap <- function(filename) {
    openxlsx::read.xlsx(filename, colNames = FALSE, skipEmptyCols = FALSE, skipEmptyRows = FALSE)
    }

#Create MS file function
MakeMSfile <- function(platemap, foldername, platesize, path = "D:\\MassHunter\\Data\\Dingyin\\") {
    if (platesize == 384) {
        #Make empty MS file
        MSfile <- as.data.frame(matrix(ncol=2, nrow=384))
        platecols <- c(rep(1:24, each=16))
        MSfile[,1] <- paste0("P2-", LETTERS[1:16], platecols)
        #Add plate map data
        unlisted <- unlist(platemap, use.names = FALSE)
        unlisted <- c(unlisted, rep(NA, 384 - length(unlisted)))
        MSfile[,2] <- paste0(path, foldername, "\\", unlisted, "-Column", rep(1:24, each = 16), "-r", 1:16, ".d")
        MSfile <- MSfile[!is.na(unlisted), ]
        return(MSfile)
    } else if (platesize == 96) {
        #Make empty MS file
        MSfile <- as.data.frame(matrix(ncol=2, nrow=96))
        platecols <- c(rep(1:12, each=8))
        MSfile[,1] <- paste0("P2-", LETTERS[1:8], platecols)
        #Add plate map data
        unlisted <- unlist(platemap, use.names = FALSE)
        unlisted <- c(unlisted, rep(NA, 96 - length(unlisted)))
        MSfile[,2] <- paste0(path, foldername, "\\", unlisted, "-Column", rep(1:12, each = 8), "-r", 1:8, ".d")
        MSfile <- MSfile[!is.na(unlisted), ]
        return(MSfile)
    } else {
        print("Error! Platesize must be 96 or 384.")
    }
}

testfunc <- function(x) {
    if (x == "Bad") {
        stop()
    }
}

#Master function
MSfile <- function(x) {
    test1 <- interact1()
    testfunc(test1)
    plate <- interact2()
    testfunc(plate)
    filename <- interact3()
    testfunc(filename)
    cat("\n Loading required packages...\n")
    loadpackagesmap()
    cat("\n Processing your file(s)...\n")
    MakeMSfile(map, foldername = folder, platesize = plate)
    openxlsx::write.xlsx(MSfile, file = outputfile, colNames = FALSE, rowNames = FALSE)
}
