#Step 1: Download required package
#Note: You only need to do this the first time!
install.packages("openxlsx")

#Step 2: Enter your variables
filename <- "Plate map.xlsx"  #Name of your plate map Excel file
plate <- 384  #Number of wells in your plate (must be 384 or 96)
folder <- "20181128"  #Name of the folder where your MS data will be stored, usually today's date
outputfile <- "Linearity test plate map.xlsx"  #Name for the file that the program will output

#Step 3: Navigate to the directory containing your data file
#Find it in the files tab, then click More and Set As Working Directory

#Step 4: Run the code
#Press control a to select all the code, and then press control enter to run.
#You can find the exported excel sheet and heat map in your current directory!


#Load required package
require(openxlsx)

#Open file
map <- openxlsx::read.xlsx(filename, colNames = FALSE, skipEmptyCols = FALSE, skipEmptyRows = FALSE)

#Create MS file function
MakeMSfile <- function(platemap, foldername, platesize, path = "D:\\MassHunter\\Data\\Dingyin\\") {  #path must use double slashes!
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

#Convert plate map to MS file
MSfile <- MakeMSfile(map, foldername = folder, platesize = plate)

#Output to new excel file
openxlsx::write.xlsx(MSfile, file = outputfile, colNames = FALSE, rowNames = FALSE)
