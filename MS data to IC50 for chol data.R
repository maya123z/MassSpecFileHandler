### MS Data to IC50 ###

#Step 1: Enter your variables
datafile <- "Chol06212019D CT 4.xlsx"    #The name of your data file. Must be formatted in platemap style!
plate <- 384   #The number of wells in your plate(s)
numplates <- 1   #The number of separate plates in your data file.
platemap <- "PM RT4.xlsx"   #Name of your compound platemap
numplatemaps <- 1   #If you have more than one plate map, input that here.
sheetname <- "Sheet1"  #If you have only one plate map, give the sheet name

#Step 2: Navigate to the directory containing your data file
#Find it in the files tab, then click More and Set As Working Directory

#Step 3: Run the code
#Press control a to select all the code, and then press control enter to run.
#You can find the exported files in your current directory!
#Note: Don't worry if the code prints a few error messages, this is normal!


#####################################################################################################


#Load required packages
check.packages <- function(pkg){
    new.pkg <- pkg[!(pkg %in% installed.packages()[, "Package"])]
    if (length(new.pkg)) 
        install.packages(new.pkg, dependencies = TRUE)
    sapply(pkg, require, character.only = TRUE)
}
packages <- c("openxlsx", "stringr", "drc", "grDevices", "stats")
check.packages(packages)

#Workbook to contain IC50 data sheets
wb <- createWorkbook("IC50 Values.xlsx")
mylist <- list()

#Define functions
load <- function(data, sheet) {
    openxlsx::read.xlsx(data, sheet=sheet, colNames=FALSE, skipEmptyCols=FALSE, skipEmptyRows=FALSE)
}

fill <- function(x) {   #Fills empty wells of data frame
    if (plate==384 & ncol(x) < 24) {
        x[,(ncol(x)+1):24] <- NA
    } else if (plate==96 & ncol(x) < 12) {
        x[,(ncol(x)+1):12] <- NA
    }
    if (plate==384 & nrow(x) < 16) {
        x[,(nrow(x)+1):16] <- NA
    } else if (plate==96 & nrow(x) < 8) {
        x[,(nrow(x)+1):8] <- NA
    }
    return(x)
}

getic50 <- function(x) {
    tryCatch(drc::drm(Value ~ Conc, data = x, na.action=na.omit,
                      fct = LL.4(names = c("Slope", "Lower Limit", "Upper Limit", "ED50"))),
             error=function(e){cat("Warning :",conditionMessage(e), "\n")})
}

plotic50 <- function(x, name) {
    op <- par(mfrow = c(1, 2), mar=c(4,4,1.5,1), mgp=c(2.5,.7,0))
    plot(x, broken=TRUE, bty="l", main=name,
         xlab="Concentration of Compound (uM)", ylab="Ratio Chol/IS")
}

listic50 <- function(x, y, z) {
    samples <- names(y)
    ic <- sapply(y, function(q) q$coefficients[4])
    oefficacy <- vector()
    for (i in 1:length(z)) {
        oe <- avg[[i]][1] - tail(avg[[i]], n=1)
        oefficacy <- c(oefficacy, oe)
    }
    pefficacy <- sapply(y, function(q) q$coefficients[3] - q$coefficients[2])
    ic50 <- data.frame(samples, as.numeric(ic), as.numeric(oefficacy), as.numeric(pefficacy),
                       stringsAsFactors=FALSE)
    rownames(ic50) <- NULL
    colnames(ic50) <- c("Sample", "IC50 (uM)", "Observed Efficacy", "Predicted Efficacy")
    for (i in 1:length(names(x))) {
        if (!(names(x)[i] %in% names(y))) {
            add <- c(names(x)[i], NA, NA, NA)
            ic50 <- rbind(ic50, add)
        }
    }
    ic50 <- ic50[order(ic50$Sample),]
    return(ic50)
}

getsdv <- function(x) {
    sdv <- list()
    for (i in 1:length(x)) {
        sp <- split(x[[i]], x[[i]]$Conc)
        z <- sapply(sp, function(y) stats::sd(y$Value))
        sdv[[i]] <- z
        names(sdv)[i] <- x[[i]][1,1]
    }
    return(sdv)
}

getavg <- function(x) {
    avg <- list()
    for (i in 1:length(x)) {
        sp <- split(x[[i]], x[[i]]$Conc)
        z <- sapply(sp, function(y) mean(y$Value))
        avg[[i]] <- z
        names(avg)[i] <- x[[i]][1,1]
    }
    return(avg)
}

#Process data
for (p in 1:numplates) {
    #Load files
    data <- fill(load(datafile, p))
    #data <- data[20:35,]
    
    if (numplatemaps == 1) {
        pm <- fill(load(platemap, sheetname))
    } else {
        pm <- fill(load(platemap, sheetname[p]))
    }
    
    #Pull out control wells
    max <- vector()
    for (cols in 1:ncol(data)) {
        for (i in 1:nrow(data)) {
            if ((!is.na(data[i, cols]) & !is.na(pm[i, cols])) & (pm[i, cols]=="MAX" | pm[i, cols]=="max")) {
                max <- c(max, data[i, cols])
            }
        }
    }
    
    
    #Normalize data
    min.avg <- 0
    max.avg <- mean(max) - min.avg
    data.norm <- sapply(data, function(x) (x - min.avg)/max.avg*100)
    
    #Organize into one dataframe
    df <- data.frame(Sample=character(), Conc=character(), Row=character(), Col=character())
    for (cols in 1:ncol(pm)) {
        for (i in 1:nrow(pm)) {
            if (!is.na(pm[i, cols]) & grepl(";", pm[i, cols])) {
                split <- stringr::str_split(pm[i, cols], ";")
                info <- c(split[[1]][1], split[[1]][2], i, cols)
                df <- rbind(df, info, stringsAsFactors=FALSE)
            }
        }
    }
    colnames(df) <- c("Sample", "Conc", "Row", "Col")
    df$Conc <- as.numeric(df$Conc)
    df$Row <- as.integer(df$Row)
    df$Col <- as.integer(df$Col)
    df$Value <- NA
    for (i in 1:nrow(df)) {
        df$Value[i] <- data.norm[df$Row[i], df$Col[i]]
    }
    if (numplatemaps > 1) {
        #Plot IC50 curves
        folder <- paste("Plate", p, "IC50 Curves/", sep=" ")
        dir.create(folder)
        
        ls <- split(df, df$Sample)
        ls2 <- lapply(ls, getic50)
        ls2 <- Filter(Negate(is.null), ls2)
        ls3 <- ls[c(which(names(ls) %in% names (ls2)))]
        
        sdv <- getsdv(ls3)
        avg <- getavg(ls3)
        
        for (i in 1:length(ls2)) {
            grDevices::jpeg(filename=paste0(getwd(), "/", folder, names(ls2)[i], ".jpg"), width=1200,
                            height=800)
            plotic50(ls2[[i]], names(ls2)[i])
            grDevices::dev.off()
        }
        
        #List IC50 Values
        ic50 <- listic50(ls, ls2, ls3)
        
        openxlsx::addWorksheet(wb, sheetName = paste0("Sheet", p))
        openxlsx::writeData(wb, paste0("Sheet", p), ic50, keepNA = TRUE)
        saveWorkbook(wb, "IC50 Values.xlsx", overwrite = TRUE)
    } else {
        mylist[[p]] <- df
    }
}

if (numplatemaps == 1) {
    reps <- do.call("rbind", mylist)
    ls <- split(reps, reps$Sample)
    
    #Plot IC50 curves
    folder <- "IC50 Curves/"
    dir.create(folder)
    
    ls2 <- lapply(ls, getic50)
    ls2 <- Filter(Negate(is.null), ls2)
    ls3 <- ls[c(which(names(ls) %in% names (ls2)))]
    
    sdv <- getsdv(ls3)
    avg <- getavg(ls3)
    
    for (i in 1:length(ls2)) {
        grDevices::jpeg(filename=paste0(getwd(), "/", folder, names(ls2)[i], ".jpg"), width=1200,
                        height=800)
        plotic50(ls2[[i]], names(ls2)[i])
        arrows(x0=sort(unique(ls3[[i]]$Conc)), y0=avg[[i]]-sdv[[i]], y1=avg[[i]]+sdv[[i]], length=0.05, angle=90,
               code=3)
        grDevices::dev.off()
    }
    
    #List IC50 Values
    ic50 <- listic50(ls, ls2, ls3)
    
    openxlsx::addWorksheet(wb, sheetName = paste0("Sheet", p))
    openxlsx::writeData(wb, paste0("Sheet", p), ic50, keepNA = TRUE)
    saveWorkbook(wb, "IC50 Values.xlsx", overwrite = TRUE)
}
