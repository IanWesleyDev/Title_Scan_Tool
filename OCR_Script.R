require(tidyverse)
require(tesseract)
require(pdftools)
require(magick)
require(tictoc)
require(RDCOMClient)

# Header ------------------------
# This script takes jpegs of washington state vehicle titles, scanned in greyscale at 300 dpi resoloution and uses
# OCR to extract the VIN and registered owner information from the titles.
###

# Setup -------------------------

WorkingDirectory <- "G:/Title_Scan_Tool"
EmailAddress        <- "wesleyi@wsdot.wa.gov;BinkleM@wsdot.wa.gov"
TechnicalMainancePersonName <- "Ian"

# Email functions ---------------

GenerateEmailError <- function(MyErrMsg, err, myTechnicalMainancePersonName = TechnicalMainancePersonName){
  myHTML <- paste0("<html><h2>There was an error with your title scan</h2>",
                   "<p> There is no need to do anything. ", myTechnicalMainancePersonName, " has been notified of the error and will work to correct it.",
                   "<br><br> Error Message: <br>", MyErrMsg, "<br>", err)
  myHTML
}

SendEmailReport <- function(myHTML, isErr = FALSE, myEmailAddress = EmailAddress) {
  subject <- "Your title scan is completed"
  
  if (isErr) {
    subject <- "There was an error in your title scan"
  }
  
  OutApp <- COMCreate("Outlook.Application")
  outMail = OutApp$CreateItem(0)
  outMail[["To"]] = myEmailAddress
  outMail[["subject"]] = subject
  outMail[["HTMLbody"]] =  myHTML                
  outMail$Send()  
}

#Set working directory ---------
tryCatch(setwd(WorkingDirectory), error= function(cond) {
  errMsg <- "Could not find working directory."
  myHTML <- GenerateEmailError(errMsg, err)
  SendEmailReport(myHTML, isErr = TRUE)
  stop()
})


# VIN Check --------------------
# Function to take a vin and confirm its check digit
####
vin_check <- function(vin) {
  
  if(nchar(vin) != 17){
    return("Invalid VIN length")
  }
  
  weights <- c(
    8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2
  )
  
  conversions <- list(
    A = 1,
    B = 2,
    C = 3,
    D = 4,
    E = 5,
    F = 6,
    G = 7,
    H = 8,
    I = '-',
    J = 1,
    K = 2,
    L = 3,
    M = 4,
    N = 5,
    O = '-',
    P = 7,
    Q = '-',
    R = 9,
    S = 2,
    T = 3,
    U = 4,
    V = 5,
    W = 6,
    X = 7,
    Y = 8,
    Z = 9
  )
  
  vin_split   <- strsplit(vin, "")
  
  check_digit <- strsplit(vin, "") %>% unlist()
  check_digit <- check_digit[9] %>% as.character() %>% toupper() 
  
  vin_replace <- vin
  for(i in 1:length(conversions)){
    vin_replace <- gsub(names(conversions)[i], conversions[i], vin_replace)
  }
  
  vin_split_number <- strsplit(vin_replace, "") %>% unlist() %>% as.integer()
  
  
  calc_check_digit <- sum(vin_split_number * weights) %% 11
  print(calc_check_digit) 
  
  if(is.na(calc_check_digit)){
    return("Invalid VIN")
  }
  
  if(calc_check_digit == 10){
    calc_check_digit <- "X"
  }
  
  calc_check_digit <- calc_check_digit %>% as.character() %>% toupper()
  
  if(calc_check_digit == check_digit){
    return("Valid VIN")
  } else {
    return("Invalid VIN")
  }
  
}

# Image preprocess -------------
image_size_crop_preprocess <- function(img_path){
  
  image_scale_factor <- 1
  
  resize_string <- paste0(3000 * image_scale_factor, "x")
  crop_string   <- paste0(2620 * image_scale_factor, "x", 1975 * image_scale_factor,"+", 195 * image_scale_factor, "+", 195 * image_scale_factor)
  
  tic("image loading and croping")  
  img <- image_read(img_path)%>%
    image_resize(resize_string) %>%
    image_crop(crop_string) %>% 
    image_draw()
  toc()
  tic("drawing rectagles")
  rect(0, 0, 750 * image_scale_factor, 255 * image_scale_factor, col='white', border = NA)
  rect(xleft=0, ybottom=1690 * image_scale_factor, xright=2700 * image_scale_factor, ytop=975 * image_scale_factor, col='white', border = NA)
  rect(xleft=0, ybottom=2115 * image_scale_factor, xright=1280 * image_scale_factor, ytop=975 * image_scale_factor, col='white', border = NA)
  toc()
  tic("image capture, flatten, and preprocess")
  img <- image_capture() %>% 
    image_flatten() %>% 
    image_contrast(1)
  toc()
  tic("image save")
  
  path_split <- strsplit(img_path, "/") %>% unlist()
  name_split <- strsplit(path_split[4], "_") %>% unlist()
  final_name <- paste0(path_split[3], "_", name_split[2])
  
  image_write(img, 
              path=paste0('./Code/images_processed/', gsub(".jpg", "", final_name), ".png"),
              format = 'png', density = '300x300')
  toc()
  dev.off()
  img <- NA
  gc()
}

# OCR and text extraction -----

ocr_image <- function(FilePath){
  tic("image read")
  input <- image_read(FilePath)
  toc()
  tic("OCR image")
  
  text <-input %>% 
    ocr()
  toc()
  img <- NA
  gc()
  return(text)
}

ocr_text_check <- function(ocr_text, processed_files){
  
  lits_ocr_text <- ocr_text %>% unlist()
  
  ocr_text_checked <- data.frame("file_path"      = processed_files,
                                 "ocr_text"       = ocr_text %>% unlist(), 
                                 "can_find_vin"   = grepl("Vehicle Identification Number", ocr_text) | grepl("VIN", ocr_text),
                                 "can_find_owner" = grepl("Registered Owner", ocr_text)
  )
  
  return(ocr_text_checked)
}

extract_vin <- function(ocr_text_check){
  
  
  get_vin <- function(my_ocr_text) {
    # browser()
    my_ocr_text <- my_ocr_text %>% as.character()
    ocr_lines <- strsplit (my_ocr_text, "\n") %>% unlist()
    
    target_line_number <- grep("Vehicle Identification Number", ocr_lines) + 1
    
    if(length(target_line_number) == 0){
      target_line_number <- grep("VIN", ocr_lines) + 1
    } 
    
    target_line <- strsplit (ocr_lines[target_line_number], " ") %>% unlist()
    # browser()
    
    vin <- target_line[1] %>% toupper()
    if(nchar(target_line[1]) != 17){
      # browser()
      search <- TRUE
      found  <- FALSE
      target_text <- paste0(target_line, collapse = "")
      my_nchar <- 1
      while (search) {
        t_vin <- substr(target_text, start = my_nchar, stop = my_nchar + 16)
        t_vin <- t_vin %>% toupper()
        t_vin <- gsub("O", 0, t_vin)
        t_vin <- gsub("I", 1, t_vin)
        t_vin <- gsub("Q", 0, t_vin)
        
        if(vin_check(t_vin) == "Valid VIN") {
          vin <- t_vin
          search <- FALSE
          found  <- TRUE
        }  
        
        my_nchar <- my_nchar + 1
        if(my_nchar >= nchar(target_text) - 17){
          search <- FALSE
        }
      }
    }
    
    vin <- vin %>% toupper()
    vin <- gsub("O0", 0, vin)
    vin <- gsub("0O", 0, vin)
    vin <- gsub("O", 0, vin)
    vin <- gsub("I", 1, vin)
    vin <- gsub("Q", 0, vin)
    
    return(vin)
    
  }
  
  output <- ocr_text_checked
  
  output$VIN <- "Cannot find VIN"
  output$valid_vin <- "Cannot find VIN"
  
  output$VIN[output$can_find_vin == TRUE] <- lapply(output$ocr_text[output$can_find_vin == TRUE], get_vin)
  
  output$valid_vin[output$can_find_vin == TRUE] <- lapply(output$VIN[output$can_find_vin == TRUE],  vin_check) 
  return(output)
}

extract_registered_owner <- function(ocr_data_frame) {
  
  get_owner <- function(my_ocr_text) {
    my_ocr_text <- my_ocr_text %>% as.character()
    ocr_lines <- strsplit (my_ocr_text, "\n") %>% unlist()
    target_line_number <- grep("Registered Owner", ocr_lines) + 1
    
    if(nchar(ocr_lines[target_line_number]) < 1){
      output <- ocr_lines[target_line_number + 1]
    } else {
      output <- ocr_lines[target_line_number]
    }
    
    return(output)
  } 
  output <- ocr_data_frame
  
  output$registered_owner <- "Not Found"
  output$registered_owner[output$can_find_owner] <- lapply(output$ocr_text[output$can_find_owner == TRUE], get_owner)
  
  return(output)
}

# Run -------------------
# Step 1 - Size, Crop, and Preprocess Images

raw_files  <- list.files('./Titles - Input', full.names = TRUE, recursive = TRUE)

if(length(raw_files) < 1){
  errMsg <- "No scanned Files found."
  myHTML <- GenerateEmailError(errMsg, err)
  SendEmailReport(myHTML, isErr = TRUE)
  stop()
}

tryCatch(mapply(image_size_crop_preprocess, raw_files), error= function(cond) {
  errMsg <- "Error preprocessing images"
  myHTML <- GenerateEmailError(errMsg, err)
  SendEmailReport(myHTML, isErr = TRUE)
  stop()
})

# Step 2 - OCR Images
processed_files <- list.files('./Code/images_processed', full.names = TRUE) 
processed_files <- processed_files[grepl("png", processed_files)]

if(length(processed_files) < 1){
  errMsg <- "No prpcessed files found"
  myHTML <- GenerateEmailError(errMsg, err)
  SendEmailReport(myHTML, isErr = TRUE)
  stop()
}

tryCatch(ocr_text <- lapply(processed_files, ocr_image), error= function(cond) {
  errMsg <- "Error running OCR"
  myHTML <- GenerateEmailError(errMsg, err)
  SendEmailReport(myHTML, isErr = TRUE)
  stop()
})


# Step 4 - Extract Text 

tryCatch(ocr_text <- lapply(processed_files, ocr_image), error= function(cond) {
  errMsg <- "Error running OCR"
  myHTML <- GenerateEmailError(errMsg, err)
  SendEmailReport(myHTML, isErr = TRUE)
  stop()
})

tryCatch(ocr_text_checked <- ocr_text_check(ocr_text, processed_files), error= function(cond) {
  errMsg <- "Error checking text"
  myHTML <- GenerateEmailError(errMsg, err)
  SendEmailReport(myHTML, isErr = TRUE)
  stop()
})


tryCatch(final_vin <- ocr_text_checked %>% extract_vin(), error= function(cond) {
  errMsg <- "Error extracting VIN"
  myHTML <- GenerateEmailError(errMsg, err)
  SendEmailReport(myHTML, isErr = TRUE)
  stop()
})

tryCatch(final_data_frame <- final_vin %>% extract_registered_owner(), error= function(cond) {
  errMsg <- "Error extracting registered owner"
  myHTML <- GenerateEmailError(errMsg, err)
  SendEmailReport(myHTML, isErr = TRUE)
  stop()
})


final_data_frame$valid_vin <- final_data_frame$valid_vin %>% unlist()
final_data_frame$VIN <- final_data_frame$VIN %>% unlist()
final_data_frame$registered_owner <- final_data_frame$registered_owner %>% unlist()

final_data_frame <- final_data_frame %>% arrange(valid_vin)

write.csv(final_data_frame, sprintf("./Excel - Output/Scanned Titles %s.csv", Sys.Date()))

unlink(processed_files)


success_html <-  paste0("<html><h2>Your title scan is complete!</h2>",
                        "<p> You may view the excel out put here: <a href='", 
                        sprintf("%s/Excel - Output/Scanned Titles %s.csv", getwd(), Sys.Date()), "'>",
                        sprintf("%s/Excel - Output/Scanned Titles %s.csv", getwd(), Sys.Date())
                        , "</a>")

SendEmailReport(success_html)