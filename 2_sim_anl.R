#install.packages('readxl')
#install.packages('stringr')
#install.packages('readtext')

library('readxl')
library('stringr')
library('readtext')


# path_1 is where the main files are saved

path_1 <- "C:\\Users\\Mahmood\\Desktop\\alizad_1995\\military\\sample\\files\\test"

files <- list.files(path_1)

path_2 <- "C:\\Users\\Mahmood\\Desktop\\alizad_1995\\military\\input\\keywords.xlsx"

# path_2 is where the keywords file is saved 

keywords <- read_excel(path_2)

nword <- dim(keywords)[1]

fin_res <- matrix(NA, nword, length(files))
colnames(fin_res) <- str_remove(files, ".docx")

res <- matrix(NA, nword, 2)


for(i in 1:length(files)) {
  
  
  filename <- files[i]
  
  cntnt <- unlist(readtext(paste0(path_1, "\\", filename)))[2]
  
  ## removing the TOC (if applicable)
  
  if(length(unlist(str_locate_all(cntnt, "PAGEREF")))!=0) {
    
    pt_1 <- unlist(str_locate_all(cntnt, "PAGEREF"))[1] - 2
    pt_2 <- unlist(str_locate_all(cntnt, "PAGEREF"))[length(unlist(str_locate_all(cntnt, "PAGEREF")))] + 20
    
    part_1 <- str_sub(cntnt, 1, pt_1-1)
    part_2 <- str_sub(cntnt, pt_2+1, nchar(cntnt))
    
    cntnt <- str_c(part_1, part_2, sep="\n")
    
  }
  
  
  # processing which farsi words are there in the document
  
  keywords_fa <- paste0(" ", unlist(keywords[, 2]))
  
  for(j in 1:length(keywords_fa)) {
    res[j, 1] <- str_detect(cntnt, keywords_fa[j])
  }
  
  
  ##
  
  # processing which english words are there in the document
  
  keywords_en <- paste0(" ", unlist(keywords[, 1]))
  
  for(j in 1:length(keywords_en)) {
    res[j, 2] <- str_detect(cntnt, keywords_en[j])
  }
  
  
  ##
  
  
  # finding out whether a word or its translation exist in a document or not.
  fin_res[, i] <- (res[, 1] + res[, 2]) != 0
  
  
  
}
