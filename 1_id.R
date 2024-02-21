#install.packages('rebus)
#install.packages('readxl')
#install.packages('stringr')
#install.packages('readtext')

library('rebus')
library('readxl')
library('stringr')
library('readtext')


## importing the main file

path_1 <- "C:\\Users\\Mahmood\\Desktop\\alizad_1995\\military\\sample\\files\\"
filename <- "تاثیر چند کشتی همزمان بر عملکرد دانه ارقام گلرنگ در منطقه آران و بیدگل.docx"

cntnt <- unlist(readtext(paste0(path_1, filename)))[2]

## removing the TOC (if applicable)

if(length(unlist(str_locate_all(cntnt, "PAGEREF")))!=0) {
  
  pt_1 <- unlist(str_locate_all(cntnt, "PAGEREF"))[1] - 2
  pt_2 <- unlist(str_locate_all(cntnt, "PAGEREF"))[length(unlist(str_locate_all(cntnt, "PAGEREF")))] + 20
  
  part_1 <- str_sub(cntnt, 1, pt_1-1)
  part_2 <- str_sub(cntnt, pt_2+1, nchar(cntnt))
  
  cntnt <- str_c(part_1, part_2, sep="\n")
  
}

# # # # #


NA_trim <- function(char) {
  
  ## returns the trimmed character, i.e., removes the ending NAs
  
  res <- vector("character", 0)
  
  if(length(char)!=0) {
    for(i in 1:length(char)) {
      if(!is.na(char[i])) {
        res <- c(res, char[i])
      } 
    }
  }
  
  return(res)
  
}


# # # # # #


integ <- function(chr_lst = chr_lst) {
  
  ## receives a character set, replaces:
  ##"ي" and "ئ" with "ی"
  ##"ك" with "ک" 
  ## "ؤ" with "و"
  ##"¬" with " " 
  ##and returns the integrated character set
  
  int <- vector('character', length(chr_lst))
  
  for(i in 1:length(chr_lst)) {
    str <- str_replace_all(chr_lst[i], "ي", "ی")
    str <- str_replace_all(str, "ئ", "ی")
    str <- str_replace_all(str, "ك", "ک")
    str <- str_replace_all(str, "ؤ", "و")
    str <- str_replace_all(str, "¬", " ")
    int[i] <- str
  }
  
  return(int)
  
}



# # # # #


min_len <- function(chrlst, minchar) {

  if(length(chrlst)!=0) {
    for(i in 1:length(chrlst)) {
      if(nchar(chrlst[i]) < minchar) {
        chrlst[i] <- NA
      }
    }
  }
  
  return(chrlst)
  
}


# # # # #



max_len <- function(chrlst, maxchar) {

  if(length(chrlst)!=0) {
    for(i in 1:length(chrlst)) {
      if(nchar(chrlst[i]) > maxchar) {
        chrlst[i] <- NA
      }
    }
  } 
  
  return(chrlst)
  
}



# # # # #


## integrating the text
cntnt <- integ(cntnt)

path_2 <- "C:\\Users\\Mahmood\\Desktop\\alizad_1995\\military\\input\\meta.xlsx"
meta <- read_excel(path_2)

org <- integ(NA_trim(meta$org))
grp <- integ(NA_trim(meta$grp))
sbj <- integ(NA_trim(meta$sbj))
prf <- integ(NA_trim(meta$prf))
exe <- integ(NA_trim(meta$exe))
yer <- integ(NA_trim(meta$yer))

key <- integ(NA_trim(meta$key))


abs <- integ(NA_trim(meta$abs))


list <- list(org, grp, sbj, prf, exe, yer)

id <- vector("list", length(list))

# extracting abs out of columns
# -2 is because of filling abs and key later, not with the other meta keyowrds
names(id) <- colnames(meta)[1:(length(meta)-2)]

# # # # # #


fnd_wrd <- function(text, word, textlen=5000) {
  
  ## returns the phrase after a WORD in a TEXT up until reaching another text
  #note: a maximum default length of 750 initial words is considered for the calculation speed considerations
  
  text <- str_sub(text, 1, textlen)
  
  res <- vector("character", 0)
  
  str <- vector("numeric", 0)
  fin <- vector("numeric", 0)
  
  # description
  bas <- "[^ ]" %R% word %R% zero_or_more(SPACE) %R% optional(":") %R% zero_or_more(SPACE) %R% zero_or_more(NEWLINE)
  
  # description + desired part
  pat <- bas %R% "(.*?)(?=\\r?\\n|\\r$)"
  
  # getting  the indices of the start and finish position of the target text
  for(i in 1:length(word)) {
    str <- unlist(str_locate(text, pat[i]))[1] + 
      (unlist(str_locate(text, bas[i]))[2] - unlist(str_locate(text, bas[i]))[1] + 1)
    fin <- unlist(str_locate(text, pat[i]))[2]
    
    res[i] <- as.character(str_sub(text, str, fin))
    
  }
  
  
  
  return(res)
  
}


# # # # # #



fnd_abs <- function(text, abs, textlen=17500, charlen=175) {
  
  ## returns abstract which comes after the words bas for the beginning texnlen characters of the text
  ## finding abstracts which are longer than charlen characters
  
  text <- str_sub(text, 1, textlen)
  
  res <- vector("character", 0)
  
  # description
  bas <- "[^ ]" %R% abs %R% zero_or_more(SPACE) %R% optional(":") %R% zero_or_more(SPACE) %R% zero_or_more(NEWLINE)
  
  # description + desired part
  pat <- bas %R% "(.*?)(?=\\r?\\n|\\r$)"
  
  for(i in 1:length(abs)) {
    
    len <- vector("numeric", 0)
    
    mat_pat <- matrix(unlist(str_locate_all(text, pat[i])), ncol=2)
    mat_bas <- matrix(unlist(str_locate_all(text, bas[i])), ncol=2)
    nrw <- nrow(mat_pat)
    
    if(nrw>0) {
      for(k in 1:nrw) {
        len[k] <- mat_pat[k, 2] - mat_pat[k, 1]
      }
      
      for(p in 1:length(len)) {
        if(len[p] < charlen) {
          
          mat_pat[p, 1] <- NA
          mat_pat[p, 2] <- NA
          
          mat_bas[p, 1] <- NA
          mat_bas[p, 2] <- NA
        }
      }
      
      mat_pat <- matrix(as.integer(NA_trim(mat_pat)), ncol=2)
      mat_bas <- matrix(as.integer(NA_trim(mat_bas)), ncol=2)
      
      if(nrow(mat_pat)>0) {
        for(j in 1:nrow(mat_pat)) {
          res <- c(res, str_sub(text, mat_bas[j, 2] + 1, mat_pat[j, 2]))
        }
      }
    }
    
  }
  
  return(res)
  
}



# # # # # #


# finding the phrases we're looking for

for (i in 1:length(list)) {
  id[[i]] <- fnd_wrd(cntnt, unlist(list[i]))
}

id$key <- fnd_wrd(cntnt, key, 17500)

# purifying exe based on length less than 30


# removing NA values
for(i in 1:length(id)) {
  id[[i]] <- NA_trim(id[[i]])
}


# exe number of characters is less than 30
id$exe <- max_len(id$exe, 30)

id$exe <- NA_trim(id$exe)


# sbj number of characters is more than 15
id$sbj <- min_len(id$sbj, 15)

id$sbj <- NA_trim(id$sbj)


# yer should have at least two numbers
for(i in 1:length(id$yer)) {
  if(!str_detect(id$yer[i], DIGIT %R% DIGIT)) {
    id$yer[i] <- NA
  }
}

id$yer <- NA_trim(id$yer)


id$abs <- fnd_abs(cntnt, abs)