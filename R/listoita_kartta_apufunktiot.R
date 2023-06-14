
## ----Kirjastot-----------------------------------------------------------------------------------------------
# library(plyr)
# library(openxlsx)
# library(stringr)
# library(data.table)
# library(haven)
# library(dplyr)


## ----Apufunktiot---------------------------------------------------------------------------------------------
# Haetaan kartasta kaikki näytteet ja tehdään niistä lista
#' @export
rect2list <- function(x, nimi, i, checkmode=F){
  
  k <- i
  # Tyhjät välilyönnit pois kaikkien solujen aluista ja lopuista
  trim <- function (x) gsub("^\\s+|\\s+$", "", x)
  
  # Etsitään kaikki rivit jotka sisältävät ainakin yhden solun jonka sisältö on "1"
  x[] <- sapply(x[], trim)   
  ind <- apply(x == 1, 1, which)
  
  # Onhan sivulla rasioita
  if(length(ind)>0){
    
    for(a in 1:length(ind)){
      if(length(names(ind[[a]]))>0){
        for(b in 1:length(names(ind[[a]]))){
          names(ind[[a]])[b]<-paste(a,names(ind[[a]])[b],sep='.')
        }
      }
    }
    
    # Haetaan kyseisiltä riveiltä kaikkien solujen koordinaatit joiden sisältö on "1"
    ind <- unlist(ind)
    ind <- cbind(sapply(names(ind), function(x) strsplit(x, ".", fixed = TRUE)[[1]][1]), ind)
    
    ind<-data.frame(ind)
    
    # Muutetaan koordinaatit numeerisiksi
    if(nrow(ind)==1){
      ind[,1]<-as.numeric(as.character(ind[,1]))
      ind[,2]<-as.numeric(as.character(ind[,2]))
    }else{
      ind <- apply(ind, 2, as.double) 
    }
    
    
    # Käydään läpi kaikki koordinaatit, etsitään rasian dimensiot ja luodaan rasian näytteistä lista
    complete <- vector("list", nrow(ind))
    
    rasianum<-0
    
    for(a in seq(complete)){
      
      rasianum<-rasianum+1
      
      # Haetaan rasian vasen yläkulma
      block <- x[(ind[a, 1]-2):min(ind[a, 1]+9, nrow(x)), ind[a, 2]:min(ind[a, 2]+12, ncol(x))]
      
      # Rasian sarakkeiden nimet
      rasia <- block[1, grepl("rasia", block[1, ], ignore.case = TRUE)]
      
      if(length(block[1, grepl("rasia", block[1, ], ignore.case = TRUE)])>0){
        rasia <- block[1, grepl("rasia", block[1, ], ignore.case = TRUE)]
      }else if(length(block[1, grepl("lista", block[1, ], ignore.case = TRUE)])>0){
        rasia <- block[1, grepl("lista", block[1, ], ignore.case = TRUE)]
      }else{
        rasia<-''
      }
      
      if(is.na(block[1,1])){
        block[1,1]<-'Tyhja'
      }
      
      # Pudotetaan duplikaatit sarake/rivinimet pois vasemmalta oikealle ja ylhäältä alaspäin, tällöin ei haittaa jos alunperin tuli mukaan sarakkeita seuraavasta rasiasta, sillä ne putoavat tässä pois
      block <- block[!duplicated(as.character(block[, 1])), !duplicated(as.character(block[2, ]))]
      
      #21052019
      # Pudotetaan sarakkeet/rivit, jotka eivät sisällä (a,b,c,...) tai (1,2,3,...) eli rivit jotka eivät kuulu rasian näyteosaan
      if(ncol(block)>6){
        for(b in ncol(block):7){
          if(!(as.character(block[2,b]) %in% letters) && !(as.character(block[2,b]) %in% LETTERS)){
            block<-block[,c(-b)]
          }
        }
      }
      if(nrow(block)>10){
        for(b in nrow(block):11){
          if(is.na(as.numeric(block[b,1]))){
            block<-block[c(-b),]
          }
        }
      }
      
      # Irroitetaan näyteosa erikseen
      samples <- block[3:nrow(block), 2:ncol(block)]
      colnames(samples) <- paste(block[2, 2:ncol(block)], "_", sep = "")
      
      samples <- unlist(samples)
      Encoding(names(samples)) <- "latin1"
      
      headerrow<-trimws(unlist(block[1,]))
      headerrow[grepl('^X[1-9]',headerrow)]<-NA
      headerrow[is.na(headerrow)]<-''
      headerrow[headerrow=='']<-NA
      rasiawi<-length(headerrow)-1
      
      if(length(headerrow)==6){
        headerrow<-c(headerrow[1:4], NA, headerrow[5], NA, headerrow[6])
      }
      
      if(length(headerrow)<11)
        headerrow<-c(headerrow,rep(NA,11-length(headerrow)))
      
      if(ncol(block)>11)
        cat('Rasiassa  yli 11 saraketta:',i, block[1,1],block[1,2], samples[1],'\n')
      
      headerrow<-headerrow[1:11]
      headerrow<-c(headerrow, rasiawi, paste0(nimi,'-',k,'-',rasianum))
      
      headerrow<-data.frame(t(headerrow))
      complete[[a]] <- data.frame(headerrow, names(samples), samples)
      
      
      
      colnames(complete[[a]]) <- 1:ncol(complete[[a]])
    }
    complete <- plyr::rbind.fill(complete)
    
  }else{ # Sivulla ei ole rasioita
    complete<-data.frame()
  }
  
  return(complete)
}


# Luetaan määritelty excelin sivu ja haetaan siitä kaikki näytteet
#' @export
lueSheet <- function(tiedosto,nimi,i,checkmode=F){
  x <- openxlsx::read.xlsx(tiedosto, i,skipEmptyRows=F,skipEmptyCols = F)
  
  # Tarkistetaan että eihän sivu ole tyhjä
  if(length(x)>0){
    x <- rbind(colnames(x), x)
    y <- rect2list(x,nimi,i,checkmode)
    
  }else{
    y<-data.frame()
  }
  
  return(y)
}


# Luetaan näytekartta
#' @export
lueNaytekartta<-function(tiedosto,nimi,sheetnums,checkmode=F){
  naytedata<-data.frame()
  
  for(i in 1:length(sheetnums)){
    naytedata<-plyr::rbind.fill(naytedata,lueSheet(tiedosto,nimi,i,checkmode))
    
  }
  
  names(naytedata) <-c('rasianimi', "rasia", "hylly", "rakki", 'header5', "rakkihylly",'header7', "pakastin", 'header9', 'header10','header11','rasialeveys','rasiaid', "rasiapaikka", "naytenimi")
  
  return(naytedata)
}


# Lasketaan näytemäärät
#' @export
naytemaarat <- function(i,data1,  lista){
  if(length(lista[[i]])>0){
    nn <- vector("numeric")
    for(k in 1:length(lista[[i]])){
      n=length(data1[which(data1$nayte%in%lista[[i]][k]), "nayte"])
      nn <- c(nn, n)
      nn
    }
  }else nn <- 0
  nn
}


#' @export
haeAineisto <- function(hakuaineisto='',naytepolku='',naytekarttalista){
  cat('# Aineisto:',hakuaineisto,'\n')
  cat('Ladataan naytekarttaa > ')
  
  naytedata<-data.frame()
  
  # Aineisto annettu, mutta ei naytepolkua
  if(hakuaineisto!='' && naytepolku==''){
    
    # Haetaan taulukko kaikkien näytekarttojen sijainneista
    karttatable<-openxlsx::read.xlsx(naytekarttalista,1,skipEmptyRows = F,skipEmptyCols = F)
    
    # Match aineiston nimi
    karttatable<-karttatable[karttatable[,1]==hakuaineisto,]
    
    loydetty<-F
    
    # Löydettiin aineisto
    if(nrow(karttatable)>0){
      # Valitaan aina ensimmäinen polku, jos ainesto oli useaan kertaan näytekarttalistalla
      naytepolku<-karttatable[1,2]
      if(!is.na(naytepolku) && naytepolku!=''){
        loydetty<-T
      }
    }
  }else if(hakuaineisto!='' && naytepolku!=''){
    loydetty<-T
  }
  
  if(loydetty){
    
    shnames<-getSheetNames(naytepolku)
    shnum<-length(shnames)
    
    naytedata<-lueNaytekartta(tiedosto = naytepolku, nimi=hakuaineisto, c(1:shnum))
  }
  
  if(loydetty){
    cat('Valmistellaan naytekoodeja > ')
    
    # Id:t ulos viivakoodeista
    naytedata$nekuidn <- as.character(naytedata[,'naytenimi'])
    naytedata$nekuidn <- toupper(naytedata$nekuidn)
    #  naytedata <- within(naytedata, {
    naytedata$nekuid <- gsub("\\D", " ", naytedata$nekuidn)
    naytedata$nekuid <- as.numeric(gsub(" .*", "", naytedata$nekuid))
    # })
    naytedata <- naytedata[!is.na(naytedata$nekuid),]
    
    # Irroitetaan viivakoodeista lopppuosa
    naytedata$nayte <- gsub("^\\d*", "", naytedata$nekuidn)
    naytedata$nayte <- toupper(naytedata$nayte)
    naytedata$nayte <- gsub(" ", "", naytedata$nayte)
    
    naytedata$lisatietoa<-naytedata$nayte

    # Näytteisiin vain näytekoodi
    naytedata$nayte<-gsub('_(.*)','',naytedata$nayte)
    
    # Lisätietoa sarakkeeseen näytteen leimat
    naytedata$lisatietoa<-paste0(naytedata$lisatietoa,' _')
    naytedata$lisatietoa<-gsub('^(.*?)_','',naytedata$lisatietoa)
    naytedata$lisatietoa<-gsub('([^ ]+$)','',naytedata$lisatietoa)
    
    # Haetaan paikkatiedot
    naytedata$pakastinno <- as.numeric(gsub("\\D", "", naytedata$pakastin))
    naytedata$hyllyno <- as.numeric(gsub("\\D","", naytedata$hylly))
    naytedata$rakkino <- as.numeric(gsub("\\D","", naytedata$rakki))
    naytedata$rakkihyllyno <- gsub("\\D","", naytedata$rakkihylly)
    naytedata$rakkihyllyno1 <- substr(naytedata$rakkihyllyno, 1, 1)
    naytedata$rakkihyllyno2 <- substr(naytedata$rakkihyllyno, 2, 2)
    naytedata$rasiano <- as.numeric(gsub("\\D","", naytedata$rasia))
    naytedata$rasiapaikkaki <- gsub("\\d", "", naytedata$rasiapaikka)
    naytedata$rasiapaikkano <- as.numeric(gsub("\\D", "", naytedata$rasiapaikka))
    
  }
  
  cat('Valmis!\n')
  
  return(naytedata) 
}


#' @export
siisti<-function(strvec){
  strvec<-as.character(strvec)
  strvec[is.na(strvec)]<-''
  strvec<- stringr::str_replace_all(strvec,'\\.',' ')
  strvec<- stringr::str_replace_all(strvec,'\\s+',' ')
  strvec<-tolower(strvec)
  
  return(strvec)
}




