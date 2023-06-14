#' Kartoita lista
#'
#' Funktio muuttaa listamuotoon kootun näytekartan takaisin karttamuotoon. 
#' Argumentteja ovat kartoitettavan listan (.csv-muotoinen) tiedostopolku, rasioiden rivi- ja sarakemäärät,
#' tiedostopolku tallennettavalle kartalle, sekä valinnainen solujen värit määräävä taulu.
#' Tehtävälle kartalle lihavoidaan otsikkorivit, ja 
#'
#' @param kartoitettava_lista_fp Polku, jossa listamuotoinen kartta on. Tiedostomuoto oltava csv.
#' @param datrows Tehtävän kartan rivimäärä. Oletuksena 10, mutta käytössä on myös 8 ja 12.
#' @param datcols Tehtävän kartan sarakemäärä. Oletuksena 10, mutta käytössä on myös 8 ja 12.
#' @param target_fp Polku, johon valmis kartta tallennetaan. Tiedostomuoto xlsx.
#' @param color_vector Dataframe, jossa sarakkeet 'naytenimi' ja 'color'. 'naytenimi' on jonkin näytteen nimi sellaisena kuin se kartalle merkitään. 'color' on openxlsx-yhteensopiva väri, esim 'red', tai hex-koodi. 
#' @export
kartoita_lista <- function(kartoitettava_lista_fp, datrows=10, datcols=10, target_fp, color_vector=NA){
  #Luetaan data
  input <- read.csv(kartoitettava_lista_fp)
  
  #Kootaan kartta
  kartta <- data.frame()
  kartta_colnames <- glue("header{1:datcols}")
  for(i in 1:length(unique(input$rasiaid))){
    block <- input[input$rasiaid == unique(input$rasiaid)[i],]  
    karttablock <- data.frame(rivinum=1:datrows)
    for(j in 1:datcols){
      #Kartan rasia kerrallaan kootaan dataframe rasian ID:istä.
      newcol <- block[block$rasiapaikka_x == unique(block$rasiapaikka_x)[j],c("naytenimi","rasiapaikka_y")]
      karttablock <- dplyr::left_join(karttablock, newcol, by=c("rivinum"="rasiapaikka_y"))
    }
    #Lisätään rasialle sarakeotsikot
    names(karttablock) <- unique(block$rasiapaikka_x)
    header_row2 <- unique(block[,c('rasianimi','rasia','hylly','rakki','header5','rakkihylly','header7','pakastin','header9','header10','header11')])
    header_row1 <- c(NA, letters[1:datcols])
    karttablock <- rbind(setNames(header_row1, kartta_colnames), setNames(karttablock,kartta_colnames))
    karttablock <- rbind(setNames(header_row2, kartta_colnames), karttablock)
    names(karttablock)  <- glue("header{1:ncol(karttablock)-1}")
    #Lisätään rasia karttatauluun
    karttablock <- rbind(karttablock, setNames(rep(NA, ncol(karttablock)), names(karttablock)))
    kartta <- rbind(kartta, karttablock)
  }
  
  #Tallennetaan kartta excelinä.
  wb<-openxlsx::createWorkbook()
  datestamp <- format(Sys.time(), "%y-%m-%d")
  openxlsx::addWorksheet(wb,datestamp)
  openxlsx::writeData(wb,sheet = 1,kartta, rowNames = F, colNames=F)
  #Käydään läpi värivektori. Väritetään pyydetyt solut.
  if(!is.na(color_vector)){
    for(k in 1:length(unique(color_vector$color))){
      fill_color <- unique(color_vector$color)[k]
      style <- openxlsx::createStyle(fgFill = fill_color)
      #matchaavat koordinaatit
      coords_to_fill <- sapply(color_vector$naytenimi[color_vector$color == fill_color], FUN=function(x) which(kartta==x, arr.ind=TRUE))
      coords_to_fill <- do.call(rbind, coords_to_fill)
      row.names(coords_to_fill)<- NULL
      colrows <- coords_to_fill[,1]
      colcols <- coords_to_fill[,2]
      openxlsx::addStyle(wb, sheet = 1, style, rows = colrows, cols = colcols, stack=TRUE)
    }
  }

  #boldataan otsikkorivit ja -sarakkete
  boldstyle <- openxlsx::createStyle(textDecoration=c('bold'))
  boldrivit <- rep(c(which(kartta[,1]==1)-1, which(kartta[,1]==1)-2), each=datcols+1)
  boldsarakkeet <- rep(1:(datcols+1), 2*length(which(kartta[,1]==1)))
  openxlsx::addStyle(wb, sheet = 1, boldstyle, rows = boldrivit, cols = boldsarakkeet , stack=TRUE)
  
  #Lisätään reunat laatikoille
  borderstyle <- openxlsx::createStyle(border = "TopBottomLeftRight")
  borderrows <- which(rowSums(is.na(kartta)) != 11)
  bordercols <- 1:(datrows+1)
  openxlsx::addStyle(wb, sheet = 1, borderstyle, rows = borderrows, cols = bordercols, gridExpand = TRUE, stack=TRUE)
  
  #Tallennetaan
  openxlsx::saveWorkbook(wb,target_fp,overwrite = T)
}



