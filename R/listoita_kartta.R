#' Listoita kartta
#'
#' Lukee xlsx-muotoisen näytekartan, ja muokkaa sen Poimuri-yhteensopivaksi listaksi. 
#' Funktio käyttää apuna listoita_kartta_apufunktiot.R -tiedoston funktioita.
#' 
#' 
#' 
#' @param naytekarttalista_fp Polku naytekarttalistalle, joka on excel-tiedosto missä on sarakkeet Aineisto ja Sijainti. Vaihtoehtoinen parametrin 'naytekartat' kanssa.
#' @param naytekartat Dataframe, jossa sarakkeet Aineisto ja Sijainti. Sisältö sama kuin 'naytekarttalista_fp' -parametrin exceltiedstossa. Vaihtoehtoinen parametrin 'naytekarttalista_fp' kanssa.
#' @param export_path Polku, johon .csv-muotoinen listoitettu kartta tallennetaan.
#' @param export_rds_path Polku, johon .Rds-muotoinen listoitettu kartta tallennetaan.
#' @param write TRUE/FALSE. Kirjoitetaanko csv ja rds-tiedostoja?
#' @param doreturn TRUE/FALSE. Palautetaanko koottu lista?
#' @return Listoitettu kartta
#' @export
listoita_kartta <- function(naytekarttalista_fp=NULL, naytekartat=NULL,export_path= './Kohortti_kaikki_naytteet.csv', export_rds_path='./Kohortti_kaikki_naytteet.Rds',write=TRUE, doreturn=FALSE){
  ## Funktiossa vaihtoehtoiset argumentit naytekarttalista_fp, ja naytekartat. Toiseen siis excelpolku, toiseen vastaava dataframe.
  
  if(is.null(naytekartat)){ #jos listaa ei ole tauluna..
    if(is.null(naytekarttalista_fp)){#jos listaa ei ole tauluna TAI polkuna, annetaan virhe.
      stop("Anna joko näytekartta-excel polku, tai naytekartat-taulu, jossa sarakkeet Aineisto ja Sijainti.")
    }
    if(!is.null(naytekarttalista_fp)){
      naytekartat<-openxlsx::read.xlsx(naytekarttalista_fp) #jos listaa ei ole tauluna, mutta ON polkuna, luetaan polun excel.
    }
  }
  
  # Käydään kaikki määritellyt kartat läpi loopissa
  naytedata<-data.frame()
  
  for(i in 1:nrow(naytekartat)){
    t00<-suppressWarnings(haeAineisto(hakuaineisto = naytekartat[i,1], naytepolku = naytekartat[i,2]))
    t00$aineisto<-naytekartat[i,1]
    t00$naytekartta<-naytekartat[i,2]
    naytedata<-plyr::rbind.fill(naytedata,t00)
  }
  
  
  # Poistetaan sarakkeista 1-11 kaikki X<numero> alkuiset tiedot, nämä ovat turhia ja vain sekoittavat
  for(i in 1:11){
    naytedata[grepl('^X[1-9]',naytedata[,i]),i]<-NA
  }
  
  # Nimetään sarakkeet paremmin
  colnames(naytedata)<-c('rasianimi', 'rasia', 'hylly', 'rakki', 'header5', 'rakkihylly', 'header7', 'pakastin', 'header9', 'header10', 'header11', 'rasialeveys', 'rasiaid', 'rasiapaikka', 'naytenimi', 'keraysidn', 'keraysid','nayte','lisatietoa','pakastinnro','hyllynro','rakkinro','rakkihyllynro','rakkihyllynro1','rakkihyllynro2','rasianro','rasiapaikka_x','rasiapaikka_y','aineisto','naytekartta')
  
  # Pudotetaan _ -merkki pois rasiapaikka_x:stä
  naytedata$rasiapaikka_x<-stringr::str_replace_all(naytedata$rasiapaikka_x,'_','')
  
  # Pudotetaan lisätiedot pois keraysidn:stä
  naytedata$keraysidn<-stringr::str_extract(naytedata$keraysidn,'[^_]*')
  
  # Järjestetään sarakkeet paremmin
  col_order<-c(c('keraysidn','naytenimi','keraysid','nayte','lisatietoa'),colnames(naytedata)[!colnames(naytedata) %in% c('keraysidn', 'naytenimi', 'keraysid', 'nayte', 'lisatietoa')])
  naytedata<-naytedata[,col_order]
  
  # Pudotetaan ylimääräiset pisteet ja välilyönnit halutuista sarakkeista
  for(coln in c('rasianimi','rasia','hylly','rakki','header5','rakkihylly','header7','pakastin','header9','header10','header11')){
    naytedata[,coln]<-siisti(naytedata[,coln])
  }
  
  if(write==TRUE){
    # Tallennetaan
    data.table::fwrite(naytedata,export_path)
    saveRDS(naytedata, export_rds_path)
  }
  if(doreturn==TRUE){
    return(naytedata)
  }
}
