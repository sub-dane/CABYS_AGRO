#' @export
# Frutas
# Cargar la biblioteca readxl

f_Frutas<-function(directorio,mes,anio){

  #Cargar librerias
  library(readxl)
  library(dplyr)
  library(openxlsx)
  library(zoo)

  #identificar la carpeta
  carpeta=nombre_carpeta(mes,anio)



  # Exportaciones ------------------------------------------------------------------
  nombre_archivos=read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Doc/Nombres_archivos_",nombres_meses[mes],".xlsx"),sheet = "Nombres")
  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Exportaciones","NOMBRE"]


  # Especifica la ruta del archivo de Excel
  archivos=list.files(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE"))
  elementos_seleccionados <- archivos[grepl("Expos e", archivos) ]
  # Especifica la ruta del archivo de Excel
  Frutas <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                       sheet = "TOTAL EXPO_KTES")


  n_fila=which(Frutas == "010499" |Frutas == "010403",arr.ind = TRUE)[,"row"]
  n_fila=c(n_fila[[1]],n_fila[[2]])
  n_col_1=which(Frutas== paste0(anio," ",mes),arr.ind = TRUE)[,"col"]
  n_col_2=which(Frutas== paste0((anio-1)," ",1),arr.ind = TRUE)[,"col"]



  Frutas2 <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                      sheet = "CTES FOBPES")
  n_fila_2=which(Frutas2 == "010499" |Frutas2 == "010403",arr.ind = TRUE)[,"row"]
  n_fila_2=c(n_fila[[1]],n_fila[[2]])
  n_col_1_2=which(Frutas2== paste0(anio," ",mes_0[mes]),arr.ind = TRUE)[,"col"]
  n_col_2_2=which(Frutas2== paste0((anio-1)," 01"),arr.ind = TRUE)[,"col"]

  Frutas3 <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/",elementos_seleccionados,"/",archivo),
                       sheet = "IP_EXPO")


   valor_exportaciones=as.data.frame(cbind(t(Frutas[n_fila,n_col_2[1]:n_col_1[1]]),t(Frutas2[n_fila_2,n_col_2_2[1]:n_col_1_2[1]]),t(Frutas3[n_fila_2,n_col_2_2[1]:n_col_1_2[1]])))
   for (i in 1:6) {
     valor_exportaciones[,i]=as.numeric(valor_exportaciones[,i])
   }






  # Consumo interno ---------------------------------------------------------
archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Clasificacion","NOMBRE"]
   clasificacion=read_excel(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Frutas/",archivo))
   vector_otras=clasificacion$`Otras frutas`
   vector_citricas=clasificacion$`Frutas citricas`
   archivo=nombre_archivos[nombre_archivos$PRODUCTO=="Microdato","NOMBRE"] 
  Frutas <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/Datos_SIPSA/Microdatos desde 2013/",archivo),sheet = "2.1",startRow = 6)

  
  Frutas[,"Fecha"] <- as.Date(Frutas[,"Fecha"], origin = "1899-12-30")
  Frutas$Cant.Kg=as.numeric(Frutas$Cant.Kg)
  otras_frutas= Frutas %>%
    mutate(año = year(Fecha), mes = month(Fecha)) %>%
    filter(Alimento %in% vector_otras,Departamento.Proc.!="OTRO") %>% 
    group_by(año, mes) %>% 
    mutate(cantidad=sum(Cant.Kg)) %>% 
    select(año,mes,cantidad) %>% 
    unique()
valor_otras=otras_frutas[otras_frutas$año==anio & otras_frutas$mes==mes,"cantidad"]

citricas= Frutas %>%
  mutate(año = year(Fecha), mes = month(Fecha)) %>%
  filter(Alimento %in% vector_citricas,Departamento.Proc.!="OTRO") %>% 
  group_by(año, mes) %>% 
  mutate(cantidad=sum(Cant.Kg)) %>% 
  select(año,mes,cantidad) %>% 
  unique()
valor_citricas=citricas[citricas$año==anio & citricas$mes==mes,"cantidad"]
    

  Valor_Frutas=as.numeric(c(valor_otras,valor_citricas))




# IPP ---------------------------------------------------------------------

  archivo=nombre_archivos[nombre_archivos$PRODUCTO=="IPP","NOMBRE"]

  IPP <- read.xlsx(paste0(directorio,"/ISE/",anio,"/",carpeta,"/Data/consolidado_ISE/Precios/",archivo),
                       sheet = "3.1")

  n_col=which(IPP == "CODIGO",arr.ind = TRUE)[, "col"]

  dos_digitos=anio %% 100
  #identificar las columna donde dice total general y peso en pie
  #identificar las columna donde dice total general y peso en pie



  columna1=max(which(grepl(paste0(nombres_siglas[1],"-",(dos_digitos-2)),IPP),arr.ind = TRUE))
  columna2=which(grepl(paste0(nombres_siglas[mes],"-",dos_digitos),IPP),arr.ind = TRUE)

  fila1=which(grepl("01310",IPP[,n_col]),arr.ind = TRUE)


  #Tomar el valor que nos interesa
  Valor_IPP=as.data.frame(t(IPP[c(fila1[1]),c(columna1:columna2[1])]))
  tamaño=nrow(Valor_IPP)
  filas=c(seq(1, tamaño, by = 2),tamaño)
  Valor_IPP=as.numeric(Valor_IPP[filas,1])
  # Agrupar datos -----------------------------------------------------------


  return(list(variacion = valor_exportaciones, vector = Valor_Frutas,IPP=Valor_IPP))
}
