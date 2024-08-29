#ruta archivo
file.choose()

#Cargando librerias necesarias
install.packages("readxl")
install.packages("dplyr")
install.packages("writexl")
install.packages("purrr")
install.packages("openxlsx")
library(readxl)
library(dplyr)
library(writexl)
library(purrr)
library(openxlsx)
# Mapeo de departamentos a regiones en Colombia. Importante en este ejercicio es primordial depurar solo DEPARTAMENTOS y poner en mayuscula los MUNICIPIOS solamente 
departamentos_a_regiones <- data.frame(
  DEPARTAMENTOS = c("ANTIOQUIA", "ATLÁNTICO", "ATLANTICO", "BOGOTÁ D.C.", "BOGOTA D.C.","BOGOTÁ, D.C.", 
                    "BOLÍVAR", "BOLIVAR", "BOYACÁ", "BOYACA", "CALDAS", "CAQUETÁ", "CAQUETA", "CAUCA", 
                    "CESAR", "CÓRDOBA", "CORDOBA", "CUNDINAMARCA", "CHOCÓ", "CHOCO", "HUILA", "LA GUAJIRA", 
                    "MAGDALENA", "META", "NARIÑO", "NARINO", "NORTE DE SANTANDER", "QUINDÍO", "QUINDIO", 
                    "RISARALDA", "SANTANDER", "SUCRE", "TOLIMA", "VALLE DEL CAUCA", "ARAUCA", "CASANARE", 
                    "PUTUMAYO", "SAN ANDRÉS", "SAN ANDRES", "AMAZONAS", "GUAINÍA", "GUAINIA", "GUAVIARE", 
                    "VAUPÉS", "VAUPES", "VICHADA", "ARCHIPIÉLAGO DE SAN ANDRÉS, PROVIDENCIA Y SANTA CATALINA",
                    "ARCHIPIELAGO DE SAN ANDRÉS, PROVIDENCIA Y SANTA CATALINA", "ARCHIPIELAGO DE SAN ANDRES, PROVIDENCIA Y SANTA CATALINA", 
                    "BOLIVAR", "GUANÍA", "GUANIA", "CHOCO", "CAQUETÁ", "CAQUETA", "BOYACÁ", "BOYACA",
                    "ATLÁNTICO", "ATLANTICO", "CÓRDOBA", "CORDOBA") ,
  
  REGIONES = c("REGIÓN ANDINA", "REGIÓN CARIBE", "REGIÓN CARIBE", "REGIÓN ANDINA", "REGIÓN ANDINA", "REGIÓN ANDINA", 
               "REGIÓN CARIBE", "REGIÓN CARIBE", "REGIÓN ANDINA", "REGIÓN ANDINA", 
               "REGIÓN ANDINA", "REGIÓN AMAZONÍA", "REGIÓN AMAZONÍA", "REGIÓN PACÍFICO", 
               "REGIÓN CARIBE", "REGIÓN CARIBE", "REGIÓN CARIBE", 
               "REGIÓN ANDINA", "REGIÓN PACÍFICO", "REGIÓN PACÍFICO", "REGIÓN ANDINA", "REGIÓN CARIBE", "REGIÓN CARIBE", 
               "REGIÓN ORINOQUÍA", "REGIÓN PACÍFICO", "REGIÓN PACÍFICO", "REGIÓN ANDINA", "REGIÓN ANDINA", "REGIÓN ANDINA", "REGIÓN ANDINA", 
               "REGIÓN ANDINA", "REGIÓN CARIBE", "REGIÓN ANDINA", "REGIÓN PACÍFICO", "REGIÓN ORINOQUÍA", 
               "REGIÓN ORINOQUÍA", "REGIÓN AMAZONÍA", "REGIÓN CARIBE", "REGIÓN CARIBE", "REGIÓN AMAZONÍA", 
               "REGIÓN AMAZONÍA", "REGIÓN AMAZONÍA", "REGIÓN AMAZONÍA", "REGIÓN AMAZONÍA", "REGIÓN AMAZONÍA", 
               "REGIÓN AMAZONÍA", "REGIÓN INSULAR", "REGIÓN INSULAR", "REGIÓN INSULAR", 
               "REGIÓN CARIBE", "REGIÓN AMAZONÍA", "REGIÓN AMAZONÍA", "REGIÓN PACÍFICO", "REGIÓN AMAZONÍA", 
               "REGIÓN AMAZONÍA", "REGIÓN ANDINA", "REGIÓN ANDINA", 
               "REGIÓN CARIBE", "REGIÓN CARIBE", "REGIÓN CARIBE", "REGIÓN CARIBE")
)

# Nombres de las hojas
hojas <- c("FEBRERO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE")

# Crear una lista para almacenar los resultados de cada hoja
resultados <- list()

# Iterar sobre cada hoja, procesar y almacenar el resultado
for (hoja in hojas) {
  # Leer la hoja actual
  datos <- read_excel("C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\PREDICIONDENGUE2023D.xlsx", sheet = hoja)
  # Añadir la columna REGIONES a partir del mapeo de departamentos
  datos <- datos %>%
    left_join(departamentos_a_regiones, by = "DEPARTAMENTOS") %>%  # Unir con el mapeo de regiones
    select(REGIONES, everything())  # Reordenar las columnas para que REGIONES sea la primera
  # Calcular las frecuencias
  frecuencias <- datos %>%
    group_by(REGIONES,`CODIGO_DEPARTAMENTO`, DEPARTAMENTOS) %>%
    summarise(
      AUMENTOS = sum(PREDICCION == "AUMENTO"),
      ESPERADOS = sum(PREDICCION == "ESPERADO"),
      DECREMENTOS = sum(PREDICCION == "DECREMENTO")
    )
  
  # Determinar la variable con mayor frecuencia
  frecuencias <- frecuencias %>%
    mutate(
      Mayor_Frecuencia = case_when(
        AUMENTOS > ESPERADOS & AUMENTOS > DECREMENTOS ~ "AUMENTO",
        ESPERADOS > AUMENTOS & ESPERADOS > DECREMENTOS ~ "ESPERADO",
        DECREMENTOS > AUMENTOS & DECREMENTOS > ESPERADOS ~ "DECREMENTO",
        AUMENTOS == ESPERADOS & AUMENTOS > DECREMENTOS ~ "AUMENTO",
        AUMENTOS == DECREMENTOS & AUMENTOS > ESPERADOS ~ "AUMENTO",
        DECREMENTOS == ESPERADOS ~ "ESPERADO",
        TRUE ~ "AUMENTO"
      )
    )
  
  # Guardar el resultado en la lista
  resultados[[hoja]] <- frecuencias
}

# Escribir todas las hojas procesadas en un solo archivo Excel
write_xlsx(resultados, path = "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\PREDICCIONDENGUE2023REGIONES.xlsx")


# Cargar los archivos
observado_path <- "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\OBSERVADOD.xlsx"
prediccion_path <- "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\PREDICCIONDENGUE2023REGIONES.xlsx"

# Definir las combinaciones de columnas y hojas
combinaciones <- list(
  list(columna = "FEBREROO", hoja = "FEBRERO"),
  list(columna = "ABRILO", hoja = "ABRIL"),
  list(columna = "MAYOO", hoja = "MAYO"),
  list(columna = "JUNIOO", hoja = "JUNIO"),
  list(columna = "JULIOO", hoja = "JULIO"),
  list(columna = "AGOSTOO", hoja = "AGOSTO"),
  list(columna = "SEPTIEMBREO", hoja = "SEPTIEMBRE"),
  list(columna = "OCTUBREO", hoja = "OCTUBRE"),
  list(columna = "NOVIEMBREO", hoja = "NOVIEMBRE"),
  list(columna = "DICIEMBREO", hoja = "DICIEMBRE")
)

# Función para procesar cada combinación
procesar_combinacion <- function(columna, hoja) {
  # Leer las columnas necesarias del archivo observado
  hoja_observado <- read_excel(observado_path, sheet = "Sheet 1")
  rango_seleccionado <- hoja_observado %>%
    select(CODIGO_DEPARTAMENTO, DEPARTAMENTOS, FEBREROO, MARZOO, ABRILO, MAYOO, JUNIOO, JULIOO, AGOSTOO, SEPTIEMBREO, OCTUBREO, NOVIEMBREO, DICIEMBREO) # Selecciona las columnas específicas por nombre o posición
  
  # Filtrar los datos de la columna específica del archivo OBSERVADO
  rangos_combinados <- rango_seleccionado %>%
    select(DEPARTAMENTOS, !!sym(columna)) # Selecciona solo la columna específica
  
  # Leer y seleccionar las columnas de la hoja correspondiente del archivo de PREDICCIONES
  prediccion_combinada <- read_excel(prediccion_path, sheet = hoja) %>%
    select(REGIONES, CODIGO_DEPARTAMENTO, DEPARTAMENTOS, Mayor_Frecuencia) %>%
    mutate(HOJA = hoja) # Añadir una columna para identificar la hoja
  
  # Realizar el inner join para encontrar los departamentos en común
  comunes <- inner_join(rangos_combinados, prediccion_combinada, by = "DEPARTAMENTOS")
  
  # Crear la columna OPERACION
  resultadosc <- comunes %>%
    mutate(OPERACION = case_when(
      !!sym(columna) == "ESPERADO" & Mayor_Frecuencia == "ESPERADO" ~ "SI",
      !!sym(columna) == "ESPERADO" & Mayor_Frecuencia != "ESPERADO" ~ "NO",
      !!sym(columna) == "AUMENTO" & Mayor_Frecuencia == "AUMENTO" ~ "SI",
      !!sym(columna) == "AUMENTO" & Mayor_Frecuencia != "AUMENTO" ~ "NO",
      !!sym(columna) == "DECREMENTO" & Mayor_Frecuencia == "DECREMENTO" ~ "SI",
      TRUE ~ "NO"
    )) %>%
    select(REGIONES, CODIGO_DEPARTAMENTO, DEPARTAMENTOS, !!sym(columna), Mayor_Frecuencia, OPERACION)
  
  return(resultadosc)
}

# Procesar todas las combinaciones y guardarlas en un archivo Excel
resultados_finales <- map(combinaciones, ~ {
  procesar_combinacion(.x$columna, .x$hoja)
})

# Crear un named list para escribir múltiples hojas
resultados_finales_named <- setNames(resultados_finales, sapply(combinaciones, `[[`, "hoja"))


# Guardar los resultados en un archivo Excel
write_xlsx(resultados_finales_named, "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\COMPARATIVO2023.xlsx")

#NOTA::**** EN ESTA PARTE se debe revisar el documento !!!CONTINUAR!!! EN especial FEBREROO... A OBSERVADO

# Ruta. Aquí vamos a sacar el porcentaje por REGIONES
ruta_archivo <- "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\COMPARATIVO2023.xlsx"

# Nombres de las hojas de Excel a procesar
hojas <- c("FEBRERO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE")

# Crear un nuevo archivo Excel
wb <- createWorkbook()

# Función para leer y procesar cada hoja
procesar_hoja <- function(hoja) {
  # Leer la hoja de Excel
  datos <- read_excel(ruta_archivo, sheet = hoja)
  
  # Filtrar las filas según las combinaciones especificadas
  aumento_aumento <- datos %>% filter(Mayor_Frecuencia == "AUMENTO" & Mayor_Frecuencia == "AUMENTO")
  esperado_esperado <- datos %>% filter(Mayor_Frecuencia == "ESPERADO" & Mayor_Frecuencia == "ESPERADO")
  decremento_decremento <- datos %>% filter(Mayor_Frecuencia == "DECREMENTO" & Mayor_Frecuencia == "DECREMENTO")
  
  # Calcular el porcentaje por regiones para cada categoría
  porcentaje_aumento <- aumento_aumento %>%
    group_by(REGIONES) %>%
    summarise(total = n()) %>%
    mutate(Porcentaje_AUMENTO = (total / sum(total)) * 100)
  
  porcentaje_esperado <- esperado_esperado %>%
    group_by(REGIONES) %>%
    summarise(total = n()) %>%
    mutate(Porcentaje_ESPERADO = (total / sum(total)) * 100)
  
  porcentaje_decremento <- decremento_decremento %>%
    group_by(REGIONES) %>%
    summarise(total = n()) %>%
    mutate(Porcentaje_DECREMENTO = (total / sum(total)) * 100)
  
  # Crear un cuadro resumen
  cuadro_resumen <- full_join(porcentaje_aumento, porcentaje_esperado, by = "REGIONES") %>%
    full_join(., porcentaje_decremento, by = "REGIONES") %>%
    select(REGIONES, Porcentaje_AUMENTO, Porcentaje_ESPERADO, Porcentaje_DECREMENTO)
  
  return(cuadro_resumen)
}

# Procesar todas las hojas y escribir cada una en una hoja nueva del archivo Excel
for (i in seq_along(hojas)) {
  hoja <- hojas[i]
  datos_resumen <- procesar_hoja(hoja)
  
  # Añadir una hoja al workbook con un nombre único
  addWorksheet(wb, paste(hoja, "_", i, sep = ""))
  
  # Escribir los datos en la hoja correspondiente
  writeData(wb, sheet = paste(hoja, "_", i, sep = ""), datos_resumen)
}

# Guardar el archivo Excel con múltiples hojas
saveWorkbook(wb, "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\%REGIONES2023.xlsx", overwrite = TRUE)

#como importar archivos de excel a R
file.choose()
#como ver contenido tablas, Hojas y cadenas
head(resultados)
excel_sheets(comparativo_abril)
colnames(comparativo_abril) "CODIGO PARA CONOCER LOS titulos en los rangos"
str(rangos_combinados)
str(prediccion_combinada)
print(resultados)
summary(resultados)