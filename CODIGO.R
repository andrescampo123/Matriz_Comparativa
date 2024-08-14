#Cargando librerias necesarias
install.packages("readxl")
library(readxl)
install.packages("dplyr")
library(dplyr)
install.packages("writexl")
library(writexl)
install.packages("purrr")
library(purrr)

# Mapeo de departamentos a regiones en Colombia
departamentos_a_regiones <- data.frame(
  DEPARTAMENTOS = c(
    "AMAZONAS", "ANTIOQUIA", "ARAUCA", "ATLÁNTICO", "BOLÍVAR", "BOYACÁ", "CALDAS", "CAQUETÁ", 
    "CASANARE", "CAUCA", "CESAR", "CHOCÓ", "CÓRDOBA", "CUNDINAMARCA", "GUAINÍA", "GUAVIARE", 
    "HUILA", "LA GUAJIRA", "MAGDALENA", "META", "NARIÑO", "NORTE DE SANTANDER", "PUTUMAYO", 
    "QUINDÍO", "RISARALDA", "SAN ANDRÉS Y PROVIDENCIA", "SANTANDER", "SUCRE", "TOLIMA", 
    "VALLE DEL CAUCA", "VAUPÉS", "VICHADA"
  ),
  REGIONES = c(
    "REGIÓN AMAZONÍA", "REGIÓN ANDINA", "REGIÓN ORINOQUÍA", "REGIÓN CARIBE", "REGIÓN CARIBE", 
    "REGIÓN ANDINA", "REGIÓN ANDINA", "REGIÓN AMAZONÍA", "REGIÓN ORINOQUÍA", "REGIÓN PACÍFICA", 
    "REGIÓN CARIBE", "REGIÓN PACÍFICA", "REGIÓN CARIBE", "REGIÓN ANDINA", "REGIÓN AMAZONÍA", 
    "REGIÓN AMAZONÍA", "REGIÓN ANDINA", "REGIÓN CARIBE", "REGIÓN CARIBE", "REGIÓN ORINOQUÍA", 
    "REGIÓN PACÍFICA", "REGIÓN ANDINA", "REGIÓN AMAZONÍA", "REGIÓN ANDINA", "REGIÓN ANDINA", 
    "REGIÓN CARIBE", "REGIÓN ANDINA", "REGIÓN CARIBE", "REGIÓN ANDINA", "REGIÓN PACÍFICA", 
    "REGIÓN AMAZONÍA", "REGIÓN ORINOQUÍA"
  )
)
# Nombres de las hojas
hojas <- c("FEBRERO", "ABRIL", "MAYO", "JUNIO","JULIO","AGOSTO","SEPTIEMBRE","DICIEMBRE")

# Crear una lista para almacenar los resultados de cada hoja
resultados <- list()

# Iterar sobre cada hoja, procesar y almacenar el resultado
for (hoja in hojas) {
# Leer la hoja actual
datos <- read_excel("C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\PREDICCIONDENGUE2023.xlsx", sheet = hoja)
# Añadir la columna REGIONES a partir del mapeo de departamentos
datos <- datos %>%
left_join(departamentos_a_regiones, by = "DEPARTAMENTOS") %>%  # Unir con el mapeo de regiones
select(REGIONES, everything())  # Reordenar las columnas para que REGIONES sea la primera
# Calcular las frecuencias
  frecuencias <- datos %>%
    group_by(`CODIGO DEPARTAMENTO`, REGIONES, DEPARTAMENTOS, MUNICIPIOS) %>%
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
write_xlsx(resultados, path = "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\PREDICCIONDENGUE2023FILTRADO.xlsx")


# Cargar los archivos
observado_path <- "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\OBSERVADO2023R.xlsx"
prediccion_path <- "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\PREDICCIONDENGUE2023FILTRADO.xlsx"

# Definir las combinaciones de columnas y hojas
combinaciones <- list(
  list(columna = "FEBREROO", hoja = "FEBRERO"),
  list(columna = "ABRILO", hoja = "ABRIL"),
  list(columna = "MAYOO", hoja = "MAYO"),
  list(columna = "JUNIOO", hoja = "JUNIO"),
  list(columna = "JULIOO", hoja = "JULIO"),
  list(columna = "AGOSTOO", hoja = "AGOSTO"),
  list(columna = "SEPTIEMBREO", hoja = "SEPTIEMBRE"),
  list(columna = "DICIEMBREO", hoja = "DICIEMBRE")
)

# Función para procesar cada combinación
procesar_combinacion <- function(columna, hoja) {
  # Leer las columnas necesarias del archivo observado
  hoja_observado <- read_excel(observado_path, sheet = "Hoja1")
  rango_seleccionado <- hoja_observado %>%
    select(DEPARTAMENTOS, MUNICIPIOS, FEBREROO, ABRILO, MAYOO, JUNIOO, JULIOO, AGOSTOO, SEPTIEMBREO, DICIEMBREO) # Selecciona las columnas específicas por nombre o posición
  
  
  # Filtrar los datos de la columna específica del archivo OBSERVADO
  rangos_combinados <- rango_seleccionado %>%
    select(DEPARTAMENTOS, MUNICIPIOS, !!sym(columna)) # Selecciona solo la columna específica
  
  # Leer y seleccionar las columnas de la hoja correspondiente del archivo de PREDICCIONES
  prediccion_combinada <- read_excel(prediccion_path, sheet = hoja) %>%
    select(REGIONES,DEPARTAMENTOS, MUNICIPIOS, Mayor_Frecuencia) %>%
    mutate(HOJA = hoja) # Añadir una columna para identificar la hoja
  
  # Realizar el inner join para encontrar los municipios y departamentos en común
  comunes <- inner_join(rangos_combinados, prediccion_combinada, by = c("DEPARTAMENTOS", "MUNICIPIOS"))
  
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
    select(REGIONES, DEPARTAMENTOS, MUNICIPIOS, !!sym(columna), Mayor_Frecuencia, OPERACION)
  
  return(resultadosc)
}

# Procesar todas las combinaciones y guardarlas en un archivo Excel
resultados_finales <- map(combinaciones, ~ {
  procesar_combinacion(.x$columna, .x$hoja)
})

# Crear un named list para escribir múltiples hojas
resultados_finales_named <- setNames(resultados_finales, sapply(combinaciones, `[[`, "hoja"))

# Guardar los resultados en un archivo Excel
write_xlsx(resultados_finales, "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\COMPARATIVO2023.xlsx")


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