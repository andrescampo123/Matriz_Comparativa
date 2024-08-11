#Cargando librerias necesarias
install.packages("readxl")
library(readxl)
install.packages("dplyr")
library(dplyr)
install.packages("writexl")
library(writexl)
install.packages("purrr")
library(purrr)

# Nombres de las hojas
hojas <- c("FEBRERO", "ABRIL", "MAYO", "JUNIO","JULIO","AGOSTO","SEPTIEMBRE","DICIEMBRE")

# Crear una lista para almacenar los resultados de cada hoja
resultados <- list()

# Iterar sobre cada hoja, procesar y almacenar el resultado
for (hoja in hojas) {
  # Leer la hoja actual
  datos <- read_excel("C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\PREDICCIONDENGUE2023.xlsx", sheet = hoja)
  
  # Calcular las frecuencias
  frecuencias <- datos %>%
    group_by(`CODIGO DEPARTAMENTO`, DEPARTAMENTOS, MUNICIPIOS) %>%
    summarise(
      INCREMENTOS = sum(PREDICCION == "INCREMENTOS"),
      ESPERADO = sum(PREDICCION == "ESPERADO"),
      DECREMENTO = sum(PREDICCION == "DECREMENTO")
    )
  
  # Determinar la variable con mayor frecuencia
  frecuencias <- frecuencias %>%
    mutate(
      Mayor_Frecuencia = case_when(
        INCREMENTOS > ESPERADO & INCREMENTOS > DECREMENTO ~ "INCREMENTOS",
        ESPERADO > INCREMENTOS & ESPERADO > DECREMENTO ~ "ESPERADO",
        DECREMENTO > INCREMENTOS & DECREMENTO > ESPERADO ~ "DECREMENTO",
        INCREMENTOS == ESPERADO & INCREMENTOS > DECREMENTO ~ "INCREMENTOS",
        INCREMENTOS == DECREMENTO & INCREMENTOS > ESPERADO ~ "INCREMENTOS",
        DECREMENTO == ESPERADO ~ "ESPERADO",
        TRUE ~ "INCREMENTOS"
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
    select(DEPARTAMENTOS, MUNICIPIOS, Mayor_Frecuencia) %>%
    mutate(HOJA = hoja) # Añadir una columna para identificar la hoja
  
  # Realizar el inner join para encontrar los municipios y departamentos en común
  comunes <- inner_join(rangos_combinados, prediccion_combinada, by = c("DEPARTAMENTOS", "MUNICIPIOS"))
  
  # Crear la columna OPERACION
  resultadosc <- comunes %>%
    mutate(OPERACION = case_when(
      !!sym(columna) == "ESPERADO" & Mayor_Frecuencia == "ESPERADO" ~ "SI",
      !!sym(columna) == "ESPERADO" & Mayor_Frecuencia != "ESPERADO" ~ "NO",
      !!sym(columna) == "BROTE" & Mayor_Frecuencia == "BROTE" ~ "SI",
      !!sym(columna) == "BROTE" & Mayor_Frecuencia != "BROTE" ~ "NO",
      !!sym(columna) == "AUMENTO" & Mayor_Frecuencia == "AUMENTO" ~ "SI",
      !!sym(columna) == "AUMENTO" & Mayor_Frecuencia != "AUMENTO" ~ "NO",
      !!sym(columna) == "DECREMENTO" & Mayor_Frecuencia == "DECREMENTO" ~ "SI",
      TRUE ~ "NO"
    )) %>%
    select(DEPARTAMENTOS, MUNICIPIOS, !!sym(columna), Mayor_Frecuencia, OPERACION)
  
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
