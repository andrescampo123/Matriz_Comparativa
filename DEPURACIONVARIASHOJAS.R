#ruta archivo
file.choose()

install.packages("readxl")
install.packages("writexl")
install.packages("dplyr")
install.packages("readxl")
library(readxl)
library(writexl)
library(dplyr)
library(readxl)
# Cargar el archivo Excel
archivo <- "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\PREDICCIONDENGUE2023.xlsx"  


# Listar las hojas a procesar
hojas <- c("FEBRERO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "DICIEMBRE")

# Función para eliminar paréntesis y su contenido
eliminar_parentesis <- function(texto) {
  gsub("\\(.*?\\)", "", texto)
}

# Inicializar una lista para guardar los datos modificados
datos_modificados <- list()

# Procesar cada hoja
for (hoja in hojas) {
  # Leer la hoja
  data <- read_excel(archivo, sheet = hoja)
  
  # Modificar la columna "MUNICIPIOS"
  if ("MUNICIPIOS" %in% colnames(data)) {
    data <- data %>%
      mutate(MUNICIPIOS = eliminar_parentesis(MUNICIPIOS))
  }
  
  # Guardar los datos modificados en la lista
  datos_modificados[[hoja]] <- data
}

# Guardar el archivo modificado
write_xlsx(datos_modificados, "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2023\\PREDICCIONDENGUE2023.xlsx")