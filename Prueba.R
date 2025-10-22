# Paquete necesario
library(readxl)

# 1. Cargar el archivo Excel (.xls)
ruta_archivo <- "Data/040_12e_R1.xls"

# 2. Listar todas las hojas
hojas <- excel_sheets(ruta_archivo)
print(hojas)

# 3. Recorrer cada hoja y exportar si tiene contenido
for (s in hojas) {
  cat("Extrayendo hoja:", s, "\n")
  
  # Leer la hoja completa
  df <- tryCatch(
    read_excel(ruta_archivo, sheet = s),
    error = function(e) NULL
  )
  
  # Guardar si hay datos
  if (!is.null(df) && nrow(df) > 0) {
    nombre_archivo <- paste0(gsub("[^A-Za-z0-9]", "_", s), "_export.csv")
    write.csv(df, nombre_archivo, row.names = FALSE)
    cat("✔ Guardado:", nombre_archivo, "\n")
  } else {
    cat("⚠ Hoja vacía o sin datos visibles:", s, "\n")
  }
}
