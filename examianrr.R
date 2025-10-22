library(readxl)

ruta_archivo <- "Data/040_12e_R1.xls"
hojas <- excel_sheets(ruta_archivo)

# Resumen de filas y columnas por hoja
resumen <- data.frame(
  hoja = character(),
  filas = numeric(),
  columnas = numeric(),
  stringsAsFactors = FALSE
)

for (s in hojas) {
  df <- tryCatch(read_excel(ruta_archivo, sheet = s), error = function(e) NULL)
  if (!is.null(df)) {
    resumen <- rbind(resumen, data.frame(
      hoja = s,
      filas = nrow(df),
      columnas = ncol(df)
    ))
  }
}

print(resumen)
