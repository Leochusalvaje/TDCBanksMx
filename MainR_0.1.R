library(skimr)
library(janitor)
library(tidyverse)
library(ggplot2)
library(readxl)

data_TDC <- read_excel("Data/Dataclean.xlsx",sheet = 1)
DLC_Actinver <- read_excel("Data/Dataclean.xlsx",sheet = 3)


df <- read_excel("Data/040_12e_R1.xls", sheet = "BD")
write.csv(df, "BD_export.csv", row.names = FALSE)
# Ver las primeras filas y columnas
head(df)
dim(df)   # n
view(df)