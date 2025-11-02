import pandas as pd
import openpyxl
import os
import scripts.routes as gv
import matplotlib.pyplot as plt


# Leer la hoja principal (por defecto lee la primera)
df = pd.read_excel(gv.ARCHIVOS_POR_CARPETA[0][0], engine="openpyxl")  # usa engine="xlrd" si es .xls

# Ver las primeras filas
print(df.head())
df['Column1'].hist()
plt.show()
