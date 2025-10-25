import pyautogui
import time
import pandas as pd
import openpyxl
import win32com.client as win32

# -------------------------------
# CONFIGURACIÓN
# -------------------------------
# Ruta del libro CNBV
ruta_cnbv = r"C:\Users\espar\Desktop\Portafolio\TDCBanksMx\Data\040_12e_R1Distribución de tarjetas por límite de crédito.xls"
# Ruta del libro maestro donde guardaremos la tabla
ruta_maestro = r"C:\Users\espar\Desktop\Portafolio\TDCBanksMx\LibroControl.xlsx"

# Coordenadas de los clicks (ejemplo, medir con mouseInfo)
coord_banco = (393, 363)
coord_fecha = (76, 409)

# -------------------------------
# ABRIR LIBRO CNBV
# -------------------------------
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True  # mantener Excel visible
wb_cnbv = excel.Workbooks.Open(ruta_cnbv)
ws_cnbv = wb_cnbv.Sheets(1)

# Esperar a que Excel cargue completamente
time.sleep(5)

# -------------------------------
# HACER CLICK EN BANCO Y FECHA
# -------------------------------
pyautogui.moveTo(coord_banco[0], coord_banco[1], duration=0.5)
pyautogui.click()
time.sleep(2)  # esperar que cargue

pyautogui.moveTo(coord_fecha[0], coord_fecha[1], duration=0.5)
pyautogui.click()
time.sleep(5)  # esperar que cargue la tabla

# -------------------------------
# COPIAR LA TABLA DE CNBV
# -------------------------------
# Aquí asumimos que la tabla está en rango A1:F20 (modificar según tu caso)
data = ws_cnbv.Range("A1:F20").Value

# Convertir a DataFrame
df = pd.DataFrame(data[1:], columns=data[0])

# -------------------------------
# GUARDAR EN LIBRO MAESTRO
# -------------------------------
try:
    wb_maestro = openpyxl.load_workbook(ruta_maestro)
    ws_maestro = wb_maestro.active
except FileNotFoundError:
    wb_maestro = openpyxl.Workbook()
    ws_maestro = wb_maestro.active

# Escribir datos en la hoja
for r_idx, row in enumerate(df.values, 2):  # comienza en fila 2 para encabezado
    for c_idx, value in enumerate(row, 1):
        ws_maestro.cell(row=r_idx, column=c_idx, value=value)

# Encabezados
for c_idx, header in enumerate(df.columns, 1):
    ws_maestro.cell(row=1, column=c_idx, value=header)

# Guardar libro maestro
wb_maestro.save(ruta_maestro)

# -------------------------------
# CERRAR LIBRO CNBV
# -------------------------------
wb_cnbv.Close(SaveChanges=False)
excel.Quit()

print("Tabla copiada correctamente al libro maestro.")
