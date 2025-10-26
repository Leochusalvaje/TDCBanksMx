import pyautogui
import time
import pandas as pd
import openpyxl
import win32com.client as win32
import pygetwindow as gw
import pyautogui
import time
import os

# -------------------------------
# CONFIGURACIÓN
# -------------------------------
# Rutas
ruta_base = os.path.dirname(os.path.dirname(__file__))  # sube un nivel desde scripts
BotonObtener = os.path.join(ruta_base, "Images", "Get.png")
ruta_cnbv=os.path.join(ruta_base,"Data","040_12e_R1Distribución de tarjetas por límite de crédito.xls")
ruta_maestro = os.path.join(ruta_base, "LibroControl.xlsx")


# Coordenadas de los clicks (ejemplo, medir con mouseInfo)
coord_abajo=(362,425)
coord_arriba=(362,388)
coord_banco = (393, 363)
coord_obtener = (76, 409)
coord_desmarcar=(155,387)
coord_fecha=(155,423)

#Funciones auxialiares



def activar_ventana_al_frente(titulo, timeout=10):

    start_time = time.time()
    while True:
        ventanas = gw.getWindowsWithTitle(titulo)
        if ventanas:
            ventana = ventanas[0]
            ventana.activate()
            # Revisamos si la ventana está activa
            if ventana.isActive:
                break
        # Timeout
        if time.time() - start_time > timeout:
            print(f"No se pudo activar la ventana '{titulo}' en {timeout} segundos")
            break
        time.sleep(0.2)  # Espera corta antes de intentar de nuevo

def abrir_excel(ruta):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(ruta)
    excel.Visible = True   
    activar_ventana_al_frente(wb.Name)
     
    return excel, wb

def keepClick(t):
    pyautogui.mouseDown()
    time.sleep(t)
    pyautogui.mouseUp()


def moveToClickAndWait(x,y,t,tc):
    pyautogui.moveTo(x, y, t)
    keepClick(tc)
    time.sleep(t) 

def ClickByImages(rutaimagen):
    pos = pyautogui.locateOnScreen(rutaimagen)  # busca en toda la pantalla
    if pos:
        center_x, center_y = pyautogui.center(pos)
        pyautogui.click(center_x, center_y)



#################################
excel,wb=abrir_excel(ruta_cnbv)
# -------------------------------
# ABRIR LIBRO CNBV
# -------------------------------

time.sleep(5)


#hasta aqui todo funciona



wb_cnbv = excel.Workbooks.Open(ruta_cnbv)
ws_cnbv = wb_cnbv.Sheets(1)

# Esperar a que Excel cargue completamente

# -------------------------------
# HACER CLICK EN BANCO Y FECHA
# -------------------------------



valor_anterior = None

while True:
        valor_actual = ws_cnbv.Range("D10").Value

        if valor_actual is None or valor_actual == "":
            print("Celda vacía, ejecutando clicks...")
            
            
            # Aquí llamas a tu función de clicks, por ejemplo:
            moveToClickAndWait(coord_desmarcar[0],coord_desmarcar[1],2,0)
            moveToClickAndWait(coord_banco[0],coord_banco[1],2,0)
            moveToClickAndWait(coord_banco[0],coord_banco[1],2,0)
            moveToClickAndWait(coord_abajo[0],coord_abajo[1],2,6)
            moveToClickAndWait(coord_fecha[0],coord_fecha[1],2,0)
            
            ClickByImages(BotonObtener)





            # ClickEnCasilla(x, y)
        elif valor_actual != valor_anterior:
            print(f"Celda cambió: {valor_actual}")
            # Continuar con siguiente acción o rango
            break
    
        valor_anterior = valor_actual
        time.sleep(1)  # espera 2 segund



time.sleep(5)
data = ws_cnbv.Range("B9").CurrentRegion.Value
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