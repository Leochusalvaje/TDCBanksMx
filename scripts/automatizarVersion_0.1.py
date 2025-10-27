import pyautogui
import time
import pandas as pd
import openpyxl
import win32com.client as win32
import pygetwindow as gw
import pyautogui
import time
import os
import win32gui, win32con
# -------------------------------
# CONFIGURACIÓN
# -------------------------------
# Rutas
ruta_base = os.path.dirname(os.path.dirname(__file__))  # sube un nivel desde scripts
BotonObtener = os.path.join(ruta_base, "Images","Get.png")
ruta_cnbv=os.path.join(ruta_base,"Data","040_12e_R1Distribución de tarjetas por límite de crédito.xls")
ruta_maestro = os.path.join(ruta_base,"Output", "LibroControl.xlsx")
ruta_Output = os.path.join(ruta_base,"Output")
ruta_data=os.path.join(ruta_base,"Data")
# Coordenadas de los clicks (ejemplo, medir con mouseInfo)
coord_abajo=(362,425)
coord_arriba=(362,388)
coord_banco = (393, 363)
coord_obtener = (76, 409)
coord_desmarcar=(155,387)
coord_fecha=(155,423)
coord_ultima=(155,385)
coord_penultima=(156,404)

#Funciones auxialiares
def activar_ventana_al_frente(hwnd, timeout=5):
    start = time.time()
    while time.time() - start < timeout:
        if hwnd:
            # 💡 CAMBIO CLAVE: Usamos SW_MAXIMIZE en lugar de SW_RESTORE
            # SW_MAXIMIZE (3) maximiza la ventana.
            # SW_RESTORE (9) la restaura a su tamaño anterior (si estaba minimizada).
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE) 
            
            # Traer al frente
            win32gui.SetForegroundWindow(hwnd) 
            break
        time.sleep(0.2)
    else:
        print(f"No se pudo activar/maximizar la ventana con hwnd {hwnd}")
def abrir_excel(ruta):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(ruta)
    excel.Visible = True 

    # 🌟 PASO CRUCIAL 1: Dale un momento a Windows para que dibuje la ventana.
    time.sleep(0.1) 
    
    # 🌟 PASO CRUCIAL 2: Obtener el Handle (hwnd) del objeto Excel.Application
    # La propiedad Hwnd está disponible en el objeto Application.
    hwnd_excel = excel.Hwnd
    
    # 🌟 PASO CRUCIAL 3: Llamar a la función con el Handle, no con el nombre.
    activar_ventana_al_frente(hwnd_excel)
    
    wb.Sheets(1).Activate()
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
def marcarSiguienteFecha(n):
    for i in range(n):
        moveToClickAndWait(coord_abajo[0],coord_abajo[1],0,0)
    moveToClickAndWait(coord_fecha[0],coord_fecha[1],0,0)
    moveToClickAndWait(coord_arriba[0],coord_arriba[1],0,0)
    moveToClickAndWait(coord_fecha[0],coord_fecha[1],0,0)
    moveToClickAndWait(coord_obtener[0],coord_obtener[1],0,0)
# Lasiguiente funcion entrega dos ruats una del archivo origen y genra la carpeta destino de la limpieza
def generar_ruta_destino(nombre_archivo_cnbv_dentro_data):
    """
    Crea una carpeta dentro de Output basada en el nombre del archivo CNBV.
    Devuelve la ruta completa donde se guardará el nuevo archivo Excel.
    """
    # Nombre base sin extensión (para usarlo como nombre de carpeta)
    nombre_base = os.path.splitext(nombre_archivo_cnbv_dentro_data)[0]

    # Crear subcarpeta dentro de Output
    carpeta_destino = os.path.join(ruta_Output, nombre_base)
    os.makedirs(carpeta_destino, exist_ok=True)

    # Ruta final del nuevo archivo
    return os.path.join(carpeta_destino),os.path.join(ruta_data,nombre_archivo_cnbv_dentro_data),nombre_archivo_cnbv_dentro_data
#################################
# -------------------------------
# ABRIR LIBRO CNBV
# -------------------------------

def limpieza_datos_por_archivo(rutaArchivoParaLimpiar,rutadestino,NumerodeBimestresEnArchivo,nombreDocumento):

        
    excel, wb_cnbv = abrir_excel(rutaArchivoParaLimpiar)
    ws_cnbv = wb_cnbv.Sheets(1)




    def traer_tabla(origen):

        #se seleciona B9 para poder seleccioanr toda la tabla ya que ahi inicia y esto no cambia
        data = origen.Range("B9").CurrentRegion.Value
        df = pd.DataFrame(data[1:], columns=data[0])
        #La siguiente linea seleciona el origen que va a tomar el nombre de los archivos aqui debes selecioanr una celda en la que siempre haya la fecha de la consulta
        #para los arhcivos de uan tablas es C9 y los demas es c10
        nombre_extra = str(int(origen.Range("C10").Value))
        nombre_documento_origen=nombreDocumento
        # Crear el nombre del archivo
        nombre_archivo = f"{nombre_documento_origen}_{nombre_extra}.xlsx"

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

        nuevo_archivo = os.path.join(rutadestino, nombre_archivo)

        # Guardar libro maestro
        wb_maestro.save(nuevo_archivo)

        # -------------------------------
        # CERRAR LIBRO CNBV
        # -------------------------------
        activar_ventana_al_frente(excel.Hwnd)

        print("Tabla copiada correctamente al libro maestro.")


    max_iteraciones = NumerodeBimestresEnArchivo
    contador = 0




    while contador < max_iteraciones:
            valor_actual = ws_cnbv.Range("D37").Value

                # Aquí llamas a tu función de clicks, por ejemplo:
            if contador==0:
                print("Celda vacía, ejecutando clicks...")  
                moveToClickAndWait(coord_desmarcar[0],coord_desmarcar[1],0,0)
                #algunos archivos de la cnbv no es necesario marcar y desmarcar la casilal de bancos para selccioanr todos basta con un solo marcado
                #agregar o quitar la siguiente linea segun sea el caso
                moveToClickAndWait(coord_banco[0],coord_banco[1],0,0)
                moveToClickAndWait(coord_banco[0],coord_banco[1],2,0)
                moveToClickAndWait(coord_abajo[0],coord_abajo[1],0,6)
                moveToClickAndWait(coord_fecha[0],coord_fecha[1],0,0)
                moveToClickAndWait(coord_obtener[0],coord_obtener[1],0,0)

            elif contador<=max_iteraciones-3:
                print(f"Celda cambió: {valor_actual}")
                # Continuar con siguiente acción o rango
                traer_tabla(ws_cnbv)
                marcarSiguienteFecha(20-contador)
            else:
                traer_tabla(ws_cnbv)
                moveToClickAndWait(coord_fecha[0],coord_fecha[1],0,0)
                moveToClickAndWait(coord_penultima[0],coord_penultima[1],0,0)
                moveToClickAndWait(coord_obtener[0],coord_obtener[1],0,0)
                traer_tabla(ws_cnbv)
                moveToClickAndWait(coord_penultima[0],coord_penultima[1],0,0)
                moveToClickAndWait(coord_ultima[0],coord_ultima[1],0,0)
                moveToClickAndWait(coord_obtener[0],coord_obtener[1],0,0)
                traer_tabla(ws_cnbv)
                break
            contador += 1
            time.sleep(1)  # espera 1 segund
#solo hay que cambiar el nombre del archivo que hay en la carpe4ta Data para los casos donde solo hay un cuadro activex debes modificar la celda como esta documentado previamente
pr=generar_ruta_destino("040_12e_R50Distribución de tarjetas por porcentaje de pago mínimo respecto a la línea de crédito")
print(pr[1])
print(pr[0])
limpieza_datos_por_archivo(pr[1],pr[0],85,pr[2])