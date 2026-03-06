
from pathlib import Path
from config import Paths
import os



RUTA_BASE = Path(__file__).parent.parent
RUTA_PRIME=RUTA_BASE/"Routes"
RUTA_DATA = RUTA_BASE  / "Data"
RUTAS_DATA_TDC=RUTA_DATA/"datosTarjetasCredito"
RUTA_OUTPUT = RUTA_BASE / "Output"
RUTA_DATASET=RUTA_BASE/"Datasets"
RUTA_LIBRO_CONTROL=RUTA_OUTPUT/ "LibroControl.xlsx"
# ruta.is_file() comprueba que sea un archivo (no carpeta).
# ruta.suffix devuelve la extensión, que comparamos con ['.xlsx', '.xls'].
# iterdir() recorre todos los elementos dentro de la carpeta.
# 1. Crear el directorio de forma segura
#ruta_carpeta.mkdir(parents=True, exist_ok=True)
#.iterdir() lista todo lo que está dentro de la carpeta (archivos y subcarpetas).
#.is_dir() filtra solo las carpetas.

# ARCHIVOS_POR_CARPETA = [
#     [f for f in CARPETAS_OUTPUT.iterdir() if f.is_file()]
#     for CARPETAS_OUTPUT in CARPETAS_OUTPUT
# ]
def rutas_de_archivos_en_carpeta(rutacarpeta):
    mat=[]
    for i in rutacarpeta.iterdir():
        if i.is_file():
            mat.append(i)
    return mat
#Compresion de listas
# "Voy a recorrer cada ruta que encuentre en RUTA_OUTPUT.iterdir()..."
# "...y por cada una de ellas, voy a fabricar un diccionario con estas 3 llaves..."
# "...y el resultado final de todos esos diccionarios se guardará automáticamente en una lista llamada RUTAS_OUTPUT_BRONZE."
RUTAS_Y_ARCHIVOS_OUTPUT_BRONZE=[
    {"proceso":ruta.name,"ruta":ruta,"archivos":list(ruta.glob("*.xls*"))}
    for ruta in RUTA_OUTPUT.iterdir()
]

#ARCHIVOS_DATA_RAW={str(RUTAS_DATA_TDC):[file for file in RUTAS_DATA_TDC.rglob("*.xls")]}

ARCHIVOS_DATA_RAW=[
    #Archivos es un objeto Path
    {"origen":archivos.name,"ruta":archivos}
    for archivos in RUTAS_DATA_TDC.iterdir()


]




def limpiar_texto(texto):
    """
    Reemplaza vocales acentuadas (con tilde, diéresis, etc.) 
    por sus equivalentes simples.
    """
    # 1. Definir el mapeo de caracteres a reemplazar
    mapeo = {
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U',
        'ü': 'u', 'Ü': 'U', 'ý': 'y', 'Ý': 'Y' 
    }
    
    # 2. Iterar sobre el mapeo y aplicar el reemplazo
    texto_limpio = texto
    for acentuada, simple in mapeo.items():
        texto_limpio = texto_limpio.replace(acentuada, simple)
        
    return texto_limpio.replace(' ', '_').strip().replace("\n","").replace("\r","")
#genera una lista de nombres sin la extension de los archivos contenidos en una carpeta
def obtener_lista_contenidas_en_carpetas(carpeta):
    lista=[]
    for i in carpeta.iterdir():
        
        if i.is_file():
            x=limpiar_texto(Path(i).stem)
            lista.append(x)
    return lista
def generar_carpeta(ruta_base, nombre_carpeta):

    ruta_completa = ruta_base / nombre_carpeta
    ruta_completa.mkdir(parents=True, exist_ok=True)
    return ruta_completa
def generar_carpetas_de_una_lista_en_una_ruta(ruta,lista):
    lista_de_rutas=[]    
    for i in lista:

        x=ruta/i
        lista_de_rutas.append(x)
        Path(x).mkdir(parents=True, exist_ok=True)
    return lista_de_rutas




def fix_long_path(ruta: str) -> str:

    if os.name == 'nt':  # Solo aplica en Windows
        # Normaliza la ruta y elimina comillas si las hubiera
        ruta = os.path.normpath(ruta.strip('"'))
        # Agrega prefijo solo si aún no lo tiene
        if not ruta.startswith('\\\\?\\'):
            ruta = '\\\\?\\' + ruta
    return ruta
