# INICIACIÓN DEL LOGGING
# ------------------------------------------------------------------------------
# Para que el logging capture correctamente los eventos de todos los módulos
# (especialmente decoradores de clase como @crear_directorios y configuraciones
# que se ejecutan al importar), la configuración de logging.basicConfig DEBE ser
# lo PRIMERO que se ejecute en el script principal, antes de los imports.
#
# Se usa 'force=True' para asegurar que esta configuración sea la mande mayor jerarquia y 
# no sea bloqueada por imports previos o librerías externas.
# ==============================================================================
from __future__ import annotations
import logging
logger=logging.getLogger("Automata")
logger.propagate = False
logger.handlers = [] # Limpia manejadores previos si los hay
logger.propagate = False 

# 3. Creas un manejador de consola (StreamHandler)
console_handler = logging.StreamHandler()

# 4. Le das un formato amigable para identificar tus pruebas
formatter = logging.Formatter('--- TEST: %(message)s ---')
console_handler.setFormatter(formatter)

# 5. Lo vinculas
logger.addHandler(console_handler)
logger.setLevel(logging.DEBUG)
from dataclasses import dataclass
import win32com.client as win32
import pyautogui
import time
import win32gui, win32con
import routes as gv
from pathlib import Path
import pywintypes as wty
import re
from typing import Any,TYPE_CHECKING
from config import RegexpPatrones, TipoPatron
from enum import Enum
if TYPE_CHECKING:
    from config import Paths,DescripcionArchivosCriticos,RegexpPatrones,ConfigManager

pyautogui.FAILSAFE = True


CONFIG_COORDS = {
    "abajo": (362, 425),
    "arriba": (362, 388), 
    "bancos": (393, 363),
    "obtener": (76, 409), 
    "fecha_top": (155, 387),
    "fecha_bottom": (155, 423),
    "penultima": (156, 404),
    "tiempo_espera_consulta": 1,
    "tiempo_espera_entre_clicks":0.3
}

coord_abajo=(362,425)
coord_arriba=(362,388)
coord_banco = (393, 363)
coord_obtener = (76, 409)
coord_desmarcar=(155,387)
coord_fecha=(155,423)
coord_ultima=(155,385)
coord_penultima=(156,404)

#Funciones auxialiares
def activar_ventana_al_frente(hwnd, timeout=5)-> None:
    start = time.time()
    while time.time() - start < timeout:
        if hwnd:

            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)             
            win32gui.SetForegroundWindow(hwnd) 
            logger.info(f"Ventana {hwnd} traida al frente y maximizada con exito")
            break
        time.sleep(0.2)
    else:
        logger.critical(f"No se pudo activar/maximizar la ventana con hwnd {hwnd}, asgeurar que no se esten ejecutando procesos externos durante al ejecucion")


def extraer_datos(celda:Any,rutadestino:Path,ventanaexcel)->None:
    rango:Any=celda.CurrentRegion()
    rutadellibrocontrol:Path=gv.Paths.output/"LibroControl.xlsx"
    nombreventana:str=rutadellibrocontrol.stem
    try:
        libroabierto=ventanaexcel.Workbooks(nombreventana)
        libroabierto.Close(SaveChanges=False)
        print("Se cierra la ventana previa del LibroControl si es que esta abierto para evitar inyectar informacion erronea")
    except:
        pass


    librocontrol=ventanaexcel.Workbooks.Open(str(rutadellibrocontrol))
    librocontrol.Activate()

    filas=len(rango)
    columnas=len(rango[0])
    hoja=librocontrol.Sheets(1)

    print(type(rango))
    celdas_destino=hoja.Range(hoja.Cells(1, 1),hoja.Cells(filas, columnas))
    celdas_destino.Value=rango
    try:
        ventanaexcel.DisplayAlerts = False
        librocontrol.SaveAs(str(rutadestino))
        librocontrol.Close(SaveChanges=False)
    except Exception as e:
        print(f"Error en el metodo SaveAs , no se pudo guardar {type(e).__name__}: {e}\n")
    finally:
            if librocontrol is not None:
                try:
                    ventanaexcel.DisplayAlerts = False 
                    librocontrol.Close(SaveChanges=False)
                except:
                    ventanaexcel.DisplayAlerts = True
                    pass     
    return None

def string_eureka_fechas(libro:Any,patronValidador:RegexpPatrones)->tuple:
    
    primerhoja=libro.Sheets(1)
    primerhoja.Activate()

    buscar_1=primerhoja.Range("B1").End(-4121).End(-4121)

    if not primerhoja.Range("B1").End(-4121).Value=="Notas":
        raise Exception("\nCambio el formato de la CNBV actualizar saltos de linea y espaciados de nuevo para encontrar la tabla, se debe modificar los pasos de la libreria win32com.client.")

    a=buscar_1.CurrentRegion
    a.Activate()
    tupla_datos=a.Value

    if not len(tupla_datos[0])>=2:
        raise Exception("\nEs muy probable que cambio la formato de la CNBV pues no se esta alcanzando la tabla de datos para su extraccion.")
    try:
    #celda=primerhoja.Cell
        str_fecha=str(int(tupla_datos[0][2]))

        print("eureka")        
        print("Formato estandar detectado no se necesito aplicar logica adicional")
    except (IndexError, ValueError, TypeError):
        str_fecha = str(int(tupla_datos[1][2]))
        print("Formato alternativo detectado")

    if not re.match(patronValidador.patrones["fecha"],str_fecha):
        raise Exception("\nEl valor para buscar la fecha y nombrar las carpetas no se cumple. Validar la celda de donde se esta extrayendo la fecha.")


    str_fecha=str_fecha
    celda_tabla=buscar_1
    hayMasDeUnBancoEnLaHoja=len(tupla_datos[0])>3

    return str_fecha,celda_tabla,hayMasDeUnBancoEnLaHoja


def lista_verdad(bimestreInicial:str,bimestreFinal,patronValidador:RegexpPatrones)->list:
    
    fecha=patronValidador.revisar_patron(TipoPatron.FECHA,bimestreFinal)

    try:
        yearbase=int(bimestreInicial[:4])
        yearUltimobimestre=int(fecha[:4])
        yearsTranscurridos=yearUltimobimestre-yearbase
        bimestreActual=(14-int(fecha[-2:]))/2
    except (ValueError, TypeError) as e:
        # Usamos .exception para que guarde el rastro (traceback)
        # Y metemos type(e).__name__ para que el mensaje claro
        logger.exception(f"[{type(e).__name__}] Error fatal en fechas: {bimestreInicial} y {bimestreFinal} revisar que se esten ingresando como strings. Detalle: {e}")
        # Esto es lo que hace que el programa "muera" o falle rápido
        raise

    bimestreBase=int(bimestreInicial[-2:])/2
    bimestresEnArchivo=int(6*(yearsTranscurridos-1)+(bimestreBase+bimestreActual))
    logger.info(f"Bimestres contenidos en cada archivo excel de la CNBV: {bimestresEnArchivo} Este parametro es crucial y se genera con los parametros del config.json ingresados por el USUARIO humano para el funcionamiento del automata, si no se esta extrayendo la cantidad correcta de bimestres revisar que las fechas de entrada sean correctas y que se este cumpliendo el patron establecido en la clase RegexpPatrones")
    #para resolver el seguimiento de fechas se solcuiona con una transformacion lineal que planteo su servidor y tener una lista de todos los bimestres 
    #que estan contenidos en el archivo excel al momento de abrirlo
    y=[]
    yearconstante=int(bimestreInicial)
    c=0
    for _ in range(1,bimestresEnArchivo+1):

        yearconstante+=2    
        if (yearconstante-201100-100*c)==14:
            c+=1 
            yearconstante+=100
            yearconstante-=12

        y.append(yearconstante)
    y.reverse()


    return y

class rutaCarpetas:
    def __init__(self,ruta_carpeta_output:Path,archivo_excel:Path,patronvalidador:RegexpPatrones)->None:
        self.rutaArchivoExcel=archivo_excel
        self.nombreArchivExcel=archivo_excel.stem
        self.nombreSerie=patronvalidador.busqueda_patron_para_nombrar_carpetas(self.nombreArchivExcel)
        self.carpetaDestino=ruta_carpeta_output / self.nombreSerie
        self.carpetaDestino.mkdir(parents=True, exist_ok=True)
    
    @property
    def revisar_archivos_en_carpeta(self)->list:
        listaarchivos=[]
        for archivos in self.carpetaDestino.iterdir():
            if archivos.is_file():
                listaarchivos.append(archivos.stem)
        return listaarchivos
@dataclass  
class tareaPendiente:
    rutaArchivoExcel:Path
    rutaCarpetaDestino:Path
    nombreSerie:str
    listaPendientes:list

    def __repr__(self):
        return (f"Serie: {self.nombreSerie} | "
                f"Archivo: {self.rutaArchivoExcel.name} | "
                f"Pendientes: {len(self.listaPendientes)}")

class carajo:
    def __init__(self,mapa:rutaCarpetas,listaFechas:list)->None: 
        #La lista de los bimestres contenidos en el archivo excel
        self.listaFechas=listaFechas
        #uso de la clase rutaCarpetas
        self.mapa=mapa
        self.archivosCreados=mapa.revisar_archivos_en_carpeta
    #Con la lista de fechas y el nombre de la serie se crea una lista de los archivos que deberian estar creados al finalizar la consulta, esto es crucial para que el vigia pueda revisar cuales archivos faltan y generar las tareas pendientes para el automata
    @property
    def dummy_archivos(self)->list:
        return [f"{self.mapa.nombreSerie}_{i}" for i in self.listaFechas]       

    
    #El vigia va a revisar los pendientes que hay en la carpeta
    @property    
    def pendientes_de_consulta(self)->list:
        pendientes=set(self.dummy_archivos)-set(self.archivosCreados)
        return list(pendientes)
    
    #Esta funcion ejecuta en tiempo real la tarea pendiente que se genero con la propiedad anterior, esto es crucial para 
    #que el automata pueda ejecutar las tareas pendientes sin necesidad de que el USUARIO humano este revisando constantemente 
    #la carpeta de destino y generando las tareas pendientes manualmente, ademas de que al generar la tarea pendiente en tiempo real
    #se evita que se generen tareas pendientes erroneas por archivos que se crean despues de revisar los pendientes pero antes de 
    #ejecutar el automata, lo que podria generar errores o confusiones en el proceso de consulta.
    @property    
    def mapa_de_pendientes(self):
        indices_pendientes=[]
        for k, i in enumerate(self.dummy_archivos):
            for j in self.pendientes_de_consulta:
                if i==j:
                    indices_pendientes.append(k)
        nuevaTarea=tareaPendiente(rutaArchivoExcel=self.mapa.rutaArchivoExcel,
                    rutaCarpetaDestino=self.mapa.carpetaDestino,
                    nombreSerie=self.mapa.nombreSerie,
                    listaPendientes=indices_pendientes)
        logger.info(f"Tarea lista para ejecución: {nuevaTarea}")

        return nuevaTarea
#Esta clase servira para saber si la cordenada dentro del checkbox esta activa o no, y nos porporcionara un boton mental


class TipoLibro(Enum):
    LIBRO_CONTROL ="libro_control"
    LIBRO_CONSULTA ="libro_consulta"

class libroExcelManager:
    def __init__(self,libro:Any,patronValidador:RegexpPatrones,posicion_hoja_excel:int=1) -> None:

        self.libro=libro
        #Dejo la propeidad de hoja como opcional por si en algun momento se necesita acceder a otra hoja del libro, 
        #aunque por el formato estandar de la CNBV no es necesario y se puede acceder a la hoja directamente con el indice 1
        self.hoja_Excel=self.libro.Sheets(posicion_hoja_excel)
        self.patronValidador=patronValidador
        
    def __getattr__(self, nombre):
        # Si se solicita un metodo que LibroExcelManager no tiene, automáticamente lo tomara de self.libro que es un objeto COM
        return getattr(self.libro, nombre)

    @property
    def funcion_eureka_buscar_tabla(self)->tuple:

        primerhoja=self.hoja_Excel
        primerhoja.Activate()
        if not primerhoja.Range("B1").End(-4121).Value=="Notas":
            raise Exception("\nCambio el formato de la CNBV actualizar saltos de linea y espaciados de nuevo para encontrar la tabla, se debe modificar los pasos de la libreria win32com.client.")
        celda_vigia=primerhoja.Range("B1").End(-4121).End(-4121)

        #Esta linea solo sirve para el debugger revisar que se esta alcanzando la tabla de datos
        #celda_vigia.CurrentRegion.Activate() 
        
        return celda_vigia.CurrentRegion.Value
    
    @property
    def numero_columnas_tabla(self)->int:
        return len(self.funcion_eureka_buscar_tabla[0])
    
    @property
    def fecha_en_tabla(self)->str:
        tabla=self.funcion_eureka_buscar_tabla
        try:
        #celda=primerhoja.Cell
            str_fecha=str(int(tabla[0][2]))
            logger.debug("Formato estandar detectado no se necesito aplicar logica adicional")
        except (IndexError, ValueError, TypeError):
            str_fecha = str(int(tabla[1][2]))
            logger.debug("Formato alternativo detectado")
            

        if not self.patronValidador.revisar_patron(TipoPatron.FECHA,str_fecha):
            raise Exception("\nEl valor para buscar la fecha y nombrar las carpetas no se cumple. Validar la celda de donde se esta extrayendo la fecha.")
        logger.debug("Patrón de fecha válido, se puede proceder")

        return str_fecha


class ExcelManager: 
    def __init__(self,patron_validador:RegexpPatrones,libro_control_ruta:Path) -> None:
        self.patron_validador=patron_validador
        self.libro_control_ruta=libro_control_ruta
        self.ventana_excel = None
        self.hwnd = None  
        self.libro:dict[TipoLibro,libroExcelManager]={} 



    def _abrir_instancia_excel(self) -> None:
        try:
            self.ventana_excel = win32.Dispatch("Excel.Application")
            self.hwnd = self.ventana_excel.Hwnd
        except Exception as e:
            logger.critical(f"Error al abrir instancia de excel validar permisos o conflictos con la libreria win32.client: {e}")
            raise

    def traer_instancia_excel_al_frente(self)-> None:
        if self.hwnd is None:
            logger.warning("Se intenta traer una ventana de excel que no esta instanciada, se invocara el metodo \"_abrir_instancia_excel\"")
            self._abrir_instancia_excel()
        try:
            self.ventana_excel.Visible = True
            win32gui.ShowWindow(self.hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(self.hwnd)
            logger.info("Ventana de Excel abierta y traída al frente")
        except Exception as e:
            logger.error(f"Error forzando traer ventana al frente: {e}")
            raise        

    def traer_libro_al_frente(self,tipo_libro:TipoLibro) -> None:
        if self.libro is not None:
            try:
                self.libro[tipo_libro].Activate()
                logger.info(f"Libro {tipo_libro} traído al frente")
            except Exception as e:
                logger.error(f"Error al traer el libro al frente: {e}")
            

    def abrir_libro(self, tipo_libro:TipoLibro, ruta_archivo: Path) -> None:
        logger.info(f"Abriendo archivo Excel {ruta_archivo.stem}")
        
        if not ruta_archivo.exists():
            logger.error(f"El archivo {ruta_archivo} no existe.")
            raise FileNotFoundError(f"Archivo no encontrado: {ruta_archivo}")

        if self.ventana_excel is None:
            self._abrir_instancia_excel()
        
        if tipo_libro is TipoLibro.LIBRO_CONTROL:
            nuevo_libro=self.ventana_excel.Woorbooks.Open(str(self.libro_control_ruta))
            self.libro[tipo_libro]=libroExcelManager(nuevo_libro,self.patron_validador)
            
        else:
            nuevo_libro=self.ventana_excel.Workbooks.Open(str(ruta_archivo))
            self.libro[tipo_libro]=libroExcelManager(nuevo_libro,self.patron_validador)
        


    def cerrar_libro_excel(self, nombre_libro: TipoLibro, ruta_guardar: Path | None = None) -> None:
        

        #Cierra un libro específico de Excel manejando la seguridad de macros y alertas.
        #Si se proporciona ruta_guardar, realiza un SaveAs antes de cerrar.

        if not self.ventana_excel or not self.libro:
            logger.info("No hay sesión de Excel activa para cerrar.") 
            return

        try:
            # Configuración de seguridad para el guardado silencioso
            self.ventana_excel.EnableEvents = False
            self.ventana_excel.DisplayAlerts = False
            
            libro_com = self.libro[nombre_libro]

            # Fix para Office 365: El autoguardado en la nube suele bloquear el SaveAs
            try:
                libro_com.AutoSaveOn = False 
            except:
                pass 
            
            if ruta_guardar:
                libro_com.Activate()
                # resolve() asegura rutas absolutas evitando errores de permisos en Windows
                ruta_final = str(ruta_guardar.resolve())
                libro_com.SaveAs(ruta_final)
                logger.info(f"Libro guardado en: {ruta_final}")

            libro_com.Close(SaveChanges=False)
            del self.libro[nombre_libro]

        except Exception as e:
            logger.error(f"Fallo al cerrar/guardar el libro '{nombre_libro.name}': {e}")
        
        finally:
            # Restaurar estado de la aplicación siempre
            if self.ventana_excel:
                self.ventana_excel.EnableEvents = True
                self.ventana_excel.DisplayAlerts = True

    def extraer_tabla_y_guardar_en_ruta(self,ruta_destino:Path,libro_control_ruta:Path):
        libro_c=TipoLibro.LIBRO_CONTROL
        #abrimos y ponemos el libro de extraccion en una varible 
        self.abrir_libro(libro_c,libro_control_ruta)
        libro_control=self.libro[libro_c]

        try:
            libro_consultas=self.libro[TipoLibro.LIBRO_CONSULTA]
        except KeyError:
            logger.critical("No se ha abierto el libro de consulta, no se puede extraer la tabla. Asegurar que se ha abierto correctamente con el metodo \"abrir_libro\" y que se esta pasando la ruta correcta del libro de consulta.")
            raise

        tabla=libro_consultas.funcion_eureka_buscar_tabla
        filas=len(tabla)
        columnas=len(tabla[0])
        
        hoja=libro_control.hoja_Excel
        celdas_destino=hoja.Range(hoja.Cells(1, 1),hoja.Cells(filas, columnas))
        celdas_destino.Value=tabla
        self.cerrar_libro_excel(TipoLibro.LIBRO_CONTROL, ruta_destino)
"""
Esta clases se hizo para resolver el problema de que el automata no tiene forma de saber si una cordenada dentro del checkbox
esta activa o no, con esta clase se puede simular un boton mental que cambia su estado cada vez que se hace click en la cordenada del checkbox,
de esta forma el automata puede saber si la cordenada esta activa o no y tomar decisiones en base a eso.
"""
class boton:
    def __init__(self,fecha:int,posicion:int)->None:
        self.fecha=fecha
        self.posicion=posicion
        self.estado=False
    @property
    def estado_actual(self)->bool:
        return self.estado
    
    def cambiar_estado(self)->None:
        self.estado=not self.estado
"""
coordenadaSimple es una clase que usa la libreria pyautogui para hacer click en una coordenada especifica,
esta clase se hizo para resolver el problema de que el automata no tiene forma de saber si una cordenada dentro
del checkbox y tarbajar con la coordenada de obtener esta activa o no.
"""
class coordenadaSimple:

    def __init__(self,coordenada:tuple,intervalo:int)->None:
        self.coordenada=coordenada
        self.intervalo=intervalo 

    def click_en_coordenada(self,boton:boton|None=None)->None:
        pyautogui.click(self.coordenada,interval=self.intervalo)
        if boton:
            boton.cambiar_estado()

@dataclass
class CheckerCoordenada:
    abajo: tuple
    arriba: tuple
    bancos: tuple
    obtener: tuple
    fecha_top: tuple
    fecha_bottom: tuple
    penultima: tuple

    def __post_init__(self):
        self.abajo = tuple(self.abajo)
        self.arriba = tuple(self.arriba)
        self.bancos = tuple(self.bancos)
        self.obtener = tuple(self.obtener)
        self.fecha_top = tuple(self.fecha_top)
        self.fecha_bottom = tuple(self.fecha_bottom)
        self.penultima = tuple(self.penultima)

@dataclass
class TiemposConfig:
    tiempo_espera_consulta: float
    tiempo_espera_entre_clicks: float

@dataclass
class ConfigVisual:
    coords: CheckerCoordenada
    tiempos: TiemposConfig


class pointCheckboxManager(coordenadaSimple):

    def __init__(self,config_visual:ConfigVisual)->None:
        self._despachador=config_visual

        super().__init__(self._despachador.coords.fecha_top, self._despachador.tiempos.tiempo_espera_entre_clicks)
        
    def click_auxiliar_en_coordenada(self,punto:tuple)->None:
        pyautogui.click(punto,interval=self._despachador.tiempos.tiempo_espera_entre_clicks)
        time.sleep(self._despachador.tiempos.tiempo_espera_consulta)   

    def siguiente_cordenada(self,pasos:int):
        pyautogui.click(self._despachador.coords.abajo,clicks=abs(pasos),interval=self._despachador.tiempos.tiempo_espera_entre_clicks)

    def  anterior_cordenada(self,pasos:int):
        pyautogui.click(self._despachador.coords.arriba,clicks=abs(pasos),interval=self._despachador.tiempos.tiempo_espera_entre_clicks)


    def obtener(self)->bool:
        pyautogui.click(self._despachador.coords.obtener)
        time.sleep(self._despachador.tiempos.tiempo_espera_consulta)

class automataPrime:
    def __init__(self,tarea:tareaPendiente,manipulacion:ExcelManager)->None:
        self.tarea=tarea
        self.manipualcion=manipulacion
        self.matrizBotones:dict[int,boton]={}

class automataConsultas:
    def __init__(self,punto_actual:pointCheckboxManager,obervador:carajo,ventanaExcel:Any)->None:
        #Esta es una ventana y un objeto para manipualr el excel posterior a esto si obtendremos el handle
        self.ventanaExcel=ventanaExcel
        self.hdwm_excel=ventanaExcel.Hwnd
        self.observador=obervador
        #iniciamos la matriz de fechas que hay dentro del archivo excel que se esta consultando
        self.matrizFechas=self.observador.listaFechas
        #con las fechas creamos botones con todos inicializados en falso, osea desmarcados y solo el primero queda marcado tal y como esta el estandar
        #del archivo excel, la fecha mas reciente esta marcada y consultada
        self.matrizBotones:dict[int,boton]={}
        self._navegante=punto_actual

        self._propiocepcion = 0
        #obtenemos la hoja excel del observador
        self.hojaExcel=self.observador.hojaExcel

        self.propiocepcion_actual

        try:
            self.crear_botones()
            # Como la primer fecha siemproe esta marcada el primer boton queda activo tal como los archivos
            #accedemeos manualmenete al dictioanro con el metodo iter para ir uno a uno y para el primero encapsualmos en un next
            # se implemnta nueva busqueda por la hoja excel en vez de inicializar las variables, no descarto la inicializacion de variables de momento se quitara
            primer_boton = next(iter(self.matrizBotones.values()))
            primer_boton.estado = True
        except Exception as e:
            print(f"Error critico no se pudieron crear los botones para guiar al automata, revisar la propiedad \"self.matrizFechas\"\n{e}")
            self.matrizBotones={}

        #con lo que hacemos la propiocepcion es con la fecha que se obtiene al iniciar el automata
        self.ojos=string_eureka_fechas(self.hojaExcel)[0]
    
    def crear_botones(self):


        for k, fechas in enumerate(self.matrizFechas):
            x=boton(fechas,k)
            self.matrizBotones[fechas]=x
        
    def avanzar_a_fecha(self,pasos:int)->None:

        if pasos == 0:
            return

        if pasos>0:
            self._navegante.siguiente_cordenada(pasos)
        elif pasos<0:
            self._navegante.anterior_cordenada(pasos)

        self._propiocepcion+=pasos
        logger.debug("propiocepcion",self._propiocepcion)


    
    def ir_apagar_boton(self,boton:boton):

        if boton.estado:
            self.avanzar_a_fecha(boton.posicion-self._propiocepcion)
            self._navegante.click_en_coordenada(boton)
        
    def ir_apagar_todos_los_botones_encendidos(self)->None:
        for boton in self.matrizBotones.values():
            if boton.estado:
                self.ir_apagar_boton(boton)

    def invertir_estado_varios_botones(self,cambios:list)->None:
        for botonFecha in cambios:
            self.matrizBotones[botonFecha].cambiar_estado()

    @property            
    def botones_encendidos(self)->list:
        listaFechas=[]
        for boton in self.matrizBotones.values():
            if boton.estado:
                listaFechas.append(boton.fecha)
        return listaFechas

    @property
    def propiocepcion_actual(self)->int:
        print (self._propiocepcion)
        return self._propiocepcion

    def homming(self)->None:
        activar_ventana_al_frente(self.hdwm_excel)
        self.avanzar_a_fecha(-self._propiocepcion)
        self._navegante.click_en_coordenada()

    def validar_fecha_actual_y_extraer(self,fechaRevision:int,celda)->bool:

        try:
            
            print(f"fecha actual obtenida: {fechaRevision}, si coincide con la propiacepcion {self.matrizFechas[self._propiocepcion]}")
            print("fecha coincide con la actual, obtener")
            if not fechaRevision==self.matrizFechas[self._propiocepcion]: 
                print("Implemnetando nueva logica")

                return True
            else: 
                archivo=f"{self.observador.dummy_archivos[self._propiocepcion]}.xlsx"
                rutaArchivo=self.observador.mapa.carpetaDestino/archivo
                extraer_datos(celda,rutaArchivo,self.ventanaExcel)
                
                return False
            

        except IndexError as e:
            print(f"IndexError en obtener fecha actual, probablemente la propiocepcion se salio de rango No es un error critico continua {e}")
            return False
        

    def recorrer_pendientes(self,lisapendientes:list)->None:
        print(self.matrizFechas)
        print("lista de pendientes a consultar",lisapendientes)

        #Configuracion para marcar todos los bancos al inicio de la consulta
        self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["bancos"])
        self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["bancos"])
        self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["obtener"])
        #Manejo de caso atipico para marcar todos los bancos
        if not string_eureka_fechas(self.hojaExcel)[2]:
            self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["bancos"])


        cicloRoto=False
        posicion=0
        for k,indice in enumerate(lisapendientes):

            if k==0 and self.matrizFechas[k] not in [self.matrizFechas[i] for i in lisapendientes]:
                self.ir_apagar_boton(self.matrizBotones[self.matrizFechas[k]])


            self.avanzar_a_fecha(indice-self._propiocepcion)
            fechaEnBucle=self.matrizFechas[indice]
            
            #validamos que el boton no es te seleccionado  
            if not self.matrizBotones[fechaEnBucle].estado:
                self._navegante.click_en_coordenada(boton=self.matrizBotones[fechaEnBucle])


            self._navegante.obtener()
            fechaEnHoja,celda,*_=string_eureka_fechas(self.hojaExcel)
            fechaEnHoja=int(fechaEnHoja)
            

            if self.validar_fecha_actual_y_extraer(fechaEnHoja,celda):  
                cicloRoto=True


                ajustePropiocepcion=self.matrizBotones[fechaEnHoja].posicion-self._propiocepcion
                self._propiocepcion+=ajustePropiocepcion

                #cuando se desajusta para encontrar el boton que se encendio por error, es la diferencia entre el anterior indice menos el indice actual 
                #de la lista de pendientes en la depuracion de una lista ejemplo [3,7,11], pasos=7-3=4 que es el borton que se encendio por error 
                pasos=lisapendientes[k]-lisapendientes[k-1]
                self.invertir_estado_varios_botones([fechaEnHoja,fechaEnBucle,self.matrizFechas[lisapendientes[k-1]],self.matrizFechas[pasos]])
                self.ir_apagar_todos_los_botones_encendidos()

                print("entro al caso particular")
                break

            self.ir_apagar_boton(self.matrizBotones[fechaEnHoja])
            posicion+=1   
            print("propiocepcion despues de avanzar",self._propiocepcion)
        print("fecha de cambio",fechaEnHoja)

        if cicloRoto:
            fechaQueRompeElBucle=fechaEnHoja

            for indice in lisapendientes[posicion:]:
                if indice==len(self.matrizFechas)-2:
                    self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["penultima"])
                    self._navegante.obtener()
                    fechaEnHoja,celda,*_=string_eureka_fechas(self.hojaExcel)
                    fechaEnHoja=int(fechaEnHoja)
                    self.validar_fecha_actual_y_extraer(fechaEnHoja,celda) 
                    self._propiocepcion-=self.matrizBotones[fechaQueRompeElBucle].posicion+self._propiocepcion
                    continue
                if indice==len(self.matrizFechas)-1:
                    self.avanzar_a_fecha(indice-self._propiocepcion)
                    self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["penultima"])
                    self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["fecha_bottom"])             
                    self._navegante.obtener()
                    fechaEnHoja,celda,*_=string_eureka_fechas(self.hojaExcel)
                    fechaEnHoja=int(fechaEnHoja)
                    self.validar_fecha_actual_y_extraer(fechaEnHoja,celda) 
                    self._propiocepcion-=self.matrizBotones[fechaQueRompeElBucle].posicion+self._propiocepcion
                    continue                 
                fechaEnBucle=self.matrizFechas[indice]
                #Pruebas para ajustar la propiocepcion con ayuda del debugger , simular los clciks manualmente mientras se testea
                #print(indice)
                #print(self.observador.listaFechas[-self._propiocepcion])
                if indice==11:
                    print("para revisar con el debugger")
                # 201406, 201310, 201302,201206,
                self.avanzar_a_fecha(indice-self._propiocepcion)
                self._navegante.click_en_coordenada(boton=self.matrizBotones[fechaEnBucle])
                
                #print(self.observador.listaFechas[-self._propiocepcion])
                if indice<len(self.matrizFechas)-2:

                    self._navegante.obtener()
                    fechaEnHoja,celda,*_=string_eureka_fechas(self.hojaExcel)
                    fechaEnHoja=int(fechaEnHoja)
                    self.validar_fecha_actual_y_extraer(fechaEnHoja,celda) 
                    self._propiocepcion-=self.matrizBotones[fechaQueRompeElBucle].posicion+self._propiocepcion
                    
                    self.ir_apagar_boton(self.matrizBotones[fechaEnHoja])
        self.homming()
        self.observador.hojaExcel.Close(SaveChanges=False)
        self.ventanaExcel.DisplayAlerts = False
        self.ventanaExcel.Quit()  

        print("Proceso de consultas finalizado")

# pat=RegexpPatrones()
# rutaPrueba=gv.RUTA_BASE/ "scripts"/ "prueba.xlsx"

# x=gv.ARCHIVOS_DATA_RAW[3]["ruta"]
# prueba=ExcelManager(pat)
# prueba._abrir_instancia_excel()

# prueba.traer_instancia_excel_al_frente()
# prueba.traer_instancia_excel_al_frente()
# prueba.abrir_libro("libro1",x)
# libro=prueba.libro["libro1"]

# prueba.traer_libro_al_frente("libro1")


# prueba.extraer_tabla_y_guardar_en_ruta("libro1",rutaPrueba,gv.RUTA_LIBRO_CONTROL)

# prueba.cerrar_libro_excel("libro1")
# time.sleep(2)   
# prueba.ventana_excel.Quit()
# prueba.ventana_excel = None
# casilla1=pointCheckboxManager(CONFIG_COORDS)
# mapa=rutaCarpetas(rutasInstaciadas,x)

# libro,ventana,*_=abrir_excel_por_ruta(x)


# observador=carajo(libro,mapa)
# observador.archivosCreados

# hoja1=automataConsultas(casilla1,observador,ventana)



# hoja1.recorrer_pendientes([0])


# print(hoja1.matrizFechas)

# ################################### 
# Implementacion de decoradores
# def mi_primer_decorador(fechaBuscada:str):

#     def decorador_real(funcion):

#             def envoltura(*args, **kwargs):
#                 pass




