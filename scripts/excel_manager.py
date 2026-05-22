import logging
logger=logging.getLogger("ExcelManager")

from dataclasses import dataclass
import win32com.client as win32
from pathlib import Path
from enum import Enum
import win32gui, win32con
from config import RegexpPatrones, TipoPatron
from typing import Any

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


