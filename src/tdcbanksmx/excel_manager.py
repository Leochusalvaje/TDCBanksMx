import logging
logger=logging.getLogger("ExcelManager")

from dataclasses import dataclass, field
import win32com.client as win32
from pathlib import Path
from enum import Enum
import win32gui, win32con
from tdcbanksmx.config import RegexpPatrones, TipoPatron
from typing import Any
import tdcbanksmx.exceptions as ex
import time





#solo en este esquema usaremos slots para optimizar el uso de memoria ya que esta clase se va a instanciar muchas veces para generar las tareas
#  pendientes, ademas de que al ser una clase de datos no necesitamos la flexibilidad que ofrecen los atributos dinamicos de python, 
# por lo que el uso de slots es adecuado para optimizar el rendimiento.
@dataclass(slots=True)
class Pendiente:
    """Representa una tarea pendiente que debe ser procesada.
    
    Contiene la información necesaria para identificar y rastrear
    una sub-tarea específica dentro del proceso de automatización.
    
    Attributes:
        fecha (int): Identificador o índice de la fecha/período al que pertenece la tarea.
        posicion (int): Posición o ubicación de la tarea dentro del archivo Excel.
        completado (bool): Estado de completación de la tarea.
    """
    
    fecha:int
    posicion:int
    completado:bool

@dataclass(slots=True)
class TareaPendiente:

    """Representa una tarea específica que el autómata debe ejecutar.

    Contiene la información necesaria para identificar el archivo Excel relacionado,
    las carpetas de destino y el desglose de los pendientes.

    Attributes:
        ruta_archivo_excel (Path): La ruta del archivo Excel que se va a procesar.
        ruta_carpeta_destino (Path): La ruta de la carpeta donde se van a guardar
            los archivos generados por el autómata.
        nombre_serie (str): El nombre de la serie que se va a procesar. Este nombre
            se extrae del archivo Excel y se utiliza para nombrar las carpetas.
        diccionario_pendientes (dict[int,Pendiente]): Un diccionarios de objetos Pendiente que
            representan las sub-tareas que el autómata ejecutará. Cada Pendiente
            aporta los siguientes atributos (para evitar tener que abrir la clase
            Pendiente en otros puntos del código):
                - fecha (int): Identificador o índice de la fecha/período.
                - posicion (int): Posición o fila/columna dentro del Excel.
                - completado (bool): Estado de completado del sub-proceso.
    """
    ruta_archivo_excel:Path
    ruta_carpeta_destino:Path
    nombre_serie:str
    diccionario_pendientes:dict[int,Pendiente]=field(default_factory=dict)

    def __repr__(self):
        return (f"Serie: {self.nombre_serie} | "
                f"Archivo: {self.ruta_archivo_excel.name} | "
                f"Pendientes: {len(self.diccionario_pendientes)}")


class rutaCarpetas:
    """Maneja rutas y creación de carpetas/archivos para el proceso de consulta.

    Crea la carpeta de destino para la serie extraída del nombre del archivo Excel
    y expone utilidades para listar los archivos ya existentes en esa carpeta.

    Args:
        ruta_carpeta_output (Path): Carpeta base donde se crearán las subcarpetas
            por serie (se lee desde config.json en el flujo principal).
        archivo_excel (Path): Ruta del archivo Excel que se va a procesar.
        patronvalidador (RegexpPatrones): Objeto/servicio que extrae el nombre de
            la serie a partir del nombre del archivo (patrón de validación).
    """
    def __init__(self,ruta_carpeta_output:Path,archivo_excel:Path,patronvalidador:RegexpPatrones)->None:

        self.ruta_archivo_excel=archivo_excel
        self.nombreArchivExcel=archivo_excel.stem
        self.nombre_serie=patronvalidador.busqueda_patron_para_nombrar_carpetas(self.nombreArchivExcel)
        self.carpeta_destino=ruta_carpeta_output / self.nombre_serie
        self.carpeta_destino.mkdir(parents=True, exist_ok=True)
    
    @property
    def revisar_archivos_en_carpeta(self)->list:
        listaarchivos=[]
        for archivos in self.carpeta_destino.iterdir():
            if archivos.is_file():
                listaarchivos.append(archivos.stem)
        return listaarchivos
    


class ManagerPendientes:
    def __init__(self,mapa:rutaCarpetas,lista_fechas:list)->None: 
        #La lista de los bimestres contenidos en el archivo excel
        self.lista_fechas=lista_fechas
        #uso de la clase rutaCarpetas
        self.mapa=mapa
        self.archivos_creados=mapa.revisar_archivos_en_carpeta
    #Con la lista de fechas y el nombre de la serie se crea una lista de los archivos que deberian estar creados al finalizar la consulta, esto es crucial para que el vigia pueda revisar cuales archivos faltan y generar las tareas pendientes para el automata
    @property
    def dummy_archivos(self)->list:
        return [f"{self.mapa.nombre_serie}_{i}" for i in self.lista_fechas]       

    
    #El vigia va a revisar los pendientes que hay en la carpeta
    @property    
    def pendientes_de_consulta(self)->list:
        pendientes=set(self.dummy_archivos)-set(self.archivos_creados)
        return list(pendientes)
    
    @property
    def mapa_de_pendientes_alfa(self)->dict[int,TareaPendiente]:
        diccionario_de_pendientes={}
        pendientes_consultar= {int(i[-6:]) for i in self.pendientes_de_consulta}


        for posicion, valor in dict(enumerate(self.lista_fechas)).items():
            if valor in pendientes_consultar:
                diccionario_de_pendientes[valor]=Pendiente(fecha=valor,posicion=posicion,completado=False)

        nueva_tarea=TareaPendiente(ruta_archivo_excel=self.mapa.ruta_archivo_excel,
                    ruta_carpeta_destino=self.mapa.carpeta_destino,
                    nombre_serie=self.mapa.nombre_serie,
                    diccionario_pendientes=diccionario_de_pendientes)
        logger.info(f"Tarea lista para ejecución: {nueva_tarea}")

        return nueva_tarea

    #Esta funcion ejecuta en tiempo real la tarea pendiente que se genero con la propiedad anterior, esto es crucial para 
    #que el automata pueda ejecutar las tareas pendientes sin necesidad de que el USUARIO humano este revisando constantemente 
    #la carpeta de destino y generando las tareas pendientes manualmente, ademas de que al generar la tarea pendiente en tiempo real
    #se evita que se generen tareas pendientes erroneas por archivos que se crean despues de revisar los pendientes pero antes de 
    #ejecutar el automata, lo que podria generar errores o confusiones en el proceso de consulta.
    @property    
    def mapa_de_pendientes(self)-> TareaPendiente:
        indices_pendientes=[]
        for k, i in enumerate(self.dummy_archivos):
            for j in self.pendientes_de_consulta:
                if i==j:
                    indices_pendientes.append(k)


        nueva_tarea=TareaPendiente(ruta_archivo_excel=self.mapa.ruta_archivo_excel,
                    ruta_carpeta_destino=self.mapa.carpeta_destino,
                    nombre_serie=self.mapa.nombre_serie,
                    diccionario_pendientes=indices_pendientes)
        logger.info(f"Tarea lista para ejecución: {nueva_tarea}")

        return nueva_tarea



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
    """Administra la interacción con instancias y libros de Excel.

    Esta clase es responsable de inicializar y controlar una aplicación Excel a través de
    la librería win32com, abrir libros de trabajo específicos, y traer ventanas o libros al frente.

    Attributes:
        patron_validador (RegexpPatrones): Validador de patrones para datos extraídos del libro.
        libro_control_ruta (Path): Ruta del libro de control que debe abrirse.
        ventana_excel: Instancia de la aplicación Excel.
        hwnd: Identificador de ventana de la instancia Excel.
        libro (dict[TipoLibro, libroExcelManager]): Mapa de tipos de libro a sus gestores correspondientes.

    Methods:
        abrir_instancia_excel(): Inicializa y obtiene una nueva instancia de Excel.
        traer_instancia_excel_al_frente(): Hace visible la ventana de Excel y la trae al frente.
        traer_libro_al_frente(tipo_libro): Activa el libro especificado dentro de la instancia de Excel.
        abrir_libro(tipo_libro, ruta_archivo): Abre el libro indicado y lo registra en el gestor.
        cerrar_libro_excel(nombre_libro, ruta_guardar): Cierra el libro especificado, guardando si se indica una ruta.
        extraer_tabla_y_guardar_en_ruta(ruta_destino, libro_control_ruta): Extrae una tabla del libro de consulta y la guarda en el libro de control, luego cierra el libro de control.

    """

    def __init__(self,patron_validador:RegexpPatrones,libro_control_ruta:Path) -> None:
        self.patron_validador=patron_validador
        self.libro_control_ruta=libro_control_ruta
        self.ventana_excel = None
        self.hwnd = None  
        self.libro:dict[TipoLibro,libroExcelManager]={} 



    def abrir_instancia_excel(self) -> None:
        try:
            self.ventana_excel = win32.Dispatch("Excel.Application")
            time.sleep(1)
            self.hwnd = self.ventana_excel.Hwnd
        except Exception as e:
            logger.critical(f"Error al abrir instancia de excel validar permisos o conflictos con la libreria win32.client: {e}")
            raise

    def traer_instancia_excel_al_frente(self)-> None:
        if self.hwnd is None:
            logger.warning("Se intenta traer una ventana de excel que no esta instanciada, se invocara el metodo \"abrir_instancia_excel\"")
            self.abrir_instancia_excel()
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
            except ex.ErrorAlTraerLibroAlFrente as e:
                logger.error(f"Error al traer el libro al frente: {e}")
            

    def abrir_libro(self, tipo_libro:TipoLibro, ruta_archivo: Path) -> None:
        logger.info(f"Abriendo archivo Excel {ruta_archivo.stem}")
        
        if not ruta_archivo.exists():
            logger.error(f"El archivo {ruta_archivo} no existe.")
            raise FileNotFoundError(f"Archivo no encontrado: {ruta_archivo}")

        if self.ventana_excel is None:
            self.abrir_instancia_excel()
        
        if tipo_libro is TipoLibro.LIBRO_CONTROL:
            nuevo_libro=self.ventana_excel.Workbooks.Open(str(self.libro_control_ruta))
            self.libro[tipo_libro]=libroExcelManager(nuevo_libro,self.patron_validador)
            
        else:
            nuevo_libro=self.ventana_excel.Workbooks.Open(str(ruta_archivo))
            self.libro[tipo_libro]=libroExcelManager(nuevo_libro,self.patron_validador)
        


    def cerrar_libro_excel(self, nombre_libro: TipoLibro, ruta_guardar: Path | None = None) -> None:
        

        #Cierra un libro específico de Excel manejando la seguridad de macros y alertas.
        #Si se proporciona ruta_guardar, realiza un SaveAs antes de cerrar.
        try:
            libro_com = self.libro[nombre_libro]
        except KeyError:
            logger.warning(f"Intentando cerrar un libro que no está abierto: {nombre_libro}")
            return
        if not self.ventana_excel or not self.libro:
            logger.info("No hay sesión de Excel activa para cerrar.") 
            return

        try:
            # Configuración de seguridad para el guardado silencioso
            self.ventana_excel.EnableEvents = False
            self.ventana_excel.DisplayAlerts = False
            


            # Fix para Office 365: El autoguardado en la nube suele bloquear el SaveAs
            try:
                libro_com.AutoSaveOn = False 
            except ex.ErrorAlDesactivarAlertasExcel as e:
                logger.warning(f"No se pudo desactivar AutoSave, es posible que el guardado silencioso falle: {e}")
                pass 
            
            if ruta_guardar:
                libro_com.Activate()
                # resolve() asegura rutas absolutas evitando errores de permisos en Windows
                ruta_final = str(ruta_guardar.resolve())
                libro_com.SaveAs(ruta_final)
                logger.info(f"Libro guardado en: {ruta_final}")

            libro_com.Close(SaveChanges=False)
            del self.libro[nombre_libro]

        except ex.ErrorAlCerrarLibro as e:
            logger.error(f"Fallo al cerrar/guardar el libro '{nombre_libro.name}': {e}")
        
        finally:
            # Restaurar estado de la aplicación siempre
            if self.ventana_excel:
                self.ventana_excel.EnableEvents = True
                self.ventana_excel.DisplayAlerts = True

    def extraer_tabla_y_guardar_en_ruta(self,ruta_destino:Path):
        libro_c=TipoLibro.LIBRO_CONTROL
        #abrimos y ponemos el libro de extraccion en una varible 
        self.abrir_libro(libro_c,self.libro_control_ruta)
        time.sleep(2)
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

    def __del__(self):
        # Asegura que la instancia de Excel se cierre al eliminar el objeto ExcelManager
        if self.ventana_excel:
            try:
                self.ventana_excel.Quit()
                logger.info("Instancia de Excel cerrada correctamente.")
            except ex.ErrorAlCerrarExcel as e:
                logger.error(f"Error al cerrar la instancia de Excel: {e}")
    
    def cerrar_instancia_excel(self):
        self.cerrar_libro_excel(TipoLibro.LIBRO_CONTROL)
        self.cerrar_libro_excel(TipoLibro.LIBRO_CONSULTA)
        if self.ventana_excel:
            try:
                self.ventana_excel.Quit()
                logger.info("Instancia de Excel cerrada correctamente.")
            except ex.ErrorAlCerrarExcel as e:
                logger.error(f"Error al cerrar la instancia de Excel: {e}")