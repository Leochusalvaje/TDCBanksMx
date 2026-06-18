from dataclasses import dataclass

from config import ConfigManager,Paths,DescripcionArchivosCriticos,RegexpPatrones
from pathlib import Path
from automata import TiemposConfig,CheckerCoordenada,ConfigVisual,pointCheckboxManager,automataConsultas
from automata import lista_verdad,AutomataPrime,PointCheckManagerPrime
import logging
from excel_manager import rutaCarpetas,ManagerPendientes,ExcelManager,TareaPendiente
#Configuracion del nombre del logger para que se muestre en los logs siempre va despues de las importaciones solo en caso de utilizar decoradores poner antes del decorador
logger=logging.getLogger("Orquestador")

@dataclass
class ContextoAutomata:
    """ContextoAutomata es una clase de datos que encapsula toda la información y configuraciones necesarias para que el automata pueda ejecutar las consultas de manera eficiente y organizada. Esta clase actúa como un contenedor centralizado para todas las configuraciones, rutas, patrones y tiempos que el automata necesita para funcionar correctamente.
    
    Atributos:
    - config_visual (ConfigVisual): Contiene las coordenadas y tiempos configurados para el automata, utilizados para realizar los clics y esperar los tiempos adecuados durante la ejecución.
    - lista_fechas (list[str]): Una lista de fechas generada a partir de la configuración del bimestre inicial y final, utilizada para navegar por las fechas dentro del archivo Excel.
    - ventana_excel (ExcelManager): Una instancia de ExcelManager que se encarga de manejar la interacción con los archivos Excel, incluyendo la apertura, lectura y escritura de datos.
    """
    tareas:list[TareaPendiente]
    manipulacion:ExcelManager
    navegante:PointCheckManagerPrime
    lista_fechas:list[str]



class PipelineOrquestador:
    def __init__(self, ruta_config: Path):
        self.despachador = ConfigManager(ruta_config)
        self.descripciones = DescripcionArchivosCriticos()
        self.paths = Paths(self.despachador.cargar_ruta_base, self.descripciones)
        self.patrones = RegexpPatrones(self.despachador.cargar_patrones)

        # Esta lista es generada por el usuario en el config.json y sirve para que el automata pueda navegar por las fechas dentro del archivo excel,
        # es crucial para el funcionamiento del automata ya que sin esta lista el automata no sabria como navegar por las fechas dentro del archivo excel 
        # y no podria extraer la información correctamente, por lo que es importante revisar que se este generando correctamente y que se este cumpliendo el patron 
        # establecido en la clase RegexpPatrones.
        self.lista_fechas = lista_verdad(
            self.despachador.cargar_bimestre_inicial,
            self.despachador.cargar_bimestre_final,
            self.patrones
        )
        self.ventana_excel = ExcelManager(self.patrones,self.paths.libroControl)
        #Cargamos la configuracion de coordenadas y tiempos para el automata
        self.coordenadas = CheckerCoordenada(**self.despachador.cargar_coordenadas)
        self.tiempos = TiemposConfig(**self.despachador.cargar_tiempos)

        #cargamos la configuracion visual en una clase para que el automata pueda acceder a ella de forma centralizada y organizada
        self.config_visual = ConfigVisual(self.coordenadas,self.tiempos)
        
        self.UI=PointCheckManagerPrime(self.config_visual)

        self.tareas_pendientes: list[TareaPendiente] = []

    def generar_tareas_pendientes(self)->None:
        lista_tareas_pendientes:list[TareaPendiente]=[]

        for rutaExcel in self.paths.dataTdc.iterdir():
            if rutaExcel.is_file() and rutaExcel.suffix in ['.xlsx', '.xls']:
                logger.info(f"Analizando archivo: {rutaExcel.name}")
                try:
                    # aqui utilziamos la logica conteninida en Excelmanager para procesar cada archivo excel y generar la lista de tareas pendientes, 
                    # esta logica se encarga de leer el archivo excel, validar su formato, extraer la información necesaria y generar la lista de tareas
                    # pendientes que el automata utilizara para realizar las consultas en la pagina de la CNBV, es importante revisar que esta logica este 
                    # funcionando correctamente.
                    gestorRutas = rutaCarpetas(self.paths.output, rutaExcel, self.patrones)
                    gestorPendientes = ManagerPendientes(gestorRutas, self.lista_fechas)
                    lista_tareas_pendientes.append(gestorPendientes.mapa_de_pendientes_alfa) 
                except Exception as e:
                    logger.error(f"Error al procesar {rutaExcel.name}: {e}")
            else:
                logger.warning(f"Archivo no válido (no es Excel): {rutaExcel.name}")
        self.tareas_pendientes=lista_tareas_pendientes

    
    def procesar_archivos(self)->None:
        aut_debugger=AutomataPrime(self.tareas_pendientes[0],self.ventana_excel,self.UI)
        aut_debugger.procesar_tarea()
        pass