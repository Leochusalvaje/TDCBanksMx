from config import ConfigManager,Paths,DescripcionArchivosCriticos,RegexpPatrones
from pathlib import Path
from automata import TiemposConfig,CheckerCoordenada,ConfigVisual,pointCheckboxManager,automataConsultas
from automata import lista_verdad,AutomataPrime,PointCheckManagerPrime
import logging
from excel_manager import rutaCarpetas,carajo,ExcelManager,tareaPendiente
#Configuracion del nombre del logger para que se muestre en los logs siempre va despues de las importaciones solo en caso de utilizar decoradores poner antes del decorador
logger=logging.getLogger("Orquestador")




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




    def procesar_archivos(self):
        lista_tareas_pendientes:list[tareaPendiente]=[]

        for rutaExcel in self.paths.dataTdc.iterdir():
            if rutaExcel.is_file() and rutaExcel.suffix in ['.xlsx', '.xls']:
                logger.info(f"Analizando archivo: {rutaExcel.name}")
                try:
                    # aqui utilziamos la logica conteninida en Excelmanager para procesar cada archivo excel y generar la lista de tareas pendientes, 
                    # esta logica se encarga de leer el archivo excel, validar su formato, extraer la información necesaria y generar la lista de tareas
                    # pendientes que el automata utilizara para realizar las consultas en la pagina de la CNBV, es importante revisar que esta logica este 
                    # funcionando correctamente.
                    gestorRutas = rutaCarpetas(self.paths.output, rutaExcel, self.patrones)
                    gestorPendientes = carajo(gestorRutas, self.lista_fechas)
                    lista_tareas_pendientes.append(gestorPendientes.mapa_de_pendientes) 
                except Exception as e:
                    logger.error(f"Error al procesar {rutaExcel.name}: {e}")
            else:
                logger.warning(f"Archivo no válido (no es Excel): {rutaExcel.name}")
        
        aut_debugger=AutomataPrime(lista_tareas_pendientes,self.ventana_excel,self.UI,self.lista_fechas)
        return aut_debugger