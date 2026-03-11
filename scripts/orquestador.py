from config import ConfigManager,Paths,DescripcionArchivosCriticos,RegexpPatrones
from pathlib import Path
from automata import rutaCarpetas,carajo,libroExcelManager,TipoLibro,TiemposConfig,CheckerCoordenada,ConfigVisual,ExcelManager,pointCheckboxManager,automataConsultas
from automata import lista_verdad
import logging
#Configuracion del nombre del logger para que se muestre en los logs siempre va despues de las importaciones solo en caso de utilizar decoradores poner antes del decorador
logger=logging.getLogger("Orquestador")




class PipelineOrquestador:
    def __init__(self, ruta_config: Path):
        self.despachador = ConfigManager(ruta_config)
        self.descripciones = DescripcionArchivosCriticos()
        self.paths = Paths(self.despachador.cargar_ruta_base, self.descripciones)
        self.patrones = RegexpPatrones(self.despachador.cargar_patrones)
        self.listaFechas = lista_verdad(
            self.despachador.cargar_bimestre_inicial,
            self.despachador.cargar_bimestre_final,
            self.patrones
        )
        self.ventana_excel = ExcelManager(self.patrones,self.paths.libroControl)
        #Cargamos la configuracion de coordenadas y tiempos para el automata
        self.coordenadas = CheckerCoordenada(**self.despachador.cargar_coordenadas)
        self.tiempos = TiemposConfig(**self.despachador.cargar_tiempos)
        self.config_visual = ConfigVisual(self.coordenadas,self.tiempos)
        
        self.UI=pointCheckboxManager(self.config_visual)
        self.automata=automataConsultas(self.UI)

    def procesar_archivos(self):
        lista_tareas_pendientes = []

        for rutaExcel in self.paths.dataTdc.iterdir():
            if rutaExcel.is_file() and rutaExcel.suffix in ['.xlsx', '.xls']:
                logger.info(f"Analizando archivo: {rutaExcel.name}")
                try:
                    gestorRutas = rutaCarpetas(self.paths.output, rutaExcel, self.patrones)
                    gestorPendientes = carajo(gestorRutas, self.listaFechas)
                    lista_tareas_pendientes.append(gestorPendientes.mapa_de_pendientes)
                except Exception as e:
                    logger.error(f"Error al procesar {rutaExcel.name}: {e}")
            else:
                logger.warning(f"Archivo no válido (no es Excel): {rutaExcel.name}")

        return lista_tareas_pendientes