from config import ConfigManager,Paths,DescripcionArchivosCriticos,RegexpPatrones
from pathlib import Path
from automata import rutaCarpetas,carajo
from automata import lista_verdad
import logging
#Configuracion del nombre del logger para que se muestre en los logs siempre va despues de las importaciones solo en caso de utilizar decoradores poner antes del decorador
logger=logging.getLogger("Orquestador")

def main():
    logger.info("--- Iniciando validación de rutas ---")
    #---Inicializamos el despachador---# 
    ruta_config = Path(__file__).parent / "config.json"
    despachador=ConfigManager(ruta_config)
    despachador.cargarConfiguracion()
    #---DESPACHADOR PARA SERVIR LOS OBJETOS DE RUTAS Y VALIDACIONES---# 
    bimestreFinal=despachador.cargarBimestreFinal
    bimestreInicial=despachador.datos.get("bimestre_inicial")

    descripciones = DescripcionArchivosCriticos()
    #---CONFIGURACIÓN DE RUTAS---# 
    paths = Paths(despachador.cargarRutaBase,descripciones)
    patron=RegexpPatrones()

    listaFechas=lista_verdad(bimestreInicial,bimestreFinal,patron)

    listaTareasPendientes=[]
    for rutaExcel in paths.dataTdc.iterdir():
        if rutaExcel.is_file() and rutaExcel.suffix in ['.xlsx', '.xls']:
            logger.info(f"Analizando archivo: {rutaExcel.name}")
            try:
                gestorRutas=rutaCarpetas(paths.output,rutaExcel,patron)
                gestorPendientes=carajo(gestorRutas,listaFechas)
                listaTareasPendientes.append(gestorPendientes.mapa_de_pendientes)
            except Exception as e:
                logger.error(f"Error al procesar {rutaExcel.name}: {e}")
        else:
            logger.warning(f"Archivo no válido (no es Excel): {rutaExcel.name}")


    
    # Aquí es donde paths.verificar_todo() debería disparar los logs

if __name__ == "__main__":
    # La configuración se hace UNA sola vez al principio
    logging.basicConfig(
        force=True,  # Asegura que esta configuración se aplique incluso si logging ya ha sido configurado por otro módulo
        level=logging.INFO,
        format='%(asctime)s - [%(name)s] - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("informacion.log",mode='w',encoding='utf-8'), # Escribe en el archivo
            logging.StreamHandler()                # También muestra en consola
        ]

    )
    
    main()