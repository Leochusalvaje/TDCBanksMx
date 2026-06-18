import logging
from pathlib import Path
from orquestador import PipelineOrquestador

#Configuracion del nombre del logger para que se muestre en los logs siempre va despues de las importaciones solo en caso de utilizar decoradores poner antes del decorador
logger=logging.getLogger("main")

def main():
    logger.info("--- Iniciando validación de rutas ---")
    #---Inicializamos el despachador---# 
    ruta_config = Path(__file__).parent / "config.json"
    automata=PipelineOrquestador(ruta_config)
    automata.generar_tareas_pendientes()
    
    pendientes=automata.procesar_archivos()
    logger.info("--- Validación de rutas completada ---")
    #---Aquí se podrían agregar más pasos del proceso, como iniciar el automata o procesar los pendientes---#

    
    # Aquí es donde paths.verificar_todo() debería disparar los logs

if __name__ == "__main__":
    # La configuración se hace UNA sola vez al principio
    logging.basicConfig(
        force=True,  # Asegura que esta configuración se aplique incluso si logging ya ha sido configurado por otro módulo
        level=logging.DEBUG,
        format='%(asctime)s - [%(name)s] - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("informacion.log",mode='w',encoding='utf-8'), # Escribe en el archivo
            #logging.StreamHandler()                # mostrar en consola
        ]

    )
    
    main()