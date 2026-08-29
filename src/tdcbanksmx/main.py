import logging
import sys
from pathlib import Path
from tdcbanksmx.orquestador import PipelineOrquestador
from tdcbanksmx.cli import build_parser

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
def configurar_logging(nivel=logging.DEBUG, modo_consola=True) -> None:

    
    """Configura los handlers globales de logging."""
    logging.basicConfig(
        force=True,  # Sobrescribe cualquier configuración previa de logging
        level=nivel,
        format='%(asctime)s - [%(name)s] - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("informacion.log", mode='w', encoding='utf-8'),
            logging.StreamHandler(sys.stdout) if modo_consola else None
        ]
    )

def main() -> None:
    # 1. Configuras el logging global una sola vez al entrar
    configurar_logging()

    # 2. Parseas los argumentos y ejecutas el handler correspondiente
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)

if __name__ == "__main__":
    main()