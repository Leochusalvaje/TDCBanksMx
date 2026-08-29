import argparse
from tdcbanksmx.orquestador import PipelineOrquestador,ServiceTasks
import logging
from pathlib import Path

logger=logging.getLogger("main")

def handler_validar_rutas(args:argparse.Namespace) -> None:
    ruta_config=Path(args.config)
    orquestador=PipelineOrquestador(ruta_config)

def handler_crear_pendientes(args:argparse.Namespace) -> None:
    tareas=ServiceTasks(args.config)
    tareas.generar_tareas_pendientes()


def build_parser() -> argparse.ArgumentParser:

    parser = argparse.ArgumentParser( prog="tdcbanksmx",
                                    description="Pruebas de validación en TDCBanksMx"
                                    )

    subparsers = parser.add_subparsers(dest="task", required=True)

    parser_val=subparsers.add_parser("validar-rutas", 
                                    help="Valida las rutas de los archivos críticos según la configuración proporcionada.")

    parser_val.add_argument(
        "--config",
        type=Path,
        default=Path("src/tdcbanksmx/config.json"),
        help="Ruta al archivo de configuración"
        )

    parser_val.set_defaults(func=handler_validar_rutas)

    parser_pendientes=subparsers.add_parser("pendientes", help="Se generan los pendientes que consumira el automata")
    parser_pendientes.add_argument(
        "--config",
        type=Path,
        default=Path("src/tdcbanksmx/config.json"),
        help="Ruta al archivo de configuración"
        )
    parser_pendientes.set_defaults(func=handler_crear_pendientes)

    return parser
