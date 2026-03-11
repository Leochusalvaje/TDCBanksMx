
from pathlib import Path
import re
import logging
import json
from enum import Enum

logger=logging.getLogger(__name__)


class ConfigManager:
    def __init__(self,rutaJson:Path):
        self.rutaJson=rutaJson
        self.datos={}
        self._cargarConfiguracion()

    def _cargarConfiguracion(self)->None:


        if not self.rutaJson.exists():
            logger.critical(f"Error al buscar el archivo de configuracion Json")
            raise FileNotFoundError(f"validar que {self.rutaJson} exista")
        
        with open(self.rutaJson,encoding="utf-8") as f:

            try:
                self.datos=json.load(f)

            except FileNotFoundError:
                logger.critical(f"ARCHIVO NO ENCONTRADO: No existe el archivo en {self.rutaJson}")
                raise
                
            except json.JSONDecodeError as e:
                logger.critical(f"JSON CORRUPTO: Error de sintaxis en {self.rutaJson}. Detalle: {e}")
                raise
                
            except Exception as e:
                logger.critical(f"ERROR INESPERADO al abrir {self.rutaJson}: {e}")
                raise
            
        if not isinstance(self.datos, dict):
            logger.critical("El archivo config.json no tiene el formato correcto validar documentación")
            raise ValueError("El JSON raíz debe ser un objeto (dict)")
        
    def obtener_dato(self,llave:str)->dict:
        validar=self.datos[llave]  
        if validar is None:
            logger.critical(f"La llave '{llave}' no existe en el archivo de configuración.")
            raise ValueError(f"La llave '{llave}' no existe en el archivo de configuración.")
        return validar

    @property
    def cargar_ruta_base(self):
        return Path(self.obtener_dato("ruta_base"))
    @property
    def cargar_bimestre_inicial(self):
        return self.obtener_dato("bimestre_inicial")
    @property
    def cargar_bimestre_final(self):
        return self.obtener_dato("bimestre_final")
    @property
    def cargar_patrones(self):
        return self.obtener_dato("patrones")
    @property
    def cargar_coordenadas(self):
        return self.obtener_dato("coordenadas")
    @property
    def cargar_tiempos(self):
        return self.obtener_dato("tiempos")

class DescripcionArchivosCriticos:
    def __init__(self):
        self.configJson = "config.json"
        self.libroControl = "LibroControl.xlsx"
        
        # Diccionario de ayuda para el usuario
        self.guia_archivos = {
            self.configJson: {
                "nombre": "Archivo de Configuración (JSON)",
                "descripcion": "Contiene las coordenadas de clic y rutas de carpetas  y segmentacion de codigo escencial para la ejecucion.",
                "instruccion": "Lee el readme para entender la estructura del JSON y completa las coordenadas necesarias para tu pantalla. Asegúrate de que el archivo esté en la ruta base /scripts."
            },
            self.libroControl: {
                "nombre": "Libro de Control (Excel)",
                "descripcion": "Este archivo le ayuda al automata a separar las consultas para descargarlas en su respectiva carpeta.",
                "instruccion": "Crea un archivo Excel vacío llamado 'LibroControl.xlsx' en la carpeta ruta base / Output."
            }
        }

class Paths:
    def __init__(self, rutaBase:str,descripciones: DescripcionArchivosCriticos=DescripcionArchivosCriticos()):

        #---CARPETAS PRINCIPALES---#
        # Se usa .cwd() (Current Working Directory) para obtener la carpeta desde donde se 
        # lanza el programa, evitando que Path(__file__) use la ubicación de este archivo de clase.
        # Se añade .resolve() para "limpiar" la ruta: convierte el punto "." o rutas relativas 
        # en rutas absolutas de Windows (C:\...), asegurando que el logger y el sistema 
        # encuentren las carpetas sin ambigüedades.
        if rutaBase == ".":
            self.base = Path.cwd().resolve()
        else:
            self.base = Path(rutaBase).resolve()


        self.data=self.base/"Data"
        self.dataTdc=self.data/"datosTarjetasCredito"
        self.output=self.base/"Output"
        self.dataset=self.base/"Datasets"
        self.scripts=self.base/"scripts"
        #---LISTA DE TODAS LAS CARPETAS---#
        self.ALL = [
            self.base,
            self.data,
            self.dataTdc,
            self.output,
            self.dataset,
            self.scripts
        ]
        #---ARCHIVOS---#
        self.configJson=self.scripts/"config.json"
        self.libroControl=self.output/"LibroControl.xlsx"
        #---CREAR DIRECTORIOS---#
        self._descripciones=descripciones
        try:
            self._crearDirectorios()
        except Exception as e:
            logger.critical(f"Error al crear directorios: {e} revisar los permisos de escritura en la ruta {self.base}")
        self._validarArchivosCriticos()

    def rutaParaLogger(self, ruta_objetivo):

        try:
            # self.base.name captura "TDCBanksMx"
            # .relative_to(self.base) captura lo que sigue
            return Path(self.base.name) / ruta_objetivo.relative_to(self.base)
        except (ValueError, AttributeError):
            # Si la ruta no pertenece al proyecto o no es objeto Path, la devuelve tal cual
            return ruta_objetivo

    def _crearDirectorios(self):
        for _, valor in vars(self).items():
            if isinstance(valor, Path):
                if not valor.suffix:  # Solo crear directorios para atributos que son rutas sin extensión
                    logger.info(f"Asegurando directorio: {self.rutaParaLogger(valor)}")
                    valor.mkdir(parents=True, exist_ok=True) 
    
    def _validarArchivosCriticos(self):
        archivos_criticos = [self.configJson,self.libroControl]
        info = self._descripciones.guia_archivos
        for archivo in archivos_criticos:
            if not archivo.exists():
                logger.critical(f"Archivo crítico no encontrado: {self.rutaParaLogger(archivo)}. Asegúrate de que el archivo exista en la ruta especificada.")
                logger.info(f"NOMBRE: {info[archivo.stem]['nombre']}")
                logger.info(f"DESCRIPCIÓN: {info[archivo.stem]['descripcion']}")
                logger.info(f"ACCIÓN: {info[archivo.stem]['instruccion']}")
                raise FileNotFoundError(f"Archivo crítico no encontrado: {archivo}")
            else:
                logger.info(f"Archivo crítico encontrado en: {self.rutaParaLogger(archivo)}")

class TipoPatron(Enum):
    FECHA = "fecha"
    ARCHIVO_CNBV = "archivo_cnbv"
        
class RegexpPatrones:


    def __init__(self,patrones:dict):
        
        if not patrones:
            logger.error("Se debe pasar un diccionario de patrones")
            raise ValueError("Se debe pasar un diccionario de patrones")

        # Validar claves contra Enum
        enum_values={e.value for e in TipoPatron}
        for clave in patrones:
            if clave not in enum_values:
                logger.error(f"Patrón desconocido: {clave}, anexarlo en la clase TipoPatron")
                raise ValueError(f"Patrón desconocido: {clave}")

        # Compilar regex y guardar en diccionario
        self._patrones={
            TipoPatron(clave): re.compile(regex) for clave, regex in patrones.items()
        }

    @property
    def revisar_patrones(self):

        return self._patrones

    def revisar_patron(self, tipo: TipoPatron, cadena: str) -> str | None:
        
        patron = self._patrones.get(tipo)
        if not patron.match(cadena):
            logger.error(f"El string '{cadena}' no cumple con el patrón '{tipo.name}'")
            return None
        logger.debug(f"El string '{cadena}' cumple con el patrón '{tipo.name}'")
        return cadena

    def busqueda_patron_para_nombrar_carpetas(self,cadena:str)->str|None:

        patron = self._patrones[TipoPatron.ARCHIVO_CNBV]
        resultado = patron.search(cadena)

        if resultado:
            serie_nombre = resultado.group()
            logger.info(f"Serie extraída: {serie_nombre}")
            return serie_nombre

        logger.warning(f"No se encontró patrón en la cadena: {cadena}")
        return None



