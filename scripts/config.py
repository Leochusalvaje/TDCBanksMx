
from pathlib import Path
import re

def crear_directorios(cls):

    # Recorremos todos los atributos de la clase
    for nombre, valor in cls.__dict__.items():
        if isinstance(valor, Path):
            print(f"Asegurando directorio: {valor}")
            valor.mkdir(parents=True, exist_ok=True)
    
    return cls

@crear_directorios
class Paths:

    base = Path(__file__).parent.parent
    data = base  / "Data"
    dataTdc=data/"datosTarjetasCredito"
    output = base / "Output"
    dataset=base/"Datasets"

    ALL = [
        base,
        data,
        dataTdc,
        output,
        dataset

    ]

        
class regexp_patrones:
    def __init__(self):
        self.FECHAS=r"^20\d{2}(0[1-9]|1[0-2])$"
        self.ARCHIVOSCNBV=r"^[^_]+_[^_]+_([^_]+)"
    # Metodo para revisar patrones
    def revisar_patron(self,patron:str,cadena:str)->str:
        try:
            if not re.match(patron,cadena):
                raise ValueError("Patron invalido")
            print("Patron Valido, se puede proceder")
        except ValueError as e:
            print(f"Error en el patron: {e}")
            return None
        return cadena
    
    def busqueda_patron_para_nombrar_carpetas(self,cadena:str)-> str:
            # Explicación del patrón:
        # ^[^_]+  -> Salta el primer bloque (040)
        # _[^_]+  -> Salta el segundo bloque (12e)
        # _([^_]+) -> Captura lo que hay en el tercer bloque hasta el siguiente guion bajo
        patron = self.ARCHIVOSCNBV

        resultado = re.search(patron, cadena)

        if resultado:
            serie_nombre = resultado.group()
            print(f"Serie extraída: {serie_nombre}") 
        return serie_nombre


patron=regexp_patrones()
patrones={"fecha":r"^20\d{2}(0[1-9]|1[0-2])$","archivos_excel":r"^[^_]+_[^_]+_([^_]+)"}

