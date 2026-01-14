
from pathlib import Path
import re

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

    def revisar_patron(self,patron,cadena):
        try:
            if not re.match(patron,cadena):
                raise ValueError("Patron invalido")
            print("Patron Valido, se puede proceder")
        except ValueError as e:
            print(f"Error en el patron: {e}")
            return None
        return cadena


patron=regexp_patrones()
patrones={"fecha":r"^20\d{2}(0[1-9]|1[0-2])$","archivos_excel":r"^[^_]+_[^_]+_([^_]+)"}