
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

# #    for indice in lisapendientes:

# #             if indice==7:
# #                 print("para para el debugger")


# #             self.avanzar_a_fecha(indice-self._propiocepcion)
# #             fechaEnBucle=self.matrizFechas[indice]
            
# #             validamos que el boton no es te seleccionado  
# #             if not self.matrizBotones[fechaEnBucle].estado:
# #                 self._navegante.click_en_coordenada(boton=self.matrizBotones[fechaEnBucle])


# #             self._navegante.obtener()
# #             fechaEnHoja=int(string_eureka_fechas(self.hojaExcel)[0])

# #             if self.obtener_fecha_actual(fechaEnHoja):  
# #                 cicloRoto=True


# #                 ajustePropiocepcion=self.matrizBotones[fechaEnHoja].posicion-self._propiocepcion
# #                 self._propiocepcion+=ajustePropiocepcion

# #                 self.apagar_boton(self.matrizBotones[fechaEnHoja])
# #                 self.matrizBotones[fechaEnBucle].cambiar_estado()
# #                 self.matrizBotones[fechaEnHoja].cambiar_estado()
# #                 if posicion!=1:

# #                     self.apagar_boton(self.matrizBotones[fechaEnHoja])

# #                 print("entro al caso particular")
# #                 break

# #             self.apagar_boton(self.matrizBotones[fechaEnHoja])



# #             posicion+=1    

# #             print("propiocepcion despues de avanzar",self._propiocepcion)



# #         print("fecha de cambio",fechaEnHoja)

# #         if cicloRoto:
# #             fechaQueRompeElBucle=fechaEnHoja

# #             ajuste=ajustePropiocepcion
# #             for indice in lisapendientes[posicion:]:

# #                 fechaEnBucle=self.matrizFechas[indice]
# #                 Pruebas para ajustar la propiocepcion con ayuda del debugger , simular los clciks manualmente mientras se testea
# #                 print(indice)
# #                 print(self.observador.listaFechas[-self._propiocepcion])
# #                 if indice==3:
# #                     print("para revisar con el debugger")
                
# #                 self.avanzar_a_fecha(indice-self._propiocepcion)
# #                 self._navegante.click_en_coordenada(boton=self.matrizBotones[fechaEnBucle])
                
# #                 print(self.observador.listaFechas[-self._propiocepcion])
# #                 if  not self.matrizFechas[self._propiocepcion] in [201108,201106]:
# #                     self._navegante.obtener()
# #                     self.obtener_fecha_actual(fechaEnHoja)

# #                     fechaEnHoja=int(string_eureka_fechas(self.hojaExcel)[0])

# #                     self._propiocepcion+=ajuste
# #                     self.apagar_boton(self.matrizBotones[fechaEnHoja])
# #                     ajuste-=1


# #                 elif self.matrizFechas[self._propiocepcion]==201108:
# #                     self._navegante.click_en_coordenada()
# #                     self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["penultima"])
# #                     self._navegante.obtener()
# #                     continue

# #                 elif self.matrizFechas[self._propiocepcion]==201106:
# #                     self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["penultima"])
# #                     self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["fecha_bottom"])             
# #                     self._navegante.obtener()
# #                     break 