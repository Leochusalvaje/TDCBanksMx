# INICIACIÓN DEL LOGGING
# ------------------------------------------------------------------------------
# Para que el logging capture correctamente los eventos de todos los módulos
# (especialmente decoradores de clase como @crear_directorios y configuraciones
# que se ejecutan al importar), la configuración de logging.basicConfig DEBE ser
# lo PRIMERO que se ejecute en el script principal, antes de los imports.
#
# Se usa 'force=True' para asegurar que esta configuración sea la mande mayor jerarquia y 
# no sea bloqueada por imports previos o librerías externas.
# ==============================================================================
from __future__ import annotations
import logging
logger=logging.getLogger("Automata")

from dataclasses import dataclass
import win32com.client as win32
import pyautogui
import time
import win32gui, win32con
import routes as gv
from pathlib import Path
import pywintypes as wty
import re
from typing import Any,TYPE_CHECKING
from config import RegexpPatrones, TipoPatron

from excel_manager import tareaPendiente,ExcelManager,carajo

if TYPE_CHECKING:
    from config import RegexpPatrones

pyautogui.FAILSAFE = True


CONFIG_COORDS = {
    "abajo": (362, 425),
    "arriba": (362, 388), 
    "bancos": (393, 363),
    "obtener": (76, 409), 
    "fecha_top": (155, 387),
    "fecha_bottom": (155, 423),
    "penultima": (156, 404),
    "tiempo_espera_consulta": 1,
    "tiempo_espera_entre_clicks":0.3
}

coord_abajo=(362,425)
coord_arriba=(362,388)
coord_banco = (393, 363)
coord_obtener = (76, 409)
coord_desmarcar=(155,387)
coord_fecha=(155,423)
coord_ultima=(155,385)
coord_penultima=(156,404)

#Funciones auxialiares
def activar_ventana_al_frente(hwnd, timeout=5)-> None:
    start = time.time()
    while time.time() - start < timeout:
        if hwnd:

            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)             
            win32gui.SetForegroundWindow(hwnd) 
            logger.info(f"Ventana {hwnd} traida al frente y maximizada con exito")
            break
        time.sleep(0.2)
    else:
        logger.critical(f"No se pudo activar/maximizar la ventana con hwnd {hwnd}, asgeurar que no se esten ejecutando procesos externos durante al ejecucion")


def extraer_datos(celda:Any,rutadestino:Path,ventanaexcel)->None:
    rango:Any=celda.CurrentRegion()
    rutadellibrocontrol:Path=gv.Paths.output/"LibroControl.xlsx"
    nombreventana:str=rutadellibrocontrol.stem
    try:
        libroabierto=ventanaexcel.Workbooks(nombreventana)
        libroabierto.Close(SaveChanges=False)
        print("Se cierra la ventana previa del LibroControl si es que esta abierto para evitar inyectar informacion erronea")
    except:
        pass


    librocontrol=ventanaexcel.Workbooks.Open(str(rutadellibrocontrol))
    librocontrol.Activate()

    filas=len(rango)
    columnas=len(rango[0])
    hoja=librocontrol.Sheets(1)

    print(type(rango))
    celdas_destino=hoja.Range(hoja.Cells(1, 1),hoja.Cells(filas, columnas))
    celdas_destino.Value=rango
    try:
        ventanaexcel.DisplayAlerts = False
        librocontrol.SaveAs(str(rutadestino))
        librocontrol.Close(SaveChanges=False)
    except Exception as e:
        print(f"Error en el metodo SaveAs , no se pudo guardar {type(e).__name__}: {e}\n")
    finally:
            if librocontrol is not None:
                try:
                    ventanaexcel.DisplayAlerts = False 
                    librocontrol.Close(SaveChanges=False)
                except:
                    ventanaexcel.DisplayAlerts = True
                    pass     
    return None

def string_eureka_fechas(libro:Any,patronValidador:RegexpPatrones)->tuple:
    
    primerhoja=libro.Sheets(1)
    primerhoja.Activate()

    buscar_1=primerhoja.Range("B1").End(-4121).End(-4121)

    if not primerhoja.Range("B1").End(-4121).Value=="Notas":
        raise Exception("\nCambio el formato de la CNBV actualizar saltos de linea y espaciados de nuevo para encontrar la tabla, se debe modificar los pasos de la libreria win32com.client.")

    a=buscar_1.CurrentRegion
    a.Activate()
    tupla_datos=a.Value

    if not len(tupla_datos[0])>=2:
        raise Exception("\nEs muy probable que cambio la formato de la CNBV pues no se esta alcanzando la tabla de datos para su extraccion.")
    try:
    #celda=primerhoja.Cell
        str_fecha=str(int(tupla_datos[0][2]))

        print("eureka")        
        print("Formato estandar detectado no se necesito aplicar logica adicional")
    except (IndexError, ValueError, TypeError):
        str_fecha = str(int(tupla_datos[1][2]))
        print("Formato alternativo detectado")

    if not re.match(patronValidador.patrones["fecha"],str_fecha):
        raise Exception("\nEl valor para buscar la fecha y nombrar las carpetas no se cumple. Validar la celda de donde se esta extrayendo la fecha.")


    str_fecha=str_fecha
    celda_tabla=buscar_1
    hayMasDeUnBancoEnLaHoja=len(tupla_datos[0])>3

    return str_fecha,celda_tabla,hayMasDeUnBancoEnLaHoja


# "lista_verdad" esta función se encarga de generar la lista de fechas que se encuentran en cada archivo excel de la CNBV, esta lista es crucial para el funcionamiento del automata 
# ya que es la que se utiliza para navegar por las fechas dentro del archivo excel y para validar que se esta extrayendo la información correcta, si esta lista no se 
# genera correctamente el automata no funcionara correctamente, por lo que es importante revisar que se este generando correctamente y que se este cumpliendo el patron 
# establecido en la clase RegexpPatrones.
def lista_verdad(bimestreInicial:str,bimestreFinal,patronValidador:RegexpPatrones)->list:
    
    fecha=patronValidador.revisar_patron(TipoPatron.FECHA,bimestreFinal)

    try:
        yearbase=int(bimestreInicial[:4])
        yearUltimobimestre=int(fecha[:4])
        yearsTranscurridos=yearUltimobimestre-yearbase
        bimestreActual=(14-int(fecha[-2:]))/2
    except (ValueError, TypeError) as e:
        # Usamos .exception para que guarde el rastro (traceback)
        # Y metemos type(e).__name__ para que el mensaje claro
        logger.exception(f"[{type(e).__name__}] Error fatal en fechas: {bimestreInicial} y {bimestreFinal} revisar que se esten ingresando como strings. Detalle: {e}")
        # Esto es lo que hace que el programa "muera" o falle rápido
        raise

    bimestreBase=int(bimestreInicial[-2:])/2
    bimestresEnArchivo=int(6*(yearsTranscurridos-1)+(bimestreBase+bimestreActual))
    logger.info(f"Bimestres contenidos en cada archivo excel de la CNBV: {bimestresEnArchivo} Este parametro es crucial y se genera con los parametros del config.json ingresados por el USUARIO humano para el funcionamiento del automata, si no se esta extrayendo la cantidad correcta de bimestres revisar que las fechas de entrada sean correctas y que se este cumpliendo el patron establecido en la clase RegexpPatrones")
    #para resolver el seguimiento de fechas se solcuiona con una transformacion lineal que planteo su servidor y tener una lista de todos los bimestres 
    #que estan contenidos en el archivo excel al momento de abrirlo
    y=[]
    yearconstante=int(bimestreInicial)
    c=0
    for _ in range(1,bimestresEnArchivo+1):

        yearconstante+=2    
        if (yearconstante-201100-100*c)==14:
            c+=1 
            yearconstante+=100
            yearconstante-=12

        y.append(yearconstante)
    y.reverse()


    return y
"""
boton es la clase para abtarer un boton y que el automata pueda tener una referencia mental de cada boton y su estado, 
esto es crucial para el funcionamiento del automata ya que sin esta clase el automata no tendria forma de saber si un boton esta activo o no y no 
podria tomar decisiones en base a eso, ademas de que al tener una referencia mental de cada boton se puede implementar logica adicional 
como apagar todos los botones encendidos o invertir el estado de varios botones a la vez, lo que facilita la navegacion 
por las fechas dentro del archivo excel y la extraccion de la informacion correcta.
"""

class boton:
    def __init__(self,fecha:int,posicion:int)->None:
        self.fecha=fecha
        self.posicion=posicion
        self.estado=False
    @property
    def estado_actual(self)->bool:
        return self.estado
    
    def cambiar_estado(self)->None:
        self.estado=not self.estado



"""
coordenadaSimple es una clase que usa la libreria pyautogui para hacer click en una coordenada especifica,
esta clase se hizo para resolver el problema de que el automata no tiene forma de saber si una cordenada dentro
del checkbox y tarbajar con la coordenada de obtener esta activa o no adicionalmente se deeja la logica de la manipulacion del click en esta funcion por si en el futuro
se decide agregar una libreria diferente para un entorno que no sa windows.
"""
class coordenadaSimple:

    def __init__(self,coordenada:tuple,intervalo:int)->None:
        self.coordenada=coordenada
        self.intervalo=intervalo 

    def click(self,boton:boton|None=None,tiempo_espera:float|None=None,numero_clicks:int|None=1)->None:

        pyautogui.click(self.coordenada,interval=self.intervalo,clicks=numero_clicks)

        if tiempo_espera:
            time.sleep(tiempo_espera)

        if boton:
            boton.cambiar_estado()

"""
CheckerCoordenada es un esquema de datos para almacenar las coordenadas de los botones y cordenadas que el automata va a utilizar para navegar 
en el archivo excel de la CNBV.
"""
@dataclass
class CheckerCoordenada:
    abajo: tuple
    arriba: tuple
    bancos: tuple
    obtener: tuple
    fecha_top: tuple
    fecha_bottom: tuple
    penultima: tuple

    def __post_init__(self):
        self.abajo = tuple(self.abajo)
        self.arriba = tuple(self.arriba)
        self.bancos = tuple(self.bancos)
        self.obtener = tuple(self.obtener)
        self.fecha_top = tuple(self.fecha_top)
        self.fecha_bottom = tuple(self.fecha_bottom)
        self.penultima = tuple(self.penultima)

#asignamios los tiempos y la configuracion de cordenadas que se cargan desde el config.json a la clase ConfigVisual para que el automata pueda acceder a ellos de forma centralizada y 
#organizada, ademas de que al cargar la configuracion desde el config.json se puede modificar facilmente sin necesidad de tocar el codigo del automata,
# lo que facilita la adaptacion a diferentes pantallas o cambios en la interfaz de la CNBV.
@dataclass
class TiemposConfig:
    tiempo_espera_consulta: float
    tiempo_espera_entre_clicks: float

@dataclass
class ConfigVisual:
    coordenadas: CheckerCoordenada
    tiempos: TiemposConfig



class PointCheckManagerPrime:
    def __init__(self,config_visual:ConfigVisual):
        self.despachador=config_visual
        self.tiempo_clicks=config_visual.tiempos.tiempo_espera_entre_clicks
        self.tiempo_consulta=config_visual.tiempos.tiempo_espera_consulta


    def click_auxiliar_en_cordenada(self,punto:tuple)->None:
        click=coordenadaSimple(punto,self.tiempo_clicks)
        click.click()
    
    def click_en_coordenada_relativista(self,boton:boton)->None:
        click=coordenadaSimple(self.despachador.coordenadas.fecha_top,self.tiempo_clicks)
        click.click(boton=boton)

    def siguiente_coordenada(self,numero_clicks:int)->None:
        click=coordenadaSimple(self.despachador.coordenadas.abajo,self.tiempo_clicks)
        click.click(numero_clicks=abs(numero_clicks))
    
    def anterior_coordenada(self,numero_clicks:int)->None:
        click=coordenadaSimple(self.despachador.coordenadas.arriba,self.tiempo_clicks)
        click.click(numero_clicks=abs(numero_clicks))
    
    def obtener(self)->None:
        click=coordenadaSimple(self.despachador.coordenadas.obtener,self.tiempo_clicks)
        click.click(tiempo_espera=self.despachador.tiempos.tiempo_espera_consulta)
        
class pointCheckboxManager(coordenadaSimple):

    def __init__(self,config_visual:ConfigVisual)->None:
        self._despachador=config_visual

        super().__init__(self._despachador.coordenadas.fecha_top, self._despachador.tiempos.tiempo_espera_entre_clicks)
        
    def click_auxiliar_en_coordenada(self,punto:tuple)->None:
        pyautogui.click(punto,interval=self._despachador.tiempos.tiempo_espera_entre_clicks)
        time.sleep(self._despachador.tiempos.tiempo_espera_consulta)   

    def siguiente_cordenada(self,pasos:int):
        pyautogui.click(self._despachador.coordenadas.abajo,clicks=abs(pasos),interval=self._despachador.tiempos.tiempo_espera_entre_clicks)

    def  anterior_cordenada(self,pasos:int):
        pyautogui.click(self._despachador.coordenadas.arriba,clicks=abs(pasos),interval=self._despachador.tiempos.tiempo_espera_entre_clicks)


    def obtener(self)->bool:
        pyautogui.click(self._despachador.coordenadas.obtener)
        time.sleep(self._despachador.tiempos.tiempo_espera_consulta)

class AutomataPrime:
    def __init__(self,tarea:tareaPendiente,manipulacion:ExcelManager,navegante:PointCheckManagerPrime)->None:
        self.tarea=tarea
        self.manipualcion=manipulacion
        self.matrizBotones:dict[int,boton]={}
        self.navegante=navegante

class automataConsultas:
    def __init__(self,punto_actual:pointCheckboxManager,obervador:carajo,ventanaExcel:Any)->None:
        #Esta es una ventana y un objeto para manipualr el excel posterior a esto si obtendremos el handle
        self.ventanaExcel=ventanaExcel
        self.hdwm_excel=ventanaExcel.Hwnd
        self.observador=obervador
        #iniciamos la matriz de fechas que hay dentro del archivo excel que se esta consultando
        self.matrizFechas=self.observador.listaFechas
        #con las fechas creamos botones con todos inicializados en falso, osea desmarcados y solo el primero queda marcado tal y como esta el estandar
        #del archivo excel, la fecha mas reciente esta marcada y consultada
        self.matrizBotones:dict[int,boton]={}
        self._navegante=punto_actual

        self._propiocepcion = 0
        #obtenemos la hoja excel del observador
        self.hojaExcel=self.observador.hojaExcel

        self.propiocepcion_actual

        try:
            self.crear_botones()
            # Como la primer fecha siemproe esta marcada el primer boton queda activo tal como los archivos
            #accedemeos manualmenete al dictioanro con el metodo iter para ir uno a uno y para el primero encapsualmos en un next
            # se implemnta nueva busqueda por la hoja excel en vez de inicializar las variables, no descarto la inicializacion de variables de momento se quitara
            primer_boton = next(iter(self.matrizBotones.values()))
            primer_boton.estado = True
        except Exception as e:
            print(f"Error critico no se pudieron crear los botones para guiar al automata, revisar la propiedad \"self.matrizFechas\"\n{e}")
            self.matrizBotones={}

        #con lo que hacemos la propiocepcion es con la fecha que se obtiene al iniciar el automata
        self.ojos=string_eureka_fechas(self.hojaExcel)[0]
    
    def crear_botones(self):


        for k, fechas in enumerate(self.matrizFechas):
            x=boton(fechas,k)
            self.matrizBotones[fechas]=x
        
    def avanzar_a_fecha(self,pasos:int)->None:

        if pasos == 0:
            return

        if pasos>0:
            self._navegante.siguiente_cordenada(pasos)
        elif pasos<0:
            self._navegante.anterior_cordenada(pasos)

        self._propiocepcion+=pasos
        logger.debug("propiocepcion",self._propiocepcion)


    
    def ir_apagar_boton(self,boton:boton):

        if boton.estado:
            self.avanzar_a_fecha(boton.posicion-self._propiocepcion)
            self._navegante.click_en_coordenada(boton)
        
    def ir_apagar_todos_los_botones_encendidos(self)->None:
        for boton in self.matrizBotones.values():
            if boton.estado:
                self.ir_apagar_boton(boton)

    def invertir_estado_varios_botones(self,cambios:list)->None:
        for botonFecha in cambios:
            self.matrizBotones[botonFecha].cambiar_estado()

    @property            
    def botones_encendidos(self)->list:
        listaFechas=[]
        for boton in self.matrizBotones.values():
            if boton.estado:
                listaFechas.append(boton.fecha)
        return listaFechas

    @property
    def propiocepcion_actual(self)->int:
        print (self._propiocepcion)
        return self._propiocepcion

    def homming(self)->None:
        activar_ventana_al_frente(self.hdwm_excel)
        self.avanzar_a_fecha(-self._propiocepcion)
        self._navegante.click_en_coordenada()

    def validar_fecha_actual_y_extraer(self,fechaRevision:int,celda)->bool:

        try:
            
            print(f"fecha actual obtenida: {fechaRevision}, si coincide con la propiacepcion {self.matrizFechas[self._propiocepcion]}")
            print("fecha coincide con la actual, obtener")
            if not fechaRevision==self.matrizFechas[self._propiocepcion]: 
                print("Implemnetando nueva logica")

                return True
            else: 
                archivo=f"{self.observador.dummy_archivos[self._propiocepcion]}.xlsx"
                rutaArchivo=self.observador.mapa.carpetaDestino/archivo
                extraer_datos(celda,rutaArchivo,self.ventanaExcel)
                
                return False
            

        except IndexError as e:
            print(f"IndexError en obtener fecha actual, probablemente la propiocepcion se salio de rango No es un error critico continua {e}")
            return False
        

    def recorrer_pendientes(self,lisapendientes:list)->None:
        print(self.matrizFechas)
        print("lista de pendientes a consultar",lisapendientes)

        #Configuracion para marcar todos los bancos al inicio de la consulta
        self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["bancos"])
        self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["bancos"])
        self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["obtener"])
        #Manejo de caso atipico para marcar todos los bancos
        if not string_eureka_fechas(self.hojaExcel)[2]:
            self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["bancos"])


        cicloRoto=False
        posicion=0
        for k,indice in enumerate(lisapendientes):

            if k==0 and self.matrizFechas[k] not in [self.matrizFechas[i] for i in lisapendientes]:
                self.ir_apagar_boton(self.matrizBotones[self.matrizFechas[k]])


            self.avanzar_a_fecha(indice-self._propiocepcion)
            fechaEnBucle=self.matrizFechas[indice]
            
            #validamos que el boton no es te seleccionado  
            if not self.matrizBotones[fechaEnBucle].estado:
                self._navegante.click_en_coordenada(boton=self.matrizBotones[fechaEnBucle])


            self._navegante.obtener()
            fechaEnHoja,celda,*_=string_eureka_fechas(self.hojaExcel)
            fechaEnHoja=int(fechaEnHoja)
            

            if self.validar_fecha_actual_y_extraer(fechaEnHoja,celda):  
                cicloRoto=True


                ajustePropiocepcion=self.matrizBotones[fechaEnHoja].posicion-self._propiocepcion
                self._propiocepcion+=ajustePropiocepcion

                #cuando se desajusta para encontrar el boton que se encendio por error, es la diferencia entre el anterior indice menos el indice actual 
                #de la lista de pendientes en la depuracion de una lista ejemplo [3,7,11], pasos=7-3=4 que es el borton que se encendio por error 
                pasos=lisapendientes[k]-lisapendientes[k-1]
                self.invertir_estado_varios_botones([fechaEnHoja,fechaEnBucle,self.matrizFechas[lisapendientes[k-1]],self.matrizFechas[pasos]])
                self.ir_apagar_todos_los_botones_encendidos()

                print("entro al caso particular")
                break

            self.ir_apagar_boton(self.matrizBotones[fechaEnHoja])
            posicion+=1   
            print("propiocepcion despues de avanzar",self._propiocepcion)
        print("fecha de cambio",fechaEnHoja)

        if cicloRoto:
            fechaQueRompeElBucle=fechaEnHoja

            for indice in lisapendientes[posicion:]:
                if indice==len(self.matrizFechas)-2:
                    self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["penultima"])
                    self._navegante.obtener()
                    fechaEnHoja,celda,*_=string_eureka_fechas(self.hojaExcel)
                    fechaEnHoja=int(fechaEnHoja)
                    self.validar_fecha_actual_y_extraer(fechaEnHoja,celda) 
                    self._propiocepcion-=self.matrizBotones[fechaQueRompeElBucle].posicion+self._propiocepcion
                    continue
                if indice==len(self.matrizFechas)-1:
                    self.avanzar_a_fecha(indice-self._propiocepcion)
                    self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["penultima"])
                    self._navegante.click_auxiliar_en_coordenada(self._navegante._despachador["fecha_bottom"])             
                    self._navegante.obtener()
                    fechaEnHoja,celda,*_=string_eureka_fechas(self.hojaExcel)
                    fechaEnHoja=int(fechaEnHoja)
                    self.validar_fecha_actual_y_extraer(fechaEnHoja,celda) 
                    self._propiocepcion-=self.matrizBotones[fechaQueRompeElBucle].posicion+self._propiocepcion
                    continue                 
                fechaEnBucle=self.matrizFechas[indice]
                #Pruebas para ajustar la propiocepcion con ayuda del debugger , simular los clciks manualmente mientras se testea
                #print(indice)
                #print(self.observador.listaFechas[-self._propiocepcion])
                if indice==11:
                    print("para revisar con el debugger")
                # 201406, 201310, 201302,201206,
                self.avanzar_a_fecha(indice-self._propiocepcion)
                self._navegante.click_en_coordenada(boton=self.matrizBotones[fechaEnBucle])
                
                #print(self.observador.listaFechas[-self._propiocepcion])
                if indice<len(self.matrizFechas)-2:

                    self._navegante.obtener()
                    fechaEnHoja,celda,*_=string_eureka_fechas(self.hojaExcel)
                    fechaEnHoja=int(fechaEnHoja)
                    self.validar_fecha_actual_y_extraer(fechaEnHoja,celda) 
                    self._propiocepcion-=self.matrizBotones[fechaQueRompeElBucle].posicion+self._propiocepcion
                    
                    self.ir_apagar_boton(self.matrizBotones[fechaEnHoja])
        self.homming()
        self.observador.hojaExcel.Close(SaveChanges=False)
        self.ventanaExcel.DisplayAlerts = False
        self.ventanaExcel.Quit()  

        print("Proceso de consultas finalizado")

# pat=RegexpPatrones()
# rutaPrueba=gv.RUTA_BASE/ "scripts"/ "prueba.xlsx"

# x=gv.ARCHIVOS_DATA_RAW[3]["ruta"]
# prueba=ExcelManager(pat)
# prueba._abrir_instancia_excel()

# prueba.traer_instancia_excel_al_frente()
# prueba.traer_instancia_excel_al_frente()
# prueba.abrir_libro("libro1",x)
# libro=prueba.libro["libro1"]

# prueba.traer_libro_al_frente("libro1")


# prueba.extraer_tabla_y_guardar_en_ruta("libro1",rutaPrueba,gv.RUTA_LIBRO_CONTROL)

# prueba.cerrar_libro_excel("libro1")
# time.sleep(2)   
# prueba.ventana_excel.Quit()
# prueba.ventana_excel = None
# casilla1=pointCheckboxManager(CONFIG_COORDS)
# mapa=rutaCarpetas(rutasInstaciadas,x)

# libro,ventana,*_=abrir_excel_por_ruta(x)


# observador=carajo(libro,mapa)
# observador.archivosCreados

# hoja1=automataConsultas(casilla1,observador,ventana)



# hoja1.recorrer_pendientes([0])


# print(hoja1.matrizFechas)

# ################################### 
# Implementacion de decoradores
# def mi_primer_decorador(fechaBuscada:str):

#     def decorador_real(funcion):

#             def envoltura(*args, **kwargs):
#                 pass




