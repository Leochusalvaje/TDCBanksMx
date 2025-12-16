import pandas as pd
import numpy as np
import openpyxl
import os
import routes as gv
from typing import Optional, Callable,Any


#tipos de datos para tener mejor control como en el lenguaje de c++
dtf=pd.DataFrame
dts=pd.Series
# df.iloc[fila, columna] -> Accede al valor por índice de fila y columna (0-indexed)
# df["Columna"].iloc[fila] -> Accede al valor por nombre de columna y posición de fila
# df.iat[fila, columna] -> Acceso rápido a un solo valor por índice de fila y columna
#drop() elimina una columan
#banks=lambda x:x["banks"].astype("Int32")  es uan funciona anonima para declarar el tipo de datos de la columna 
#Float32
#Prueba con decorador fallida complica mas las cosas !!!!°!
#Prueba con decorador para asignar nombres a a las columnas 
# def endulcoradorparacambiarnombreytipodecolumna(nombreColumna,posciionColumna,tipoDato):

#     def endulcorador1(funcion):
#         @wraps(funcion) # Buenas prácticas: Preservar metadatos
#         def endulcorador2(*args, **kwargs):
            
#             # 1. Obtener el DataFrame de la función original
#             df = funcion(*args, **kwargs)
#             nombre_actual = df.columns[posciionColumna]
#             df_modificado = (
#                 df.rename(columns={nombre_actual: nombreColumna})

#                 .assign(**{
#                     nombreColumna: lambda x: pd.to_numeric(x[nombreColumna], errors='coerce').astype(tipoDato)
#                 })
#             )
#             return df_modificado
            
#         return endulcorador2
#     return endulcorador1

# Esta funcion facilitara la asignacion de datos al final del dataframe por columna 
def recorrer_asignar_tipos(df:dtf,*tiposdedatos,**tipospordiccionario)->dtf:

    columnas=[nombrescol for nombrescol in df]
    if tiposdedatos:
        if len(tiposdedatos) != len(columnas):
            raise ValueError("El numero de tipos debe coincidir con el numero de columans ")
        creardictioanrio=dict(zip(columnas,tiposdedatos))
        return df.astype(creardictioanrio)
    if tipospordiccionario:
        columnas_faltantes = set(tipospordiccionario) - set(df.columns)
        if columnas_faltantes:
            raise KeyError(f"Columnas inexistentes: {columnas_faltantes}")
        return df.astype(tipospordiccionario)

    




def limpiar_datos_de_carpeta_040_12d_R1(rutaArchivo, 
                                        transformacion_adicional: Optional[Callable] = None):

    df = pd.read_excel(rutaArchivo, engine="openpyxl")  # usa engine="xlrd" si es .xls


    df_limpio=df.drop(columns=["Unnamed: 0"]) \
                .rename(columns={"Banco":"banks"})\
                .assign(year=str(df.columns[2])[:4],
                        month=str(df.columns[2])[-2:],)\
                .assign(
#asignacion de tipo de dato de columnas


                        year=lambda x:x["year"].astype("Int32"),
                        month=lambda x:x["month"].astype("Int32"),
                        banks=lambda x:x["banks"].astype(str),)\
                .iloc[:-1]
    
    if transformacion_adicional:
        # Usamos .pipe() para integrar la función adicional en la cadena.
        # Esto le da el DataFrame df_limpio como primer argumento.
        df_limpio = df_limpio.pipe(transformacion_adicional)

    return df_limpio


#Esta funcion va a retornar una Dataframe juntando todo el contenido preprocesado con la funcion de limpieza
def limpiar_datos_de_carpeta(funcionlimpieza,rutarchivo,**kwargs):
    df_unido=pd.DataFrame()

    for archivo in rutarchivo:
        k=funcionlimpieza(archivo,**kwargs)
        df_unido=pd.concat([df_unido,k],ignore_index=True)
    
    return df_unido
###################

def pipeline(*args: Callable) -> Callable:

    def salida(datos_iniciales: Any):
        resultado = datos_iniciales
        
        for func in args:
            resultado = func(resultado) 
            
        return resultado
    return salida

def asignar_y_tipificar(df, nombre_columna_nuevo, posicion_columna, tipo_dato):
 
    nombre_actual = df.columns[posicion_columna]
    df=df.rename(columns={nombre_actual:nombre_columna_nuevo}



    ).assign()


    return (
        df.rename(columns={nombre_actual: nombre_columna_nuevo})
        .assign(**{
            nombre_columna_nuevo: lambda x: pd.to_numeric(x[nombre_columna_nuevo], errors='coerce').astype(tipo_dato)
        })
    )

def transformar_monto(df):
    return asignar_y_tipificar(df, "quantyCreditCards", 1, "Int32")
            
# importantge no borrar------------------------------------------> carpeta 1 Limpia

# Numero_de_tarjetas_de_credito_por_institucion = limpiar_datos_de_carpeta(
#     limpiar_datos_de_carpeta_040_12d_R1,
#     gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[0],
#     transformacion_adicional=transformar_monto
# )
# importantge no borrar------------------------------------------>





#funcion paraobtenr las columans de año y mes acepata dos argumentos optionales que le pasemos la posicuion de la fecha si esta en la cabecera de una columna o por localizacion de indice en el excel


def crear_columna_year_month(
    df:dtf,
    indexcolumnafecha: Optional[int] = None,
    collumnafechaexiste: Optional[int]=None
)->dtf:

    if indexcolumnafecha is not None:

        df=df.assign(
        Date=str(df.columns[indexcolumnafecha]),
        Year=str(df.columns[indexcolumnafecha])[:4],
        Month=str(df.columns[indexcolumnafecha])[-2:]

        ). assign(

            Year=lambda x: x["Year"].astype("int16"),
            Month=lambda x:x["Month"].astype("int8"),
            Bimester=lambda x:(x["Month"] - 1) // 2 + 1,
            Date=lambda x: pd.PeriodIndex( x["Date"],freq="M")

        ).assign(Bimester=lambda x:x["Bimester"].astype("int8"))


        
    elif collumnafechaexiste is not None:
        posicion=collumnafechaexiste
        nombrecolumna=df.columns[posicion]
        year = df[nombrecolumna].str[:4].astype("int16")
        month = df[nombrecolumna].str[4:].astype("int8")
        
        df.insert(posicion + 1, "Year", year)
        df.insert(posicion + 2, "Month", month)
        bim=((df["Month"] - 1) // 2 + 1).astype("int8")

        df.insert(posicion + 3, "Bimester", bim) 
        df[nombrecolumna]=pd.PeriodIndex(df["Date"], freq="M")

    else:
        raise ValueError("Debes pasar indexcolumnafecha o collumnafechaexiste.")


    return df

def limpiar_datos_de_carpeta_040_12d_R2(
        #Ruta donde se encuentra el archivo a limpiar
        ruta:str,
        #Nombre de la columna que vamos obtener la informacion
        nombrecolumnaenarchivo:str,
        #Columnas que se van exportar desde el readexcel ejemplo "A:Z"
        colexcelexportar:str,
        #Trasnformacion adicional en caso de ser necesario se debe definir y crear una funcion invocable que retorne un Dataframe de pandas
        transformacion_adicional:Optional[callable]=None,
        #De ser necesario se ponen una tupla en orden de aparicion de las columnas para assignar  un tipo de dato a cada columna
        *tiposDatosOrdenCol,
        #Por medio de un diccionario le pasamos el nombre de una columna que sabemos existe en dataframe para assignarle un tipo de dato
        **datospordiccionario):
    
    df=pd.read_excel(ruta,usecols=colexcelexportar).iloc[:-1]
    df=crear_columna_year_month(df,1)

    nombrescolumnas=list(df.columns)
    df=df.rename(columns={nombrescolumnas[1]:nombrecolumnaenarchivo,"Banco":"Banks"}).fillna(0)


    extraer=df.pop(nombrecolumnaenarchivo)
    #insertar
    df.insert(5,nombrecolumnaenarchivo,extraer)
    if transformacion_adicional:
        df = df.pipe(transformacion_adicional)

    if tiposDatosOrdenCol or datospordiccionario:
        df=recorrer_asignar_tipos(df,*tiposDatosOrdenCol,**datospordiccionario)
    return df



def abc(df):
   return asignar_y_tipificar(df,"creditCardBalanceByInstitution",1,"Float32")


##Carpeta 2 Limpia
# unido=limpiar_datos_de_carpeta(limpiar_datos_de_carpeta_040_12d_R2,gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[1],transformacion_adicional=abc).fillna(0).rename(columns={"Banco":"banks"})
# unido.head(5)

def limpiardatsetslargos(ruta:str,

colexcelexportar:str,
nombreIntervalo:str,
*tiposDatosOrdenCol,
**datospordiccionario
)->dtf:


    #Se lee la informacion y se asigna a un data frame, se seleccionan las columnas que va a leer y se trasponen los datos para asignar un formato largo 
    df=pd.read_excel(ruta,usecols=colexcelexportar).T.reset_index()
    #Elimino la columna  final
    ultima_col = df.columns[-1]
    df = df.drop(columns=ultima_col)
    #se asignan componentes especificos de la matriz el valor necesario para la limpieza
    df.iloc[0,1]="Date"
    df.iloc[0,0]="Banks"
    df.columns=df.iloc[0]

    #se elimina una fila especifica del dataframe
    df=df.drop(index=0)
    df=df.melt(
    id_vars=['Banks','Date'],
    var_name=nombreIntervalo,
    value_name="Number_Of_Creditcards"
    ).assign(
    Date=lambda x:x["Date"].astype(int).astype(str),
    Number_Of_Creditcards=lambda x:x["Number_Of_Creditcards"].astype("Int64")

    )
    df=crear_columna_year_month(df,collumnafechaexiste=1).fillna(0)
    if tiposDatosOrdenCol or datospordiccionario:
        df=recorrer_asignar_tipos(df,*tiposDatosOrdenCol,**datospordiccionario)
    df.head()
    return df


#Pruebas
#-------------------------------------------------------------------------------------------------------------------------------------->

#Limpiar datos carpeta 3
#x=limpiardatsetslargos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[2][0],"B:AO","Credit_Limit_Range_KMXN",Credit_Limit_Range_KMXN="category")
#Limpiar Datos carpeta 2
#x=limpiar_datos_de_carpeta_040_12d_R2(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[1][0],"Credit_Card_Balance","B:C",Credit_Card_Balance="float64")
#Limpiar datos carpeta 1
#x=limpiar_datos_de_carpeta_040_12d_R2(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[0][0],"Number_Of_Creditcards","B:C",Number_Of_Creditcards="int")

#Limpiar datos carpeta 4
#Credit_Utilization_Range
#x=limpiardatsetslargos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[3][0],"B:AO","Credit_Utilization_Range")
x=pd.read_excel(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[3][0])
x.head()