import pandas as pd
import numpy as np
import tdcbanksmx.routes as gv
from typing import Optional, Callable,Any 


#tipos de datos para tener mejor control como en el lenguaje de c++
dtf=pd.DataFrame
dts=pd.Series
logicaDatos=Callable[..., pd.DataFrame]

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

#Algunas conevrsiones de columna pueden fallar por lo que esta funcion intenta convertir a Int64 y si falla lo deja en Float64
def conversion_segura(series):
    try:

        return series.astype("Int64")
    except:
        # Si falla, convertimos a Float64
        return series.astype("Float64")    




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

def limpiar_datos_largos(
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
# unido=limpiar_datos_de_carpeta(limpiar_datos_largos,gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[1],transformacion_adicional=abc).fillna(0).rename(columns={"Banco":"banks"})
# unido.head(5)

def limpiar_datos_anchos(ruta:str,

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
    Number_Of_Creditcards=lambda x: conversion_segura(x["Number_Of_Creditcards"])

    )
    df=crear_columna_year_month(df,collumnafechaexiste=1).fillna(0)
    if tiposDatosOrdenCol or datospordiccionario:
        df=recorrer_asignar_tipos(df,*tiposDatosOrdenCol,**datospordiccionario)
    df.head()
    return df


#Pruebas
#-------------------------------------------------------------------------------------------------------------------------------------->

#Limpiar datos carpeta 3  040_12e_R10_Distribucion_de_tarjetas_por_probabilidad_de_incumplimiento_201106.xlsx
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[2][0],"B:AO","Credit_Limit_Range_KMXN",Credit_Limit_Range_KMXN="category")
#Limpiar Datos carpeta 2 040_12d_R2_Saldo_de_tarjetas_de_credito_por_institucion_201106.xlsx
#x=limpiar_datos_largos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[1][0],"Credit_Card_Balance","B:C",Credit_Card_Balance="float64")
#Limpiar datos carpeta 1 040_12d_R1_Numero_de_tarjetas_de_credito_por_institucion_201106.xlsx
#x=limpiar_datos_largos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[0][0],"Number_Of_Creditcards","B:C",Number_Of_Creditcards="int")
#Limpiar datos carpeta 4 040_12e_R11_Distribucion_de_tarjetas_por_impagos_consecutivos_201106.xlsx
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[3][0],"B:AO","Consecutive_Delinquency_Bucket",Consecutive_Delinquency_Bucket="category")
#Limpiar datos carpeta 5 040_12e_R1_Distribucion_de_tarjetas_por_limite_de_credito_201106.xlsx Credit_Limit_Range_KMXN
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[4][0],"B:AO","Credit_Limit_Range_KMXN",Credit_Limit_Range_KMXN="category")
#Limpiar datos carpeta 6 040_12e_R2_Distribucion_de_tarjetas_por_porcentaje_de_uso_de_linea_201106.xlsx Credit_Line_Utilization_Range_Pct
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[5][0],"B:AO","Credit_Line_Utilization_Range_Pct",Credit_Line_Utilization_Range_Pct="category")
#Limpiar datos carpeta 7 040_12e_R3_Porcentaje_de_pago_minimo_exigido_con_respecto_al_saldo_a_pagar_201106.xlsx
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[6][0],"B:AO","Minimum_Payment_Percentage",Minimum_Payment_Percentage="category")
#Limpiar datos carpeta 8 040_12e_R4_Porcentaje_de_pago_realizado_con_respecto_al_saldo_a_pagar_201106.xlsx
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[7][0],"B:AO","Actual_Payment_Percentage",Actual_Payment_Percentage="category")
#x.head()
#Limpiar datos carpeta 9 040_12e_R50_Distribucion_de_tarjetas_por_porcentaje_de_pago_minimo_respecto_a_la_linea_de_credito_201106.xlsx
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[8][0],"B:AO","Minimum_Payment_Percentage_Range",Minimum_Payment_Percentage_Range="category")
#Limpiar datos carpeta 10 040_12e_R5_Porcentaje_de_pago_realizado_vs_pago_minimo_exigido_201106.xlsx
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[9][0],"B:AO","Actual_vs_Minimum_Payment_Percentage",Actual_vs_Minimum_Payment_Percentage="category")
#Limpiar datos carpeta 11 040_12e_R6_ConsumoRevolvente_TarjetasPagoSinIntereses_201106.xlsx
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[10][0],"B:AO","Minimum_Payment_Ratio_Interval",Minimum_Payment_Ratio_Interval="category")
#Limpiar datos carpeta 12 040_12e_R7_Consumo_Revolvente_Porcentaje_pago_realizado_vs_PPNGI_201106.xlsx
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[11][0],"B:AO","Payment_Balance_Ratio_Interval",Payment_Balance_Ratio_Interval="category")
#Limpiar datos carpeta 13 040_12e_R8_Consumo_Revolvente_desde_la_apertura_de_la_cuenta_201106.xlsx
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[12][0],"B:AO","Months_Since_Opening",Months_Since_Opening="category")
#Limpiar datos carpeta 14 040_12e_R9_Consumo_Revolvente_Distribucion_de_tarjetas_por_perdida_esperada_201106.xlsx
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[13][0],"B:AO","Expected_Loss_Range",Expected_Loss_Range="category")
#Limpiar datos carpeta 15 040_12e_R9_Consumo_Revolvente_Distribucion_de_tarjetas_por_perdida_esperada_201106.xlsx #aqui se hizo un renombre de la columna Number_Of_Creditcards a Interest_Rate
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[14][0],"B:AO","Expected_Loss_Probability",Expected_Loss_Probability="category").rename(columns={"Number_Of_Creditcards":"Interest_Rate"})
#Limpiar datos carpeta 16 040_12h_R2_Porcentaje_de_uso_de_linea_por_perdida_esperada_201106.xlsx aqui se hizo un renombre de la columna Number_Of_Creditcards a credit_line_utilization_rate
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[15][0],"B:AO","Expected_Loss_Probability",Expected_Loss_Probability="category").rename(columns={"Number_Of_Creditcards":"credit_line_utilization_rate"})
#Limpiar datos carpeta 17 040_12h_R3_Porcentaje_de_pago_minimo_entre_saldo_a_pagar_por_perdida_esperada_201106.xlsx aqui se hizo un renombre de la columna Number_Of_Creditcards a Min_Payment_To_Balance_Ratio
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[16][0],"B:AO","Expected_Loss_Probability",Expected_Loss_Probability="category").rename(columns={"Number_Of_Creditcards":"Min_Payment_To_Balance_Ratio"})
#Limpiar datos carpeta 18 040_12h_R5_Porcentaje_de_pago_realizado_entre_pago_minimo_por_perdida_esperada_201106.xlsx aqui se hizo un renombre de la columna Number_Of_Creditcards a Min_Payment_Coverage_Percentage
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[17][0],"B:AO","Expected_Loss_Probability",Expected_Loss_Probability="category").rename(columns={"Number_Of_Creditcards":"Min_Payment_Coverage_Percentage"})
#Limpiar datos carpeta 19 040_12h_R8_Impagos_consecutivos_por_perdida_esperada_201106.xlsx 
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[18][0],"B:AO","Expected_Loss_Probability",Expected_Loss_Probability="category").rename(columns={"Number_Of_Creditcards":"Months_In_Default"})
#Limpiar datos carpeta 20 040_12h_R9_Meses_transcurridos_desde_la_apertura_por_perdida_esperada_201106.xlsx en esta se renombro la columna Number_Of_Creditcards a portfolio_maturity_months
#x=limpiar_datos_anchos(gv.ARCHIVOS_DENTRO_CARPETAS_OUTPUT[19][0],"B:AO","Expected_Loss_Probability",Expected_Loss_Probability="category").rename(columns={"Number_Of_Creditcards":"portfolio_maturity_months"})
#-------------->ok
class DataFrameLeo:
            # Tipo: Recibe (DataFrame ) -> Devuelve (DataFrame)
# Callable: Significa "algo que se puede llamar" (una función o método).

# Primeros corchetes [...]: Es una lista de los tipos de datos que la función espera recibir. Si recibiera dos cosas, sería [[pd.DataFrame, int], ...].

# La coma final ,: Separa los datos de entrada del dato de salida.

# Último elemento: Es el tipo de dato que obtendrás cuando ejecutes la función

    def __init__(self, ruta:str ,usecols:str,ESTRATEGIA:Callable[[dtf],dtf],dtf)->None:
        self.ruta:str=ruta
        self.usecols:str=usecols


        self.df:dtf=self._leer_excel()
        
    def _leer_excel(self)->dtf:
        return pd.read_excel(self.ruta,usecols=self.usecols)

