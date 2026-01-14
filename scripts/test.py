import pytest
import automatizar_v1_0 as aut
import pathlib as path

def test_obtener_serie():
    nombre="040_12e_R10_Distribución de tarjetas por probabilidad de incumplimiento"
    resultado=aut.busqueda_patron(nombre)
    assert resultado=="040_12e_R10"

def test_obtener_serie_corto():
    nombre="040_12e_R6_ConsumoRevolvente_Tarjetas_Pagominimo_vs_pagoRealizado"
    resultado=aut.busqueda_patron(nombre)
    assert resultado=="040_12e_R6"

def test_obtener_serie_2():
    nombre="040_12e_R50_Distribución de tarjetas por porcentaje de pago mínimo respecto a la línea de crédito"
    resultado=aut.busqueda_patron(nombre)
    assert isinstance(resultado,str)
    assert resultado=="040_12e_R50"
    assert len(resultado)>0
def test_revisar_que_se_cree_carpeta(tmp_path):
    ruta_base=tmp_path/"Output"