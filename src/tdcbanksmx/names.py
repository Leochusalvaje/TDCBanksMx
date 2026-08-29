import re
import tdcbanksmx.routes as gv
from pathlib import Path
from typing import List, Tuple

def renombrar_y_limpiar_archivos(ruta_carpeta: Path) -> List[Tuple[Path, Path]]:
    """
    Recorre los archivos en la carpeta, elimina el patrón Distribucion_de_tarjetas_por_,
    e inserta "_Porcentaje_" capitalizando la P.
    """
    if not ruta_carpeta.is_dir():
        print(f"Error: La ruta {ruta_carpeta} no es un directorio válido.")
        return []
    
    cambios_realizados = []
    
    # Patrón de sustitución fijo:
    # Captura la cadena que quieres eliminar (incluyendo el guion bajo y el 'por').
    PATRON_A_ELIMINAR = r'_Consumo_Revolvente_desde_la_apertura_de_la_cuenta_' 
    
    # Cadena de reemplazo: Guion bajo seguido de "Porcentaje" (la P en mayúscula)
    CADENA_REEMPLAZO = r'_Consumo_Revolvente_Distribución_por_meses_desde_apertura_' 

    for ruta_original in ruta_carpeta.iterdir():
        
        if ruta_original.is_file():
            
            nombre_original = ruta_original.name
            
            # 1. Aplicar la sustitución Regex
            # Busca el patrón y lo reemplaza por la cadena de reemplazo.
            nombre_nuevo = re.sub(PATRON_A_ELIMINAR, CADENA_REEMPLAZO, nombre_original)
            
            # 2. Verificar si hubo un cambio para evitar renombrar innecesariamente
            if nombre_original != nombre_nuevo:
                
                # 3. Construir la ruta nueva y completa
                # ruta_original.parent es la carpeta (WindowsPath)
                ruta_nueva = ruta_original.parent / nombre_nuevo
                
                # 4. Renombrar el archivo en el disco
                ruta_original.rename(ruta_nueva)
                cambios_realizados.append((ruta_original, ruta_nueva))
                
    return cambios_realizados

# --- Ejemplo de Uso ---

# NOTA: Asegúrate de que esta ruta apunte a la carpeta que contiene los archivos
# RUTA_A_PROCESAR = gv.RUTAS_OUTPUT_CONSULTAS[12]
# print(RUTA_A_PROCESAR)
# resultados = renombrar_y_limpiar_archivos(RUTA_A_PROCESAR)

# print(f"Archivos renombrados: {len(resultados)}")
# if resultados:
#     print("\nPrimeros 3 cambios:")
#     for original, nuevo in resultados[:3]:
#         print(f"  De: {original.name}\n  A:  {nuevo.name}")