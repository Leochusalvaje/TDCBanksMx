# Instrucciones para Documentación de Python

## Estilo de Docstrings
- Usa el formato **Google Style Python Docstrings**, pero adaptado para listas con guiones.
- Está **permitido** usar guiones (`-`) al inicio de cada atributo, por ejemplo: `- nombre_variable (tipo): descripción`.
- La sintaxis clave debe ser: `nombre_variable (tipo): descripción`. Asegúrate de que los dos puntos `:` vayan pegados al paréntesis del tipo de dato, sin espacios antes de ellos, para que el editor mantenga la línea unida.
- Si la descripción se extiende a varias líneas, la sangría debe alinearse correctamente debajo del texto para no romper la viñeta.

## Tono y Extensión
- El resumen inicial de la clase o método puede ser tan detallado como sea necesario para explicar procesos complejos o flujos de negocio. No te limites a dos líneas si el contexto lo amerita.
- Sé claro, técnico y preciso, evitando explicaciones redundantes pero manteniendo la profundidad del proceso.

## Ejemplo de Referencia Obligatorio
Al generar un docstring para una clase, básate exactamente en este formato dinámico:

```python
class EjemploClase:
    """Aquí va la explicación detallada de la clase. Puede extenderse los párrafos
    que sean necesarios para describir la lógica de negocio, reglas o el comportamiento
    global de la automatización dentro del sistema.

    Attributes:
        - ruta_archivo (Path): La ruta del archivo que se va a procesar.
        - lista_fechas (list[str]): Lista de fechas en formato YYYY-MM-DD para la navegación.
        - total_pendientes (int): Cantidad de elementos restantes por procesar con una descripción 
          más larga que puede saltar de línea de forma segura.
    """