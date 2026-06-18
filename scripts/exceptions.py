class ErrorAlCerrarExcel(Exception):
    """
    Ocurre cuando Excel no puede cerrarse correctamente.
    """
    pass


class ErrorConfiguracion(Exception):
    """
    Error relacionado con el archivo de configuración.
    """
    pass


class ErrorAutomata(Exception):
    """
    Error general del autómata.
    """
    pass

class ErrorAlTraerLibroAlFrente(Exception):
    """
    Ocurre cuando no se puede traer el libro de Excel al frente.
    """
    pass
class ErrorAlCerrarLibro(Exception):
    """
    Ocurre cuando no se puede cerrar un libro de Excel correctamente.
    """
    pass

class ErrorAlDesactivarAlertasExcel(Exception):
    """
    Ocurre cuando no se pueden desactivar las alertas de Excel.
    """
    pass

class ErrorAlCrearBotones:
    """Ocurre cuando no se pudieron generar de manera correctalos botones al momento de leer la lista de pendientes
    """
    pass