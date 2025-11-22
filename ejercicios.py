d = {"a": 1, "b": 2}
print(d)

def sumar(a, b):
    return a + b
parametros = {"a": 10, "b": 20}
resultado = sumar(**parametros)
print(resultado)

parametros = {"a": 5, "b": 7}
print(sumar(**parametros))

#parametros = {"a": 1, "b": 2, "c": 3}
#print(sumar(**parametros))

def mostrar_kwargs(**kwargs):
    print(kwargs)
mostrar_kwargs(nombre="Leo", edad=27, carrera="Datos")

def funcion(a, b, **otros):
    print("a:", a)
    print("b:", b)
    print("otros:", otros)
funcion(10, 20, c=30, d=40)
def crear_objeto(**atributos):
    return atributos

obj = crear_objeto(nombre="Leo", nivel=1, lenguaje="Python")
print(obj)
