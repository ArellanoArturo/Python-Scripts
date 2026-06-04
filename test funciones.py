import time
from functools import reduce

# Generador de temperaturas de ciudades mexicanas
def leer_temperaturas():
    yield( "CDMX", 26 )
    yield( "Monterrey", 35 )
    yield( "Guadalajara", 28 )
    yield( "Cancún", 32 )
    yield( "Mérida", 33 )
    yield( "Puebla", 29 )

# Filtrar temperaturas mayores a 30 grados
Temperatura = list(filter(lambda x: x[1] > 30, leer_temperaturas()))

# Ordenar temperaturas de menor a mayor
Temperatura = sorted(Temperatura, key=lambda x: x[1])

# Generar reporte de alertas de calor, Tuplas a string
reporte = map(lambda x: f"Alerta de calor en {x[0]}: {x[1]}°C", Temperatura)
print(list(reporte))

# Calcular temperatura promedio de las ciudades con alerta de calor > 30°C

promedio = reduce(lambda x, y: x + y[1], Temperatura, 0) / len(Temperatura)
print(f"Temperatura promedio de alertas: {promedio:.2f}°C")

# Decorador para auditar el tiempo de ejecución y la cantidad de llamadas de una función
def auditar_funcion(funcion):
    n=0
    def wrapper():
        nonlocal n
        inicio = time.time()
        funcion()
        fin = time.time()
        n = n + 1
        print(f"Tiempo de ejecución: {fin - inicio:.2f} segundos")
        print(f"Nombre de la función: {funcion.__name__}")
        print(f"Cantidad de veces que se ha ejecutado: {n}")
    return wrapper



    