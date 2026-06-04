from typing import Counter, OrderedDict
from collections import OrderedDict


compras = [
    "Luis", "Ana", "Luis", "Carlos", "Marta", "Ana", 
    "Sofía", "Elena", "Luis", "Carlos"
]

registros = ["Ana", "Carlos", "Marta", "Elena"]

# Obteniendo valores únicos de compras
compras_unicas = list(OrderedDict.fromkeys(compras))
print(f"\nCompradores únicos: {compras_unicas} \n")

# Usuarios no registrados
diferencia = set(compras_unicas).difference(set(registros))
print(f"Usuarios no registrados: {diferencia} \n")

# Contando la cantidad de compras por persona
Conteo = Counter(compras)
print(f"Conteo de compras por cliente: {Conteo} \n")

# Filtrar personas con más de 3 compras
compradores_filtrados = [persona for persona, cantidad in Conteo.items() if cantidad > 1]
print(f"Compradores con más de 1 compras: {compradores_filtrados} \n")
