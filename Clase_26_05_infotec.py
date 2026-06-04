
lista_de_precios = [23,35,41,52,64,5]

nuevos_precios = []

for e in lista_de_precios:
    if e > 30:
        nuevos_precios.append(round(e*1.16, 2))

print(nuevos_precios)

#También puede ser
nuevos_precios = [f"{e*1.16:.2f}" for e in lista_de_precios if e > 30]
print(nuevos_precios)   