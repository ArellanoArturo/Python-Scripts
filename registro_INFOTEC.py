"""
User Registration Form - Infotec
=================================
Terminal form that collects and validates name, age, and email.
Allows up to 4 attempts per field before exiting. Built as a class
to support unit testing without modifying core logic.

Usage:
    python registro_INFOTEC.py

Requirements:
    sys (standard library)
"""

import sys

class FormularioUsuario:
    # El usuario tiene 4 intentos para ingresar datos válidos
    MAX_INTENTOS = 4
    
    # Definimos el constructor de la clase, inicializando el contador de intentos en 0
    def __init__(self):
        self.intentos = 0
        self.entrada = input #<- lo dejamos por si requerimos hacer otros tipos de tests
    
    # Método para validar datos al usuario, con validación y manejo de errores
    def validar(self, mensaje, condición, error, tipo=str):
        try:
            valor = tipo(self.entrada(mensaje).strip())
            if condición(valor): # <- Validación de datos de acuerdo a la condición de cada uno, regresa True o False
                return valor
            print(error) # <- Si la condición no se cumple, se muestra el mensaje de error específico
        except ValueError: # <- Manejo de expeción de valor
            print("Entrada no válida. Verifique el tipo de dato.")
        except KeyboardInterrupt: # <- Manejo de interrupción del programa por parte del usuario
            print("\nPrograma interrumpido.\n")
            sys.exit(0)
        return None
    
    # Función sigue preguntando al usuario hasta que ingrese un dato válido o se agoten los intentos
    def contador(self, mensaje, condición, error, tipo=str):
        while self.intentos < self.MAX_INTENTOS:
            valor = self.validar(mensaje, condición, error, tipo)
            if valor is not None: 
                self.intentos = 0 #<- Si el valor es válido, se reinicia el contador de intentos para la siguiente pregunta
                return valor
            self.intentos += 1  #<- Si la condición no se cumple, valor es None, por lo que se sigue preguntando al usuario
            restantes = self.MAX_INTENTOS - self.intentos
            if restantes > 0:
                print(f"Intentos restantes: {restantes}\n")
        print("Demasiados intentos fallidos. Programa terminado, vuelva a intentarlo.")
        sys.exit(0)

    # Función que recopila los datos del usuario, la condición y el mensaje de error para cada uno 
    def recopilar(self):          
        print("\nBienvenido al formulario de registro. Por favor, ingrese sus datos correctamente.\n")         
        self.edad = self.contador( # <- Contador empieza y guarda el número de intentos
            "Ingrese su edad: ",
            condición=lambda x: 0 < x <= 120,
            error="Edad fuera de rango. Ingrese un número entre 1 y 120.",
            tipo=int # <- Esto ya maneja que el dato sea un número entero, si no lo es, se lanzará una excepción ValueError que se captura en el método validar
            )
            
        self.nombre = self.contador(
            "Ingrese su nombre: ",
            condición=lambda x: x.replace(" ", "").isalpha() and 2 <= len(x) <= 70, 
            # Usamos replace() ya que es un requerimiento para isalpha() que no haya espacios, función que checa que el nombre sólo contenga letras
            error="Nombre no válido. Solo letras y espacios, máximo 70 caracteres."
            )
            
        self.correo = self.contador(
            "Ingrese su correo: ",
            condición=lambda x: "@" in x and "." in x and 7 < len(x) <= 150,
            # Checamos que el email tenga los requerimientos mínimos 
            error="Correo no válido. Debe contener '@' y '.'."
            )
           
    
    def mostrar(self):
        print(f"\nDatos registrados:")
        print(f"  Nombre: {self.nombre}")
        print(f"  Edad:   {self.edad} años")
        print(f"  Correo: {self.correo}")

           
# Llamada 
if __name__ == "__main__": # <- Línea para evitar que el programa se ejecute al importar la clase en los tests
    formulario = FormularioUsuario()
    formulario.recopilar()
    formulario.mostrar()
