"""
Pruebas Unitarias para registro_INFOTEC.py
==========================================
Verifica el comportamiento de FormularioUsuario ante entradas incorrectas:
mensajes de error, contador de reintentos y valor final aceptado para
cada campo (edad, nombre, correo). Usa unittest.mock para simular inputs
sin intervención del usuario.

Uso:
    python test_registro_INFOTEC.py

Requisitos:
    unittest, unittest.mock (librería estándar)
"""

import unittest
from unittest.mock import patch
from io import StringIO

from registro_INFOTEC import FormularioUsuario


class TestFormulario(unittest.TestCase): #<- unittest.TestCase hace que esta clase sea reconocida por unittest

    print("\nIniciando pruebas del programa de registro...\n")

    # Función que verifica que los mensajes de error sean los correctos o los "esperados"
    def verificar_secuencia(self, texto, esperados, entradas):
        """
        Verifica que todos los mensajes aparezcan
        y en el orden correcto.
        """
        print("Los siguientes son los valores incorrectos ingresados:")
        print(*entradas,sep="\n")    
        print("\n")    
        for i,mensaje in enumerate(esperados):

            self.assertEqual( #<- assertEqual = Confirma que sea igual, sino... 
                mensaje,
                texto[i+3], #<- el contador empieza 3 líneas después por el mensaje de bienvenida al inicio
                f"No se encontró el mensaje:\n{mensaje}"
            )

            print(f"✓ Mensaje: {mensaje} ----- OK")
            
            
    # EMPIEZAN LOS TESTS
    @patch("builtins.input") # <- @patch sustituye los inputs por mock_input, de esta manera cuando llamamos a formulario.recopilar(), en lugar de esperar la entrada de usuario, utiliza los datos del mock_input
    def test_a_edad(self, mock_input):
        print("\n")
        entradas = [ #<- Definimos los valores erroneos
            "abc",          # ValueError
            "-5",           # Fuera de rango
            "5!",           # ValueError
            "25",           # Correcto

            "Juan Perez",
            "juan@test.com"
        ]
        mock_input.side_effect = entradas  #<- .side_effect itera los valores dentro del mock_input

        salida = StringIO() #<- definimos esta variable para extraer los datos después

        with patch("sys.stdout", salida): #<- Hacemos el parche para cambiar el apuntador de la consola a salida(StringIO)

            formulario = FormularioUsuario()
            formulario.recopilar()

        texto = salida.getvalue() # <- Extrayendo los datos y guardandolos en una variable real
        texto = texto.splitlines() #<- Convertimos los print que originalmente se guardaron como una sola cadena de caracteres a una lista por línea impresa

        esperados = [ # Mensajes que se espera que esten guardados (en orden) en la variable texto
            "Entrada no válida. Verifique el tipo de dato.",
            "Intentos restantes: 3",
            "",

            "Edad fuera de rango. Ingrese un número entre 1 y 120.",
            "Intentos restantes: 2",
            "",

            "Entrada no válida. Verifique el tipo de dato.",
            "Intentos restantes: 1",
            "",
        ]

        self.verificar_secuencia( # <- Verificamos
            texto,
            esperados,
            entradas
        )

        self.assertEqual( #<- Verificamos que el valor final sea el correcto
            formulario.edad,
            25
        )

        print("\n✓ test_edad OK")

    @patch("builtins.input")
    def test_b_nombre(self, mock_input):
        print("\n")
        entradas = [
            "25",

            "12345",
            "Juan123",
            "Juan!",

            "Juan Perez",

            "juan@test.com"
        ]
        mock_input.side_effect = entradas

        salida = StringIO()

        with patch("sys.stdout", salida):

            formulario = FormularioUsuario()
            formulario.recopilar()

        texto = salida.getvalue()
        texto = texto.splitlines()

        esperados = [
            "Nombre no válido. Solo letras y espacios, máximo 70 caracteres.",
            "Intentos restantes: 3",
            "",

            "Nombre no válido. Solo letras y espacios, máximo 70 caracteres.",
            "Intentos restantes: 2",
            "",

            "Nombre no válido. Solo letras y espacios, máximo 70 caracteres.",
            "Intentos restantes: 1",
            ""
        ]

        self.verificar_secuencia(
            texto,
            esperados,
            entradas
        )

        self.assertEqual( #<- para qué ocupas este si ya lanzas la validación anterior?
            formulario.nombre,
            "Juan Perez"
        )

        print("\n✓ test_nombre OK")

    @patch("builtins.input")
    def test_c_correo(self, mock_input):
        print("\n")
        entradas = [
            "25",
            "Juan Perez",

            "correo",
            "correo@",
            "a@b.c",

            "juan@test.com"
        ]
        mock_input.side_effect = entradas
        
        salida = StringIO()

        with patch("sys.stdout", salida):

            formulario = FormularioUsuario()
            formulario.recopilar()

        texto = salida.getvalue()
        texto = texto.splitlines()
       
        esperados = [
            "Correo no válido. Debe contener '@' y '.'.",
            "Intentos restantes: 3",
            "",

            "Correo no válido. Debe contener '@' y '.'.",
            "Intentos restantes: 2",
            "",

            "Correo no válido. Debe contener '@' y '.'.",
            "Intentos restantes: 1"
            "",
        ]

        self.verificar_secuencia(
            texto,
            esperados,
            entradas
        )

        self.assertEqual(
            formulario.correo,
            "juan@test.com"
        )

        print("\n✓ test_correo OK")


if __name__ == "__main__": #<- línea para poder llamar al ejecutable unittest desde test_registro.py
    unittest.main(verbosity=2)
