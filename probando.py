# Importamos el ModuMódulo

import pywhatkit

# Usamos Un try-except
try:

    # Enviamos el mensaje

    pywhatkit.sendwhatmsg("+5491150189948",
                          "Mensaje De Prueba",
                          15, 27)

    print("Mensaje Enviado")



except:

    print("Ocurrio Un Error")