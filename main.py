import pyodbc
import pandas as pd
import os
from plyer import  notification

direccion_servidor = 'BOLSAGESTION\CUCCHIARA'
nombre_bd = 'testBolsa'
nombre_usuario = 'sa'
password = 'Bolsa411'

#CREAMOS LA CONEXION A LA BASE DE DATOS
conexion = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' +
                              direccion_servidor+';DATABASE='+nombre_bd+';UID='+nombre_usuario+';PWD=' + password)
cursor = conexion.cursor()
print("Se ha conectado con exito")

# GENERAMOS LA CONSULTA EN LA BASE DE DATOS Y LO TRAEMOS A PYTHON
sqlConsulta = "SELECT * FROM ClientesSinCondomino WHERE Productor = '8'"

df = pd.read_sql_query(sql=sqlConsulta, con=conexion, index_col=None)
print(df)

try:
    #EXPORTAMOS LA CONSULTA AL ESCRITORIO
    df.to_excel(os.environ["userprofile"] + "\\Desktop\\Kikiproyect\\" + "ClientesKiki" + ".xlsx" , index = False)

    #NOTIFICAMOS QUE TERMINO
    notification.notify(title = "Estado de reporte", message = f"Se ha guardado correctamente se exportaron un: \nTotal filas: {df.shape[0]}\nTotal Columnas: {df.shape[1]}", timeout = 10)
except ValueError as error:
    notification.notify(title="Estado de reporte",
                        message=f"No se pudo exportar el documento \nAsegurese de tener el archivo cerrado\nERROR: {error}",timeout=10)
