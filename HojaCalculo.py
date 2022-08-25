import pandas
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import xlsxwriter
from datetime import datetime

ID = ""
Nombre_Documento = ""
with open('LeerDocumento.json') as file:
    data = json.load(file)
    for Datos_Exportacion in data['Archivo_Operaciones']:
        ID = Datos_Exportacion['ID']
        Nombre_Documento = Datos_Exportacion['Nombre_Documento']

Fecha = datetime.now()

url = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
       "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
Credenciales = ServiceAccountCredentials.from_json_keyfile_name("Credenciales.json", url)

Autorizar = gspread.authorize(Credenciales)

AbrirExcel = Autorizar.open(Nombre_Documento)

worksheet = AbrirExcel.worksheet("Hoja 1")

list_of_dicts = worksheet.get_all_values()

print(list_of_dicts)

workbook = xlsxwriter.Workbook('Operaciones.xlsx')
worksheet = workbook.add_worksheet()

data_frame = pandas.DataFrame(list_of_dicts)
print(list(data_frame))
for row, _dict in enumerate(list_of_dicts):
    for col, key in enumerate(list(data_frame)):
        worksheet.write(row, col, _dict[key])
workbook.close()


