from openpyxl import load_workbook, Workbook
import warnings

nombre_modelo = str(input("Ingrese el nombre del modelo: "))
nombre1_pago = str(input("Ingrese el nombre de la planilla de pagos: "))

nombre_modelo = f'{nombre_modelo}.xlsx'
nombre_pago = f'{nombre1_pago}.xlsx'

modelo = load_workbook(nombre_modelo, data_only=True)
pago = load_workbook(nombre_pago, data_only=True)

ws_modelo = modelo['Base']
ws_pago = pago['Sheet0']

dnis = []
montos = []
dni_montos = {}
for column_data in ws_modelo['A']:
    if column_data.value != None and column_data.value != 'DNI':
        dnis.append(str(column_data.value))

with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    for column_data in ws_modelo['AI']:
        if column_data.value != None and column_data.value != 'BANCO':
            montos.append(str(column_data.value))

contador = 0
print(len(dnis))
print(len(montos))
for dni in dnis:
    dni_montos[dni] = montos[contador]
    contador += 1 

for row in ws_pago.iter_rows():
    
    if row[2].value in dnis:
        row[5].value = dni_montos[row[2].value]
    else:
        print('Formato incorrecto')

pago.save(f"{nombre1_pago}-INTEGRADO.xlsx")