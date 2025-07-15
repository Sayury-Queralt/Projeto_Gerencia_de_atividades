from openpyxl import load_workbook
from collections import defaultdict

# Carregar o arquivo Excel
wb = load_workbook('Export_SIGITM_real.xlsx')
ws = wb.active

#Deleta a ultima coluna (VTP PK)
ws.delete_cols(ws.max_column)

dados = [cell.value for cell in ws[2]]
element = []
List_TP = []

# Processar linhas do Excel
# Iterar sobre as linhas (ignorando o cabeçalho se houver)
for row in ws.iter_rows(min_row=3, max_row=8, values_only=True):
    if dados[1] != row[1]:
        print(dados)
        dados[-1] = row[:-2]
        element = []
        element.append(row[-1])
        print(element)
        dados.append(element)
    else:
        element.append(row[-1])
        dados[-1] = element
        print(element)
    print(element)        

