from openpyxl import load_workbook
from collections import defaultdict

# Carregar o arquivo Excel
wb = load_workbook('Export_atividades.xlsx')
ws = wb.active

#Deleta a ultima coluna (VTP PK)
ws.delete_cols(ws.max_column)

Classe_TP = []

#Criar a classe TP
class TP:
    def __init__(self, dados):
        self.Data_Criacao = dados[0]
        self.Origem = dados[1]
        self.Descricao = dados[2]
        self.Data = dados[3]
        self.Tipo = dados[4]
        self.Status = dados[5]
        self.Executor = dados[7]
        self.Afetacao = dados[8]
        self.Elementos = dados[-1]

#Manipulação do formato dos elementos
def ElementosFinal(ElementoPlanilha):
    SortElement = ElementoPlanilha
    SortElement.sort()
    ElementosSplit = [elem.split("_") for elem in SortElement]
    Elemento = [subelem[0]+ "-" +subelem[2] for subelem in ElementosSplit]
    return(Elemento)

#Manipulação e preparação 
def Criar_TP(dados):
    Tipo_TP = dados[4]
    if Tipo_TP == "S":
        Tipo_TP = "Pré-aprovada"
    else:
        Tipo_TP = "Programada"
    dados[4]=Tipo_TP
    ElementoFinal = ElementosFinal(dados[-1])    
    dados[-1] = ElementoFinal
    Classe_TP.append(TP(dados))

#Adiciona a primeira linha em dados
dados = [cell.value for cell in ws[2]]
element = [dados[-1]]
dados[-1] = element

# Processar as demais linhas do Excel
# Iterar sobre as linhas (ignorando o cabeçalho e a primeira linha)
for row in ws.iter_rows(min_row=3, values_only=True):
    if dados[1] != row[1]:
        Criar_TP(dados)
        dados = list(row[:-2])
        element = []
        element.append(row[-1])
        element.sort
        dados.append(element)
        
    else:
        element.append(row[-1])
        dados[-1] = element            
Criar_TP(dados)


#Função de Busca
def buscar_tp():
    while True:
        numero = input("\nDigite o número da TP (ou 'sair' para encerrar): ")
        
        if numero.lower() == 'sair':
            print("Encerrando...")
            break
            
        encontrada = False
        for tp in Classe_TP:
            if str(tp.Origem) == numero:
                print("\n" + "="*50)
                print(f"TP Nº: {tp.Origem}")
                print(f"Data Criação: {tp.Data_Criacao}")
                print(f"Descrição: {tp.Descricao}")
                print(f"Data Prevista: {tp.Data}")
                print(f"Tipo: {tp.Tipo}")
                print(f"Status: {tp.Status}")
                print(f"Executor: {tp.Executor}")
                print(f"Afetação: {tp.Afetacao}")
                print(f"Elementos: {tp.Elementos}")
                
                encontrada = True
                break
                
        if not encontrada:
            print(f"\nATENÇÃO: TP {numero} não encontrada!\n")

buscar_tp()