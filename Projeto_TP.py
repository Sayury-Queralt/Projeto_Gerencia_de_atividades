from openpyxl import load_workbook
from collections import defaultdict
from openpyxl import Workbook
import re

Afetacao_list = [["0","Não Afeta Elemento nem Serviços"],
["1.1","Sem Afetação Elemento e Afetação Parcial Serviços"],
["1.2","Afetação Parcial Elemento e sem Afetação Serviços"],
["3","Afetação Parcial Elemento e Serviços"],
["4.1","Afetação Total Elemento e sem Afetação Serviços"],
["4.2","Sem Afetação Elemento e Afetação Total Serviços"],
["5.1","Afetação Parcial Elemento e Total Serviços"],
["5.2","Afetação Total Elemento e Parcial Serviços"],
["6","Afetação Total Elemento e Serviços"]]

# Carregar o arquivo Excel
wb = load_workbook('TP_Criação_Host_Sayury.xlsx')
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
        self.ElementosFull = dados [-2]
        self.ElementosAgr = dados[-1]

#Manipulação do formato dos elementos
def ElementosFinal(ElementoPlanilha):
    Elemento_Final = {}
    ElementoInput = ElementoPlanilha
    ElementosSplit = [[p[0],p[1],p[2][:3]] for p in (re.split(r"[_-]",elem) for elem in ElementoInput)]
    SortNF = sorted(ElementosSplit, key=lambda x: x[-1])

    for sub in SortNF:
        K = sub[2]
        V = sub[0]
        if K not in Elemento_Final:
            Elemento_Final[K] = []
        Elemento_Final[K].append(V)
    return Elemento_Final

#Manipulação e preparação 
def Criar_TP(dados):
    Tipo_TP = dados[4]
    if Tipo_TP == "S":
        Tipo_TP = "PA"
    else:
        Tipo_TP = "PR"
    dados[4]=Tipo_TP

    Afetacao = dados[8]
    for A in Afetacao_list:
        if A[1] == Afetacao:
            dados[8] = A[0]
     
    Elemento_entrada_def = dados[-1]
    ElementoFinal = ElementosFinal(Elemento_entrada_def)    
    dados.append(ElementoFinal)
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
                print(f"Elementos: {tp.ElementosAgr}")
                
                encontrada = True
                break
                
        if not encontrada:
            print(f"\nATENÇÃO: TP {numero} não encontrada!\n")

#buscar_tp()

wb = Workbook()
ws = wb.active
ws.title = "Teste"
#cabecalho = list(vars(Classe_TP[0]).keys())
cabecalho = ["TP","Descrição","Data validação","Validador","Executor","Tipo","Afetação","Status","Férias","Calendário","Aberto por","POP","Elemento"]
ws.append(cabecalho)
for tp in Classe_TP:
    El = tp.ElementosAgr
    K = El.keys()
    V = El.values()
    K_String = str(K)
    V_String = str(V)
    ws.append([tp.Origem,tp.Descricao,tp.Data,"",tp.Executor,tp.Tipo,tp.Afetacao,"Em análise","","Pendente","",K_String,V_String])
wb.save("Saida.xlsx")