from openpyxl import Workbook, load_workbook

#Carrega planilha excel
wb = load_workbook("Export_SIGITM.xlsx")
ws = wb["Planilha1"]
TP_List = []

cabecalho = [celula.value for celula in ws[1]]

for row in ws.iter_rows(min_row=2, values_only=True):
    TP_List.append(list(row))
    
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
        self.Elementos = dados[9]

Objeto_TP = []
for item in TP_List:
    try:
        Objeto_TP.append(TP(item))
    except IndexError:
        print(f"Linha incompleta ignorada: {item}")

def mostrar_todas_TPs():
    print("\nRELATÓRIO DE TODAS AS TPs:")
    for TP in Objeto_TP:
        print("\n" + "-"*30)
        for attr, value in TP.__dict__.items():
            print(f"{attr}: {value}")

def buscar_tp():
    numero = input("\nDigite o número da TP: ")
    
    for tp in Objeto_TP:
        if str(tp.numTP) == str(numero):  # Agora usando numTP
            print(f"\nNúmero: {tp.numTP}")
            print(f"Descrição: {tp.desc}")  # E aqui desc
            print(f"Data: {tp.data}")
            print(f"Elementos: {tp.elementos}")
            return
    
    print(f"\nTP {numero} não encontrada!")

#mostrar_todas_TPs()        
buscar_tp()