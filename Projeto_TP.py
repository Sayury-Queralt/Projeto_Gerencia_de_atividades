from openpyxl import load_workbook
from collections import defaultdict

# Carregar o arquivo Excel
wb = load_workbook('Export_SIGITM_real.xlsx')
ws = wb.active

# Dicionário para agrupar linhas pelo ID: {id: [linha_consolidada]}
dados_agrupados = defaultdict(lambda: {'dados': None, 'elementos': []})

# Processar linhas do Excel
# Iterar sobre as linhas (ignorando o cabeçalho se houver)
for row in ws.iter_rows(min_row=2, values_only=True):
    try:
        id_referencia = row[1]  # Coluna B (ID)
        elemento = row[-1]       # Última coluna (Elemento)
        
        # Se é a primeira ocorrência deste ID, armazena os dados
        if dados_agrupados[id_referencia]['dados'] is None:
            dados_agrupados[id_referencia]['dados'] = row[:-1]  # Todos os campos exceto o último
            
        # Adiciona o elemento à lista (se for válido)
        if elemento and str(elemento).strip():
            dados_agrupados[id_referencia]['elementos'].append(str(elemento).strip())
    except Exception as e:
        print(f"Erro ao processar linha: {row}. Erro: {e}")

# Converter para lista final (formato desejado)
#TP_List = list(dados_agrupados.values())
    
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
        self.Elementos = elemento

# Criar lista de objetos TP
Objeto_TP = []
for id_tp, grupo in dados_agrupados.items():
    if grupo['dados'] and grupo['elementos']:  # Só cria se tiver dados e elementos
        try:
            Objeto_TP.append(TP(grupo['dados'], grupo['elementos']))
        except Exception as e:
            print(f"Erro ao criar TP {id_tp}: {str(e)}")

#Função para listar todas as TPs
def mostrar_todas_TPs():
    print("\nRELATÓRIO DE TODAS AS TPs:")
    
    for TP in Objeto_TP:
        print("\n" + "-"*30)
        for attr, value in TP.__dict__.items():
            print(f"{attr}: {value}")

#Função de Busca
def buscar_tp():
    while True:
        numero = input("\nDigite o número da TP (ou 'sair' para encerrar): ")
        
        if numero.lower() == 'sair':
            print("Encerrando...")
            break
            
        encontrada = False
        for tp in Objeto_TP:
            if str(tp.Origem) == numero:
                print("\n" + "="*50)
                print(f"TP Nº: {tp.Origem}")
                print(f"Data/Hora: {tp.Data_Criacao}")
                print(f"Descrição: {tp.Descricao}")
                print(f"Data Prevista: {tp.Data}")
                print(f"Tipo: {tp.Tipo}")
                print(f"Status: {tp.Status}")
                print(f"Executor: {tp.Executor}")
                print(f"Afetação: {tp.Afetacao}")
                print("Elementos Relacionados:")
                for i, elemento in enumerate(tp.Elementos, 1):
                    print(f"  {i}. {elemento}")
                print("="*50)
                encontrada = True
                break
                
        if not encontrada:
            print(f"\nATENÇÃO: TP {numero} não encontrada!\n")


# Verificação rápida dos dados processados
print("\n=== VERIFICAÇÃO INICIAL ===")
print(f"Total de TPs processadas: {len(Objeto_TP)}")
if Objeto_TP:
    print(f"\nExemplo da primeira TP:")
    print(f"ID: {Objeto_TP[0].Origem}")
    print(f"Elementos: {Objeto_TP[0].Elementos}")
    print(f"Quantidade de elementos: {len(Objeto_TP[0].Elementos)}")

#mostrar_todas_TPs()        
buscar_tp()
#print(TP_List)