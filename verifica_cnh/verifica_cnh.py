################### Imports ###################################################

import pandas as pd
import datetime
# Pode ser necessário fazer instalações, se não tiver os pacotes basta rodar as linhas comentadas abaixo no terminal
# pip install pandas
# pip install openpyxl
# pip intall xlrd==2.0.1

################### Leitura dos dados #########################################

finalizar = 1

while finalizar == 1:
    try:
        dados = pd.read_excel('RelValidadeCNH.xls', None)
        finalizar = len(dados) + 1
    except:
        try:
            print('Verifique se o arquivo está na mesma pasta que o programa e se o nome é o original "RelValidadeCNH"')
            caminho = input('Caso esteja na mesma pasta, digite o nome do arquivo \n')
            dados = pd.read_excel(f'{caminho}.xls', None)
            finalizar = len(dados) + 1
        except:
            pass

if len(dados) > 1:
  cont = 0
  for i in dados.keys():
    if cont <= 0:
      _dados = dados[i]
    else:
      _dados = pd.concat([_dados, dados[i]], axis=0)
  dados = _dados
else:
  dados = dados['RelVctoCNH']

################### Corrigindo os dados ########################################


dados.reset_index(inplace=True)
dados = dados[['Nome Condutor', 'SIAPE Condutor', 'Unnamed: 2', 'Unnamed: 14', 'Unnamed: 12', ]].dropna()
dados.columns = ['condutor', 'siape', 'status', 'Validade CNH', 'data_fim', ]
dados.reset_index(inplace=True)
dados.drop(columns='index', inplace=True)
dados['siape'] = dados['siape'].apply(int)

################### LENDO OS SIAPES DA AGÊNCIA #################################

siapes = pd.read_csv('siapes.txt', sep=',', header = None)
siapes = siapes.loc[0].to_list()
dados = dados.query(f'siape in {siapes} & status == "PEDIDO_CONCLUIDO"')

################### Calculando o Tempo restante ################################

tempo_restante = []
for i in dados.data_fim:
    _ = [i.split('/')]
    dias = (datetime.date(int(_[0][2]), int(_[0][1]), int(_[0][0])) - datetime.date.today()).days

    if dias <= 0:
        tempo_restante.append('VENCIDO')
    else:
        tempo_restante.append(dias)

################### Tratamento para vizualização ###############################

dados['tempo_restante'] = tempo_restante
dados.drop_duplicates('siape', keep='last', inplace=True)
dados.columns = ['Funcionário', 'Siape', 'status', 'Validade CNH', 'Vencimento', 'Dias restantes']
dados.set_index('Funcionário')
dados.drop(columns='status', inplace=True)
dados.sort_values('Funcionário',inplace = True)

################### MAIN ########################################################


print(f'Tabela geral:\n{dados}\n\n\n\n')
dados.to_excel('Vencimento Carteiras.xlsx', index=False)
print('Os seguintes funcionários estão com a carteira vencida:\n\n')

print(dados[dados['Dias restantes'] == 'VENCIDO'])

finalizar = '_'
while (finalizar == '_'):
    finalizar = input('\n\nAperte enter para finalizar')
