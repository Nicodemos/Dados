import pandas as pd
import openpyxl
import os
import xlrd

# Uma observação, o backlog é por dia e não por estado. Não associar o backlog ao estado. Associar ao dia, mês e ano
def qtd_linhas_estados(df):
     cont=0
     for x in df.iloc[:,:1].values:
          if 'BIOMOL_AMOSTRAS' in x[0]:
               cont+=1
     return cont

def trata_coluna_estado(col):
     aux = str(col).split('_')[2:]
     if len(aux)>1:
          return aux[0]+' '+aux[-1]
     return aux[0]

colunas = ['Data','Estado','Processadas','Recebidas','Positivas']
data_frame_temp = pd.DataFrame(columns=colunas)
data_frame_global = pd.DataFrame(columns=colunas)
cont_row = 0
colunas_temp = ['MNEMÔNICO INDICADOR','RESULTADO', 'COMPETÊNCIA']

lista_reteste = []
lista_datas = []
lista_backlog = []
conta_erro = 0
backlog = 0

# Retorna, o diretório como str, nome de subdiretórios em dirs, e nome de arquivos do diretório alvo em files.
path, dirs, files = next(os.walk(r"\\fioce-d-ca11\Compartilhamento\Relatório Biomol"))
for pastas_mes in dirs:
     diretorio, lista_pastas, lista_arq = next(os.walk(r"\\fioce-d-ca11\Compartilhamento\Relatório Biomol\{}".format(pastas_mes)))

     for pastas_final in lista_pastas:
          dir_final, pastas_finais, arquivos = next(os.walk(r"{}\{}".format(diretorio,pastas_final)))

          for i in arquivos:
               if 'INTERACT' in i:
                    j = 0
                    cont_linhas = 0
                    try:
                         data_frame_temp = pd.DataFrame(columns=colunas)
                         pl = pd.read_excel(r"{}\{}".format(dir_final,i),sheet_name='Resultados')
                         df_temp = pl[colunas_temp]  ## indexa apenas os dados escolhidos no arquivo
                         colunas_file = ['MNEMONICO','RESULTADO', 'COMPETENCIA'] # retirando os acentos

                         aux = df_temp.values
                         df_temp = pd.DataFrame(aux, columns=colunas_file)
                         tamanho = qtd_linhas_estados(df_temp)
                         data_frame_temp['Estado'] = df_temp.iloc[:tamanho, :1]

                         data_frame_temp['Data'] = df_temp['COMPETENCIA'].iloc[:tamanho]
                         data_frame_temp['Data'] = data_frame_temp['Data'].map(lambda x: str(x).replace('-','/'))

                         data_frame_temp['Processadas'] = df_temp['RESULTADO'].iloc[:tamanho]

                         tamanho_pos = tamanho*2
                         data_frame_temp['Positivas'] = df_temp['RESULTADO'].iloc[tamanho:tamanho_pos].values
                         tamanho_rec = tamanho_pos*2
                         recebidas = df_temp['RESULTADO'].iloc[tamanho_pos:tamanho_rec]

                         if len(recebidas.values)<2: # Nem não é em todas as planilhas que tem a informação de amostras recebidas,
                              pass
                         else:
                              # faz o cálculo do backlog das amostras
                              data_frame_temp['Recebidas'] = pd.Series(recebidas.values)
                              if data_frame_temp['Processadas'].astype(int).sum() <= (data_frame_temp['Recebidas'].astype(int).sum() + backlog):
                                   backlog = data_frame_temp['Processadas'].astype(int).sum() - (data_frame_temp['Recebidas'].astype(int).sum() + backlog)  # cria o campo backlog
                                   if backlog < 0:
                                     backlog *= -1
                              else:
                                   backlog = 0
                              lista_backlog.append(backlog)
                              j=1  ## variavel controladora, funciona como interruptor
                         if j==0:
                              lista_backlog.append(0)

                         data_frame_global = pd.concat([data_frame_global,data_frame_temp],ignore_index=True)
                         lista_datas.append(str(df_temp['COMPETENCIA'].iloc[1:2].values[0]).replace('-','/'))
                         lista_reteste.append(df_temp['RESULTADO'].iloc[:-1].values[0])
                    except Exception as inst:
                         print(inst)
                         conta_erro+=1

data_frame_global = data_frame_global.fillna(0)

# criando outros dataframe auxiliares
df_reteste = pd.DataFrame({"Data":lista_datas,"Reteste":lista_reteste})
df_backlog = pd.DataFrame({"Data":lista_datas,"Backlog":lista_backlog})
data_frame_global['Estado'] = data_frame_global['Estado'].map(trata_coluna_estado)
data_frame_global['Positividade'] = (data_frame_global['Positivas'] / data_frame_global['Processadas'])*100
data_frame_global['Positividade'] = data_frame_global['Positividade'].map(lambda x: round(x, 2))
data_frame_global = data_frame_global.fillna(0)

# gerando planilhas com dados agregados
data_frame_global.to_excel(r"\Users\nicodemos.freitas\dados_pcr\acumulado_pcr.xlsx", sheet_name='historico',index=False)
df_backlog.to_excel(r"\Users\nicodemos.freitas\dados_pcr\backlog_pcr.xlsx", sheet_name='backlog',index=False)
df_reteste.to_excel(r"\Users\nicodemos.freitas\dados_pcr\retestes.xlsx", sheet_name='reteste',index=False)

print('# ------ FIM ------- #')
print('# TOTAL DE ERROS NO PRCESSAMENTO DAS PLANILHAS = ',conta_erro)
