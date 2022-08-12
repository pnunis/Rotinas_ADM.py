import pandas as pd
import os

# Obter o diretório do arquivo de entrada
diretorio = os.path.dirname('D:/Users/paulo.souza/Desktop/TESTE PYEXCEL/TEXTO_JUI.xlsx')

# Leitura do arquivo Excel
#df = pd.read_excel('TEXTO_TESTE.xlsx')
df = pd.read_excel('D:/Users/paulo.souza/Desktop/TESTE PYEXCEL/TEXTO_JUI.xlsx')
# Agrupamento por número de pedido e junção do texto das linhas
df_agrupado = df.groupby('Documento de compras').apply(lambda x: '; '.join(x['Linha de texto'])).reset_index()

# Concatenar o nome do arquivo de saída com o diretório do arquivo de entrada
caminho_saida = os.path.join(diretorio, 'resultado.xlsx')

# Criação de um novo arquivo Excel com o resultado
df_agrupado.to_excel(caminho_saida, index=False)
