import pandas as pd
import win32com.client as win32

# importar a base de dados
tabela_vendas = pd.read_excel('vendas.xlsx')
# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-' * 50)
# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)
# quantidade por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)
# ticket medio por produto  em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'paulo_nunis2@hotmail.com'
mail.Subject = 'Relatório de vendas por Loja'
mail.HTMLBody = f'''
<p>Prezado,</p>

<p>segue relatório de vendas por cada loja</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticker Médio do produto em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Méido': 'R${:,.2f}'.format})}

<p>Qual quer duvida estou a disposição</p>
'''

mail.Send()
