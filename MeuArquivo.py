import pandas as pd
import openpyxl
import win32com.client as win32

#Importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#Visualizar base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

#print(tabela_vendas[['ID Loja', 'Valor Final']])

print("-"*50)
#Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print("-"*50)
#Quantidade de produtos vendidos por loja
quantidade_produto = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade_produto)

print("-"*50)
#Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final']/quantidade_produto['Quantidade']).to_frame()
print(round(ticket_medio,3))

#Enviar email com relatorio


outlook = win32.Dispath('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'paulooorrs@gmail.com'
mail.Subject = 'Relatório de vendas por loja.'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento}

<p>Quantidade Vendida:</p>
{quantidade_produto}

<p>Ticket Médio dos Produtos em cada Loja</p>
{ticket_medio}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Paulo</p>
'''

mail.Send()

print("EMAIL ENVIADO.")