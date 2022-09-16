import pandas as pd
import openpyxl
import win32com.client as win32

#Importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#Visualizar base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)
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
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(round(ticket_medio,3))

#Enviar email com relatorio


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'erikacristine0702@gmail.com'
mail.Subject = 'Relatório de vendas por loja.'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade_produto.to_html()}

<p>Ticket Médio dos Produtos em cada Loja</p>
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Paulo</p>
'''

mail.Send()

print("EMAIL ENVIADO.")