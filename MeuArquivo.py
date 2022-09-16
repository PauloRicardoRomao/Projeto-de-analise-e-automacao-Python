import pandas as pd
import openpyxl

#Importar base de dados
tabela_vendas = pd.read_excel('vendas.xlsx')

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