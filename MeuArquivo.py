import pandas as pd
import openpyxl

#Importar base de dados
tabela_vendas = pd.read_excel('vendas.xlsx')
pd.set_option('display.max_columns', None)
#pd.set_option('display.max_rows', None)

#Visualizar base de dados

#print(tabela_vendas[['ID Loja', 'Valor Final']])

#Faturamento por loja
Faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(Faturamento)

#Quantidade de produtos vendidos por loja

#Ticket médio por produto em cada loja

#Enviar email com relatorio