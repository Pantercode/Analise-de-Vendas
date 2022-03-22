import pandas as pd
import  win32com.client as win32

# importar a base de dados
tabela_de_vendas = pd.read_excel('Vendas.xlsx')

# vizualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_de_vendas)

# Faturamento por loja
faturamento = tabela_de_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produto vendido por loja
quantidade = tabela_de_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('_' * 50)

# ticket médio por produto em cada loja(faturamento/ pela quantidade de produtos vendido por loja)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns= {0 : "Ticket Médio"})
print(ticket_medio)


# Enviar um email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'marcelloliveirafull@gmail.com'
mail.Subject = 'Relátorio de Vendas por Loja Mensal'
mail.HtmlBody = f'''
<p>Prezados,</p>

<p>Segue o relátorio de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,..2f}'.format})}


 
<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,..2f}'.format})}


<p>Qualquer dúvida estou a disposição.</p>

<p>Att..</p>
<p>Marcell </p>
'''
mail.Send()
print('Email Enviado')
