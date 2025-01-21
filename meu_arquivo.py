import pandas as pd

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None) ### Mostrar o máximo de colunas
print(tabela_vendas)

# Faturamento por loja

# Quantidade de produtos vendidos por loja

# Ticket médio por produto em cada loja

# Enviar um email com o relatório