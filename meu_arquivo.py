import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None) ### Mostrar o máximo de colunas

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)
print('-' * 50)

# Configurações do servidor SMTP (usando Gmail como exemplo)
smtp_server = 'smtp.gmail.com'
smtp_port = 587
email_user = 'seuemail@gmail.com'  # Seu e-mail
email_password = 'senha_app'  # Senha gerada para aplicativos, caso tenha 2FA no Gmail

# Destinatário
to_email = 'destinatario@mail.com'

# Criando a mensagem
msg = MIMEMultipart()
msg['From'] = email_user
msg['To'] = to_email
msg['Subject'] = 'Relatório de Vendas por Loja'

# Corpo do e-mail em HTML
html_body = f'''
<p>Prezados,</p>
<p>Segue o Relatório de Vendas por cada Loja.</p>

<p><strong>Faturamento:</strong></p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p><strong>Quantidade Vendida:</strong></p>
{quantidade.to_html()}

<p><strong>Ticket Médio dos Produtos em cada Loja:</strong></p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>
<p>Att.,</p>
<p>Rael Viana</p>
'''

msg.attach(MIMEText(html_body, 'html'))

# Enviando o e-mail via SMTP
try:
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()  # Criptografando a conexão
        server.login(email_user, email_password)  # Fazendo login
        server.sendmail(email_user, to_email, msg.as_string())  # Enviando o e-mail
        print("E-mail enviado com sucesso!")
except Exception as e:
    print(f"Ocorreu um erro: {e}")
