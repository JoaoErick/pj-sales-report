import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# -----------------Constantes-----------------
# Configuração de E-mail
HOST = '<Host>'
PORT = '<Port>'
USER = '<Username>'
PASSWORD = '<Password>'
# --------------------------------------------


# Importar a base de dados
sales_table = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(sales_table)

print('-' * 50)

# Faturamento por loja
revenues = sales_table[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(revenues)

print('-' * 50)

# Quantidade de produtos vendidos por loja
amount = sales_table[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(amount)

print('-' * 50)

# Ticket médio por produto em cada loja
average_ticket = (revenues['Valor Final'] / amount['Quantidade']).to_frame()
average_ticket = average_ticket.rename(columns={0: 'Ticket Médio'})
print(average_ticket)

# Enviar um e-mail com o relatório

# Criando objeto
server = smtplib.SMTP(HOST, PORT)

# Login com servidor
server.ehlo()
server.login(USER, PASSWORD)

# Criando mensagem
message = f'''
<html>
  <head></head>
  <body>
    <p>Prezados,</p>

    <p>Segue o Relatório de Vendas por cada Loja.</p>

    <p><h3>Faturamento:</h3></p>
    {revenues.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format})}

    <p><h3>Quantidade Vendida:</h3></p>
    {amount.to_html()}

    <p><h3>Ticket Médio dos Produtos em cada Loja:</h3></p>
    {average_ticket.to_html(formatters={'Ticket Médio': 'R$ {:,.2f}'.format})}

    <p>Qualquer dúvida, estou à disposição.</p>

    <p>Att.,</p>
    <p>João Erick Barbosa.</p>
  </body>
</html>
'''

email_msg = MIMEMultipart()
email_msg['From'] = USER
email_msg['To'] = 'jerick1415@outlook.com'
email_msg['Subject'] = 'Relatório de Vendas'

email_msg.attach(MIMEText(message, 'html'))

# Enviando mensagem
print('\nEnviando mensagem...')
server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
print('Mensagem enviada!')
server.quit()