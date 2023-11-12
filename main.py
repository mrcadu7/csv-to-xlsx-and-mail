import os
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

caminho = 'bases/'
arquivos = os.listdir(caminho)

tabela = pd.DataFrame()

for nome_arquivo in arquivos:
    tabela_vendas = pd.read_csv(os.path.join(caminho, nome_arquivo))
    tabela_vendas['Data de Venda'] = pd.to_datetime('01/01/1900') + pd.to_timedelta(
        tabela_vendas['Data de Venda'], unit='d'
    )
    tabela = pd.concat([tabela, tabela_vendas])
    
tabela = tabela.sort_values(by='Data de Venda')
tabela = tabela.reset_index(drop=True)
tabela.to_excel('Vendas.xlsx', index=False)



# Configurações do email
meu_email = 'xxx@gmail.com'  # Substitua pelo seu email
minha_senha = 'xxxx xxxx xxxx xxxx'  # Substitua pela sua senha de app gmail
destinatario_email = 'yyy@xxx.com'  # Substitua pelo email do destinatário
data_hoje = datetime.today().strftime('%d/%m/%Y')

msg = MIMEMultipart()
msg['From'] = meu_email
msg['To'] = destinatario_email
msg['Subject'] = f'Relatório de Vendas {data_hoje}'

msg.attach(MIMEText(f"""
Prezados,

segue o relatório de vendas de {data_hoje} atualizado.
Qualquer duvida estou a disposição.
abs,
ADM
""", 'plain'))

# Anexar o arquivo
part = MIMEBase('application', "octet-stream")
part.set_payload(open("Vendas.xlsx", "rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment', filename="Vendas.xlsx")  # ou outro nome que você quiser
msg.attach(part)

# Enviar o email
server = smtplib.SMTP('smtp.gmail.com: 587')
server.starttls()
server.login(meu_email, minha_senha)
server.sendmail(meu_email, destinatario_email, msg.as_string())
server.quit()
