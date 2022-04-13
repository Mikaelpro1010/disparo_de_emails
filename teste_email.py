import pandas as pd
import os
import smtplib
from email.message import EmailMessage

df = pd.read_excel('C:\\Users\\Prefeitura\\Desktop\\MIkael\\Programacao\\planilha-disparo-de-e-mails.xlsx')

df_index = df.iloc[0].index



#Configurando o e-mail e a senha do usuário
email_address = 'mikaelantonio398@gmail.com'
email_password = '********'

for i in range(len(df)):
    message = EmailMessage()
    message['Subject'] = df.at[i,'Assunto']
    message['From'] = 'mikaelantonio398@gmail.com'
    message['to'] = df.at[i, 'Destinatário']
    message.set_content(df.at[i, 'Conteúdo'])

    with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
        smtp.login(email_address, email_password)
        smtp.send_message(message)
    print('enviado o email: '+ str(i + 1))
