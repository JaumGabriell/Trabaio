import smtplib
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

#Exemplo de uso com anexos
caminho = r'C:\Users\joaog\Desktop\alpha\destino'

arquivos_destino = os.listdir(caminho)

caminho_completo = []

if arquivos_destino:
    for arquivo in arquivos_destino:
        caminho_arquivo = os.path.join(caminho, arquivo)
        if os.path.isfile(caminho_arquivo):
            # Filtrar para não incluir o arquivo modelo base
            if not arquivo == "Modelo Planilha Guia operação Mestra.xlsx":
                caminho_completo.append(caminho_arquivo)

def enviar_email(anexos=None):
    corpo_email = """
    <p>Olá, <b>teste de envio de email!<b></p>
    """
    
    # Criar mensagem multipart para suportar anexos
    msg = MIMEMultipart()
    msg['Subject'] = 'Teste de Envio de Email'
    msg['From'] = 'Email de quem envia '
    msg['To'] = 'email de quem recebe'
    
    # vá em segurança do google e gere uma senha de app
    password = 'Sua senha gerada no google apps'
    
    # Adicionar corpo do email
    msg.attach(MIMEText(corpo_email, 'html'))
    
    # Adicionar anexos se fornecidos
    if anexos:
        for caminho_arquivo in anexos:
            if os.path.isfile(caminho_arquivo):
                with open(caminho_arquivo, "rb") as attachment:
                    # Obter nome e extensão do arquivo
                    nome_arquivo = os.path.basename(caminho_arquivo)
                    
                    # Configurar MIME type baseado na extensão
                    if nome_arquivo.lower().endswith('.xlsx'):
                        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    elif nome_arquivo.lower().endswith('.xls'):
                        part = MIMEBase('application', 'vnd.ms-excel')
                    elif nome_arquivo.lower().endswith('.pdf'):
                        part = MIMEBase('application', 'pdf')
                    else:
                        part = MIMEBase('application', 'octet-stream')
                    
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                
                # Adicionar header para o anexo com nome correto
                part.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=nome_arquivo
                )
                
                msg.attach(part)
                #print(f"Anexo adicionado: {nome_arquivo}")
            else:
                print(f"Arquivo não encontrado: {caminho_arquivo}")
    
    s = smtplib.SMTP('smtp.gmail.com', 587)
    
    s.starttls()
    
    s.login(msg['From'], password)
    
    s.sendmail(msg['From'], msg['To'], msg.as_string())
    
    print("Email enviado com sucesso!")


# Chamar a função com todos os arquivos encontrados
if caminho_completo:
    enviar_email(caminho_completo)
else:
    print("Nenhum arquivo encontrado para anexar")
    enviar_email()  # Enviar sem anexos
