import smtplib
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Configuração do arquivo modelo (deve ser igual ao copia2.py)
ARQUIVO_MODELO_COMPLETO = "Modelo Planilha Guia operação Mestra.xlsx"

def enviar_email(anexos=None, caminho_pasta=None):
    # Se caminho_pasta for fornecido, buscar arquivos automaticamente
    if caminho_pasta and not anexos:
        caminho_completo = []
        try:
            arquivos_destino = os.listdir(caminho_pasta)
            for arquivo in arquivos_destino:
                caminho_arquivo = os.path.join(caminho_pasta, arquivo)
                if os.path.isfile(caminho_arquivo):
                    # Filtrar para não incluir o arquivo modelo base
                    if not arquivo == ARQUIVO_MODELO_COMPLETO:
                        caminho_completo.append(caminho_arquivo)
            anexos = caminho_completo
        except FileNotFoundError:
            print(f"Pasta não encontrada: {caminho_pasta}")
            return False
    
    # Verificar se há arquivos para anexar (exceto o modelo)
    if not anexos or len(anexos) == 0:
        print("❌ Nenhum arquivo processado encontrado para enviar por email.")
        print("📋 Apenas o arquivo modelo está presente na pasta destino.")
        print("🚫 Email não será enviado.")
        return False
    
    print("📧 Iniciando envio do email...")
    print(f"📎 Anexando {len(anexos)} arquivo(s) processado(s)")
    
    corpo_email = """
    <p>Olá, <b>teste de envio de email!<b></p>
    """
    
    # Criar mensagem multipart para suportar anexos
    msg = MIMEMultipart()
    msg['Subject'] = 'Teste de Envio de Email'
    msg['From'] = 'email remetente'
    msg['To'] = 'email a qual deseja enviar'
    password = 'sua senha gmail apps'
    
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
    
    print("✅ Email enviado com sucesso!")
    return True


# Execução apenas quando chamado diretamente (não quando importado)
if __name__ == "__main__":
    # Chamar a função com pasta de anexos
    caminho_destino = r'C:\Users\joaog\OneDrive - alphasubsea.com\destino'
    enviar_email(caminho_pasta=caminho_destino)
