import os

# Pasta onde está o modelo e onde serão salvos os processados (destino)
pasta_destino = r"C:\Users\joaog\OneDrive - alphasubsea.com\destino"

PREFIXO_ARQUIVOS_GERADOS = "Feito"  # Ou use NOME_ARQUIVO_MODELO para manter o nome original
EXTENSAO_MODELO = ".xlsx"  # Extensão do arquivo modelo

# Verificar e limpar pasta de destino - manter apenas o arquivo modelo base
if os.path.exists(pasta_destino):
    for arquivo in os.listdir(pasta_destino):
        caminho_arquivo = os.path.join(pasta_destino, arquivo)
        if os.path.isfile(caminho_arquivo):
            # Manter apenas o arquivo modelo base (sem sufixo adicional)
            # Remove todas as cópias que têm sufixo após o nome do modelo
            if arquivo.startswith(f"{PREFIXO_ARQUIVOS_GERADOS}_") and arquivo.endswith(EXTENSAO_MODELO):
                try:
                    os.remove(caminho_arquivo)
                except Exception as e:
                    pass
                
# Pasta onde estão os arquivos para processar (origem)
caminho_origem = r"C:\Users\joaog\OneDrive - alphasubsea.com\Anexos"

if os.path.exists(caminho_origem):
    for arquivo in os.listdir(caminho_origem):
        caminho_arquivo = os.path.join(caminho_origem, arquivo)
        if os.path.isfile(caminho_arquivo):
            os.remove(caminho_arquivo)