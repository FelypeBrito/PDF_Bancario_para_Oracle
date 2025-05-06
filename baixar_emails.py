import poplib
import time
import os
import re
import email
import hashlib
from email.parser import BytesParser
from email.header import decode_header
from datetime import datetime

# Configurações do servidor e e-mail
EMAIL = ""
PASSWORD = ""
POP3_SERVER = ""
POP3_PORT = 
PASTA_ANEXOS = r"etl_pix\insercao_dados_pix\Comprovante"
INTERVALO_SEGUNDOS = 30
ARQUIVO_HASHES = "emails_processados.txt"

# Garante que a pasta de anexos existe
os.makedirs(PASTA_ANEXOS, exist_ok=True)

# Carrega hashes de e-mails já processados
def carregar_hashes():
    if not os.path.exists(ARQUIVO_HASHES):
        return set()
    with open(ARQUIVO_HASHES, "r") as f:
        return set(linha.strip() for linha in f.readlines())

# Salva novo hash
def salvar_hash(hash_msg):
    with open(ARQUIVO_HASHES, "a") as f:
        f.write(hash_msg + "\n")

# Decodifica headers
def decodificar(texto):
    partes = decode_header(texto or "")
    resultado = ""
    for parte, encoding in partes:
        if isinstance(parte, bytes):
            resultado += parte.decode(encoding or "utf-8", errors="ignore")
        else:
            resultado += parte
    return resultado

# Salva os anexos
from datetime import datetime

def salvar_anexos(msg):
    for parte in msg.walk():
        if parte.get_content_maintype() == "multipart":
            continue
        if parte.get("Content-Disposition") is None:
            continue

        nome_arquivo = parte.get_filename()
        if nome_arquivo:
            nome_arquivo = decodificar(nome_arquivo)
            nome_arquivo = re.sub(r'[<>:"/\\|?*\n\r]', '_', nome_arquivo).strip()
            if nome_arquivo.lower().endswith(".pdf"):
                caminho = os.path.join(PASTA_ANEXOS, nome_arquivo)
                with open(caminho, "wb") as f:
                    f.write(parte.get_payload(decode=True))
                print(f"✔ PDF salvo: {caminho}")
            else:
                print(f"⏭ Ignorado (não é PDF): {nome_arquivo}")
                log_ignorado(nome_arquivo)

def log_ignorado(nome_arquivo):
    log_path = "anexos_ignorados.txt"
    data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"[{data_hora}] {nome_arquivo}\n")


# Verifica e processa novos e-mails
def verificar_emails(hashes_existentes):
    try:
        conexao = poplib.POP3_SSL(POP3_SERVER, POP3_PORT)
        conexao.user(EMAIL)
        conexao.pass_(PASSWORD)
        total_mensagens = len(conexao.list()[1])
        print(f"📬 Verificando {total_mensagens} mensagens...")

        for i in range(total_mensagens):
            resposta = conexao.retr(i + 1)
            conteudo_email = b"\n".join(resposta[1])

            # Gera hash do conteúdo do e-mail
            hash_msg = hashlib.md5(conteudo_email).hexdigest()
            if hash_msg in hashes_existentes:
                continue  # Já foi processado

            msg = BytesParser().parsebytes(conteudo_email)
            assunto = decodificar(msg["Subject"])
            if "Novo Empréstimo" in assunto:
                print(f"📨 Novo e-mail: {assunto}")
                salvar_anexos(msg)
            else:
                print(f"⏭ Ignorado: {assunto}")

            salvar_hash(hash_msg)
            hashes_existentes.add(hash_msg)

        conexao.quit()
    except Exception as e:
        print("❌ Erro ao verificar e-mails:", e)

# Loop contínuo
if __name__ == "__main__":
    hashes = carregar_hashes()
    verificar_emails(hashes)
    
