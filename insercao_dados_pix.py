import pdfplumber
import os
import time
import re
import cx_Oracle
import shutil
import unicodedata
import string
import logging
import subprocess
from datetime import datetime  

# Configuração do banco de dados Oracle
USUARIO_ORACLE = ""
SENHA_ORACLE = ""
NOME_SERVIDOR = ""
NOME_SERVICO = ""


# Configuração do logging
log_folder = "logs"
os.makedirs(log_folder, exist_ok=True)
log_filename = os.path.join(log_folder, f"log_{datetime.now().strftime('%Y%m%d')}.log")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(log_filename, encoding='utf-8'),
        logging.StreamHandler()
    ]
)


# Função para conectar ao banco

def conectar_banco():
    dsn = cx_Oracle.makedsn(NOME_SERVIDOR, port=1521, service_name=NOME_SERVICO)
    try:
        conexao = cx_Oracle.connect(USUARIO_ORACLE, SENHA_ORACLE, dsn)
        print("Conexão com o banco de dados bem-sucedida!")
        return conexao
    except cx_Oracle.DatabaseError as e:
        logging.exception("Erro ao conectar ao banco de dados:")
        return None

# Função para identificar o formato do comprovante
def identificar_formato(linhas):
    linhas_lower = [linha.lower() for linha in linhas]

    for linha in linhas_lower:
        if "dados de quem está recebendo" in linha:
            return "primeira_via"
        elif "dados do recebedor" in linha:
            return "segunda_via"
    
    print("Formato não identificado no comprovante.")
    return None

def remover_acentos_e_pontuacao(texto):
    texto = "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")
    texto = "".join(c for c in texto if c not in string.punctuation)  # Remove pontuação
    return texto

# Função para extrair os dados do comprovante
def extrair_dados(linhas, formato):
    if formato == "primeira_via":
        dados = {
            "Comprovante de transferência": {
                "Dados de quem está recebendo": {
                    "Nome": "",
                    "Chave": "",
                    "Instituição": "",
                    "Agência": "",
                    "Conta": "",
                    "Tipo de conta": "",
                    "CCB": ""
                },
                "Dados da transação": {
                    "Valor": "",
                    "Data da transferência": "",
                    "Tipo de Pagamento": ""
                },
                "Autenticação no comprovante": "",
                "ID da transação": "",
                "Controle": "",
                "Efetuada em": "",
                "Via": "1"
            }
        }
    else:  # formato == "segunda_via"
        dados = {
            "Comprovante de Transferência": {
                "dados do recebedor": {
                    "nome do recebedor": "",
                    "chave": "",
                    "instituição": "",
                    "agência": "",
                    "conta": "",
                    "tipo de conta": "",
                    "CCB": ""
                },
                "dados da transação": {
                    "valor": "",
                    "data da transferência": "",
                    "tipo de pagamento": ""
                },
                "autenticação no comprovante": "",
                "ID da transação": "",
                "controle": "",
                "transação efetuada em": "",
                "via": "2"
            }
        
        }
    bloco_recebedor = False

    for i, linha in enumerate(linhas):
        linha_lower = linha.lower()
        
        if formato == "primeira_via":
            if "dados de quem está recebendo" in linha_lower:
                    bloco_recebedor = True
            elif "dados de quem está pagando" in linha_lower:
                    bloco_recebedor = False
        else:
            if "dados do recebedor" in linha_lower:
                    bloco_recebedor = True
            elif "dados do pagador" in linha_lower:
                    bloco_recebedor = False

        if formato == "primeira_via":
            if bloco_recebedor is True:
                if "Nome" in linha:
                    dados["Comprovante de transferência"]["Dados de quem está recebendo"]["Nome"] = " ".join(linha.split(" ")[1:]).strip()
                elif "Chave" in linha:
                    dados["Comprovante de transferência"]["Dados de quem está recebendo"]["Chave"] = " ".join(linha.split(" ")[1:]).strip()
                elif "Instituição" in linha:
                    instituicao = " ".join(linha.split(" ")[1:]).strip()
                    instituicao = instituicao.replace('.', ' ')
                    dados["Comprovante de transferência"]["Dados de quem está recebendo"]["Instituição"] = remover_acentos_e_pontuacao(instituicao).rstrip().upper()
                elif "Valor" in linha:
                    valor = " ".join(linha.split(" ")[1:]).strip()
                    valor = re.sub(r'[^0-9,]', '', valor)  # Remove tudo que não é número
                    dados["Comprovante de transferência"]["Dados da transação"]["Valor"] = valor
                elif "Data da transferência" in linha:
                    dados["Comprovante de transferência"]["Dados da transação"]["Data da transferência"] = " ".join(linha.split(" ")[3:]).strip()
                elif "Tipo de Pagamento" in linha:
                    tipo_pagamento_1 = " ".join(linha.split(" ")[3:]).strip()
                    if tipo_pagamento_1 == "PIX - pagamento instantâneo":
                        dados["Comprovante de transferência"]["Dados da transação"]["Tipo de Pagamento"] = "PIX"
                elif "Autenticação no comprovante" in linha:
                    dados["Comprovante de transferência"]["Autenticação no comprovante"] = linhas[i + 1].strip()
                elif "ID da transação" in linha:
                    dados["Comprovante de transferência"]["ID da transação"] = linhas[i + 1].strip()
                elif "Controle" in linha:
                    dados["Comprovante de transferência"]["Controle"] = linhas[i + 1].strip().lstrip('0')
                elif "Efetuada em" in linha:
                    match = re.search(r'(\d{2}/\d{2}/\d{4})\s+às\s+(\d{2}:\d{2}:\d{2})', linha.strip())
                    if match:
                        dados["Comprovante de transferência"]["Efetuada em"] = f"{match.group(1)} {match.group(2)}"
            
        else:  # formato == "segunda_via"
            if bloco_recebedor is True:
                if "nome do recebedor" in linha_lower:
                    dados["Comprovante de Transferência"]["dados do recebedor"]["nome do recebedor"] = linha.split(":")[-1].strip()
                elif "chave" in linha_lower:
                    dados["Comprovante de Transferência"]["dados do recebedor"]["chave"] = linha.split(":")[-1].strip()
                elif "instituição" in linha_lower:
                    instituicao = linha.split(":")[-1].strip()
                    instituicao = instituicao.replace('.', ' ')
                    dados["Comprovante de Transferência"]["dados do recebedor"]["instituição"] = remover_acentos_e_pontuacao(instituicao).rstrip().upper()
                elif "valor" in linha_lower:
                    valor = linha.split(":")[-1].strip()
                    valor = re.sub(r'[^0-9,]', '', valor)  # Remove tudo que não é número
                    dados["Comprovante de Transferência"]["dados da transação"]["valor"] = valor
                elif "data da transferência" in linha_lower:
                    dados["Comprovante de Transferência"]["dados da transação"]["data da transferência"] = linha.split(":")[-1].strip()
                elif "tipo de pagamento" in linha_lower:
                    tipo_pagamento_2 = linha.split(":")[-1].strip()
                    if tipo_pagamento_2 == 'PIX TRANSFERENCIA':
                        dados["Comprovante de Transferência"]["dados da transação"]["tipo de pagamento"] = "PIX"
                    else:
                        dados["Comprovante de Transferência"]["dados da transação"]["tipo de pagamento"] = tipo_pagamento_2
                elif "autenticação no comprovante" in linha_lower:
                    dados["Comprovante de Transferência"]["autenticação no comprovante"] = linhas[i + 1].strip()
                elif "id da transação" in linha_lower:
                    dados["Comprovante de Transferência"]["ID da transação"] = linhas[i + 1].strip()
                elif "controle" in linha_lower:
                    dados["Comprovante de Transferência"]["controle"] = linhas[i + 1].strip().lstrip('0')
                elif "transação efetuada em" in linha_lower:
                    match = re.search(r'(\d{2}/\d{2}/\d{4})\s+às\s+(\d{2}:\d{2}:\d{2})', linha.strip())
                    if match:
                        dados["Comprovante de Transferência"]["transação efetuada em"] = f"{match.group(1)} {match.group(2)}"
            
        if formato == "primeira_via":
            if bloco_recebedor is True:
                if not dados["Comprovante de transferência"]["Dados de quem está recebendo"]["Chave"]:
                    if("agência e conta" in linha_lower):
                        dados["Comprovante de transferência"]["Dados de quem está recebendo"]["Agência"]= linha.split(" ")[3].split('/')[0].strip().lstrip('0')
                        dados["Comprovante de transferência"]["Dados de quem está recebendo"]["Conta"]= remover_acentos_e_pontuacao(linha.split(" ")[3].split('/')[1].strip().lstrip('0'))
                    elif ("tipo de conta" in linha_lower):
                        dados["Comprovante de transferência"]["Dados de quem está recebendo"]["Tipo de conta"] = " ".join(linha.split(" ")[3:]).replace("_", " ").strip().upper()
        elif formato == "segunda_via":
            if bloco_recebedor is True:
                if not dados["Comprovante de Transferência"]["dados do recebedor"]["chave"]:
                    if "agência/conta" in linha_lower:
                        dados["Comprovante de Transferência"]["dados do recebedor"]["agência"]= linha.split(":")[-1].strip().split('/')[0].strip().lstrip('0')
                        dados["Comprovante de Transferência"]["dados do recebedor"]["conta"]= remover_acentos_e_pontuacao(linha.split(":")[-1].strip().split('/')[1].strip().lstrip('0'))
                    elif ("tipo de conta" in linha_lower):
                        dados["Comprovante de Transferência"]["dados do recebedor"]["tipo de conta"] = linha.split(":")[-1].replace("_", " ").strip().upper()   
        
        
    return dados

# Função para inserir os dados no banco de dados Oracle
def inserir_dados_banco(dados, formato):
    conexao = conectar_banco()
    if not conexao:
        return
    
    cursor = conexao.cursor()

    # Verifica se o CCB já existe no banco
    ccb = None
    if formato == "primeira_via":
        ccb = dados["Comprovante de transferência"]["Dados de quem está recebendo"].get("CCB")
    else:
        ccb = dados["Comprovante de Transferência"]["dados do recebedor"].get("CCB")

    if not ccb:
        print("CCB não encontrado nos dados. Ignorando inserção.")
        return

    cursor.execute("SELECT COUNT(*) FROM vt_dados_pix WHERE ccb = :1", (ccb,))
    resultado = cursor.fetchone()

    if resultado[0] > 0:
        print(f"O CCB {ccb} já existe no banco. Nenhum dado foi inserido.")
        return

    sql = """
        INSERT INTO vt_dados_pix (
            nome_recebedor, chave, instituicao, valor, data_transferencia,
            tipo_pagamento, autenticacao, id_transacao, controle, efetuada_em, 
            agencia, conta, tipo_conta, via, ccb
        ) VALUES (:1, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11, :12, :13, :14, :15)
    """
    
    try:
        if formato == "primeira_via":
            cursor.execute(sql, (
                dados["Comprovante de transferência"]["Dados de quem está recebendo"].get("Nome", "N/A"),
                dados["Comprovante de transferência"]["Dados de quem está recebendo"].get("Chave", "N/A"),
                dados["Comprovante de transferência"]["Dados de quem está recebendo"].get("Instituição", "N/A"),
                dados["Comprovante de transferência"]["Dados da transação"].get("Valor", "0,00"),
                dados["Comprovante de transferência"]["Dados da transação"].get("Data da transferência", "N/A"),
                dados["Comprovante de transferência"]["Dados da transação"].get("Tipo de Pagamento", "N/A"),
                dados["Comprovante de transferência"].get("Autenticação no comprovante", "N/A"),
                dados["Comprovante de transferência"].get("ID da transação", "N/A"),
                dados["Comprovante de transferência"].get("Controle", "N/A"),
                dados["Comprovante de transferência"].get("Efetuada em", "N/A"),
                dados["Comprovante de transferência"]["Dados de quem está recebendo"].get("Agência", "N/A"),
                dados["Comprovante de transferência"]["Dados de quem está recebendo"].get("Conta", "N/A"),
                dados["Comprovante de transferência"]["Dados de quem está recebendo"].get("Tipo de conta", "N/A"),
                dados["Comprovante de transferência"].get("Via", "N/A"),
                ccb
            ))
        else:
            cursor.execute(sql, (
                dados["Comprovante de Transferência"]["dados do recebedor"].get("nome do recebedor", "N/A"),
                dados["Comprovante de Transferência"]["dados do recebedor"].get("chave", "N/A"),
                dados["Comprovante de Transferência"]["dados do recebedor"].get("instituição", "N/A"),
                dados["Comprovante de Transferência"]["dados da transação"].get("valor", "0,00"),
                dados["Comprovante de Transferência"]["dados da transação"].get("data da transferência", "N/A"),
                dados["Comprovante de Transferência"]["dados da transação"].get("tipo de pagamento", "N/A"),
                dados["Comprovante de Transferência"].get("autenticação no comprovante", "N/A"),
                dados["Comprovante de Transferência"].get("ID da transação", "N/A"),
                dados["Comprovante de Transferência"].get("controle", "N/A"),
                dados["Comprovante de Transferência"].get("transação efetuada em", "N/A"),
                dados["Comprovante de Transferência"]["dados do recebedor"].get("agência", "N/A"),
                dados["Comprovante de Transferência"]["dados do recebedor"].get("conta", "N/A"),
                dados["Comprovante de Transferência"]["dados do recebedor"].get("tipo de conta", "N/A"),
                dados["Comprovante de Transferência"].get("via", "N/A"),
                ccb
            ))

        conexao.commit()
        print("Dados inseridos com sucesso!")
    except cx_Oracle.Error as e:
        logging.exception(f"Erro ao inserir dados no banco:")
        conexao.rollback()
    finally:
        cursor.close()
        conexao.close()

# Função para processar um PDF
def processar_pdf(caminho_pdf):
    try:
        if not os.path.exists(caminho_pdf):
            logging.warning(f"Arquivo {caminho_pdf} não encontrado.")
            return

        linhas_extraidas = []
        nome_arquivo = os.path.basename(caminho_pdf)

        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if texto:
                    linhas_extraidas.extend(texto.split("\n"))

        if not linhas_extraidas:
            logging.warning(f"Nenhum texto extraído do arquivo {caminho_pdf}.")
            return

        formato = identificar_formato(linhas_extraidas)
        if not formato:
            logging.warning(f"Formato desconhecido para o arquivo {caminho_pdf}")
            return

        dados = extrair_dados(linhas_extraidas, formato)
        if formato == "primeira_via":
            dados["Comprovante de transferência"]["Dados de quem está recebendo"]["CCB"] = nome_arquivo.split(" ")[1]
        elif formato == "segunda_via":
            dados["Comprovante de Transferência"]["dados do recebedor"]["CCB"] = nome_arquivo.split(" ")[1]

        inserir_dados_banco(dados, formato)

        pasta_processados = r"etl_pix\insercao_dados_pix\Comprovantes processados"
        os.makedirs(pasta_processados, exist_ok=True)
        destino = os.path.join(pasta_processados, nome_arquivo)

        if os.path.exists(destino):
            os.remove(destino)

        shutil.move(caminho_pdf, destino)
        print(f"Arquivo movido para {pasta_processados}")
    except Exception as e:
        logging.exception(f"Erro ao processar o PDF: {caminho_pdf}")

# Função para monitorar a pasta de comprovantes
def monitorar_pasta():
    pasta_origem = r"etl_pix\insercao_dados_pix\Comprovante"
    print("🚀 Iniciando monitoramento...")
    
    while True:
        subprocess.run(["python", "baixar_emails.py"])  # executa o script
        for arquivo in os.listdir(pasta_origem):
            if arquivo.lower().endswith(".pdf"):
                caminho_pdf = os.path.join(pasta_origem, arquivo)
                processar_pdf(caminho_pdf)
        print(f"⏱ Aguardando {10}s...\n")
        time.sleep(10)

if __name__ == "__main__":
    monitorar_pasta()
