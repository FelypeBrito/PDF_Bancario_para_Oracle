
Extrator e Inseridor de Dados de Comprovantes PIX (PDF) via Email
---------------------------------------------------------------------------

📌 Visão Geral

Este projeto consiste em um sistema Python automatizado que realiza duas tarefas principais:

1. Leitura de e-mails via protocolo POP3 e salvamento de anexos (comprovantes em PDF) localmente.
2. Extração de informações de comprovantes de transferência PIX (primeira ou segunda via) desses PDFs 
   e inserção dos dados extraídos em uma tabela Oracle.

📂 Estrutura do Projeto

- extrair_e_inserir.py: Script principal que:
  - Lê os comprovantes (PDFs) salvos localmente;
  - Identifica se são primeira ou segunda via;
  - Extrai os dados relevantes como nome do recebedor, chave Pix, valor, data, ID da transação, etc.;
  - Insere esses dados na tabela VT_DADOS_PIX, se ainda não houver CCB duplicado.

- monitorar_email.py: Script auxiliar que:
  - Acessa uma conta de e-mail via POP3;
  - Lê as mensagens recebidas;
  - Baixa e salva os arquivos PDF anexados que correspondem a comprovantes de transferência;
  - Move os arquivos processados para a pasta de backup após salvar.

- logs/: Pasta onde são gerados arquivos de log com todos os eventos (execuções, erros, confirmações de inserções).

📥 Informações Extraídas

O sistema suporta comprovantes em dois formatos diferentes (primeira via e segunda via) e coleta os seguintes dados:

- Nome do recebedor
- Chave PIX
- Instituição
- Agência / Conta
- Tipo de conta
- Valor da transferência
- Data e hora
- Tipo de pagamento (PIX)
- ID da transação
- Controle
- Autenticação
- CCB (Cédula de Crédito Bancário)

🧠 Funcionalidades Inteligentes

- Identificação automática do formato do comprovante.
- Normalização de texto, com remoção de acentos e pontuação para padronização.
- Verificação de duplicidade via CCB no banco antes da inserção.
- Log detalhado de todas as execuções e exceções no diretório logs/.

🛠️ Requisitos

- Python 3.8+
- Oracle Client e biblioteca cx_Oracle
- Bibliotecas Python:
  - pdfplumber
  - cx_Oracle
  - logging
  - poplib
  - email
  - unicodedata
  - re, os, time, datetime, etc.

⚙️ Configuração

Edite no script as seguintes variáveis:

Para conexão com o banco Oracle:

USUARIO_ORACLE = "seu_usuario"
SENHA_ORACLE = "sua_senha"
NOME_SERVIDOR = "seu_host"
NOME_SERVICO = "servico_ou_sid"

Para acesso ao e-mail:

No script de monitoramento, configure:

EMAIL = "seu_email"
SENHA = "sua_senha"
SERVIDOR_POP = "pop.seuprovedor.com"
PORTA_POP = 995

🧪 Execução

1. Execute o script de leitura de e-mails (se desejar automatizar o download dos PDFs):
   python monitorar_email.py

2. Execute o script de extração e inserção:
   python extrair_e_inserir.py

📋 Observações

- Os comprovantes devem estar em formato de texto legível (não imagem).
- Certifique-se de que os nomes dos arquivos não se repitam ou sobrescrevam.
- Os logs são gravados com data, hora e nível de severidade (INFO/ERROR) no diretório logs/.

