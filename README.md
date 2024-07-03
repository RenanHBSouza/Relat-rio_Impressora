# Relatorio_Impressora

## Motivação

A motivação para a criação deste sistema foi a necessidade de monitorar e analisar o uso das impressoras de uma empresa de uma forma mais barata que diversos softwares do mercado, visando controlar os custos de impressão e identificar padrões de uso entre os funcionários. O sistema permite coletar dados detalhados sobre as impressões realizadas, gerando relatórios que auxiliam na gestão eficiente dos recursos de impressão.


## Descrição do Sistema

O sistema é composto por três partes principais:

1. **Arquivo PowerShell**: Coleta logs de impressão (eventos com ID 307) e salva em um arquivo CSV.
2. **Arquivo Python para Envio de E-mail**: Envia o CSV gerado via e-mail utilizando as bibliotecas `smtplib` e `email`.
3. **Arquivo Python para Geração de Relatório**: Gera um relatório a partir do CSV utilizando as bibliotecas `pandas`, `tkinter`, `os` e `docx`.


## Tecnologias Utilizadas

- **PowerShell**: Utilizado para a coleta de logs de impressão.
- **Python**: Linguagem principal para o envio de e-mails e geração de relatórios.
  - **smtplib**: Biblioteca para envio de e-mails.
  - **email**: Biblioteca para manipulação de conteúdo de e-mail.
  - **pandas**: Biblioteca para manipulação e análise de dados.
  - **tkinter**: Biblioteca para criação de interfaces gráficas.
  - **os**: Biblioteca para interações com o sistema operacional.
  - **python-docx**: Biblioteca para criação e manipulação de documentos do Word.


## Funcionamento

### 1. Coleta de Logs de Impressão (PowerShell)

Um script PowerShell coleta logs de impressão (eventos com ID 307) e salva esses dados em um arquivo CSV. Este arquivo contém informações detalhadas sobre cada evento de impressão, incluindo o ID do usuário, o número de páginas impressas e o nome do documento.

### 2. Envio de CSV por E-mail (Python)

Um script Python lê o arquivo CSV gerado pelo PowerShell e envia-o para um e-mail específico utilizando as bibliotecas `smtplib` e `email`. Este script é responsável por garantir que os dados coletados sejam entregues ao destinatário desejado para análise posterior.

### 3. Geração de Relatório (Python)

Outro script Python lê o arquivo CSV, processa os dados utilizando a biblioteca `pandas` e gera um relatório detalhado em formato de documento do Word (`.docx`). A interface gráfica para seleção de arquivos é feita com `tkinter`, e interações com o sistema de arquivos são realizadas com a biblioteca `os`.


## Passo-a-Passo de Utilização

### 1. Instalação e configuração

- Execute o arquivo install.bat para instalar todas as bibliotecas necessárias para a aplicação.
- Personalize o script PowerShell /Source/export_log.ps1 alterando o caminho de $logPath para o local que deseja salvar o arquivo CSV.
- Personalize o arquivo /Email/Send_email.py alterando as variáveis de configuração para o que você necessita:
  - Configuração do servidor SMTP;
  - Configuração de email e senha do remetente;
  - E-mail destinatário, assunto, corpo e caminho do csv para ser anexado;
- Edite o cabeçalho do arquivo Gerar_report.py para as suas necessidades, como valor de impressão e impressora buscada.

### 2. Coleta de Logs

- Execute o arquivo /Source/log_start.bat para dar start na aplicação e gerar o arquivo CSV. 

### 3. Envio de E-mail

- Execute o arquivo /Email/Start.bat para fazer o envio do email.

### 4. Geração do relatório

- Execute o arquivo Start.bat para gerar um relatório da impressora.


## Conclusão

O sistema de monitoramento de impressora oferece uma solução completa para monitorar custos, promover a conscientização e aprimorar a gestão de recursos de impressão em uma empresa. Através da coleta automatizada de dados, geração de relatórios e envio de emails, a equipe obtém as informações necessárias para tomar decisões conscientes e buscar reduzir despesas desnecessárias.
