# GET TRANSACTION FROM API TO EXCEL

## Descrição

Este projeto tem como objetivo automatizar a extração de dados de transações a partir de uma API e armazená-los em um arquivo Excel no formato `.xlsx`.

O sistema realiza a coleta dos dados e organiza automaticamente as informações em diferentes planilhas, de acordo com o estado de pagamento das transações, tais como:

- Sucesso
- Falha
- Expirada
- Outros estados.

Essa abordagem facilita a análise, organização e acompanhamento das transações.

## Tecnologias e Bibliotecas

- `requests` : Consumo de APIs  
- `pandas` : Manipulação e estruturação de dados  
- `openpyxl` : Criação e edição de arquivos Excel  
- `os` : Interação com o sistema operativo  
- `python-dotenv` : Gestão de variáveis de ambiente  
- `logging` : Monitoramento e registro de eventos do sistema 

## Funcionalidades
- Conexão com API para obtenção de dados  
- Extração automática de transações  
- Organização dos dados por status de pagamento  
- Geração de múltiplas planilhas no Excel  
- Registro de logs para acompanhamento do processo  
