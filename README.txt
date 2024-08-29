# AVAYA Phone Provisioning Script

Este script foi desenvolvido para automatizar o provisionamento de aparelhos telefônicos AVAYA, utilizando dados fornecidos em planilhas Excel. O script gera arquivos TXT formatados, prontos para serem importados no sistema de gerenciamento AVAYA (OSV), configurando automaticamente os aparelhos na rede.

## Funcionalidades

- **Processamento de múltiplas planilhas:** O script busca automaticamente todas as planilhas com extensão `.xlsx` no diretório especificado e processa cada uma delas individualmente.
- **Geração automática de arquivos TXT:** Para cada ramal e MAC presentes nas planilhas, o script gera um arquivo TXT formatado conforme o padrão AVAYA.
- **Solicitação de informações adicionais:** O script solicita a sobrevivência e o prefixo de ramal caso necessário, garantindo a configuração correta dos aparelhos.
- **Diretório de saída organizado:** Para cada planilha processada, é criado um diretório de saída com o nome da planilha, onde os arquivos gerados são armazenados.

## Requisitos

- **Python 3.6+**
- **Bibliotecas:** `openpyxl`, `tkinter`

## Como Usar

1. **Estrutura do Diretório:**
    - Certifique-se de que as planilhas a serem processadas estejam localizadas no diretório:
      ```
      C:\Users\SEU_USUARIO\Projetos\Importação_AVAYA
      ```
    - Coloque o arquivo `Padrão.txt` no mesmo diretório, contendo a formatação padrão para o AVAYA.

2. **Formato da Planilha:**
    - As planilhas devem conter:
        - **RAMAIS** na coluna **A**.
        - **MAC's** na coluna **B**.
    - **Atenção:** O script não formata a planilha. Certifique-se de que os dados estejam corretamente posicionados.

3. **Executar o Script:**
    - Execute o script Python:
      ```bash
      python avaya_provisioning.py
      ```
    - Um ambiente gráfico aparecerá mostrando o status de processamento e solicitando informações adicionais, como sobrevivência e prefixo.

4. **Saída:**
    - Para cada planilha processada, será criado um diretório de saída com o nome da planilha no diretório base. 
    - Cada diretório conterá os arquivos TXT gerados prontos para serem importados no sistema AVAYA.

## Estrutura do Diretório

