# Importação Automática de Aparelhos Telefônicos AVAYA

Este aplicativo Python foi desenvolvido para simplificar o processo de importação de aparelhos telefônicos AVAYA. Ele automatiza a geração de arquivos de configuração com base em informações contidas em uma planilha de entrada e um modelo de script padrão AVAYA. Isso permite que você importe eficientemente vários aparelhos no sistema AVAYA com base nos dados fornecidos pela equipe técnica.

## Como Usar

1. **Estrutura de Diretórios:**

   Coloque a planilha de entrada (no formato XLSX) no diretório base do código. Os arquivos gerados serão salvos em subdiretórios com o mesmo nome das planilhas.

2. **Executando o Aplicativo:**

   Basta executar o aplicativo compilado sem a necessidade de instalar ou configurar qualquer ambiente Python adicional. Isso iniciará uma interface gráfica (GUI) onde você pode interagir com o aplicativo.

3. **Processamento de Arquivos:**

   - Clique no botão "Gerar Arquivos" para iniciar o processamento das planilhas.
   - O programa solicitará informações adicionais:
     - **Prefixo:** Insira o prefixo a ser adicionado ao número do ramal (caso necessário).
     - **Sobrevivência:** Insira as informações de sobrevivência a serem incluídas nos arquivos de configuração.
   - O aplicativo irá gerar arquivos de configuração individuais para cada linha da planilha.

4. **Resultados:**

   O programa exibirá o progresso do processamento e informará quais planilhas foram processadas com sucesso. Os arquivos gerados estarão disponíveis nos subdiretórios correspondentes às planilhas.

## Funções Principais

- **Mostrar Mensagens Iniciais:** O aplicativo lista as planilhas disponíveis e exibe a quantidade de linhas de dados em cada uma.

- **Processar Arquivos:** Esta é a função principal do aplicativo. Ela processa as planilhas, solicita informações do usuário e gera arquivos de configuração para importação no sistema AVAYA.

- **Ler Linhas Padrão:** O aplicativo lê um arquivo de texto chamado "Padrão.txt", que contém um modelo de configuração a ser usado na geração dos arquivos.

- **Criar Arquivo:** Esta função cria um arquivo de texto com base nas linhas de configuração modificadas e o salva no diretório de saída correspondente.

- **Converter Títulos para Minúsculas:** Os títulos das colunas na planilha são convertidos para letras minúsculas para facilitar o processamento.

- **Solicitar Sobrevivência e Prefixo:** O aplicativo solicita ao usuário que insira informações importantes, como prefixo e sobrevivência, que serão incorporados aos arquivos de configuração.

- **Processar Planilha:** Esta função realiza o processamento específico da planilha, criando arquivos de configuração com base nos dados da planilha e nas informações fornecidas pelo usuário.

## Contribuindo

Sinta-se à vontade para contribuir com melhorias ou correções para este aplicativo. Basta abrir um problema ou enviar uma solicitação pull.

## Licença

Este aplicativo é distribuído sob a [Licença MIT](LICENSE).


Desenvolvido por Julius Lopes - MetodoTelecom.
Juliuscxlopes@gmail.com
                        
