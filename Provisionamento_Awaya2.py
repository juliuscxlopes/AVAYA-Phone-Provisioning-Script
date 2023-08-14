import openpyxl
import os


# Pergunta o nome do projeto
nome_projeto = input("Qual o nome do projeto? ")

# Diretório base onde os arquivos estão localizados
diretorio_base = os.path.join(os.path.expanduser('~'), 'Projetos', 'Importação_AVAYA')

# Caminho completo para a planilha Excel e arquivo padrão
caminho_planilha = os.path.join(diretorio_base, 'Padrão.xlsx')
caminho_padrao = os.path.join(diretorio_base, 'Padrão.txt')

# Diretório de saída para os arquivos modificados
diretorio_saida = os.path.join(diretorio_base, nome_projeto)

# Verifica se o diretório de saída existe, senão cria
if not os.path.exists(diretorio_saida):
    os.makedirs(diretorio_saida)

# Abrindo Planilha e Buscando Folha.
book = openpyxl.load_workbook(caminho_planilha)
sheet = book.active  # Obtém a folha ativa (primeira folha por padrão)

# Lê o conteúdo do arquivo Padrão.txt
with open(caminho_padrao, 'r') as txt:
    linhas_padrao = txt.readlines()

# Loop para percorrer toda a planilha.
for row in sheet.iter_rows():
    telefone_concatenado = str(row[0].value)
    mac = str(row[1].value)

    # Remove os ":" do endereço MAC.
    mac_s_P = mac.replace(":", "")

    # Alterações de Texto no TXT PADRÃO.
    linhas_modificadas = []
    for linha in linhas_padrao:
        if 'SET FORCE_SIP_USERNAME' in linha:
            linhas_modificadas.append(f'SET FORCE_SIP_USERNAME {telefone_concatenado}\n')
        elif 'SET FORCE_SIP_EXTENSION' in linha:
            linhas_modificadas.append(f'SET FORCE_SIP_EXTENSION {telefone_concatenado}\n')
        elif 'SET SIP_USER_ACCOUNT' in linha:
            linhas_modificadas.append(f'SET SIP_USER_ACCOUNT {telefone_concatenado}@10.159.0.54\n')
        elif 'SET PREV_SIP_USER_ACCOUNT' in linha:
            linhas_modificadas.append(f'SET PREV_SIP_USER_ACCOUNT {telefone_concatenado}@10.159.0.54\n')
        elif 'SET DISPLAY_NAME' in linha:
            linhas_modificadas.append(f'SET DISPLAY_NAME {telefone_concatenado}\n')
        else:
            linhas_modificadas.append(linha)

    nome_arquivo = f"{mac_s_P}.txt"

    # Constrói o caminho completo para o destino.
    caminho_completo = os.path.join(diretorio_saida, nome_arquivo)

    with open(caminho_completo, 'w') as txt:
        txt.writelines(linhas_modificadas)

    print(f"Arquivo {nome_arquivo} modificado e salvo em {caminho_completo}")

print("Todos os arquivos foram modificados e salvos.")