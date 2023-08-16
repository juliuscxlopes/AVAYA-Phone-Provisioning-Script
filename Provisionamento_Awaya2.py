import openpyxl
import os
import glob
import tkinter as tk
from tkinter import messagebox, scrolledtext

# Função para fechar a janela Tkinter
def fechar_janela():
    root.destroy()

# Função principal para processar os arquivos
def processar_arquivos():
    # Diretório base onde os arquivos estão localizados
    diretorio_base = os.path.join(os.path.expanduser('~'), 'Projetos', 'Importação_AVAYA')

    # Busca por arquivos XLSX no diretório base
    arquivos_xlsx = glob.glob(os.path.join(diretorio_base, '*.xlsx'))

    # Verifica se foram encontrados arquivos XLSX
    if not arquivos_xlsx:
        output_text.insert(tk.END, "Nenhum arquivo XLSX encontrado no diretório base.\n")
    else:
        # Lê o conteúdo do arquivo Padrão.txt
        caminho_padrao = os.path.join(diretorio_base, 'Padrão.txt')
        with open(caminho_padrao, 'r') as txt:
            linhas_padrao = txt.readlines()

        # Loop para processar cada planilha encontrada
        for caminho_planilha in arquivos_xlsx:
            # Extrai o nome da planilha (sem extensão)
            nome_planilha = os.path.splitext(os.path.basename(caminho_planilha))[0]

            # Diretório de saída para os arquivos modificados
            diretorio_saida = os.path.join(diretorio_base, nome_planilha)

            # Verifica se o diretório de saída existe, senão cria
            if not os.path.exists(diretorio_saida):
                os.makedirs(diretorio_saida)

            # Carrega a planilha
            planilha = openpyxl.load_workbook(caminho_planilha)
            folha = planilha.active  # Obtém a folha ativa (primeira folha por padrão)

            # Lista para armazenar nomes de arquivos gerados com sucesso
            arquivos_gerados = []

            # Loop para percorrer toda a planilha.
            for row in folha.iter_rows():
                telefone_concatenado = str(row[0].value)
                mac = str(row[1].value)
                
                # Verifica se a linha está vazia ou se não há telefone ou mac
                if not telefone_concatenado or not mac:
                    continue

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

                    # Armazena o nome do arquivo gerado com sucesso
                    arquivos_gerados.append(nome_arquivo)

            output_text.insert(tk.END, f"Todos os ramais da \"{nome_planilha}\" foram criados e salvos:\n")
            for arquivo in arquivos_gerados:
                output_text.insert(tk.END, f"- {arquivo}\n")
            output_text.insert(tk.END, "Arquivos gerados com sucesso!!\n\n")

# Cria uma janela Tkinter
root = tk.Tk()
root.title("Status de Processamento")

# Cria um widget de texto rolável para exibir a saída
output_text = scrolledtext.ScrolledText(root, wrap=tk.WORD)
output_text.pack(fill=tk.BOTH, expand=True)

# Botão para processar os arquivos
botao_processar = tk.Button(root, text="Processar Arquivos", command=processar_arquivos)
botao_processar.pack()

# Inicia o loop principal do Tkinter
root.mainloop()
