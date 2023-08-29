Codigo desenvolvido para Atualização e configuração de aparelhos AVAYA

        - O codigo gera a partir de uma planilha com RAMAIS e MAC's arquivos TXT formatados para Importação AVAYA.

        - A formatação da Planilha é fundamental para o funcinamento do Codigo.
                -RAMAIS posicionado na coluna A
                -MAC's na coluna B
                        - O codigo nao faz formatação da planilha. Precisa que a mesma esteja com as informações formatadas conforme acima.

        - A planilha deve estar no mesmo directorio do Projeto.
                - C:\Users\' USUARIO'\Projetos\Importação_AVAYA
                - Assim o Codigo irá buscar todos os arquivos que estão com a extensão .xlsx
                - O codigo está preparado para receber diversos fluxos de planilhas. Ou seja, Mais de uma planilha que estiver no directorio.
        
        - O Codigo Cria um directorio de Saida com o mesmo nome da planilha do qual ele extrai os dados.
                - em caso de varias planilhas o codigo criará uma pasta para cada planilha.

        - Padrão.txt
                Oferece a formatação necessaria AVAYA. Qualquer alteração na formatação, não funcionará no proximo esstagio de importação.

Desenvolvido por Julius Lopes.                
