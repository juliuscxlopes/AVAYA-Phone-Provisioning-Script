
README

Provisionamento_AVAYA -> Configuração e Atualização 


# 1)- EXPLICAÇÃO 
        Com o objetivo de automatizar a atualização e configuração de aparelhos AVAYA, o codigo basea-se em um TXT Padrão que esta com a formatação disponibilizada pela AVAYA como protocolo de atualização e configuração
        O codigo pega o Padrão.txt e atualiza as informações necessarias de ramais e MAC.
        Necessita de uma planilha fornecida pela area tecnica que indica ramais e mac utilizados na implantação ou na manutenção executante.
        No dictorio há uma planilha Padrão.XLSX que é a formatação pela qual deve-se seguir para o funcionamento do codigo.
                exprimindo 2 informações basicas, pré-fixo + Ramal na coluna A.
                na coluna B o MAC do aparelho.

        O Codigo susten-se na padronização de cada linha da planilha em um txt eclusivo para atualização remota de aparelhos AVAYA.



#2)- ATUALIZE sua planilha dentro dos padrões já previsto na planilha Padrão.XLSX

#3)- Ao Rodar o codigo, será criado um directorio cujo o nome é SAIDA
        Retire os TXT's e aloque nos devidos lugares para seu uso.
        exclua a pasta saida para não haver erro (directorio existente)




Atualização

#1)- Planilha não precisa ter nome 'padrão.xlsx' Basta que esteja no formato XLSX e esteja dentro do directorio base do projeto.

#2)- directorio Saída com o mesmo nome da planilha fornecida.

#3)- Janela Com mensagem de conclusão!!




        





