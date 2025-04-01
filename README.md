#Bot_tesouraria

##Surgiu da necessidade de automatizar a tarefa de digitação dos caixas 
no setor da tesouraria. Com o aumento do número das lojas estava se tornando cada vez
mais complicado cumprir os prazos do fechamento deste setor daí foi criada
esta automação no sentido de otimizar a tarefa e conseguir cumprir os prazos.

##Publicação.
A aplicação esta disponível para uso no servidor 64.181.170.59 / 10.29.141.60 no caminho c:\Projetos_python\TESOURARIA\dist. Sendo 64.181.170.59 o ip externo e o 10.29.141.60 o ip interno

##Bibliotecas utilizadas.
vide arquivo requirements.txt


## Como funciona.
O Bot funciona fazendo a leitura dos dados da planilha de movimento dos pdv´s enviada pelas lojas e digitando
no módulo tesouraria da c5. 
A planilha tem que estar no formato xlsx na pasta \\10.11.10.3\arcomixfs$\Financeiro\digitacao, daí ele faz uma
copia do arquivo para uma pasta local do projeto C:\Projetos_Python\TESOURARIA\arquivos\executar e começa a digitação.
Uma vez que os arquivos das lojas estejam na pasta digitação, basta executar o ícone Financeiro_digitação na área de trabalho do
servidor ou acessar o caminho do executável : c:\Projetos_python\TESOURARIA\dist

##
Desenvolvimento interno TI-Arcomix.



