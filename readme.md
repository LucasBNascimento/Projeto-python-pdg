# Projeto
  Projeto desenvolvido em Python com o intuito de coletar dados sobre a empresa PDG e suas empresas satélite.
 
 
## Como funciona:
  O programa irá pegar uma base de CNPJ's de uma planilha do Excel e fará uma pesquisa no site da Receita federal para extrair dados. A partir desses dados, uma nova planilha será montada.
  
  
## Como utilizar o programa: 
  1. Na linha 9, escrever o local da planilha com os CNPJ's que serão pesquisados no site da Receita Federal.
      * Por exemplo: `caminho = '/Users/Lucas Nascimento/Documents/Documentação PDG/CNPJ-PDG.xlsx'`
  2. Na linha 16, usar o valor de Z para definir a partir de qual linha o programa irá ler os CNPJ's
  3. Na linha 19, colocar de 'column' para a coluna que você deseja que os dados sejam extraidos.
      * Por exemplo: `column = 1` para coluna A, `column = 2` para a coluna B etc.
  4. Na linha 65, escolher o nome da planinha com as informações extraidas do site da Receita Federal.
      * Por exemplo: `arquivo_excel.save("Informações PDG.xlsx")`
 
 
## Bibliotecas necessárias 
  * openpyxl
  * requests
  * time
