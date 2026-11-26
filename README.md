# Como montar folhetos

Este script transforma apresentações do PowerPoint em arquivos PDF prontos para imprimir como folheto. Ele monta **3 slides por página**, com linhas ao lado para anotações.

## Antes de começar

- Use um computador com **Windows** e o **Microsoft PowerPoint** instalado.
- Instale o **Python 3.8 ou superior**.
- Instale as bibliotecas necessárias. Abra o Prompt de Comando e execute:

  ```bat
  py -m pip install pywin32 Pillow
  ```

## Passo a passo

1. Coloque o arquivo `folhetomaker.py` e as apresentações `.ppt` ou `.pptx` na mesma pasta.
2. Feche os arquivos do PowerPoint que serão convertidos.
3. Abra o Prompt de Comando nessa pasta.
4. Execute:

   ```bat
   py folhetomaker.py
   ```

5. Aguarde o processamento. O resultado aparece na tela para cada apresentação.
6. Quando a mensagem pedir, pressione **Enter** para fechar a janela.
7. Abra a pasta de saída criada, por exemplo `slides1`. Os PDFs estarão dentro dela.

## O que o script faz

- Processa todos os arquivos `.ppt` e `.pptx` encontrados na pasta.
- Cria uma pasta nova para cada execução (`slides1`, `slides2` e assim por diante), sem substituir as pastas de resultados anteriores.
- Tenta primeiro usar a exportação de folhetos do PowerPoint. Se esse método falhar, tenta montar o PDF por um método alternativo.
- Mostra `OK` quando um arquivo é convertido e `ERRO` quando não consegue convertê-lo.

## Se algo der errado

- **“Nenhum arquivo PPTX encontrado”**: confira se as apresentações estão na mesma pasta que o script e se são arquivos `.ppt` ou `.pptx`.
- **Erro ao abrir ou converter**: confirme que o PowerPoint está instalado e que o arquivo abre normalmente. Feche também qualquer janela do PowerPoint que esteja usando o arquivo.
- **Biblioteca ausente**: execute novamente `py -m pip install pywin32 Pillow`.
- **Um arquivo apresentou erro**: veja os detalhes exibidos no Prompt de Comando. Os outros arquivos ainda podem ser processados.

Não apague nem mova os arquivos originais durante o processamento. O script cria PDFs; ele não altera as apresentações.
