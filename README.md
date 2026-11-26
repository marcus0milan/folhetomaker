# Folhetomaker

Transforma apresentações do PowerPoint **e arquivos PDF** em folhetos A4 prontos para imprimir: por padrão **3 slides por página**, com linhas ao lado para anotações.

O folheto é gerado em formato vetorial, então o texto fica nítido na impressão e o arquivo final fica leve.

## Requisitos

- **Python 3.8 ou superior**.
- Bibliotecas (no Prompt de Comando):

  ```bat
  py -m pip install pymupdf pywin32
  ```

- **Microsoft PowerPoint** (Windows), necessário só para converter apresentações. Para PDFs não é preciso.

## Formatos aceitos

- PowerPoint: `.pptx`, `.ppt`, `.pptm`, `.ppsx`, `.pps`, `.ppsm`, `.potx`, `.pot`, `.potm`, `.odp`
- PDF: `.pdf` (cada página é tratada como um slide)

## Como usar

**Jeito mais simples:** coloque o `folhetomaker.py` na pasta das apresentações e dê dois cliques nele, ou rode:

```bat
py folhetomaker.py
```

**Arrastar e soltar:** arraste arquivos ou pastas para cima do `folhetomaker.py`.

**Pela linha de comando:**

```bat
py folhetomaker.py aula1.pptx apostila.pdf      :: só esses arquivos
py folhetomaker.py C:\Aulas                       :: todos os arquivos de uma pasta
py folhetomaker.py C:\Aulas -n 6                  :: 6 slides por página
py folhetomaker.py aula.pptx -n 1                 :: 1 slide grande + linhas embaixo
py folhetomaker.py C:\Aulas -o C:\Impressao       :: escolher a pasta de saída
```

### Opções

| Opção | O que faz |
|---|---|
| `-n`, `--por-pagina` | Slides por página: `1`, `2`, `3`, `4`, `6` ou `9` (padrão `3`). Há linhas de anotação nos modos `1` e `3`. |
| `-o`, `--saida` | Pasta de saída. Sem ela, é criada uma pasta nova (`slides1`, `slides2`...) ao lado dos arquivos. |
| `--sem-linhas` | Não desenha as linhas de anotação. |
| `--sem-cabecalho` | Não coloca o nome do arquivo no topo nem o número da página no rodapé. |
| `--incluir-ocultos` | Inclui os slides ocultos (por padrão eles ficam de fora, como na impressão do PowerPoint). |
| `--sem-pausa` | Não espera o Enter no final (útil em scripts e automações). |

## O que o script faz

- Converte cada arquivo e mostra `OK` ou `ERRO`. Um erro em um arquivo não interrompe os demais.
- Nunca altera as apresentações originais (elas são abertas como somente leitura).
- Nunca sobrescreve resultados: se já existir um PDF com o mesmo nome, cria `nome (2).pdf`.
- Se o PowerPoint já estava aberto, ele **não** é fechado no final.
- Para gerar o PDF dos slides, tenta a exportação do PowerPoint; se ela falhar, salva como PDF; em último caso, exporta cada slide como imagem.
- Termina com código `0` (tudo OK), `1` (nada encontrado ou faltam bibliotecas) ou `2` (algum arquivo com erro).

## Se algo der errado

- **"Nenhuma apresentação ou PDF encontrado"**: confira a pasta e a extensão dos arquivos.
- **"PyMuPDF não instalada" / "pywin32 não instalado"**: rode `py -m pip install pymupdf pywin32`.
- **"não foi possível iniciar o PowerPoint"**: confirme que o PowerPoint está instalado e abre normalmente.
- **Erro em um arquivo específico**: verifique se ele abre no PowerPoint (ou no leitor de PDF) e se não tem senha. Para ver os detalhes técnicos do erro, rode antes `set FOLHETO_DEBUG=1`.
