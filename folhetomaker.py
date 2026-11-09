import os
import re
import sys
import traceback
import win32com.client

# Pasta onde o próprio .py está salvo (funciona com duplo clique, VS Code, etc.)
PASTA = os.path.dirname(os.path.abspath(__file__))


def main():
    # Descobre o próximo número: slides1, slides2, slides3...
    numeros = [int(m.group(1)) for d in os.listdir(PASTA)
               if (m := re.fullmatch(r"slides(\d+)", d)) and os.path.isdir(os.path.join(PASTA, d))]
    execucao = max(numeros, default=0) + 1

    arquivos = [f for f in os.listdir(PASTA)
                if f.lower().endswith((".pptx", ".ppt")) and not f.startswith("~$")]

    if not arquivos:
        print(f"Nenhum arquivo PPTX encontrado em: {PASTA}")
        return

    SAIDA = os.path.join(PASTA, f"slides{execucao}")
    os.makedirs(SAIDA)

    ppt = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for i, nome in enumerate(arquivos, 1):
            origem = os.path.join(PASTA, nome)
            destino = os.path.join(SAIDA, os.path.splitext(nome)[0] + ".pdf")
            try:
                pres = ppt.Presentations.Open(origem, True, False, False)
                # PDF=2, Print=2, FrameSlides=False, HandoutOrder=1, 3 slides/página
                pres.ExportAsFixedFormat(destino, 2, 2, False, 1, 3)
                pres.Close()
                print(f"[{i}/{len(arquivos)}] OK: {nome}")
            except Exception as e:
                print(f"[{i}/{len(arquivos)}] ERRO em {nome}: {e}")
    finally:
        ppt.Quit()

    print(f"\nExecução nº {execucao} concluída! PDFs em: {SAIDA}")


try:
    main()
except Exception:
    traceback.print_exc()

# Mantém a janela aberta para você ler o resultado
input("\nPressione Enter para fechar...")