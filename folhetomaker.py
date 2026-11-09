import os
import re
import traceback
import win32com.client

PASTA = os.path.dirname(os.path.abspath(__file__))


def main():
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

    ppt = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
    try:
        for i, nome in enumerate(arquivos, 1):
            origem = os.path.join(PASTA, nome)
            destino = os.path.join(SAIDA, os.path.splitext(nome)[0] + ".pdf")
            try:
                pres = ppt.Presentations.Open(
                    FileName=origem, ReadOnly=-1, Untitled=0, WithWindow=0)
                pres.ExportAsFixedFormat(
                    Path=destino,
                    FixedFormatType=2,   # PDF
                    Intent=2,            # qualidade de impressão
                    FrameSlides=-1,      # moldura nos slides
                    HandoutOrder=1,      # ordem vertical
                    OutputType=3)        # 3 slides por página com linhas
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

input("\nPressione Enter para fechar...")