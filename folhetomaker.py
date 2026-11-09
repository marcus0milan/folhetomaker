import os
import sys
import win32com.client

# Pasta com os PPTX (pode passar como argumento: python folhetos.py "C:\\Aulas")
PASTA = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
SAIDA = os.path.join(PASTA, "Folhetos")
os.makedirs(SAIDA, exist_ok=True)

arquivos = [f for f in os.listdir(PASTA)
            if f.lower().endswith((".pptx", ".ppt")) and not f.startswith("~$")]

ppt = win32com.client.Dispatch("PowerPoint.Application")

try:
    for i, nome in enumerate(arquivos, 1):
        origem = os.path.abspath(os.path.join(PASTA, nome))
        destino = os.path.abspath(os.path.join(SAIDA, os.path.splitext(nome)[0] + ".pdf"))
        try:
            # ReadOnly=True, Untitled=False, WithWindow=False
            pres = ppt.Presentations.Open(origem, True, False, False)
            # (Caminho, PDF=2, Intent=Print=2, FrameSlides=False, HandoutOrder=1, OutputType=3 slides/página)
            pres.ExportAsFixedFormat(destino, 2, 2, False, 1, 3)
            pres.Close()
            print(f"[{i}/{len(arquivos)}] OK: {nome}")
        except Exception as e:
            print(f"[{i}/{len(arquivos)}] ERRO em {nome}: {e}")
finally:
    ppt.Quit()

print(f"\nConcluído! PDFs salvos em: {SAIDA}")