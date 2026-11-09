import os
import sys
import win32com.client

# Pasta com os PPTX (pode passar como argumento: python folhetos.py "C:\\Aulas")
PASTA = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
SAIDA = os.path.join(PASTA, "Folhetos")
os.makedirs(SAIDA, exist_ok=True)

arquivos = [f for f in os.listdir(PASTA)
            if f.lower().endswith((".pptx", ".ppt")) and not f.startswith("~$")]