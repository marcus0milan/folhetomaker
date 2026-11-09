import os
import re
import tempfile
import traceback
import win32com.client.dynamic
from PIL import Image, ImageDraw

PASTA = os.path.dirname(os.path.abspath(__file__))

# True = tenta o método nativo primeiro | False = vai direto pro alternativo
USAR_NATIVO = True

# Layout A4 a 200 dpi (método alternativo)
DPI = 200
PW, PH = 1654, 2339
MARGEM = 130
CAIXA_W, CAIXA_H = 760, 600

estado = {"nativo": USAR_NATIVO}


def montar_folheto(imagens, destino):
    """3 slides por página, com linhas de anotação ao lado de cada um."""
    linha_h = (PH - 2 * MARGEM) // 3
    paginas = []
    for ini in range(0, len(imagens), 3):
        pag = Image.new("RGB", (PW, PH), "white")
        d = ImageDraw.Draw(pag)
        for pos, caminho in enumerate(imagens[ini:ini + 3]):
            img = Image.open(caminho).convert("RGB")
            img.thumbnail((CAIXA_W, CAIXA_H), Image.LANCZOS)
            x0 = MARGEM
            y0 = MARGEM + pos * linha_h + (linha_h - img.height) // 2
            pag.paste(img, (x0, y0))
            d.rectangle([x0 - 1, y0 - 1, x0 + img.width, y0 + img.height],
                        outline="black", width=2)
            xl = x0 + CAIXA_W + 60
            for k in range(1, 6):
                y = y0 + img.height * k // 5
                d.line([(xl, y), (PW - MARGEM, y)], fill="gray", width=2)
        paginas.append(pag)
    paginas[0].save(destino, "PDF", resolution=DPI, save_all=True,
                    append_images=paginas[1:], quality=90)


def converter(ppt, origem, destino):
    # ReadOnly=-1 (sim), Untitled=0, WithWindow=0 (sem janela)
    pres = ppt.Presentations.Open(origem, -1, 0, 0)
    try:
        # 1) Método nativo
        if estado["nativo"]:
            try:
                # PDF=2, Intent=2, FrameSlides=0, HandoutOrder=1, 3 slides/página
                pres.ExportAsFixedFormat(destino, 2, 2, 0, 1, 3)
                return "nativo"
            except Exception as e:
                print(f"    nativo falhou ({e}) -> usando método alternativo")
                estado["nativo"] = False
                if os.path.exists(destino):
                    os.remove(destino)

        # 2) Método alternativo: slides -> PNG -> folheto montado em Python
        w = pres.PageSetup.SlideWidth
        h = pres.PageSetup.SlideHeight
        pw = 1600
        ph = int(pw * h / w)
        with tempfile.TemporaryDirectory() as tmp:
            imagens = []
            for i in range(1, pres.Slides.Count + 1):
                p = os.path.join(tmp, f"{i:03d}.png")
                pres.Slides.Item(i).Export(p, "PNG", pw, ph)
                imagens.append(p)
            montar_folheto(imagens, destino)
        return "alternativo"
    finally:
        pres.Close()


def main():
    numeros = [int(m.group(1)) for d in os.listdir(PASTA)
               if (m := re.fullmatch(r"slides(\d+)", d))
               and os.path.isdir(os.path.join(PASTA, d))]
    execucao = max(numeros, default=0) + 1

    arquivos = [f for f in os.listdir(PASTA)
                if f.lower().endswith((".pptx", ".ppt")) and not f.startswith("~$")]
    if not arquivos:
        print(f"Nenhum arquivo PPTX encontrado em: {PASTA}")
        return

    saida = os.path.join(PASTA, f"slides{execucao}")
    os.makedirs(saida)

    ppt = win32com.client.dynamic.Dispatch("PowerPoint.Application")
    try:
        for i, nome in enumerate(arquivos, 1):
            origem = os.path.abspath(os.path.join(PASTA, nome))
            destino = os.path.abspath(os.path.join(saida, os.path.splitext(nome)[0] + ".pdf"))
            try:
                metodo = converter(ppt, origem, destino)
                print(f"[{i}/{len(arquivos)}] OK ({metodo}): {nome}")
            except Exception:
                print(f"[{i}/{len(arquivos)}] ERRO em {nome}:")
                traceback.print_exc()
    finally:
        ppt.Quit()

    print(f"\nExecução nº {execucao} concluída! PDFs em: {saida}")


try:
    main()
except Exception:
    traceback.print_exc()

input("\nPressione Enter para fechar...")