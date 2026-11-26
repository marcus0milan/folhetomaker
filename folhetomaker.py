"""
Folhetomaker - transforma apresentações (PowerPoint) e PDFs em folhetos A4
prontos para imprimir, com N slides por página e linhas para anotação.

Uso rápido:
    py folhetomaker.py                      -> converte tudo da pasta do script
    py folhetomaker.py aula.pptx slides.pdf -> converte só esses arquivos
    py folhetomaker.py C:\\Aulas -n 6        -> 6 slides/página, pasta inteira
    (também dá para arrastar arquivos/pastas em cima do folhetomaker.py)

Veja todas as opções com:  py folhetomaker.py --help
"""

import argparse
import os
import re
import sys
import tempfile
import time
import traceback

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:
    try:
        import fitz  # PyMuPDF antigo
    except ImportError:
        fitz = None

PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))

EXT_POWERPOINT = (".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm",
                  ".pot", ".potx", ".potm", ".odp")
EXT_PDF = (".pdf",)

# Página A4 em pontos (1 pt = 1/72 pol)
A4_W, A4_H = 595.28, 841.89
MARGEM = 36
CABECALHO = 16
RODAPE = 16
ESPACO = 14

# slides por página -> (colunas, linhas, onde ficam as linhas de anotação)
LAYOUTS = {
    1: (1, 1, "abaixo"),
    2: (1, 2, None),
    3: (1, 3, "direita"),
    4: (2, 2, None),
    6: (2, 3, None),
    9: (3, 3, None),
}

# Constantes do PowerPoint
PP_FIXED_PDF = 2
PP_INTENT_PRINT = 2
PP_SAVE_AS_PDF = 32
PP_ALERTS_NONE = 1
MSO_TRUE, MSO_FALSE = -1, 0


# --------------------------------------------------------------------------
# Montagem do folheto (vetorial, via PyMuPDF)
# --------------------------------------------------------------------------

def encaixar(origem_w, origem_h, area, alinhar_esq=False, alinhar_topo=False):
    """Retângulo com a proporção do slide, cabendo dentro de `area`."""
    escala = min(area.width / origem_w, area.height / origem_h)
    w, h = origem_w * escala, origem_h * escala
    x0 = area.x0 if alinhar_esq else area.x0 + (area.width - w) / 2
    y0 = area.y0 if alinhar_topo else area.y0 + (area.height - h) / 2
    return fitz.Rect(x0, y0, x0 + w, y0 + h)


def desenhar_linhas(pagina, x0, x1, y0, y1, espaco):
    y = y0 + espaco
    while y <= y1 + 0.1:
        pagina.draw_line((x0, y), (x1, y), color=(0.6, 0.6, 0.6), width=0.5)
        y += espaco


def montar_folheto(fonte, destino, por_pagina=3, linhas=True,
                   cabecalho=True, titulo=""):
    """Monta o folheto a partir de um PDF de slides (1 slide por página)."""
    cols, rows, notas = LAYOUTS[por_pagina]
    if not linhas:
        notas = None
    total = fonte.page_count
    if total == 0:
        raise ValueError("o arquivo não tem nenhum slide/página")

    saida = fitz.open()
    n_paginas = (total + por_pagina - 1) // por_pagina
    topo = MARGEM + (CABECALHO if cabecalho else 0)
    base = A4_H - MARGEM - (RODAPE if cabecalho else 0)
    cel_w = (A4_W - 2 * MARGEM - ESPACO * (cols - 1)) / cols
    cel_h = (base - topo - ESPACO * (rows - 1)) / rows

    for np_ in range(n_paginas):
        pag = saida.new_page(width=A4_W, height=A4_H)
        if cabecalho:
            pag.insert_text((MARGEM, MARGEM + 8), titulo[:110],
                            fontsize=8, color=(0.35, 0.35, 0.35))
            rotulo = f"{np_ + 1} / {n_paginas}"
            pag.insert_text((A4_W - MARGEM - 4.4 * len(rotulo), A4_H - MARGEM),
                            rotulo, fontsize=8, color=(0.35, 0.35, 0.35))

        for pos in range(por_pagina):
            idx = np_ * por_pagina + pos
            if idx >= total:
                break
            c, r = pos % cols, pos // cols
            cx0 = MARGEM + c * (cel_w + ESPACO)
            cy0 = topo + r * (cel_h + ESPACO)
            celula = fitz.Rect(cx0, cy0, cx0 + cel_w, cy0 + cel_h)
            sw, sh = fonte[idx].rect.width, fonte[idx].rect.height

            if notas == "direita":
                area = fitz.Rect(celula.x0, celula.y0,
                                 celula.x0 + celula.width * 0.55, celula.y1)
                ret = encaixar(sw, sh, area, alinhar_esq=True)
            elif notas == "abaixo":
                area = fitz.Rect(celula.x0, celula.y0,
                                 celula.x1, celula.y0 + celula.height * 0.55)
                ret = encaixar(sw, sh, area, alinhar_topo=True)
            else:
                ret = encaixar(sw, sh, celula)

            pag.show_pdf_page(ret, fonte, idx)
            pag.draw_rect(ret, color=(0.3, 0.3, 0.3), width=0.6)

            if notas == "direita":
                desenhar_linhas(pag, ret.x1 + 16, celula.x1,
                                ret.y0, ret.y1, 20)
            elif notas == "abaixo":
                desenhar_linhas(pag, celula.x0, celula.x1,
                                ret.y1 + 6, celula.y1, 24)

    saida.save(destino, garbage=3, deflate=True)
    saida.close()
    return n_paginas


# --------------------------------------------------------------------------
# PowerPoint -> PDF (1 slide por página)
# --------------------------------------------------------------------------

class PowerPoint:
    """Abre o PowerPoint só quando necessário e não fecha o do usuário."""

    def __init__(self):
        self.app = None
        self.ja_aberto = False

    def obter(self):
        if self.app is None:
            try:
                import pythoncom
                import win32com.client
            except ImportError:
                raise RuntimeError(
                    "pywin32 não instalado. Rode: py -m pip install pywin32")
            pythoncom.CoInitialize()
            try:
                win32com.client.GetActiveObject("PowerPoint.Application")
                self.ja_aberto = True
            except Exception:
                self.ja_aberto = False
            try:
                self.app = win32com.client.DispatchEx("PowerPoint.Application")
            except Exception as e:
                raise RuntimeError(
                    f"não foi possível iniciar o PowerPoint ({e}). "
                    "Ele está instalado?")
            try:
                self.app.DisplayAlerts = PP_ALERTS_NONE
            except Exception:
                pass
        return self.app

    def fechar(self):
        if self.app is None:
            return
        try:
            if not self.ja_aberto and self.app.Presentations.Count == 0:
                self.app.Quit()
        except Exception:
            pass
        self.app = None


def remover_ocultos(pres, pdf_path):
    """Apaga do PDF as páginas dos slides ocultos (se ele tiver 1 por slide)."""
    total = pres.Slides.Count
    ocultos = [i - 1 for i in range(1, total + 1)
               if pres.Slides.Item(i).SlideShowTransition.Hidden == MSO_TRUE]
    if not ocultos:
        return
    doc = fitz.open(pdf_path)
    try:
        if doc.page_count != total:
            return  # o PowerPoint já os excluiu
        if len(ocultos) == total:
            raise ValueError("todos os slides estão ocultos")
        doc.delete_pages(ocultos)
        tmp = pdf_path + ".tmp.pdf"
        doc.save(tmp)
    finally:
        doc.close()
    os.replace(tmp, pdf_path)


def powerpoint_para_pdf(ppt, origem, pdf_tmp, pasta_tmp, ocultos=False):
    """Gera um PDF simples (1 slide/página). Retorna o método usado."""
    app = ppt.obter()
    # FileName, ReadOnly, Untitled, WithWindow
    pres = app.Presentations.Open(origem, MSO_TRUE, MSO_FALSE, MSO_FALSE)
    try:
        if pres.Slides.Count == 0:
            raise ValueError("a apresentação não tem slides")
        erros = []

        # 1) ExportAsFixedFormat (respeita a opção de slides ocultos)
        try:
            pres.ExportAsFixedFormat(
                pdf_tmp, PP_FIXED_PDF, PP_INTENT_PRINT, MSO_FALSE,
                1, 1, MSO_TRUE if ocultos else MSO_FALSE)
            if os.path.getsize(pdf_tmp) > 0:
                return "PowerPoint/export"
        except Exception as e:
            erros.append(f"export: {e}")

        # 2) Salvar como PDF (pode incluir ocultos: removemos depois)
        try:
            if ocultos:
                # só na cópia em memória; o arquivo original não é salvo
                for i in range(1, pres.Slides.Count + 1):
                    pres.Slides.Item(i).SlideShowTransition.Hidden = MSO_FALSE
            pres.SaveCopyAs(pdf_tmp, PP_SAVE_AS_PDF)
            if os.path.getsize(pdf_tmp) > 0:
                if not ocultos:
                    remover_ocultos(pres, pdf_tmp)
                return "PowerPoint/salvar"
        except Exception as e:
            erros.append(f"salvar: {e}")

        # 3) Último recurso: cada slide vira imagem PNG
        w = pres.PageSetup.SlideWidth
        h = pres.PageSetup.SlideHeight
        px_w = 2000
        px_h = int(px_w * h / w)
        doc = fitz.open()
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides.Item(i)
            if not ocultos and slide.SlideShowTransition.Hidden == MSO_TRUE:
                continue
            png = os.path.join(pasta_tmp, f"slide{i:04d}.png")
            slide.Export(png, "PNG", px_w, px_h)
            with fitz.open(png) as img:
                pdf_bytes = img.convert_to_pdf()
            with fitz.open("pdf", pdf_bytes) as pg:
                doc.insert_pdf(pg)
        if doc.page_count == 0:
            raise ValueError("todos os slides estão ocultos")
        doc.save(pdf_tmp)
        doc.close()
        print("      (aviso: exportação em PDF falhou; usando imagens: "
              + " | ".join(erros) + ")")
        return "PowerPoint/imagens"
    finally:
        try:
            pres.Saved = MSO_TRUE  # descarta qualquer alteração em memória
            pres.Close()
        except Exception:
            pass


# --------------------------------------------------------------------------
# Programa principal
# --------------------------------------------------------------------------

def coletar_arquivos(entradas):
    arquivos = []
    validos = EXT_POWERPOINT + EXT_PDF
    for e in entradas:
        e = os.path.abspath(e.strip('"'))
        if os.path.isdir(e):
            for f in sorted(os.listdir(e), key=str.lower):
                p = os.path.join(e, f)
                if (os.path.isfile(p) and f.lower().endswith(validos)
                        and not f.startswith("~$")):
                    arquivos.append(p)
        elif os.path.isfile(e):
            if e.lower().endswith(validos):
                arquivos.append(e)
            else:
                print(f"Ignorado (formato não suportado): {e}")
        else:
            print(f"Não encontrado: {e}")
    # remove duplicados mantendo a ordem
    vistos, unicos = set(), []
    for a in arquivos:
        k = os.path.normcase(a)
        if k not in vistos:
            vistos.add(k)
            unicos.append(a)
    return unicos


def proxima_pasta_saida(base):
    numeros = [int(m.group(1)) for d in os.listdir(base)
               if (m := re.fullmatch(r"slides(\d+)", d))
               and os.path.isdir(os.path.join(base, d))]
    return os.path.join(base, f"slides{max(numeros, default=0) + 1}")


def nome_livre(pasta, nome):
    destino = os.path.join(pasta, nome + ".pdf")
    k = 2
    while os.path.exists(destino):
        destino = os.path.join(pasta, f"{nome} ({k}).pdf")
        k += 1
    return destino


def converter(caminho, destino, args, ppt):
    titulo = os.path.basename(caminho)
    with tempfile.TemporaryDirectory() as tmp:
        if caminho.lower().endswith(EXT_PDF):
            fonte = fitz.open(caminho)
            if fonte.needs_pass:
                raise ValueError("PDF protegido por senha")
            metodo = "PDF"
        else:
            pdf_tmp = os.path.join(tmp, "slides.pdf")
            metodo = powerpoint_para_pdf(ppt, caminho, pdf_tmp, tmp,
                                         args.incluir_ocultos)
            fonte = fitz.open(pdf_tmp)
        try:
            n = montar_folheto(fonte, destino, args.por_pagina,
                               not args.sem_linhas, not args.sem_cabecalho,
                               titulo)
            return metodo, fonte.page_count, n
        finally:
            fonte.close()


def main():
    parser = argparse.ArgumentParser(
        description="Transforma apresentações e PDFs em folhetos A4.")
    parser.add_argument("entradas", nargs="*",
                        help="arquivos ou pastas (padrão: pasta do script)")
    parser.add_argument("-n", "--por-pagina", type=int, default=3,
                        choices=sorted(LAYOUTS),
                        help="slides por página (padrão: 3)")
    parser.add_argument("-o", "--saida",
                        help="pasta de saída (padrão: nova pasta slidesN)")
    parser.add_argument("--sem-linhas", action="store_true",
                        help="não desenhar linhas de anotação")
    parser.add_argument("--sem-cabecalho", action="store_true",
                        help="sem nome do arquivo e número de página")
    parser.add_argument("--incluir-ocultos", action="store_true",
                        help="incluir slides ocultos do PowerPoint")
    parser.add_argument("--sem-pausa", action="store_true",
                        help="não esperar Enter ao final")
    args = parser.parse_args()

    if fitz is None:
        print("Biblioteca PyMuPDF não instalada. Rode:\n"
              "    py -m pip install pymupdf pywin32")
        return 1

    entradas = args.entradas or [PASTA_SCRIPT]
    arquivos = coletar_arquivos(entradas)
    if not arquivos:
        print("Nenhuma apresentação ou PDF encontrado em: "
              + ", ".join(entradas))
        return 1

    if args.saida:
        saida = os.path.abspath(args.saida)
    else:
        primeira = entradas[0] if os.path.isdir(entradas[0]) \
            else os.path.dirname(os.path.abspath(entradas[0]))
        saida = proxima_pasta_saida(primeira)
    os.makedirs(saida, exist_ok=True)

    print(f"{len(arquivos)} arquivo(s) | {args.por_pagina} slides por página"
          f" | saída: {saida}\n")

    ppt = PowerPoint()
    ok = falhas = 0
    try:
        for i, caminho in enumerate(arquivos, 1):
            nome = os.path.splitext(os.path.basename(caminho))[0]
            destino = nome_livre(saida, nome)
            t0 = time.time()
            print(f"[{i}/{len(arquivos)}] {os.path.basename(caminho)} ...",
                  flush=True)
            try:
                metodo, slides, paginas = converter(caminho, destino, args, ppt)
                ok += 1
                print(f"      OK ({metodo}): {slides} slides -> {paginas} "
                      f"página(s) em {time.time() - t0:.1f}s")
            except Exception as e:
                falhas += 1
                if os.path.exists(destino):
                    try:
                        os.remove(destino)
                    except OSError:
                        pass
                print(f"      ERRO: {e}")
                if os.environ.get("FOLHETO_DEBUG"):
                    traceback.print_exc()
    finally:
        ppt.fechar()

    print(f"\nConcluído: {ok} OK, {falhas} com erro. PDFs em: {saida}")
    return 0 if falhas == 0 else 2


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    codigo = 1
    try:
        codigo = main()
    except KeyboardInterrupt:
        print("\nCancelado.")
    except Exception:
        traceback.print_exc()
    if "--sem-pausa" not in sys.argv and sys.stdin and sys.stdin.isatty():
        try:
            input("\nPressione Enter para fechar...")
        except EOFError:
            pass
    sys.exit(codigo)
