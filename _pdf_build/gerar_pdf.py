# -*- coding: utf-8 -*-
"""
Gera o PDF de documentação técnica do projeto de extração de tickets de táxi.

Uso:  python _pdf_build/gerar_pdf.py
Saída: Documentacao_TicketsTaxis.pdf (raiz do projeto)
"""
from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle,
    Image, KeepTogether, PageBreak, Flowable,
)

RAIZ = Path(__file__).resolve().parent.parent
PASTA_IMG = Path(__file__).resolve().parent / "img"
SAIDA = RAIZ / "Documentacao_TicketsTaxis.pdf"

# --------------------------------------------------------------------------- #
# Fontes
# --------------------------------------------------------------------------- #
F = r"C:\Windows\Fonts"
pdfmetrics.registerFont(TTFont("App", f"{F}\\segoeui.ttf"))
pdfmetrics.registerFont(TTFont("AppBold", f"{F}\\segoeuib.ttf"))
pdfmetrics.registerFont(TTFont("AppSemi", f"{F}\\seguisb.ttf"))
pdfmetrics.registerFont(TTFont("Mono", f"{F}\\consola.ttf"))
pdfmetrics.registerFontFamily("App", normal="App", bold="AppBold",
                              italic="App", boldItalic="AppBold")

# --------------------------------------------------------------------------- #
# Paleta
# --------------------------------------------------------------------------- #
AZUL_ESCURO = colors.HexColor("#12233F")
AZUL = colors.HexColor("#1F4E79")
AZUL_MEDIO = colors.HexColor("#2E75B6")
AZUL_CLARO = colors.HexColor("#DEEBF7")
AZUL_PALIDO = colors.HexColor("#F2F7FC")
VERDE = colors.HexColor("#1E7A45")
VERDE_CLARO = colors.HexColor("#E2F2E8")
LARANJA = colors.HexColor("#C05A1B")
LARANJA_CLARO = colors.HexColor("#FCE9DD")
AMARELO_CLARO = colors.HexColor("#FFF6DD")
VERMELHO = colors.HexColor("#9C2B2B")
VERMELHO_CLARO = colors.HexColor("#FBE4E4")
CINZA = colors.HexColor("#5A6672")
CINZA_CLARO = colors.HexColor("#F1F3F5")
CINZA_BORDA = colors.HexColor("#C9D2DA")
BRANCO = colors.white
TEXTO = colors.HexColor("#22303C")

PAGINA = A4
MARGEM_L = 1.9 * cm
MARGEM_R = 1.9 * cm
MARGEM_T = 2.4 * cm
MARGEM_B = 1.8 * cm
LARGURA = PAGINA[0] - MARGEM_L - MARGEM_R

# --------------------------------------------------------------------------- #
# Estilos
# --------------------------------------------------------------------------- #
ss = getSampleStyleSheet()


def st(name, **kw):
    base = dict(fontName="App", fontSize=9.6, leading=14.2, textColor=TEXTO,
                alignment=TA_JUSTIFY, spaceAfter=5)
    base.update(kw)
    return ParagraphStyle(name, **base)


S_BODY = st("body")
S_BODY_SM = st("body_sm", fontSize=8.6, leading=12.4)
S_H2 = st("h2", fontName="AppBold", fontSize=12.4, leading=16, textColor=AZUL,
          alignment=TA_LEFT, spaceBefore=8, spaceAfter=4)
S_H3 = st("h3", fontName="AppSemi", fontSize=10.2, leading=13.5, textColor=AZUL_ESCURO,
          alignment=TA_LEFT, spaceBefore=6, spaceAfter=2)
S_CAP = st("cap", fontSize=7.8, leading=10.5, textColor=CINZA, alignment=TA_CENTER,
           spaceBefore=3, spaceAfter=2)
S_CELL = st("cell", fontSize=8.4, leading=11.2, alignment=TA_LEFT, spaceAfter=0)
S_CELL_B = st("cellb", fontName="AppSemi", fontSize=8.4, leading=11.2, alignment=TA_LEFT,
              spaceAfter=0)
S_CELL_C = st("cellc", fontSize=8.4, leading=11.2, alignment=TA_CENTER, spaceAfter=0)
S_CELL_H = st("cellh", fontName="AppBold", fontSize=8.4, leading=11.2, textColor=BRANCO,
              alignment=TA_LEFT, spaceAfter=0)
S_COD = st("cod", fontName="Mono", fontSize=7.2, leading=10.3, alignment=TA_LEFT,
           textColor=colors.HexColor("#D9E4EE"), backColor=colors.HexColor("#1E2B38"),
           borderPadding=(5, 6, 6, 4), spaceAfter=0)
S_LEGENDA_MONO = st("lm", fontName="Mono", fontSize=7.6, leading=10.5, textColor=AZUL_ESCURO)


def P(txt, style=S_BODY):
    return Paragraph(txt, style)


def B(txt):
    return f"<b>{txt}</b>"


def code(txt):
    return f'<font name="Mono" size="7.8" color="#1F4E79">{txt}</font>'


def esc(txt):
    return (txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# --------------------------------------------------------------------------- #
# Componentes visuais
# --------------------------------------------------------------------------- #

class Barra(Flowable):
    def __init__(self, largura, altura=2.2, cor=AZUL_MEDIO):
        super().__init__()
        self.largura, self.altura, self.cor = largura, altura, cor
        self.width, self.height = largura, altura

    def draw(self):
        c = self.canv
        c.setFillColor(self.cor)
        c.roundRect(0, 0, self.largura, self.altura, self.altura / 2, stroke=0, fill=1)


class TituloSecao(Flowable):
    def __init__(self, numero, titulo, largura=LARGURA):
        super().__init__()
        self.numero, self.titulo = numero, titulo
        self.width, self.height = largura, 30

    def draw(self):
        c = self.canv
        c.setFillColor(AZUL_ESCURO)
        c.roundRect(0, 0, self.width, self.height, 5, stroke=0, fill=1)
        c.setFillColor(AZUL_MEDIO)
        c.roundRect(0, 0, 44, self.height, 5, stroke=0, fill=1)
        c.setFillColor(BRANCO)
        c.setFont("AppBold", 16)
        c.drawCentredString(22, 9.5, f"{self.numero:02d}")
        c.setFont("AppBold", 12.6)
        c.drawString(56, 10.5, self.titulo.upper())


class Caixa(Flowable):
    def __init__(self, titulo, paragrafos, cor_fundo, cor_borda, largura=LARGURA):
        super().__init__()
        self.titulo = titulo
        self.paras = paragrafos
        self.cor_fundo, self.cor_borda = cor_fundo, cor_borda
        self.largura = largura
        self.pad = 9

    def wrap(self, aw, ah):
        self._aw = min(self.largura, aw)
        inner = self._aw - 2 * self.pad - 6
        self._h_titulo = 0
        self._titulo_para = None
        if self.titulo:
            p = Paragraph(self.titulo, st("ct", fontName="AppBold", fontSize=9.3,
                                          leading=12, textColor=self.cor_borda,
                                          alignment=TA_LEFT, spaceAfter=0))
            self._titulo_para = p
            self._h_titulo = p.wrap(inner, ah)[1] + 5
        self._alturas = []
        total = 0
        for p in self.paras:
            h = p.wrap(inner, ah)[1]
            self._alturas.append(h)
            total += h + 2.5
        self.height = total + self._h_titulo + 2 * self.pad - 2
        self.width = self._aw
        return self.width, self.height

    def draw(self):
        c = self.canv
        c.setFillColor(self.cor_fundo)
        c.setStrokeColor(self.cor_borda)
        c.setLineWidth(0.9)
        c.roundRect(0, 0, self.width, self.height, 5, stroke=1, fill=1)
        c.setFillColor(self.cor_borda)
        c.roundRect(0, 0, 3.6, self.height, 1.8, stroke=0, fill=1)
        y = self.height - self.pad
        x = self.pad + 7
        if self._titulo_para:
            y -= self._h_titulo
            self._titulo_para.drawOn(c, x, y + 3)
            y -= 5
        for p, h in zip(self.paras, self._alturas):
            y -= h
            p.drawOn(c, x, y)
            y -= 2.5


def caixa(titulo, texto, tipo="info"):
    mapa = {
        "info": (AZUL_PALIDO, AZUL_MEDIO),
        "ok": (VERDE_CLARO, VERDE),
        "warn": (AMARELO_CLARO, colors.HexColor("#A87500")),
        "erro": (VERMELHO_CLARO, VERMELHO),
        "dica": (LARANJA_CLARO, LARANJA),
    }
    fundo, borda = mapa[tipo]
    textos = texto if isinstance(texto, list) else [texto]
    paras = [Paragraph(t, S_BODY_SM) for t in textos]
    return Caixa(titulo, paras, fundo, borda)


def tabela(dados, larguras, cabecalho=True, zebra=True, alinhamentos=None):
    linhas = []
    for li, linha in enumerate(dados):
        nova = []
        for ci, cel in enumerate(linha):
            if isinstance(cel, str):
                if li == 0 and cabecalho:
                    nova.append(Paragraph(cel, S_CELL_H))
                else:
                    estilo = S_CELL
                    if alinhamentos and ci < len(alinhamentos):
                        al = alinhamentos[ci]
                        if al == "c":
                            estilo = S_CELL_C
                        elif al == "b":
                            estilo = S_CELL_B
                    nova.append(Paragraph(cel, estilo))
            else:
                nova.append(cel)
        linhas.append(nova)

    t = Table(linhas, colWidths=larguras, repeatRows=1 if cabecalho else 0)
    cmds = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 0), (-1, -1), 0.4, CINZA_BORDA),
        ("BOX", (0, 0), (-1, -1), 0.6, CINZA_BORDA),
    ]
    if cabecalho:
        cmds += [("BACKGROUND", (0, 0), (-1, 0), AZUL)]
    if zebra:
        for i in range(1, len(linhas)):
            if i % 2 == 0:
                cmds.append(("BACKGROUND", (0, i), (-1, i), AZUL_PALIDO))
    t.setStyle(TableStyle(cmds))
    return t


def bloco_codigo(titulo, codigo_txt):
    linhas = codigo_txt.strip("\n").split("\n")
    corpo = "<br/>".join(
        f'<font color="#D9E4EE">{esc(ln).replace(" ", "&nbsp;") or "&nbsp;"}</font>'
        for ln in linhas
    )
    partes = []
    if titulo:
        titulo_esc = esc(titulo)
        cab = Table(
            [[Paragraph(f'<font name="Mono" size="7.4" color="#7FB6E8">▍</font>'
                        f'<font name="AppSemi" size="8.2" color="#12233F"> {titulo_esc}</font>',
                        st("lc", alignment=TA_LEFT, spaceAfter=0))]],
            colWidths=[LARGURA - 4],
        )
        cab.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), AZUL_CLARO),
            ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#B7CFE6")),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ]))
        partes.append(cab)
        partes.append(Spacer(1, 2))
    partes.append(Paragraph(corpo, S_COD))
    return KeepTogether(partes)


def figura(nome, largura, legenda=""):
    caminho = PASTA_IMG / nome
    if not caminho.exists():
        return None
    from PIL import Image as PILImage
    with PILImage.open(caminho) as im:
        w, h = im.size
    altura = h * (largura / w)
    img = Image(str(caminho), width=largura, height=altura)
    if legenda:
        return KeepTogether([img, P(f"▲ {legenda}", S_CAP)])
    return img


def fluxo_etapas(etapas, largura_total=LARGURA, altura_caixa=26, altura_desc=34):
    """Diagrama horizontal de etapas com setas."""
    n = len(etapas)
    gap = 0.28 * cm
    largura_caixa = (largura_total - gap * (n - 1)) / n
    topo, desc = [], []
    for i, (num, tit) in enumerate(etapas):
        topo.append(Paragraph(
            f'<font color="white" size="8"><b>{num}</b></font>',
            st("fx", alignment=TA_CENTER, fontSize=8, textColor=BRANCO, spaceAfter=0)))
        desc.append(Paragraph(
            f'<font color="#12233F" size="7.6"><b>{tit}</b></font>',
            st("fx2", alignment=TA_CENTER, fontSize=7.6, textColor=AZUL_ESCURO,
               leading=9.6, spaceAfter=0)))
    t1 = Table([topo], colWidths=[largura_caixa] * n)
    estilo1 = [
        ("BACKGROUND", (0, 0), (-1, -1), AZUL_MEDIO),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("INNERGRID", (0, 0), (-1, -1), gap, BRANCO),
    ]
    for i in range(n):
        estilo1.append(("ROUNDEDCORNERS", [4, 4, 4, 4]) if False else ("BOX", (i, 0), (i, 0),
                                                                        0, BRANCO))
    t1.setStyle(TableStyle(estilo1))
    t2 = Table([desc], colWidths=[largura_caixa] * n)
    t2.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), AZUL_PALIDO),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("INNERGRID", (0, 0), (-1, -1), gap, BRANCO),
    ]))
    return KeepTogether([t1, t2])


# --------------------------------------------------------------------------- #
# Cabeçalho / rodapé
# --------------------------------------------------------------------------- #

def desenhar_capa(c, doc):
    c.saveState()
    c.setFillColor(AZUL_ESCURO)
    c.rect(0, 0, PAGINA[0], PAGINA[1], stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#183053"))
    c.circle(PAGINA[0] * 0.95, PAGINA[1] * 0.92, 155, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#152A48"))
    c.circle(PAGINA[0] * 0.05, PAGINA[1] * 0.13, 135, stroke=0, fill=1)
    c.setFillColor(AZUL_MEDIO)
    c.rect(0, PAGINA[1] - 0.5 * cm, PAGINA[0], 0.5 * cm, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#F0A02C"))
    c.rect(0, PAGINA[1] - 0.62 * cm, PAGINA[0], 0.12 * cm, stroke=0, fill=1)
    c.restoreState()


def desenhar_pagina(c, doc):
    c.saveState()
    c.setFillColor(AZUL_ESCURO)
    c.rect(0, PAGINA[1] - 1.2 * cm, PAGINA[0], 1.2 * cm, stroke=0, fill=1)
    c.setFillColor(AZUL_MEDIO)
    c.rect(0, PAGINA[1] - 1.27 * cm, PAGINA[0], 0.07 * cm, stroke=0, fill=1)
    c.setFillColor(BRANCO)
    c.setFont("AppBold", 8)
    c.drawString(MARGEM_L, PAGINA[1] - 0.8 * cm, "EXTRATOR DE TICKETS DE TÁXI")
    c.setFont("App", 8)
    c.setFillColor(colors.HexColor("#A9C6E2"))
    c.drawRightString(PAGINA[0] - MARGEM_R, PAGINA[1] - 0.8 * cm,
                      "Documentação técnica · OCR local")
    c.setStrokeColor(CINZA_BORDA)
    c.setLineWidth(0.5)
    c.line(MARGEM_L, 1.2 * cm, PAGINA[0] - MARGEM_R, 1.2 * cm)
    c.setFillColor(CINZA)
    c.setFont("App", 7.4)
    c.drawString(MARGEM_L, 0.82 * cm,
                 "projeto_ticket · Python 3.13 · OpenCV · Tesseract 5 · Tkinter · PyInstaller")
    c.drawRightString(PAGINA[0] - MARGEM_R, 0.82 * cm, f"pág. {doc.page - 1}")
    c.restoreState()


def criar_doc():
    doc = BaseDocTemplate(
        str(SAIDA), pagesize=PAGINA,
        leftMargin=MARGEM_L, rightMargin=MARGEM_R,
        topMargin=MARGEM_T, bottomMargin=MARGEM_B,
        title="Extrator de Tickets de Táxi — Documentação Técnica",
        author="projeto_ticket", subject="OCR de tickets com Python",
    )
    frame_capa = Frame(MARGEM_L, MARGEM_B, LARGURA,
                       PAGINA[1] - MARGEM_T - MARGEM_B, id="capa")
    frame = Frame(MARGEM_L, MARGEM_B, LARGURA,
                  PAGINA[1] - MARGEM_T - MARGEM_B, id="normal")
    doc.addPageTemplates([
        PageTemplate(id="capa", frames=[frame_capa], onPage=desenhar_capa),
        PageTemplate(id="normal", frames=[frame], onPage=desenhar_pagina),
    ])
    return doc
