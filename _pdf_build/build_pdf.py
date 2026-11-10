# -*- coding: utf-8 -*-
"""Monta o PDF final juntando todas as seções."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from gerar_pdf import criar_doc, SAIDA  # noqa: E402
import conteudo as c1               # noqa: E402
import conteudo2 as c2              # noqa: E402
import conteudo3 as c3              # noqa: E402

from reportlab.platypus import Spacer  # noqa: E402


def construir():
    """Devolve a lista de Flowables na ordem final do documento."""
    historia = []
    historia += c1.capa()
    historia += c1.sumario()
    historia += c1.secao_problema()
    historia += c1.secao_arquitetura()
    historia += c2.secao_pipeline()
    historia += c2.secao_bibliotecas()
    historia += c2.secao_decisoes()
    historia += c3.secao_interface()
    historia += c3.secao_excel()
    historia += c3.secao_distribuicao()
    historia += c3.secao_desempenho()
    historia += c3.secao_limites()
    historia += c3.secao_glossario()
    return historia


def compactar(historia, fator: float):
    """Encolhe os Spacer verticais brandos por `fator` (0..1)."""
    for fl in historia:
        h = getattr(fl, "height", None)
        if isinstance(fl, Spacer) and h is not None and 0 < h <= 30:
            fl.height = h * fator
    return historia


def montar(fator: float = 1.0) -> Path:
    doc = criar_doc()
    historia = compactar(construir(), fator)
    doc.build(historia)
    _adicionar_bookmarks(SAIDA)
    return SAIDA


# Títulos das seções na ordem em que são montadas (para os bookmarks).
TITULOS = [
    "Capa", "Sumário",
    "01 · O problema",
    "02 · Arquitetura da solução",
    "03 · O pipeline de OCR em 6 etapas",
    "04 · Bibliotecas escolhidas e por quê",
    "05 · Decisões de engenharia",
    "06 · Interface e concorrência",
    "07 · Saída em Excel com openpyxl",
    "08 · Distribuição: um único .exe",
    "09 · Desempenho e resultados",
    "10 · Limitações e próximos passos",
    "11 · Glossário e consulta rápida",
]


def _adicionar_bookmarks(caminho: Path) -> None:
    """Cria o sumário navegável a partir dos banners de número de seção."""
    import re
    import pymupdf

    doc = pymupdf.open(str(caminho))
    toc = [[1, "Capa", 1], [1, "Sumário", 2]]
    # Cada banner de seção aparece no texto como "01", "02", ... isolado.
    for indice, pagina in enumerate(doc):
        if indice < 2:
            continue
        inicio = pagina.get_text("text", clip=pymupdf.Rect(0, 0, pagina.rect.width, 130))
        achado = re.search(r"\b(\d{2})\b", inicio)
        if achado:
            numero = int(achado.group(1))
            if 1 <= numero <= 11 and numero < len(TITULOS) - 1:
                titulo = TITULOS[numero + 1]
                if not any(t[1] == titulo for t in toc):
                    toc.append([1, titulo, indice + 1])
    doc.set_toc(toc)
    doc.saveIncr()
    doc.close()


if __name__ == "__main__":
    fator = float(sys.argv[1]) if len(sys.argv) > 1 else 1.0
    caminho = montar(fator)
    print(f"PDF gerado (fator {fator:.2f}): {caminho}")
