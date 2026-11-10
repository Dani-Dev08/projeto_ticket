# -*- coding: utf-8 -*-
"""Gera o PDF com ajuste automático de espaçamento.

Testa vários fatores de compactação de Spacer, mede cada resultado e escolhe
aquele com:
  1) nenhuma página quase vazia (pouco texto = sobra órfã de quebra de página);
  2) menos páginas no total.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import pymupdf                        # noqa: E402
import build_pdf                      # noqa: E402


# Abaixo disso a página é considerada "órfã" (só um resto de seção).
LIMITE_ORFA = 900
# Abaixo disso é capa ou página de fecho — legítima, não conta como órfã.
PAGINAS_ESPERADAS_CURTAS = {1}


def analisar(caminho: Path):
    doc = pymupdf.open(str(caminho))
    orfas, tamanhos = [], []
    for i, page in enumerate(doc):
        n = len(page.get_text().strip())
        tamanhos.append(n)
        if n < LIMITE_ORFA and (i + 1) not in PAGINAS_ESPERADAS_CURTAS:
            orfas.append((i + 1, n))
    total = doc.page_count
    doc.close()
    return total, orfas, tamanhos


def main() -> int:
    resultados = []
    for fator in [1.00, 0.92, 0.85, 0.78, 0.72, 0.66, 0.60]:
        try:
            caminho = build_pdf.montar(fator)
        except Exception as exc:                     # noqa: BLE001
            print(f"fator {fator:.2f}: FALHOU -> {exc}")
            continue
        total, orfas, _ = analisar(caminho)
        resultados.append((len(orfas), total, fator, orfas))
        print(f"fator {fator:.2f}: {total} páginas, órfãs: {orfas if orfas else 'nenhuma'}")

    if not resultados:
        print("Nenhuma configuração válida gerada.")
        return 1

    # Melhor = menos órfãs, depois menos páginas (não queremos esmagar o PDF).
    resultados.sort(key=lambda r: (r[0], r[1], -r[2]))
    orfas, total, fator, lista = resultados[0]

    caminho = build_pdf.montar(fator)
    print()
    print(f"ESCOLHIDO: fator {fator:.2f} -> {total} páginas, "
          f"{'sem órfãs' if not lista else 'órfãs: ' + str(lista)}")
    print(f"Arquivo: {caminho}")

    if lista:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
