#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Extrator de Nº do TICKET e VALOR de tickets de táxi.

Lê todas as imagens da pasta `img_ticket/` e imprime no terminal o número do
ticket e o valor total. Usa OCR 100% local (Tesseract 5) com pré-processamento
OpenCV em várias binarizações e votação por linha, para lidar tanto com tickets
impressos nítidos quanto com tickets de texto cinza claro (pouca tinta).

Uso:
    python extrair_tickets.py                    # lê ./img_ticket
    python extrair_tickets.py <pasta>            # lê outra pasta
    python extrair_tickets.py <pasta> --json     # também gera tickets.json
    python extrair_tickets.py --debug            # mostra as linhas reconhecidas

Requisitos: tesseract (com por.traineddata), opencv-python, pytesseract.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import statistics
import subprocess
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Callable, Iterable, Sequence

import cv2
import numpy as np
import pytesseract

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# Configuração inicial do Tesseract (antes de qualquer uso)
# --------------------------------------------------------------------------- #

def _pasta_base_aplicacao() -> Path:
    """
    Pasta onde o programa está rodando.

    - Script normal: pasta do arquivo .py
    - Executável (PyInstaller onefile): pasta onde o .exe está, não a pasta
      temporária _MEIPASS, para que o usuário possa trocar a pasta do
      Tesseract ao lado do .exe.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _pasta_recursos() -> Path:
    """Pasta de recursos embutidos (PyInstaller _MEIPASS) ou a pasta do app."""
    meipass = getattr(sys, "_MEIPASS", None)
    return Path(meipass) if meipass else _pasta_base_aplicacao()


def configurar_tesseract() -> None:
    """Garante que o executável do Tesseract e o TESSDATA_PREFIX estejam corretos."""
    base = _pasta_base_aplicacao()
    recursos = _pasta_recursos()

    candidatos_exe = [
        # Tesseract embutido no .exe / ao lado dele (portátil).
        base / "Tesseract-OCR" / "tesseract.exe",
        base / "tesseract" / "tesseract.exe",
        base / "tesseract.exe",
        recursos / "Tesseract-OCR" / "tesseract.exe",
        # Instalação padrão do sistema.
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    ]
    exe = os.environ.get("TESSERACT_EXE")
    if not exe or not Path(exe).is_file():
        exe = next((str(c) for c in candidatos_exe if c.is_file()), shutil.which("tesseract"))

    if not exe:
        return

    pytesseract.pytesseract.tesseract_cmd = exe
    pasta_exe = Path(exe).parent

    # tessdata: prefere a pasta ao lado do tesseract; no Anaconda fica em
    # Library/share/tessdata.
    candidatos_tessdata = [
        pasta_exe / "tessdata",
        pasta_exe.parent / "share" / "tessdata",
        base / "Tesseract-OCR" / "tessdata",
        recursos / "Tesseract-OCR" / "tessdata",
        Path(r"C:\Program Files\Tesseract-OCR\tessdata"),
    ]
    for tessdata in candidatos_tessdata:
        if tessdata.is_dir() and any(tessdata.glob("*.traineddata")):
            os.environ["TESSDATA_PREFIX"] = str(tessdata)
            break

configurar_tesseract()

# No .exe (modo janela, sem console) cada chamada ao tesseract.exe abriria uma
# janela do CMD piscando. STARTUPINFO + CREATE_NO_WINDOW impede isso.
_SEM_JANELA = None
if os.name == "nt":
    _si = subprocess.STARTUPINFO()
    _si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    _si.wShowWindow = subprocess.SW_HIDE
    _SEM_JANELA = {
        "startupinfo": _si,
        "creationflags": subprocess.CREATE_NO_WINDOW,
    }


def _tesseract_sem_janela(*args, **kwargs):
    """Substitui subprocess.Popen para nunca exibir janela do console."""
    if _SEM_JANELA:
        for chave, valor in _SEM_JANELA.items():
            kwargs.setdefault(chave, valor)
    return _POPEN_ORIGINAL(*args, **kwargs)


# O pytesseract chama o tesseract por subprocess.Popen em toda leitura: sem
# isso reaparece uma janela preta para cada ticket processado.
# (guardamos a referência original: reatribuir o atributo no módulo faria a
# função chamar a si mesma)
_POPEN_ORIGINAL = subprocess.Popen
pytesseract.pytesseract.subprocess.Popen = _tesseract_sem_janela


def idioma_disponivel() -> str:
    """Retorna o idioma preferido ('por') se disponível, senão 'eng' ou outro."""
    try:
        idiomas = set(pytesseract.get_languages(config=""))
        if LANG_PREFERIDA in idiomas:
            return LANG_PREFERIDA
        if "eng" in idiomas:
            return "eng"
        return list(idiomas)[0] if idiomas else LANG_PREFERIDA
    except Exception:
        return LANG_PREFERIDA


# --------------------------------------------------------------------------- #
# Configuração
# --------------------------------------------------------------------------- #

EXTENSOES = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"}

LANG_PREFERIDA = "por"

# Lado máximo (px) da imagem enviada ao Tesseract. Acima disso o OCR só fica
# mais lento, sem ganho de precisão: 1700px já deixa o texto com altura de sobra.
LADO_MAXIMO_OCR = 1700

# Rótulos conhecidos (o OCR costuma estragar as primeiras letras por causa
# dos ícones/figuras que ficam à esquerda do texto).
ROTULOS = ("ped", "bagag", "valor", "total")

# Linha "VALOR" (valor da corrida) — o OCR às vezes perde letras isoladas.
# Cobre "VALOR", "VALOR DA CORRIDA", "VALOR":
VALOR_TXT = re.compile(r"v[a4][l1i][o0]r")
CORRIDA_TXT = re.compile(r"c[o0]rr[i1l]da|c[o0]re[i1l]da|bande[i1l]ra")

# Linha "VALOR TOTAL A PAGAR" / "TOTAL A PAGAR" / "TOTAL".
# Este teste deve vir ANTES do de "VALOR", pois "VALOR TOTAL" também casa
# com VALOR_TXT.
TOTAL_TXT = re.compile(r"t[o0]t[ai1l]")

# Padrões para descartar datas/horas quando procurando valores monetários.
RE_DATA = re.compile(r"\b\d{2}[./]\d{2}[./]\d{2,4}\b")
RE_HORA = re.compile(r"\b\d{1,2}[:.]\d{2}(?:[:.]\d{2})?\b")

# Confiança mínima do OCR para aceitar o Nº do ticket / o valor principal.
# Abaixo disso a leitura é tratada como ilegível ("Ñ IDENTIFICADO").
MIN_CONF_NUMERO = 55.0
MIN_CONF_VALOR = 45.0

# Piso de valor: ticket de táxi nunca custa menos que isto. Leitura abaixo é
# tratada como erro de OCR -> "Ñ IDENTIFICADO".
MIN_VALOR_TICKET = 20.0

# Leituras que, na prática, são ano/data vazando para o campo do Nº.
ANOS_SUSPEITOS = {"2020", "2021", "2022", "2023", "2024", "2025", "2026", "2027",
                  "2028", "2029", "2030"}

# Marcador de leitura não confiável.
NAO_IDENTIFICADO = "Ñ IDENTIFICADO"


@dataclass
class Item:
    """Uma linha de valor do ticket (pedágio, bagagem, corrida, total...)."""

    rotulo: str
    texto_ocr: str
    valor: float | None
    confianca: float
    y: int
    # Recorte da região do valor, para revalidação em caso de conflito.
    recorte: np.ndarray | None = None
    caixa: tuple[int, int, int, int] | None = None


@dataclass
class Ticket:
    arquivo: str
    numero: str | None = None
    valor: float | None = None
    valor_corrida: float | None = None
    pedagio: float | None = None
    bagagem: float | None = None
    total: float | None = None
    valor_pagamento: float | None = None
    data_hora: str | None = None
    forma_pagamento: str | None = None
    origem: str = ""
    confianca: str = "baixa"
    observacoes: list[str] = field(default_factory=list)

    def como_dicionario(self) -> dict:
        d = asdict(self)
        d["valor"] = None if self.valor is None else round(self.valor, 2)
        d["valor_formatado"] = NAO_IDENTIFICADO if self.valor is None else moeda(self.valor)
        d["numero"] = self.numero or NAO_IDENTIFICADO
        return d


# --------------------------------------------------------------------------- #
# Pré-processamento e OCR
# --------------------------------------------------------------------------- #

def carregar_imagem(caminho: Path) -> np.ndarray:
    """Lê a imagem em tons de cinza (via imdecode, aceita caminhos acentuados)."""
    dados = np.fromfile(str(caminho), dtype=np.uint8)
    img = cv2.imdecode(dados, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f"Não foi possível abrir a imagem: {caminho.name}")
    return cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)


def _escala_para_texto(cinza: np.ndarray, altura_alvo: float = 60.0) -> float:
    """Estima a ampliação necessária para o texto ficar com ~altura_alvo px."""
    binaria = cv2.threshold(cinza, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    _, _, stats, _ = cv2.connectedComponentsWithStats(binaria, 8)
    alturas = [s[3] for s in stats[1:] if 6 <= s[3] <= cinza.shape[0] / 6 and s[4] > 8]
    altura_texto = statistics.median(alturas) if alturas else 18.0
    return float(np.clip(altura_alvo / max(altura_texto, 1.0), 1.0, 6.0))


def binarizacoes(cinza: np.ndarray) -> list[tuple[str, np.ndarray, float]]:
    """
    Devolve variantes (nome, imagem_binária, escala) da imagem inteira.

    - `otsu`:       texto escuro normal (tickets nítidos).
    - `open`:       Otsu + abertura: limpa/regulariza dígitos degradados.
    - `soft`:       preserva texto cinza claro (pouca tinta).
    - `clahe_open`: contraste local para texto muito apagado.

    A ampliação é limitada por LADO_MAXIMO_OCR: acima disso o Tesseract só fica
    mais lento sem ganho de acurácia (o texto já tem altura suficiente).
    """
    escala = _escala_para_texto(cinza)
    altura, largura = cinza.shape[:2]
    escala = min(escala, LADO_MAXIMO_OCR / max(altura, largura))

    if escala > 1.05:
        ampliada = cv2.resize(cinza, None, fx=escala, fy=escala, interpolation=cv2.INTER_CUBIC)
    else:
        ampliada = cinza.copy()

    variantes: list[tuple[str, np.ndarray, float]] = []

    otsu = cv2.threshold(ampliada, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    variantes.append(("otsu", otsu, escala))

    kernel = np.ones((2, 2), np.uint8)
    variantes.append(("open", cv2.morphologyEx(otsu, cv2.MORPH_OPEN, kernel), escala))

    borrado = cv2.GaussianBlur(ampliada, (0, 0), 1.2)
    variantes.append(("soft", cv2.threshold(borrado, 205, 255, cv2.THRESH_BINARY)[1], escala))

    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(ampliada)
    clahe = cv2.GaussianBlur(clahe, (3, 3), 0)
    adapt = cv2.adaptiveThreshold(
        clahe, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 55, 18
    )
    variantes.append(("clahe_open", cv2.morphologyEx(adapt, cv2.MORPH_OPEN, kernel), escala))

    return variantes


def _iou(a: dict, b: dict) -> float:
    """Interseção sobre união de duas caixas."""
    x1, y1 = max(a["x"], b["x"]), max(a["y"], b["y"])
    x2 = min(a["x"] + a["w"], b["x"] + b["w"])
    y2 = min(a["y"] + a["h"], b["y"] + b["h"])
    if x2 <= x1 or y2 <= y1:
        return 0.0
    inter = (x2 - x1) * (y2 - y1)
    uniao = a["w"] * a["h"] + b["w"] * b["h"] - inter
    return inter / uniao if uniao else 0.0


def _mesmo_simbolo(a: dict, b: dict, limite: float = 0.35) -> bool:
    """Duas leituras representam o mesmo símbolo quando as caixas coincidem."""
    if _iou(a, b) >= limite:
        return True
    # Com dígitos iguais, tolera pequenos deslocamentos causados pela binarização.
    da, db = re.sub(r"\D", "", a["texto"]), re.sub(r"\D", "", b["texto"])
    if not da or not db:
        return False
    cx_a, cx_b = a["x"] + a["w"] / 2, b["x"] + b["w"] / 2
    cy_a, cy_b = a["y"] + a["h"] / 2, b["y"] + b["h"] / 2
    proximo = (abs(cx_a - cx_b) < max(a["w"], b["w"]) * 0.7 and
               abs(cy_a - cy_b) < max(a["h"], b["h"]) * 0.7)
    return proximo and (da == db or da in db or db in da)


def ocr_palavras(cinza: np.ndarray, idioma: str) -> list[dict]:
    """Roda o OCR em todas as binarizações e devolve uma única leitura por símbolo."""
    candidatos: list[dict] = []

    for nome, img, escala in binarizacoes(cinza):
        dados = pytesseract.image_to_data(
            img,
            lang=idioma,
            config="--psm 6 -c preserve_interword_spaces=1",
            output_type=pytesseract.Output.DICT,
        )
        for i, bruto in enumerate(dados["text"]):
            texto = (bruto or "").strip()
            conf = float(dados["conf"][i])
            if not texto or conf < 20:
                continue
            candidatos.append({
                "texto": texto,
                "x": int(dados["left"][i] / escala),
                "y": int(dados["top"][i] / escala),
                "w": max(1, int(dados["width"][i] / escala)),
                "h": max(1, int(dados["height"][i] / escala)),
                "conf": conf,
                "fonte": nome,
            })

    # Melhores leituras primeiro; descarta as que repetem um símbolo já aceito.
    candidatos.sort(key=lambda p: -p["conf"])
    aceitas: list[dict] = []
    for palavra in candidatos:
        if any(_mesmo_simbolo(palavra, outra) for outra in aceitas):
            continue
        aceitas.append(palavra)

    return aceitas


def agrupar_linhas(palavras: Sequence[dict], tolerancia: float = 0.5) -> list[list[dict]]:
    """
    Agrupa palavras em linhas usando o centro vertical das caixas.

    A tolerância usa a *altura mediana* do documento (não a maior caixa da
    linha) para não fundir a linha de "VALOR DA CORRIDA" com a de "VALOR TOTAL"
    quando algum token sai com caixa exagerada.
    """
    if not palavras:
        return []

    altura_mediana = statistics.median(p["h"] for p in palavras)
    ordenadas = sorted(palavras, key=lambda p: (p["y"] + p["h"] / 2, p["x"]))

    linhas: list[list[dict]] = []
    for palavra in ordenadas:
        centro = palavra["y"] + palavra["h"] / 2
        if linhas:
            ultima = linhas[-1]
            centro_ultima = statistics.mean(p["y"] + p["h"] / 2 for p in ultima)
            if abs(centro - centro_ultima) <= altura_mediana * tolerancia:
                ultima.append(palavra)
                continue
        linhas.append([palavra])
    return [sorted(linha, key=lambda p: p["x"]) for linha in linhas]


def normalizar_numero(texto: str) -> str | None:
    """Corrige confusões típicas do OCR: O->0, l/I/|->1, S->5, B->8."""
    tabela = str.maketrans({
        "O": "0", "o": "0", "D": "0", "Q": "0",
        "l": "1", "I": "1", "|": "1", "i": "1", "!": "1",
        "S": "5", "s": "5", "B": "8", "Z": "2", "z": "2",
        "G": "6", "T": "7",
    })
    limpo = re.sub(r"[^\dOoDlI|iSsBZzQG!]", "", texto).translate(tabela)
    if not 3 <= len(limpo) <= 12:
        return None
    return limpo or None


def reler_valor(recorte: np.ndarray) -> float | None:
    """
    Relê o recorte de um valor com limiares variados e devolve o valor votado.

    Usado quando a leitura da linha conflita com as demais fontes (ex.: total
    não bate com corrida+extras). A votação entre limiares evita confiar em
    uma única leitura na qual um '2' degradado foi lido como '7'.
    """
    if recorte is None or recorte.size == 0:
        return None

    votos: dict[float, float] = {}
    for fx in (6.0, 10.0):
        ampliado = cv2.resize(recorte, None, fx=fx, fy=fx, interpolation=cv2.INTER_CUBIC)
        suave = cv2.GaussianBlur(ampliado, (0, 0), 1.2)
        for limite in (175, 190, 205, 220, 235):
            binaria = cv2.threshold(suave, limite, 255, cv2.THRESH_BINARY)[1]
            binaria = cv2.copyMakeBorder(binaria, 30, 30, 40, 40,
                                         cv2.BORDER_CONSTANT, value=255)
            for psm in (7, 8):
                texto = pytesseract.image_to_string(binaria, config=f"--psm {psm}").strip()
                valor = normalizar_valor(texto)
                if valor is not None and 0.0 <= valor <= 100_000:
                    votos[round(valor, 2)] = votos.get(round(valor, 2), 0.0) + 1

    if not votos:
        return None
    return max(votos.items(), key=lambda kv: kv[1])[0]


def normalizar_valor(texto: str) -> float | None:
    """Extrai o último valor monetário do texto ('R$ 1.234,56' -> 1234.56)."""
    candidatos = list(re.finditer(r"(\d[\d.\s]*)[.,:]?\s*(\d{2})(?!\d)", texto))
    if not candidatos:
        return None
    achado = candidatos[-1]
    inteiro = re.sub(r"\D", "", achado.group(1))
    if not inteiro:
        return None
    try:
        return float(f"{int(inteiro)}.{achado.group(2)}")
    except ValueError:
        return None


def _distancia(a: str, b: str) -> int:
    """Distância de Levenshtein (usada para tolerar erros do OCR nos rótulos)."""
    if a == b:
        return 0
    if not a or not b:
        return max(len(a), len(b))
    anterior = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        atual = [i]
        for j, cb in enumerate(b, 1):
            atual.append(min(
                anterior[j] + 1,
                atual[j - 1] + 1,
                anterior[j - 1] + (ca != cb),
            ))
        anterior = atual
    return anterior[-1]


def rotulo_da_linha(linha: Sequence[dict]) -> str:
    """Classifica a linha pelo rótulo que aparece à esquerda do valor."""
    texto = " ".join(p["texto"] for p in linha).lower()
    texto = re.sub(r"[^\wÀ-ÿ\s$.,:>|\-]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()

    # Ordem importa: "VALOR TOTAL A PAGAR" contém "VALOR", então o teste de
    # TOTAL precisa vir primeiro para não ser classificado como "valor".
    if TOTAL_TXT.search(texto):
        return "total"
    if VALOR_TXT.search(texto) or CORRIDA_TXT.search(texto):
        return "valor"
    # Compara com os rótulos tolerando 1-2 caracteres errados pelo OCR.
    for palavra in re.findall(r"[a-zà-ÿ]+", texto):
        for chave in ROTULOS:
            if palavra.startswith(chave) or (len(palavra) >= 4 and _distancia(palavra[:6], chave) <= 1):
                return chave
    return "?"


# --------------------------------------------------------------------------- #
# Extração
# --------------------------------------------------------------------------- #

def extrair_numero(linhas: Sequence[Sequence[dict]], largura: int) -> tuple[str | None, float]:
    """
    Nº do ticket fica na linha do rótulo 'Nº TICKET', na coluna da direita.

    Se o rótulo 'TICKET' sair ilegível ('TICKCI', 'TICKEI'), identifica a linha
    pela estrutura: um único número longo isolado na metade direita, numa
    faixa que contém 'Nº' à esquerda.

    O número é apenas informativo: quando a leitura é duvidosa (baixa confiança
    do OCR, poucos dígitos, ou leitura de um possível ano/data), devolvemos
    `None` em vez de "adivinhar" um número — evita inventar dados.
    """
    candidatos: list[tuple[int, str, float]] = []  # (prioridade, número, conf)

    # Contextos que indicam número de cartão/autorização/CNPJ, não o Nº do ticket.
    RE_CONTEXTO_RUIDO = re.compile(
        r"CREDITO|DEBITO|MASTERCARD|VISA|ELO|CARD|CARTAO|"
        r"AEROCOOP|ELECTRON|VISTA|AUTORIZ|PAN|"
        r"\d{2}/\d{2}"  # data
    )

    for linha in linhas:
        texto = " ".join(p["texto"] for p in linha).upper()
        tem_rotulo = "TICKET" in texto or bool(
            re.search(r"T[I1|]?C[KX][CE][IT]", texto)
        ) or bool(re.search(r"N[º°oª]", texto))

        # Sem o rótulo da linha do Nº, só aceitamos números de uma linha que
        # claramente não é de cartão/pagamento/CNPJ — evita capturar o PAN.
        linha_ruidosa = bool(RE_CONTEXTO_RUIDO.search(texto))
        if not tem_rotulo and linha_ruidosa:
            continue

        coluna = [p for p in linha if p["x"] > largura * 0.45]
        if not coluna:
            continue
        for palavra in sorted(coluna, key=lambda p: -(p["x"] + p["w"])):
            if not re.fullmatch(r"[\d\s]{4,12}", palavra["texto"]):
                continue
            # Confiança baixa indica dígitos apagados: não confiar no número.
            if palavra["conf"] < MIN_CONF_NUMERO:
                continue
            numero = normalizar_numero(palavra["texto"])
            if not numero or not 5 <= len(numero) <= 10:
                continue
            # Descarta leituras que parecem ano/data (ex.: 2020, 2026, 1610...).
            if numero in ANOS_SUSPEITOS or numero[:4] in ANOS_SUSPEITOS:
                continue
            prioridade = 2 if tem_rotulo else 1
            candidatos.append((prioridade, numero, palavra["conf"]))
            break

    # Só confia em número sem rótulo se não houver nenhum candidato com rótulo
    # E houver confiança razoável — caso contrário, prefere "Ñ IDENTIFICADO".
    if candidatos:
        com_rotulo = [c for c in candidatos if c[0] == 2]
        if com_rotulo:
            candidatos = com_rotulo
        else:
            candidatos = [c for c in candidatos if c[2] >= MIN_CONF_NUMERO + 10]

    if not candidatos:
        return None, 0.0
    candidatos.sort(key=lambda c: (c[0], c[2]), reverse=True)
    return candidatos[0][1], candidatos[0][2]


def _melhor_valor_da_coluna(coluna: Sequence[dict]) -> tuple[float | None, float]:
    """
    Escolhe a leitura monetária mais plausível dentro da coluna da direita.

    Tokens com `,`/`.` decimal explícito têm prioridade sobre sequências de
    dígitos que, sem o separador, poderiam formar um valor absurdo
    (ex.: '124,00' + '144' -> 112500.00).
    """
    opcoes: list[tuple[int, float, float]] = []  # (prioridade, valor, confiança)

    for palavra in coluna:
        if not re.search(r"\d", palavra["texto"]):
            continue
        if RE_DATA.search(palavra["texto"]) or RE_HORA.search(palavra["texto"]):
            continue
        valor = normalizar_valor(palavra["texto"])
        if valor is None:
            continue
        explicito = bool(re.search(r"(\d)[.,](\d{2})$", palavra["texto"]))
        prioridade = 2 if explicito else 1
        opcoes.append((prioridade, valor, palavra["conf"]))

    # Também tenta a concatenação dos dígitos na ordem (ex.: '144' + '00').
    juntos = re.sub(r"[^\d.,]", "", "".join(p["texto"] for p in coluna))
    if juntos:
        valor = normalizar_valor(juntos)
        if valor is not None and re.search(r"(\d)[.,](\d{2})$", juntos):
            opcoes.append((1, valor, statistics.mean(p["conf"] for p in coluna)))

    if not opcoes:
        return None, 0.0
    opcoes.sort(key=lambda o: (o[0], o[2]), reverse=True)
    return opcoes[0][1], opcoes[0][2]


def extrair_valores(linhas: Sequence[Sequence[dict]], cinza: np.ndarray) -> list[Item]:
    """Lê o valor monetário da coluna da direita de cada linha rotulada."""
    itens: list[Item] = []
    largura = cinza.shape[1]
    limite_esq = largura * 0.55

    for linha in linhas:
        if not linha:
            continue
        coluna = [p for p in linha if p["x"] >= limite_esq]
        if not coluna:
            continue

        rotulo = rotulo_da_linha(linha)
        if rotulo == "?":
            continue

        # Ordena por borda direita: leituras coladas antes se anexam ao final do
        # valor ('124,00' termina em 672, '144' em 638), então o sufixo vem antes.
        coluna = sorted(coluna, key=lambda p: p["x"] + p["w"])

        valor, conf = _melhor_valor_da_coluna(coluna)
        if valor is None or not 0.0 <= valor <= 100_000:
            continue

        # Recorta a região do valor (com folga) para eventual revalidação.
        x0 = max(0, min(p["x"] for p in coluna) - 12)
        y0 = max(0, min(p["y"] for p in coluna) - 6)
        x1 = min(largura, max(p["x"] + p["w"] for p in coluna) + 8)
        y1 = min(cinza.shape[0], max(p["y"] + p["h"] for p in coluna) + 6)
        recorte = cinza[y0:y1, x0:x1].copy()

        itens.append(
            Item(
                rotulo=rotulo,
                texto_ocr=" ".join(p["texto"] for p in linha),
                valor=valor,
                confianca=conf,
                y=linha[0]["y"],
                recorte=recorte,
                caixa=(x0, y0, x1, y1),
            )
        )

    return itens


def extrair_data_hora(linhas: Sequence[Sequence[dict]]) -> str | None:
    padrao = re.compile(r"(\d{2}[./]\d{2}[./]\d{4}).*?(\d{1,2}[:.]\d{2}[:.]\d{2})")
    for linha in linhas:
        achado = padrao.search(" ".join(p["texto"] for p in linha))
        if achado:
            return f"{achado.group(1).replace('.', '/')} {achado.group(2).replace('.', ':')}"
    return None


def extrair_forma_pagamento(linhas: Sequence[Sequence[dict]]) -> str | None:
    for linha in linhas:
        texto = " ".join(p["texto"] for p in linha).upper()
        if re.search(r"DEB[I1l]T", texto):
            return "Cartão de débito"
        if re.search(r"CR[E3]D[I1l]T", texto):
            return "Cartão de crédito"
        if re.search(r"ESPECIE|DINHEIRO", texto):
            return "Dinheiro"
        if "PIX" in texto:
            return "PIX"
    return None


def extrair_origem(linhas: Sequence[Sequence[dict]]) -> str:
    """Recupera origem/destino (opcional, útil para auditoria)."""
    partes = []
    for linha in linhas:
        texto = " ".join(p["texto"] for p in linha).upper()
        if re.search(r"AEROPORTO|COPACABANA|JANEIRO|DESTINO", texto):
            partes.append(re.sub(r"\s+", " ", texto))
    return " | ".join(partes[:3])


def interpretar(arquivo: str, linhas: Sequence[Sequence[dict]], cinza: np.ndarray) -> Ticket:
    ticket = Ticket(arquivo=arquivo)
    itens = extrair_valores(linhas, cinza)

    ticket.numero, conf_numero = extrair_numero(linhas, cinza.shape[1])
    ticket.data_hora = extrair_data_hora(linhas)
    ticket.forma_pagamento = extrair_forma_pagamento(linhas)
    ticket.origem = extrair_origem(linhas)

    if not itens:
        ticket.observacoes.append("Nenhuma linha de valor reconhecida.")
        return ticket

    # Separa a seção da corrida da seção de pagamento.
    y_pagamento = 0
    for linha in linhas:
        if "PAGAMENTO" in " ".join(p["texto"] for p in linha).upper():
            y_pagamento = linha[0]["y"]
            break

    corrida = [i for i in itens if y_pagamento == 0 or i.y < y_pagamento]
    pagamento = [i for i in itens if y_pagamento and i.y > y_pagamento]

    def primeiro(lista: Iterable[Item], chave: str) -> Item | None:
        return next((i for i in lista if i.rotulo == chave), None)

    pedagio = primeiro(corrida, "ped")
    bagagem = primeiro(corrida, "bagag")
    totais = [i for i in corrida if i.rotulo == "total"]
    valores = [i for i in corrida if i.rotulo == "valor"]
    valor_pag = pagamento[-1] if pagamento else None

    # O ticket tem duas vias: a do passageiro (VALOR DA CORRIDA + VALOR TOTAL
    # A PAGAR) e a da empresa, que costuma repetir "VALOR R$ ..." mais abaixo.
    # O "VALOR" que interessa é o que precede o "VALOR TOTAL A PAGAR".
    y_primeiro_total = totais[0].y if totais else float("inf")
    valores_antes_total = [i for i in valores if i.y < y_primeiro_total]
    valor_corrida_item = (valores_antes_total or valores or [None])[0]

    ticket.pedagio = pedagio.valor if pedagio else None
    ticket.bagagem = bagagem.valor if bagagem else None
    ticket.valor_pagamento = valor_pag.valor if valor_pag else None
    ticket.valor_corrida = valor_corrida_item.valor if valor_corrida_item else None
    ticket.total = totais[0].valor if totais else None

    # ------------------------------------------------------------------ #
    # Releitura do VALOR quando ele aparenta leitura ruim.
    # Em tickets térmicos apagados o "VALOR DA CORRIDA" pode sair truncado
    # (ex.: '4U,00' -> 14,00). Se o TOTAL estiver confirmado pelo valor pago,
    # tentamos reler o recorte do VALOR antes de decidir pela prioridade.
    # ------------------------------------------------------------------ #
    if (
        valor_corrida_item is not None
        and ticket.total is not None
        and ticket.valor_pagamento is not None
        and abs(ticket.total - ticket.valor_pagamento) <= 0.01
        and abs(ticket.valor_corrida - ticket.total) > 0.01
    ):
        # O total "oficial" é total - extras; usa-o como alvo da releitura.
        alvo = round(ticket.total - (ticket.bagagem or 0.0) - (ticket.pedagio or 0.0), 2)
        releitura = reler_valor(valor_corrida_item.recorte)
        if (
            releitura is not None
            and abs(releitura - alvo) < 0.01
            and abs(releitura - ticket.valor_corrida) > 0.01
        ):
            ticket.observacoes.append(
                f"VALOR relido por recorte: {ticket.valor_corrida:.2f} -> {releitura:.2f}"
            )
            ticket.valor_corrida = releitura
            valor_corrida_item.valor = releitura

    # ------------------------------------------------------------------ #
    # Revalidação: se o total não fecha com corrida+extras, relê o recorte.
    # Só aceita a releitura se ela for consistente com o contexto da linha
    # (mesmos dígitos finais) — descarta leituras truncadas como '4,00'.
    # ------------------------------------------------------------------ #
    if totais and ticket.valor_corrida is not None:
        esperado = round(
            ticket.valor_corrida + (ticket.bagagem or 0.0) + (ticket.pedagio or 0.0), 2
        )
        if abs(esperado - ticket.total) > 0.01:
            releitura = reler_valor(totais[0].recorte)
            if releitura is not None and abs(releitura - esperado) < 0.01:
                ticket.observacoes.append(
                    f"Total relido por recorte: {ticket.total:.2f} -> {releitura:.2f}"
                )
                ticket.total = releitura
                totais[0].valor = releitura

    # Revalida o valor pago apenas se ele divergir de um total já confiável.
    if valor_pag and ticket.total is not None and abs(valor_pag.valor - ticket.total) > 0.01:
        releitura = reler_valor(valor_pag.recorte)
        if releitura is not None and abs(releitura - ticket.total) < 0.01:
            ticket.observacoes.append(
                f"Valor pago relido por recorte: {valor_pag.valor:.2f} -> {releitura:.2f}"
            )
            valor_pag.valor = releitura
            ticket.valor_pagamento = releitura

    # ------------------------------------------------------------------ #
    # Descarte de leituras impossíveis.
    #
    # Ticket de táxi nunca custa R$ 0,00. Quando o OCR "lê" zero numa seção
    # (típico de dígito apagado: 'RS /0,00'), a linha foi interrompida e a
    # leitura é inválida. Zeramos essas seções para que a escolha do valor
    # caia na próxima fonte legível (ex.: VALOR TOTAL A PAGAR).
    # ------------------------------------------------------------------ #
    if ticket.valor_corrida is not None and ticket.valor_corrida <= 0:
        ticket.observacoes.append(
            f"VALOR ilegível (leitura {ticket.valor_corrida:.2f}); descartado."
        )
        ticket.valor_corrida = None
        if valor_corrida_item is not None:
            valor_corrida_item.valor = None
    if ticket.total is not None and ticket.total <= 0:
        ticket.observacoes.append(
            f"TOTAL ilegível (leitura {ticket.total:.2f}); descartado."
        )
        ticket.total = None
    if ticket.valor_pagamento is not None and ticket.valor_pagamento <= 0:
        ticket.observacoes.append(
            f"Valor pago ilegível (leitura {ticket.valor_pagamento:.2f}); descartado."
        )
        ticket.valor_pagamento = None

    # Se o VALOR foi zerado e existe TOTAL legível, tenta reler o VALOR
    # usando o TOTAL como alvo (recupera '1x,00' -> '76,00' em tickets ruins).
    if valor_corrida_item is not None and ticket.valor_corrida is None and ticket.total is not None:
        alvo = round(ticket.total - (ticket.bagagem or 0.0) - (ticket.pedagio or 0.0), 2)
        releitura = reler_valor(valor_corrida_item.recorte)
        if releitura is not None and abs(releitura - alvo) < 0.01:
            ticket.observacoes.append(
                f"VALOR relido por recorte (zerado): -> {releitura:.2f}"
            )
            ticket.valor_corrida = releitura
            valor_corrida_item.valor = releitura

    # ------------------------------------------------------------------ #
    # Escolha do valor principal por PRIORIDADE explícita (regra final):
    #
    #   1) VALOR (seção INFORMAÇÕES DE PAGAMENTO)  -> ticket.valor_pagamento
    #   2) VALOR TOTAL A PAGAR (seção INFORMAÇÕES DA CORRIDA) -> ticket.total
    #   3) VALOR DA CORRIDA (seção INFORMAÇÕES DA CORRIDA) -> ticket.valor_corrida
    #   4) nenhum legível -> Ñ IDENTIFICADO
    #
    # Piso: qualquer valor final abaixo de R$ 20,00 é tratado como Ñ IDENTIFICADO
    # (ticket de táxi não custa menos que isso; leitura < 20 é erro de OCR).
    # Não se usa soma de campos nem cálculo: o valor vem direto da seção lida.
    # ------------------------------------------------------------------ #
    escolhido: tuple[float, str, float] | None = None  # (valor, fonte, confiança)

    # 1) VALOR (pagamento)
    if ticket.valor_pagamento is not None:
        conf = valor_pag.confianca if valor_pag else 0.0
        escolhido = (ticket.valor_pagamento, "VALOR", conf)

    # 2) VALOR TOTAL A PAGAR
    if escolhido is None and ticket.total is not None:
        escolhido = (ticket.total, "VALOR TOTAL A PAGAR", totais[0].confianca)

    # 3) VALOR DA CORRIDA
    if escolhido is None and ticket.valor_corrida is not None:
        conf = valor_corrida_item.confianca if valor_corrida_item else 0.0
        escolhido = (ticket.valor_corrida, "VALOR DA CORRIDA", conf)

    if escolhido is None:
        ticket.observacoes.append(
            "VALOR, VALOR TOTAL A PAGAR e VALOR DA CORRIDA ilegíveis: "
            "valor Ñ IDENTIFICADO."
        )
        ticket.confianca = "baixa"
        return ticket

    valor_escolhido, fonte_principal, conf_fonte = escolhido

    # Piso de R$ 20,00: nada abaixo disso é aceito (erro de OCR).
    if valor_escolhido < MIN_VALOR_TICKET:
        ticket.observacoes.append(
            f"Valor {valor_escolhido:.2f} ({fonte_principal}) abaixo do piso "
            f"R$ {MIN_VALOR_TICKET:.2f}; valor Ñ IDENTIFICADO."
        )
        ticket.confianca = "baixa"
        return ticket

    # Quantas outras fontes confirmam esse mesmo valor (tolerância de 1 centavo).
    confirmacoes = sum(
        1 for v in (ticket.valor_corrida, ticket.total, ticket.valor_pagamento)
        if v is not None and abs(v - valor_escolhido) <= 0.01
    ) - 1  # desconta a própria fonte
    confirmacoes = max(confirmacoes, 0)

    ticket.valor = valor_escolhido
    ticket.observacoes.append(
        f"valor principal: {fonte_principal} ({ticket.valor:.2f}, conf. {conf_fonte:.0f})"
    )

    # Registra divergências entre as fontes conhecidas (útil para auditoria).
    outras = {
        "VALOR TOTAL A PAGAR": ticket.total,
        "VALOR DA CORRIDA": ticket.valor_corrida,
        "VALOR (pagamento)": ticket.valor_pagamento,
    }
    for nome, valor_outro in outras.items():
        if valor_outro is None:
            continue
        if abs(valor_outro - ticket.valor) > 0.01:
            ticket.observacoes.append(
                f"Divergência: {nome} = {valor_outro:.2f} x {fonte_principal} = {ticket.valor:.2f}"
            )

    # NÃO sobrescreve ticket.total nem ticket.valor_pagamento: eles guardam a
    # leitura de cada seção para auditoria/memória de cálculo. Apenas
    # ticket.valor recebe o valor escolhido pela prioridade acima.

    # Confiança: alta só quando uma segunda fonte confirma o mesmo valor.
    if confirmacoes >= 1:
        ticket.confianca = "alta" if conf_fonte >= 70 else "média"
    elif conf_fonte >= 70:
        ticket.confianca = "média"
    else:
        ticket.confianca = "baixa"

    if ticket.valor_corrida is not None and ticket.total is not None:
        esperado = round(ticket.valor_corrida + (ticket.bagagem or 0.0) + (ticket.pedagio or 0.0), 2)
        if abs(esperado - ticket.total) > 0.01:
            ticket.observacoes.append(
                f"Total ({ticket.total:.2f}) difere de corrida+extras ({esperado:.2f})."
            )

    return ticket


def processar(caminho: Path, idioma: str, debug: bool = False) -> Ticket:
    cinza = carregar_imagem(caminho)
    linhas = agrupar_linhas(ocr_palavras(cinza, idioma))
    if debug:
        print(f"  -- linhas reconhecidas em {caminho.name} --")
        for linha in linhas:
            texto = " ".join(p["texto"] for p in linha)
            if re.search(r"\d", texto):
                print(f"     y={linha[0]['y']:4d} | {texto}")
    return interpretar(caminho.name, linhas, cinza)


# --------------------------------------------------------------------------- #
# Saída
# --------------------------------------------------------------------------- #

def moeda(valor: float | None) -> str:
    if valor is None:
        return NAO_IDENTIFICADO
    return "R$ " + f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def imprimir(tickets: Sequence[Ticket]) -> None:
    largura_total = 92
    print("=" * largura_total)
    print("RESULTADO DA EXTRAÇÃO".center(largura_total))
    print("=" * largura_total)
    print(f"{'ARQUIVO':<18} {'Nº DO TICKET':<18} {'VALOR':>14}  {'CONF.':<6}")
    print("-" * largura_total)
    for t in tickets:
        numero = t.numero or NAO_IDENTIFICADO
        print(f"{t.arquivo:<18} {numero:<18} {moeda(t.valor):>14}  {t.confianca:<6}")
    print("-" * largura_total)

    for t in tickets:
        detalhes = []
        if t.valor_corrida is not None:
            detalhes.append(f"corrida {moeda(t.valor_corrida)}")
        if t.bagagem:
            detalhes.append(f"bagagem {moeda(t.bagagem)}")
        if t.pedagio:
            detalhes.append(f"pedágio {moeda(t.pedagio)}")
        if t.total is not None:
            detalhes.append(f"total {moeda(t.total)}")
        if t.valor_pagamento is not None:
            detalhes.append(f"pago {moeda(t.valor_pagamento)}")
        print(f"* {t.arquivo}")
        print(f"    número ....: {t.numero or NAO_IDENTIFICADO}")
        print(f"    valor .....: {moeda(t.valor)}")
        if detalhes:
            print(f"    memória ...: {' | '.join(detalhes)}")
        if t.data_hora:
            print(f"    data/hora .: {t.data_hora}")
        if t.forma_pagamento:
            print(f"    pagamento .: {t.forma_pagamento}")
        if t.observacoes:
            print(f"    notas .....: {'; '.join(t.observacoes)}")
    print("=" * largura_total)

    total_geral = sum(t.valor for t in tickets if t.valor is not None)
    print(f"Tickets processados: {len(tickets)}   |   Total: {moeda(round(total_geral, 2))}")
    # O Nº do ticket é informativo: só o valor precisa ser revisado manualmente.
    pendentes = [t.arquivo for t in tickets if t.valor is None]
    sem_numero = [t.arquivo for t in tickets if t.numero is None]
    if pendentes:
        print(f"ATENÇÃO: valor Ñ IDENTIFICADO -> revisar: {', '.join(pendentes)}")
    if sem_numero:
        print(f"Obs.: Nº Ñ IDENTIFICADO (não afeta o valor): {', '.join(sem_numero)}")
    print("=" * largura_total)


def pasta_area_de_trabalho() -> Path:
    """Retorna a Área de Trabalho do usuário (Desktop), com fallback para a pasta do script."""
    home = Path.home()
    candidatos = [
        home / "Desktop",
        home / "Área de Trabalho",
        home / "OneDrive" / "Desktop",
        home / "OneDrive" / "Área de Trabalho",
    ]
    for c in candidatos:
        if c.is_dir():
            return c
    return Path(__file__).resolve().parent


def salvar_excel(tickets: Sequence[Ticket], destino: Path | None = None) -> Path:
    """Salva o relatório de tickets em uma planilha .xlsx na Área de Trabalho.

    Estrutura:
      * Título "RELATÓRIO DE TICKETS" (linha 1, mesclado).
      * Topo fixo (linha 3, congelado): VALOR TOTAL (fórmula dinâmica, soma os valores).
      * Tabela com 7 pares (Nº DO TICKET | VALOR) lado a lado (linhas zebradas).
      * Todos os valores em negrito.
      * Cores: não lido (Ñ IDENTIFICADO) e com dúvida (confiança baixa/média).

    O VALOR TOTAL é gravado como fórmula SUM sobre todas as células de valor, de
    modo que qualquer valor digitado manualmente na planilha é somado
    automaticamente ao abrir/editar no Excel.
    """
    PARES_POR_LINHA = 5
    LINHA_FOLGA_MANUAL = 3  # linhas extras em branco para lançamentos manuais

    # Paleta harmônica (tons suaves).
    COR_TITULO = "1F3864"        # azul-escuro
    COR_TOPO_FUNDO = "2E75B6"    # azul
    COR_CABECALHO = "D9E1F2"     # azul-claro
    COR_ZEBRA = "F2F2F2"         # cinza-claro (linha alternada)
    COR_NAO_LIDO = "F8CBAD"      # laranja-suave (destaque forte)
    COR_DUVIDA = "FFE699"        # amarelo-suave
    COR_BORDA = "BFBFBF"

    wb = Workbook()
    ws = wb.active
    ws.title = "Tickets"

    n_colunas = PARES_POR_LINHA * 2  # cada par ocupa 2 colunas

    borda = Border(
        left=Side(style="thin", color=COR_BORDA),
        right=Side(style="thin", color=COR_BORDA),
        top=Side(style="thin", color=COR_BORDA),
        bottom=Side(style="thin", color=COR_BORDA),
    )
    centro = Alignment(horizontal="center", vertical="center")

    # ---------------------------------------------------------------- #
    # Título
    # ---------------------------------------------------------------- #
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_colunas)
    cel_titulo = ws.cell(row=1, column=1, value="RELATÓRIO DE TICKETS")
    cel_titulo.font = Font(bold=True, size=16, color="FFFFFF")
    cel_titulo.fill = PatternFill("solid", fgColor=COR_TITULO)
    cel_titulo.alignment = centro
    ws.row_dimensions[1].height = 28

    # ---------------------------------------------------------------- #
    # Topo fixo: VALOR TOTAL (fórmula dinâmica sobre todas as células de valor)
    # ---------------------------------------------------------------- #
    LINHA_CABECALHO = 5
    # As células de valor ocupam as colunas pares (2, 4, 6, ...) nas linhas
    # de dados. A fórmula SUM é montada com todas as linhas reservadas para
    # dados (inclusive as vazias), permitindo soma de valores digitados à mão.
    LINHA_PRIMEIRA_DADO = LINHA_CABECALHO + 1
    colunas_valor = [(i * 2 + 2) for i in range(PARES_POR_LINHA)]

    # Reserva linhas suficientes para os tickets (mínimo de 1) + folga para
    # lançamentos manuais.
    total_linhas_dados = max(1, -(-len(tickets) // PARES_POR_LINHA)) + LINHA_FOLGA_MANUAL
    LINHA_ULTIMO_DADO = LINHA_PRIMEIRA_DADO + total_linhas_dados - 1

    partes_formula: list[str] = []
    for col in colunas_valor:
        letra = get_column_letter(col)
        partes_formula.append(f"{letra}{LINHA_PRIMEIRA_DADO}:{letra}{LINHA_ULTIMO_DADO}")
    formula_total = "=SUM(" + ",".join(partes_formula) + ")"

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=3)
    cel_rotulo = ws.cell(row=3, column=1, value="VALOR TOTAL")
    cel_rotulo.font = Font(bold=True, size=12, color="FFFFFF")
    cel_rotulo.fill = PatternFill("solid", fgColor=COR_TOPO_FUNDO)
    cel_rotulo.alignment = centro

    ws.merge_cells(start_row=3, start_column=4, end_row=3, end_column=6)
    cel_total = ws.cell(row=3, column=4, value=formula_total)
    cel_total.number_format = 'R$ #,##0.00'
    cel_total.font = Font(bold=True, size=12, color="FFFFFF")
    cel_total.fill = PatternFill("solid", fgColor=COR_TOPO_FUNDO)
    cel_total.alignment = centro
    ws.row_dimensions[3].height = 22

    # Linha 4 em branco como respiro.

    # ---------------------------------------------------------------- #
    # Cabeçalho da tabela (linha 5): 7 pares Nº DO TICKET | VALOR
    # ---------------------------------------------------------------- #
    for i in range(PARES_POR_LINHA):
        col_num = i * 2 + 1
        col_val = i * 2 + 2
        for col, texto in ((col_num, "Nº DO TICKET"), (col_val, "VALOR")):
            cel = ws.cell(row=LINHA_CABECALHO, column=col, value=texto)
            cel.font = Font(bold=True, size=11, color=COR_TITULO)
            cel.fill = PatternFill("solid", fgColor=COR_CABECALHO)
            cel.alignment = centro
            cel.border = borda
    ws.row_dimensions[LINHA_CABECALHO].height = 20

    # ---------------------------------------------------------------- #
    # Dados: preenche lado a lado, desce de linha ao completar
    # ---------------------------------------------------------------- #
    # Remove o ticket de "reprocessamento" (imagens que falharam) não é
    # necessário; apenas itera na ordem de leitura (sem ordenar).
    linha_atual = LINHA_CABECALHO + 1  # índice da primeira linha de dados

    for indice, t in enumerate(tickets):
        posicao = indice % PARES_POR_LINHA
        if posicao == 0 and indice > 0:
            linha_atual += 1

        # Fundo branco em toda a linha (sem zebra).
        fundo_linha = "FFFFFF"
        for i in range(PARES_POR_LINHA):
            for col in (i * 2 + 1, i * 2 + 2):
                cel_bg = ws.cell(row=linha_atual, column=col)
                cel_bg.fill = PatternFill("solid", fgColor=fundo_linha)
                cel_bg.alignment = centro
                cel_bg.border = borda

        # Define a cor de destaque, se houver, por célula.
        nao_lido = t.valor is None
        duvida = (not nao_lido) and (t.confianca in ("baixa", "média"))
        cor_destaque = COR_NAO_LIDO if nao_lido else (COR_DUVIDA if duvida else fundo_linha)

        col_num = posicao * 2 + 1
        col_val = posicao * 2 + 2

        # Nº DO TICKET
        cel_num = ws.cell(row=linha_atual, column=col_num, value=t.numero or NAO_IDENTIFICADO)
        cel_num.alignment = centro
        cel_num.border = borda
        cel_num.fill = PatternFill("solid", fgColor=cor_destaque)
        cel_num.font = Font(bold=True)

        # VALOR (sempre em negrito)
        if t.valor is None:
            cel_val = ws.cell(row=linha_atual, column=col_val, value=NAO_IDENTIFICADO)
        else:
            cel_val = ws.cell(row=linha_atual, column=col_val, value=round(t.valor, 2))
            cel_val.number_format = 'R$ #,##0.00'
        cel_val.font = Font(bold=True)
        cel_val.alignment = centro
        cel_val.border = borda
        cel_val.fill = PatternFill("solid", fgColor=cor_destaque)

    # Linhas extras em branco (com bordas) para preenchimento manual.
    ultima_linha_preenchida = linha_atual
    for linha_extra in range(ultima_linha_preenchida + 1, LINHA_ULTIMO_DADO + 1):
        for i in range(PARES_POR_LINHA):
            for col in (i * 2 + 1, i * 2 + 2):
                cel = ws.cell(row=linha_extra, column=col)
                cel.alignment = centro
                cel.border = borda
                cel.fill = PatternFill("solid", fgColor="FFFFFF")
                if col % 2 == 0:  # coluna de VALOR
                    cel.number_format = 'R$ #,##0.00'
                    cel.font = Font(bold=True)

    # ---------------------------------------------------------------- #
    # Larguras e congelamento (5 pares: colunas mais largas, tabela arejada)
    # ---------------------------------------------------------------- #
    for i in range(PARES_POR_LINHA):
        ws.column_dimensions[get_column_letter(i * 2 + 1)].width = 18
        ws.column_dimensions[get_column_letter(i * 2 + 2)].width = 15

    # Congela título + topo + cabeçalho (linhas 1..5).
    ws.freeze_panes = ws.cell(row=LINHA_CABECALHO + 1, column=1)

    # Força o Excel/Calc a recalcular as fórmulas ao abrir o arquivo.
    wb.calculation.fullCalcOnLoad = True

    # Salva na Área de Trabalho. Se o arquivo estiver aberto no Excel (ou
    # bloqueado), grava "RELATORIO_DE_TICKETS_2.xlsx", "_3.xlsx", etc., para
    # o usuário nunca perder o resultado.
    if destino is None:
        base = pasta_area_de_trabalho() / "RELATORIO_DE_TICKETS.xlsx"
    else:
        base = Path(destino)

    candidatos = [base] + [
        base.with_name(f"{base.stem}_{n}{base.suffix}") for n in range(2, 100)
    ]
    for destino in candidatos:
        try:
            wb.save(destino)
            return destino
        except PermissionError:
            continue
    raise PermissionError(
        f"Não foi possível salvar: '{base.name}' está aberto em outro "
        f"programa. Feche a planilha e tente novamente."
    )


def _processar_seguro(argumentos: tuple[str, str, bool]) -> Ticket:
    """Wrapper para uso em processos: nunca levanta exceção, devolve Ticket."""
    caminho_txt, idioma, debug = argumentos
    caminho = Path(caminho_txt)
    try:
        return processar(caminho, idioma, debug=debug)
    except Exception as erro:  # pragma: no cover
        ticket = Ticket(arquivo=caminho.name)
        ticket.observacoes.append(f"falha no processamento: {erro}")
        return ticket


def processar_lote(
    arquivos: Sequence[Path],
    idioma: str,
    debug: bool = False,
    processos: int | None = None,
    progresso: Callable[[int, int, Path, Ticket], None] | None = None,
) -> list[Ticket]:
    """
    Processa as imagens em paralelo (mantendo a ordem original).

    O Tesseract é um processo externo, então rodar vários em paralelo escala
    quase linearmente com os núcleos da CPU. Com processos=1 cai no modo serial.

    `progresso` é chamado a cada imagem concluída como
    (concluídas, total, caminho, ticket).
    """
    n = processos if processos is not None else max(1, (os.cpu_count() or 2) - 1)
    n = max(1, min(n, len(arquivos))) if arquivos else 1
    total = len(arquivos)

    def avisar(indice: int, caminho: Path, ticket: Ticket) -> None:
        if progresso:
            try:
                progresso(indice + 1, total, caminho, ticket)
            except Exception:
                pass

    def sem_valor(ticket: Ticket) -> bool:
        """Falha real = valor não identificado (observações são normais)."""
        return ticket.valor is None and ticket.numero is None

    tarefas = [(str(p), idioma, debug) for p in arquivos]

    if n == 1 or total <= 1:
        tickets = []
        for indice, caminho in enumerate(arquivos):
            ticket = _processar_seguro(tarefas[indice])
            tickets.append(ticket)
            estado = "FALHA" if sem_valor(ticket) else "ok"
            print(f"  [{estado}] {caminho.name}")
            avisar(indice, caminho, ticket)
        return tickets

    tickets = []
    with ProcessPoolExecutor(max_workers=n) as executor:
        # as_completed entrega na ordem de término (barra de progresso fluida),
        # mas guardamos pela posição original para manter a ordem do relatório.
        resultados: list[Ticket | None] = [None] * total
        futuros = {executor.submit(_processar_seguro, tarefa): i
                   for i, tarefa in enumerate(tarefas)}
        concluidos = 0
        for futuro in as_completed(futuros):
            indice = futuros[futuro]
            caminho = arquivos[indice]
            try:
                ticket = futuro.result()
            except Exception as erro:  # pragma: no cover
                ticket = Ticket(arquivo=caminho.name)
                ticket.observacoes.append(f"falha no processamento: {erro}")
            resultados[indice] = ticket
            concluidos += 1
            estado = "FALHA" if sem_valor(ticket) else "ok"
            print(f"  [{estado}] {caminho.name}")
            avisar(concluidos - 1, caminho, ticket)
        tickets = [t if t is not None else Ticket(arquivo=arquivos[i].name)
                   for i, t in enumerate(resultados)]
    return tickets


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extrai Nº do ticket e valor das imagens de ticket de táxi."
    )
    parser.add_argument("pasta", nargs="?", default="img_ticket",
                        help="Pasta com as imagens (padrão: img_ticket)")
    parser.add_argument("--json", metavar="ARQUIVO", nargs="?", const="tickets.json",
                        help="Salva o resultado em JSON (padrão: tickets.json)")
    parser.add_argument("--debug", action="store_true",
                        help="Mostra todas as linhas reconhecidas pelo OCR")
    parser.add_argument("--excel", metavar="ARQUIVO", nargs="?", const="",
                        help="Salva o relatório em Excel (padrão: Área de Trabalho)")
    parser.add_argument("--sem-excel", action="store_true",
                        help="Não gera a planilha Excel")
    parser.add_argument("--processos", type=int, default=None, metavar="N",
                        help="Nº de processos paralelos (padrão: núcleos - 1)")
    args = parser.parse_args(argv)

    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    idioma = idioma_disponivel()

    pasta = Path(args.pasta)
    if not pasta.is_absolute():
        pasta = Path(__file__).resolve().parent / pasta
    if not pasta.is_dir():
        print(f"ERRO: pasta não encontrada: {pasta}", file=sys.stderr)
        return 2

    arquivos = sorted(p for p in pasta.iterdir() if p.suffix.lower() in EXTENSOES)
    if not arquivos:
        print(f"ERRO: nenhuma imagem encontrada em {pasta}", file=sys.stderr)
        return 2

    print(f"Pasta: {pasta}")
    print(f"Idioma OCR: {idioma} | Imagens: {len(arquivos)}")
    print()

    n_proc = args.processos if args.processos is not None else max(1, (os.cpu_count() or 2) - 1)
    n_proc = max(1, min(n_proc, len(arquivos)))
    if len(arquivos) > 1:
        print(f"Processando {len(arquivos)} imagem(ns) com {n_proc} processo(s) em paralelo...")
        print()

    inicio = time.perf_counter()
    tickets = processar_lote(arquivos, idioma, debug=args.debug, processos=n_proc)
    duracao = time.perf_counter() - inicio
    print()
    print(f"Tempo total: {duracao:.1f}s "
          f"({duracao / max(1, len(tickets)):.2f}s por ticket)")
    print()

    imprimir(tickets)

    if args.json:
        destino = Path(args.json)
        if not destino.is_absolute():
            destino = Path(__file__).resolve().parent / destino
        destino.write_text(
            json.dumps([t.como_dicionario() for t in tickets], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"JSON salvo em: {destino}")

    # Planilha Excel na Área de Trabalho (gerada por padrão).
    if not args.sem_excel:
        destino_excel = Path(args.excel) if args.excel else None
        if destino_excel is not None and not destino_excel.is_absolute():
            destino_excel = Path(__file__).resolve().parent / destino_excel
        try:
            caminho_excel = salvar_excel(tickets, destino_excel)
            print(f"Planilha salva em: {caminho_excel}")
        except Exception as erro:  # pragma: no cover
            print(f"ERRO ao salvar a planilha: {erro}", file=sys.stderr)

    # O número do ticket é informativo; o código de saída depende só do valor.
    return 0 if all(t.valor is not None for t in tickets) else 1


if __name__ == "__main__":
    # Necessário no Windows para o processamento paralelo (spawn).
    from multiprocessing import freeze_support
    freeze_support()
    raise SystemExit(main())
