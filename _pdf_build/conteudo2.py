# -*- coding: utf-8 -*-
"""Seções 3 a 11 do documento. Importado por gerar_pdf.py."""
from __future__ import annotations

from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.enums import TA_CENTER, TA_LEFT

from gerar_pdf import (
    AZUL, AZUL_CLARO, AZUL_ESCURO, AZUL_MEDIO, AZUL_PALIDO, AMARELO_CLARO, BRANCO,
    CINZA, CINZA_BORDA, LARGURA, LARANJA, VERDE, VERDE_CLARO, VERMELHO, VERMELHO_CLARO,
    TEXTO, S_BODY, S_BODY_SM, S_CAP, S_CELL, S_CELL_C, S_CELL_H, S_H2, S_H3,
    Barra, TituloSecao, caixa, tabela, bloco_codigo, figura, fluxo_etapas, st, P, B, code,
)


# =========================================================================== #
def secao_pipeline():
    out = [TituloSecao(3, "O pipeline de OCR em 6 etapas"), Spacer(1, 10)]

    out.append(P("Esta é a parte central da automação. Cada imagem percorre seis etapas; qualquer erro "
                 "cometido em uma delas se propaga. Por isso o desenho prioriza <b>redundância</b> (várias "
                 "tentativas de leitura) seguida de <b>consenso</b>."))

    # ---- 3.1
    out.append(P("Etapa 1 — Leitura robusta do arquivo", S_H2))
    out.append(P("A imagem é carregada em tons de cinza. Detalhe importante: usa-se "
                 f"{code('np.fromfile')} + {code('cv2.imdecode')} em vez de {code('cv2.imread')}, porque "
                 "o imread do OpenCV falha silenciosamente com caminhos que contêm acentos (comum no "
                 "Windows: \"Área de Trabalho\", \"Faturamento\")."))

    # ---- 3.2
    out.append(P("Etapa 2 — Normalização de escala e 4 binarizações", S_H2))
    out.append(P("Antes de qualquer OCR, o sistema estima o tamanho do texto usando componentes conexos "
                 "e amplia a imagem para que o texto fique com cerca de 60 px de altura — o Tesseract é "
                 "muito sensível a texto pequeno. A ampliação é limitada a 1700 px de lado, pois acima "
                 "disso o OCR só fica mais lento, sem ganho de acurácia."))
    out.append(P("Em seguida, a imagem é binarizada de <b>quatro</b> maneiras diferentes. Cada uma "
                 "favorece um tipo de defeito; depois o sistema vota entre elas."))
    out.append(Spacer(1, 4))

    bins = [
        ["Variante", "Como é gerada", "Para que serve"],
        ["<b>otsu</b>", "Limiar automático global (Otsu)",
         "Tickets nítidos, texto escuro normal. É a base de todas as outras."],
        ["<b>open</b>", "Otsu + abertura morfológica (kernel 2×2)",
         "Remove grão e 'cabelinhos' do papel térmico, regularizando dígitos degradados."],
        ["<b>soft</b>", "Blur leve + limiar fixo suave (205)",
         "Preserva texto cinza muito claro que o Otsu apagaria por completo."],
        ["<b>clahe_open</b>", "CLAHE (contraste local) + threshold adaptativo + abertura",
         "Resgata texto apagado em regiões com iluminação desigual."],
    ]
    out.append(tabela(bins, [2.5 * cm, 6.2 * cm, LARGURA - 8.7 * cm]))
    out.append(Spacer(1, 8))

    out.append(P("Veja o efeito nas imagens reais (ticket com tinta degradada):", S_BODY_SM))
    out.append(Spacer(1, 2))

    # galeria de binarizacoes
    minis = []
    rotulos = [("bin_otsu.png", "otsu"), ("bin_open.png", "open"),
               ("bin_soft.png", "soft"), ("bin_clahe_open.png", "clahe_open")]
    largura_mini = (LARGURA - 3 * 0.3 * cm) / 4
    celulas, legendas = [], []
    for arq, nome in rotulos:
        f = figura(arq, largura_mini)
        if f is not None:
            celulas.append(f)
            legendas.append(Paragraph(f'<font name="Mono" size="7.4" color="#1F4E79">{nome}</font>',
                                      S_CAP))
    if celulas:
        t = Table([celulas, legendas], colWidths=[largura_mini] * len(celulas))
        t.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, 0), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("GRID", (0, 0), (-1, 0), 0.5, CINZA_BORDA),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#FAFCFE")),
        ]))
        out.append(t)
        out.append(P("▲ As quatro binarizações aplicadas ao mesmo ticket. Repare como a variante 'soft' "
                     "mantém traços que o 'otsu' elimina — é isso que permite ler recibos apagados.",
                     S_CAP))
    out.append(Spacer(1, 10))

    # ---- 3.3
    out.append(P("Etapa 3 — OCR com caixas por palavra", S_H2))
    out.append(P("Cada binarização é enviada ao Tesseract com "
                 f"{code('--psm 6')} (assume um bloco uniforme de texto) e "
                 f"{code('preserve_interword_spaces=1')} (mantém os espaços entre colunas, essencial "
                 "para separar rótulo de valor). O idioma é <b>por</b> (português). Usa-se "
                 f"{code('image_to_data')} para obter, além do texto, a <b>caixa</b> (x, y, w, h) e a "
                 "<b>confiança</b> de cada palavra."))
    out.append(bloco_codigo("extrair_tickets.py — OCR configurado para recibos",
                            """dados = pytesseract.image_to_data(
    img, lang=idioma,
    config="--psm 6 -c preserve_interword_spaces=1",
    output_type=pytesseract.Output.DICT,
)
for i, bruto in enumerate(dados["text"]):
    texto = (bruto or "").strip()
    conf  = float(dados["conf"][i])
    if not texto or conf < 20:      # descarta ruído de baixa confiança
        continue
    candidatos.append({... "conf": conf, "fonte": nome_da_binarizacao})"""))
    out.append(Spacer(1, 8))

    # ---- 3.4
    out.append(P("Etapa 4 — Consenso entre as leituras", S_H2))
    out.append(P("As quatro binarizações podem ler o mesmo número de quatro formas diferentes. Em vez de "
                 "escolher uma, o sistema reúne todas as leituras e resolve em três passos:"))

    passos = [
        ["Passo", "Regra", "Efeito prático"],
        ["<b>1. Ordenar</b>", "Todas as palavras por confiança decrescente",
         "A leitura mais confiável é considerada primeiro."],
        ["<b>2. Deduplicar</b>",
         "Descarta leituras que representam o mesmo símbolo físico, comparando as caixas por "
         "<b>IoU</b> (interseção sobre união) ≥ 0,35 ou dígitos iguais em posição muito próxima",
         "Uma única palavra no papel vira uma única entrada, mesmo que as 4 binarizações a tenham "
         "lido 4 vezes."],
        ["<b>3. Agrupar em linhas</b>",
         "Agrupa por centro vertical, com tolerância baseada na <b>altura mediana</b> do documento",
         "Não funde \"VALOR DA CORRIDA\" com \"VALOR TOTAL\" quando um token sai com caixa exagerada."],
    ]
    out.append(tabela(passos, [2.5 * cm, 7.2 * cm, LARGURA - 9.7 * cm]))
    out.append(Spacer(1, 8))
    out.append(bloco_codigo("extrair_tickets.py — deduplicação por IoU",
                            """candidatos.sort(key=lambda p: -p["conf"])   # melhores primeiro
aceitas = []
for palavra in candidatos:
    if any(_mesmo_simbolo(palavra, outra) for outra in aceitas):
        continue                               # já temos essa palavra
    aceitas.append(palavra)                    # vence a de maior confiança"""))
    out.append(Spacer(1, 8))

    # ---- 3.5
    out.append(P("Etapa 5 — Regras de negócio: dar sentido ao texto", S_H2))
    out.append(P("Com as linhas reconstruídas, entra a parte que não é visão computacional, e sim "
                 "conhecimento do domínio do recibo. O sistema:"))

    regras = [
        "<b>Identifica o rótulo</b> de cada linha com expressões regulares tolerantes a erros de OCR: "
        + code("v[a4][l1i][o0]r") + " encontra VALOR mesmo lendo \"VA4LOR\" ou \"VAL0R\"; "
        + code("t[o0]t[ai1l]") + " encontra TOTAL; " + code("c[o0]rr[i1l]da") + " encontra CORRIDA.",
        "<b>Separa as seções</b> do recibo usando a linha que contém PAGAMENTO como divisor: o que está "
        "acima é a corrida; abaixo é o pagamento.",
        "<b>Escolhe o VALOR certo</b>: entre vários \"VALOR\", pega o que aparece imediatamente antes do "
        "primeiro \"VALOR TOTAL A PAGAR\" (a via do passageiro).",
        "<b>Lê valores da coluna da direita</b> (x ≥ 55% da largura), dando prioridade a números com "
        "vírgula decimal explícita — sem isso, \"124,00\" + \"144\" poderia virar 112500.00.",
        "<b>Descarta datas e horas</b> ao procurar valores, para não confundir \"12/03/2026\" com dinheiro.",
        "<b>Corrige dígitos</b> do número do ticket: O→0, l/I/|→1, S→5, B→8, Z→2, G→6.",
        "<b>Rejeita falsos positivos</b>: números parecidos com ano (2020–2030) não são aceitos como "
        "número de ticket.",
        "<b>Descarta leituras de baixa confiança</b>: abaixo de 55 (número) ou 45 (valor), em vez de "
        "arriscar, o campo vira \"Ñ IDENTIFICADO\".",
    ]
    for r in regras:
        out.append(P(f'<font color="#2E75B6">▪</font> {r}', S_BODY_SM))
    out.append(Spacer(1, 8))

    out.append(bloco_codigo("extrair_tickets.py — validação aritmética do ticket",
                            """esperado = round(valor_corrida + bagagem + pedagio, 2)
if abs(esperado - total) > 0.01:
    # o total não fecha: relê o recorte da linha com vários limiares
    releitura = reler_valor(totais[0].recorte)
    if releitura is not None and abs(releitura - esperado) < 0.01:
        ticket.total = releitura
        ticket.observacoes.append("Total relido por recorte")"""))
    out.append(Spacer(1, 8))

    out.append(caixa("A regra mais poderosa: aritmética como validador",
                     "Um recibo é internamente consistente: <b>VALOR DA CORRIDA + PEDÁGIO + BAGAGEM = "
                     "TOTAL</b>, e o TOTAL deveria bater com o valor pago. Quando a leitura do OCR fura essa "
                     "igualdade, há erro — o sistema detecta, recorta só aquela região, relê com cinco "
                     "limiares diferentes em duas escalas e aceita a leitura que fecha a conta. É "
                     "checagem de integridade aplicada a OCR.", "ok"))
    out.append(Spacer(1, 10))

    # ---- 3.6
    out.append(P("Etapa 6 — Objeto Ticket e saída", S_H2))
    out.append(P("O resultado de tudo isso é um objeto <b>Ticket</b> (dataclass) com campos tipados: "
                 "número, valor da corrida, pedágio, bagagem, total, valor pago, data/hora, forma de "
                 "pagamento, origem, confiança e uma lista de observações explicando cada correção "
                 "feita. Esse objeto alimenta o terminal, o JSON opcional e a planilha."))
    out.append(bloco_codigo("extrair_tickets.py — o contrato de saída",
                            """@dataclass
class Ticket:
    arquivo: str
    numero: str | None = None
    valor: float | None = None
    valor_corrida: float | None = None
    total: float | None = None
    data_hora: str | None = None
    forma_pagamento: str | None = None
    origem: str = ""
    confianca: str = "baixa"
    observacoes: list[str] = field(default_factory=list)"""))
    out.append(PageBreak())
    return out


# =========================================================================== #
def secao_bibliotecas():
    out = [TituloSecao(4, "Bibliotecas escolhidas e por quê"), Spacer(1, 10)]

    out.append(P("Toda a pilha foi escolhida com um critério em mente: <b>rodar 100% offline, em Windows, "
                 "dentro de um único .exe</b>. Isso elimina dependências pesadas de nuvem e favorece "
                 "bibliotecas maduras, com pacotes binários prontos para Windows."))

    principal = [
        ["Biblioteca", "Versão", "Papel no projeto"],
        ["<b>Python</b>", "3.13.2", "Linguagem-base. Ecossistema completo de visão computacional e "
                                    "excelente suporte do PyInstaller."],
        ["<b>OpenCV</b> (<i>opencv-python</i>)", "5.0.0",
         "Pré-processamento: tons de cinza, threshold, morfologia, CLAHE, componentes conexos, "
         "resize. Também decodifica imagens de forma segura com caminhos acentuados."],
        ["<b>numpy</b>", "2.1.3", "Base numérica do OpenCV. Usado para medir a mediana de alturas de "
                                  "texto e lidar com os pixels."],
        ["<b>Tesseract</b> (OCR)", "5.5.2", "Motor de OCR propriamente dito. Roda em processo separado, "
                                            "o que permite paralelismo real."],
        ["<b>pytesseract</b>", "0.3.13", "Ponte Python → tesseract.exe. Dá acesso ao modo TSV/DICT com "
                                         "caixas e confianças por palavra."],
        ["<b>Tkinter</b>", "nativo", "Interface gráfica. Já vem no Python, zero dependências extras — "
                                     "decisivo para o .exe ficar simples."],
        ["<b>openpyxl</b>", "3.1.5", "Gera o .xlsx com fórmulas, cores, bordas e mesclagem de células."],
        ["<b>PyInstaller</b>", "6.x", "Empacota tudo em um .exe de arquivo único, sem console."],
    ]
    out.append(tabela(principal, [4.0 * cm, 1.7 * cm, LARGURA - 5.7 * cm]))
    out.append(Spacer(1, 12))

    out.append(P("Por que OpenCV e não Pillow ou scikit-image?", S_H2))
    comp = [
        ["Critério", "OpenCV", "Pillow", "scikit-image"],
        ["Velocidade em operações por pixel", "Altíssima (código C++ otimizado, uso de SIMD)", "Boa, "
         "porém limitada a operações básicas", "Boa, mas mais lenta em loops grandes"],
        ["Binarização adaptativa, CLAHE, morfologia", "Nativo e rápido", "Não tem nativo", "Tem, mais lento"],
        ["Componentes conexos com estatísticas", "Uma chamada (<i>connectedComponentsWithStats</i>)", "—", "Sim"],
        ["Leitura de imagem com caminho acentuado", "Resolvida com <i>imdecode</i>", "Sim", "Via Pillow"],
        ["Tamanho no pacote final", "Médio (mas necessário)", "Pequeno", "Grande (depende de numpy/scipy)"],
    ]
    out.append(tabela(comp, [4.6 * cm, 4.6 * cm, 3.6 * cm, LARGURA - 12.8 * cm]))
    out.append(Spacer(1, 6))
    out.append(P("<i>Pillow ainda é usado — mas apenas para a interface (exibir ícones no Tkinter) e não "
                 "no processamento pesado.</i>", S_BODY_SM))
    out.append(Spacer(1, 10))

    out.append(P("Por que Tesseract e não EasyOCR / PaddleOCR / OCR de nuvem?", S_H2))
    comp2 = [
        ["Opção", "Prós", "Contras para este projeto"],
        ["<b>Tesseract 5</b><br/>(escolhido)", "Leve, offline, rápido, gratuito, ótimo em texto impresso "
                                               "nítido, modelo português pronto, fácil de embutir no .exe",
         "Sofre com texto muito degradado — mitigado pelas 4 binarizações e pela validação aritmética"],
        ["EasyOCR", "Bom em cenas complexas e inclinadas", "Depende de PyTorch (centenas de MB), lento em "
                                                           "CPU, .exe gigante"],
        ["PaddleOCR", "Excelente acurácia", "Dependências pesadas e complicadas de empacotar no Windows"],
        ["OCR de nuvem", "Reconhecimento de ponta", "Requer internet, custo por chamada e envio de dados "
                                                     "sensíveis a terceiros"],
    ]
    out.append(tabela(comp2, [3.2 * cm, 6.4 * cm, LARGURA - 9.6 * cm]))
    out.append(Spacer(1, 12))

    out.append(caixa("Uma decisão de arquitetura que muda tudo: Tesseract como processo externo",
                     "O pytesseract não 'importa' o motor de OCR — ele chama <b>tesseract.exe</b> "
                     "como subprocesso. Isso tem duas consequências valiosas: (1) o Python não fica "
                     "preso ao GIL durante o OCR, então vários processos podem rodar OCR em paralelo de "
                     "verdade; (2) é trivial embutir o tesseract.exe e sua pasta tessdata dentro do .exe "
                     "e apontar o caminho na inicialização.", "info"))
    out.append(Spacer(1, 10))

    out.append(P("Configuração do Tesseract na inicialização", S_H3))
    out.append(P("Em vez de exigir instalação no computador do usuário, o programa procura o "
                 "tesseract.exe em uma lista de locais — inclusive dentro do pacote PyInstaller — e "
                 "define as variáveis de ambiente corretas antes de qualquer OCR."))
    out.append(bloco_codigo("extrair_tickets.py — localização automática do Tesseract",
                            """candidatos_exe = [
    base / "Tesseract-OCR" / "tesseract.exe",     # ao lado do .exe (portátil)
    recursos / "Tesseract-OCR" / "tesseract.exe", # dentro do .exe (_MEIPASS)
    Path(r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"),
]
exe = next((str(c) for c in candidatos_exe if c.is_file()),
           shutil.which("tesseract"))
pytesseract.pytesseract.tesseract_cmd = exe
os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)"""))
    out.append(PageBreak())
    return out


# =========================================================================== #
def secao_decisoes():
    out = [TituloSecao(5, "Decisões de engenharia"), Spacer(1, 10)]

    out.append(P("Esta seção reúne as escolhas que mais influenciaram a acurácia e o desempenho — e o "
                 "raciocínio por trás de cada uma. São bons pontos de partida para estudos futuros."))

    decisoes = [
        ("1", "Votar entre 4 binarizações em vez de escolher uma",
         "Cada binarização 'vê' o papel de um jeito. O consenso por confiança + IoU combina o melhor "
         "de todas: onde o Otsu falha, o soft pode acertar, e a deduplicação evita contar a mesma "
         "palavra várias vezes.",
         "Acurácia bem maior em tickets apagados; custo de ~4× o tempo de OCR por imagem — aceitável "
         "para lotes pequenos."),
        ("2", "IoU como critério de 'mesma palavra'",
         "Caixas do mesmo símbolo (fundamentalmente) se sobrepõem muito; caixas de símbolos distintos "
         "não. O limiar de 0,35 é tolerante o bastante para pequenas variações de binarização.",
         "Elimina duplicatas sem precisar de OCR pós-processado manual."),
        ("3", "Confiança mínima e 'Ñ IDENTIFICADO'",
         "É melhor admitir que não leu do que registrar um número errado. Erros em planilha de "
         "faturamento são caros.",
         "Nenhum valor é inventado; a planilha destaca em cor os casos não lidos."),
        ("4", "Aritmética do ticket como validador",
         "A relação corrida + extras = total é uma verificação de integridade gratuita. Quando falha, "
         "há suspeita concreta de erro de leitura.",
         "Corrige trocas como '2' → '7' que nenhuma regra sintática pegaria."),
        ("5", "Multiprocessamento, não multithreading, para o OCR",
         "O gargalo é o tesseract.exe. Vários processos independentes escalam quase linearmente com "
         "os núcleos; threads compartilhariam apenas o GIL.",
         "Corte de ~3,5 s para ~1 s por ticket em lotes grandes."),
        ("6", "Limitar 0,55 × largura para a coluna de valores",
         "Nos recibos, os valores estão sempre alinhados à direita. Restringir a busca evita ler "
         "números do miolo do texto (placa, telefone, código de autorização).",
         "Menos falsos positivos sem perder nenhum valor real."),
        ("7", "Ampliar o texto até 60 px, com teto de 1700 px",
         "O Tesseract precisa de altura de texto suficiente; acima de certo ponto, ganho zero e custo "
         "alto.",
         "Equilíbrio entre acurácia e velocidade."),
        ("8", "console=False no .exe",
         "O tesseract.exe é chamado de dentro do programa; sem cuidado, cada chamada abriria uma "
         "janela de console piscando.",
         "Execução silenciosa via STARTUPINFO + CREATE_NO_WINDOW."),
    ]
    for num, tit, motivo, efeito in decisoes:
        dados = Table(
            [[Paragraph(f'<font color="white" size="12"><b>{num}</b></font>',
                        st("num", alignment=TA_CENTER, spaceAfter=0)),
              Paragraph(f"<b>{tit}</b>", st("dt", fontName="AppSemi", fontSize=9.6, leading=12.5,
                                            textColor=AZUL, spaceAfter=1))]],
            colWidths=[0.95 * cm, LARGURA - 0.95 * cm - 2],
        )
        dados.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), AZUL_MEDIO),
            ("BACKGROUND", (1, 0), (1, 0), AZUL_PALIDO),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("BOX", (0, 0), (-1, -1), 0.6, AZUL_CLARO),
        ]))
        corpo = Table(
            [[Paragraph(f"<b>Motivo.</b> {motivo} "
                        f"<font color='#1E7A45'><b>Efeito.</b> {efeito}</font>", S_BODY_SM)]],
            colWidths=[LARGURA - 2],
        )
        corpo.setStyle(TableStyle([
            ("TOPPADDING", (0, 0), (0, 0), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#E3E8ED")),
        ]))
        out.append(KeepTogether([dados, corpo, Spacer(1, 5)]))

    out.append(Spacer(1, 2))
    out.append(caixa("Como o erro é tratado: exemplo real",
                     "No ticket com tinta degradada, a linha VALOR DA CORRIDA saiu como \"4U,00\" — leitura "
                     "que daria 4,00. Mas o TOTAL (76,00) batia com o valor pago. A conta "
                     "76,00 − extras = 76,00 apontava para valor de corrida 76,00. O sistema então recortou "
                     "apenas aquela região, releu em 2 escalas × 5 limiares e obteve 76,00, fechando a "
                     "conta. O ticket foi salvo com observação: 'VALOR relido por recorte'.", "dica"))
    out.append(PageBreak())
    return out
