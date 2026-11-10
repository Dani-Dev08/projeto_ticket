# -*- coding: utf-8 -*-
"""Conteúdo (seções) do PDF de documentação. Importado por gerar_pdf.py."""
from __future__ import annotations

from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether

from gerar_pdf import (
    AZUL, AZUL_CLARO, AZUL_ESCURO, AZUL_MEDIO, AZUL_PALIDO, AMARELO_CLARO, BRANCO,
    CINZA, CINZA_BORDA, LARGURA, LARANJA, VERDE, VERDE_CLARO, VERMELHO, VERMELHO_CLARO,
    TEXTO, S_BODY, S_BODY_SM, S_CAP, S_CELL, S_CELL_C, S_CELL_H, S_H2, S_H3,
    Barra, TituloSecao, caixa, tabela, bloco_codigo, figura, fluxo_etapas, st, P, B, code,
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT


# =========================================================================== #
def capa():
    out = [Spacer(1, 3.0 * cm)]
    out.append(P('<font color="#7FB6E8" size="10"><b>DOCUMENTAÇÃO TÉCNICA · ESTUDO DE CASO</b></font>',
                 st("x", alignment=TA_LEFT, spaceAfter=10)))
    out.append(P('<font color="white" size="33"><b>Extrator de<br/>Tickets de Táxi</b></font>',
                 st("x", alignment=TA_LEFT, leading=39, spaceAfter=0)))
    out.append(Spacer(1, 0.3 * cm))
    out.append(Barra(7.5 * cm, 4, colors.HexColor("#F0A02C")))
    out.append(Spacer(1, 0.7 * cm))
    out.append(P('<font color="#C9DCEE" size="12">Automação de leitura de recibos com OCR 100% local —<br/>'
                 'arquitetura, bibliotecas, decisões de engenharia e resultados</font>',
                 st("x", alignment=TA_LEFT, leading=17)))

    out.append(Spacer(1, 2.0 * cm))
    cartoes = [
        ["<font color='#7FB6E8'><b>44</b></font><br/><font color='#D6E4F0' size='8'>imagens processadas</font>",
         "<font color='#7FB6E8'><b>~3,5 s</b></font><br/><font color='#D6E4F0' size='8'>por ticket (serial)</font>"],
        ["<font color='#7FB6E8'><b>4</b></font><br/><font color='#D6E4F0' size='8'>binarizações + votação</font>",
         "<font color='#7FB6E8'><b>0</b></font><br/><font color='#D6E4F0' size='8'>dados enviados à nuvem</font>"],
    ]
    t = Table(
        [[Paragraph(c, st("cc", fontSize=17, leading=20, alignment=TA_CENTER, textColor=BRANCO))
          for c in linha] for linha in cartoes],
        colWidths=[5.3 * cm, 5.3 * cm], rowHeights=[1.7 * cm, 1.7 * cm],
    )
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1B3355")),
        ("BOX", (0, 0), (0, 0), 0.8, colors.HexColor("#33587F")),
        ("BOX", (1, 0), (1, 0), 0.8, colors.HexColor("#33587F")),
        ("BOX", (0, 1), (0, 1), 0.8, colors.HexColor("#33587F")),
        ("BOX", (1, 1), (1, 1), 0.8, colors.HexColor("#33587F")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    out.append(t)

    out.append(Spacer(1, 2.3 * cm))
    out.append(P('<font color="#7F95AC" size="8.5">Python 3.13.2 · OpenCV 5.0 · Tesseract 5.5.2 (por.traineddata) · '
                 'pytesseract 0.3.13 · Tkinter · openpyxl 3.1.5 · PyInstaller</font>',
                 st("x", alignment=TA_LEFT, spaceAfter=2)))
    out.append(P('<font color="#7F95AC" size="8.5">Novembro de 2026</font>',
                 st("x", alignment=TA_LEFT)))
    out.append(PageBreak())
    return out


# =========================================================================== #
def sumario():
    out = [Spacer(1, 0.2 * cm)]
    out.append(P('<font color="#1F4E79" size="19"><b>Sumário</b></font>',
                 st("x", alignment=TA_LEFT, spaceAfter=2)))
    out.append(Barra(5.2 * cm, 3.5, AZUL_MEDIO))
    out.append(Spacer(1, 0.45 * cm))

    itens = [
        ("01", "O problema", "O que precisa ser lido, dificuldades reais e por que OCR local"),
        ("02", "Arquitetura da solução", "Camadas, módulos, fluxo de dados ponta a ponta"),
        ("03", "O pipeline de OCR em 6 etapas", "Da imagem crua ao ticket estruturado"),
        ("04", "Bibliotecas escolhidas e por quê", "Comparativo técnico de cada dependência"),
        ("05", "Decisões de engenharia", "Votação entre binarizações, IoU, revalidação por recorte"),
        ("06", "Interface e concorrência", "Tkinter, threads, multiprocessamento e tempo restante"),
        ("07", "Saída em Excel", "openpyxl, fórmula SUM dinâmica e sinalização de erros"),
        ("08", "Distribuição: um único .exe", "PyInstaller, Tesseract embutido e modo headless"),
        ("09", "Desempenho e resultados", "Medições reais, paralelismo e limites"),
        ("10", "Limitações e próximos passos", "Onde o sistema erra e como evoluir com ML"),
        ("11", "Glossário e consulta rápida", "Termos, comandos, arquivos e mapa do código"),
    ]
    linhas = []
    for num, tit, desc in itens:
        linhas.append([
            Paragraph(f'<font color="#2E75B6" size="13"><b>{num}</b></font>', S_CELL),
            Paragraph(f"<b>{tit}</b><br/><font color='#5A6672' size='7.8'>{desc}</font>", S_CELL),
        ])
    t = Table(linhas, colWidths=[1.3 * cm, LARGURA - 1.3 * cm])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LINEBELOW", (0, 0), (-1, -2), 0.4, colors.HexColor("#E3E8ED")),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
    ]))
    out.append(t)
    out.append(Spacer(1, 0.5 * cm))
    out.append(caixa("Como ler este documento",
                     "Cada seção combina explicação conceitual, trechos reais do código e imagens "
                     "geradas a partir das próprias fotos processadas pelo sistema. Os quadros "
                     "coloridos resumem o ponto principal: <b>azul</b> = contexto e funcionamento, "
                     "<b>verde</b> = por que a escolha foi boa, <b>amarelo</b> = atenção ou "
                     "limitação, <b>laranja</b> = dica prática de uso.", "info"))
    out.append(Spacer(1, 6))
    out.append(caixa("Estrutura de arquivos do projeto",
                     [code("projeto_ticket/") + "<br/>"
                      + code("├── extrair_tickets.py") + " &nbsp;— núcleo: OCR, regras, Excel, CLI<br/>"
                      + code("├── app_tickets.py") + " &nbsp;&nbsp;&nbsp;&nbsp;— interface Tkinter<br/>"
                      + code("├── TicketsTaxis.spec") + " &nbsp;&nbsp;&nbsp;— receita do PyInstaller<br/>"
                      + code("├── build_exe.ps1") + " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;— script de build do .exe<br/>"
                      + code("├── img_ticket/") + " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;— imagens de entrada<br/>"
                      + code("├── README.md") + " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;— instruções de uso<br/>"
                      + code("└── Documentacao_TicketsTaxis.pdf") + "  — este documento"], "ok"))
    out.append(PageBreak())
    return out


# =========================================================================== #
def secao_problema():
    out = [TituloSecao(1, "O problema"), Spacer(1, 10)]

    out.append(P("O objetivo é transformar um lote de fotos de recibos de táxi (<i>tickets</i> térmicos) "
                 f"em uma planilha com duas informações essenciais: o {B('número do ticket')} e o "
                 f"{B('valor total pago')}. Parece simples, mas os recibos reais têm uma combinação de "
                 "defeitos que atacam justamente os caracteres numéricos — exatamente os que importam."))

    out.append(P("Dificuldades reais encontradas nas imagens", S_H3))

    problemas = [
        ("Impressão térmica degradada",
         "Papel térmico perde tinta com o tempo e o calor. Dígitos ficam com traços finos e "
         "incompletos (o 8 'abre' e vira 0; o 6 vira 5)."),
        ("Texto cinza muito claro",
         "Alguns tickets saem com pouquíssima tinta. Uma binarização única (Otsu) apaga a linha "
         "inteira e o valor simplesmente desaparece."),
        ("Duas vias no mesmo papel",
         "O recibo tem a via do passageiro (VALOR DA CORRIDA + VALOR TOTAL A PAGAR) e a via da "
         "empresa, que repete \"VALOR R$\" mais abaixo. É preciso escolher o valor certo."),
        ("Ruído estrutural no lugar de letras",
         "Ícones e figuras à esquerda do texto fazem o OCR trocar as primeiras letras dos rótulos "
         "(\"PEDÁGIO\" → \"EDÁGIO\"), exigindo tolerância na identificação do rótulo."),
        ("Foto com inclinação, sombra e baixa resolução",
         "As imagens vêm de celular, com iluminação irregular e sem garantia de orientação ou "
         "distância fixa."),
        ("Confusão clássica de dígitos",
         "O → 0, l/I/| → 1, S → 5, B → 8, Z → 2. É a maior fonte de erro em números de ticket."),
    ]
    linhas = [[Paragraph("Desafio", S_CELL_H), Paragraph("Por que atrapalha", S_CELL_H)]]
    for tit, desc in problemas:
        linhas.append([Paragraph(f"<b>{tit}</b>", S_CELL), Paragraph(desc, S_CELL)])
    out.append(tabela(linhas, [5.2 * cm, LARGURA - 5.2 * cm]))
    out.append(Spacer(1, 10))

    out.append(caixa("Por que não usar um serviço de OCR na nuvem?",
                     ["<b>1. Privacidade.</b> Os tickets contêm dados de corridas; enviá-los a terceiros é um risco "
                      "desnecessário. <b>2. Custo.</b> O volume é pequeno e esporádico — API por chamada não compensa "
                      "frente a um processo local gratuito. <b>3. Portabilidade.</b> O programa precisa rodar em um "
                      "computador comum do escritório, sem internet, como um único .exe. O OCR local do "
                      "<b>Tesseract</b> resolve os três pontos de uma vez."], "ok"))
    out.append(Spacer(1, 12))

    out.append(P("O que o sistema entrega", S_H3))
    out.append(P("Uma janela onde o usuário escolhe a pasta de imagens, acompanha barra de progresso e "
                 "tempo restante, e ao final recebe uma planilha <b>.xlsx</b> pronta, com todos os valores "
                 "somados por uma fórmula viva. Nada de digitação manual, nada de dependência de internet."))
    out.append(PageBreak())
    return out


# =========================================================================== #
def secao_arquitetura():
    out = [TituloSecao(2, "Arquitetura da solução"), Spacer(1, 10)]

    out.append(P("O projeto foi dividido em <b>três camadas</b> com responsabilidades bem separadas. É essa "
                 "divisão que permite usar o mesmo núcleo de OCR pelo terminal (script) e pela interface "
                 "gráfica — e empacotar tudo em um único executável."))

    camadas = [
        ["Camada", "Arquivo", "Responsabilidade"],
        ["<b>Núcleo (motor)</b>", code("extrair_tickets.py"),
         "Pré-processamento, OCR, interpretação do texto, validação dos valores e geração do Excel. "
         "Não conhece interface gráfica e também funciona como CLI."],
        ["<b>Interface</b>", code("app_tickets.py"),
         "Janela Tkinter: pasta, barra de progresso, tempo decorrido/restante, tabela de resultados e "
         "botão para abrir a planilha. Consome o núcleo em threads e processos."],
        ["<b>Distribuição</b>", code("TicketsTaxis.spec") + "<br/>" + code("build_exe.ps1"),
         "Receita do PyInstaller: embute Python, OpenCV, numpy, Pillow, openpyxl, Tkinter e o próprio "
         "Tesseract no .exe, com console=False (sem janela preta)."],
    ]
    out.append(tabela(camadas, [2.6 * cm, 4.3 * cm, LARGURA - 6.9 * cm]))
    out.append(Spacer(1, 12))

    out.append(P("Fluxo de dados ponta a ponta", S_H2))
    out.append(fluxo_etapas([
        ("1", "Pasta de<br/>imagens"),
        ("2", "Pré-<br/>processamento"),
        ("3", "OCR<br/>Tesseract"),
        ("4", "Reconstrução<br/>de linhas"),
        ("5", "Regras de<br/>negócio"),
        ("6", "Excel<br/>+ fórmula"),
    ]))
    out.append(P("▲ Cada etapa é uma função independente do núcleo, o que facilita testar e ajustar uma "
                 "sem quebrar as demais.", S_CAP))
    out.append(Spacer(1, 8))

    out.append(P("Como as peças se encaixam", S_H2))
    out.append(P("A interface nunca faz OCR diretamente: ela chama "
                 f"{code('nucleo.processar_lote()')}, que distribui as imagens entre processos e devolve "
                 "objetos <b>Ticket</b> prontos via callback de progresso. Do outro lado, o modo CLI chama a "
                 "mesma função com progresso no terminal. Ou seja, existe <b>uma única implementação</b> da "
                 "lógica de leitura."))

    out.append(bloco_codigo("app_tickets.py — a interface só consome o núcleo",
                            """import extrair_tickets as nucleo

idioma = nucleo.idioma_disponivel()          # "por" se o tessdata existir
tickets = nucleo.processar_lote(
    arquivos,
    idioma,
    progresso=lambda feitos, total, caminho, ticket: ...
)
caminho_xlsx = nucleo.salvar_excel(tickets)  # planilha pronta"""))
    out.append(Spacer(1, 8))

    out.append(caixa("Princípio de projeto: o núcleo é puro",
                     "Toda a inteligência vive em `extrair_tickets.py`, que não importa Tkinter. Isso "
                     "significa que a lógica pode ser testada em scripts, reutilizada em outra interface "
                     "ou rodada em servidor sem mudar uma linha. A interface é apenas uma 'casca' fina.",
                     "info"))
    out.append(PageBreak())
    return out
