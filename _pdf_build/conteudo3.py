# -*- coding: utf-8 -*-
"""Seções 6 a 11 do documento."""
from __future__ import annotations

from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, CondPageBreak
from reportlab.lib.enums import TA_CENTER, TA_LEFT

from gerar_pdf import (
    AZUL, AZUL_CLARO, AZUL_ESCURO, AZUL_MEDIO, AZUL_PALIDO, AMARELO_CLARO, BRANCO,
    CINZA, CINZA_BORDA, LARGURA, LARANJA, VERDE, VERDE_CLARO, VERMELHO, VERMELHO_CLARO,
    TEXTO, S_BODY, S_BODY_SM, S_CAP, S_CELL, S_CELL_C, S_CELL_H, S_H2, S_H3,
    Barra, TituloSecao, caixa, tabela, bloco_codigo, figura, fluxo_etapas, st, P, B, code,
)


# =========================================================================== #
def secao_interface():
    out = [TituloSecao(6, "Interface e concorrência"), Spacer(1, 10)]

    out.append(P("A janela é construída com <b>Tkinter/ttk</b>. A regra de ouro de qualquer interface com "
                 "processamento pesado: <b>a thread da interface nunca faz o trabalho</b>. Se fizesse, a "
                 "janela congelaria e o Windows exibiria \"Não responde\"."))

    out.append(P("Arquitetura da janela", S_H2))
    blocos = [
        ["Camada", "Papel"],
        ["<b>Thread principal (Tk)</b>", "Desenha a janela, responde a cliques e lê uma <b>fila</b> a cada "
                                         "120 ms para atualizar progresso."],
        ["<b>Thread de trabalho</b>", "Chama o núcleo e envia mensagens para a fila "
                                      "('total', 'progresso', 'concluido', 'erro') — nunca toca em widgets."],
        ["<b>Processos</b>", "O núcleo distribui as imagens entre vários processos (um por núcleo, "
                             "reservando um para a interface), cada um rodando seu tesseract.exe."],
    ]
    out.append(tabela(blocos, [4.0 * cm, LARGURA - 4.0 * cm]))
    out.append(Spacer(1, 5))

    out.append(P("Comunicação interface ↔ trabalho", S_H2))
    out.append(P("A comunicação é unidirecional e não bloqueante: a thread de trabalho coloca eventos em "
                 f"uma {code('queue.Queue')}; a thread principal os consome com "
                 f"{code('after(120, self._drenar_fila)')}. Assim, nenhuma atualização de tela acontece "
                 "fora da thread do Tk — evitando travamentos e erros clássicos de Tkinter."))
    out.append(bloco_codigo("app_tickets.py — fila + after = interface fluida",
                            """def _trabalho(self):
    def enviar(tipo, **dados):
        self._fila.put((tipo, dados))       # worker só enfileira

    enviar("total", total=total)
    nucleo.processar_lote(
        arquivos, idioma, processos=...,
        progresso=lambda f, t, arq, tk: enviar("progresso", feitos=f, total=t, ticket=tk),
    )
    enviar("concluido")

# thread principal:
self.after(120, self._drenar_fila)          # consome e atualiza widgets"""))
    out.append(Spacer(1, 5))

    out.append(P("Estimativa de tempo restante", S_H2))
    out.append(P("Mostrar apenas \"x de y\" é pouco útil. O sistema mantém uma lista de marcos "
                 "(quantos concluídos e quando) e calcula a <b>média móvel</b> de segundos por ticket, "
                 "estimando o restante como \"3m 12s\". Como o tempo varia conforme a qualidade da "
                 "imagem, a média é recalculada a cada ticket — a previsão melhora com o andamento."))
    out.append(bloco_codigo("app_tickets.py — duração formatada e média móvel",
                            """def formatar_duracao(segundos: float) -> str:
    segundos = max(0, int(round(segundos)))
    horas, resto = divmod(segundos, 3600)
    minutos, segs = divmod(resto, 60)
    if horas:   return f"{horas}h {minutos:02d}m {segs:02d}s"
    if minutos: return f"{minutos:02d}m {segs:02d}s"
    return f"{segs}s\""""))
    out.append(Spacer(1, 4))

    out.append(P("Por que reservar um núcleo", S_H2))
    out.append(P("Usar 12 processos de OCR em 12 núcleos deixa a interface com fome de CPU e ela parece "
                 "travada. Reservar um núcleo "
                 f"({code('processos = os.cpu_count() - 1')}) mantém a janela responsiva sem afetar "
                 "o tempo total."))
    out.append(PageBreak())
    return out


# =========================================================================== #
def secao_excel():
    out = [TituloSecao(7, "Saída em Excel com openpyxl"), Spacer(1, 7)]

    out.append(P("A planilha não é só um despejo de linhas: ela foi desenhada para ser usada no dia a dia. "
                 "Tem título, total fixo no topo, pares (Nº | VALOR) lado a lado, linhas zebradas, valores "
                 "em negrito e destaque de cores para problemas."))

    out.append(Spacer(1, 6))
    out.append(caixa("Tkinter: por que não Qt ou uma interface web?",
                     "Tkinter vem <b>dentro</b> do Python: nenhum pacote extra, nenhuma DLL externa e um "
                     ".exe final bem menor do que com PyQt/PySide. Para uma barra de progresso, uma tabela "
                     "e três botões, o ttk entrega visual nativo do Windows sem complicar a distribuição.",
                     "ok"))
    out.append(Spacer(1, 8))

    out.append(P("Layout da planilha", S_H2))
    layout = [
        ["Elemento", "Posição", "Detalhe"],
        ["Título \"RELATÓRIO DE TICKETS\"", "linha 1 (mesclada)", "fundo azul-escuro, texto branco em negrito"],
        ["<b>VALOR TOTAL</b>", "linha 3 (congelada)", "gravado como <b>fórmula SUM</b>, não como número — "
                                                      "soma qualquer valor digitado depois"],
        ["Cabeçalho Nº DO TICKET | VALOR", "linha 5", "5 pares por linha, azul-claro"],
        ["Dados", "linhas 6+", "linhas alternadas em cinza-claro (zebra)"],
        ["Linhas extras em branco", "3 no fim", "espaço para lançamentos manuais"],
    ]
    out.append(tabela(layout, [5.0 * cm, 3.4 * cm, LARGURA - 8.4 * cm]))
    out.append(Spacer(1, 10))

    out.append(P("Um detalhe importante: fórmula, não valor", S_H2))
    out.append(P("Se o total fosse gravado como número fixo, qualquer lançamento manual ficaria de fora. "
                 "O openpyxl permite gravar a fórmula, e o Excel a recalcula ao abrir. Assim, o total "
                 "permanece sempre correto mesmo que o usuário acrescente valores à mão."))
    out.append(bloco_codigo("extrair_tickets.py — fórmula SUM dinâmica",
                            """parcelas = [f"{get_column_letter(col)}{LINHA_PRIMEIRA_DADO}:"
                        f"{get_column_letter(col)}{LINHA_ULTIMO_DADO}"
                        for col in colunas_valor]
cel_total = ws.cell(row=3, column=4, value="=SUM(" + ",".join(parcelas) + ")")
cel_total.number_format = 'R$ #,##0.00'
cel_total.font = Font(bold=True, size=12, color="FFFFFF")"""))
    out.append(Spacer(1, 8))

    out.append(P("Cores com significado", S_H2))
    cores = [
        ("#F8CBAD", "Laranja-suave", "\"Ñ IDENTIFICADO\" — OCR não conseguiu ler", "conferir à mão"),
        ("#FFE699", "Amarelo-suave", "Confiança média — leitura duvidosa", "revisar se possível"),
        ("#FFFFFF", "Sem cor", "Leitura confiável", "nenhuma"),
    ]
    linhas = []
    for cor_hex, nome, sig, acao in cores:
        amostra = Table([[""]], colWidths=[0.7 * cm], rowHeights=[0.42 * cm])
        amostra.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(cor_hex)),
            ("BOX", (0, 0), (-1, -1), 0.5, CINZA_BORDA),
        ]))
        linhas.append([amostra, Paragraph(f"<b>{nome}</b> <font color='#5A6672' size='7.5'>{cor_hex}</font>"
                                          f"<br/>{sig}", S_CELL), Paragraph(acao, S_CELL)])
    t = Table([[Paragraph("Amostra", S_CELL_H), Paragraph("Significado", S_CELL_H),
                Paragraph("Ação esperada", S_CELL_H)]] + linhas,
              colWidths=[1.4 * cm, 8.6 * cm, LARGURA - 10 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 0), (-1, -1), 0.4, CINZA_BORDA),
        ("BOX", (0, 0), (-1, -1), 0.6, CINZA_BORDA),
    ]))
    out.append(t)
    out.append(Spacer(1, 10))

    out.append(P("Por que openpyxl", S_H2))
    out.append(P("É a biblioteca pura Python mais completa para <b>escrever</b> .xlsx: suporta fontes, cores, "
                 "bordas, mesclagem, largura de coluna, congelamento de painéis, formato de moeda e, "
                 "principalmente, <b>fórmulas</b>. Alternativas como pandas exigiriam trazer a base toda e "
                 "não dariam o mesmo controle visual; o xlsxwriter é ótimo, mas só escreve (não lê), o que "
                 "limita extensões futuras."))
    out.append(CondPageBreak(0))
    return out


# =========================================================================== #
def secao_distribuicao():
    out = [TituloSecao(8, "Distribuição: um único .exe"), Spacer(1, 10)]

    out.append(P("O destinatário final não tem Python instalado e não deveria precisar instalar nada. A "
                 "entrega é um único arquivo <b>TicketsTaxis.exe</b> que contém o interpretador, as "
                 "bibliotecas, as fontes e o próprio Tesseract-OCR."))

    out.append(fluxo_etapas([
        ("1", "Código<br/>Python"),
        ("2", "Analysis<br/>(imports)"),
        ("3", "PYZ +<br/>datas"),
        ("4", "EXE<br/>(onefile)"),
        ("5", "Dist/<br/>TicketsTaxis.exe"),
    ], altura_caixa=22))
    out.append(Spacer(1, 10))

    out.append(P("O que a receita do PyInstaller faz", S_H2))
    itens_pi = [
        "<b>Embute o Tesseract inteiro</b>: copia tesseract.exe, todas as DLLs e apenas os idiomas "
        "usados (<i>por</i>, <i>eng</i>, <i>osd</i>) para dentro do pacote — mantendo o .exe enxuto.",
        "<b>Declara imports ocultos</b> (<i>cv2</i>, <i>numpy</i>, <i>openpyxl</i>, <i>PIL.Image</i>, "
        "<i>pytesseract</i>), que o PyInstaller não detecta sozinho.",
        "<b>Exclui o que não é usado</b> (matplotlib, pandas, scipy, PyQt, IPython), reduzindo o "
        "tamanho final.",
        "<b>console=False</b>: o .exe abre só a janela, sem console preto atrás.",
        "<b>Modo onefile</b>: tudo dentro de um arquivo; na execução o PyInstaller extrai em uma pasta "
        "temporária (_MEIPASS).",
    ]
    for it in itens_pi:
        out.append(P(f'<font color="#2E75B6">▪</font> {it}', S_BODY_SM))
    out.append(Spacer(1, 8))

    out.append(P("O detalhe que evita 'janelinhas piscando'", S_H2))
    out.append(P("No modo console=False, cada chamada ao tesseract.exe abriria uma janela de console por "
                 "um instante. A solução é informar ao sistema operacional que o subprocesso não deve criar "
                 "janela:"))
    out.append(bloco_codigo("extrair_tickets.py — executar o Tesseract sem console",
                            """_CRIAR_SEM_JANELA = 0x08000000  # CREATE_NO_WINDOW

def _tesseract_sem_janela(*args, **kwargs):
    info = subprocess.STARTUPINFO()
    info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    kwargs["startupinfo"] = info
    kwargs["creationflags"] = _CRIAR_SEM_JANELA
    return subprocess.Popen(*args, **kwargs)

pytesseract.pytesseract.subprocess = ...  # patch aplicado na inicialização"""))
    out.append(Spacer(1, 8))

    out.append(P("Build", S_H2))
    out.append(P("O build é disparado por "
                 f"{code('build_exe.ps1')} ou diretamente por {code('pyinstaller --noconfirm TicketsTaxis.spec')}. "
                 "O resultado fica em <b>dist\\TicketsTaxis.exe</b> e pode ser copiado para qualquer "
                 "computador Windows 10/11 sem instalação."))

    out.append(caixa("Portabilidade do Tesseract",
                     ["O .exe procura o tesseract.exe em quatro lugares, em ordem: ao lado do próprio .exe "
                      "(pasta Tesseract-OCR), dentro do pacote (_MEIPASS), e nas pastas padrão do sistema. "
                      "Assim, se você quiser trocar a versão do motor ou adicionar um idioma, basta colocar "
                      "uma pasta Tesseract-OCR ao lado do executável — sem reconstruí-lo.",
                      "Em máquinas Anaconda, o tessdata costuma ficar em "
                      + code("anaconda3\\share\\tessdata") + " — por isso o programa também testa essa "
                      "localização (pasta pai / share / tessdata)."], "info"))
    out.append(PageBreak())
    return out


# =========================================================================== #
def secao_desempenho():
    out = [TituloSecao(9, "Desempenho e resultados"), Spacer(1, 10)]

    out.append(P("As medições abaixo foram feitas neste computador, com o lote real de 44 imagens do "
                 "projeto. Elas servem de referência para dimensionar expectativas."))

    out.append(P("Ambiente de medição", S_H2))
    amb = [
        ["Item", "Valor"],
        ["Processador", "Intel Core i5-11400 @ 2,60 GHz (6 núcleos / 12 threads, até 4,4 GHz)"],
        ["Memória", "16 GB DDR4"],
        ["Sistema", "Windows 11 · Python 3.13.2 (conda-forge) · Tesseract 5.5.2 (Leptonica 1.87)"],
        ["Lote de teste", "44 imagens de tickets (jpg), densidade de texto variável"],
    ]
    out.append(tabela(amb, [3.4 * cm, LARGURA - 3.4 * cm]))
    out.append(Spacer(1, 10))

    out.append(P("Tempo medido (modo serial, 1 processo)", S_H2))
    tempos = [
        ["Amostra", "Tamanho da imagem ampliada", "Tempo", "Resultado"],
        ["N_apagado.jpg", "1700 × 914 px", "3,51 s", "Valor 76,00 (número não lido)"],
        ["V_apagado.jpg", "1700 × 914 px", "4,21 s", "Valor 101,00"],
        ["tickets001.jpg", "1700 × 914 px", "3,39 s", "Ticket 19725 · R$ 95,00"],
    ]
    out.append(tabela(tempos, [3.6 * cm, 5.2 * cm, 1.9 * cm, LARGURA - 10.7 * cm]))
    out.append(Spacer(1, 6))
    out.append(P(f"<b>Média: ~3,5 s por ticket em modo serial.</b> Em paralelo "
                 f"({code('os.cpu_count() - 1')} = 11 processos), o lote de 44 imagens leva "
                 "cerca de 44 × 3,5 ÷ 11 ≈ <b>~14 segundos</b>.", S_BODY_SM))
    out.append(Spacer(1, 10))

    out.append(P("Por que o paralelismo escala bem aqui", S_H2))
    out.append(P("O custo dominante é o tesseract.exe — um processo separado. Ao rodar vários ao mesmo "
                 "tempo, cada núcleo da CPU fica ocupado com seu próprio motor de OCR, sem disputa pelo "
                 "GIL do Python. Diferente de uma tarefa puramente Python em threads, aqui o ganho é "
                 "praticamente linear até o número de núcleos físicos."))
    dados = [["Processos", "Tempo estimado (lote de 44)", "Ganho"],
             ["1 (serial)", "~2 min 34 s", "1×"],
             ["4", "~39 s", "4×"],
             ["11 (padrão)", "~14 s", "~11×"]]
    out.append(tabela(dados, [3.6 * cm, 6.4 * cm, LARGURA - 10 * cm], alinhamentos=["b", "c", "c"]))
    out.append(Spacer(1, 10))

    out.append(caixa("Onde está o tempo: OCR, não Python",
                     "Das quatro binarizações, a chamada ao Tesseract consome a maior parte do tempo. Isso "
                     "explica duas escolhas: (1) ampliar a imagem até o máximo útil (1700 px) em vez de "
                     "gigantes; (2) paralelizar por processos. Otimizações no código Python isoladamente "
                     "teriam impacto pequeno.", "warn"))
    out.append(PageBreak())
    return out


# =========================================================================== #
def secao_limites():
    out = [TituloSecao(10, "Limitações e próximos passos"), Spacer(1, 10)]

    out.append(P("Ser honesto sobre os limites é parte da engenharia. Abaixo, o que o sistema ainda não "
                 "faz bem e caminhos concretos de melhoria."))

    limita = [
        ["Limitação atual", "Impacto", "Caminho de melhoria"],
        ["Ticket muito borrado / fora de foco", "Notifica \"Ñ IDENTIFICADO\"", "Desfoque de movimento "
         "é difícil de reverter; exigir nova foto com app de orientação"],
        ["Inclinação (skew) forte", "Linhas agrupadas de forma errada", "Detecção de ângulo (Hough "
         "ou minAreaRect) + deskew antes do OCR"],
        ["Outros layouts de recibo", "Rótulos/variantes fora das regex atuais", "Ajustar ou externalizar "
         "as regex em um arquivo de configuração"],
        ["Números muito degradados", "Troca de dígitos (2↔7, 6↔5)", "Pós-processamento de recorte "
         "por caractere com validação"],
        ["Só Windows testado", "Outros SOs não validados", "Empacotar com PyInstaller para Linux/macOS "
         "(o núcleo é multiplataforma)"],
    ]
    out.append(tabela(limita, [4.6 * cm, 4.6 * cm, LARGURA - 9.2 * cm]))
    out.append(Spacer(1, 12))

    out.append(P("Ideias de evolução, em ordem de custo-benefício", S_H2))
    ideias = [
        ("1", "Deskew automático",
         "Detecta a inclinação do papel e corrige antes de binarizar. Baixo esforço (poucas linhas de "
         "OpenCV) e grande ganho em fotos tortas."),
        ("2", "Detecção de região de interesse",
         "Localizar o bloco de texto com contornos antes do OCR reduz ruído e acelera."),
        ("3", "Aprendizado da correção de valor",
         "Guardar leituras \"corrigidas\" (antes/depois) para descobrir padrões de erro e automatizar "
         "mais correções com segurança."),
        ("4", "Modelo de OCR especializado",
         "Treinar/ajustar um modelo com tickets reais (fine-tune de Tesseract com <i>tesstrain</i> ou um "
         "CRNN leve) tende a ser o maior salto de acurácia."),
        ("5", "Dashboard de qualidade",
         "Registrar quantos tickets vieram de cada nível de confiança ao longo do tempo, para acompanhar "
         "a melhoria do processo e do lote de fotos."),
    ]
    for num, tit, desc in ideias:
        out.append(P(f'<font color="#2E75B6"><b>{num}.</b></font> <b>{tit}.</b> {desc}', S_BODY_SM))
    out.append(Spacer(1, 10))

    out.append(caixa("Como medir melhorias",
                     "Antes de mudar qualquer coisa, monte um pequeno conjunto dourado (20–30 tickets com "
                     "número e valor conferidos à mão) e um script que roda o núcleo sobre ele, comparando "
                     "resultado x esperado. Sem essa métrica, não há como saber se um ajuste melhorou ou "
                     "apenas mudou o erro de lugar.", "dica"))
    out.append(PageBreak())
    return out


# =========================================================================== #
def secao_glossario():
    out = [TituloSecao(11, "Glossário e consulta rápida"), Spacer(1, 10)]

    out.append(P("Termos técnicos", S_H2))
    gloss = [
        ["Termo", "Definição prática"],
        ["OCR", "<i>Optical Character Recognition</i> — conversão de imagem de texto em texto editável."],
        ["Binarização", "Transformar a imagem em preto e branco; separa texto do fundo."],
        ["Otsu", "Método que escolhe automaticamente o melhor limiar de binarização."],
        ["Abertura morfológica", "Erosão seguida de dilatação; remove pontos isolados e suaviza traços."],
        ["CLAHE", "Equalização de contraste adaptativa com limitação — realça texto apagado regionalmente."],
        ["Componente conexo", "Bloco de pixels vizinhos com a mesma cor; usado para estimar altura de texto."],
        ["IoU", "<i>Intersection over Union</i> — medida de sobreposição entre duas caixas (0 a 1)."],
        ["PSM", "<i>Page Segmentation Mode</i> — como o Tesseract interpreta a página; usamos 6 (bloco "
                "uniforme)."],
        ["tessdata", "Pasta com os modelos de idioma do Tesseract (ex.: por.traineddata)."],
        ["GIL", "<i>Global Interpreter Lock</i> — limita threads em Python; por isso usamos processos."],
        ["onefile", "Modo do PyInstaller que gera um único .exe, extraído em pasta temporária ao rodar."],
    ]
    out.append(tabela(gloss, [3.8 * cm, LARGURA - 3.8 * cm]))
    out.append(Spacer(1, 12))

    out.append(P("Comandos úteis", S_H2))
    out.append(bloco_codigo("Terminal (PowerShell) — do projeto",
                            """# rodar o script sobre a pasta padrão (./img_ticket)
python extrair_tickets.py

# pasta específica + gerar JSON além do terminal
python extrair_tickets.py "C:\\caminho\\das\\imagens" --json

# abrir com debug mostrando as linhas reconhecidas
python extrair_tickets.py --debug

# abrir a interface gráfica já com uma pasta
python app_tickets.py "C:\\caminho\\das\\imagens"

# gerar o executável (Windows)
pyinstaller --noconfirm TicketsTaxis.spec   # -> dist\\TicketsTaxis.exe"""))
    out.append(Spacer(1, 10))

    out.append(P("Mapa do código — onde está cada coisa", S_H2))
    mapa = [
        ["Função / símbolo", "Arquivo", "O que faz"],
        ["configurar_tesseract()", "extrair_tickets.py", "localiza tesseract.exe e define TESSDATA_PREFIX"],
        ["_escala_para_texto()", "extrair_tickets.py", "estima ampliação para texto de ~60 px"],
        ["binarizacoes()", "extrair_tickets.py", "gera as 4 variantes (otsu, open, soft, clahe_open)"],
        ["ocr_palavras()", "extrair_tickets.py", "roda OCR nas 4 e deduplica por IoU"],
        ["agrupar_linhas()", "extrair_tickets.py", "agrupa palavras em linhas por centro vertical"],
        ["extrair_numero()", "extrair_tickets.py", "lê e corrige o nº do ticket"],
        ["extrair_valores()", "extrair_tickets.py", "lê a coluna de valores de cada linha rotulada"],
        ["reler_valor()", "extrair_tickets.py", "relê o recorte com vários limiares e vota"],
        ["interpretar()", "extrair_tickets.py", "regras de negócio → objeto Ticket"],
        ["processar_lote()", "extrair_tickets.py", "processa em paralelo com progresso"],
        ["salvar_excel()", "extrair_tickets.py", "gera a planilha com fórmula SUM"],
        ["JanelaTickets._trabalho()", "app_tickets.py", "thread que chama o núcleo e enfileira eventos"],
        ["JanelaTickets._drenar_fila()", "app_tickets.py", "atualiza a interface a partir da fila"],
    ]
    out.append(tabela(mapa, [5.4 * cm, 3.2 * cm, LARGURA - 8.6 * cm]))
    out.append(Spacer(1, 14))

    out.append(Spacer(1, 4))
    out.append(caixa("Resumo em uma frase",
                     "Reaproveitamos um motor de OCR consagrado (Tesseract) e resolvemos a fragilidade dele "
                     "com <b>redundância inteligente</b>: quatro formas de olhar a imagem, consenso por "
                     "confiança, regras de domínio e a aritmética do próprio recibo como juiz — tudo local, "
                     "rápido e distribuído como um único .exe.", "ok"))

    out.append(Spacer(1, 18))
    out.append(Table([[""]], colWidths=[LARGURA], rowHeights=[1],
                     style=TableStyle([("LINEABOVE", (0, 0), (-1, -1), 0.8, AZUL_MEDIO)])))
    out.append(Spacer(1, 6))
    out.append(P('<font color="#5A6672" size="8">Documento gerado automaticamente a partir do código-fonte '
                 'e das imagens processadas pelo próprio sistema. '
                 'Versões: Python 3.13.2 · OpenCV 5.0.0 · Tesseract 5.5.2 · numpy 2.1.3 · pytesseract 0.3.13 · '
                 'openpyxl 3.1.5 · Pillow 12.2.0.</font>',
                 st("fim", fontSize=8, textColor=CINZA, alignment=TA_LEFT)))
    return out
