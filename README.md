# 🚕 Extrator de Tickets de Táxi — OCR + Excel automatizado

> **Da foto do recibo à planilha pronta em ~1 segundo por ticket — 100% offline, sem enviar nenhum dado para a nuvem.**

Automação que lê um lote de fotos de cupons fiscais de táxi, identifica **número do ticket** e **valores**, valida a aritmética do recibo, destaca o que não conseguiu ler e exporta tudo em uma planilha Excel pronta para o setor de faturamento.

[![Python](https://img.shields.io/badge/Python-3.13-3776AB?style=flat&logo=python&logoColor=white)](https://www.python.org/)
[![OpenCV](https://img.shields.io/badge/OpenCV-4.x-5C3EE8?style=flat&logo=opencv&logoColor=white)](https://opencv.org/)
[![Tesseract](https://img.shields.io/badge/Tesseract-5.x-FF6B35?style=flat)](https://github.com/tesseract-ocr/tesseract)
[![Tkinter](https://img.shields.io/badge/GUI-Tkinter-2C6E9B?style=flat)](https://docs.python.org/3/library/tkinter.html)
[![Excel](https://img.shields.io/badge/Excel-openpyxl-217346?style=flat&logo=microsoftexcel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![Licença](https://img.shields.io/badge/licença-MIT-green?style=flat)](LICENSE)

---

## 📊 Resultados medidos

| Métrica | Resultado |
|---|---|
| Imagens processadas (lote real) | **44** |
| Tempo total | **43,8 s** (~1,0 s por ticket) |
| Valores extraídos corretamente | **44 / 44** (100%) |
| Números do ticket identificados | **42 / 44** (95%)* |
| Total faturado conferido | **R$ 5.275,95** |
| Dependência de internet | **nenhuma** |

\* Os 2 casos restantes tinham tinta fisicamente degradada no papel; foram marcados explicitamente na planilha em vez de gerar um número inventado.

> Esses números não são estimativas: são o resultado da execução do lote real de 44 recibos incluído em `img_ticket/`.

---

## 🎯 O problema

O processo manual era: abrir cada foto de cupom térmico, ler o número do ticket e o valor, digitar em uma planilha. Com um lote grande, isso é **lento, repetitivo e propenso a erro de digitação** — e um erro em faturamento custa caro.

O desafio real não é "usar OCR". É que **cupons térmicos são um caso adverso**:

- papel amassado, torto e com sombras;
- tinta que desbota e some em partes do papel;
- fonte matricial estreita, onde `0` parece `O`, `1` parece `l` e `5` parece `S`;
- texto que se funde com o fundo.

Um OCR genérico entrega lixo nesse cenário. Foi preciso construir uma camada de tratamento em volta dele.

---

## 💡 Como funciona

```
 foto do ticket
      │
      ▼
[1] Leitura + correção de orientação
      │
      ▼
[2] Escala inteligente (texto até ~60 px)
      │
      ▼
[3] 4 binarizações diferentes
      │
      ▼
[4] OCR com caixas por palavra (Tesseract, --psm 6)
      │
      ▼
[5] Votação por confiança + IoU  ← deduplica e escolhe as melhores leituras
      │
      ▼
[6] Interpretação semântica + validação aritmética
      │
      ▼
 planilha .xlsx formatada
```

### A decisão que mudou o jogo: votar, não escolher

Nenhuma binarização é boa para tudo. Em vez de eleger uma, o sistema roda **quatro** e faz cada palavra aparecer em todas elas:

| Binarização | Para que serve |
|---|---|
| `otsu` | texto escuro em papel nítido |
| `open` | limpa dígitos degradados (abertura morfológica) |
| `soft` | preserva texto cinza muito claro |
| `clahe_open` | contraste local para tinta quase apagada |

Depois, as leituras são agrupadas por **sobreposição de caixas (IoU ≥ 0,35)** e vence a de maior confiança. O resultado prático: onde o Otsu falha, o `soft` acerta — e a deduplicação impede contar a mesma palavra duas vezes.

### Validação que não depende de OCR

A relação **`corrida + extras = total`** é uma conferência de integridade gratuita. Quando a conta não fecha, o sistema sabe que leu algo errado — e então **recorta só aquela região** e relê em 2 escalas × 5 limiares até a conta fechar. Correções desse tipo pegam erros que nenhuma regra sintática pegaria (ex.: `2` lido como `7`).

### Honestidade embutida no design

Quando não dá para ler, o sistema **não chuta**: registra `Ñ IDENTIFICADO`, pinta a célula em laranja na planilha e escreve o motivo nas observações. Preferir um campo vazio destacado a um número errado é uma escolha de engenharia, não uma limitação.

---

## ✨ Funcionalidades

### Núcleo de extração
- ✅ 4 binarizações + votação por confiança e IoU
- ✅ Correção de caracteres típicos de OCR (`O→0`, `l/I/|→1`, `S→5`, `B→8`, `Z→2`, `G→6`)
- ✅ Filtro de falsos positivos: telefone e código de autorização não viram valor
- ✅ Restrição da coluna de valores a 55% da largura (valores ficam à direita)
- ✅ Releitura por recorte quando a aritmética do ticket não fecha
- ✅ Extração de data/hora, forma de pagamento e origem/destino

### Interface desktop
- ✅ Janela Tkinter/ttk com seleção de pasta, barra de progresso e log
- ✅ Tempo restante estimado por **média móvel** (`"3m 12s"`)
- ✅ Interface nunca congela — thread de trabalho + `queue.Queue` + `after(120 ms)`
- ✅ Fila de eventos unidirecional: o worker nunca toca em widgets

### Performance e distribuição
- ✅ `ProcessPoolExecutor` para paralelismo real (o gargalo é o `tesseract.exe`, não o Python)
- ✅ Um núcleo reservado para a interface continuar responsiva
- ✅ Empacotamento em **um único `.exe`** com PyInstaller (usuário final não precisa de Python)
- ✅ Tesseract chamado sem abrir janelas de console (`CREATE_NO_WINDOW`)

---

## 🧠 Decisões de engenharia (o que eu aprendi)

| Decisão | Motivo |
|---|---|
| Processos em vez de threads | O gargalo é um processo externo; threads compartilhariam o GIL e não escalariam |
| IoU para "é a mesma palavra?" | Caixas do mesmo símbolo se sobrepõem muito; símbolos diferentes, não |
| Confiança mínima por tipo de campo | Número do ticket exige 55; valor aceita 45 (mais tolerante a tinta ruim) |
| Fórmula `SUM` no Excel, não número | Se o usuário digitar um valor manual, o total se recalcula sozinho |
| `console=False` + `STARTUPINFO` | Sem isso, cada chamada ao Tesseract pisca uma janela preta na tela |
| Valores ordenados por borda direita | Leituras coladas (`'124'` + `'00'`) se anexam no fim do valor |

---

## 🚀 Como rodar

### Requisitos
- Python 3.10+
- **Tesseract 5** com o idioma português (`por.traineddata`) — [download](https://github.com/UB-Mannheim/tesseract/wiki)

```powershell
pip install opencv-python numpy pytesseract openpyxl
```

### Linha de comando

```powershell
# processa a pasta e gera o Excel na área de trabalho
python extrair_tickets.py img_ticket

# com JSON de saída e depuração visual
python extrair_tickets.py img_ticket --json tickets.json --debug

# escolhendo o número de processos
python extrair_tickets.py img_ticket --processos 6
```

### Interface gráfica

```powershell
python app_tickets.py
```

### Gerar o executável

```powershell
.\build_exe.ps1
```

---

## 📁 Estrutura

```
projeto_ticket/
├── extrair_tickets.py      # núcleo: OCR, votação, interpretação, Excel
├── app_tickets.py          # interface Tkinter + orquestração de processos
├── build_exe.ps1           # empacotamento com PyInstaller
├── TicketsTaxis.spec       # especificação do executável
├── img_ticket/             # lote de imagens de entrada (44 recibos)
├── Documentacao_TicketsTaxis.pdf   # documentação técnica completa (17 páginas)
└── _pdf_build/             # gerador da documentação em PDF
```

---

## 📄 Documentação técnica

O arquivo [`Documentacao_TicketsTaxis.pdf`](Documentacao_TicketsTaxis.pdf) (17 páginas) detalha a arquitetura, o pipeline etapa por etapa, as bibliotecas comparadas, os resultados medidos e as limitações conhecidas — incluindo o trecho de código real de cada decisão.

---

## 🛠️ Stack

`Python` · `OpenCV` · `NumPy` · `Tesseract 5` · `pytesseract` · `Tkinter/ttk` · `openpyxl` · `PyInstaller` · `ReportLab`

---

## 🔭 Próximos passos

- [ ] Testes automatizados com imagens de referência e comparação de valores
- [ ] Suporte a PDF (recibos digitalizados em lote)
- [ ] Aprendizado dos limiares a partir de correções do usuário
- [ ] Versão CLI assíncrona para integração com outros sistemas

---

## 👤 Autor

Desenvolvido como projeto de automação real, com foco em **processamento de imagens, OCR aplicado e engenharia de software para uso final** — não apenas um notebook de demonstração, mas uma ferramenta empacotada que roda no computador de quem não tem Python instalado.

**Licença:** MIT
