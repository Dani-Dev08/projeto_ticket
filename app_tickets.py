#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Interface gráfica do extrator de tickets de táxi.

Permite escolher a pasta das imagens, acompanhar o andamento da leitura com
barra de progresso, tempo decorrido e tempo previsto para terminar, e ao final
abre/gera a planilha Excel com o resultado.

Uso:
    python app_tickets.py            # abre a janela
    python app_tickets.py <pasta>    # abre a janela já com a pasta selecionada
"""

from __future__ import annotations

import os
import queue
import statistics
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Garante que os imports do núcleo funcionem tanto rodando pelo .py quanto
# empacotado em .exe (PyInstaller onefile).
if getattr(sys, "frozen", False):
    sys.path.insert(0, str(Path(sys.executable).resolve().parent))

import extrair_tickets as nucleo


# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #

def formatar_duracao(segundos: float) -> str:
    """Formata segundos como '1h 02m 03s', '02m 03s' ou '03s'."""
    segundos = max(0, int(round(segundos)))
    horas, resto = divmod(segundos, 3600)
    minutos, segs = divmod(resto, 60)
    if horas:
        return f"{horas}h {minutos:02d}m {segs:02d}s"
    if minutos:
        return f"{minutos:02d}m {segs:02d}s"
    return f"{segs}s"


def abrir_arquivo(caminho: Path) -> None:
    """Abre o arquivo no aplicativo padrão do sistema (Excel, etc.)."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(caminho))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(caminho)])
        else:
            subprocess.Popen(["xdg-open", str(caminho)])
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Janela principal
# --------------------------------------------------------------------------- #

class JanelaTickets(tk.Tk):
    def __init__(self, pasta_inicial: str | None = None) -> None:
        super().__init__()

        self.title("Extrator de Tickets de Táxi")
        self.geometry("760x600")
        self.minsize(680, 540)

        self.pasta: Path | None = None
        self.arquivos: list[Path] = []
        self.tickets: list[tuple[str, str, str]] = []  # (arquivo, número, valor)

        self._fila: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._cancelar = threading.Event()
        self._inicio = 0.0
        self._duracao_media = 0.0  # segundos por ticket (aprendido no run)
        self._tempo_restante = 0.0
        self._marcas: list[tuple[int, float]] = []  # (concluídos, instante)
        self._ultimo_restante = 0.0  # último valor de "restante" exibido

        self._criar_widgets(pasta_inicial)
        self.after(120, self._drenar_fila)
        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)

    # ------------------------------------------------------------------ #
    # Construção da interface
    # ------------------------------------------------------------------ #

    def _criar_widgets(self, pasta_inicial: str | None) -> None:
        estilo = ttk.Style(self)
        try:
            estilo.theme_use("vista" if sys.platform.startswith("win") else "clam")
        except tk.TclError:
            pass

        container = ttk.Frame(self, padding=14)
        container.pack(fill="both", expand=True)

        # ---------------- Seleção de pasta ----------------
        quadro_pasta = ttk.LabelFrame(container, text=" Pasta das imagens ", padding=10)
        quadro_pasta.pack(fill="x")

        linha = ttk.Frame(quadro_pasta)
        linha.pack(fill="x")

        self.var_pasta = tk.StringVar(value=str(pasta_inicial or ""))
        self.entrada_pasta = ttk.Entry(linha, textvariable=self.var_pasta)
        self.entrada_pasta.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.botao_escolher = ttk.Button(
            linha, text="Escolher pasta...", command=self._escolher_pasta
        )
        self.botao_escolher.pack(side="left")

        self.rotulo_info_pasta = ttk.Label(
            quadro_pasta, text="Nenhuma pasta selecionada.", foreground="#666666"
        )
        self.rotulo_info_pasta.pack(anchor="w", pady=(8, 0))

        # ---------------- Ações ----------------
        quadro_acoes = ttk.Frame(container)
        quadro_acoes.pack(fill="x", pady=(12, 0))

        self.botao_iniciar = ttk.Button(
            quadro_acoes, text="▶  Iniciar leitura", command=self._iniciar
        )
        self.botao_iniciar.pack(side="left")

        self.botao_cancelar = ttk.Button(
            quadro_acoes, text="Cancelar", command=self._cancelar_leitura,
            state="disabled",
        )
        self.botao_cancelar.pack(side="left", padx=(8, 0))

        self.botao_excel = ttk.Button(
            quadro_acoes, text="Abrir planilha", command=self._abrir_planilha,
            state="disabled",
        )
        self.botao_excel.pack(side="right")

        self.caminho_excel: Path | None = None

        # ---------------- Progresso ----------------
        quadro_prog = ttk.LabelFrame(container, text=" Andamento ", padding=10)
        quadro_prog.pack(fill="x", pady=(12, 0))

        self.barra = ttk.Progressbar(quadro_prog, mode="determinate", maximum=100)
        self.barra.pack(fill="x")

        self.var_status = tk.StringVar(value="Aguardando início.")
        ttk.Label(quadro_prog, textvariable=self.var_status).pack(anchor="w", pady=(8, 0))

        grade = ttk.Frame(quadro_prog)
        grade.pack(fill="x", pady=(8, 0))
        for col in (1, 3):
            grade.columnconfigure(col, weight=1)

        ttk.Label(grade, text="Decorrido:").grid(row=0, column=0, sticky="w")
        self.var_decorrido = tk.StringVar(value="00s")
        ttk.Label(grade, textvariable=self.var_decorrido).grid(row=0, column=1, sticky="w")

        ttk.Label(grade, text="Restante estimado:").grid(row=0, column=2, sticky="w", padx=(20, 0))
        self.var_restante = tk.StringVar(value="--")
        ttk.Label(grade, textvariable=self.var_restante).grid(row=0, column=3, sticky="w")

        self.rotulo_atual = ttk.Label(
            quadro_prog, text="", foreground="#666666", anchor="w"
        )
        self.rotulo_atual.pack(fill="x", pady=(8, 0))

        # ---------------- Resultados ----------------
        quadro_res = ttk.LabelFrame(container, text=" Resultado ", padding=10)
        quadro_res.pack(fill="both", expand=True, pady=(12, 0))

        colunas = ("arquivo", "numero", "valor")
        self.tabela = ttk.Treeview(
            quadro_res, columns=colunas, show="headings", height=8
        )
        self.tabela.heading("arquivo", text="Arquivo")
        self.tabela.heading("numero", text="Nº do ticket")
        self.tabela.heading("valor", text="Valor")
        self.tabela.column("arquivo", width=260, anchor="w")
        self.tabela.column("numero", width=160, anchor="center")
        self.tabela.column("valor", width=130, anchor="e")

        rolagem = ttk.Scrollbar(quadro_res, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=rolagem.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        rolagem.pack(side="right", fill="y")

        # ---------------- Rodapé ----------------
        self.var_rodape = tk.StringVar(value="")
        ttk.Label(
            container, textvariable=self.var_rodape, font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", pady=(10, 0))

        if pasta_inicial:
            self.var_pasta.set(str(pasta_inicial))
            self._carregar_pasta(Path(pasta_inicial), silencioso=True)

    # ------------------------------------------------------------------ #
    # Seleção de pasta
    # ------------------------------------------------------------------ #

    def _escolher_pasta(self) -> None:
        inicial = self.var_pasta.get().strip() or str(Path.home())
        escolhida = filedialog.askdirectory(
            title="Escolha a pasta com as imagens dos tickets",
            initialdir=inicial,
            mustexist=True,
        )
        if not escolhida:
            return
        self.var_pasta.set(escolhida)
        self._carregar_pasta(Path(escolhida))

    def _carregar_pasta(self, pasta: Path, silencioso: bool = False) -> None:
        if not pasta.is_dir():
            if not silencioso:
                messagebox.showerror("Pasta inválida", f"A pasta não existe:\n{pasta}")
            return

        try:
            arquivos = sorted(
                p for p in pasta.iterdir()
                if p.is_file() and p.suffix.lower() in nucleo.EXTENSOES
            )
        except OSError as erro:
            if not silencioso:
                messagebox.showerror("Erro ao ler a pasta", str(erro))
            return

        self.pasta = pasta
        self.arquivos = arquivos

        if arquivos:
            self.rotulo_info_pasta.configure(
                text=f"{len(arquivos)} imagem(ns) encontrada(s) nesta pasta.",
                foreground="#1F6F3D",
            )
            self.botao_iniciar.configure(state="normal")
        else:
            self.rotulo_info_pasta.configure(
                text="Nenhuma imagem encontrada nesta pasta "
                     "(use .jpg, .png, .tif, .bmp ou .webp).",
                foreground="#B23B3B",
            )
            self.botao_iniciar.configure(state="disabled")

    # ------------------------------------------------------------------ #
    # Execução
    # ------------------------------------------------------------------ #

    def _iniciar(self) -> None:
        if self._thread and self._thread.is_alive():
            return

        pasta_txt = self.var_pasta.get().strip()
        if not pasta_txt:
            messagebox.showwarning("Pasta não escolhida", "Escolha a pasta das imagens.")
            return

        pasta = Path(pasta_txt)
        self._carregar_pasta(pasta, silencioso=True)
        if not self.arquivos:
            messagebox.showwarning(
                "Sem imagens",
                "A pasta escolhida não contém imagens suportadas.",
            )
            return

        # Reset da tela.
        self.tabela.delete(*self.tabela.get_children())
        self.tickets.clear()
        self.barra["value"] = 0
        self.var_status.set("Preparando...")
        self.var_decorrido.set("00s")
        self.var_restante.set("--")
        self.var_rodape.set("")
        self.caminho_excel = None
        self.botao_excel.configure(state="disabled")
        self.botao_iniciar.configure(state="disabled")
        self.botao_escolher.configure(state="disabled")
        self.botao_cancelar.configure(state="normal")

        self._cancelar.clear()
        self._inicio = time.perf_counter()
        self._marcas.clear()
        self._ultimo_restante = 0.0
        self._thread = threading.Thread(target=self._trabalho, daemon=True)
        self._thread.start()

    def _trabalho(self) -> None:
        """Roda em thread separada: OCR é pesado e não pode travar a janela."""
        def enviar(tipo: str, **dados) -> None:
            self._fila.put((tipo, dados))

        try:
            arquivos = list(self.arquivos)
            total = len(arquivos)
            enviar("total", total=total)

            idioma = nucleo.idioma_disponivel()
            enviar("idioma", idioma=idioma)

            # Reserva 1 núcleo para a interface continuar responsiva.
            processos = max(1, min((os.cpu_count() or 2) - 1, total))

            iniciados = 0
            duracoes: list[float] = []
            ultimo_inicio = time.perf_counter()

            def ao_progresso(concluidos: int, total_: int, caminho: Path, ticket) -> None:
                nonlocal ultimo_inicio
                agora = time.perf_counter()
                duracoes.append(agora - ultimo_inicio)
                ultimo_inicio = agora
                enviar(
                    "progresso",
                    concluidos=concluidos,
                    total=total_,
                    arquivo=caminho.name,
                    numero=ticket.numero,
                    valor=ticket.valor,
                )

            tickets = nucleo.processar_lote(
                arquivos,
                idioma,
                debug=False,
                processos=processos,
                progresso=ao_progresso,
            )

            if self._cancelar.is_set():
                enviar("cancelado")
                return

            enviar("salvando", total=total)
            caminho_excel: Path | None = None
            erro_excel = ""
            try:
                caminho_excel = nucleo.salvar_excel(tickets)
            except Exception as erro:  # pragma: no cover
                erro_excel = str(erro)

            enviar("concluido", tickets=tickets, excel=caminho_excel, erro_excel=erro_excel)
        except Exception:
            enviar("falha", detalhe=traceback.format_exc())

    def _cancelar_leitura(self) -> None:
        self._cancelar.set()
        self.var_status.set("Cancelando... aguardando os tickets em andamento.")

    def _drenar_fila(self) -> None:
        """Processa as mensagens da thread de trabalho na thread da interface."""
        try:
            while True:
                tipo, dados = self._fila.get_nowait()
                self._tratar(tipo, dados)
        except queue.Empty:
            pass
        self.after(120, self._drenar_fila)

    def _tratar(self, tipo: str, dados: dict) -> None:
        if tipo == "total":
            self._total = dados["total"]
            self.var_status.set(f"Preparando OCR ({dados['total']} imagens)...")

        elif tipo == "idioma":
            pass

        elif tipo == "progresso":
            concluidos = dados["concluidos"]
            total = dados["total"]
            self.barra["value"] = concluidos / total * 100

            agora = time.perf_counter()
            decorrido = agora - self._inicio
            self.var_decorrido.set(formatar_duracao(decorrido))

            # Tempo restante: 45% média geral + 55% mediana recente.
            # A mediana absorve picos (tickets difíceis) que faziam a
            # estimativa "pular"; o prazo mínimo entre atualizações evita
            # que o texto fique trocando por 1s de diferença.
            if concluidos:
                self._marcas.append((agora, concluidos))
                if len(self._marcas) > 15:
                    self._marcas.pop(0)

                if concluidos < total:
                    base = self._marcas[0]
                    intervalo_tempo = agora - base[0]
                    intervalo_tickets = concluidos - base[1]

                    if intervalo_tickets >= 3 and intervalo_tempo > 0:
                        taxa_recente = intervalo_tickets / intervalo_tempo
                        # mediana das velocidades entre marcas consecutivas
                        velocidades = []
                        for (t0, c0), (t1, c1) in zip(self._marcas,
                                                      self._marcas[1:]):
                            if c1 > c0 and t1 > t0:
                                velocidades.append((c1 - c0) / (t1 - t0))
                        taxa_mediana = (statistics.median(velocidades)
                                        if velocidades else taxa_recente)
                        taxa = 0.45 * (concluidos / decorrido) + \
                            0.30 * taxa_recente + 0.25 * taxa_mediana
                    else:
                        taxa = concluidos / decorrido

                    restante = (total - concluidos) / taxa

                    # Só redesenha quando muda de verdade (evita cintilação).
                    if abs(restante - self._ultimo_restante) >= 2 or \
                            self._ultimo_restante == 0:
                        self.var_restante.set(formatar_duracao(restante))
                        self._ultimo_restante = restante
                else:
                    self.var_restante.set("0s")

            numero = dados["numero"] or nucleo.NAO_IDENTIFICADO
            if dados["valor"] is None:
                valor_txt = nucleo.NAO_IDENTIFICADO
            else:
                valor_txt = nucleo.moeda(dados["valor"])

            self.tabela.insert(
                "", "end",
                values=(dados["arquivo"], numero, valor_txt),
            )
            self.tabela.yview_moveto(1.0)
            self.var_status.set(
                f"Lendo tickets... {concluidos} de {total} "
                f"({concluidos / total * 100:.0f}%)"
            )
            self.rotulo_atual.configure(text=f"Último arquivo: {dados['arquivo']}")

        elif tipo == "salvando":
            self.barra["value"] = 100
            self.var_status.set("Gerando planilha Excel...")

        elif tipo == "concluido":
            self._finalizar(dados)

        elif tipo == "cancelado":
            self._restaurar_botoes()
            self.var_status.set("Leitura cancelada pelo usuário.")
            self.var_restante.set("--")

        elif tipo == "falha":
            self._restaurar_botoes()
            self.var_status.set("Ocorreu um erro.")
            messagebox.showerror(
                "Erro inesperado",
                "Não foi possível concluir a leitura.\n\n"
                "A planilha não foi gerada. Detalhes técnicos abaixo:\n\n"
                + dados.get("detalhe", "")[-1500:],
            )

    def _finalizar(self, dados: dict) -> None:
        self._restaurar_botoes()
        tickets = dados["tickets"]
        excel = dados["excel"]
        erro_excel = dados.get("erro_excel") or ""

        decorrido = time.perf_counter() - self._inicio
        self.barra["value"] = 100
        self.var_decorrido.set(formatar_duracao(decorrido))
        self.var_restante.set("0s")

        com_valor = [t for t in tickets if t.valor is not None]
        total_geral = sum(t.valor for t in com_valor if t.valor is not None)
        sem_valor = [t for t in tickets if t.valor is None]

        self.var_status.set(
            f"Concluído: {len(tickets)} tickets em {formatar_duracao(decorrido)}."
        )
        self.var_rodape.set(
            f"Total: {nucleo.moeda(round(total_geral, 2))}"
            + (f"   |   {len(sem_valor)} ticket(s) sem valor lido" if sem_valor else "")
        )
        self.rotulo_atual.configure(text="")

        if excel:
            self.caminho_excel = Path(excel)
            self.botao_excel.configure(state="normal")
            sem_valor_txt = ""
            if sem_valor:
                nomes = ", ".join(t.arquivo for t in sem_valor[:12])
                if len(sem_valor) > 12:
                    nomes += ", ..."
                sem_valor_txt = (
                    f"\n\nATENÇÃO: {len(sem_valor)} ticket(s) sem valor lido "
                    f"(em destaque laranja na planilha):\n{nomes}"
                )
            messagebox.showinfo(
                "Leitura concluída",
                f"{len(tickets)} tickets processados.\n"
                f"Tempo: {formatar_duracao(decorrido)}\n\n"
                f"Planilha salva em:\n{excel}{sem_valor_txt}",
            )
            abrir_arquivo(self.caminho_excel)
        elif erro_excel:
            messagebox.showerror(
                "Erro ao gerar a planilha",
                f"Os tickets foram lidos, mas a planilha não pôde ser salva:\n\n{erro_excel}",
            )

    def _restaurar_botoes(self) -> None:
        self.botao_iniciar.configure(state="normal")
        self.botao_escolher.configure(state="normal")
        self.botao_cancelar.configure(state="disabled")

    def _abrir_planilha(self) -> None:
        if self.caminho_excel and self.caminho_excel.exists():
            abrir_arquivo(self.caminho_excel)
        else:
            messagebox.showinfo(
                "Planilha não encontrada",
                "A planilha ainda não foi gerada ou foi movida.",
            )

    def _ao_fechar(self) -> None:
        if self._thread and self._thread.is_alive():
            if not messagebox.askyesno(
                "Sair durante a leitura",
                "A leitura está em andamento. Deseja realmente sair e cancelar?",
            ):
                return
            self._cancelar.set()
        self.destroy()


# --------------------------------------------------------------------------- #
# Ponto de entrada
# --------------------------------------------------------------------------- #

def _pasta_argumento() -> str | None:
    for arg in sys.argv[1:]:
        if not arg.startswith("-") and Path(arg).is_dir():
            return str(Path(arg).resolve())
    return None


def main() -> int:
    if sys.platform.startswith("win"):
        # Evita desfoque em telas com escala (DPI alto).
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    try:
        janela = JanelaTickets(_pasta_argumento())
    except Exception as erro:  # pragma: no cover
        try:
            import tkinter as _tk
            raiz = _tk.Tk()
            raiz.withdraw()
            messagebox.showerror("Erro ao iniciar", str(erro))
            raiz.destroy()
        except Exception:
            print(erro, file=sys.stderr)
        return 1

    janela.mainloop()
    return 0


if __name__ == "__main__":
    from multiprocessing import freeze_support
    freeze_support()
    raise SystemExit(main())
