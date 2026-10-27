# -*- coding: utf-8 -*-
"""
Janela da automação de capas.

Cole os alunos direto do Excel (NOME | CURSO | PLANO), confira na tabela
e clique em Validar ou Gerar capas. A janela grava uma planilha temporária
e roda AUTOMAÇAO-PASTA.py em um processo separado.
"""

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import unicodedata
from datetime import datetime
from pathlib import Path

import openpyxl
from PIL import Image, ImageTk


# ============================================================
# CAMINHOS
# ============================================================

PASTA = Path(__file__).resolve().parent

SCRIPT = PASTA / "AUTOMAÇAO-PASTA.py"

CONFIG = PASTA / "automacao_capas_config.json"

RELATORIO = PASTA / "resultado_capas.csv"

RELATORIO_PREFLIGHT = PASTA / "preflight_capas.csv"
SCRIPT_RECORTE = PASTA / "recortar_fotos.py"
RELATORIO_RECORTE = PASTA / "recorte_fotos.csv"
PREVIEW_RECORTE = Path(tempfile.gettempdir()) / "recorte_preview"

PASTA_TEMP = Path(tempfile.gettempdir()) / "capas_ui"

PLANILHA_TEMP = PASTA_TEMP / "planilha_colada.xlsx"

PLANILHA_LINK = PASTA_TEMP / "planilha_do_link.xlsx"

ULTIMA_COLAGEM = PASTA_TEMP / "ultima_colagem.txt"

PARAR_ARQUIVO = PASTA_TEMP / "parar_capas.flag"

TIPOS_SAIDA = ("JPG e PSD", "Só JPG", "Só PSD")
CURSOS_POR_FOTO_FILE = PASTA / "cursos_por_foto.json"

# Cadastro de uma sessão: ao fechar a janela vai para historico_sessoes e a próxima começa limpa
HISTORICO_SESSOES = PASTA / "historico_sessoes"
SESSOES_GUARDADAS = 30
ARQUIVOS_SESSAO = (
    CURSOS_POR_FOTO_FILE,
    PASTA / "recorte_manual.json",
    PASTA / "ajustes_nome.json",
    RELATORIO,
    RELATORIO_PREFLIGHT,
    RELATORIO_RECORTE,
    ULTIMA_COLAGEM,
    PLANILHA_TEMP,
    PLANILHA_LINK,
)


def limpar_cadastro_sessao():
    """Move o cadastro da sessão para historico_sessoes/<data_hora>. Devolve os arquivos que não saíram."""

    existentes = [a for a in ARQUIVOS_SESSAO if a.is_file()]
    if not existentes:
        return []
    destino = HISTORICO_SESSOES / datetime.now().strftime("%Y-%m-%d_%H%M%S")
    destino.mkdir(parents=True, exist_ok=True)
    presos = []
    for arquivo in existentes:
        try:
            shutil.move(str(arquivo), str(destino / arquivo.name))
        except OSError:
            presos.append(arquivo)
    antigas = sorted(p for p in HISTORICO_SESSOES.iterdir() if p.is_dir())
    for pasta in antigas[:-SESSOES_GUARDADAS]:
        shutil.rmtree(pasta, ignore_errors=True)
    return presos

# Paleta do tema "Cursor Dark Midnight"
COR_FUNDO = "#191c22"
COR_SUPERFICIE = "#1e2127"
COR_CAMPO = "#23272f"
COR_BORDA = "#2c313b"
COR_BOTAO = "#2a2f3a"
COR_BOTAO_HOVER = "#343a47"
COR_SELECAO = "#2f3646"
COR_TEXTO = "#d8dee9"
COR_TEXTO_FORTE = "#eceff4"
COR_SECUNDARIO = "#7b88a1"
COR_APAGADO = "#4b5163"
COR_DESTAQUE = "#88c0d0"
COR_DESTAQUE_HOVER = "#98d5e7"
COR_VERDE = "#a3be8c"
COR_AMARELO = "#ebcb8b"
COR_LARANJA = "#d08770"
COR_VERMELHO = "#bf616a"
COR_AZUL = "#81a1c1"


def chave_caminho_foto(caminho):

    return os.path.normcase(os.path.abspath(os.path.normpath(str(caminho))))


# ============================================================
# LEITURA DA COLAGEM
# ============================================================

def chave(valor):

    texto = re.sub(r"\s+", " ", str(valor or "").strip()).upper()

    return "".join(
        c
        for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


# Mesmos sinônimos de encontrar_cabecalho no script
CABECALHO_NOME = {
    "NOME", "NOME DO ALUNO", "NOME DO FORMANDO",
    "NOME COMPLETO", "ALUNO", "FORMANDO",
}

CABECALHO_CURSO = {
    "CURSO", "FORMACAO",
}

CABECALHO_PLANO = {
    "PLANO", "ALBUM", "PRODUTO", "TIPO",
    "TIPO DE ALBUM",
}


def _detectar_cabecalho(celulas):
    """Retorna {"nome": i, "curso": i, "plano": i} se a linha for cabeçalho."""

    colunas = {}

    for i, valor in enumerate(celulas):

        k = chave(valor)

        if k in CABECALHO_NOME:
            colunas.setdefault("nome", i)
        elif k in CABECALHO_CURSO:
            colunas.setdefault("curso", i)
        elif k in CABECALHO_PLANO:
            colunas.setdefault("plano", i)

    if "nome" in colunas and "plano" in colunas:
        return colunas

    return None


def formatar_nome(nome):
    """Só a primeira letra de cada palavra maiúscula: "MARIA DA SILVA" -> "Maria Da Silva"."""
    texto = re.sub(r"\s+", " ", str(nome or "")).strip()
    return re.sub(r"[^\W\d_]+", lambda m: m.group(0)[0].upper() + m.group(0)[1:].lower(), texto)


def interpretar_colagem(texto):
    """
    Converte o texto colado do Excel em lista de alunos.

    Cada aluno: {"linha": n, "nome": ..., "curso": ..., "plano": ...}
    "linha" é o número que o script vai usar (cabeçalho na 1, alunos a partir da 2).
    """

    linhas = [
        l for l in texto.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if l.strip()
    ]

    if not linhas:
        return []

    colunas = {"nome": 0, "curso": 1, "plano": 2}

    primeira = linhas[0].split("\t")
    cabecalho = _detectar_cabecalho(primeira)

    if cabecalho:
        colunas = cabecalho
        linhas = linhas[1:]

    def pegar(celulas, campo):
        i = colunas.get(campo)
        if i is None or i >= len(celulas):
            return ""
        return re.sub(r"\s+", " ", celulas[i].strip())

    alunos = []

    for numero, linha in enumerate(linhas, start=2):

        celulas = linha.split("\t")

        alunos.append({
            "linha": numero,
            "nome": formatar_nome(pegar(celulas, "nome")),
            "curso": pegar(celulas, "curso"),
            "plano": pegar(celulas, "plano"),
        })

    return alunos


def gravar_planilha_temporaria(alunos, destino=PLANILHA_TEMP):

    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ALUNOS"

    ws.append(["NOME", "CURSO", "PLANO"])

    for aluno in alunos:
        ws.append([aluno["nome"], aluno["curso"], aluno["plano"]])

    wb.save(destino)

    return destino


def parse_faixa_linhas(texto):
    """Aceita '10', '10-20', '10,12,15-18'. Lança ValueError se inválido."""

    texto = (texto or "").strip()

    if not texto:
        return None

    for parte in texto.split(","):
        parte = parte.strip()
        if not parte:
            continue
        if "-" in parte:
            a, b = parte.split("-", 1)
            int(a.strip())
            int(b.strip())
        else:
            int(parte)

    return texto


# ============================================================
# EXECUÇÃO DO SCRIPT
# ============================================================

def python_console():
    """python.exe ao lado do pythonw.exe atual (a saída precisa ser capturada)."""

    exe = Path(sys.executable)

    if exe.name.lower() == "pythonw.exe":
        candidato = exe.with_name("python.exe")
        if candidato.is_file():
            return str(candidato)

    return str(exe)


RE_RESUMO = re.compile(
    r"OK:\s*(\d+)\s*\|\s*SKIP:\s*(\d+)\s*\|\s*ERRO:\s*(\d+)"
)

RE_RESUMO_PREFLIGHT = re.compile(
    r"OK:\s*(\d+)\s*\|\s*AVISO:\s*(\d+)\s*\|\s*ERRO:\s*(\d+)"
)


# ============================================================
# JANELA
# ============================================================

_recorte = None


def modulo_recorte():
    """recortar_fotos.py (config, links do Google, recorte)."""

    global _recorte

    if _recorte is None:
        if str(PASTA) not in sys.path:
            sys.path.insert(0, str(PASTA))
        import recortar_fotos
        _recorte = recortar_fotos

    return _recorte


_automacao = None


def modulo_automacao():
    """AUTOMAÇAO-PASTA.py (cursos e cores padrão)."""

    global _automacao

    if _automacao is None:
        import importlib.util
        spec = importlib.util.spec_from_file_location("automacao_capas", SCRIPT)
        modulo = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(modulo)
        _automacao = modulo

    # relê o config a cada uso: cores trocadas na aba "Cores dos cursos" valem na hora
    _automacao.carregar_config()
    return _automacao


def passos_drive():
    """Baixa as fotos novas do Drive antes de rodar, quando as fotos vêm de um link."""

    dados = modulo_recorte().ler_config_bruto()

    if (dados.get("fotos_origem") == "drive" and dados.get("photos_link")
            and dados.get("drive_sincronizar_antes", True)):
        return [(SCRIPT_RECORTE, ["--sincronizar-drive"])]

    return []


def planilha_dir_config():

    try:
        dados = json.loads(CONFIG.read_text(encoding="utf-8"))
        if dados.get("planilha_dir"):
            return dados["planilha_dir"]
    except Exception:
        pass

    return str(PASTA)


class JanelaCapas:

    def __init__(self, root):

        import tkinter as tk
        from tkinter import ttk

        self.tk = tk
        self.ttk = ttk
        self.root = root
        self.processo = None
        self.relatorio_atual = None
        self.resumo = None
        self.alunos = []

        root.title("Automação de Capas")
        root.minsize(820, 620)
        try:
            root.state("zoomed")
        except tk.TclError:
            root.geometry("1280x800")

        self._configurar_estilo()

        self.origem = tk.StringVar(value="fotos")
        self.planilha = tk.StringVar()
        self.curso_todos = tk.StringVar()
        self.plano_todos = tk.StringVar()
        self.plano_todos_aplicado = ""
        self.aluno = tk.StringVar()
        self.plano = tk.StringVar()
        self.linhas = tk.StringVar()
        self.lote = tk.StringVar()
        self.status_filtro = tk.StringVar()
        self.so_erros = tk.BooleanVar(value=False)
        self.refazer = tk.BooleanVar(value=False)
        self.berlim = tk.StringVar(value="Ambos")
        self.tipo_saida = tk.StringVar(value=self._tipo_saida_config())
        self.parando = False
        self.status = tk.StringVar(
            value="Sem planilha: as capas saem das fotos da pasta. Cole alunos ou escolha uma planilha se quiser."
        )

        self.fila = []

        self._montar()
        self._carregar_ultima_colagem()
        self.planilha.trace_add("write", lambda *a: self._planilha_mudou())
        root.protocol("WM_DELETE_WINDOW", self._fechar)

    def _fechar(self):

        from tkinter import messagebox

        if self.processo is not None:
            if not messagebox.askyesno(
                "Automação de Capas",
                "Ainda tem coisa rodando.\n\nParar agora e fechar?\n"
                "(o cadastro desta sessão vai para historico_sessoes)",
            ):
                return
            self.fila = []
            subprocess.run(
                ["taskkill", "/PID", str(self.processo.pid), "/T", "/F"],
                capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            try:
                self.processo.wait(timeout=10)
            except Exception:
                pass

        try:
            presos = limpar_cadastro_sessao()
        except Exception as e:
            presos = [str(e)]
        if presos and not messagebox.askokcancel(
            "Automação de Capas",
            "Estes arquivos do cadastro não puderam ser limpos (abertos no Excel?):\n\n"
            + "\n".join(str(p) for p in presos)
            + "\n\nOK = fechar mesmo assim | Cancelar = voltar, fechar o Excel e tentar de novo",
        ):
            return
        self.root.destroy()

    @staticmethod
    def _tipo_saida_config():

        try:
            dados = json.loads(CONFIG.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return TIPOS_SAIDA[0]
        jpg, psd = dados.get("export_jpg", True), dados.get("export_psd", True)
        if jpg and not psd:
            return "Só JPG"
        if psd and not jpg:
            return "Só PSD"
        return TIPOS_SAIDA[0]

    def _configurar_estilo(self):

        tk = self.tk
        ttk = self.ttk
        estilo = ttk.Style(self.root)

        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass

        fonte = ("Segoe UI", 10)
        root = self.root
        root.configure(background=COR_SUPERFICIE)
        for chave_opcao, valor in (
            ("*Text.background", COR_FUNDO),
            ("*Text.foreground", COR_TEXTO),
            ("*Text.insertBackground", COR_DESTAQUE),
            ("*Text.selectBackground", COR_SELECAO),
            ("*Text.selectForeground", COR_TEXTO_FORTE),
            ("*Text.relief", "flat"),
            ("*Text.borderWidth", 0),
            ("*Text.highlightThickness", 1),
            ("*Text.highlightBackground", COR_BORDA),
            ("*Text.highlightColor", COR_APAGADO),
            ("*Text.padX", 8),
            ("*Text.padY", 6),
            ("*Text.font", ("Cascadia Mono", 9)),
            ("*Toplevel.background", COR_SUPERFICIE),
            ("*Canvas.highlightThickness", 0),
            ("*TCombobox*Listbox.background", COR_CAMPO),
            ("*TCombobox*Listbox.foreground", COR_TEXTO),
            ("*TCombobox*Listbox.selectBackground", COR_SELECAO),
            ("*TCombobox*Listbox.selectForeground", COR_TEXTO_FORTE),
            ("*TCombobox*Listbox.borderWidth", 0),
            ("*TCombobox*Listbox.font", fonte),
            ("*Menu.background", COR_CAMPO),
            ("*Menu.foreground", COR_TEXTO),
            ("*Menu.activeBackground", COR_SELECAO),
            ("*Menu.relief", "flat"),
        ):
            root.option_add(chave_opcao, valor)

        plano = {"bordercolor": COR_BORDA, "lightcolor": COR_BORDA, "darkcolor": COR_BORDA}
        estilo.configure(".", background=COR_SUPERFICIE, foreground=COR_TEXTO, font=fonte,
                         focuscolor=COR_SUPERFICIE, troughcolor=COR_FUNDO,
                         selectbackground=COR_SELECAO, selectforeground=COR_TEXTO_FORTE, **plano)
        estilo.configure("TFrame", background=COR_SUPERFICIE)
        estilo.configure("TLabel", background=COR_SUPERFICIE, foreground=COR_TEXTO)
        estilo.configure("Dica.TLabel", foreground=COR_SECUNDARIO)
        estilo.configure("Titulo.TLabel", font=("Segoe UI Semibold", 14), foreground=COR_TEXTO_FORTE)
        estilo.configure("TLabelframe", background=COR_SUPERFICIE, borderwidth=1, relief="solid", **plano)
        estilo.configure("TLabelframe.Label", background=COR_SUPERFICIE, foreground=COR_SECUNDARIO,
                         font=("Segoe UI Semibold", 9))
        estilo.configure("TSeparator", background=COR_BORDA)

        botao = {"bordercolor": COR_BOTAO, "lightcolor": COR_BOTAO, "darkcolor": COR_BOTAO}
        estilo.configure("TButton", background=COR_BOTAO, foreground=COR_TEXTO, padding=(12, 5),
                         relief="flat", borderwidth=0, focusthickness=0, **botao)
        estados_botao = [("disabled", COR_SUPERFICIE), ("pressed", COR_SELECAO), ("active", COR_BOTAO_HOVER)]
        estilo.map("TButton",
                   background=estados_botao, bordercolor=estados_botao,
                   lightcolor=estados_botao, darkcolor=estados_botao,
                   foreground=[("disabled", COR_APAGADO)])
        primario = {"bordercolor": COR_DESTAQUE, "lightcolor": COR_DESTAQUE, "darkcolor": COR_DESTAQUE}
        estilo.configure("Primario.TButton", background=COR_DESTAQUE, foreground=COR_FUNDO,
                         font=("Segoe UI Semibold", 10), **primario)
        estados_primario = [("disabled", COR_BOTAO), ("pressed", COR_AZUL), ("active", COR_DESTAQUE_HOVER)]
        estilo.map("Primario.TButton",
                   background=estados_primario, bordercolor=estados_primario,
                   lightcolor=estados_primario, darkcolor=estados_primario,
                   foreground=[("disabled", COR_APAGADO)])
        campo = {"fieldbackground": COR_CAMPO, "foreground": COR_TEXTO, "insertcolor": COR_DESTAQUE,
                 "bordercolor": COR_BORDA, "lightcolor": COR_CAMPO, "darkcolor": COR_CAMPO}
        foco = [("focus", COR_APAGADO)]
        estilo.configure("TEntry", padding=(8, 5), **campo)
        estilo.map("TEntry", bordercolor=foco, fieldbackground=[("disabled", COR_SUPERFICIE)],
                   foreground=[("disabled", COR_APAGADO)])
        estilo.configure("TCombobox", background=COR_CAMPO, arrowcolor=COR_SECUNDARIO, padding=(8, 4),
                         arrowsize=12, **campo)
        estilo.map("TCombobox", bordercolor=foco,
                   fieldbackground=[("readonly", COR_CAMPO), ("disabled", COR_SUPERFICIE)],
                   background=[("active", COR_BOTAO_HOVER)],
                   foreground=[("disabled", COR_APAGADO), ("readonly", COR_TEXTO)],
                   selectbackground=[("readonly", COR_CAMPO)],
                   selectforeground=[("readonly", COR_TEXTO)],
                   arrowcolor=[("disabled", COR_APAGADO)])
        estilo.configure("TSpinbox", background=COR_CAMPO, arrowcolor=COR_SECUNDARIO, padding=(6, 3),
                         arrowsize=10, **campo)
        estilo.map("TSpinbox", bordercolor=foco, fieldbackground=[("disabled", COR_SUPERFICIE)],
                   foreground=[("disabled", COR_APAGADO)], arrowcolor=[("disabled", COR_APAGADO)])

        marcador = {"indicatorbackground": COR_CAMPO, "indicatorforeground": COR_DESTAQUE,
                    "upperbordercolor": COR_BORDA, "lowerbordercolor": COR_BORDA, "indicatormargin": (0, 0, 6, 0)}
        for nome_estilo in ("TCheckbutton", "TRadiobutton"):
            estilo.configure(nome_estilo, background=COR_SUPERFICIE, foreground=COR_TEXTO,
                             focusthickness=0, **marcador)
            estilo.map(nome_estilo, background=[("active", COR_SUPERFICIE)],
                       foreground=[("disabled", COR_APAGADO), ("active", COR_TEXTO_FORTE)],
                       indicatorbackground=[("selected", COR_CAMPO), ("active", COR_BOTAO)])

        sem_borda = {"bordercolor": COR_FUNDO, "lightcolor": COR_FUNDO, "darkcolor": COR_FUNDO}
        estilo.configure("TNotebook", background=COR_FUNDO, borderwidth=0, tabmargins=(0, 0, 0, 0), **sem_borda)
        estilo.configure("TNotebook.Tab", background=COR_FUNDO, foreground=COR_SECUNDARIO,
                         padding=(18, 9), borderwidth=0, focuscolor=COR_FUNDO, **sem_borda)
        estilo.map("TNotebook.Tab",
                   background=[("selected", COR_SUPERFICIE), ("active", COR_FUNDO)],
                   foreground=[("selected", COR_TEXTO_FORTE), ("active", COR_TEXTO)],
                   bordercolor=[("selected", COR_SUPERFICIE)],
                   lightcolor=[("selected", COR_SUPERFICIE)],
                   darkcolor=[("selected", COR_SUPERFICIE)],
                   expand=[("selected", (0, 0, 0, 0))])

        lista = {"background": COR_FUNDO, "fieldbackground": COR_FUNDO, "foreground": COR_TEXTO,
                 "borderwidth": 0, **sem_borda}
        estilo.configure("Treeview", rowheight=28, **lista)
        estilo.configure("Review.Treeview", rowheight=58, **lista)
        for nome_estilo in ("Treeview", "Review.Treeview"):
            estilo.map(nome_estilo, background=[("selected", COR_SELECAO)],
                       foreground=[("selected", COR_TEXTO_FORTE)])
        estilo.configure("Treeview.Heading", background=COR_FUNDO, foreground=COR_SECUNDARIO,
                         font=("Segoe UI Semibold", 9), padding=(8, 6), relief="flat", borderwidth=0,
                         **sem_borda)
        estilo.map("Treeview.Heading", background=[("active", COR_SUPERFICIE)],
                   foreground=[("active", COR_TEXTO)])
        try:
            estilo.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
            estilo.layout("Review.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        except tk.TclError:
            pass

        barra = {"background": COR_BOTAO, "troughcolor": COR_FUNDO, "arrowsize": 10, "borderwidth": 0,
                 "gripcount": 0, "relief": "flat", **sem_borda}
        estados_barra = [("pressed", COR_APAGADO), ("active", COR_BOTAO_HOVER)]
        for orientacao, lado in (("Vertical", "ns"), ("Horizontal", "we")):
            nome_estilo = orientacao + ".TScrollbar"
            try:
                estilo.layout(nome_estilo, [(orientacao + ".Scrollbar.trough", {
                    "sticky": lado,
                    "children": [(orientacao + ".Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})],
                })])
            except tk.TclError:
                pass
            estilo.configure(nome_estilo, **barra)
            estilo.map(nome_estilo, background=estados_barra)
        estilo.configure("Horizontal.TScale", background=COR_BOTAO, troughcolor=COR_CAMPO,
                         sliderthickness=14, borderwidth=0, **sem_borda)
        estilo.map("Horizontal.TScale", background=[("active", COR_DESTAQUE)])

    # --------------------------------------------------------
    # Montagem
    # --------------------------------------------------------

    def _montar(self):

        tk = self.tk
        ttk = self.ttk

        abas = ttk.Notebook(self.root)
        abas.pack(fill="both", expand=True)

        principal = ttk.Frame(abas, padding=16)
        abas.add(principal, text="Capas")

        quadro_fotos = ttk.Frame(abas, padding=12)
        abas.add(quadro_fotos, text="Fotos (recorte)")
        self.quadro_fotos = quadro_fotos

        quadro_revisao = ttk.Frame(abas, padding=16)
        abas.add(quadro_revisao, text="Revisão")
        self.aba_fotos = AbaFotos(self, quadro_fotos, quadro_revisao)

        self.abas = abas
        self.quadro_revisao = quadro_revisao
        abas.bind("<<NotebookTabChanged>>", self._aba_principal_mudou)

        quadro_pastas = ttk.Frame(abas, padding=16)
        abas.add(quadro_pastas, text="Pastas e links")
        self.aba_pastas = AbaPastas(self, quadro_pastas)

        quadro_cores = ttk.Frame(abas, padding=16)
        abas.add(quadro_cores, text="Cores dos cursos")
        self.aba_cores = AbaCores(self, quadro_cores)

        principal.columnconfigure(0, weight=1)
        principal.rowconfigure(1, weight=1)
        principal.rowconfigure(4, weight=1)

        # ---------------- Origem ----------------

        origem = ttk.Frame(principal)
        origem.grid(row=0, column=0, sticky="we")

        ttk.Radiobutton(
            origem, text="Sem planilha (usar as fotos da pasta)", value="fotos",
            variable=self.origem, command=self._escolheu_origem,
        ).pack(side="left")

        ttk.Radiobutton(
            origem, text="Colar alunos", value="colar",
            variable=self.origem, command=self._escolheu_origem,
        ).pack(side="left", padx=(16, 0))

        ttk.Radiobutton(
            origem, text="Planilha (arquivo ou link do Google)", value="arquivo",
            variable=self.origem, command=self._escolheu_origem,
        ).pack(side="left", padx=(16, 4))

        self.campo_planilha = ttk.Entry(origem, textvariable=self.planilha, width=60)
        self.campo_planilha.pack(side="left", fill="x", expand=True, padx=4)

        self.botao_planilha = ttk.Button(
            origem, text="Escolher...", command=self._escolher_planilha
        )
        self.botao_planilha.pack(side="left")

        # ---------------- Colagem + tabela ----------------

        colagem = ttk.LabelFrame(
            principal,
            text="Alunos  (copie do Excel as colunas NOME | CURSO | PLANO e cole aqui)",
            padding=6,
        )
        colagem.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        colagem.columnconfigure(0, weight=1)
        colagem.columnconfigure(1, weight=2)
        colagem.rowconfigure(0, weight=1)

        self.texto_colagem = tk.Text(colagem, height=10, wrap="none", undo=True)
        self.texto_colagem.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.texto_colagem.bind("<<Paste>>", lambda e: self.root.after(50, self._colou_texto))

        colunas = ("linha", "nome", "curso", "plano")
        self.tabela = ttk.Treeview(colagem, columns=colunas, show="headings", height=10)
        for col, titulo, largura in (
            ("linha", "Linha", 50),
            ("nome", "Nome", 220),
            ("curso", "Curso", 160),
            ("plano", "Plano", 170),
        ):
            self.tabela.heading(col, text=titulo)
            self.tabela.column(col, width=largura, anchor="w", stretch=(col != "linha"))
        self.tabela.tag_configure("aviso", foreground=COR_VERMELHO)
        self.tabela.grid(row=0, column=1, sticky="nsew")

        barra = ttk.Scrollbar(colagem, orient="vertical", command=self.tabela.yview)
        barra.grid(row=0, column=2, sticky="ns")
        self.tabela.configure(yscrollcommand=barra.set)

        botoes_colagem = ttk.Frame(colagem)
        botoes_colagem.grid(row=1, column=0, columnspan=3, sticky="we", pady=(6, 0))

        self.botao_colar = ttk.Button(
            botoes_colagem, text="Colar da área de transferência",
            command=self._colar_clipboard,
        )
        self.botao_colar.pack(side="left")

        self.botao_carregar = ttk.Button(
            botoes_colagem, text="Carregar", command=self.carregar
        )
        self.botao_carregar.pack(side="left", padx=6)

        self.botao_limpar = ttk.Button(
            botoes_colagem, text="Limpar", command=self._limpar
        )
        self.botao_limpar.pack(side="left")

        self.botao_ler_fotos = ttk.Button(
            botoes_colagem, text="Ler fotos da pasta", command=self.ler_fotos
        )
        self.botao_ler_fotos.pack(side="left", padx=(16, 4))

        ttk.Label(botoes_colagem, text="Curso p/ todos:").pack(side="left", padx=(6, 2))
        campo_curso = ttk.Entry(botoes_colagem, textvariable=self.curso_todos, width=22)
        campo_curso.pack(side="left")
        campo_curso.bind("<Return>", lambda e: self._aplicar_curso_todos())
        ttk.Button(
            botoes_colagem, text="Aplicar", command=self._aplicar_curso_todos
        ).pack(side="left", padx=(4, 0))

        ttk.Label(botoes_colagem, text="Plano p/ todos:").pack(side="left", padx=(16, 2))
        self.campo_plano_todos = ttk.Combobox(
            botoes_colagem, textvariable=self.plano_todos, width=24, postcommand=self._listar_planos)
        self.campo_plano_todos.pack(side="left")
        self.campo_plano_todos.bind("<Return>", lambda e: self._aplicar_plano_todos())
        self.campo_plano_todos.bind("<<ComboboxSelected>>", lambda e: self._aplicar_plano_todos())
        ttk.Button(
            botoes_colagem, text="Aplicar", command=self._aplicar_plano_todos
        ).pack(side="left", padx=(4, 0))

        self.total_alunos = ttk.Label(botoes_colagem, text="")
        self.total_alunos.pack(side="right")

        # ---------------- Filtros e opções ----------------

        filtros = ttk.LabelFrame(principal, text="Filtros e opções", padding=6)
        filtros.grid(row=2, column=0, sticky="we", pady=(8, 0))

        ttk.Label(filtros, text="Aluno").grid(row=0, column=0, sticky="w")
        ttk.Entry(filtros, textvariable=self.aluno, width=24).grid(row=0, column=1, sticky="w", padx=(4, 12))

        ttk.Label(filtros, text="Plano").grid(row=0, column=2, sticky="w")
        ttk.Entry(filtros, textvariable=self.plano, width=20).grid(row=0, column=3, sticky="w", padx=(4, 12))

        ttk.Label(filtros, text="Linhas").grid(row=0, column=4, sticky="w")
        ttk.Entry(filtros, textvariable=self.linhas, width=14).grid(row=0, column=5, sticky="w", padx=(4, 4))
        ttk.Label(filtros, text="ex.: 10-20 ou 5,8,12", style="Dica.TLabel").grid(row=0, column=6, sticky="w")

        ttk.Checkbutton(
            filtros, text="Só quem deu erro na última vez", variable=self.so_erros
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

        ttk.Checkbutton(
            filtros, text="Refazer mesmo se a capa já existe", variable=self.refazer
        ).grid(row=1, column=3, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Label(filtros, text="Berlim").grid(row=1, column=5, sticky="e", pady=(6, 0))
        ttk.Combobox(
            filtros, textvariable=self.berlim, state="readonly", width=12,
            values=["Ambos", "Só 24x34", "Só 20x20"],
        ).grid(row=1, column=6, sticky="w", pady=(6, 0))

        ttk.Label(filtros, text="Salvar").grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Combobox(
            filtros, textvariable=self.tipo_saida, state="readonly", width=12, values=TIPOS_SAIDA,
        ).grid(row=3, column=1, sticky="w", padx=(4, 12), pady=(6, 0))

        ttk.Label(filtros, text="Lote").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(filtros, textvariable=self.lote, width=24).grid(row=2, column=1, sticky="w", padx=(4, 12), pady=(6, 0))

        ttk.Label(filtros, text="Status").grid(row=2, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(filtros, textvariable=self.status_filtro, width=20).grid(
            row=2, column=3, sticky="w", padx=(4, 12), pady=(6, 0))

        ttk.Label(
            filtros, text="Lote e Status: só com planilha (ex.: lote 16 ou 15,16; status Aprovado)",
            style="Dica.TLabel",
        ).grid(row=2, column=4, columnspan=3, sticky="w", pady=(6, 0))

        # ---------------- Ações ----------------

        acoes = ttk.Frame(principal)
        acoes.grid(row=3, column=0, sticky="we", pady=(8, 0))

        self.botao_validar = ttk.Button(
            acoes, text="Validar (sem Photoshop)", command=lambda: self.executar(True)
        )
        self.botao_validar.pack(side="left")

        self.botao_gerar = ttk.Button(
            acoes, text="Gerar capas", style="Primario.TButton", command=lambda: self.executar(False)
        )
        self.botao_gerar.pack(side="left", padx=6)

        self.botao_parar = ttk.Button(acoes, text="Parar", command=self.parar, state="disabled")
        self.botao_parar.pack(side="left", padx=(0, 6))

        self.botao_relatorio = ttk.Button(
            acoes, text="Abrir relatório", command=self._abrir_relatorio, state="disabled"
        )
        self.botao_relatorio.pack(side="left")

        ttk.Separator(acoes, orient="vertical").pack(side="left", fill="y", padx=10)

        self.botao_recorte_simular = ttk.Button(
            acoes, text="Recortar fotos (simular)", command=lambda: self.recortar(True)
        )
        self.botao_recorte_simular.pack(side="left")

        self.botao_recorte = ttk.Button(
            acoes, text="Recortar fotos", command=lambda: self.recortar(False)
        )
        self.botao_recorte.pack(side="left", padx=6)

        self.botao_miniaturas = ttk.Button(
            acoes, text="Miniaturas", command=self._abrir_miniaturas
        )
        self.botao_miniaturas.pack(side="left")

        ttk.Label(acoes, textvariable=self.status).pack(side="left", padx=12)

        # ---------------- Saída ----------------

        saida = ttk.LabelFrame(principal, text="Saída", padding=6)
        saida.grid(row=4, column=0, sticky="nsew", pady=(8, 0))
        saida.columnconfigure(0, weight=1)
        saida.rowconfigure(0, weight=1)

        self.texto_saida = tk.Text(saida, height=12, wrap="word", state="disabled",
                                   background=COR_FUNDO, foreground=COR_TEXTO)
        self.texto_saida.grid(row=0, column=0, sticky="nsew")

        barra_saida = ttk.Scrollbar(saida, orient="vertical", command=self.texto_saida.yview)
        barra_saida.grid(row=0, column=1, sticky="ns")
        self.texto_saida.configure(yscrollcommand=barra_saida.set)

        self._atualizar_origem()

    def _aba_principal_mudou(self, evento=None):

        if self.abas.select() == str(self.quadro_revisao):
            self.aba_fotos.abrir_revisao()

    # --------------------------------------------------------
    # Colagem
    # --------------------------------------------------------

    def _carregar_ultima_colagem(self):

        try:
            if ULTIMA_COLAGEM.is_file():
                texto = ULTIMA_COLAGEM.read_text(encoding="utf-8")
                if texto.strip():
                    self.texto_colagem.insert("1.0", texto)
                    self.origem.set("colar")
                    self.carregar()
                    return
        except Exception:
            pass
        self.root.after(300, self.ler_fotos)

    def _colar_clipboard(self):

        try:
            conteudo = self.root.clipboard_get()
        except Exception:
            self.status.set("Área de transferência vazia.")
            return

        self.origem.set("colar")
        self._atualizar_origem()
        self.texto_colagem.delete("1.0", "end")
        self.texto_colagem.insert("1.0", conteudo)
        self._colou_texto()

    def _colou_texto(self):

        self.plano_todos_aplicado = ""

        if self.origem.get() != "colar":
            self.origem.set("colar")
            self._atualizar_origem()
            self.status.set("Alunos colados: usando a lista colada.")
        self.carregar()

    def ler_fotos(self):
        """Monta a lista de capas pelas fotos da pasta (baixa as novas do Drive antes, se for o caso)."""

        if getattr(self, "_lendo_fotos", False):
            return
        self._lendo_fotos = True
        self.botao_ler_fotos.configure(state="disabled")
        self.status.set("Lendo as fotos da pasta...")

        def trabalho():
            erro = None
            alunos = []
            try:
                rf = modulo_recorte()
                cfg = rf.carregar_config()
                if cfg.get("photos_link"):
                    try:
                        rf.sincronizar_drive(cfg["photos_link"], cfg["fotos"], aviso=lambda *a: None)
                    except Exception as e:
                        erro = "Não consegui baixar as fotos do Drive (usando as que já estão na pasta): " + str(e)
                if not Path(cfg["fotos"]).is_dir():
                    raise RuntimeError("Pasta de fotos não encontrada: " + str(cfg["fotos"]))
                alunos = rf.alunos_das_fotos(cfg["fotos"])
            except Exception as e:
                erro = str(e)
                alunos = None
            self.root.after(0, lambda: self._fotos_lidas(alunos, erro))

        threading.Thread(target=trabalho, daemon=True).start()

    def _fotos_lidas(self, alunos, erro):

        self._lendo_fotos = False
        self.botao_ler_fotos.configure(state="normal")

        if alunos is None:
            self.status.set("Sem fotos: " + (erro or ""))
            return

        if self.origem.get() == "arquivo":
            self.status.set("Fotos lidas, mas a planilha escolhida continua valendo.")
            return

        self.origem.set("fotos")
        self._atualizar_origem()
        self.plano_todos_aplicado = ""
        self._preencher_texto(alunos, self.curso_todos.get().strip())

        texto = str(len(alunos)) + " capa(s) montada(s) pelas fotos da pasta (sem planilha)."
        if erro:
            texto += "  " + erro
        self.status.set(texto)

    def _preencher_texto(self, alunos, curso=""):

        linhas = ["NOME\tCURSO\tPLANO"]
        for a in alunos:
            linhas.append(a["nome"] + "\t" + (a.get("curso") or curso) + "\t" + a["plano"])

        estado = str(self.texto_colagem.cget("state"))
        self.texto_colagem.configure(state="normal")
        self.texto_colagem.delete("1.0", "end")
        self.texto_colagem.insert("1.0", "\n".join(linhas) + "\n")
        self.texto_colagem.configure(state=estado)
        self.carregar()

    def _aplicar_curso_todos(self):

        if self.origem.get() == "arquivo":
            return
        alunos = self.carregar()
        curso = self.curso_todos.get().strip()
        for a in alunos:
            a["curso"] = curso
        self._preencher_texto(alunos)
        self.status.set("Curso aplicado em " + str(len(alunos)) + " aluno(s).")

    def _listar_planos(self):

        try:
            automacao = modulo_automacao()
            planos = set(automacao.MODELOS) | {"ESSENCIAL BOOK 20X20"}
        except Exception:
            planos = {"PRIME BOOK 30X20", "PRIME BOOK 30X22", "ESSENCIAL BOOK 20X20", "STUDIO DAY BERLIM 24X34"}
        self.campo_plano_todos.configure(values=sorted(planos))

    def _aplicar_plano_todos(self):

        if self.origem.get() == "arquivo":
            self.status.set("Com planilha escolhida, o plano vem da planilha. Use \"Colar alunos\" para trocar.")
            return
        plano = self.plano_todos.get().strip()
        if not plano:
            self.status.set("Escolha um plano para aplicar em todos.")
            return
        try:
            nome_modelo, modelo = modulo_automacao().encontrar_modelo(plano)
        except Exception:
            nome_modelo, modelo = plano, True
        if not modelo:
            self.status.set("Plano não reconhecido: " + plano)
            return
        alunos = self.carregar()
        for a in alunos:
            a["plano"] = plano
        self.plano_todos_aplicado = plano
        self._preencher_texto(alunos)
        self.status.set("Plano " + str(nome_modelo) + " aplicado em " + str(len(alunos)) + " aluno(s).")

    def _limpar(self):

        self.plano_todos_aplicado = ""
        self.texto_colagem.delete("1.0", "end")
        self.carregar()

    def carregar(self):

        texto = self.texto_colagem.get("1.0", "end")
        self.alunos = interpretar_colagem(texto)

        self.tabela.delete(*self.tabela.get_children())

        avisos = 0

        for aluno in self.alunos:
            incompleto = not aluno["nome"] or not aluno["plano"]
            if incompleto:
                avisos += 1
            self.tabela.insert(
                "", "end",
                values=(aluno["linha"], aluno["nome"], aluno["curso"], aluno["plano"]),
                tags=("aviso",) if incompleto else (),
            )

        texto_total = "Total: " + str(len(self.alunos)) + " aluno(s)"
        if avisos:
            texto_total += "  |  " + str(avisos) + " sem nome ou plano (em vermelho)"
        self.total_alunos.configure(text=texto_total)
        if hasattr(self, "aba_fotos"):
            self.aba_fotos.agendar_revisao()

        return self.alunos

    # --------------------------------------------------------
    # Origem
    # --------------------------------------------------------

    def _atualizar_origem(self):

        estado_lista = "disabled" if self.origem.get() == "arquivo" else "normal"

        self.texto_colagem.configure(state=estado_lista)
        for botao in (self.botao_carregar, self.botao_limpar):
            botao.configure(state=estado_lista)

    def _escolheu_origem(self):

        origem = self.origem.get()
        self._atualizar_origem()
        self.plano_todos_aplicado = ""

        if origem == "fotos":
            self.ler_fotos()
        elif origem == "colar":
            self.texto_colagem.delete("1.0", "end")
            try:
                if ULTIMA_COLAGEM.is_file():
                    self.texto_colagem.insert("1.0", ULTIMA_COLAGEM.read_text(encoding="utf-8"))
            except Exception:
                pass
            self.carregar()
        elif not self.planilha.get().strip():
            self.status.set("Escolha a planilha ou cole o link do Google no campo ao lado.")
            self.campo_planilha.focus_set()

    def _planilha_mudou(self):
        """Planilha preenchida -> passa a usar a planilha; apagada -> volta para as fotos da pasta."""

        tem = bool(self.planilha.get().strip())

        if tem and self.origem.get() != "arquivo":
            self.origem.set("arquivo")
            self._atualizar_origem()
            self.status.set("Planilha identificada: as capas vão sair dela.")
        elif not tem and self.origem.get() == "arquivo":
            self.origem.set("fotos")
            self._atualizar_origem()
            self.ler_fotos()

    def _escolher_planilha(self):

        from tkinter import filedialog

        caminho = filedialog.askopenfilename(
            title="Planilha",
            initialdir=planilha_dir_config(),
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
        )

        if caminho:
            self.planilha.set(caminho)

    # --------------------------------------------------------
    # Execução
    # --------------------------------------------------------

    def _montar_argumentos(self, validar):

        from tkinter import messagebox

        args = []

        if self.origem.get() in ("colar", "fotos"):

            alunos = self.carregar()

            if not any(a["nome"] for a in alunos):
                if self.origem.get() == "fotos":
                    messagebox.showwarning(
                        "Automação de Capas",
                        "Nenhuma foto encontrada na pasta.\n"
                        "Confira a pasta de fotos na aba \"Pastas e links\" e clique em \"Ler fotos da pasta\".",
                    )
                else:
                    messagebox.showwarning("Automação de Capas", "Cole pelo menos um aluno antes de rodar.")
                return None

            PASTA_TEMP.mkdir(parents=True, exist_ok=True)
            if self.origem.get() == "colar":
                ULTIMA_COLAGEM.write_text(self.texto_colagem.get("1.0", "end"), encoding="utf-8")

            try:
                gravar_planilha_temporaria(alunos)
            except PermissionError:
                messagebox.showerror(
                    "Automação de Capas",
                    "Não foi possível gravar a planilha temporária.\n"
                    "Feche o arquivo se estiver aberto no Excel:\n\n" + str(PLANILHA_TEMP),
                )
                return None

            args += ["--planilha", str(PLANILHA_TEMP)]

        else:

            caminho = self.planilha.get().strip()
            if not caminho:
                messagebox.showwarning(
                    "Automação de Capas",
                    "Escolha a planilha ou cole o link do Google.\n"
                    "Sem planilha? Marque \"Sem planilha (usar as fotos da pasta)\".",
                )
                return None
            if caminho:
                if modulo_recorte().eh_link(caminho):
                    self.status.set("Baixando a planilha do Google...")
                    self.root.update_idletasks()
                    try:
                        caminho = str(modulo_recorte().baixar_planilha_link(caminho, PLANILHA_LINK))
                    except Exception as e:
                        messagebox.showerror("Automação de Capas", "Não consegui baixar a planilha do link:\n\n" + str(e))
                        self.status.set("Falha ao baixar a planilha.")
                        return None
                elif not Path(caminho).is_file():
                    messagebox.showwarning("Automação de Capas", "Planilha não encontrada:\n\n" + caminho)
                    return None
                args += ["--planilha", caminho]

        try:
            linhas = parse_faixa_linhas(self.linhas.get())
        except ValueError:
            messagebox.showwarning(
                "Automação de Capas",
                "Campo Linhas inválido.\nUse, por exemplo: 10  ou  10-20  ou  5,8,12-15",
            )
            return None

        if self.aluno.get().strip():
            args += ["--aluno", self.aluno.get().strip()]

        if self.plano.get().strip():
            args += ["--plano", self.plano.get().strip()]

        if self.origem.get() == "arquivo":
            if self.lote.get().strip():
                args += ["--lote", self.lote.get().strip()]
            if self.status_filtro.get().strip():
                args += ["--status", self.status_filtro.get().strip()]

        if linhas:
            args += ["--linhas", linhas]

        if self.so_erros.get():
            args.append("--so-erros")

        if self.refazer.get():
            args.append("--no-skip-existing")

        berlim = self.berlim.get()
        if berlim == "Só 24x34":
            args.append("--berlim-24")
        elif berlim == "Só 20x20":
            args.append("--berlim-20")
        else:
            args.append("--berlim-ambos")

        tipo = self.tipo_saida.get()
        if tipo == "Só JPG":
            args.append("--no-export-psd")
        elif tipo == "Só PSD":
            args += ["--export-psd", "--no-export-jpg"]
        else:
            args.append("--export-psd")

        if validar:
            args.append("--preflight")
        else:
            args += ["--parar-arquivo", str(PARAR_ARQUIVO)]

        return args

    def executar(self, validar):

        from tkinter import messagebox

        if self.processo is not None:
            return

        if not SCRIPT.is_file():
            messagebox.showerror("Automação de Capas", "Script não encontrado:\n\n" + str(SCRIPT))
            return

        args = self._montar_argumentos(validar)
        if args is None:
            return

        self._rodar(
            SCRIPT,
            args,
            RELATORIO_PREFLIGHT if validar else RELATORIO,
            "Validando..." if validar else "Gerando capas...",
            antes=passos_drive(),
        )

    def recortar(self, simular):

        from tkinter import messagebox

        if self.processo is not None:
            return

        if not SCRIPT_RECORTE.is_file():
            messagebox.showerror("Automação de Capas", "Script não encontrado:\n\n" + str(SCRIPT_RECORTE))
            return

        if not simular and not messagebox.askyesno(
            "Recortar fotos",
            "As fotos da pasta FOTO que ainda não estão no tamanho do molde "
            "serão recortadas e substituídas.\n\n"
            "O original de cada uma vai para a pasta _ORIGINAIS.\n\nContinuar?",
        ):
            return

        args = []
        if simular:
            args.append("--simular")
        if self.aluno.get().strip():
            args += ["--aluno", self.aluno.get().strip()]

        self._rodar(
            SCRIPT_RECORTE,
            args,
            RELATORIO_RECORTE,
            "Simulando recorte..." if simular else "Recortando fotos...",
            antes=passos_drive(),
        )

    def _rodar(self, script, args, relatorio, texto_status, antes=None):
        """Roda o script; `antes` é uma lista de (script, args) executados primeiro."""

        self.relatorio_atual = relatorio
        self.resumo = None
        self.parando = False
        try:
            PARAR_ARQUIVO.unlink()
        except OSError:
            pass

        self._limpar_saida()
        self.fila = list(antes or []) + [(script, args)]

        self._habilitar_botoes(False)
        self.status.set(texto_status)
        self._iniciar_proximo()

    def _iniciar_proximo(self):

        script, args = self.fila.pop(0)
        self._escrever("> " + " ".join([script.name] + args) + "\n\n")

        ambiente = dict(os.environ)
        ambiente["PYTHONIOENCODING"] = "utf-8"
        ambiente["PYTHONUNBUFFERED"] = "1"

        try:
            self.processo = subprocess.Popen(
                [python_console(), "-u", str(script)] + args,
                cwd=str(PASTA),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                env=ambiente,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as e:
            self.processo = None
            self.fila = []
            self._habilitar_botoes(True)
            self.status.set("Não foi possível iniciar o script.")
            self._escrever(str(e) + "\n")
            return

        threading.Thread(target=self._ler_saida, daemon=True).start()

    def _ler_saida(self):

        processo = self.processo

        for bruto in iter(processo.stdout.readline, b""):
            linha = bruto.decode("utf-8", errors="replace")
            self.root.after(0, self._escrever, linha)

        codigo = processo.wait()
        self.root.after(0, self._finalizar, codigo)

    def _finalizar(self, codigo):

        self.processo = None

        if codigo == 0 and self.fila:
            self._escrever("\n")
            self._iniciar_proximo()
            return
        self.fila = []

        self._habilitar_botoes(True)

        if self.parando:
            self.parando = False
            try:
                PARAR_ARQUIVO.unlink()
            except OSError:
                pass
            self.status.set("Parado por você. As capas que já estavam prontas foram mantidas.")
        elif codigo != 0:
            self.status.set("Terminou com erro (código " + str(codigo) + "). Veja a saída.")
        elif self.resumo:
            self.status.set("Concluído: " + self.resumo)
        else:
            self.status.set("Concluído.")

        if self.relatorio_atual and self.relatorio_atual.is_file():
            self.botao_relatorio.configure(state="normal")

        if self.relatorio_atual == RELATORIO_RECORTE:
            self.aba_fotos.atualizar_lista()

    # --------------------------------------------------------
    # Utilidades
    # --------------------------------------------------------

    def _habilitar_botoes(self, habilitar):

        estado = "normal" if habilitar else "disabled"
        for botao in (
            self.botao_validar, self.botao_gerar,
            self.botao_recorte_simular, self.botao_recorte,
        ):
            botao.configure(state=estado)
        self.botao_parar.configure(state="disabled" if habilitar else "normal", text="Parar")
        if not habilitar:
            self.botao_relatorio.configure(state="disabled")

    def parar(self):
        """1º clique: para antes da próxima capa. 2º clique: encerra na hora."""

        from tkinter import messagebox

        processo = self.processo
        if processo is None:
            return
        self.fila = []
        gerando_capas = self.relatorio_atual == RELATORIO

        if gerando_capas and not self.parando:
            self.parando = True
            try:
                PASTA_TEMP.mkdir(parents=True, exist_ok=True)
                PARAR_ARQUIVO.write_text("parar", encoding="utf-8")
            except OSError:
                pass
            self.botao_parar.configure(text="Forçar parada")
            self.status.set("Parando: termina a capa atual e não começa a próxima...")
            self._escrever("\n>>> Parada pedida: o processo para depois da capa atual.\n")
            return

        if not messagebox.askyesno(
            "Parar agora",
            "Encerrar o processo imediatamente?\n\n"
            "A capa que está sendo feita agora pode ficar incompleta (e aberta no Photoshop)."
            if gerando_capas else "Encerrar o processo imediatamente?",
        ):
            return
        self.parando = True
        try:
            processo.kill()
        except OSError:
            pass
        self._escrever("\n>>> Processo encerrado por você.\n")

    def _limpar_saida(self):

        self.texto_saida.configure(state="normal")
        self.texto_saida.delete("1.0", "end")
        self.texto_saida.configure(state="disabled")

    def _escrever(self, texto):

        m = RE_RESUMO.search(texto)
        if m:
            self.resumo = "OK " + m.group(1) + " | Pulados " + m.group(2) + " | Erro " + m.group(3)

        m = RE_RESUMO_PREFLIGHT.search(texto)
        if m:
            self.resumo = "OK " + m.group(1) + " | Aviso " + m.group(2) + " | Erro " + m.group(3)

        self.texto_saida.configure(state="normal")
        self.texto_saida.insert("end", texto)
        self.texto_saida.see("end")
        self.texto_saida.configure(state="disabled")

    def _abrir_relatorio(self):

        from tkinter import messagebox

        if self.relatorio_atual and self.relatorio_atual.is_file():
            os.startfile(str(self.relatorio_atual))
        else:
            messagebox.showinfo("Automação de Capas", "Relatório ainda não foi gerado.")

    def _abrir_miniaturas(self):

        from tkinter import messagebox

        if PREVIEW_RECORTE.is_dir() and any(PREVIEW_RECORTE.iterdir()):
            os.startfile(str(PREVIEW_RECORTE))
        else:
            messagebox.showinfo(
                "Automação de Capas",
                "Nenhuma miniatura ainda.\nRode \"Recortar fotos (simular)\" primeiro.",
            )


# ============================================================
# ABA FOTOS (recorte manual com prévia)
# ============================================================

TAMANHOS = ("20X20", "30X20", "24X34")
SAIDAS = TAMANHOS + ("30X22",)

SITUACAO_MOLDE = "no molde"
SITUACAO_RECORTAR = "precisa recortar"
SITUACAO_MANUAL = "ajustada à mão"
SITUACAO_SEM_ROSTO = "sem rosto"
SITUACAO_BAIXA = "resolução baixa"

LADO_EDITOR = 1600
LADO_PREVIA_REVISAO = 1400
INTERVALO_VIGIA_MS = 2000
ALCA_PX = 10


class AbaFotos:

    def __init__(self, janela, quadro, quadro_revisao):

        import tkinter as tk
        from tkinter import ttk

        self.janela = janela
        self.root = janela.root
        self.tk = tk
        self.ttk = ttk
        self.quadro = quadro
        self.quadro_revisao = quadro_revisao

        self.rf = None
        self.cfg = None
        self.calib = None

        self.fotos = []
        self.situacoes = {}

        self.arquivo = None
        self.fonte = None
        self.reduzida = None
        self.fator_red = 1.0
        self.caixa = None
        self.caixa_auto = None
        self.rosto = None
        self.historico = []

        self.escala_tela = 1.0
        self.offset = (0, 0)
        self._img_tela = None
        self._img_tela_chave = None
        self._img_previa = None
        self._arrasto = None
        self._previa_pendente = None
        self._previas_capa = {}
        self._presets_curso = {}
        self._presets_carregando = set()
        self.ocupado = False
        self.cursos_manuais = self._ler_cursos_foto()
        self.cursos_por_foto = {}
        self._revisao_por_curso = {}
        self._revisao_imagens = {}
        self._revisao_pendente = False
        self._revisao_suja = True
        self._revisao_agendada = None
        self._assinaturas = {}
        self._vigiando = False
        self._revisao_carregando = False
        self._curso_revisao_selecionado = None
        self._img_revisao_capa = None
        self.cursos_disponiveis = ()
        self._capa_revisao_atual = None
        self._alunos_por_iid = {}
        self._aluno_nome = None
        self._aluno_escolhido = None
        self._previa_aluno = None
        self._render_nome_pendente = None
        self._escala_tocada = {1: False, 2: False}
        self._preenchendo_nome = False
        self._nome_editado = False
        self._revisao_grupo_por_iid = {}
        self._excluidos_revisao = set()

        self.filtro_nome = tk.StringVar()
        self.filtro_tam = tk.StringVar(value="Todos")
        self.filtro_sit = tk.StringVar(value="Todas")
        self.tam = tk.StringVar(value="30X20")
        self.largura = tk.StringVar()
        self.altura = tk.StringVar()
        self.modo_previa = tk.StringVar(value="capa")
        self.capa_previa = tk.StringVar(value="30X20")
        self.zoom = tk.DoubleVar(value=100.0)
        self.info = tk.StringVar(value="")
        self.alerta = tk.StringVar(value="")
        self.status = tk.StringVar(value="Clique em Atualizar para listar as fotos da pasta FOTO.")
        self.curso_para_foto = tk.StringVar()
        self.curso_revisao_nome = tk.StringVar(value="Selecione um curso")
        self.curso_revisao_cor = tk.StringVar(value="Cor não definida")
        self.status_revisao = tk.StringVar(value="Carregue as fotos para montar a revisão.")
        self.foto_revisao_atual = tk.StringVar(value="Nenhuma foto selecionada")
        self.so_nomes_aviso = tk.BooleanVar(value=False)
        self.nome_revisao_titulo = tk.StringVar(value="Selecione um aluno na lista para ajustar o nome.")
        self.corte_nome = tk.StringVar()
        self.escala1_nome = tk.StringVar(value="100")
        self.escala2_nome = tk.StringVar(value="100")
        self.status_nome = tk.StringVar(value="")

        self._montar()
        self.root.after(300, self.atualizar_lista)
        self.root.after(INTERVALO_VIGIA_MS, self._vigiar_mudancas)
        janela.origem.trace_add("write", lambda *a: self.agendar_revisao())
        janela.planilha.trace_add("write", lambda *a: self.agendar_revisao())

    # --------------------------------------------------------
    # Módulo de recorte (carregado só quando a aba é usada)
    # --------------------------------------------------------

    def _modulo(self):

        if self.rf is None:
            if str(PASTA) not in sys.path:
                sys.path.insert(0, str(PASTA))
            import recortar_fotos
            self.rf = recortar_fotos

        self.cfg = self.rf.carregar_config()
        return self.rf

    def _calibracao(self):

        if self.calib is None:
            rf = self._modulo()
            if rf.CALIBRACAO_FILE.is_file():
                self.calib = json.loads(rf.CALIBRACAO_FILE.read_text(encoding="utf-8"))
            else:
                self.calib = {t: dict(rf.CALIBRACAO_PADRAO) for t in TAMANHOS}
        return self.calib

    # --------------------------------------------------------
    # Montagem
    # --------------------------------------------------------

    def _montar(self):

        tk = self.tk
        ttk = self.ttk
        q = self.quadro
        q.columnconfigure(1, weight=1)
        q.rowconfigure(0, weight=1)

        # ---------------- Lista ----------------

        esquerda = ttk.LabelFrame(q, text="Fotos da pasta FOTO", padding=6)
        esquerda.grid(row=0, column=0, sticky="nsw")
        esquerda.rowconfigure(2, weight=1)

        filtros = ttk.Frame(esquerda)
        filtros.grid(row=0, column=0, columnspan=2, sticky="we")
        ttk.Label(filtros, text="Nome").grid(row=0, column=0, sticky="w")
        campo = ttk.Entry(filtros, textvariable=self.filtro_nome, width=22)
        campo.grid(row=0, column=1, columnspan=3, sticky="we", padx=4)
        campo.bind("<KeyRelease>", lambda e: self._preencher_lista())
        ttk.Label(filtros, text="Tamanho").grid(row=1, column=0, sticky="w", pady=(4, 0))
        combo_tam = ttk.Combobox(filtros, textvariable=self.filtro_tam, state="readonly", width=8,
                                 values=("Todos",) + TAMANHOS)
        combo_tam.grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))
        combo_tam.bind("<<ComboboxSelected>>", lambda e: self._preencher_lista())
        combo_sit = ttk.Combobox(filtros, textvariable=self.filtro_sit, state="readonly", width=16,
                                 values=("Todas", SITUACAO_RECORTAR, SITUACAO_MOLDE, SITUACAO_MANUAL,
                                         SITUACAO_SEM_ROSTO, SITUACAO_BAIXA))
        combo_sit.grid(row=1, column=2, sticky="w", pady=(4, 0))
        combo_sit.bind("<<ComboboxSelected>>", lambda e: self._preencher_lista())

        botoes = ttk.Frame(esquerda)
        botoes.grid(row=1, column=0, columnspan=2, sticky="we", pady=(6, 4))
        ttk.Button(botoes, text="Atualizar", command=self.atualizar_lista).pack(side="left")
        ttk.Button(botoes, text="Próxima", command=self.proxima).pack(side="left", padx=4)
        ttk.Button(botoes, text="Abrir pasta", command=self._abrir_pasta).pack(side="left")

        self.lista = ttk.Treeview(esquerda, columns=("arquivo", "tam", "sit"), show="headings",
                                  height=20, selectmode="browse")
        for col, titulo, largura in (("arquivo", "Arquivo", 190), ("tam", "Tam.", 52), ("sit", "Situação", 110)):
            self.lista.heading(col, text=titulo)
            self.lista.column(col, width=largura, anchor="w", stretch=(col == "arquivo"))
        self.lista.tag_configure(SITUACAO_RECORTAR, foreground=COR_LARANJA)
        self.lista.tag_configure(SITUACAO_MANUAL, foreground=COR_VERDE)
        self.lista.tag_configure(SITUACAO_SEM_ROSTO, foreground=COR_VERMELHO)
        self.lista.tag_configure(SITUACAO_BAIXA, foreground=COR_VERMELHO)
        self.lista.grid(row=2, column=0, sticky="nsew")
        barra = ttk.Scrollbar(esquerda, orient="vertical", command=self.lista.yview)
        barra.grid(row=2, column=1, sticky="ns")
        self.lista.configure(yscrollcommand=barra.set)
        self.lista.bind("<<TreeviewSelect>>", lambda e: self._abrir_selecionada())

        self.total = ttk.Label(esquerda, text="", style="Dica.TLabel")
        self.total.grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # ---------------- Editor ----------------

        centro = ttk.LabelFrame(
            q, text="Corte  (arraste a caixa, use os cantos ou a rodinha para o tamanho, setas para ajuste fino)",
            padding=6,
        )
        centro.grid(row=0, column=1, sticky="nsew", padx=6)
        centro.columnconfigure(0, weight=1)
        centro.rowconfigure(0, weight=1)

        self.tela = tk.Canvas(centro, background=COR_FUNDO, highlightthickness=0,
                              cursor="fleur")
        self.tela.grid(row=0, column=0, sticky="nsew")
        self.tela.bind("<Configure>", lambda e: self._desenhar())
        self.tela.bind("<ButtonPress-1>", self._pressionar)
        self.tela.bind("<B1-Motion>", self._arrastar)
        self.tela.bind("<ButtonRelease-1>", self._soltar)
        self.tela.bind("<MouseWheel>", self._rodinha)
        for tecla, dx, dy in (("Left", -1, 0), ("Right", 1, 0), ("Up", 0, -1), ("Down", 0, 1)):
            self.tela.bind("<" + tecla + ">", lambda e, dx=dx, dy=dy: self._mover_teclado(dx, dy, e))
        self.tela.bind("<plus>", lambda e: self._zoom_relativo(0.97))
        self.tela.bind("<minus>", lambda e: self._zoom_relativo(1.03))
        self.tela.bind("<Control-z>", lambda e: self.desfazer())

        ferramentas = ttk.Frame(centro)
        ferramentas.grid(row=1, column=0, sticky="we", pady=(6, 0))
        self.botoes_editor = []
        for texto, comando in (
            ("Auto (rosto)", self.auto_rosto),
            ("Centralizar rosto", self.centralizar_rosto),
            ("Desfazer", self.desfazer),
            ("Voltar ao automático", self.voltar_automatico),
        ):
            b = ttk.Button(ferramentas, text=texto, command=comando)
            b.pack(side="left", padx=(0, 4))
            self.botoes_editor.append(b)
        ttk.Label(ferramentas, text="Tamanho do corte").pack(side="left", padx=(12, 4))
        self.controle_zoom = ttk.Scale(ferramentas, from_=15, to=100, variable=self.zoom,
                                       orient="horizontal", length=180)
        self.controle_zoom.pack(side="left")
        for evento in ("<B1-Motion>", "<ButtonRelease-1>"):
            self.controle_zoom.bind(evento, lambda e: self.root.after_idle(self._zoom_escala, self.zoom.get()), add="+")

        ttk.Label(centro, textvariable=self.status, style="Dica.TLabel").grid(row=2, column=0, sticky="w", pady=(4, 0))

        # ---------------- Prévia e tamanho ----------------

        direita = ttk.Frame(q)
        direita.grid(row=0, column=2, sticky="nse")

        previa = ttk.LabelFrame(direita, text="Prévia do resultado", padding=6)
        previa.pack(fill="x")
        modos = ttk.Frame(previa)
        modos.pack(fill="x")
        ttk.Radiobutton(modos, text="Na capa", value="capa", variable=self.modo_previa,
                        command=self._agendar_previa).pack(side="left")
        ttk.Radiobutton(modos, text="Foto final", value="foto", variable=self.modo_previa,
                        command=self._agendar_previa).pack(side="left", padx=8)
        self.combo_capa = ttk.Combobox(modos, textvariable=self.capa_previa, width=7, state="readonly")
        self.combo_capa.pack(side="left")
        self.combo_capa.bind("<<ComboboxSelected>>", lambda e: (self._agendar_previa(), self._desenhar()))
        self.botao_gerar_previa = ttk.Button(modos, text="Gerar prévia da capa", command=self.gerar_previa_capa)
        self.botao_gerar_previa.pack(side="right")

        self._img_previa = ImageTk.PhotoImage(Image.new("RGB", (380, 380), COR_FUNDO))
        self.rotulo_previa = tk.Label(previa, background=COR_FUNDO, image=self._img_previa)
        self.rotulo_previa.pack(pady=(6, 0))

        tamanho = ttk.LabelFrame(direita, text="Tamanho da foto para o Photoshop", padding=6)
        tamanho.pack(fill="x", pady=(8, 0))
        ttk.Label(tamanho, text="Capa").grid(row=0, column=0, sticky="w")
        combo = ttk.Combobox(tamanho, textvariable=self.tam, state="readonly", width=8, values=TAMANHOS)
        combo.grid(row=0, column=1, sticky="w", padx=4)
        combo.bind("<<ComboboxSelected>>", lambda e: self._trocar_tamanho())
        ttk.Label(tamanho, text="Largura").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(tamanho, textvariable=self.largura, width=8).grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))
        ttk.Label(tamanho, text="Altura").grid(row=1, column=2, sticky="w", pady=(4, 0))
        ttk.Entry(tamanho, textvariable=self.altura, width=8).grid(row=1, column=3, sticky="w", padx=4, pady=(4, 0))
        ttk.Label(tamanho, text="px").grid(row=1, column=4, sticky="w", pady=(4, 0))
        ttk.Button(tamanho, text="Gravar tamanho", command=self.gravar_tamanho).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Button(tamanho, text="Ler tamanho do PSD", command=self.ler_tamanho_psd).grid(
            row=2, column=2, columnspan=3, sticky="w", pady=(6, 0))

        ttk.Label(direita, textvariable=self.info, justify="left").pack(fill="x", pady=(8, 0))
        tk.Label(direita, textvariable=self.alerta, foreground=COR_VERMELHO, background=COR_SUPERFICIE, justify="left",
                 wraplength=380, anchor="w").pack(fill="x")

        acoes = ttk.LabelFrame(direita, text="Gravar", padding=6)
        acoes.pack(fill="x", pady=(8, 0))
        self.botao_salvar = ttk.Button(acoes, text="Salvar esta foto", style="Primario.TButton", command=lambda: self.salvar(False))
        self.botao_salvar.grid(row=0, column=0, sticky="we")
        self.botao_salvar_prox = ttk.Button(acoes, text="Salvar e próxima", command=lambda: self.salvar(True))
        self.botao_salvar_prox.grid(row=0, column=1, sticky="we", padx=4)
        self.botao_restaurar = ttk.Button(acoes, text="Restaurar original", command=self.restaurar)
        self.botao_restaurar.grid(row=1, column=0, columnspan=2, sticky="we", pady=(4, 0))
        ttk.Label(
            acoes, style="Dica.TLabel", wraplength=360, justify="left",
            text="Grava na pasta FOTO no tamanho exato acima. O original fica em _ORIGINAIS "
                 "e o corte é lembrado para o \"Recortar fotos\" em lote.",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        self._montar_revisao(self.quadro_revisao)

    def _montar_revisao(self, pagina):

        ttk = self.ttk
        tk = self.tk
        pagina.columnconfigure(0, weight=1)
        pagina.rowconfigure(1, weight=1)

        ttk.Label(
            pagina,
            text="Preset visual aproximado do curso. A composição usa a demonstração do PSD e não substitui a prova final do Photoshop.",
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        corpo = ttk.Frame(pagina)
        corpo.grid(row=1, column=0, sticky="nsew")
        corpo.columnconfigure(0, weight=1)
        corpo.columnconfigure(1, weight=2)
        corpo.rowconfigure(0, weight=3)
        corpo.rowconfigure(1, weight=2)

        quadro_lista = ttk.LabelFrame(corpo, text="Cursos encontrados", padding=6)
        quadro_lista.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        quadro_lista.columnconfigure(0, weight=1)
        quadro_lista.rowconfigure(0, weight=1)

        self.lista_revisao = ttk.Treeview(
            quadro_lista, columns=("curso", "quantidade", "cor", "nomes"), show="tree headings",
            selectmode="browse", style="Review.Treeview",
        )
        self.lista_revisao.heading("#0", text="Brasão")
        self.lista_revisao.column("#0", width=68, anchor="center", stretch=False)
        for coluna, titulo, largura in (
            ("curso", "Curso", 170),
            ("quantidade", "Fotos", 50),
            ("cor", "Cor", 76),
            ("nomes", "Nomes", 90),
        ):
            self.lista_revisao.heading(coluna, text=titulo)
            self.lista_revisao.column(coluna, width=largura, anchor="w")
        self.lista_revisao.grid(row=0, column=0, sticky="nsew")
        barra_revisao = ttk.Scrollbar(quadro_lista, orient="vertical", command=self.lista_revisao.yview)
        barra_revisao.grid(row=0, column=1, sticky="ns")
        self.lista_revisao.configure(yscrollcommand=barra_revisao.set)
        self.lista_revisao.bind("<<TreeviewSelect>>", lambda e: self._selecionar_curso_revisao())

        quadro_alunos = ttk.LabelFrame(corpo, text="Alunos do curso (nome na capa)", padding=6)
        quadro_alunos.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(8, 0))
        quadro_alunos.columnconfigure(0, weight=1)
        quadro_alunos.rowconfigure(1, weight=1)
        ttk.Checkbutton(
            quadro_alunos, text="Só nomes com aviso (todos os cursos)",
            variable=self.so_nomes_aviso, command=self._preencher_alunos_revisao,
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.lista_alunos = ttk.Treeview(
            quadro_alunos, columns=("nome", "capa", "situacao"), show="headings", selectmode="browse",
        )
        for coluna, titulo, largura in (("nome", "Aluno", 230), ("capa", "Capa", 56), ("situacao", "Nome", 130)):
            self.lista_alunos.heading(coluna, text=titulo)
            self.lista_alunos.column(coluna, width=largura, anchor="w")
        self.lista_alunos.tag_configure("nome_reduzido", background="#36352b", foreground=COR_AMARELO)
        self.lista_alunos.tag_configure("nome_estoura", background="#3b2a30", foreground="#e3a3aa")
        self.lista_alunos.tag_configure("nome_ajustado", foreground=COR_AZUL)
        self.lista_alunos.grid(row=1, column=0, sticky="nsew")
        barra_alunos = ttk.Scrollbar(quadro_alunos, orient="vertical", command=self.lista_alunos.yview)
        barra_alunos.grid(row=1, column=1, sticky="ns")
        self.lista_alunos.configure(yscrollcommand=barra_alunos.set)
        self.lista_alunos.bind("<<TreeviewSelect>>", lambda e: self._selecionar_aluno_revisao())

        painel = ttk.Frame(corpo, padding=(4, 0, 0, 0))
        painel.grid(row=0, column=1, rowspan=2, sticky="nsew")
        painel.columnconfigure(0, weight=1)
        painel.rowconfigure(2, weight=1)

        ttk.Label(painel, textvariable=self.curso_revisao_nome,
                  style="Titulo.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(painel, text="Trocar curso...", command=self._abrir_atribuicao_curso).grid(
            row=0, column=1, sticky="e", padx=(8, 0))
        ttk.Label(painel, textvariable=self.curso_revisao_cor).grid(
            row=1, column=0, sticky="w", pady=(2, 8))
        self.amostra_cor_revisao = tk.Canvas(
            painel, width=34, height=20, background=COR_CAMPO, highlightthickness=1,
            highlightbackground=COR_BORDA,
        )
        self.amostra_cor_revisao.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(2, 8))

        imagens = ttk.Frame(painel, height=320)
        imagens.grid(row=2, column=0, columnspan=2, sticky="nsew")
        imagens.grid_propagate(False)
        imagens.columnconfigure(0, weight=1)
        imagens.rowconfigure(0, weight=1)
        self._img_revisao_capa = None
        self.previa_revisao_capa = tk.Label(imagens, background=COR_FUNDO)
        self.previa_revisao_capa.grid(row=0, column=0, sticky="nsew")
        self.quadro_previa_revisao = imagens
        imagens.bind("<Configure>", lambda e: self._mostrar_previa_revisao(self._capa_revisao_atual))
        ttk.Label(painel, textvariable=self.foto_revisao_atual).grid(
            row=3, column=0, sticky="w", pady=(6, 0))

        quadro_nome = ttk.LabelFrame(painel, text="Nome nesta capa", padding=6)
        quadro_nome.grid(row=4, column=0, columnspan=2, sticky="we", pady=(8, 0))
        quadro_nome.columnconfigure(1, weight=1)
        ttk.Label(quadro_nome, textvariable=self.nome_revisao_titulo,
                  font=("Segoe UI Semibold", 10)).grid(row=0, column=0, columnspan=6, sticky="w")
        ttk.Label(quadro_nome, text="Cortar depois de:").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.combo_corte_nome = ttk.Combobox(quadro_nome, textvariable=self.corte_nome, state="disabled")
        self.combo_corte_nome.grid(row=1, column=1, columnspan=5, sticky="we", padx=(6, 0), pady=(4, 0))
        self.combo_corte_nome.bind("<<ComboboxSelected>>", lambda e: self._nome_mudou())
        ttk.Label(quadro_nome, text="1ª linha (%):").grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.spin_escala1 = ttk.Spinbox(
            quadro_nome, from_=50, to=130, increment=5, width=6, textvariable=self.escala1_nome, state="disabled")
        self.spin_escala1.grid(row=2, column=1, sticky="w", padx=(6, 0), pady=(4, 0))
        ttk.Label(quadro_nome, text="2ª linha (%):").grid(row=2, column=2, sticky="w", padx=(12, 0), pady=(4, 0))
        self.spin_escala2 = ttk.Spinbox(
            quadro_nome, from_=50, to=130, increment=5, width=6, textvariable=self.escala2_nome, state="disabled")
        self.spin_escala2.grid(row=2, column=3, sticky="w", padx=(6, 0), pady=(4, 0))
        self.botao_nome_auto = ttk.Button(
            quadro_nome, text="Automático", command=self._ajuste_nome_automatico, state="disabled")
        self.botao_nome_auto.grid(row=2, column=4, padx=(12, 0), pady=(4, 0))
        self.botao_nome_salvar = ttk.Button(
            quadro_nome, text="Salvar", style="Primario.TButton", command=self._salvar_ajuste_nome, state="disabled")
        self.botao_nome_salvar.grid(row=2, column=5, padx=(6, 0), pady=(4, 0))
        ttk.Label(quadro_nome, textvariable=self.status_nome, wraplength=620).grid(
            row=3, column=0, columnspan=6, sticky="w", pady=(4, 0))
        self.escala1_nome.trace_add("write", lambda *a: self._escala_nome_mudou(1))
        self.escala2_nome.trace_add("write", lambda *a: self._escala_nome_mudou(2))

        acoes = ttk.Frame(painel)
        acoes.grid(row=5, column=0, sticky="we", pady=(8, 0))
        ttk.Button(acoes, text="Atualizar revisão", command=self._atualizar_revisao_manual).pack(side="left")
        ttk.Button(acoes, text="Abrir foto no recorte", command=self._abrir_foto_revisao).pack(
            side="left", padx=6)
        ttk.Button(acoes, text="Excluir da revisão", command=self._excluir_da_revisao).pack(side="left")
        ttk.Button(acoes, text="Refazer prévia", command=self._excluir_previa_revisao).pack(
            side="left", padx=6)
        ttk.Label(painel, textvariable=self.status_revisao, wraplength=640).grid(
            row=6, column=0, sticky="w", pady=(8, 0))

    def _ler_cursos_foto(self):

        try:
            dados = json.loads(CURSOS_POR_FOTO_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}
        mapa = dados.get("cursos_por_foto", {}) if isinstance(dados, dict) else {}
        if not isinstance(mapa, dict):
            return {}
        return {
            chave_caminho_foto(caminho): str(curso).strip()
            for caminho, curso in mapa.items()
            if str(curso).strip()
        }

    def _gravar_cursos_foto(self):

        CURSOS_POR_FOTO_FILE.parent.mkdir(parents=True, exist_ok=True)
        temporario = CURSOS_POR_FOTO_FILE.with_suffix(".tmp")
        dados = {"cursos_por_foto": self.cursos_manuais}
        temporario.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")
        os.replace(temporario, CURSOS_POR_FOTO_FILE)
        self._conhecer_arquivo("cursos", CURSOS_POR_FOTO_FILE)

    # --------------------------------------------------------
    # Atualização automática (pasta de fotos e arquivos da revisão)
    # --------------------------------------------------------

    @staticmethod
    def _assinatura_arquivo(caminho):

        try:
            estado = os.stat(caminho)
        except OSError:
            return None
        return estado.st_mtime_ns, estado.st_size

    @staticmethod
    def _assinatura_pasta(pasta):

        try:
            with os.scandir(pasta) as itens:
                return frozenset(
                    (item.name, item.stat().st_mtime_ns, item.stat().st_size)
                    for item in itens if item.is_file()
                )
        except OSError:
            return None

    def _conhecer_arquivo(self, nome, caminho):
        """Registra uma gravação feita pela própria UI para o vigia não tratá-la como mudança externa."""

        if self._assinaturas:
            self._assinaturas[nome] = self._assinatura_arquivo(caminho)

    def _vigiar_mudancas(self):

        rf = self.rf
        if rf is None or self.cfg is None or self._vigiando:
            self.root.after(INTERVALO_VIGIA_MS, self._vigiar_mudancas)
            return
        arquivos = {
            "config": rf.CONFIG_FILE,
            "calibracao": rf.CALIBRACAO_FILE,
            "manual": rf.MANUAL_FILE,
            "ajustes_nome": rf.AJUSTES_NOME_FILE,
            "cursos": CURSOS_POR_FOTO_FILE,
        }
        planilha = self.janela.planilha.get().strip()
        if self.janela.origem.get() == "arquivo" and planilha and not rf.eh_link(planilha):
            arquivos["planilha"] = Path(planilha)
        pasta = Path(self.cfg["fotos"])
        self._vigiando = True

        def trabalho():
            novas = {nome: self._assinatura_arquivo(caminho) for nome, caminho in arquivos.items()}
            novas["fotos"] = self._assinatura_pasta(pasta)
            self.root.after(0, self._mudancas_lidas, novas)

        threading.Thread(target=trabalho, daemon=True).start()

    def _mudancas_lidas(self, novas):

        self._vigiando = False
        antigas, self._assinaturas = self._assinaturas, novas
        self.root.after(INTERVALO_VIGIA_MS, self._vigiar_mudancas)
        if not antigas:
            return
        mudou = {nome for nome in novas.keys() | antigas.keys() if novas.get(nome) != antigas.get(nome)}
        if not mudou:
            return

        if "config" in mudou:
            self._modulo()
        if "calibracao" in mudou:
            self.calib = None
        if "cursos" in mudou:
            self.cursos_manuais = self._ler_cursos_foto()
        if mudou & {"fotos", "config", "manual"}:
            self.atualizar_lista(silencioso=True)
        if self.arquivo and mudou & {"config", "calibracao", "ajustes_nome", "cursos"}:
            self._agendar_previa()
        self.agendar_revisao()

    def _revisao_visivel(self):
        return self.janela.abas.select() == str(self.janela.quadro_revisao)

    def agendar_revisao(self, atraso=500):
        """Remonta a revisão quando algo que ela mostra muda; com a aba fechada, só na próxima abertura."""

        self._revisao_suja = True
        if self._revisao_agendada:
            self.root.after_cancel(self._revisao_agendada)
        self._revisao_agendada = self.root.after(atraso, self._rodar_revisao_agendada)

    def _rodar_revisao_agendada(self):

        self._revisao_agendada = None
        if not self._revisao_suja or not self._revisao_visivel():
            return
        if self._nome_editado:
            self.status_revisao.set("Há mudanças para mostrar: a revisão atualiza depois que você salvar o nome.")
            self._revisao_agendada = self.root.after(1500, self._rodar_revisao_agendada)
            return
        self._recarregar_revisao()

    def abrir_revisao(self):

        if self._revisao_suja:
            self._recarregar_revisao()

    def _recarregar_revisao(self):

        if self._revisao_carregando:
            self._revisao_pendente = True
            return
        self._revisao_suja = False
        try:
            rf = self._modulo()
            automacao = modulo_automacao()
            calibracao = dict(self._calibracao())
        except Exception as e:
            self.status_revisao.set("Não foi possível carregar os módulos da revisão: " + str(e))
            return

        registros_colados = [dict(registro) for registro in self.janela.alunos]
        cursos_manuais = dict(self.cursos_manuais)
        fotos = list(self.fotos)
        cfg = dict(self.cfg)
        origem = self.janela.origem.get()
        caminho_planilha = self.janela.planilha.get().strip()
        excluidos = set(self._excluidos_revisao)
        plano_extras = self.janela.plano_todos_aplicado if origem != "arquivo" else ""
        self._revisao_carregando = True
        self.status_revisao.set("Lendo cursos, fotos e brasões...")

        self.cursos_disponiveis = tuple(sorted(automacao.CURSOS))

        def trabalho():
            grupos = {}
            cursos_da_foto = {}
            erros = []
            try:
                if origem == "arquivo":
                    if not caminho_planilha:
                        raise ValueError("Escolha uma planilha antes de carregar a revisão.")
                    if rf.eh_link(caminho_planilha):
                        caminho = rf.baixar_planilha_link(caminho_planilha, PLANILHA_LINK)
                    else:
                        caminho = Path(caminho_planilha)
                    registros = automacao.ler_planilha(caminho)
                else:
                    registros = registros_colados

                entradas = []
                caminhos_vistos = set()
                for registro in registros:
                    nome = str(registro.get("nome") or "").strip()
                    if not nome:
                        continue
                    _, modelo = automacao.encontrar_modelo(registro.get("plano", ""))
                    if not modelo:
                        continue
                    tamanho_capa = modelo.get("tamanho", "30X20")
                    tamanho_foto = "30X20" if tamanho_capa == "30X22" else tamanho_capa
                    foto = None
                    erro_foto = None
                    try:
                        foto = automacao.encontrar_foto(nome, tamanho_foto)
                    except Exception as e:
                        erro_foto = str(e)

                    chave_foto = chave_caminho_foto(foto) if foto else None
                    if chave_foto:
                        caminhos_vistos.add(chave_foto)
                    curso = (cursos_manuais.get(chave_foto) if chave_foto else None) or str(
                        registro.get("curso") or ""
                    ).strip()
                    if chave_foto and curso:
                        cursos_da_foto[chave_foto] = curso
                    entradas.append({
                        "nome": nome,
                        "curso": curso,
                        "foto": foto,
                        "erro": erro_foto,
                        "tamanho_foto": tamanho_foto,
                        "tamanho_capa": tamanho_capa,
                    })

                for foto in fotos:
                    chave_foto = chave_caminho_foto(foto)
                    curso = cursos_manuais.get(chave_foto)
                    if not curso or chave_foto in caminhos_vistos:
                        continue
                    tamanho_capa = rf.tamanho_do_arquivo(foto)
                    modelo_extra = automacao.encontrar_modelo(plano_extras)[1] if plano_extras else None
                    if modelo_extra:
                        tamanho_capa = modelo_extra.get("tamanho", tamanho_capa)
                    cursos_da_foto[chave_foto] = curso
                    entradas.append({
                        "nome": rf.nome_do_aluno(foto),
                        "curso": curso,
                        "foto": foto,
                        "erro": None,
                        "tamanho_foto": "30X20" if tamanho_capa == "30X22" else tamanho_capa,
                        "tamanho_capa": tamanho_capa,
                    })

                entradas = [e for e in entradas
                            if ("aluno", e["nome"], e["tamanho_capa"]) not in excluidos]
                for entrada in entradas:
                    curso = entrada["curso"]
                    simbolo = None
                    aviso = None
                    if curso:
                        try:
                            simbolo, aviso = automacao.resolver_curso(curso)
                        except Exception as e:
                            aviso = str(e)
                    entrada["simbolo"] = simbolo
                    entrada["aviso"] = aviso
                    entrada["cor"] = automacao.obter_cor_background(simbolo) if simbolo else None
                    curso_normalizado = automacao.chave(curso) if curso else ""
                    curso_canonico = automacao.CURSO_ALIASES.get(curso_normalizado, curso)
                    id_grupo = (
                        automacao.chave(curso_canonico) if curso_canonico
                        else "sem-curso"
                    )
                    if ("curso", id_grupo) in excluidos:
                        continue
                    grupo = grupos.setdefault(id_grupo, {
                        "curso": curso_canonico or "Sem curso",
                        "simbolo": simbolo,
                        "cor": entrada["cor"],
                        "entradas": [],
                        "amostra": None,
                        "brasao": None,
                        "capa": None,
                        "avisos": [],
                    })
                    grupo["entradas"].append(entrada)
                    if aviso:
                        grupo["avisos"].append(aviso)

                ajustes_nome = rf.ler_ajustes_nome()
                for entrada in entradas:
                    tamanho_capa = entrada["tamanho_capa"]
                    ajuste = ajustes_nome.get(rf.chave_ajuste_nome(entrada["nome"], tamanho_capa)) or {}
                    try:
                        entrada["layout"] = rf.calcular_layout_nome(tamanho_capa, entrada["nome"], ajuste)
                    except Exception:
                        entrada["layout"] = None

                simbolos_por_tamanho = {}
                for grupo in grupos.values():
                    amostra = next((entrada for entrada in grupo["entradas"] if entrada["foto"]), None)
                    grupo["amostra"] = amostra
                    if amostra and grupo["simbolo"]:
                        simbolos_por_tamanho.setdefault(amostra["tamanho_capa"], set()).add(grupo["simbolo"])

                tamanhos_usados = {g["amostra"]["tamanho_capa"] for g in grupos.values() if g["amostra"]}
                for tamanho in sorted(tamanhos_usados):
                    if rf.ler_previa_capa(tamanho) is None:
                        self.root.after(0, self.status_revisao.set,
                                        "Montando a base da capa " + tamanho + " a partir do PSD (cerca de 1 minuto)...")
                        try:
                            rf.gerar_previa_capa(cfg, [tamanho])
                        except Exception as e:
                            erros.append("Base " + tamanho + ": " + str(e))

                presets_por_tamanho = {}
                for tamanho, simbolos in simbolos_por_tamanho.items():
                    cursos_rgb = {}
                    for grupo in grupos.values():
                        amostra = grupo["amostra"]
                        if (amostra and amostra["tamanho_capa"] == tamanho
                                and grupo["simbolo"] in simbolos):
                            rgb = None if tamanho == "24X34" else grupo["cor"]
                            cursos_rgb[grupo["simbolo"]] = rgb
                    try:
                        presets_por_tamanho[tamanho] = rf.gerar_presets_curso(tamanho, cursos_rgb)
                    except Exception as e:
                        erros.append("Presets " + tamanho + ": " + str(e))
                        presets_por_tamanho[tamanho] = {}

                for grupo in grupos.values():
                    amostra = grupo["amostra"]
                    if not amostra:
                        continue
                    try:
                        preset = presets_por_tamanho.get(amostra["tamanho_capa"], {}).get(grupo["simbolo"])
                        grupo["preset"] = preset
                        if preset:
                            brasao = preset["frente"]
                            caixa_brasao = brasao.getchannel("A").getbbox()
                            grupo["brasao"] = brasao.crop(caixa_brasao) if caixa_brasao else None

                        arquivo = Path(amostra["foto"])
                        origem_foto = rf.fonte_original(arquivo, cfg)
                        imagem, _, _ = rf.abrir_foto(origem_foto)
                        tamanho_foto = amostra["tamanho_foto"]
                        molde = cfg["moldes"][tamanho_foto]
                        caixa = rf.caixa_manual(arquivo, tamanho_foto, imagem.size, molde)
                        if caixa is None:
                            if rf.no_molde(imagem.size, molde):
                                caixa = rf.ajustar_caixa((0, 0, imagem.width, imagem.height), imagem.size, molde)
                            else:
                                rosto = rf.detectar_rosto_cache(origem_foto, imagem) if hasattr(rf, "detectar_rosto_cache") else rf.detectar_rosto(imagem)
                                caixa = (
                                    rf.calcular_recorte(imagem.size, rosto, molde, calibracao.get(tamanho_foto, rf.CALIBRACAO_PADRAO))
                                    if rosto else rf.ajustar_caixa((0, 0, imagem.width, imagem.height), imagem.size, molde)
                                )
                        foto_recortada = imagem.crop(tuple(round(c) for c in caixa))
                        if preset:
                            capa = rf.montar_previa_preset(foto_recortada, preset, amostra["nome"])
                            capa.thumbnail((LADO_PREVIA_REVISAO, LADO_PREVIA_REVISAO), Image.LANCZOS)
                            grupo["capa"] = capa
                        else:
                            previa = rf.ler_previa_capa(amostra["tamanho_capa"])
                            if previa:
                                capa = rf.montar_previa_capa(
                                    foto_recortada, previa, amostra["nome"], amostra["tamanho_capa"])
                                capa.thumbnail((LADO_PREVIA_REVISAO, LADO_PREVIA_REVISAO), Image.LANCZOS)
                                grupo["capa"] = capa
                            else:
                                grupo["avisos"].append(
                                    "Preset de " + amostra["tamanho_capa"] + " indisponível"
                                )
                    except Exception as e:
                        grupo["avisos"].append("Prévia indisponível: " + str(e))

                self.root.after(0, self._revisao_pronta, grupos, erros, cursos_da_foto)
            except Exception as e:
                self.root.after(0, self._revisao_pronta, {}, [str(e)], {})

        threading.Thread(target=trabalho, daemon=True).start()

    def _revisao_pronta(self, grupos, erros, cursos_da_foto=None):

        self._revisao_carregando = False
        if self._revisao_pendente:
            self._revisao_pendente = False
            self.agendar_revisao(100)
        self._revisao_por_curso = grupos
        self.cursos_por_foto = cursos_da_foto or {}
        if self.arquivo:
            self._agendar_previa()
        self.lista_revisao.delete(*self.lista_revisao.get_children())
        self._revisao_imagens = {}
        linhas = {}

        for indice, (chave_grupo, grupo) in enumerate(
                sorted(grupos.items(), key=lambda item: item[1]["curso"].casefold())):
            iid = "curso_" + str(indice)
            brasao = grupo["brasao"]
            if brasao:
                miniatura = brasao.copy()
                miniatura.thumbnail((48, 48), Image.LANCZOS)
                celula = Image.new("RGBA", (52, 52), (0, 0, 0, 0))
                celula.alpha_composite(
                    miniatura, ((52 - miniatura.width) // 2, (52 - miniatura.height) // 2))
                imagem = ImageTk.PhotoImage(celula)
                self._revisao_imagens[iid] = imagem
            else:
                imagem = ""
            cor = grupo["cor"]
            cor_texto = "#%02X%02X%02X" % tuple(cor) if cor else "Sem cor"
            self.lista_revisao.insert(
                "", "end", iid=iid, text="", image=imagem,
                values=(grupo["curso"], len(grupo["entradas"]), cor_texto, self._texto_nomes_grupo(grupo)),
            )
            linhas[iid] = chave_grupo

        self._revisao_grupo_por_iid = linhas
        if grupos:
            itens = self.lista_revisao.get_children()
            alvo = next(
                (iid for iid, chave_grupo in linhas.items()
                 if chave_grupo == self._curso_revisao_selecionado),
                itens[0],
            )
            self.lista_revisao.selection_set(alvo)
            self.lista_revisao.see(alvo)
            self._selecionar_curso_revisao()
        else:
            self._preencher_alunos_revisao()
            self._mostrar_previa_revisao(None)
            self.curso_revisao_nome.set("Nenhum curso encontrado")
            self.curso_revisao_cor.set("Atribua cursos às fotos ou carregue uma planilha com a coluna CURSO.")
            self.status_revisao.set("; ".join(erros) if erros else "Nenhum registro para revisar.")

        self._atualizar_foto_revisao_atual()
        if erros:
            self.status_revisao.set("Avisos: " + " | ".join(erros[:3]))
        elif grupos:
            total = sum(len(grupo["entradas"]) for grupo in grupos.values())
            self.status_revisao.set(str(len(grupos)) + " curso(s), " + str(total) + " registro(s).")
        if self._excluidos_revisao and not erros:
            self.status_revisao.set(
                self.status_revisao.get() + "  |  " + str(len(self._excluidos_revisao))
                + " excluído(s) da revisão (\"Atualizar revisão\" mostra de novo).")

    def _atualizar_revisao_manual(self):

        self._excluidos_revisao.clear()
        self._recarregar_revisao()

    def _excluir_da_revisao(self):
        """Tira da revisão o aluno selecionado ou, sem aluno selecionado, o curso inteiro."""

        if self._aluno_nome:
            _, entrada = self._aluno_nome
            self._excluidos_revisao.add(("aluno", entrada["nome"], entrada["tamanho_capa"]))
            for iid in self.lista_alunos.selection():
                self.lista_alunos.delete(iid)
            self._aluno_escolhido = None
            texto = entrada["nome"] + " (" + entrada["tamanho_capa"] + ")"
        else:
            selecao = self.lista_revisao.selection()
            chave_grupo = self._revisao_grupo_por_iid.get(selecao[0]) if selecao else None
            grupo = self._revisao_por_curso.get(chave_grupo)
            if not grupo:
                self.status_revisao.set("Selecione um curso ou um aluno para excluir da revisão.")
                return
            self._excluidos_revisao.add(("curso", chave_grupo))
            self.lista_revisao.delete(selecao[0])
            self.lista_alunos.delete(*self.lista_alunos.get_children())
            self._curso_revisao_selecionado = None
            texto = "O curso " + grupo["curso"]
        self._limpar_painel_nome()
        self._mostrar_previa_revisao(None)
        self.status_revisao.set(texto + " saiu da revisão. Remontando...")
        self._recarregar_revisao()

    def _selecionar_curso_revisao(self):

        selecao = self.lista_revisao.selection()
        if not selecao:
            return
        chave_grupo = self._revisao_grupo_por_iid.get(selecao[0])
        grupo = self._revisao_por_curso.get(chave_grupo)
        if not grupo:
            return
        self._curso_revisao_selecionado = chave_grupo
        self.curso_revisao_nome.set(grupo["curso"])
        amostra = grupo.get("amostra")
        if amostra and amostra.get("foto"):
            self.foto_revisao_atual.set(Path(amostra["foto"]).name)
        if grupo["simbolo"] and grupo["curso"] in self.cursos_disponiveis:
            self.curso_para_foto.set(grupo["curso"])
        else:
            self.curso_para_foto.set("")
        tamanho_amostra = (grupo.get("amostra") or {}).get("tamanho_capa")
        cor = None if tamanho_amostra == "24X34" else grupo["cor"]
        if tamanho_amostra == "24X34":
            self.curso_revisao_cor.set("Fundo original do PSD 24X34")
            self.amostra_cor_revisao.configure(background=COR_CAMPO)
        elif cor:
            cor_hex = "#%02X%02X%02X" % tuple(cor)
            self.curso_revisao_cor.set("Cor do fundo: " + cor_hex)
            self.amostra_cor_revisao.configure(background=cor_hex)
        else:
            self.curso_revisao_cor.set("Cor de fundo não definida")
            self.amostra_cor_revisao.configure(background=COR_CAMPO)

        self._mostrar_previa_revisao(grupo["capa"])
        self._preencher_alunos_revisao()

        if grupo["avisos"]:
            self.status_revisao.set(" | ".join(dict.fromkeys(grupo["avisos"]))[:500])

    def _mostrar_previa_revisao(self, capa):

        self._capa_revisao_atual = capa
        if capa is None:
            self.previa_revisao_capa.configure(image="")
            self._img_revisao_capa = None
            return
        largura = max(200, self.quadro_previa_revisao.winfo_width() - 4)
        altura = max(150, self.quadro_previa_revisao.winfo_height() - 4)
        imagem = capa.copy()
        imagem.thumbnail((largura, altura), Image.LANCZOS)
        self._img_revisao_capa = ImageTk.PhotoImage(imagem)
        self.previa_revisao_capa.configure(image=self._img_revisao_capa)

    # --------------------------------------------------------
    # Nome na capa (revisão aluno por aluno)
    # --------------------------------------------------------

    @staticmethod
    def _situacao_nome(layout):
        """(texto, tag) da situação do nome na capa."""

        if not layout:
            return "sem modelo", ""
        if layout.get("estoura"):
            return "estoura a área", "nome_estoura"
        if layout.get("manual"):
            return "ajustado à mão", "nome_ajustado"
        if layout.get("reduzido_auto"):
            partes = [str(layout["escala%d" % i]) + "%" for i in (1, 2) if layout["escala%d" % i] != 100]
            return ("reduzido " + " / ".join(partes)) if partes else "lombada reduzida", "nome_reduzido"
        return "ok", ""

    def _texto_nomes_grupo(self, grupo):

        situacoes = [self._situacao_nome(e.get("layout"))[1] for e in grupo["entradas"]]
        estouram = situacoes.count("nome_estoura")
        reduzidos = situacoes.count("nome_reduzido")
        if estouram:
            return str(estouram) + " estoura(m)"
        return str(reduzidos) + " reduzido(s)" if reduzidos else "ok"

    def _preencher_alunos_revisao(self):

        self.lista_alunos.delete(*self.lista_alunos.get_children())
        self._alunos_por_iid = {}
        if self.so_nomes_aviso.get():
            pares = [
                (grupo, entrada)
                for grupo in self._revisao_por_curso.values() for entrada in grupo["entradas"]
                if self._situacao_nome(entrada.get("layout"))[1] in ("nome_estoura", "nome_reduzido")
            ]
        else:
            grupo = self._revisao_por_curso.get(self._curso_revisao_selecionado)
            pares = [(grupo, entrada) for entrada in grupo["entradas"]] if grupo else []
        pares.sort(key=lambda par: par[1]["nome"].casefold())
        for indice, (grupo, entrada) in enumerate(pares):
            texto, tag = self._situacao_nome(entrada.get("layout"))
            iid = "aluno_" + str(indice)
            self.lista_alunos.insert(
                "", "end", iid=iid, values=(entrada["nome"], entrada["tamanho_capa"], texto),
                tags=(tag,) if tag else (),
            )
            self._alunos_por_iid[iid] = (grupo, entrada)
        self._limpar_painel_nome()
        alvo = next((iid for iid, (_, entrada) in self._alunos_por_iid.items()
                     if (entrada["nome"], entrada["tamanho_capa"]) == self._aluno_escolhido), None)
        if alvo:
            self.lista_alunos.selection_set(alvo)
            self.lista_alunos.see(alvo)

    def _limpar_painel_nome(self):

        self._aluno_nome = None
        self._previa_aluno = None
        self._nome_editado = False
        self._preenchendo_nome = True
        self.corte_nome.set("")
        self.escala1_nome.set("100")
        self.escala2_nome.set("100")
        self._preenchendo_nome = False
        for widget in (self.combo_corte_nome, self.spin_escala1, self.spin_escala2,
                       self.botao_nome_auto, self.botao_nome_salvar):
            widget.configure(state="disabled")
        self.nome_revisao_titulo.set("Selecione um aluno na lista para ajustar o nome.")
        self.status_nome.set("")

    def _preencher_controles_nome(self, entrada):

        rf = self.rf
        ajuste = rf.ler_ajuste_nome(entrada["nome"], entrada["tamanho_capa"])
        layout = rf.calcular_layout_nome(entrada["tamanho_capa"], entrada["nome"], ajuste)
        palavras = rf.palavras_nome(entrada["nome"])
        opcoes = [" ".join(palavras[:i]) + "   |   " + " ".join(palavras[i:]) for i in range(1, len(palavras))]
        self._preenchendo_nome = True
        self.combo_corte_nome.configure(values=opcoes, state="readonly" if opcoes else "disabled")
        if opcoes:
            self.combo_corte_nome.current(max(0, min(len(opcoes) - 1, layout["corte"] - 1)))
        else:
            self.corte_nome.set(entrada["nome"])
        self.escala1_nome.set(str(layout["escala1"]))
        self.escala2_nome.set(str(layout["escala2"]))
        self._preenchendo_nome = False
        self._escala_tocada = {1: "escala1" in ajuste, 2: "escala2" in ajuste}
        self._nome_editado = False
        for widget in (self.spin_escala1, self.spin_escala2, self.botao_nome_auto, self.botao_nome_salvar):
            widget.configure(state="normal")

    def _selecionar_aluno_revisao(self):

        selecao = self.lista_alunos.selection()
        if not selecao:
            return
        grupo, entrada = self._alunos_por_iid.get(selecao[0], (None, None))
        if not entrada:
            return
        rf = self.rf
        self._aluno_nome = (grupo, entrada)
        self._aluno_escolhido = (entrada["nome"], entrada["tamanho_capa"])
        self._previa_aluno = None
        self._preencher_controles_nome(entrada)
        self.curso_revisao_nome.set(grupo["curso"])
        self.nome_revisao_titulo.set(
            entrada["nome"] + "  -  capa " + entrada["tamanho_capa"] + "  (" + grupo["curso"] + ")")
        self.foto_revisao_atual.set(
            Path(entrada["foto"]).name if entrada.get("foto") else "Sem foto: prévia com fundo neutro")
        self.status_nome.set("Montando a prévia desta capa...")
        cfg = dict(self.cfg)
        calibracao = dict(self._calibracao())

        def trabalho():
            recortada = preset = previa = erro = None
            try:
                tamanho = entrada["tamanho_capa"]
                simbolo = grupo.get("simbolo")
                if simbolo:
                    cor = None if tamanho == "24X34" else grupo.get("cor")
                    preset = rf.gerar_presets_curso(
                        tamanho, {simbolo: tuple(cor) if cor else None}).get(simbolo)
                if preset is None:
                    previa = rf.ler_previa_capa(tamanho)
                if entrada.get("foto"):
                    recortada = self._recorte_revisao(
                        rf, Path(entrada["foto"]), entrada["tamanho_foto"], cfg, calibracao)
                else:
                    recortada = Image.new("RGB", (400, 300), "#808080")
            except Exception as e:
                erro = e
            self.root.after(0, self._capa_aluno_pronta, entrada, recortada, preset, previa, erro)

        threading.Thread(target=trabalho, daemon=True).start()

    @staticmethod
    def _recorte_revisao(rf, arquivo, tamanho_foto, cfg, calibracao):

        origem_foto = rf.fonte_original(arquivo, cfg)
        imagem, _, _ = rf.abrir_foto(origem_foto)
        molde = cfg["moldes"][tamanho_foto]
        caixa = rf.caixa_manual(arquivo, tamanho_foto, imagem.size, molde)
        if caixa is None:
            if rf.no_molde(imagem.size, molde):
                caixa = rf.ajustar_caixa((0, 0, imagem.width, imagem.height), imagem.size, molde)
            else:
                rosto = rf.detectar_rosto_cache(origem_foto, imagem) if hasattr(rf, "detectar_rosto_cache") else rf.detectar_rosto(imagem)
                caixa = (
                    rf.calcular_recorte(imagem.size, rosto, molde, calibracao.get(tamanho_foto, rf.CALIBRACAO_PADRAO))
                    if rosto else rf.ajustar_caixa((0, 0, imagem.width, imagem.height), imagem.size, molde)
                )
        recortada = imagem.crop(tuple(round(c) for c in caixa)).convert("RGB")
        recortada.thumbnail((1400, 1400), Image.LANCZOS)
        return recortada

    def _capa_aluno_pronta(self, entrada, recortada, preset, previa, erro):

        if not self._aluno_nome or self._aluno_nome[1] is not entrada:
            return
        if erro:
            self.status_nome.set("Prévia indisponível: " + str(erro))
            return
        if preset is None and previa is None:
            self.status_nome.set("A base da capa " + entrada["tamanho_capa"]
                                 + " ainda não foi montada: clique em \"Atualizar revisão\".")
            return
        self._previa_aluno = (recortada, preset, previa)
        self._renderizar_nome_revisao()

    def _ajuste_dos_controles(self):

        ajuste = {}
        indice = self.combo_corte_nome.current()
        if indice >= 0:
            ajuste["corte"] = indice + 1
        for linha, variavel in ((1, self.escala1_nome), (2, self.escala2_nome)):
            if not self._escala_tocada[linha]:
                continue
            try:
                valor = int(round(float(variavel.get().replace(",", "."))))
            except ValueError:
                continue
            ajuste["escala%d" % linha] = max(50, min(130, valor))
        return ajuste

    def _nome_mudou(self):

        if self._preenchendo_nome or not self._aluno_nome:
            return
        self._nome_editado = True
        if self._render_nome_pendente:
            self.root.after_cancel(self._render_nome_pendente)
        self._render_nome_pendente = self.root.after(150, self._renderizar_nome_revisao)

    def _escala_nome_mudou(self, linha):

        if self._preenchendo_nome or not self._aluno_nome:
            return
        self._escala_tocada[linha] = True
        self._nome_mudou()

    def _renderizar_nome_revisao(self):

        self._render_nome_pendente = None
        if not self._aluno_nome or not self._previa_aluno:
            return
        rf = self.rf
        _, entrada = self._aluno_nome
        recortada, preset, previa = self._previa_aluno
        tamanho = entrada["tamanho_capa"]
        ajuste = self._ajuste_dos_controles()
        if preset:
            capa = rf.montar_previa_preset(recortada, preset, entrada["nome"], ajuste)
        else:
            capa = rf.montar_previa_capa(recortada, previa, entrada["nome"], tamanho, ajuste)
        layout = rf.calcular_layout_nome(tamanho, entrada["nome"], ajuste)

        self._preenchendo_nome = True
        for linha, variavel in ((1, self.escala1_nome), (2, self.escala2_nome)):
            if not self._escala_tocada[linha]:
                variavel.set(str(layout["escala%d" % linha]))
        self._preenchendo_nome = False
        self._mostrar_previa_revisao(capa)

        partes = ["1ª linha " + str(layout["escala1"]) + "%", "2ª linha " + str(layout["escala2"]) + "%"]
        if layout["escala_lombada"] != 100:
            partes.append("lombada " + str(layout["escala_lombada"]) + "%")
        if layout["estoura"]:
            linhas = " e ".join(str(n) + "ª" for n in layout["linhas_estouradas"])
            situacao = "ESTOURA a área do nome: diminua a " + linhas + " linha ou mude o corte."
        elif layout["reduzido_auto"] and not any(self._escala_tocada.values()):
            situacao = "Reduzido automaticamente para caber."
        else:
            situacao = "Cabe na área do nome."
        self.status_nome.set(" | ".join(partes) + "  -  " + situacao
                             + ("  (alterações não salvas)" if self._nome_editado else ""))

    def _atualizar_linha_aluno(self, grupo, entrada):

        texto, tag = self._situacao_nome(entrada.get("layout"))
        for iid, (_, outra) in self._alunos_por_iid.items():
            if outra is entrada and self.lista_alunos.exists(iid):
                self.lista_alunos.set(iid, "situacao", texto)
                self.lista_alunos.item(iid, tags=(tag,) if tag else ())
        for iid, chave_grupo in self._revisao_grupo_por_iid.items():
            if self._revisao_por_curso.get(chave_grupo) is grupo and self.lista_revisao.exists(iid):
                self.lista_revisao.set(iid, "nomes", self._texto_nomes_grupo(grupo))

    def _salvar_ajuste_nome(self):

        if not self._aluno_nome:
            return
        rf = self.rf
        grupo, entrada = self._aluno_nome
        tamanho = entrada["tamanho_capa"]
        ajuste = self._ajuste_dos_controles()
        corte_auto = rf.corte_automatico_nome(rf.palavras_nome(entrada["nome"]))
        if ajuste.get("corte") == corte_auto and not any(campo.startswith("escala") for campo in ajuste):
            ajuste = {}
        try:
            rf.gravar_ajuste_nome(entrada["nome"], tamanho, ajuste)
        except OSError as e:
            self.status_nome.set("Não foi possível salvar: " + str(e))
            return
        self._conhecer_arquivo("ajustes_nome", rf.AJUSTES_NOME_FILE)
        entrada["layout"] = rf.calcular_layout_nome(tamanho, entrada["nome"], rf.ler_ajuste_nome(entrada["nome"], tamanho))
        self._nome_editado = False
        self._atualizar_linha_aluno(grupo, entrada)
        self._renderizar_nome_revisao()
        if self.arquivo:
            self._agendar_previa()
        self.status_revisao.set(
            "Nome de " + entrada["nome"] + " salvo para a capa " + tamanho
            + (": o Photoshop vai usar este corte e tamanho." if ajuste else ": volta ao automático."))

    def _ajuste_nome_automatico(self):

        if not self._aluno_nome:
            return
        rf = self.rf
        grupo, entrada = self._aluno_nome
        tamanho = entrada["tamanho_capa"]
        try:
            rf.gravar_ajuste_nome(entrada["nome"], tamanho, {})
        except OSError as e:
            self.status_nome.set("Não foi possível salvar: " + str(e))
            return
        self._conhecer_arquivo("ajustes_nome", rf.AJUSTES_NOME_FILE)
        entrada["layout"] = rf.calcular_layout_nome(tamanho, entrada["nome"], {})
        self._preencher_controles_nome(entrada)
        self._atualizar_linha_aluno(grupo, entrada)
        self._renderizar_nome_revisao()
        if self.arquivo:
            self._agendar_previa()
        self.status_revisao.set("Nome de " + entrada["nome"] + " voltou ao corte e tamanho automáticos.")

    def _excluir_previa_revisao(self):
        """Apaga a prévia em cache do curso/capa selecionado e remonta a revisão."""

        from tkinter import messagebox

        selecao = self.lista_revisao.selection()
        grupo = self._revisao_por_curso.get(self._revisao_grupo_por_iid.get(selecao[0])) if selecao else None
        if not grupo:
            self.status_revisao.set("Selecione um curso na lista para excluir a prévia dele.")
            return
        entrada = self._aluno_nome[1] if self._aluno_nome else grupo.get("amostra")
        tamanho = (entrada or {}).get("tamanho_capa")
        if not tamanho:
            self.status_revisao.set("Este curso não tem prévia montada para excluir.")
            return

        rf = self._modulo()
        if grupo.get("simbolo"):
            if not messagebox.askyesno(
                "Refazer prévia",
                "Apagar a prévia salva de " + grupo["curso"] + " na capa " + tamanho
                + " e montar de novo a partir do PSD?\n\n"
                "Use quando o PSD mudou e a prévia ficou desatualizada.",
            ):
                return
            cor = None if tamanho == "24X34" else grupo.get("cor")
            apagados = rf.excluir_preset_curso(tamanho, grupo["simbolo"], cor)
        else:
            if not messagebox.askyesno(
                "Refazer prévia",
                "Este curso usa a base da capa " + tamanho + ".\n"
                "Apagar a base salva? Ela será lida de novo do PSD (cerca de 1 minuto).",
            ):
                return
            apagados = rf.excluir_previa_capa(tamanho)

        self._presets_curso.clear()
        self._previas_capa = {}
        grupo["capa"] = None
        self._mostrar_previa_revisao(None)
        self._recarregar_revisao()
        self.status_revisao.set(
            str(len(apagados)) + " arquivo(s) da prévia de " + tamanho + " apagado(s). Remontando a revisão...")

    def _atualizar_foto_revisao_atual(self):

        if self.janela.abas.select() == str(self.janela.quadro_revisao):
            selecao = self.lista_revisao.selection()
            if selecao:
                grupo_id = self._revisao_grupo_por_iid.get(selecao[0])
                grupo = self._revisao_por_curso.get(grupo_id)
                amostra = grupo.get("amostra") if grupo else None
                if amostra and amostra.get("foto"):
                    self.foto_revisao_atual.set(Path(amostra["foto"]).name)
                    self.curso_para_foto.set(grupo["curso"] if grupo["simbolo"] else "")
                    return
        if self.arquivo:
            self.foto_revisao_atual.set(self.arquivo.name)
            chave_foto = chave_caminho_foto(self.arquivo)
            curso = self.cursos_manuais.get(chave_foto) or self.cursos_por_foto.get(chave_foto, "")
            self.curso_para_foto.set(curso)
        else:
            self.foto_revisao_atual.set("Selecione uma foto na lista do recorte")
            self.curso_para_foto.set("")

    def _abrir_foto_revisao(self):

        selecao = self.lista_revisao.selection()
        if not selecao:
            return
        chave_grupo = self._revisao_grupo_por_iid.get(selecao[0])
        grupo = self._revisao_por_curso.get(chave_grupo)
        amostra = grupo.get("amostra") if grupo else None
        if self._aluno_nome and self._aluno_nome[1].get("foto"):
            amostra = self._aluno_nome[1]
        if not amostra or not amostra.get("foto"):
            self.status_revisao.set("Não há uma foto correspondente para abrir no recorte.")
            return
        arquivo = Path(amostra["foto"])
        self.janela.abas.select(self.janela.quadro_fotos)
        if self.lista.exists(arquivo.name):
            self.lista.selection_set(arquivo.name)
            self.lista.see(arquivo.name)
        if self.arquivo != arquivo:
            self.abrir(arquivo)

    def _alvo_troca_curso(self):
        """(entrada, é_amostra): o aluno selecionado na lista ou, sem seleção, a amostra do curso."""

        if self._aluno_nome:
            return self._aluno_nome[1], False
        selecao = self.lista_revisao.selection()
        grupo = self._revisao_por_curso.get(self._revisao_grupo_por_iid.get(selecao[0])) if selecao else None
        return (grupo.get("amostra") if grupo else None), True

    def _abrir_atribuicao_curso(self):

        entrada, eh_amostra = self._alvo_troca_curso()
        if not entrada:
            self.status_revisao.set("Selecione um aluno na lista para trocar o curso dele.")
            return
        if not entrada.get("foto"):
            self.status_revisao.set(entrada["nome"] + " não tem foto: o curso é guardado por foto.")
            return
        arquivo = Path(entrada["foto"])

        tk = self.tk
        ttk = self.ttk
        janela = tk.Toplevel(self.root)
        janela.title("Trocar curso")
        janela.transient(self.root)
        janela.resizable(False, False)
        janela.columnconfigure(0, weight=1)

        ttk.Label(janela, text=entrada["nome"], style="Titulo.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(14, 0))
        ttk.Label(janela, text="Foto: " + arquivo.name + "   |   Curso atual: " + (entrada.get("curso") or "nenhum"),
                  style="Dica.TLabel").grid(row=1, column=0, columnspan=3, sticky="w", padx=14)
        if eh_amostra:
            ttk.Label(
                janela, foreground=COR_AMARELO, wraplength=420, justify="left",
                text="Nenhum aluno selecionado: esta é a foto de amostra do curso. "
                     "Para trocar outra pessoa, feche e selecione o aluno na lista de baixo.",
            ).grid(row=2, column=0, columnspan=3, sticky="w", padx=14, pady=(6, 0))
        curso = tk.StringVar(value=entrada.get("curso") or "")
        combo = ttk.Combobox(
            janela, textvariable=curso, state="readonly", values=self.cursos_disponiveis, width=48,
        )
        combo.grid(row=3, column=0, columnspan=3, sticky="we", padx=14, pady=(10, 0))
        combo.focus_set()

        def atribuir():
            if self._salvar_curso_foto(arquivo, curso.get()):
                self._aluno_escolhido = (entrada["nome"], entrada["tamanho_capa"])
                try:
                    automacao = modulo_automacao()
                    canonico = automacao.CURSO_ALIASES.get(automacao.chave(curso.get()), curso.get())
                    self._curso_revisao_selecionado = automacao.chave(canonico)
                except Exception:
                    pass
                janela.destroy()

        def remover():
            if self._remover_curso_arquivo(arquivo):
                janela.destroy()

        ttk.Button(janela, text="Trocar", style="Primario.TButton", command=atribuir).grid(
            row=4, column=0, sticky="w", padx=14, pady=14)
        ttk.Button(janela, text="Voltar ao curso da lista", command=remover).grid(
            row=4, column=1, sticky="w", pady=14)
        ttk.Button(janela, text="Cancelar", command=janela.destroy).grid(
            row=4, column=2, sticky="e", padx=14, pady=14)
        janela.bind("<Escape>", lambda evento: janela.destroy())

    def _salvar_curso_foto(self, arquivo, curso=None):

        curso = str(curso if curso is not None else self.curso_para_foto.get()).strip()
        if not curso:
            self.status_revisao.set("Escolha um curso para atribuir à foto.")
            return False
        try:
            simbolo, _ = modulo_automacao().resolver_curso(curso)
            if not simbolo:
                raise ValueError("O curso não resolve para um brasão.")
            self.cursos_manuais[chave_caminho_foto(arquivo)] = curso
            self._gravar_cursos_foto()
        except Exception as e:
            self.status_revisao.set("Não foi possível atribuir o curso: " + str(e))
            return False
        self.status_revisao.set("Curso " + curso + " salvo para " + arquivo.name + ".")
        self._recarregar_revisao()
        return True

    def _remover_curso_arquivo(self, arquivo):

        removido = self.cursos_manuais.pop(chave_caminho_foto(arquivo), None)
        if removido is None:
            self.status_revisao.set("A foto não tem uma atribuição manual para remover.")
            return False
        try:
            self._gravar_cursos_foto()
        except OSError as e:
            self.cursos_manuais[chave_caminho_foto(arquivo)] = removido
            self.status_revisao.set("Não foi possível salvar a alteração: " + str(e))
            return False
        self.curso_para_foto.set("")
        self.status_revisao.set("Atribuição manual removida de " + arquivo.name + ".")
        self._recarregar_revisao()
        return True

    # --------------------------------------------------------
    # Lista
    # --------------------------------------------------------

    def atualizar_lista(self, silencioso=False):

        if self.ocupado:
            return

        try:
            rf = self._modulo()
        except Exception as e:
            self.status.set("Não foi possível carregar recortar_fotos.py: " + str(e))
            return

        pasta = self.cfg["fotos"]
        if not silencioso:
            self.status.set("Lendo a pasta FOTO...")

        def trabalho():
            try:
                if not pasta.is_dir():
                    raise FileNotFoundError("Pasta de fotos não encontrada: " + str(pasta))
                fotos = rf.listar_fotos(pasta)
                manual = rf.ler_manual()
                situacoes = {}
                for f in fotos:
                    tam = rf.tamanho_do_arquivo(f)
                    try:
                        no_molde = rf.no_molde(rf.tamanho_rapido(f), self.cfg["moldes"][tam])
                    except Exception:
                        no_molde = False
                    if no_molde:
                        situacoes[f.name] = SITUACAO_MANUAL if f.stem in manual else SITUACAO_MOLDE
                    else:
                        situacoes[f.name] = SITUACAO_RECORTAR
                self.root.after(0, self._lista_pronta, fotos, situacoes, None, silencioso)
            except Exception as e:
                self.root.after(0, self._lista_pronta, [], {}, e, silencioso)

        threading.Thread(target=trabalho, daemon=True).start()

    def _lista_pronta(self, fotos, situacoes, erro, silencioso=False):

        if erro:
            self.status.set(str(erro))
            return

        antigas = {k: v for k, v in self.situacoes.items() if v in (SITUACAO_SEM_ROSTO, SITUACAO_BAIXA)}
        self.fotos = fotos
        self.situacoes = situacoes
        for nome, sit in antigas.items():
            if self.situacoes.get(nome) == SITUACAO_RECORTAR:
                self.situacoes[nome] = sit
        self._preencher_lista()
        if not silencioso:
            self.status.set(str(len(fotos)) + " foto(s) na pasta. Escolha uma na lista.")

    def _preencher_lista(self):

        rf = self.rf
        if rf is None:
            return

        selecionada = self.arquivo.name if self.arquivo else None
        nome = rf.chave(self.filtro_nome.get()) if self.filtro_nome.get().strip() else ""
        tam_f = self.filtro_tam.get()
        sit_f = self.filtro_sit.get()

        self.lista.delete(*self.lista.get_children())
        contagem = {}
        mostradas = 0
        for f in self.fotos:
            tam = rf.tamanho_do_arquivo(f)
            sit = self.situacoes.get(f.name, "")
            contagem[sit] = contagem.get(sit, 0) + 1
            if nome and nome not in rf.chave(f.stem):
                continue
            if tam_f != "Todos" and tam != tam_f:
                continue
            if sit_f != "Todas" and sit != sit_f:
                continue
            self.lista.insert("", "end", iid=f.name, values=(f.name, tam, sit), tags=(sit,))
            mostradas += 1

        if selecionada and self.lista.exists(selecionada):
            self.lista.selection_set(selecionada)
            self.lista.see(selecionada)

        partes = [str(mostradas) + " de " + str(len(self.fotos))]
        partes += [k + ": " + str(v) for k, v in sorted(contagem.items()) if k]
        self.total.configure(text="  |  ".join(partes))

    def _marcar(self, arquivo, situacao):

        self.situacoes[arquivo.name] = situacao
        if self.lista.exists(arquivo.name):
            valores = list(self.lista.item(arquivo.name, "values"))
            valores[2] = situacao
            self.lista.item(arquivo.name, values=valores, tags=(situacao,))

    def proxima(self):

        itens = self.lista.get_children()
        if not itens:
            return
        atual = self.lista.selection()
        indice = itens.index(atual[0]) + 1 if atual and atual[0] in itens else 0
        if indice < len(itens):
            self.lista.selection_set(itens[indice])
            self.lista.see(itens[indice])
        else:
            self.status.set("Fim da lista.")

    def _abrir_pasta(self):

        try:
            self._modulo()
            os.startfile(str(self.cfg["fotos"]))
        except Exception as e:
            self.status.set("Não foi possível abrir a pasta: " + str(e))

    # --------------------------------------------------------
    # Abrir foto
    # --------------------------------------------------------

    def _abrir_selecionada(self):

        sel = self.lista.selection()
        if not sel or self.ocupado:
            return
        arquivo = self.cfg["fotos"] / sel[0]
        if self.arquivo == arquivo and self.fonte is not None:
            self._atualizar_foto_revisao_atual()
            return
        self.abrir(arquivo)

    def abrir(self, arquivo):

        rf = self._modulo()
        calib = self._calibracao()
        tam = rf.tamanho_do_arquivo(arquivo)
        self._ocupar(True, "Abrindo " + arquivo.name + " e procurando o rosto...")

        def trabalho():
            try:
                origem = rf.fonte_original(arquivo, self.cfg)
                fonte = rf.abrir_foto(origem)
                img = fonte[0]
                reduzida = img.convert("RGB")
                fator = min(1.0, LADO_EDITOR / max(img.size))
                if fator < 1.0:
                    reduzida = reduzida.resize((round(img.width * fator), round(img.height * fator)), Image.BILINEAR)
                rosto = rf.detectar_rosto_cache(origem, img) if hasattr(rf, "detectar_rosto_cache") else rf.detectar_rosto(img)
                self.root.after(0, self._foto_aberta, arquivo, origem, fonte, reduzida, fator, rosto, tam, calib, None)
            except Exception as e:
                self.root.after(0, self._foto_aberta, arquivo, None, None, None, 1.0, None, tam, calib, e)

        threading.Thread(target=trabalho, daemon=True).start()

    def _foto_aberta(self, arquivo, origem, fonte, reduzida, fator, rosto, tam, calib, erro):

        self._ocupar(False)
        if erro:
            self.status.set("Erro ao abrir " + arquivo.name + ": " + str(erro))
            return

        rf = self.rf
        self.arquivo = arquivo
        self.fonte = fonte
        self.reduzida = reduzida
        self.fator_red = fator
        self.rosto = rosto
        self.historico = []
        self._img_tela_chave = None
        self._atualizar_foto_revisao_atual()
        self.tam.set(tam)
        self._mostrar_molde()

        img = fonte[0]
        molde = self.cfg["moldes"][tam]
        self.caixa_auto = (
            rf.calcular_recorte(img.size, rosto, molde, calib[tam]) if rosto
            else rf.ajustar_caixa((0, 0, img.width, img.height), img.size, molde)
        )

        manual = rf.caixa_manual(arquivo, tam, img.size, molde)
        ja_no_molde = rf.no_molde(img.size, molde)
        if manual:
            self.caixa = manual
            origem_caixa = "corte salvo à mão"
        elif ja_no_molde:
            self.caixa = rf.ajustar_caixa((0, 0, img.width, img.height), img.size, molde)
            origem_caixa = "foto já no tamanho do molde (sem original guardado)"
        else:
            self.caixa = self.caixa_auto
            origem_caixa = "automático pelo rosto" if rosto else "sem rosto: caixa centralizada"

        if not rosto and self.situacoes.get(arquivo.name) == SITUACAO_RECORTAR:
            self._marcar(arquivo, SITUACAO_SEM_ROSTO)
        elif (self.situacoes.get(arquivo.name) == SITUACAO_RECORTAR
              and rf.ampliacao_da_caixa(self.caixa_auto, molde) > self.cfg["ampliacao_maxima"]):
            self._marcar(arquivo, SITUACAO_BAIXA)

        usando_original = origem != arquivo
        self.status.set(
            arquivo.name + "  |  " + str(img.width) + "x" + str(img.height)
            + ("  (original de _ORIGINAIS)" if usando_original else "")
            + "  |  " + origem_caixa
        )
        self._sincronizar_zoom()
        self._desenhar()
        self._agendar_previa()
        self.tela.focus_set()

    # --------------------------------------------------------
    # Geometria
    # --------------------------------------------------------

    def _molde(self):
        return self.cfg["moldes"][self.tam.get()]

    def _tamanho_fonte(self):
        return self.fonte[0].size

    def _definir_caixa(self, caixa, guardar=True):

        if guardar and self.caixa:
            self.historico.append(self.caixa)
            del self.historico[:-50]
        self.caixa = self.rf.ajustar_caixa(caixa, self._tamanho_fonte(), self._molde())
        self._sincronizar_zoom()
        self._desenhar()
        self._agendar_previa()

    def _altura_maxima(self):
        W, H = self._tamanho_fonte()
        mw, mh = self._molde()
        return min(H, W * mh / mw)

    def _sincronizar_zoom(self):
        if self.caixa:
            self.zoom.set(100.0 * (self.caixa[3] - self.caixa[1]) / self._altura_maxima())

    def _redimensionar(self, nova_altura, guardar=True):
        l, t, r, b = self.caixa
        cx, cy = (l + r) / 2, (t + b) / 2
        mw, mh = self._molde()
        h = max(20.0, min(nova_altura, self._altura_maxima()))
        w = h * mw / mh
        self._definir_caixa((cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2), guardar)

    def _para_tela(self, x, y):
        return self.offset[0] + x * self.escala_tela, self.offset[1] + y * self.escala_tela

    def _para_foto(self, x, y):
        return (x - self.offset[0]) / self.escala_tela, (y - self.offset[1]) / self.escala_tela

    # --------------------------------------------------------
    # Desenho do editor
    # --------------------------------------------------------

    def _desenhar(self):

        tela = self.tela
        tela.delete("all")
        if not self.fonte or not self.caixa:
            return

        cw, ch = max(50, tela.winfo_width()), max(50, tela.winfo_height())
        W, H = self._tamanho_fonte()
        s = min((cw - 20) / W, (ch - 20) / H)
        self.escala_tela = s
        dw, dh = max(1, round(W * s)), max(1, round(H * s))
        self.offset = ((cw - dw) / 2, (ch - dh) / 2)

        chave_img = (self.arquivo, dw, dh)
        if self._img_tela_chave != chave_img:
            self._img_tela = ImageTk.PhotoImage(self.reduzida.resize((dw, dh), Image.BILINEAR))
            self._img_tela_chave = chave_img
        tela.create_image(self.offset[0], self.offset[1], image=self._img_tela, anchor="nw")

        x0, y0 = self.offset
        x1, y1 = x0 + dw, y0 + dh
        l, t = self._para_tela(self.caixa[0], self.caixa[1])
        r, b = self._para_tela(self.caixa[2], self.caixa[3])

        # escurece o que fica fora do corte
        for ret in ((x0, y0, x1, t), (x0, b, x1, y1), (x0, t, l, b), (r, t, x1, b)):
            if ret[2] > ret[0] and ret[3] > ret[1]:
                tela.create_rectangle(*ret, fill="black", stipple="gray50", outline="")

        w, h = r - l, b - t

        # terços
        for i in (1, 2):
            tela.create_line(l + w * i / 3, t, l + w * i / 3, b, fill="#bbbbbb", dash=(3, 5))
            tela.create_line(l, t + h * i / 3, r, t + h * i / 3, fill="#bbbbbb", dash=(3, 5))

        # margem de segurança (5%)
        m = 0.05
        tela.create_rectangle(l + w * m, t + h * m, r - w * m, b - h * m, outline="#ff6060", dash=(2, 4))
        # centro e linha dos olhos calibrados para este tamanho
        calib = self._calibracao().get(self.tam.get()) or {}
        cx = l + w * calib.get("centro_x", 0.5)
        tela.create_line(cx, t, cx, b, fill="#00d0ff", dash=(6, 4))
        if "olhos_y" in calib:
            oy = t + h * calib["olhos_y"]
            tela.create_line(l, oy, r, oy, fill="#00d0ff", dash=(6, 4))
            tela.create_text(l + 6, oy - 8, text="olhos", fill="#00d0ff", anchor="w", font=("Segoe UI", 8))

        # rosto detectado
        if self.rosto:
            fl, ft = self._para_tela(self.rosto["x"], self.rosto["y"])
            fr, fb = self._para_tela(self.rosto["x"] + self.rosto["w"], self.rosto["y"] + self.rosto["h"])
            tela.create_rectangle(fl, ft, fr, fb, outline="#40ff40")
            ex, ey = self._para_tela(self.rosto["cx"], self.rosto["olhos_y"])
            tela.create_oval(ex - 3, ey - 3, ex + 3, ey + 3, outline="#40ff40")

        # caixa e alças
        tela.create_rectangle(l, t, r, b, outline="#ffd400", width=2)
        a = ALCA_PX / 2
        for px, py in ((l, t), (r, t), (l, b), (r, b)):
            tela.create_rectangle(px - a, py - a, px + a, py + a, fill="#ffd400", outline="black")

        self._atualizar_info()

    # --------------------------------------------------------
    # Mouse e teclado
    # --------------------------------------------------------

    def _pressionar(self, e):

        self.tela.focus_set()
        if not self.caixa:
            return
        l, t = self._para_tela(self.caixa[0], self.caixa[1])
        r, b = self._para_tela(self.caixa[2], self.caixa[3])

        cantos = {(l, t): (self.caixa[2], self.caixa[3]), (r, t): (self.caixa[0], self.caixa[3]),
                  (l, b): (self.caixa[2], self.caixa[1]), (r, b): (self.caixa[0], self.caixa[1])}
        for (px, py), ancora in cantos.items():
            if abs(e.x - px) <= ALCA_PX and abs(e.y - py) <= ALCA_PX:
                self._arrasto = {"tipo": "canto", "ancora": ancora, "inicio": self.caixa}
                return

        if l <= e.x <= r and t <= e.y <= b:
            self._arrasto = {"tipo": "mover", "x": e.x, "y": e.y, "inicio": self.caixa}
            return

        # clique fora: centraliza a caixa no ponto
        fx, fy = self._para_foto(e.x, e.y)
        w, h = self.caixa[2] - self.caixa[0], self.caixa[3] - self.caixa[1]
        self._definir_caixa((fx - w / 2, fy - h / 2, fx + w / 2, fy + h / 2))
        self._arrasto = {"tipo": "mover", "x": e.x, "y": e.y, "inicio": self.caixa}

    def _arrastar(self, e):

        if not self._arrasto:
            return
        inicio = self._arrasto["inicio"]

        if self._arrasto["tipo"] == "mover":
            dx = (e.x - self._arrasto["x"]) / self.escala_tela
            dy = (e.y - self._arrasto["y"]) / self.escala_tela
            caixa = (inicio[0] + dx, inicio[1] + dy, inicio[2] + dx, inicio[3] + dy)
            self._mover_para(caixa)
        else:
            ax, ay = self._arrasto["ancora"]
            fx, fy = self._para_foto(e.x, e.y)
            mw, mh = self._molde()
            h = max(abs(fy - ay), abs(fx - ax) * mh / mw, 20.0)
            h = min(h, self._altura_maxima())
            w = h * mw / mh
            l = ax - w if fx < ax else ax
            t = ay - h if fy < ay else ay
            self.caixa = self.rf.ajustar_caixa((l, t, l + w, t + h), self._tamanho_fonte(), self._molde())
            self._sincronizar_zoom()
            self._desenhar()
            self._agendar_previa()

    def _mover_para(self, caixa):
        """Move sem mudar o tamanho, travando nas bordas da foto."""
        W, H = self._tamanho_fonte()
        w, h = caixa[2] - caixa[0], caixa[3] - caixa[1]
        l = min(max(0.0, caixa[0]), W - w)
        t = min(max(0.0, caixa[1]), H - h)
        self.caixa = (int(round(l)), int(round(t)), int(round(l + w)), int(round(t + h)))
        self._desenhar()
        self._agendar_previa()

    def _soltar(self, e):

        mudou = bool(self._arrasto and self._arrasto["inicio"] != self.caixa)
        if mudou:
            self.historico.append(self._arrasto["inicio"])
            del self.historico[:-50]
        self._arrasto = None
        if mudou:
            self._agendar_previa()

    def _rodinha(self, e):
        self._zoom_relativo(0.95 if e.delta > 0 else 1.05)

    def _zoom_relativo(self, fator):
        if self.caixa:
            self._redimensionar((self.caixa[3] - self.caixa[1]) * fator)

    def _zoom_escala(self, valor):
        if not self.caixa or self._arrasto:
            return
        alvo = self._altura_maxima() * float(valor) / 100.0
        if abs(alvo - (self.caixa[3] - self.caixa[1])) > 1:
            self._redimensionar(alvo, guardar=False)

    def _mover_teclado(self, dx, dy, e):
        if not self.caixa:
            return
        passo = (10 if e.state & 0x0001 else 1) / self.escala_tela
        self.historico.append(self.caixa)
        c = self.caixa
        self._mover_para((c[0] + dx * passo, c[1] + dy * passo, c[2] + dx * passo, c[3] + dy * passo))

    # --------------------------------------------------------
    # Ferramentas
    # --------------------------------------------------------

    def auto_rosto(self):

        if not self.fonte:
            return
        if not self.rosto:
            self.status.set("Rosto não encontrado nesta foto: ajuste a caixa à mão.")
            return
        calib = self._calibracao()[self.tam.get()]
        self._definir_caixa(self.rf.calcular_recorte(self._tamanho_fonte(), self.rosto, self._molde(), calib))

    def centralizar_rosto(self):

        if not self.fonte or not self.rosto:
            self.status.set("Rosto não encontrado nesta foto.")
            return
        calib = self._calibracao()[self.tam.get()]
        l, t, r, b = self.caixa
        w, h = r - l, b - t
        nl = self.rosto["cx"] - calib["centro_x"] * w
        nt = self.rosto["olhos_y"] - calib["olhos_y"] * h
        self.historico.append(self.caixa)
        self._mover_para((nl, nt, nl + w, nt + h))

    def desfazer(self):

        if self.historico:
            self.caixa = self.historico.pop()
            self._sincronizar_zoom()
            self._desenhar()
            self._agendar_previa()

    def voltar_automatico(self):

        if self.caixa_auto:
            self._definir_caixa(self.caixa_auto)

    # --------------------------------------------------------
    # Prévia
    # --------------------------------------------------------

    def _agendar_previa(self):

        if self._previa_pendente:
            self.root.after_cancel(self._previa_pendente)
        atraso = 24 if getattr(self, "_arrasto", None) else 75
        self._previa_pendente = self.root.after(atraso, self._atualizar_previa)

    def _previa_capa(self, tam):

        if tam not in self._previas_capa:
            self._previas_capa[tam] = self.rf.ler_previa_capa(tam)
        return self._previas_capa[tam]

    def _capas_do_tamanho(self):

        tam = self.tam.get()
        return getattr(self.rf, "CAPAS_DO_TAMANHO", {}).get(tam, [tam])

    def _capa_atual(self):
        """Capa usada na prévia: a escolhida, se a foto deste tamanho for usada nela."""

        capas = self._capas_do_tamanho()
        if tuple(self.combo_capa.cget("values")) != tuple(capas):
            self.combo_capa.configure(values=capas)
        if self.capa_previa.get() not in capas:
            self.capa_previa.set(capas[0])
        self.combo_capa.configure(state="readonly" if len(capas) > 1 else "disabled")
        return self.capa_previa.get()

    def _atualizar_previa(self):

        self._previa_pendente = None
        if not self.fonte or not self.caixa:
            return

        f = self.fator_red
        caixa_red = tuple(round(c * f) for c in self.caixa)
        mw, mh = self._molde()
        escala = min(1.0, 1400 / max(mw, mh))
        filtro_interp = Image.BILINEAR if getattr(self, "_arrasto", None) else Image.LANCZOS
        recortada = self.reduzida.crop(caixa_red).resize(
            (max(1, round(mw * escala)), max(1, round(mh * escala))), filtro_interp)

        imagem = recortada
        if self.modo_previa.get() == "capa":
            capa = self._capa_atual()
            preset = self._preset_da_foto(capa)
            previa = None if preset else self._previa_capa(capa)
            nome = self.rf.nome_do_aluno(self.arquivo) if self.arquivo else None
            if preset:
                imagem = self.rf.montar_previa_preset(recortada, preset, nome)
            elif previa:
                imagem = self.rf.montar_previa_capa(recortada, previa, nome, capa)
            else:
                self.alerta.set("Prévia da capa ainda não gerada: clique em \"Gerar prévia da capa\".")

        imagem = imagem.copy()
        imagem.thumbnail((380, 380), Image.LANCZOS)
        fundo = Image.new("RGB", (380, 380), COR_FUNDO)
        fundo.paste(imagem, ((380 - imagem.width) // 2, (380 - imagem.height) // 2))
        self._img_previa = ImageTk.PhotoImage(fundo)
        self.rotulo_previa.configure(image=self._img_previa)

    def _atualizar_info(self):

        if not self.caixa:
            return
        rf = self.rf
        mw, mh = self._molde()
        w, h = self.caixa[2] - self.caixa[0], self.caixa[3] - self.caixa[1]
        amp = rf.ampliacao_da_caixa(self.caixa, (mw, mh))
        self.info.set(
            "Corte na foto original: " + str(w) + " x " + str(h) + " px\n"
            + "Foto final: " + str(mw) + " x " + str(mh) + " px\n"
            + "Escala: " + format(amp, ".0%") + ("  (ampliando)" if amp > 1 else "  (reduzindo)")
        )
        alertas = []
        if amp > self.cfg["ampliacao_maxima"]:
            alertas.append("Resolução baixa: a foto vai ser ampliada " + format(amp, ".0%")
                           + " (limite " + format(self.cfg["ampliacao_maxima"], ".0%") + ").")
        tam_nome = rf.tamanho_do_arquivo(self.arquivo)
        if tam_nome != self.tam.get():
            alertas.append("Atenção: o nome do arquivo indica " + tam_nome
                           + ". O Photoshop usa esta foto na capa " + tam_nome + ".")
        if self.modo_previa.get() == "capa":
            capa = self._capa_atual()
            curso, simbolo, rgb = self._dados_curso(capa)
            if not (rf.PREVIA_CAPAS_DIR / (capa + ".json")).is_file():
                alertas.append("Prévia da capa ainda não gerada: clique em \"Gerar prévia da capa\".")
            elif simbolo and (capa, simbolo, rgb) not in self._presets_curso \
                    and not rf.preset_em_cache(capa, simbolo, rgb):
                alertas.append("Clique em \"Gerar prévia da capa\" para ver a cor e os brasões de " + curso + ".")
        self.alerta.set("\n".join(alertas))

    # --------------------------------------------------------
    # Preset do curso (cor + brasões) para a prévia "Na capa"
    # --------------------------------------------------------

    def _curso_da_foto(self):

        if not self.arquivo:
            return ""
        chave_foto = chave_caminho_foto(self.arquivo)
        return self.cursos_manuais.get(chave_foto) or self.cursos_por_foto.get(chave_foto) or ""

    def _dados_curso(self, capa):
        """(curso, símbolo, rgb) do aluno desta foto; símbolo None se o curso não for conhecido."""

        curso = self._curso_da_foto()
        if not curso:
            return "", None, None
        try:
            auto = modulo_automacao()
            simbolo, _ = auto.resolver_curso(curso)
        except Exception:
            return curso, None, None
        if not simbolo:
            return curso, None, None
        rgb = None if capa == "24X34" else auto.obter_cor_background(simbolo)
        return curso, simbolo, (tuple(rgb) if rgb else None)

    def _preset_da_foto(self, capa):
        """Preset pronto do curso da foto; se estiver em cache, carrega em segundo plano."""

        _, simbolo, rgb = self._dados_curso(capa)
        if not simbolo:
            return None
        chave_preset = (capa, simbolo, rgb)
        if chave_preset in self._presets_curso:
            return self._presets_curso[chave_preset]
        if chave_preset not in self._presets_carregando and self.rf.preset_em_cache(capa, simbolo, rgb):
            self._carregar_presets(capa, {simbolo: rgb})
        return None

    def _carregar_presets(self, capa, cursos_rgb, ao_terminar=None):

        chaves = {(capa, s, tuple(c) if c else None) for s, c in cursos_rgb.items()}
        self._presets_carregando |= chaves
        rf = self.rf

        def trabalho():
            erro = None
            presets = {}
            try:
                presets = rf.gerar_presets_curso(capa, cursos_rgb)
            except Exception as e:
                erro = e
            self.root.after(0, self._presets_prontos, capa, cursos_rgb, presets, erro, chaves, ao_terminar)

        threading.Thread(target=trabalho, daemon=True).start()

    def _presets_prontos(self, capa, cursos_rgb, presets, erro, chaves, ao_terminar):

        self._presets_carregando -= chaves
        for simbolo, preset in presets.items():
            cor = cursos_rgb.get(simbolo)
            self._presets_curso[(capa, simbolo, tuple(cor) if cor else None)] = preset
        if ao_terminar:
            ao_terminar(erro)
        self._agendar_previa()

    def gerar_previa_capa(self):
        """Monta a prévia pelo preset do curso da foto: cor do curso + brasões da frente e do verso."""

        from tkinter import messagebox

        rf = self._modulo()
        capa = self._capa_atual()
        curso, simbolo, rgb = self._dados_curso(capa)

        base_pronta = rf.ler_previa_capa(capa) is not None
        if not base_pronta and not messagebox.askyesno(
            "Prévia da capa",
            "A base da capa " + capa + " ainda não foi montada.\n"
            "Na primeira vez eu leio o PSD (cerca de 1 minuto); depois a prévia de cada curso sai na hora.\n\n"
            "Continuar?",
        ):
            return

        self.botao_gerar_previa.configure(state="disabled")
        self.status.set("Montando a prévia da capa " + capa
                        + (" do curso " + curso if simbolo else "") + "...")
        cfg = self.cfg

        def depois(erro):
            self.botao_gerar_previa.configure(state="normal")
            self._previas_capa = {}
            if erro:
                self.status.set("Falha ao montar a prévia: " + str(erro))
            elif simbolo:
                self.status.set("Prévia pronta: capa " + capa + " com cor e brasões de " + curso + ".")
            elif curso:
                self.status.set("Prévia da capa pronta, mas o curso \"" + curso
                                + "\" não tem brasão conhecido.")
            else:
                self.status.set("Prévia da capa pronta, sem curso: atribua o curso desta foto "
                                "na revisão para ver a cor e o brasão.")
            if self.caixa:
                self._atualizar_info()

        def trabalho():
            erro = None
            try:
                if not base_pronta:
                    rf.gerar_previa_capa(cfg, [capa])
            except Exception as e:
                erro = e
            if erro or not simbolo:
                self.root.after(0, depois, erro)
                self.root.after(0, self._agendar_previa)
            else:
                self.root.after(0, self._carregar_presets, capa, {simbolo: rgb}, depois)

        threading.Thread(target=trabalho, daemon=True).start()

    # --------------------------------------------------------
    # Tamanho do molde
    # --------------------------------------------------------

    def _mostrar_molde(self):
        mw, mh = self._molde()
        self.largura.set(str(mw))
        self.altura.set(str(mh))

    def _trocar_tamanho(self):

        self._mostrar_molde()
        if self.fonte and self.caixa:
            self._img_tela_chave = None
            self._definir_caixa(self.caixa)

    def gravar_tamanho(self):

        from tkinter import messagebox

        try:
            w, h = int(self.largura.get()), int(self.altura.get())
            if w < 100 or h < 100:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Tamanho", "Largura e altura precisam ser números em pixels (mínimo 100).")
            return
        tam = self.tam.get()
        if not messagebox.askyesno(
            "Tamanho", "Gravar " + tam + " = " + str(w) + " x " + str(h) + " px?\n\n"
            "Todas as fotos " + tam + " passam a ser gravadas nesse tamanho.",
        ):
            return
        self.rf.gravar_molde(tam, w, h)
        self._modulo()
        if self.fonte and self.caixa:
            self._definir_caixa(self.caixa)
        self.status.set("Tamanho " + tam + " gravado: " + str(w) + " x " + str(h) + " px.")

    def ler_tamanho_psd(self):

        from tkinter import messagebox

        rf = self._modulo()
        tam = self.tam.get()
        arq = rf.PREVIA_CAPAS_DIR / (tam + ".json")
        if not arq.is_file():
            messagebox.showinfo("Tamanho", "Clique antes em \"Gerar prévia da capa\": é ela que lê o PSD.")
            return
        info = json.loads(arq.read_text(encoding="utf-8"))
        so = info.get("smart_object")
        if not so:
            messagebox.showinfo("Tamanho", "Não consegui ler o tamanho do smart object FOTO no PSD.")
            return
        self.largura.set(str(so[0]))
        self.altura.set(str(so[1]))
        atual = self._molde()
        if tuple(so) == tuple(atual):
            self.status.set("O PSD " + tam + " pede " + str(so[0]) + " x " + str(so[1]) + " px: já é o tamanho usado.")
        else:
            self.gravar_tamanho()

    # --------------------------------------------------------
    # Gravar
    # --------------------------------------------------------

    def salvar(self, proxima):

        from tkinter import messagebox

        if not self.fonte or not self.caixa or self.ocupado:
            return
        rf = self.rf
        tam = self.tam.get()
        amp = rf.ampliacao_da_caixa(self.caixa, self._molde())
        if amp > self.cfg["ampliacao_maxima"] and not messagebox.askyesno(
            "Resolução baixa",
            "Esta foto vai ser ampliada " + format(amp, ".0%") + " e pode ficar sem nitidez na capa.\n\n"
            "Salvar mesmo assim?",
        ):
            return

        arquivo, caixa, fonte = self.arquivo, self.caixa, self.fonte
        self._ocupar(True, "Gravando " + arquivo.name + "...")

        def trabalho():
            try:
                destino = rf.aplicar_recorte(arquivo, caixa, tam, self.cfg, fonte)
                rf.gravar_manual(arquivo, tam, caixa, fonte[0].size)
                self.root.after(0, self._salvo, arquivo, destino, proxima, None)
            except Exception as e:
                self.root.after(0, self._salvo, arquivo, None, proxima, e)

        threading.Thread(target=trabalho, daemon=True).start()

    def _salvo(self, arquivo, destino, proxima, erro):

        from tkinter import messagebox

        self._ocupar(False)
        if erro:
            messagebox.showerror("Gravar foto", "Não foi possível gravar:\n\n" + str(erro))
            self.status.set("Erro ao gravar " + arquivo.name)
            return

        mw, mh = self._molde()
        if destino != arquivo:
            curso = self.cursos_manuais.pop(chave_caminho_foto(arquivo), None)
            if curso:
                self.cursos_manuais[chave_caminho_foto(destino)] = curso
                try:
                    self._gravar_cursos_foto()
                except OSError:
                    self.cursos_manuais.pop(chave_caminho_foto(destino), None)
                    self.cursos_manuais[chave_caminho_foto(arquivo)] = curso
                    self.status_revisao.set("O recorte foi salvo, mas não consegui atualizar o curso da nova foto.")
            idx = self.fotos.index(arquivo) if arquivo in self.fotos else None
            if idx is not None:
                self.fotos[idx] = destino
            self.situacoes.pop(arquivo.name, None)
            self.arquivo = destino
            self._preencher_lista()
        self._atualizar_foto_revisao_atual()
        self._marcar(destino, SITUACAO_MANUAL)
        self.status.set("Gravada: " + destino.name + "  (" + str(mw) + " x " + str(mh) + " px)")
        self.agendar_revisao()
        if proxima:
            self.proxima()

    def restaurar(self):

        from tkinter import messagebox

        if not self.arquivo or self.ocupado:
            return
        rf = self.rf
        if rf.backup_de(self.arquivo, self.cfg) is None:
            messagebox.showinfo("Restaurar", "Não há original guardado em _ORIGINAIS para esta foto.")
            return
        if not messagebox.askyesno("Restaurar", "Voltar o original de " + self.arquivo.name
                                   + " para a pasta FOTO e esquecer o corte manual?"):
            return
        try:
            destino = rf.restaurar_original(self.arquivo, self.cfg)
        except Exception as e:
            messagebox.showerror("Restaurar", str(e))
            return
        self.arquivo = None
        self.fonte = None
        self.status.set("Original restaurado: " + destino.name)
        self.atualizar_lista()

    # --------------------------------------------------------
    # Utilidades
    # --------------------------------------------------------

    def _ocupar(self, ocupado, texto=None):

        self.ocupado = ocupado
        estado = "disabled" if ocupado else "normal"
        for b in self.botoes_editor + [self.botao_salvar, self.botao_salvar_prox, self.botao_restaurar]:
            b.configure(state=estado)
        if texto:
            self.status.set(texto)


# ============================================================
# ABA PASTAS E LINKS
# ============================================================

PASTA_FOTOS_DRIVE_PADRAO = PASTA / "FOTO_DRIVE"


class AbaPastas:

    def __init__(self, janela, quadro):

        import tkinter as tk

        self.janela = janela
        self.root = janela.root
        self.quadro = quadro

        self.fotos_origem = tk.StringVar(value="pasta")
        self.fotos_pasta = tk.StringVar()
        self.fotos_link = tk.StringVar()
        self.fotos_drive_dir = tk.StringVar()
        self.sincronizar_antes = tk.BooleanVar(value=True)
        self.backup = tk.StringVar()
        self.psd = tk.StringVar()
        self.saidas = {t: tk.StringVar() for t in SAIDAS}
        self.planilha_dir = tk.StringVar()

        self._montar()
        self.carregar()

    # --------------------------------------------------------
    # Montagem
    # --------------------------------------------------------

    def _linha_pasta(self, pai, linha, texto, variavel, dica=""):

        ttk = self.janela.ttk
        ttk.Label(pai, text=texto).grid(row=linha, column=0, sticky="w", pady=3)
        campo = ttk.Entry(pai, textvariable=variavel, width=90)
        campo.grid(row=linha, column=1, sticky="we", padx=6, pady=3)
        ttk.Button(pai, text="Procurar...", command=lambda: self._procurar(variavel, texto)).grid(
            row=linha, column=2, pady=3)
        if dica:
            ttk.Label(pai, text=dica, style="Dica.TLabel").grid(row=linha + 1, column=1, sticky="w", padx=6)
        return campo

    def _montar(self):

        tk = self.janela.tk
        ttk = self.janela.ttk
        q = self.quadro
        q.columnconfigure(0, weight=1)

        # ---------------- Fotos ----------------

        fotos = ttk.LabelFrame(q, text="Fotos dos alunos", padding=8)
        fotos.grid(row=0, column=0, sticky="we")
        fotos.columnconfigure(1, weight=1)

        origem = ttk.Frame(fotos)
        origem.grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Radiobutton(origem, text="Pasta no computador / rede", value="pasta",
                        variable=self.fotos_origem, command=self._atualizar_estado).pack(side="left")
        ttk.Radiobutton(origem, text="Link de pasta do Google Drive", value="drive",
                        variable=self.fotos_origem, command=self._atualizar_estado).pack(side="left", padx=16)

        self.campo_fotos = self._linha_pasta(fotos, 1, "Pasta das fotos", self.fotos_pasta)

        ttk.Label(fotos, text="Link do Drive").grid(row=2, column=0, sticky="w", pady=3)
        self.campo_link = ttk.Entry(fotos, textvariable=self.fotos_link, width=90)
        self.campo_link.grid(row=2, column=1, sticky="we", padx=6, pady=3)
        self.botao_baixar = ttk.Button(fotos, text="Baixar fotos novas", command=self.baixar_fotos)
        self.botao_baixar.grid(row=2, column=2, pady=3)
        ttk.Label(
            fotos, style="Dica.TLabel",
            text="A pasta do Drive precisa estar compartilhada como \"Qualquer pessoa com o link\".",
        ).grid(row=3, column=1, sticky="w", padx=6)

        self.campo_drive_dir = self._linha_pasta(
            fotos, 4, "Guardar as fotos do Drive em", self.fotos_drive_dir,
            "As fotos são baixadas para esta pasta; o recorte e o Photoshop usam ela. "
            "Só baixa as que ainda não estão lá.",
        )
        self.check_sincronizar = ttk.Checkbutton(
            fotos, text="Baixar as fotos novas do Drive antes de Validar / Gerar / Recortar",
            variable=self.sincronizar_antes,
        )
        self.check_sincronizar.grid(row=6, column=1, sticky="w", padx=6, pady=(4, 0))

        self._linha_pasta(
            fotos, 7, "Originais (_ORIGINAIS)", self.backup,
            "Onde o recorte guarda a foto original. Vazio = pasta _ORIGINAIS dentro da pasta das fotos.",
        )

        # ---------------- Capas ----------------

        capas = ttk.LabelFrame(q, text="Modelos e capas prontas", padding=8)
        capas.grid(row=1, column=0, sticky="we", pady=(10, 0))
        capas.columnconfigure(1, weight=1)
        self._linha_pasta(capas, 0, "PSDs das capas", self.psd)
        for i, tam in enumerate(SAIDAS, start=1):
            dica = ("Frente e verso. Vazio = pasta \"CAPAS 30X22\" ao lado da pasta do 30X20."
                    if tam == "30X22" else "")
            self._linha_pasta(capas, i, "Capas prontas " + tam, self.saidas[tam], dica)

        # ---------------- Planilhas ----------------

        planilhas = ttk.LabelFrame(q, text="Planilhas", padding=8)
        planilhas.grid(row=2, column=0, sticky="we", pady=(10, 0))
        planilhas.columnconfigure(1, weight=1)
        self._linha_pasta(
            planilhas, 0, "Pasta das planilhas", self.planilha_dir,
            "Pasta aberta no \"Escolher...\". O link do Google vai direto no campo da aba Capas.",
        )

        # ---------------- Ações ----------------

        acoes = ttk.Frame(q)
        acoes.grid(row=3, column=0, sticky="we", pady=(10, 0))
        ttk.Button(acoes, text="Salvar", style="Primario.TButton", command=self.salvar).pack(side="left")
        ttk.Button(acoes, text="Conferir caminhos", command=self.conferir).pack(side="left", padx=6)
        ttk.Button(acoes, text="Desfazer alterações", command=self.carregar).pack(side="left")

        self.log = tk.Text(q, height=9, wrap="word", state="disabled",
                           background=COR_FUNDO, foreground=COR_TEXTO)
        self.log.grid(row=4, column=0, sticky="nsew", pady=(8, 0))
        q.rowconfigure(4, weight=1)

    # --------------------------------------------------------
    # Config
    # --------------------------------------------------------

    def carregar(self):

        dados = modulo_recorte().ler_config_bruto()

        self.fotos_origem.set(dados.get("fotos_origem") or "pasta")
        self.fotos_pasta.set(dados.get("photos_local_dir") or (
            dados.get("photos_dir", "") if self.fotos_origem.get() == "pasta" else ""))
        self.fotos_link.set(dados.get("photos_link") or "")
        self.fotos_drive_dir.set(dados.get("photos_drive_dir") or str(PASTA_FOTOS_DRIVE_PADRAO))
        self.sincronizar_antes.set(bool(dados.get("drive_sincronizar_antes", True)))
        self.backup.set(dados.get("recorte_backup_dir") or "")
        self.psd.set(dados.get("psd_dir") or "")
        for tam in SAIDAS:
            self.saidas[tam].set((dados.get("output_dirs") or {}).get(tam, ""))
        self.planilha_dir.set(dados.get("planilha_dir") or "")
        self._atualizar_estado()

    def _atualizar_estado(self):

        drive = self.fotos_origem.get() == "drive"
        self.campo_fotos.configure(state="disabled" if drive else "normal")
        for w in (self.campo_link, self.campo_drive_dir, self.botao_baixar, self.check_sincronizar):
            w.configure(state="normal" if drive else "disabled")

    def _procurar(self, variavel, titulo):

        from tkinter import filedialog

        inicial = variavel.get().strip()
        pasta = filedialog.askdirectory(
            title=titulo, initialdir=inicial if inicial and Path(inicial).is_dir() else str(PASTA),
        )
        if pasta:
            variavel.set(str(Path(pasta)))

    def _valores(self):

        drive = self.fotos_origem.get() == "drive"
        fotos = self.fotos_drive_dir.get().strip() if drive else self.fotos_pasta.get().strip()
        return {
            "fotos_origem": "drive" if drive else "pasta",
            "photos_dir": fotos,
            "photos_local_dir": self.fotos_pasta.get().strip(),
            "photos_link": self.fotos_link.get().strip(),
            "photos_drive_dir": self.fotos_drive_dir.get().strip(),
            "drive_sincronizar_antes": bool(self.sincronizar_antes.get()),
            "recorte_backup_dir": self.backup.get().strip() or None,
            "psd_dir": self.psd.get().strip(),
            "output_dirs": {t: self.saidas[t].get().strip() for t in SAIDAS},
            "planilha_dir": self.planilha_dir.get().strip(),
        }

    def salvar(self):

        from tkinter import messagebox

        valores = self._valores()
        rf = modulo_recorte()

        if valores["fotos_origem"] == "drive":
            if not rf.eh_link(valores["photos_link"]):
                messagebox.showwarning("Pastas e links", "Cole o link da pasta do Google Drive com as fotos.")
                return
            try:
                rf.id_google(valores["photos_link"])
            except ValueError as e:
                messagebox.showwarning("Pastas e links", str(e))
                return
            if not valores["photos_dir"]:
                messagebox.showwarning("Pastas e links", "Escolha a pasta onde guardar as fotos do Drive.")
                return
        elif not valores["photos_dir"]:
            messagebox.showwarning("Pastas e links", "Escolha a pasta das fotos.")
            return

        if self.janela.processo is not None:
            messagebox.showwarning("Pastas e links", "Espere terminar o que está rodando antes de salvar.")
            return

        rf.gravar_config(valores)
        self._escrever("Configuração salva em " + str(CONFIG) + "\n")
        self.conferir()
        self.janela.aba_fotos.atualizar_lista()

    def conferir(self):

        v = self._valores()
        itens = [("Fotos", v["photos_dir"]), ("PSDs das capas", v["psd_dir"])]
        itens += [("Capas prontas " + t, v["output_dirs"][t]) for t in SAIDAS]
        itens += [("Pasta das planilhas", v["planilha_dir"])]
        if v["recorte_backup_dir"]:
            itens.append(("Originais", v["recorte_backup_dir"]))

        self._escrever("\nConferindo caminhos:\n")
        for nome, caminho in itens:
            if not caminho:
                estado = "VAZIO"
            elif Path(caminho).is_dir():
                estado = "OK"
            elif v["fotos_origem"] == "drive" and nome == "Fotos":
                estado = "será criada ao baixar do Drive"
            else:
                estado = "NÃO ENCONTRADA"
            self._escrever(f"  {estado:<30} {nome}: {caminho}\n")
        if v["fotos_origem"] == "drive":
            self._escrever("  Link das fotos: " + (v["photos_link"] or "(vazio)") + "\n")

    # --------------------------------------------------------
    # Drive
    # --------------------------------------------------------

    def baixar_fotos(self):

        from tkinter import messagebox

        link = self.fotos_link.get().strip()
        destino = self.fotos_drive_dir.get().strip()
        rf = modulo_recorte()
        if not rf.eh_link(link) or not destino:
            messagebox.showwarning("Pastas e links", "Preencha o link do Drive e a pasta onde guardar as fotos.")
            return

        self.botao_baixar.configure(state="disabled")
        self._escrever("\nBaixando fotos do Drive...\n")

        def aviso(texto):
            self.root.after(0, self._escrever, texto + "\n")

        def trabalho():
            try:
                rf.sincronizar_drive(link, destino, aviso)
            except Exception as e:
                aviso("ERRO: " + str(e))
            self.root.after(0, self._baixado)

        threading.Thread(target=trabalho, daemon=True).start()

    def _baixado(self):

        self._atualizar_estado()
        if modulo_recorte().ler_config_bruto().get("fotos_origem") == "drive":
            self.janela.aba_fotos.atualizar_lista()

    def _escrever(self, texto):

        self.log.configure(state="normal")
        self.log.insert("end", texto)
        self.log.see("end")
        self.log.configure(state="disabled")


# ============================================================
# ABA CORES DOS CURSOS
# ============================================================

def cor_hex(rgb):
    return "#%02X%02X%02X" % tuple(rgb) if rgb else ""


def cor_do_texto(texto):
    """"#7A1A18", "7A1A18" ou "122, 26, 24" -> (r, g, b); None se inválido."""

    texto = (texto or "").strip()
    numeros = re.findall(r"\d+", texto)
    if "," in texto or (" " in texto and len(numeros) == 3):
        if len(numeros) == 3 and all(0 <= int(n) <= 255 for n in numeros):
            return tuple(int(n) for n in numeros)
        return None
    m = re.fullmatch(r"#?([0-9a-fA-F]{6})", texto)
    if m:
        h = m.group(1)
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    return None


def texto_contraste(rgb):
    r, g, b = rgb
    return "#000000" if (0.299 * r + 0.587 * g + 0.114 * b) > 150 else "#FFFFFF"


class AbaCores:
    """Troca a cor do fundo/preenchimento de cada curso (vale para todas as capas)."""

    def __init__(self, janela, quadro):

        import tkinter as tk

        self.janela = janela
        self.root = janela.root
        self.quadro = quadro

        self.busca = tk.StringVar()
        self.texto_cor = tk.StringVar()
        self.selecionado = None
        self.padrao = {}
        self.nomes = {}
        self.cores = {}
        self.salvas = {}

        self._montar()
        self.root.after(200, self.carregar)

    def _montar(self):

        tk = self.janela.tk
        ttk = self.janela.ttk
        q = self.quadro
        q.columnconfigure(0, weight=1)
        q.rowconfigure(1, weight=1)

        topo = ttk.Frame(q)
        topo.grid(row=0, column=0, columnspan=2, sticky="we")
        ttk.Label(topo, text="Procurar curso").pack(side="left")
        campo = ttk.Entry(topo, textvariable=self.busca, width=30)
        campo.pack(side="left", padx=6)
        self.busca.trace_add("write", lambda *a: self._preencher())
        ttk.Label(
            topo, style="Dica.TLabel",
            text="A cor vale para o fundo/preenchimento do curso em todas as capas (20x20, 30x20, 24x34, 30x22).",
        ).pack(side="left", padx=10)

        lista = ttk.Frame(q)
        lista.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        lista.columnconfigure(0, weight=1)
        lista.rowconfigure(0, weight=1)

        self.tabela = ttk.Treeview(
            lista, columns=("curso", "nomes", "cor", "origem"), show="headings", selectmode="browse"
        )
        for col, titulo, largura in (("curso", "Curso (brasão)", 170), ("nomes", "Nomes na planilha", 380),
                                     ("cor", "Cor", 90), ("origem", "", 110)):
            self.tabela.heading(col, text=titulo)
            self.tabela.column(col, width=largura, anchor="w")
        self.tabela.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(lista, orient="vertical", command=self.tabela.yview)
        barra.grid(row=0, column=1, sticky="ns")
        self.tabela.configure(yscrollcommand=barra.set)
        self.tabela.bind("<<TreeviewSelect>>", lambda e: self._selecionou())
        self.tabela.bind("<Double-1>", lambda e: self.escolher())

        lado = ttk.LabelFrame(q, text="Cor do curso", padding=8)
        lado.grid(row=1, column=1, sticky="ns", padx=(10, 0), pady=(8, 0))

        self.rotulo_curso = ttk.Label(lado, text="Escolha um curso na lista", font=("Segoe UI", 10, "bold"))
        self.rotulo_curso.pack(anchor="w")
        self.amostra = tk.Canvas(lado, width=200, height=110, background=COR_CAMPO,
                                 highlightthickness=1, highlightbackground=COR_BORDA)
        self.amostra.pack(pady=8)
        self.rotulo_padrao = ttk.Label(lado, text="", style="Dica.TLabel")
        self.rotulo_padrao.pack(anchor="w")

        linha = ttk.Frame(lado)
        linha.pack(fill="x", pady=(8, 0))
        ttk.Label(linha, text="Cor:").pack(side="left")
        campo_cor = ttk.Entry(linha, textvariable=self.texto_cor, width=16)
        campo_cor.pack(side="left", padx=4)
        campo_cor.bind("<Return>", lambda e: self.aplicar_texto())
        ttk.Button(linha, text="Aplicar", command=self.aplicar_texto).pack(side="left")
        ttk.Label(lado, text="#7A1A18 ou 122, 26, 24", style="Dica.TLabel").pack(anchor="w")

        self.botao_escolher = ttk.Button(lado, text="Escolher cor...", command=self.escolher)
        self.botao_escolher.pack(fill="x", pady=(10, 0))
        self.botao_padrao = ttk.Button(lado, text="Voltar à cor padrão", command=self.voltar_padrao)
        self.botao_padrao.pack(fill="x", pady=(4, 0))

        ttk.Separator(lado).pack(fill="x", pady=10)
        ttk.Button(lado, text="Salvar cores", style="Primario.TButton", command=self.salvar).pack(fill="x")
        ttk.Button(lado, text="Desfazer alterações", command=self.carregar).pack(fill="x", pady=(4, 0))

        self.status = tk.StringVar(value="")
        ttk.Label(lado, textvariable=self.status, style="Dica.TLabel", wraplength=200).pack(anchor="w", pady=(10, 0))

    # --------------------------------------------------------

    def carregar(self):

        try:
            auto = modulo_automacao()
        except Exception as e:
            self.status.set("Não consegui ler os cursos: " + str(e))
            return

        simbolos = set(auto.CORES_BACKGROUND) | set(auto.CURSOS.values())
        self.padrao = {s: auto.CORES_BACKGROUND.get(s) for s in simbolos}
        self.nomes = {}
        for nome, simbolo in auto.CURSOS.items():
            self.nomes.setdefault(simbolo, []).append(nome)

        self.salvas = {}
        for simbolo, valor in (modulo_recorte().ler_config_bruto().get("cores_cursos") or {}).items():
            rgb = cor_do_texto(valor) if isinstance(valor, str) else (
                tuple(valor) if isinstance(valor, (list, tuple)) and len(valor) == 3 else None)
            if rgb:
                self.salvas[simbolo] = rgb
                self.padrao.setdefault(simbolo, None)
        self.cores = dict(self.salvas)
        self._preencher()
        self.status.set("")

    def _cor(self, simbolo):
        return self.cores.get(simbolo) or self.padrao.get(simbolo)

    def _preencher(self):

        alvo = self.busca.get().strip().upper()
        self.tabela.delete(*self.tabela.get_children())
        for simbolo in sorted(self.padrao, key=lambda s: s.upper()):
            nomes = ", ".join(sorted(self.nomes.get(simbolo, []))).title()
            if alvo and alvo not in simbolo.upper() and alvo not in nomes.upper():
                continue
            rgb = self._cor(simbolo)
            if simbolo in self.cores:
                origem = "trocada (salva)" if self.salvas.get(simbolo) == self.cores[simbolo] else "trocada (não salva)"
            else:
                origem = "padrão" if rgb else "sem cor"
            tag = "c_" + (cor_hex(rgb)[1:] if rgb else "nenhuma")
            if rgb:
                self.tabela.tag_configure(tag, background=cor_hex(rgb), foreground=texto_contraste(rgb))
            self.tabela.insert("", "end", iid=simbolo, values=(simbolo, nomes, cor_hex(rgb) or "-", origem), tags=(tag,))
        if self.selecionado in self.tabela.get_children():
            self.tabela.selection_set(self.selecionado)
            self.tabela.see(self.selecionado)

    def _selecionou(self):

        sel = self.tabela.selection()
        if not sel:
            return
        self.selecionado = sel[0]
        self._mostrar()

    def _mostrar(self):

        simbolo = self.selecionado
        rgb = self._cor(simbolo)
        self.rotulo_curso.configure(text=simbolo)
        self.amostra.delete("all")
        if rgb:
            self.amostra.configure(background=cor_hex(rgb))
            self.amostra.create_text(100, 55, text=simbolo, fill=texto_contraste(rgb), font=("Segoe UI", 12, "bold"))
        else:
            self.amostra.configure(background=COR_CAMPO)
            self.amostra.create_text(100, 55, text="sem cor\n(fica a do PSD)", fill=COR_SECUNDARIO, justify="center")
        padrao = self.padrao.get(simbolo)
        self.rotulo_padrao.configure(
            text="Padrão: " + (cor_hex(padrao) + "  (" + ", ".join(map(str, padrao)) + ")" if padrao else "nenhuma"))
        self.texto_cor.set(cor_hex(rgb))

    def _trocar(self, rgb):

        if not self.selecionado:
            self.status.set("Escolha um curso na lista.")
            return
        if rgb is None or rgb == self.padrao.get(self.selecionado):
            self.cores.pop(self.selecionado, None)
        else:
            self.cores[self.selecionado] = tuple(rgb)
        self._preencher()
        self._mostrar()
        self.status.set("Cor alterada. Clique em \"Salvar cores\" para usar nas próximas capas.")

    def escolher(self):

        from tkinter import colorchooser

        if not self.selecionado:
            self.status.set("Escolha um curso na lista.")
            return
        atual = self._cor(self.selecionado)
        escolhido = colorchooser.askcolor(
            color=cor_hex(atual) or None, title="Cor do curso " + self.selecionado, parent=self.root)
        if escolhido and escolhido[0]:
            self._trocar(tuple(int(round(v)) for v in escolhido[0]))

    def aplicar_texto(self):

        rgb = cor_do_texto(self.texto_cor.get())
        if rgb is None:
            self.status.set("Cor inválida. Use #7A1A18 ou 122, 26, 24.")
            return
        self._trocar(rgb)

    def voltar_padrao(self):

        self._trocar(None)

    def salvar(self):

        from tkinter import messagebox

        if self.janela.processo is not None:
            messagebox.showwarning("Cores dos cursos", "Espere terminar o que está rodando antes de salvar.")
            return
        modulo_recorte().gravar_config({"cores_cursos": {s: cor_hex(rgb) for s, rgb in sorted(self.cores.items())}})
        self.salvas = dict(self.cores)
        self._preencher()
        self.status.set(str(len(self.cores)) + " cor(es) trocada(s) salvas. Valem a partir da próxima capa gerada.")


def main():

    import tkinter as tk

    root = tk.Tk()
    JanelaCapas(root)
    root.mainloop()


if __name__ == "__main__":
    main()
