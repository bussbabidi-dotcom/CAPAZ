# -*- coding: utf-8 -*-
"""
Janela da automação de capas.

Cole os alunos direto do Excel (NOME | CURSO | PLANO), confira na tabela
e clique em Validar ou Gerar capas. A janela grava uma planilha temporária
e roda AUTOMAÇAO-PASTA.py em um processo separado.
"""

import json
import multiprocessing
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

        cfg_inicial = {}
        try:
            if CONFIG.is_file():
                cfg_inicial = json.loads(CONFIG.read_text(encoding="utf-8"))
        except Exception:
            cfg_inicial = {}

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
        self.limite_cpu_pct = tk.DoubleVar(
            value=float(max(25, min(85, int(cfg_inicial.get("hardware_limite_cpu_pct", 60)))))
        )
        self.usar_gpu_opencl = tk.BooleanVar(
            value=bool(cfg_inicial.get("hardware_usar_gpu_opencl", True))
        )
        self.usar_copia_leve_local = tk.BooleanVar(
            value=bool(cfg_inicial.get("usar_copia_leve_local", True))
        )
        self.preset_hardware = tk.StringVar(value="60% (Equilibrado)")
        self.resumo_hardware = tk.StringVar(value="")
        self.parando = False
        self.status = tk.StringVar(
            value="Sem planilha: as capas saem das fotos da pasta. Cole alunos ou escolha uma planilha se quiser."
        )

        self.fila = []

        self._montar()
        self._atualizar_resumo_hardware(salvar=False)
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
        abas.add(principal, text="Capas (Lista e Gerar)")

        quadro_fotos = ttk.Frame(abas, padding=12)
        self.quadro_fotos = quadro_fotos

        quadro_revisao = ttk.Frame(abas, padding=14)
        abas.add(quadro_revisao, text="Revisão e Recorte")
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

        # ---------------- Barra Rápida (1 Clique Sem Planilha) ----------------

        topo_rapido = ttk.Frame(principal)
        topo_rapido.grid(row=0, column=0, sticky="we")

        linha_rapida_1 = ttk.Frame(topo_rapido)
        linha_rapida_1.pack(fill="x")

        self.botao_ler_fotos = ttk.Button(
            linha_rapida_1, text="Carregar Fotos da Pasta (1 Clique)", style="Primario.TButton",
            command=self.ler_fotos,
        )
        self.botao_ler_fotos.pack(side="left", padx=(0, 6))

        self.botao_escolher_fotos = ttk.Button(
            linha_rapida_1, text="Escolher Melhores Fotos (Subpastas)", style="Primario.TButton",
            command=self._escolher_melhores_fotos_subpastas,
        )
        self.botao_escolher_fotos.pack(side="left", padx=(0, 6))

        self.botao_tela_7_melhores = ttk.Button(
            linha_rapida_1,
            text="Tela: Escolher 1 entre as 7 Melhores",
            command=self.abrir_tela_sete_melhores,
        )
        self.botao_tela_7_melhores.pack(side="left", padx=(0, 10))

        ttk.Label(linha_rapida_1, text="Curso p/ todos:").pack(side="left", padx=(4, 4))
        self.campo_curso_todos = ttk.Combobox(
            linha_rapida_1, textvariable=self.curso_todos, width=24, postcommand=self._listar_cursos_todos
        )
        self.campo_curso_todos.pack(side="left")
        self.campo_curso_todos.bind("<Return>", lambda e: self._aplicar_curso_todos())
        self.campo_curso_todos.bind("<<ComboboxSelected>>", lambda e: self._aplicar_curso_todos())
        ttk.Button(
            linha_rapida_1, text="Aplicar", command=self._aplicar_curso_todos
        ).pack(side="left", padx=(4, 12))

        ttk.Label(linha_rapida_1, text="Plano p/ todos:").pack(side="left", padx=(4, 4))
        self.campo_plano_todos = ttk.Combobox(
            linha_rapida_1, textvariable=self.plano_todos, width=22, postcommand=self._listar_planos)
        self.campo_plano_todos.pack(side="left")
        self.campo_plano_todos.bind("<Return>", lambda e: self._aplicar_plano_todos())
        self.campo_plano_todos.bind("<<ComboboxSelected>>", lambda e: self._aplicar_plano_todos())
        ttk.Button(
            linha_rapida_1, text="Aplicar", command=self._aplicar_plano_todos
        ).pack(side="left", padx=(4, 12))

        ttk.Button(
            linha_rapida_1, text="Ir para Revisão e Recorte ->",
            command=lambda: self.abas.select(self.quadro_revisao),
        ).pack(side="left", padx=(4, 0))

        self.mostrar_avancado = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            linha_rapida_1, text="Usar planilha / Opções avançadas",
            variable=self.mostrar_avancado, command=self._alternar_modo_avancado,
        ).pack(side="right")

        # ---------------- Barra de Cópia Leve Local + Limite de Hardware (Sem Travar o PC) ----------------

        barra_hardware = ttk.Frame(topo_rapido)
        barra_hardware.pack(fill="x", pady=(6, 0))

        ttk.Checkbutton(
            barra_hardware,
            text="Cópia leve local p/ analisar (Original só na escolhida)",
            variable=self.usar_copia_leve_local,
            command=lambda: self._atualizar_resumo_hardware(salvar=True),
        ).pack(side="left", padx=(0, 12))

        ttk.Label(barra_hardware, text="Limite de Hardware (CPU):").pack(side="left", padx=(0, 4))
        self.escala_hardware = ttk.Scale(
            barra_hardware,
            from_=25,
            to=85,
            variable=self.limite_cpu_pct,
            orient="horizontal",
            length=135,
            command=lambda _v: self._atualizar_resumo_hardware(salvar=False),
        )
        self.escala_hardware.pack(side="left", padx=(0, 6))
        self.escala_hardware.bind("<ButtonRelease-1>", lambda e: self._atualizar_resumo_hardware(salvar=True))

        self.combo_preset_hardware = ttk.Combobox(
            barra_hardware,
            textvariable=self.preset_hardware,
            state="readonly",
            width=19,
            values=(
                "30% (Leve)",
                "50% (Metade livre)",
                "60% (Equilibrado)",
                "75% (Rápido)",
                "85% (Turbo controlado)",
            ),
        )
        self.combo_preset_hardware.pack(side="left", padx=(0, 10))
        self.combo_preset_hardware.bind("<<ComboboxSelected>>", self._escolheu_preset_hardware)

        ttk.Checkbutton(
            barra_hardware,
            text="Impulsionar c/ GPU (OpenCL)",
            variable=self.usar_gpu_opencl,
            command=lambda: self._atualizar_resumo_hardware(salvar=True),
        ).pack(side="left", padx=(0, 10))

        self.abrir_tela_7_ao_concluir = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            barra_hardware,
            text="Abrir tela das 7 melhores ao concluir",
            variable=self.abrir_tela_7_ao_concluir,
        ).pack(side="left", padx=(0, 10))

        ttk.Label(
            barra_hardware,
            textvariable=self.resumo_hardware,
            style="Dica.TLabel",
        ).pack(side="left")

        # ---------------- Painel Avançado (Oculto por padrão para não poluir) ----------------

        self.quadro_avancado = ttk.Frame(principal)

        origem = ttk.Frame(self.quadro_avancado)
        origem.pack(fill="x", pady=(6, 4))

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

        self.campo_planilha = ttk.Entry(origem, textvariable=self.planilha, width=50)
        self.campo_planilha.pack(side="left", fill="x", expand=True, padx=4)

        self.botao_planilha = ttk.Button(
            origem, text="Escolher...", command=self._escolher_planilha
        )
        self.botao_planilha.pack(side="left")

        # ---------------- Tabela limpa de Alunos ----------------

        colagem = ttk.LabelFrame(
            principal,
            text="Lista de Capas (duplo-clique em um aluno para abrir na Revisão e Recorte)",
            padding=6,
        )
        colagem.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        colagem.columnconfigure(0, weight=0)
        colagem.columnconfigure(1, weight=1)
        colagem.rowconfigure(0, weight=1)

        self.texto_colagem = tk.Text(colagem, height=10, width=38, wrap="none", undo=True)
        self.texto_colagem.bind("<<Paste>>", lambda e: self.root.after(50, self._colou_texto))

        colunas = ("linha", "nome", "curso", "plano")
        self.tabela = ttk.Treeview(colagem, columns=colunas, show="headings", height=12)
        for col, titulo, largura in (
            ("linha", "#", 46),
            ("nome", "Nome do Formando", 280),
            ("curso", "Curso (definido na prévia ou lista)", 220),
            ("plano", "Plano / Tamanho da Capa", 200),
        ):
            self.tabela.heading(col, text=titulo)
            self.tabela.column(col, width=largura, anchor="w", stretch=(col != "linha"))
        self.tabela.tag_configure("aviso", foreground=COR_AMARELO)
        self.tabela.grid(row=0, column=1, sticky="nsew")
        self.tabela.bind("<Double-1>", self._duplo_clique_tabela)

        barra = ttk.Scrollbar(colagem, orient="vertical", command=self.tabela.yview)
        barra.grid(row=0, column=2, sticky="ns")
        self.tabela.configure(yscrollcommand=barra.set)

        botoes_colagem = ttk.Frame(colagem)
        botoes_colagem.grid(row=1, column=0, columnspan=3, sticky="we", pady=(6, 0))

        self.botao_colar = ttk.Button(
            botoes_colagem, text="Colar do Excel",
            command=self._colar_clipboard,
        )
        self.botao_colar.pack(side="left")

        self.botao_carregar = ttk.Button(
            botoes_colagem, text="Atualizar tabela", command=self.carregar
        )
        self.botao_carregar.pack(side="left", padx=6)

        self.botao_limpar = ttk.Button(
            botoes_colagem, text="Limpar lista", command=self._limpar
        )
        self.botao_limpar.pack(side="left")

        ttk.Label(botoes_colagem, text="Filtrar aluno:").pack(side="left", padx=(18, 4))
        ttk.Entry(botoes_colagem, textvariable=self.aluno, width=22).pack(side="left")

        self.total_alunos = ttk.Label(botoes_colagem, text="")
        self.total_alunos.pack(side="right")

        # ---------------- Filtros avançados (ocultos por padrão) ----------------

        filtros = ttk.LabelFrame(self.quadro_avancado, text="Filtros avançados", padding=6)
        filtros.pack(fill="x", pady=(4, 0))

        ttk.Label(filtros, text="Plano").grid(row=0, column=0, sticky="w")
        ttk.Entry(filtros, textvariable=self.plano, width=20).grid(row=0, column=1, sticky="w", padx=(4, 12))

        ttk.Label(filtros, text="Linhas").grid(row=0, column=2, sticky="w")
        ttk.Entry(filtros, textvariable=self.linhas, width=14).grid(row=0, column=3, sticky="w", padx=(4, 4))

        ttk.Label(filtros, text="Berlim").grid(row=0, column=4, sticky="e", padx=(12, 4))
        ttk.Combobox(
            filtros, textvariable=self.berlim, state="readonly", width=12,
            values=["Ambos", "Só 24x34", "Só 20x20"],
        ).grid(row=0, column=5, sticky="w")

        ttk.Label(filtros, text="Lote").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(filtros, textvariable=self.lote, width=20).grid(row=1, column=1, sticky="w", padx=(4, 12), pady=(6, 0))

        ttk.Label(filtros, text="Status").grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(filtros, textvariable=self.status_filtro, width=20).grid(
            row=1, column=3, sticky="w", padx=(4, 12), pady=(6, 0))

        # ---------------- Ações Principais (Limpas e Diretas) ----------------

        acoes = ttk.Frame(principal)
        acoes.grid(row=3, column=0, sticky="we", pady=(10, 0))

        self.botao_gerar = ttk.Button(
            acoes, text="GERAR CAPAS NO PHOTOSHOP", style="Primario.TButton",
            command=lambda: self.executar(False),
        )
        self.botao_gerar.pack(side="left", padx=(0, 8))

        self.botao_recorte = ttk.Button(
            acoes, text="Recortar todas as fotos", command=lambda: self.recortar(False)
        )
        self.botao_recorte.pack(side="left", padx=(0, 8))

        self.botao_validar = ttk.Button(
            acoes, text="Validar (sem Photoshop)", command=lambda: self.executar(True)
        )
        self.botao_validar.pack(side="left", padx=(0, 8))

        self.botao_parar = ttk.Button(acoes, text="Parar", command=self.parar, state="disabled")
        self.botao_parar.pack(side="left", padx=(0, 12))

        ttk.Label(acoes, text="Salvar:").pack(side="left", padx=(4, 4))
        ttk.Combobox(
            acoes, textvariable=self.tipo_saida, state="readonly", width=10, values=TIPOS_SAIDA,
        ).pack(side="left", padx=(0, 12))

        ttk.Checkbutton(
            acoes, text="Refazer existentes", variable=self.refazer
        ).pack(side="left", padx=(0, 10))

        ttk.Checkbutton(
            acoes, text="Só erros", variable=self.so_erros
        ).pack(side="left", padx=(0, 10))

        self.botao_relatorio = ttk.Button(
            acoes, text="Relatório", command=self._abrir_relatorio, state="disabled"
        )
        self.botao_relatorio.pack(side="left")

        self.botao_recorte_simular = ttk.Button(
            self.quadro_avancado, text="Recortar fotos (simular)", command=lambda: self.recortar(True)
        )
        self.botao_miniaturas = ttk.Button(
            self.quadro_avancado, text="Miniaturas", command=self._abrir_miniaturas
        )

        ttk.Label(acoes, textvariable=self.status).pack(side="left", padx=12)

        # ---------------- Saída ----------------

        saida = ttk.LabelFrame(principal, text="Andamento / Saída", padding=6)
        saida.grid(row=4, column=0, sticky="nsew", pady=(8, 0))
        saida.columnconfigure(0, weight=1)
        saida.rowconfigure(0, weight=1)

        self.texto_saida = tk.Text(saida, height=9, wrap="word", state="disabled",
                                   background=COR_FUNDO, foreground=COR_TEXTO)
        self.texto_saida.grid(row=0, column=0, sticky="nsew")

        barra_saida = ttk.Scrollbar(saida, orient="vertical", command=self.texto_saida.yview)
        barra_saida.grid(row=0, column=1, sticky="ns")
        self.texto_saida.configure(yscrollcommand=barra_saida.set)

        self._atualizar_origem()

    def _escolheu_preset_hardware(self, evento=None):
        txt = self.preset_hardware.get()
        for num in (30, 50, 60, 75, 85):
            if str(num) in txt:
                self.limite_cpu_pct.set(float(num))
                break
        self._atualizar_resumo_hardware(salvar=True)

    def _atualizar_resumo_hardware(self, salvar=True):
        pct = max(25, min(85, int(round(float(self.limite_cpu_pct.get() or 60)))))
        usar_gpu = bool(self.usar_gpu_opencl.get())
        usar_proxy = bool(self.usar_copia_leve_local.get())

        if pct <= 35:
             self.preset_hardware.set("30% (Leve)")
        elif pct <= 54:
            self.preset_hardware.set("50% (Metade livre)")
        elif pct <= 67:
            self.preset_hardware.set("60% (Equilibrado)")
        elif pct <= 80:
            self.preset_hardware.set("75% (Rápido)")
        else:
            self.preset_hardware.set("85% (Turbo controlado)")

        nucleos_totais = max(1, os.cpu_count() or 4)
        max_permitido = max(1, nucleos_totais - 1) if nucleos_totais > 1 else 1
        nucleos_usados = min(max_permitido, max(1, int(round(nucleos_totais * (pct / 100.0)))))
        nucleos_livres = max(0, nucleos_totais - nucleos_usados)

        gpu_txt = " + GPU OpenCL" if usar_gpu else " (só CPU)"
        proxy_txt = "Proxy leve ativo" if usar_proxy else "Leitura direta"
        self.resumo_hardware.set(
            f"Usando {nucleos_usados}/{nucleos_totais} núcleos ({pct}% · {nucleos_livres} livres p/ não travar){gpu_txt} · {proxy_txt}"
        )

        if salvar:
            try:
                rf = modulo_recorte()
                rf.gravar_config({
                    "hardware_limite_cpu_pct": pct,
                    "hardware_usar_gpu_opencl": usar_gpu,
                    "usar_copia_leve_local": usar_proxy,
                })
            except Exception:
                pass

    def _alternar_modo_avancado(self):
        if self.mostrar_avancado.get():
            self.quadro_avancado.grid(row=2, column=0, sticky="we", pady=(6, 0))
            self.texto_colagem.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        else:
            self.quadro_avancado.grid_forget()
            self.texto_colagem.grid_forget()

    def _duplo_clique_tabela(self, evento=None):
        sel = self.tabela.selection()
        if not sel:
            return
        valores = self.tabela.item(sel[0], "values")
        if len(valores) >= 2:
            nome_aluno = valores[1]
            self.abas.select(self.quadro_revisao)
            if hasattr(self, "aba_fotos"):
                self.aba_fotos.selecionar_aluno_por_nome(nome_aluno)

    def _aba_principal_mudou(self, evento=None):

        if self.abas.select() == str(self.quadro_revisao):
            self.aba_fotos.abrir_revisao()

    # --------------------------------------------------------
    # Colagem e Sincronização Automática com a Revisão
    # --------------------------------------------------------

    def _carregar_ultima_colagem(self):
        """Inicia direto verificando o cache persistente (cache_analise_subpastas.json) e lendo as fotos da pasta."""
        self.origem.set("fotos")
        self._atualizar_origem()
        try:
            rf = modulo_recorte()
            if hasattr(rf, "verificar_cache_analise_subpastas"):
                info_cache = rf.verificar_cache_analise_subpastas(
                    limite_cpu_pct=int(round(self.limite_cpu_pct.get())),
                    usar_gpu_opencl=bool(self.usar_gpu_opencl.get()),
                )
                if info_cache.get("existe"):
                    self.status.set(
                        f"Cache carregado (cache_analise_subpastas.json: {info_cache.get('fotos_em_cache', 0)} foto(s), "
                        f"{info_cache.get('subpastas_em_cache', 0)} subpasta(s)). Lendo fotos..."
                    )
        except Exception:
            pass
        self.root.after(250, self.ler_fotos)

    def sincronizar_curso_aluno(self, nome_aluno, curso):
        """Grava automaticamente na lista principal o curso definido na prévia/revisão."""
        if not nome_aluno:
            return
        alvo = chave(nome_aluno)
        alunos = interpretar_colagem(self.texto_colagem.get("1.0", "end"))
        alterou = False
        for a in alunos:
            if chave(a.get("nome", "")) == alvo:
                if a.get("curso") != curso:
                    a["curso"] = curso
                    alterou = True
        if alterou:
            self._preencher_texto(alunos)
            try:
                PASTA_TEMP.mkdir(parents=True, exist_ok=True)
                ULTIMA_COLAGEM.write_text(self.texto_colagem.get("1.0", "end"), encoding="utf-8")
            except Exception:
                pass

    def sincronizar_cursos_fotos(self, mapa_cursos_manuais):
        """Sincroniza todos os cursos salvos por foto para a tabela principal de Capas."""
        if not mapa_cursos_manuais:
            return
        rf = modulo_recorte()
        por_nome = {}
        for cam, cur in mapa_cursos_manuais.items():
            if cur:
                nm = rf.nome_do_aluno(Path(cam))
                if nm:
                    por_nome[chave(nm)] = cur
        if not por_nome:
            return
        alunos = interpretar_colagem(self.texto_colagem.get("1.0", "end"))
        alterou = False
        for a in alunos:
            ch = chave(a.get("nome", ""))
            if ch in por_nome and a.get("curso") != por_nome[ch]:
                a["curso"] = por_nome[ch]
                alterou = True
        if alterou:
            self._preencher_texto(alunos)

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
        """Monta a lista de capas pelas fotos da pasta verificando antes o cache em disco (cache_analise_subpastas.json)."""

        if getattr(self, "_lendo_fotos", False):
            return
        self._lendo_fotos = True
        self.botao_ler_fotos.configure(state="disabled")
        self.status.set("Verificando cache_analise_subpastas.json e lendo as fotos da pasta...")
        limite_pct = int(round(self.limite_cpu_pct.get() or 60))
        usar_gpu = bool(self.usar_gpu_opencl.get())

        def trabalho():
            erro = None
            alunos = []
            info_cache = {}
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
                if hasattr(rf, "verificar_cache_analise_subpastas"):
                    info_cache = rf.verificar_cache_analise_subpastas(
                        cfg["fotos"],
                        processar_pendentes_pool=True,
                        limite_cpu_pct=limite_pct,
                        usar_gpu_opencl=usar_gpu,
                    )
                alunos = rf.alunos_das_fotos(cfg["fotos"])
            except Exception as e:
                erro = str(e)
                alunos = None
            self.root.after(0, lambda: self._fotos_lidas(alunos, erro, info_cache))

        threading.Thread(target=trabalho, daemon=True).start()

    def _fotos_lidas(self, alunos, erro, info_cache=None):

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

        com_curso = sum(1 for a in self.alunos if a.get("curso"))
        texto = (
            str(len(alunos)) + " capa(s) carregada(s) direto das fotos ("
            + str(com_curso) + " já com curso definido)"
        )
        if info_cache and (info_cache.get("fotos_em_cache", 0) > 0 or info_cache.get("subpastas_em_cache", 0) > 0):
            texto += (
                f" · Cache ativo ({info_cache.get('fotos_em_cache', 0)} foto(s) / "
                f"{info_cache.get('subpastas_em_cache', 0)} subpasta(s) em cache_analise_subpastas.json)."
            )
        else:
            texto += "."
        if erro:
            texto += "  " + erro
        self.status.set(texto)

    def _escolher_melhores_fotos_subpastas(self):
        """Abre uma pasta com as subpastas de cada formando, escolhe a melhor foto conforme o plano e copia p/ FOTO."""

        from tkinter import filedialog
        import time

        if getattr(self, "_escolhendo_fotos", False):
            return

        rf = modulo_recorte()
        bruto = rf.ler_config_bruto()
        pasta_inicial = bruto.get("subpastas_fotos_dir") or bruto.get("photos_dir") or str(PASTA)
        pasta = filedialog.askdirectory(
            title="Escolha a pasta principal com as subpastas dos formandos",
            initialdir=pasta_inicial if Path(pasta_inicial).is_dir() else str(PASTA),
        )
        if not pasta:
            return

        limite_pct = max(25, min(85, int(round(self.limite_cpu_pct.get() or 60))))
        usar_gpu = bool(self.usar_gpu_opencl.get())
        usar_proxy = bool(self.usar_copia_leve_local.get())

        try:
            rf.gravar_config({
                "subpastas_fotos_dir": pasta,
                "hardware_limite_cpu_pct": limite_pct,
                "hardware_usar_gpu_opencl": usar_gpu,
                "usar_copia_leve_local": usar_proxy,
            })
        except Exception:
            pass

        self._escolhendo_fotos = True
        self.botao_escolher_fotos.configure(state="disabled")
        self.botao_ler_fotos.configure(state="disabled")
        self._limpar_saida()

        info_cache = (
            rf.verificar_cache_analise_subpastas(limite_cpu_pct=limite_pct, usar_gpu_opencl=usar_gpu)
            if hasattr(rf, "verificar_cache_analise_subpastas")
            else {}
        )
        nucleos_totais = info_cache.get("nucleos_cpu") or max(1, os.cpu_count() or 4)
        nucleos_usados = info_cache.get("nucleos_usados") or max(1, int(round(nucleos_totais * (limite_pct / 100.0))))
        nucleos_livres = max(0, nucleos_totais - nucleos_usados)
        modo_proxy_txt = (
            "Cópia leve local (.cache_proxies_capaz ~45KB) p/ analisar + Original na escolhida"
            if usar_proxy
            else "Leitura direta com DCT 1/8"
        )
        self._escrever(
            f"Iniciando seleção acelerada pelo seu hardware ({nucleos_usados}/{nucleos_totais} núcleos = {limite_pct}% · {nucleos_livres} livres p/ não travar"
            + (" + GPU OpenCL" if usar_gpu else "")
            + " · Prioridade Baixa) em: "
            + pasta + "\n"
        )
        self._escrever(f"Modo de leitura da pasta: {modo_proxy_txt}\n")
        if info_cache.get("existe"):
            self._escrever(
                f"Cache em disco verificado: {info_cache.get('fotos_em_cache', 0)} foto(s), "
                f"{info_cache.get('subpastas_em_cache', 0)} subpasta(s) e "
                f"{info_cache.get('proxies_locais', 0)} cópia(s) leve(s) pronta(s) para carregamento instantâneo!\n"
            )
        self._escrever(
            "Modo Rápido: 1º Separação instantânea por ORIENTAÇÃO via cabeçalho (< 0.08ms/foto, sem decodificar) · "
            "2º Pessoa sozinha + Encaixe no molde · Parada imediata na 1ª foto boa sem demorar nos critérios!\n\n"
        )

        alunos_atuais = list(self.carregar())
        plano_padrao = self.plano_todos.get().strip() or "Prime Book 30X20"
        curso_padrao = self.curso_todos.get().strip()
        berlim_modo = self.berlim.get().strip() or "Ambos"

        def progresso(atual, total, msg):
            self.root.after(0, lambda: self.status.set(f"[{atual}/{total}] {msg}"))

        def trabalho():
            erro = None
            relatorio = []
            t0 = time.perf_counter()
            try:
                cfg = rf.carregar_config()
                calib = rf.carregar_calibracao(cfg)
                relatorio = rf.escolher_melhores_fotos_subpastas(
                    pasta,
                    cfg,
                    calib=calib,
                    alunos_planos=alunos_atuais,
                    plano_padrao=plano_padrao,
                    curso_padrao=curso_padrao,
                    berlim_modo=berlim_modo,
                    progresso_cb=progresso,
                    limite_cpu_pct=limite_pct,
                    usar_gpu_opencl=usar_gpu,
                    usar_copia_leve_local=usar_proxy,
                )
            except Exception as e:
                erro = str(e)
            tempo_seg = time.perf_counter() - t0
            self.root.after(0, lambda: self._melhores_fotos_prontas(relatorio, erro, tempo_seg))

        threading.Thread(target=trabalho, daemon=True).start()

    def _melhores_fotos_prontas(self, relatorio, erro, tempo_seg=0.0):

        self._escolhendo_fotos = False
        self.botao_escolher_fotos.configure(state="normal")
        self.botao_ler_fotos.configure(state="normal")

        if erro:
            self.status.set("Erro ao escolher fotos: " + erro)
            self._escrever("\nERRO: " + erro + "\n")
            return

        ok_count = 0
        aviso_qualidade_count = 0
        subpastas_cache_count = 0
        fotos_cache_total = 0
        fotos_novas_total = 0
        proxies_total = 0
        for item in relatorio:
            escolhidas = item.get("escolhidas") or {}
            avisos_item = item.get("avisos") or []
            do_cache = bool(item.get("do_cache", False))
            fotos_cache_total += int(item.get("fotos_em_cache", 0))
            fotos_novas_total += int(item.get("fotos_analisadas_agora", 0))
            proxies_total += int(item.get("proxies_usados", 0))
            if do_cache:
                subpastas_cache_count += 1
            tag_cache = " [CACHE INSTANTÂNEO]" if do_cache else ""
            if escolhidas:
                ok_count += 1
                resumos = []
                for tam, info in escolhidas.items():
                    av_rev = info.get("aviso_revisao") or ""
                    sufixo_aviso = f" [AVISO: {av_rev}]" if av_rev else ""
                    resumos.append(
                        f"{tam}: {info['nome_arquivo']} ({info['nota']} pts — {info['resumo']}){sufixo_aviso}"
                    )
                if avisos_item:
                    aviso_qualidade_count += 1
                    self._escrever(
                        f"AVISO QUALIDADE (menos pior){tag_cache} — {item['nome']} [{item['plano']}]: " + " | ".join(resumos) + "\n"
                    )
                else:
                    self._escrever(f"OK{tag_cache} — {item['nome']} [{item['plano']}]: " + " | ".join(resumos) + "\n")
            else:
                self._escrever(f"AVISO — {item['nome']}: nenhuma foto encontrada na subpasta.\n")

        self._escrever(
            f"\nResumo: {subpastas_cache_count}/{len(relatorio)} subpasta(s) instantâneas do cache "
            f"| {fotos_cache_total} foto(s) em cache | {fotos_novas_total} analisada(s) agora ({proxies_total} via cópia leve local) "
            f"| Original copiada apenas nas campeãs | Tempo total: {tempo_seg * 1000:.0f} ms\n"
        )

        # Atualiza a tabela com os alunos e planos processados
        novos_alunos = [
            {"nome": r["nome"], "curso": r.get("curso") or "", "plano": r.get("plano") or "Prime Book 30X20"}
            for r in relatorio if r.get("escolhidas")
        ]
        if novos_alunos:
            self.origem.set("fotos")
            self._atualizar_origem()
            self._preencher_texto(novos_alunos, self.curso_todos.get().strip())
        else:
            self.ler_fotos()

        if hasattr(self, "aba_fotos"):
            self.aba_fotos.atualizar_lista(silencioso=True)
            self.aba_fotos.agendar_revisao()

        msg_final = (
            f"Concluído em {tempo_seg * 1000:.0f} ms! {ok_count} formando(s) "
            f"({subpastas_cache_count} via cache_analise_subpastas.json)"
        )
        if aviso_qualidade_count:
            msg_final += f" · {aviso_qualidade_count} com aviso de luz/contraste na Revisão."
        else:
            msg_final += " · Clique em 'Tela: Escolher 1 entre as 7 Melhores' para revisar."
        self.status.set(msg_final)

        if ok_count > 0 and getattr(self, "abrir_tela_7_ao_concluir", None) and self.abrir_tela_7_ao_concluir.get():
            self.root.after(150, self.abrir_tela_sete_melhores)

    def abrir_tela_sete_melhores(self, nome_aluno_inicial=None):
        """
        Abre a Tela Visual Rápida para o usuário escolher com 1 clique (ou teclas 1 a 7)
        1 foto entre as 7 melhores avaliadas pela escolha rápida de orientação.
        """
        from tkinter import messagebox

        rf = modulo_recorte()
        opcoes_todas = rf.ler_opcoes_foto() if hasattr(rf, "ler_opcoes_foto") else {}
        if not opcoes_todas:
            messagebox.showinfo(
                "Escolher 1 entre as 7 Melhores",
                "Ainda não há fotos avaliadas das subpastas.\n\n"
                "Clique primeiro em \"Escolher Melhores Fotos (Subpastas)\" para fazer a separação rápida "
                "das 7 melhores fotos de cada formando.",
            )
            return

        # Fecha instância anterior se já estiver aberta para focar na nova
        if getattr(self, "_janela_7_melhores", None) is not None:
            try:
                if self._janela_7_melhores.winfo_exists():
                    self._janela_7_melhores.destroy()
            except Exception:
                pass

        tk = self.tk
        ttk = self.ttk
        cfg = rf.carregar_config()
        calib = rf.carregar_calibracao(cfg)

        jan = tk.Toplevel(self.root)
        self._janela_7_melhores = jan
        jan.title("Escolher 1 entre as 7 Melhores Fotos — Seleção Rápida por Orientação")
        jan.geometry("1420x820")
        jan.minsize(1120, 680)
        jan.configure(background=COR_FUNDO)
        jan.transient(self.root)

        # Organiza os registros ordenados por nome do formando e tamanho
        registros = []
        for chave_op, dados_op in opcoes_todas.items():
            if not isinstance(dados_op, dict):
                continue
            cands = (dados_op.get("candidatas") or [])[:7]
            if not cands:
                continue
            registros.append({
                "chave_op": chave_op,
                "nome": dados_op.get("nome") or chave_op.split("|")[0],
                "tamanho": dados_op.get("tamanho") or "30X20",
                "escolhida": str(dados_op.get("escolhida") or cands[0].get("caminho") or ""),
                "aviso_revisao": str(dados_op.get("aviso_revisao") or ""),
                "candidatas": cands,
            })
        registros.sort(key=lambda r: (rf.chave(r["nome"]), r["tamanho"]))

        if not registros:
            jan.destroy()
            return

        jan.columnconfigure(1, weight=1)
        jan.rowconfigure(0, weight=1)

        # ---------------- COLUNA ESQUERDA: Lista de Formandos ----------------
        painel_esq = ttk.Frame(jan, padding=12)
        painel_esq.grid(row=0, column=0, sticky="nsew")
        painel_esq.rowconfigure(2, weight=1)
        painel_esq.columnconfigure(0, weight=1)

        ttk.Label(
            painel_esq,
            text=f"Formandos ({len(registros)})",
            style="Titulo.TLabel",
        ).grid(row=0, column=0, sticky="w")

        quadro_busca = ttk.Frame(painel_esq)
        quadro_busca.grid(row=1, column=0, sticky="we", pady=(6, 6))
        ttk.Label(quadro_busca, text="Filtrar:").pack(side="left")
        filtro_var = tk.StringVar(value="")
        campo_busca = ttk.Entry(quadro_busca, textvariable=filtro_var, width=24)
        campo_busca.pack(side="left", fill="x", expand=True, padx=(6, 0))

        quadro_lista = ttk.Frame(painel_esq)
        quadro_lista.grid(row=2, column=0, sticky="nsew")
        quadro_lista.rowconfigure(0, weight=1)
        quadro_lista.columnconfigure(0, weight=1)

        arvore = ttk.Treeview(
            quadro_lista,
            columns=("nome", "tam", "atual"),
            show="headings",
            selectmode="browse",
            width=32,
        )
        arvore.heading("nome", text="Formando")
        arvore.heading("tam", text="Tam.")
        arvore.heading("atual", text="Escolhida")
        arvore.column("nome", width=165, anchor="w")
        arvore.column("tam", width=54, anchor="center")
        arvore.column("atual", width=95, anchor="w")
        arvore.tag_configure("trocada", foreground=COR_DESTAQUE)
        arvore.tag_configure("aviso", foreground=COR_AMARELO)
        arvore.grid(row=0, column=0, sticky="nsew")

        scroll_esq = ttk.Scrollbar(quadro_lista, orient="vertical", command=arvore.yview)
        scroll_esq.grid(row=0, column=1, sticky="ns")
        arvore.configure(yscrollcommand=scroll_esq.set)

        avancar_auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            painel_esq,
            text="Avançar p/ próximo ao clicar na foto",
            variable=avancar_auto_var,
        ).grid(row=3, column=0, sticky="w", pady=(8, 4))

        modo_miniatura_var = tk.StringVar(value="recorte")
        quadro_modo = ttk.Frame(painel_esq)
        quadro_modo.grid(row=4, column=0, sticky="we", pady=(2, 6))
        ttk.Label(quadro_modo, text="Ver:", style="Dica.TLabel").pack(side="left")
        ttk.Radiobutton(
            quadro_modo,
            text="Recorte da Capa",
            value="recorte",
            variable=modo_miniatura_var,
            command=lambda: _renderizar_formando_atual(),
        ).pack(side="left", padx=(6, 6))
        ttk.Radiobutton(
            quadro_modo,
            text="Foto Inteira",
            value="inteira",
            variable=modo_miniatura_var,
            command=lambda: _renderizar_formando_atual(),
        ).pack(side="left")

        nav_botoes = ttk.Frame(painel_esq)
        nav_botoes.grid(row=5, column=0, sticky="we", pady=(4, 0))
        nav_botoes.columnconfigure(0, weight=1)
        nav_botoes.columnconfigure(1, weight=1)

        # ---------------- COLUNA DIREITA: Grade Visual das 7 Melhores Fotos ----------------
        painel_dir = ttk.Frame(jan, padding=12)
        painel_dir.grid(row=0, column=1, sticky="nsew")
        painel_dir.columnconfigure(0, weight=1)
        painel_dir.rowconfigure(1, weight=1)

        cabecalho = tk.Frame(painel_dir, background=COR_SUPERFICIE, padx=14, pady=10, highlightthickness=1, highlightbackground=COR_BORDA)
        cabecalho.grid(row=0, column=0, sticky="we", pady=(0, 10))
        cabecalho.columnconfigure(0, weight=1)

        titulo_aluno_var = tk.StringVar(value="Selecione um formando")
        subtitulo_aluno_var = tk.StringVar(value="")
        status_escolha_var = tk.StringVar(
            value="Dica rápida: clique em qualquer uma das 7 fotos (ou aperte as teclas 1 a 7 do teclado) para escolhê-la na hora!"
        )

        tk.Label(
            cabecalho,
            textvariable=titulo_aluno_var,
            font=("Segoe UI", 13, "bold"),
            background=COR_SUPERFICIE,
            foreground=COR_TEXTO,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        tk.Label(
            cabecalho,
            textvariable=subtitulo_aluno_var,
            font=("Segoe UI", 9),
            background=COR_SUPERFICIE,
            foreground=COR_DESTAQUE,
            anchor="w",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        ttk.Button(
            cabecalho,
            text="Abrir na Revisão e Recorte Completa",
            command=lambda: _abrir_aluno_na_revisao(),
        ).grid(row=0, column=1, rowspan=2, sticky="e", padx=(12, 0))

        grade_cartoes = ttk.Frame(painel_dir)
        grade_cartoes.grid(row=1, column=0, sticky="nsew")
        for c_idx in range(4):
            grade_cartoes.columnconfigure(c_idx, weight=1, uniform="col_7")
        grade_cartoes.rowconfigure(0, weight=1)
        grade_cartoes.rowconfigure(1, weight=1)

        rodape = tk.Frame(painel_dir, background=COR_SUPERFICIE, padx=12, pady=8, highlightthickness=1, highlightbackground=COR_BORDA)
        rodape.grid(row=2, column=0, sticky="we", pady=(10, 0))
        tk.Label(
            rodape,
            textvariable=status_escolha_var,
            font=("Segoe UI", 9, "bold"),
            background=COR_SUPERFICIE,
            foreground=COR_VERDE,
            anchor="w",
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(rodape, text="Concluir e Fechar (Esc)", style="Primario.TButton", command=jan.destroy).pack(side="right")

        estado_tela = {
            "reg_por_iid": {},
            "iid_atual": None,
            "imagens_tk": [],
            "token_render": 0,
        }

        def _texto_escolhida_curto(reg):
            cands = reg["candidatas"]
            esc = reg["escolhida"]
            for idx_c, c in enumerate(cands, 1):
                if str(c.get("caminho")) == esc:
                    return f"#{idx_c} ({c.get('nota', 0):.0f}pts)"
            return "#1"

        def _popular_arvore():
            termo = rf.chave(filtro_var.get()) if filtro_var.get().strip() else ""
            arvore.delete(*arvore.get_children())
            estado_tela["reg_por_iid"].clear()
            primeiro_iid = None
            alvo_iid = None
            alvo_nome_ch = rf.chave(nome_aluno_inicial) if nome_aluno_inicial else ""

            for idx_r, reg in enumerate(registros):
                if termo and termo not in rf.chave(reg["nome"]):
                    continue
                iid = f"reg_{idx_r}"
                estado_tela["reg_por_iid"][iid] = reg
                txt_esc = _texto_escolhida_curto(reg)
                tag = "trocada" if not txt_esc.startswith("#1 ") else ("aviso" if reg.get("aviso_revisao") else "")
                arvore.insert(
                    "",
                    "end",
                    iid=iid,
                    values=(reg["nome"], reg["tamanho"], txt_esc),
                    tags=(tag,) if tag else (),
                )
                if primeiro_iid is None:
                    primeiro_iid = iid
                if alvo_nome_ch and rf.chave(reg["nome"]) == alvo_nome_ch and alvo_iid is None:
                    alvo_iid = iid

            selecionar = alvo_iid or estado_tela["iid_atual"] or primeiro_iid
            if selecionar and arvore.exists(selecionar):
                arvore.selection_set(selecionar)
                arvore.see(selecionar)
                _ao_selecionar_arvore()

        def _ir_para_vizinho(delta):
            itens = arvore.get_children()
            if not itens:
                return
            sel = arvore.selection()
            idx_at = itens.index(sel[0]) if sel and sel[0] in itens else 0
            novo_idx = max(0, min(len(itens) - 1, idx_at + delta))
            novo_iid = itens[novo_idx]
            arvore.selection_set(novo_iid)
            arvore.see(novo_iid)
            _ao_selecionar_arvore()

        ttk.Button(nav_botoes, text="< Anterior", command=lambda: _ir_para_vizinho(-1)).grid(
            row=0, column=0, sticky="we", padx=(0, 4)
        )
        ttk.Button(nav_botoes, text="Próximo >", command=lambda: _ir_para_vizinho(1)).grid(
            row=0, column=1, sticky="we", padx=(4, 0)
        )

        def _abrir_aluno_na_revisao():
            iid = estado_tela["iid_atual"]
            reg = estado_tela["reg_por_iid"].get(iid) if iid else None
            if not reg:
                return
            self.abas.select(self.quadro_revisao)
            if hasattr(self, "aba_fotos"):
                self.aba_fotos.selecionar_aluno_por_nome(reg["nome"])

        def _escolher_candidata_indice(indice_0):
            iid = estado_tela["iid_atual"]
            reg = estado_tela["reg_por_iid"].get(iid) if iid else None
            if not reg:
                return
            cands = reg["candidatas"]
            if indice_0 < 0 or indice_0 >= len(cands):
                return
            cand = cands[indice_0]
            caminho_orig = cand.get("caminho")
            if not caminho_orig:
                return

            try:
                destino = rf.trocar_foto_por_opcao(reg["nome"], reg["tamanho"], caminho_orig, cfg)
            except Exception as e:
                status_escolha_var.set("Erro ao trocar foto: " + str(e))
                return

            reg["escolhida"] = str(caminho_orig)
            reg["aviso_revisao"] = (
                cand.get("aviso_qualidade", "") if not cand.get("aprovada_qualidade", True) else ""
            )
            txt_esc = _texto_escolhida_curto(reg)
            tag = "trocada" if indice_0 > 0 else ("aviso" if reg.get("aviso_revisao") else "")
            if arvore.exists(iid):
                arvore.set(iid, "atual", txt_esc)
                arvore.item(iid, tags=(tag,) if tag else ())

            status_escolha_var.set(
                f"✓ Escolhida a {indice_0 + 1}ª opção ({cand.get('nome_arquivo')}) para {reg['nome']} ({reg['tamanho']}) — Cópia original salva em FOTO!"
            )

            # Sincroniza com a aba Revisão e Recorte em segundo plano
            if hasattr(self, "aba_fotos"):
                try:
                    if self.aba_fotos.arquivo and self.aba_fotos.arquivo.resolve() == Path(destino).resolve():
                        self.aba_fotos.abrir(destino)
                    self.aba_fotos.agendar_revisao(200)
                except Exception:
                    pass

            if avancar_auto_var.get():
                itens = arvore.get_children()
                if iid in itens and itens.index(iid) + 1 < len(itens):
                    jan.after(120, lambda: _ir_para_vizinho(1))
                    return
            _renderizar_formando_atual()

        def _renderizar_formando_atual():
            for w in grade_cartoes.winfo_children():
                w.destroy()
            estado_tela["imagens_tk"].clear()

            iid = estado_tela["iid_atual"]
            reg = estado_tela["reg_por_iid"].get(iid) if iid else None
            if not reg:
                return

            cands = reg["candidatas"][:7]
            escolhida_cam = str(reg.get("escolhida") or "")
            tam = reg["tamanho"]
            modo_vis = modo_miniatura_var.get()

            titulo_aluno_var.set(f"{reg['nome']}   —   Capa {tam}   ({len(cands)} melhores opções da pasta)")
            subtitulo_aluno_var.set(
                "Separadas por Orientação Ideal + Pessoa Sozinha + Encaixe no Molde  |  "
                "Clique no cartão da foto desejada ou pressione as teclas 1 a 7"
            )

            estado_tela["token_render"] += 1
            meu_token = estado_tela["token_render"]

            # Cria os 7 cartões imediatamente (4 na linha de cima, 3 na linha de baixo)
            labels_imagem = []
            for idx_c, cand in enumerate(cands):
                linha_grid = 0 if idx_c < 4 else 1
                col_grid = idx_c if idx_c < 4 else (idx_c - 4)
                eh_escolhida = str(cand.get("caminho") or "") == escolhida_cam
                orient_ideal = bool(cand.get("orientacao_ideal", True))

                cor_borda_card = COR_VERDE if eh_escolhida else (COR_DESTAQUE if idx_c == 0 else COR_BORDA)
                esp_borda = 2 if eh_escolhida else 1

                card = tk.Frame(
                    grade_cartoes,
                    background=COR_SUPERFICIE,
                    highlightthickness=esp_borda,
                    highlightbackground=cor_borda_card,
                    highlightcolor=cor_borda_card,
                    padx=8,
                    pady=6,
                )
                card.grid(row=linha_grid, column=col_grid, sticky="nsew", padx=5, pady=5)
                card.columnconfigure(0, weight=1)
                card.rowconfigure(1, weight=1)

                # Topo do cartão: Posição (#1 a #7) + Badge de Orientação
                topo_c = tk.Frame(card, background=COR_SUPERFICIE)
                topo_c.grid(row=0, column=0, sticky="we")

                txt_pos = (
                    f"✓ {idx_c + 1}ª ESCOLHIDA"
                    if eh_escolhida
                    else (f"1ª RECOMENDADA" if idx_c == 0 else f"{idx_c + 1}ª OPÇÃO")
                )
                cor_pos = COR_VERDE if eh_escolhida else (COR_DESTAQUE if idx_c == 0 else COR_TEXTO)
                tk.Label(
                    topo_c,
                    text=txt_pos,
                    font=("Segoe UI", 9, "bold"),
                    background=COR_SUPERFICIE,
                    foreground=cor_pos,
                ).pack(side="left")

                txt_orient = f"{cand.get('orientacao', 'Vertical')} ({'Ideal' if orient_ideal else 'Adaptada'})"
                cor_orient = COR_VERDE if orient_ideal else COR_AMARELO
                tk.Label(
                    topo_c,
                    text=txt_orient,
                    font=("Segoe UI", 8, "bold"),
                    background=COR_SUPERFICIE,
                    foreground=cor_orient,
                ).pack(side="right")

                # Área da Miniatura (clicável para escolher em 1 clique)
                lbl_img = tk.Label(
                    card,
                    text="Carregando prévia...",
                    background=COR_CAMPO,
                    foreground=COR_SECUNDARIO,
                    cursor="hand2",
                )
                lbl_img.grid(row=1, column=0, sticky="nsew", pady=(5, 5))
                lbl_img.bind("<Button-1>", lambda e, i=idx_c: _escolher_candidata_indice(i))
                labels_imagem.append((idx_c, cand, lbl_img))

                # Nome do arquivo + Nota e Encaixe
                info_txt = (
                    f"{cand.get('nome_arquivo', '')}  ·  {cand.get('nota', 0):.1f} pts  ·  "
                    f"Encaixe {cand.get('encaixe_pct', 0)}%"
                )
                tk.Label(
                    card,
                    text=info_txt,
                    font=("Segoe UI", 8, "bold"),
                    background=COR_SUPERFICIE,
                    foreground=COR_TEXTO,
                    anchor="w",
                ).grid(row=2, column=0, sticky="we")

                resumo_curto = str(cand.get("resumo") or "")
                if len(resumo_curto) > 48:
                    resumo_curto = resumo_curto[:46] + "..."
                tk.Label(
                    card,
                    text=resumo_curto,
                    font=("Segoe UI", 8),
                    background=COR_SUPERFICIE,
                    foreground=COR_SECUNDARIO,
                    anchor="w",
                ).grid(row=3, column=0, sticky="we", pady=(0, 4))

                txt_btn = (
                    f"✓ Foto #{idx_c + 1} Escolhida (Tecla {idx_c + 1})"
                    if eh_escolhida
                    else f"Escolher Foto #{idx_c + 1} (Tecla {idx_c + 1})"
                )
                btn_card = ttk.Button(
                    card,
                    text=txt_btn,
                    style="Primario.TButton" if (eh_escolhida or idx_c == 0) else "TButton",
                    command=lambda i=idx_c: _escolher_candidata_indice(i),
                )
                btn_card.grid(row=4, column=0, sticky="we")

            # Carrega as 7 miniaturas em thread de fundo usando os proxies locais (~2ms por foto)
            def _carregar_miniaturas_bg():
                for idx_c, cand, lbl_w in labels_imagem:
                    if estado_tela["token_render"] != meu_token:
                        return
                    try:
                        im_pil = rf.obter_miniatura_candidata_rapida(
                            cand["caminho"],
                            tam,
                            cfg,
                            calib=calib,
                            caixa_sugerida=cand.get("caixa"),
                            lado_max=205,
                            modo=modo_vis,
                        )
                    except Exception:
                        im_pil = None

                    if im_pil is not None and estado_tela["token_render"] == meu_token:
                        def _aplicar_img(widget=lbl_w, pil_im=im_pil):
                            if estado_tela["token_render"] != meu_token or not widget.winfo_exists():
                                return
                            tk_im = ImageTk.PhotoImage(pil_im)
                            estado_tela["imagens_tk"].append(tk_im)
                            widget.configure(image=tk_im, text="")
                        jan.after(0, _aplicar_img)

            threading.Thread(target=_carregar_miniaturas_bg, daemon=True).start()

        def _ao_selecionar_arvore(evento=None):
            sel = arvore.selection()
            if not sel:
                return
            estado_tela["iid_atual"] = sel[0]
            _renderizar_formando_atual()

        arvore.bind("<<TreeviewSelect>>", _ao_selecionar_arvore)
        filtro_var.trace_add("write", lambda *a: _popular_arvore())

        # Atalhos de teclado 1 a 7 para escolher instantaneamente e setas para navegar
        for num_tecla in range(1, 8):
            jan.bind(
                str(num_tecla),
                lambda e, n=num_tecla - 1: (
                    None if isinstance(e.widget, tk.Entry) else _escolher_candidata_indice(n)
                ),
            )
            jan.bind(
                f"<KP_{num_tecla}>",
                lambda e, n=num_tecla - 1: (
                    None if isinstance(e.widget, tk.Entry) else _escolher_candidata_indice(n)
                ),
            )
        jan.bind("<Left>", lambda e: None if isinstance(e.widget, tk.Entry) else _ir_para_vizinho(-1))
        jan.bind("<Right>", lambda e: None if isinstance(e.widget, tk.Entry) else _ir_para_vizinho(1))
        jan.bind("<Escape>", lambda e: jan.destroy())

        _popular_arvore()

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

    def _listar_cursos_todos(self):
        try:
            automacao = modulo_automacao()
            cursos = sorted(automacao.CURSOS.keys())
        except Exception:
            cursos = []
        self.campo_curso_todos.configure(values=cursos)

    def _aplicar_curso_todos(self):

        if self.origem.get() == "arquivo":
            return
        alunos = self.carregar()
        curso = self.curso_todos.get().strip()
        if not curso:
            return
        for a in alunos:
            a["curso"] = curso
        if hasattr(self, "aba_fotos") and self.aba_fotos.fotos:
            for foto in self.aba_fotos.fotos:
                self.aba_fotos.cursos_manuais[chave_caminho_foto(foto)] = curso
            try:
                self.aba_fotos._gravar_cursos_foto()
            except Exception:
                pass
        self._preencher_texto(alunos)
        self.status.set("Curso " + curso + " aplicado e gravado em " + str(len(alunos)) + " aluno(s).")

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

        sem_curso = 0

        for aluno in self.alunos:
            incompleto = not aluno["nome"] or not aluno["plano"] or not aluno.get("curso")
            if not aluno.get("curso"):
                sem_curso += 1
            self.tabela.insert(
                "", "end",
                values=(aluno["linha"], aluno["nome"], aluno.get("curso") or "(definir na Revisão)", aluno["plano"]),
                tags=("aviso",) if incompleto else (),
            )

        texto_total = "Total: " + str(len(self.alunos)) + " capa(s)"
        if sem_curso:
            texto_total += "  |  " + str(sem_curso) + " aguardando curso (defina na aba Revisão e Recorte)"
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

            if hasattr(self, "aba_fotos") and self.aba_fotos.cursos_manuais:
                self.sincronizar_cursos_fotos(self.aba_fotos.cursos_manuais)
            alunos = self.carregar()

            if not any(a["nome"] for a in alunos):
                if self.origem.get() == "fotos":
                    messagebox.showwarning(
                        "Automação de Capas",
                        "Nenhuma foto encontrada na pasta.\n"
                        "Confira a pasta de fotos na aba \"Pastas e links\" e clique em \"Carregar Fotos da Pasta\".",
                    )
                else:
                    messagebox.showwarning("Automação de Capas", "Cole pelo menos um aluno antes de rodar.")
                return None

            PASTA_TEMP.mkdir(parents=True, exist_ok=True)
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

LADO_EDITOR = 2000
LADO_PREVIA_REVISAO = 2000
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
        self._aprovados_revisao = set()
        self._cache_thumbs_7_opcoes = {}
        self._cache_recorte_revisao = {}
        self._token_selecao_aluno = 0
        self._redimensionar_previa_pendente = None
        self._ultimo_tam_previa_revisao = (0, 0)
        self._botoes_grid_7 = []

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
        self.trava_aluno_revisao = tk.BooleanVar(value=True)
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

        # Widgets internos de compatibilidade mantidos no quadro oculto
        previa = ttk.Frame(q)
        self.combo_capa = ttk.Combobox(previa, textvariable=self.capa_previa, width=7, state="readonly")
        self.botao_gerar_previa = ttk.Button(previa, text="Gerar prévia da capa", command=self.gerar_previa_capa)
        self._img_previa = ImageTk.PhotoImage(Image.new("RGB", (380, 380), COR_FUNDO))
        self.rotulo_previa = tk.Label(previa, background=COR_FUNDO, image=self._img_previa)

        self._montar_revisao(self.quadro_revisao)

    def _montar_revisao(self, pagina):

        ttk = self.ttk
        tk = self.tk
        pagina.columnconfigure(0, weight=1)
        pagina.rowconfigure(0, weight=1)

        corpo = ttk.Frame(pagina)
        corpo.grid(row=0, column=0, sticky="nsew")
        corpo.columnconfigure(0, weight=1)
        corpo.columnconfigure(1, weight=3)
        corpo.columnconfigure(2, weight=0)
        corpo.rowconfigure(0, weight=2)
        corpo.rowconfigure(1, weight=3)

        # ---------------- Coluna Esquerda: Cursos e Alunos ----------------

        quadro_lista = ttk.LabelFrame(corpo, text="Cursos encontrados", padding=6)
        quadro_lista.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        quadro_lista.columnconfigure(0, weight=1)
        quadro_lista.rowconfigure(0, weight=1)

        self.lista_revisao = ttk.Treeview(
            quadro_lista, columns=("curso", "quantidade", "cor", "nomes"), show="tree headings",
            selectmode="browse", style="Review.Treeview", height=5,
        )
        self.lista_revisao.heading("#0", text="Brasão")
        self.lista_revisao.column("#0", width=62, anchor="center", stretch=False)
        for coluna, titulo, largura in (
            ("curso", "Curso", 150),
            ("quantidade", "Fotos", 46),
            ("cor", "Cor", 70),
            ("nomes", "Nomes", 80),
        ):
            self.lista_revisao.heading(coluna, text=titulo)
            self.lista_revisao.column(coluna, width=largura, anchor="w")
        self.lista_revisao.grid(row=0, column=0, sticky="nsew")
        barra_revisao = ttk.Scrollbar(quadro_lista, orient="vertical", command=self.lista_revisao.yview)
        barra_revisao.grid(row=0, column=1, sticky="ns")
        self.lista_revisao.configure(yscrollcommand=barra_revisao.set)
        self.lista_revisao.bind("<<TreeviewSelect>>", lambda e: self._selecionar_curso_revisao())

        quadro_alunos = ttk.LabelFrame(corpo, text="Alunos (clique para ver a capa e recortar)", padding=6)
        quadro_alunos.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(8, 0))
        quadro_alunos.columnconfigure(0, weight=1)
        quadro_alunos.rowconfigure(1, weight=1)
        ttk.Checkbutton(
            quadro_alunos, text="Fila Inteligente: Só Pendências e Avisos",
            variable=self.so_nomes_aviso, command=self._preencher_alunos_revisao,
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.lista_alunos = ttk.Treeview(
            quadro_alunos, columns=("nome", "capa", "situacao"), show="headings", selectmode="browse",
        )
        for coluna, titulo, largura in (("nome", "Aluno", 210), ("capa", "Capa", 54), ("situacao", "Status", 110)):
            self.lista_alunos.heading(coluna, text=titulo)
            self.lista_alunos.column(coluna, width=largura, anchor="w")
        self.lista_alunos.tag_configure("nome_reduzido", background="#36352b", foreground=COR_AMARELO)
        self.lista_alunos.tag_configure("nome_estoura", background="#3b2a30", foreground="#e3a3aa")
        self.lista_alunos.tag_configure("nome_ajustado", foreground=COR_AZUL)
        self.lista_alunos.tag_configure("aluno_aprovado", foreground=COR_VERDE)
        self.lista_alunos.grid(row=1, column=0, sticky="nsew")
        barra_alunos = ttk.Scrollbar(quadro_alunos, orient="vertical", command=self.lista_alunos.yview)
        barra_alunos.grid(row=1, column=1, sticky="ns")
        self.lista_alunos.configure(yscrollcommand=barra_alunos.set)
        self.lista_alunos.bind("<<TreeviewSelect>>", lambda e: self._selecionar_aluno_revisao())

        # ---------------- Coluna Central (GRANDE): Prévia da Capa + Definir Curso + Nome ----------------

        painel = ttk.Frame(corpo, padding=(2, 0, 8, 0))
        painel.grid(row=0, column=1, rowspan=2, sticky="nsew")
        painel.columnconfigure(0, weight=1)
        painel.rowconfigure(2, weight=1)

        topo_curso = ttk.LabelFrame(painel, text="Curso desta Capa (ao escolher, já grava na lista principal)", padding=6)
        topo_curso.grid(row=0, column=0, columnspan=2, sticky="we", pady=(0, 6))

        ttk.Label(topo_curso, textvariable=self.curso_revisao_nome, font=("Segoe UI Semibold", 11)).pack(side="left", padx=(0, 8))
        self.amostra_cor_revisao = tk.Canvas(
            topo_curso, width=26, height=18, background=COR_CAMPO, highlightthickness=1,
            highlightbackground=COR_BORDA,
        )
        self.amostra_cor_revisao.pack(side="left", padx=(0, 10))

        self.combo_curso_direto = ttk.Combobox(
            topo_curso, textvariable=self.curso_para_foto, width=28,
            postcommand=lambda: self.combo_curso_direto.configure(values=self.cursos_disponiveis),
        )
        self.combo_curso_direto.pack(side="left", padx=(0, 6))
        self.combo_curso_direto.bind("<<ComboboxSelected>>", lambda e: self._definir_curso_direto())
        self.combo_curso_direto.bind("<Return>", lambda e: self._definir_curso_direto())

        ttk.Button(
            topo_curso, text="Gravar Curso", command=self._definir_curso_direto
        ).pack(side="left", padx=(0, 6))

        ttk.Button(
            topo_curso, text="Aplicar nos Sem Curso", command=self._aplicar_curso_sem_curso
        ).pack(side="left", padx=(0, 6))

        ttk.Button(
            topo_curso,
            text="✓ Aprovar e Ir p/ Próximo Pendente (Enter)",
            style="Primario.TButton",
            command=self.aprovar_e_proximo_pendente,
        ).pack(side="left", padx=(0, 8))

        ttk.Checkbutton(
            topo_curso,
            text="Travar no aluno atual (sem troca auto)",
            variable=self.trava_aluno_revisao,
        ).pack(side="right")

        ttk.Label(painel, textvariable=self.curso_revisao_cor, style="Dica.TLabel").grid(
            row=1, column=0, sticky="w", pady=(0, 4))

        imagens = ttk.LabelFrame(painel, text="Prévia Grande da Capa (Estabilizada em Memória)", padding=4)
        imagens.grid(row=2, column=0, columnspan=2, sticky="nsew")
        imagens.grid_propagate(False)
        imagens.columnconfigure(0, weight=1)
        imagens.rowconfigure(0, weight=1)
        self._img_revisao_capa = None
        self.previa_revisao_capa = tk.Label(imagens, background=COR_FUNDO)
        self.previa_revisao_capa.grid(row=0, column=0, sticky="nsew")
        self.quadro_previa_revisao = imagens
        imagens.bind("<Configure>", self._ao_configurar_previa_revisao)

        # Galeria Fixa em Grid das 7 Fotos Sugeridas (Abaixo da Prévia Grande — Troca em 1 Clique)
        self.quadro_grid_7_fotos = ttk.LabelFrame(
            painel,
            text="Galeria das 7 Fotos Sugeridas (Troca em 1 Clique · Indicadores Contraste / Nitidez / Brilho)",
            padding=4,
        )
        self.quadro_grid_7_fotos.grid(row=3, column=0, columnspan=2, sticky="we", pady=(6, 0))
        self._botoes_grid_7 = []
        for idx_g in range(7):
            self.quadro_grid_7_fotos.columnconfigure(idx_g, weight=1)
            btn_card = tk.Button(
                self.quadro_grid_7_fotos,
                text=f"#{idx_g + 1}\n--\nC:-- N:-- B:--",
                font=("Consolas", 8),
                background=COR_CAMPO,
                foreground=COR_TEXTO,
                activebackground=COR_DESTAQUE,
                activeforeground="#0F172A",
                relief="solid",
                bd=1,
                padx=4,
                pady=3,
                cursor="hand2",
                state="disabled",
                command=lambda i=idx_g: self._trocar_por_opcao_indice(i),
            )
            btn_card.grid(row=0, column=idx_g, sticky="nsew", padx=2, pady=1)
            self._botoes_grid_7.append(btn_card)

        quadro_nome = ttk.LabelFrame(painel, text="Nome nesta capa", padding=6)
        quadro_nome.grid(row=4, column=0, columnspan=2, sticky="we", pady=(6, 0))
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
            quadro_nome, text="Salvar Nome", style="Primario.TButton", command=self._salvar_ajuste_nome, state="disabled")
        self.botao_nome_salvar.grid(row=2, column=5, padx=(6, 0), pady=(4, 0))
        ttk.Label(quadro_nome, textvariable=self.status_nome, wraplength=540).grid(
            row=3, column=0, columnspan=6, sticky="w", pady=(4, 0))
        self.escala1_nome.trace_add("write", lambda *a: self._escala_nome_mudou(1))
        self.escala2_nome.trace_add("write", lambda *a: self._escala_nome_mudou(2))

        acoes = ttk.Frame(painel)
        acoes.grid(row=5, column=0, columnspan=2, sticky="we", pady=(6, 0))
        ttk.Button(acoes, text="Atualizar revisão", command=self._atualizar_revisao_manual).pack(side="left")
        ttk.Button(acoes, text="Refazer prévia do PSD", command=self._excluir_previa_revisao).pack(
            side="left", padx=6)
        ttk.Button(acoes, text="Ocultar da revisão", command=self._excluir_da_revisao).pack(side="left")
        ttk.Label(painel, textvariable=self.status_revisao, wraplength=560, style="Dica.TLabel").grid(
            row=6, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # ---------------- Coluna Direita (PEQUENA): Recorte Compacto da Foto ----------------

        quadro_recorte = ttk.LabelFrame(
            corpo, text="Recorte da Foto (arraste / rodinha)", padding=6
        )
        quadro_recorte.grid(row=0, column=2, rowspan=2, sticky="ns")

        self.tela = tk.Canvas(
            quadro_recorte, width=260, height=250, background=COR_FUNDO,
            highlightthickness=0, cursor="fleur",
        )
        self.tela.pack(fill="x")
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

        ferramentas = ttk.Frame(quadro_recorte)
        ferramentas.pack(fill="x", pady=(6, 0))
        self.botoes_editor = []
        for texto, comando in (
            ("Auto", self.auto_rosto),
            ("Centrar", self.centralizar_rosto),
            ("Desfazer", self.desfazer),
        ):
            b = ttk.Button(ferramentas, text=texto, command=comando)
            b.pack(side="left", expand=True, fill="x", padx=1)
            self.botoes_editor.append(b)

        quadro_zoom = ttk.Frame(quadro_recorte)
        quadro_zoom.pack(fill="x", pady=(6, 0))
        ttk.Label(quadro_zoom, text="Zoom:").pack(side="left")
        self.controle_zoom = ttk.Scale(
            quadro_zoom, from_=15, to=100, variable=self.zoom, orient="horizontal", length=190
        )
        self.controle_zoom.pack(side="left", fill="x", expand=True, padx=(6, 0))
        for evento in ("<B1-Motion>", "<ButtonRelease-1>"):
            self.controle_zoom.bind(evento, lambda e: self.root.after_idle(self._zoom_escala, self.zoom.get()), add="+")

        acoes_foto = ttk.Frame(quadro_recorte)
        acoes_foto.pack(fill="x", pady=(8, 0))
        self.botao_aprovar_pendente = ttk.Button(
            acoes_foto,
            text="✓ Aprovar e Próximo Pendente (Enter)",
            style="Primario.TButton",
            command=self.aprovar_e_proximo_pendente,
        )
        self.botao_aprovar_pendente.pack(fill="x")
        self.botao_salvar = ttk.Button(
            acoes_foto, text="Só Salvar Recorte desta Foto",
            command=lambda: self.salvar(False),
        )
        self.botao_salvar.pack(fill="x", pady=(4, 0))
        self.botao_salvar_prox = ttk.Button(
            acoes_foto, text="Salvar e Próxima Foto", command=lambda: self.salvar(True)
        )
        self.botao_salvar_prox.pack(fill="x", pady=(4, 0))
        self.botao_restaurar = ttk.Button(
            acoes_foto, text="Restaurar Foto Original", command=self.restaurar
        )
        self.botao_restaurar.pack(fill="x", pady=(4, 0))
        self.tela.bind("<Return>", lambda e: self.aprovar_e_proximo_pendente())

        quadro_opcoes = ttk.LabelFrame(
            quadro_recorte, text="7 Melhores da Pasta (Troca 1 Clique)", padding=4
        )
        quadro_opcoes.pack(fill="x", pady=(8, 0))
        self.opcao_foto_var = tk.StringVar(value="")
        self._opcoes_atuais_lista = []
        self.combo_opcoes_foto = ttk.Combobox(
            quadro_opcoes, textvariable=self.opcao_foto_var, state="disabled", width=28
        )
        self.combo_opcoes_foto.pack(fill="x")
        self.combo_opcoes_foto.bind("<<ComboboxSelected>>", lambda e: self._trocar_por_opcao_selecionada())
        self.botao_trocar_opcao = ttk.Button(
            quadro_opcoes, text="Trocar por esta Foto", command=self._trocar_por_opcao_selecionada, state="disabled"
        )
        self.botao_trocar_opcao.pack(fill="x", pady=(4, 0))
        self.botao_abrir_tela_7 = ttk.Button(
            quadro_opcoes,
            text="Abrir Tela Visual: 7 Melhores Fotos",
            command=lambda: self.janela.abrir_tela_sete_melhores(
                self.rf.nome_do_aluno(self.arquivo) if (self.arquivo and self.rf) else None
            ),
        )
        self.botao_abrir_tela_7.pack(fill="x", pady=(4, 0))

        ttk.Label(quadro_recorte, textvariable=self.foto_revisao_atual, style="Dica.TLabel", wraplength=250).pack(
            anchor="w", pady=(8, 0)
        )
        ttk.Label(quadro_recorte, textvariable=self.info, style="Dica.TLabel", justify="left", wraplength=250).pack(
            anchor="w", pady=(4, 0)
        )
        tk.Label(
            quadro_recorte, textvariable=self.alerta, foreground=COR_VERMELHO,
            background=COR_SUPERFICIE, justify="left", wraplength=250, anchor="w",
        ).pack(fill="x", pady=(2, 0))


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
        # Se a aba de Revisão já está carregada e travada no aluno atual, bloqueia reconstrução automática!
        if self._revisao_visivel() and self._revisao_por_curso and self.trava_aluno_revisao.get():
            self._revisao_suja = True
            return
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
        if self._revisao_por_curso and self.trava_aluno_revisao.get() and self._aluno_escolhido:
            return
        if self._nome_editado:
            self.status_revisao.set("Há mudanças para mostrar: a revisão atualiza depois que você salvar o nome.")
            self._revisao_agendada = self.root.after(1500, self._rodar_revisao_agendada)
            return
        self._recarregar_revisao()

    def abrir_revisao(self):

        if self._revisao_suja and not self._revisao_por_curso:
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
                if hasattr(rf, "verificar_cache_analise_subpastas"):
                    rf.verificar_cache_analise_subpastas(cfg["fotos"])
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
                        tamanho_foto = amostra["tamanho_foto"]
                        foto_recortada = self._recorte_revisao(rf, arquivo, tamanho_foto, cfg, calibracao)
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

        # Se o modo travado no aluno atual está ativo e já existe um aluno em revisão, NÃO troca a capa para a amostra!
        tem_aluno_travado = bool(self.trava_aluno_revisao.get() and self._aluno_escolhido and self._aluno_nome)

        amostra = grupo.get("amostra")
        if not tem_aluno_travado and amostra and amostra.get("foto"):
            arq_amostra = Path(amostra["foto"])
            self.foto_revisao_atual.set(arq_amostra.name)
            if self.lista.exists(arq_amostra.name):
                self.lista.selection_set(arq_amostra.name)
                self.lista.see(arq_amostra.name)
            if self.arquivo != arq_amostra and not self.ocupado:
                self.abrir(arq_amostra)
        if not tem_aluno_travado:
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

        if not tem_aluno_travado:
            self._mostrar_previa_revisao(grupo["capa"], forcar=True)
        self._preencher_alunos_revisao()

        if grupo["avisos"]:
            self.status_revisao.set(" | ".join(dict.fromkeys(grupo["avisos"]))[:500])

    def _ao_configurar_previa_revisao(self, evento=None):
        """Debounce no redimensionamento da prévia grande para nunca travar a UI."""
        if self._capa_revisao_atual is None:
            return
        largura = max(200, self.quadro_previa_revisao.winfo_width() - 4)
        altura = max(150, self.quadro_previa_revisao.winfo_height() - 4)
        ult_w, ult_h = self._ultimo_tam_previa_revisao
        if abs(largura - ult_w) < 6 and abs(altura - ult_h) < 6:
            return
        if self._redimensionar_previa_pendente:
            self.root.after_cancel(self._redimensionar_previa_pendente)
        self._redimensionar_previa_pendente = self.root.after(
            60, lambda: self._mostrar_previa_revisao(self._capa_revisao_atual, forcar=True)
        )

    def _mostrar_previa_revisao(self, capa, forcar=False):

        mesma_capa = capa is self._capa_revisao_atual
        self._capa_revisao_atual = capa
        if capa is None:
            self.previa_revisao_capa.configure(image="")
            self._img_revisao_capa = None
            return
        largura = max(200, self.quadro_previa_revisao.winfo_width() - 4)
        altura = max(150, self.quadro_previa_revisao.winfo_height() - 4)
        ult_w, ult_h = self._ultimo_tam_previa_revisao
        if mesma_capa and not forcar and abs(largura - ult_w) < 6 and abs(altura - ult_h) < 6 and self._img_revisao_capa:
            return
        self._ultimo_tam_previa_revisao = (largura, altura)
        imagem = capa.copy()
        imagem.thumbnail((largura, altura), Image.BILINEAR)
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
                if (entrada["nome"], entrada["tamanho_capa"]) not in self._aprovados_revisao
                and (
                    not grupo.get("simbolo")
                    or self._situacao_nome(entrada.get("layout"))[1] in ("nome_estoura", "nome_reduzido")
                )
            ]
        else:
            grupo = self._revisao_por_curso.get(self._curso_revisao_selecionado)
            pares = [(grupo, entrada) for entrada in grupo["entradas"]] if grupo else []
        pares.sort(key=lambda par: par[1]["nome"].casefold())
        for indice, (grupo, entrada) in enumerate(pares):
            chave_ap = (entrada["nome"], entrada["tamanho_capa"])
            if chave_ap in self._aprovados_revisao:
                texto, tag = "Aprovado ✓", "aluno_aprovado"
            elif not grupo.get("simbolo"):
                texto, tag = "Sem Curso", "nome_estoura"
            else:
                texto, tag = self._situacao_nome(entrada.get("layout"))
            iid = "aluno_" + str(indice)
            self.lista_alunos.insert(
                "", "end", iid=iid, values=(entrada["nome"], entrada["tamanho_capa"], texto),
                tags=(tag,) if tag else (),
            )
            self._alunos_por_iid[iid] = (grupo, entrada)
        if not (self.trava_aluno_revisao.get() and self._aluno_nome):
            self._limpar_painel_nome()
        alvo = next((iid for iid, (_, entrada) in self._alunos_por_iid.items()
                     if (entrada["nome"], entrada["tamanho_capa"]) == self._aluno_escolhido), None)
        if alvo:
            self.lista_alunos.selection_set(alvo)
            self.lista_alunos.see(alvo)
            if not self._aluno_nome:
                self._selecionar_aluno_revisao()

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
        self._token_selecao_aluno += 1
        token_atual = self._token_selecao_aluno
        self._aluno_nome = (grupo, entrada)
        self._aluno_escolhido = (entrada["nome"], entrada["tamanho_capa"])
        self._previa_aluno = None
        self._preencher_controles_nome(entrada)
        self.curso_revisao_nome.set(grupo["curso"])
        if grupo.get("simbolo") and grupo["curso"] in self.cursos_disponiveis:
            self.curso_para_foto.set(grupo["curso"])
        elif entrada.get("curso") and entrada["curso"] in self.cursos_disponiveis:
            self.curso_para_foto.set(entrada["curso"])
        self.nome_revisao_titulo.set(
            entrada["nome"] + "  -  capa " + entrada["tamanho_capa"] + "  (" + grupo["curso"] + ")  [Fixo]")
        if entrada.get("foto"):
            arq_aluno = Path(entrada["foto"])
            self.foto_revisao_atual.set(arq_aluno.name)
            if self.lista.exists(arq_aluno.name):
                self.lista.selection_set(arq_aluno.name)
                self.lista.see(arq_aluno.name)
            if self.arquivo != arq_aluno and not self.ocupado:
                self.abrir(arq_aluno)
        else:
            self.foto_revisao_atual.set("Sem foto: prévia com fundo neutro")
        self.status_nome.set("Montando a prévia leve desta capa em memória...")
        cfg = dict(self.cfg)
        calibracao = dict(self._calibracao())

        def trabalho():
            recortada = preset = previa = erro = None
            try:
                tamanho = entrada["tamanho_capa"]
                simbolo = grupo.get("simbolo")
                if simbolo:
                    cor = None if tamanho == "24X34" else grupo.get("cor")
                    chave_preset = (tamanho, simbolo, tuple(cor) if cor else None)
                    preset = self._presets_curso.get(chave_preset) or grupo.get("preset")
                    if preset is None:
                        preset = rf.gerar_presets_curso(
                            tamanho, {simbolo: tuple(cor) if cor else None}).get(simbolo)
                        if preset is not None:
                            self._presets_curso[chave_preset] = preset
                if preset is None:
                    previa = self._previas_capa.get(tamanho)
                    if previa is None:
                        previa = rf.ler_previa_capa(tamanho)
                        if previa is not None:
                            self._previas_capa[tamanho] = previa
                if entrada.get("foto"):
                    arq_f = Path(entrada["foto"])
                    try:
                        st_f = arq_f.stat()
                        ch_rec = (str(arq_f), st_f.st_mtime_ns, entrada["tamanho_foto"])
                    except OSError:
                        ch_rec = None
                    recortada = self._cache_recorte_revisao.get(ch_rec) if ch_rec else None
                    if recortada is None:
                        recortada = self._recorte_revisao(
                            rf, arq_f, entrada["tamanho_foto"], cfg, calibracao)
                        if ch_rec and recortada is not None:
                            if len(self._cache_recorte_revisao) > 40:
                                self._cache_recorte_revisao.clear()
                            self._cache_recorte_revisao[ch_rec] = recortada
                else:
                    recortada = Image.new("RGB", (400, 300), "#808080")
            except Exception as e:
                erro = e
            self.root.after(0, self._capa_aluno_pronta, token_atual, entrada, recortada, preset, previa, erro)

        threading.Thread(target=trabalho, daemon=True).start()

    @staticmethod
    def _recorte_revisao(rf, arquivo, tamanho_foto, cfg, calibracao):

        origem_foto = rf.fonte_original(arquivo, cfg)
        if hasattr(rf, "abrir_foto_previa_rapida"):
            imagem_red, tam_orig, _ = rf.abrir_foto_previa_rapida(origem_foto, 1400)
        else:
            imagem, _, _ = rf.abrir_foto(origem_foto)
            tam_orig = imagem.size
            fator = min(1.0, 1400.0 / max(tam_orig))
            imagem_red = (
                imagem.resize((max(1, round(tam_orig[0] * fator)), max(1, round(tam_orig[1] * fator))), Image.BILINEAR)
                if fator < 1.0 else imagem
            )
        molde = cfg["moldes"][tamanho_foto]
        caixa = rf.caixa_manual(arquivo, tamanho_foto, tam_orig, molde)
        if caixa is None:
            if rf.no_molde(tam_orig, molde):
                caixa = rf.ajustar_caixa((0, 0, tam_orig[0], tam_orig[1]), tam_orig, molde)
            else:
                rosto = rf.detectar_rosto_cache(origem_foto) if hasattr(rf, "detectar_rosto_cache") else rf.detectar_rosto(imagem_red)
                caixa = (
                    rf.calcular_recorte(tam_orig, rosto, molde, calibracao.get(tamanho_foto, rf.CALIBRACAO_PADRAO))
                    if rosto else rf.ajustar_caixa((0, 0, tam_orig[0], tam_orig[1]), tam_orig, molde)
                )
        fx = imagem_red.width / float(max(1, tam_orig[0]))
        fy = imagem_red.height / float(max(1, tam_orig[1]))
        caixa_red = (
            max(0, round(caixa[0] * fx)),
            max(0, round(caixa[1] * fy)),
            min(imagem_red.width, max(1, round(caixa[2] * fx))),
            min(imagem_red.height, max(1, round(caixa[3] * fy))),
        )
        recortada = imagem_red.crop(caixa_red).convert("RGB")
        return recortada

    def _capa_aluno_pronta(self, token, entrada, recortada, preset, previa, erro):

        if token != self._token_selecao_aluno:
            return
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
        if self.lista.exists(arquivo.name):
            self.lista.selection_set(arquivo.name)
            self.lista.see(arquivo.name)
        if self.arquivo != arquivo:
            self.abrir(arquivo)
        self.tela.focus_set()

    def selecionar_aluno_por_nome(self, nome_aluno):
        """Seleciona o aluno e seu curso na aba unificada Revisão e Recorte."""

        rf = self._modulo()
        alvo_chave = rf.chave(nome_aluno)
        for chave_grupo, grupo in self._revisao_por_curso.items():
            for entrada in grupo.get("entradas", []):
                if rf.chave(entrada.get("nome", "")) == alvo_chave:
                    self._curso_revisao_selecionado = chave_grupo
                    self._aluno_escolhido = (entrada["nome"], entrada["tamanho_capa"])
                    for iid, cg in self._revisao_grupo_por_iid.items():
                        if cg == chave_grupo and self.lista_revisao.exists(iid):
                            self.lista_revisao.selection_set(iid)
                            self.lista_revisao.see(iid)
                            break
                    self._selecionar_curso_revisao()
                    for iid, (_, ent) in self._alunos_por_iid.items():
                        if rf.chave(ent.get("nome", "")) == alvo_chave and self.lista_alunos.exists(iid):
                            self.lista_alunos.selection_set(iid)
                            self.lista_alunos.see(iid)
                            self._selecionar_aluno_revisao()
                            break
                    return

    def _definir_curso_direto(self):
        """Grava imediatamente o curso selecionado na barra acima da prévia para o aluno/foto atual e atualiza a lista."""

        curso = self.curso_para_foto.get().strip()
        if not curso:
            self.status_revisao.set("Escolha um curso na caixa acima da prévia.")
            return
        entrada, _ = self._alvo_troca_curso()
        arquivo = Path(entrada["foto"]) if entrada and entrada.get("foto") else self.arquivo
        if not arquivo:
            self.status_revisao.set("Selecione um aluno ou uma foto para gravar o curso.")
            return
        if self._salvar_curso_foto(arquivo, curso):
            if entrada:
                self._aluno_escolhido = (entrada["nome"], entrada["tamanho_capa"])
            try:
                automacao = modulo_automacao()
                canonico = automacao.CURSO_ALIASES.get(automacao.chave(curso), curso)
                self._curso_revisao_selecionado = automacao.chave(canonico)
            except Exception:
                pass

    def _aplicar_curso_sem_curso(self):
        """Aplica o curso selecionado no topo a todos os alunos que estão sem curso na revisão."""

        curso = self.curso_para_foto.get().strip()
        if not curso:
            self.status_revisao.set("Escolha um curso na caixa ao lado antes de aplicar aos alunos sem curso.")
            return
        try:
            simbolo, _ = modulo_automacao().resolver_curso(curso)
            if not simbolo:
                raise ValueError("Curso sem brasão válido.")
        except Exception as e:
            self.status_revisao.set("Curso inválido: " + str(e))
            return

        rf = self._modulo()
        alterados = 0
        for chave_grupo, grupo in list(self._revisao_por_curso.items()):
            if grupo.get("simbolo"):
                continue
            for entrada in grupo.get("entradas", []):
                if entrada.get("foto"):
                    arq = Path(entrada["foto"])
                    self.cursos_manuais[chave_caminho_foto(arq)] = curso
                    alterados += 1
                    try:
                        self.janela.sincronizar_curso_aluno(entrada["nome"], curso)
                    except Exception:
                        pass
        if not alterados:
            # Se não havia grupo sem símbolo, aplica ao grupo atualmente selecionado se for sem curso
            entrada, _ = self._alvo_troca_curso()
            if entrada and entrada.get("foto"):
                self._definir_curso_direto()
                return
            self.status_revisao.set("Todos os alunos já possuem curso definido.")
            return

        try:
            self._gravar_cursos_foto()
            automacao = modulo_automacao()
            canonico = automacao.CURSO_ALIASES.get(automacao.chave(curso), curso)
            self._curso_revisao_selecionado = automacao.chave(canonico)
        except Exception:
            pass
        self.status_revisao.set("Curso " + curso + " gravado para " + str(alterados) + " aluno(s) e salvo na lista!")
        self._recarregar_revisao()

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
        try:
            rf = self._modulo()
            nome_aluno = rf.nome_do_aluno(Path(arquivo))
            self.janela.sincronizar_curso_aluno(nome_aluno, curso)
        except Exception:
            pass
        self.status_revisao.set("Curso " + curso + " salvo para " + arquivo.name + " e gravado na lista de Capas.")
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
                if hasattr(rf, "verificar_cache_analise_subpastas"):
                    rf.verificar_cache_analise_subpastas(pasta)
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

    def aprovar_e_proximo_pendente(self):
        """Aprova a capa do formando atual (salva recorte se alterado) e pula direto para o próximo pendente."""
        if self._aluno_nome:
            grupo_at, ent_at = self._aluno_nome
            chave_at = (ent_at["nome"], ent_at["tamanho_capa"])
            self._aprovados_revisao.add(chave_at)
            if not grupo_at.get("simbolo") and self.curso_para_foto.get().strip():
                self._definir_curso_direto()

        # Procura o próximo aluno com pendência (sem curso ou nome estourando/reduzido) ainda não aprovado
        candidatos_pendentes = []
        for g in self._revisao_por_curso.values():
            for ent in g.get("entradas", []):
                ch = (ent["nome"], ent["tamanho_capa"])
                if ch in self._aprovados_revisao:
                    continue
                _, tag_n = self._situacao_nome(ent.get("layout"))
                if (not g.get("simbolo")) or tag_n in ("nome_estoura", "nome_reduzido"):
                    candidatos_pendentes.append((g, ent))

        if candidatos_pendentes:
            prox_g, prox_ent = candidatos_pendentes[0]
            self._aluno_escolhido = (prox_ent["nome"], prox_ent["tamanho_capa"])
            self._preencher_alunos_revisao()
            alvo = next(
                (iid for iid, (_, e) in self._alunos_por_iid.items() if (e["nome"], e["tamanho_capa"]) == self._aluno_escolhido),
                None,
            )
            if alvo:
                self.lista_alunos.selection_set(alvo)
                self.lista_alunos.see(alvo)
                self._selecionar_aluno_revisao()
            self.status_revisao.set(
                f"✓ Aprovado! Avançou direto para o próximo pendente: {prox_ent['nome']} ({prox_ent['tamanho_capa']})."
            )
            return

        self._preencher_alunos_revisao()
        self.proxima()
        self.status_revisao.set("✓ Capa aprovada! Todas as pendências prioritárias foram conferidas.")

    def proxima(self):

        itens_alunos = self.lista_alunos.get_children()
        if itens_alunos:
            atual_aluno = self.lista_alunos.selection()
            indice_a = itens_alunos.index(atual_aluno[0]) + 1 if atual_aluno and atual_aluno[0] in itens_alunos else 0
            if indice_a < len(itens_alunos):
                self.lista_alunos.selection_set(itens_alunos[indice_a])
                self.lista_alunos.see(itens_alunos[indice_a])
                self._selecionar_aluno_revisao()
                return
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
        self._ocupar(True, "Abrindo " + arquivo.name + "...")

        def trabalho():
            try:
                origem = rf.fonte_original(arquivo, self.cfg)
                if hasattr(rf, "abrir_foto_previa_rapida"):
                    reduzida, tam_orig, fator = rf.abrir_foto_previa_rapida(origem, LADO_EDITOR)
                    fonte = None
                else:
                    fonte = rf.abrir_foto(origem)
                    img = fonte[0]
                    tam_orig = img.size
                    reduzida = img.convert("RGB")
                    fator = min(1.0, LADO_EDITOR / max(tam_orig))
                    if fator < 1.0:
                        reduzida = reduzida.resize((round(tam_orig[0] * fator), round(tam_orig[1] * fator)), Image.BILINEAR)
                rosto = rf.detectar_rosto_cache(origem) if hasattr(rf, "detectar_rosto_cache") else rf.detectar_rosto(reduzida)
                self.root.after(0, self._foto_aberta, arquivo, origem, fonte, tam_orig, reduzida, fator, rosto, tam, calib, None)
            except Exception as e:
                self.root.after(0, self._foto_aberta, arquivo, None, None, None, None, 1.0, None, tam, calib, e)

        threading.Thread(target=trabalho, daemon=True).start()

    def _foto_aberta(self, arquivo, origem, fonte, tam_orig, reduzida, fator, rosto, tam, calib, erro):

        self._ocupar(False)
        if erro:
            self.status.set("Erro ao abrir " + arquivo.name + ": " + str(erro))
            return

        rf = self.rf
        self.arquivo = arquivo
        self.fonte = fonte if fonte is not None else True
        self.tam_fonte_orig = tam_orig
        self.reduzida = reduzida
        self.fator_red = fator
        self.rosto = rosto
        self.historico = []
        self._img_tela_chave = None
        self._atualizar_foto_revisao_atual()
        self.tam.set(tam)
        self._mostrar_molde()

        W_orig, H_orig = tam_orig
        molde = self.cfg["moldes"][tam]
        self.caixa_auto = (
            rf.calcular_recorte(tam_orig, rosto, molde, calib[tam]) if rosto
            else rf.ajustar_caixa((0, 0, W_orig, H_orig), tam_orig, molde)
        )

        manual = rf.caixa_manual(arquivo, tam, tam_orig, molde)
        ja_no_molde = rf.no_molde(tam_orig, molde)
        if manual:
            self.caixa = manual
            origem_caixa = "corte salvo à mão"
        elif ja_no_molde:
            self.caixa = rf.ajustar_caixa((0, 0, W_orig, H_orig), tam_orig, molde)
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
            arquivo.name + "  |  " + str(W_orig) + "x" + str(H_orig)
            + ("  (original de _ORIGINAIS)" if usando_original else "")
            + "  |  " + origem_caixa
        )
        self._sincronizar_zoom()
        self._desenhar()
        self._agendar_previa()
        self._atualizar_opcoes_foto_aluno()
        self.tela.focus_set()

    def _atualizar_opcoes_foto_aluno(self):
        """Preenche a caixa de Troca Rápida com as melhores fotos avaliadas da subpasta deste formando."""

        self._opcoes_atuais_lista = []
        self._aviso_opcoes_foto_atual = ""
        if not self.arquivo or not self.rf or not hasattr(self.rf, "ler_opcoes_foto"):
            self.opcao_foto_var.set("")
            self.combo_opcoes_foto.configure(values=[], state="disabled")
            self.botao_trocar_opcao.configure(state="disabled")
            return

        nome_aluno = self.rf.nome_do_aluno(self.arquivo)
        tam = self.tam.get()
        chave_op = f"{self.rf.chave(nome_aluno)}|{tam}"
        dados_op = self.rf.ler_opcoes_foto().get(chave_op)
        candidatas = (dados_op or {}).get("candidatas") or []
        if not candidatas:
            self.opcao_foto_var.set("Use 'Escolher Melhores Fotos'")
            self.combo_opcoes_foto.configure(values=[], state="disabled")
            self.botao_trocar_opcao.configure(state="disabled")
            for idx_b, btn in enumerate(getattr(self, "_botoes_grid_7", [])):
                btn.configure(text=f"#{idx_b + 1}\n--\nC:-- N:-- B:--", state="disabled", background=COR_CAMPO)
            return

        self._aviso_opcoes_foto_atual = str((dados_op or {}).get("aviso_revisao") or "")
        escolhida_atual = str((dados_op or {}).get("escolhida") or "")
        rotulos = []
        indice_atual = 0
        for idx, cand in enumerate(candidatas[:7]):
            tag_aviso = " [AVISO]" if not cand.get("aprovada_qualidade", True) else ""
            c_pct = int(round(float(cand.get("contraste_score", 75))))
            n_pct = int(round(float(cand.get("nitidez_score") or min(100.0, float(cand.get("nitidez", 120)) / 1.8))))
            b_pct = int(round(float(cand.get("brilho_score", 85))))
            rotulo = (
                f"{idx + 1}ª ({cand.get('nota', 0)} pts){tag_aviso} "
                f"[C:{c_pct}% N:{n_pct}% B:{b_pct}%] "
                f"{cand.get('nome_arquivo', '')} — {cand.get('resumo', '')}"
            )
            rotulos.append(rotulo)
            self._opcoes_atuais_lista.append(cand)
            if str(cand.get("caminho", "")) == escolhida_atual:
                indice_atual = idx

        # Atualiza os 7 cartões da Galeria Fixa abaixo da Prévia Grande (com indicador de 1 Rosto)
        for idx_b, btn in enumerate(getattr(self, "_botoes_grid_7", [])):
            if idx_b < len(self._opcoes_atuais_lista):
                c_item = self._opcoes_atuais_lista[idx_b]
                ativa = idx_b == indice_atual
                c_pct = int(round(float(c_item.get("contraste_score", 75))))
                n_pct = int(round(float(c_item.get("nitidez_score") or min(100.0, float(c_item.get("nitidez", 120)) / 1.8))))
                b_pct = int(round(float(c_item.get("brilho_score", 85))))
                qtd_p = int(c_item.get("qtd_pessoas") or 1)
                tag_rosto = "1Rosto✓" if qtd_p == 1 else f"{qtd_p}Pess!"
                orient_curta = "Vert✓" if c_item.get("orientacao_ideal") else str(c_item.get("orientacao", ""))[:4]
                txt_card = (
                    f"{'★ ' if ativa else ''}#{idx_b + 1} · {tag_rosto} ({orient_curta})\n"
                    f"{str(c_item.get('nome_arquivo', ''))[:16]}\n"
                    f"C:{c_pct}%  N:{n_pct}%  B:{b_pct}%"
                )
                btn.configure(
                    text=txt_card,
                    state="normal",
                    background="#1e3a5f" if ativa else COR_CAMPO,
                    foreground="#4ade80" if ativa else COR_TEXTO,
                )
            else:
                btn.configure(text=f"#{idx_b + 1}\n--\nC:-- N:-- B:--", state="disabled", background=COR_CAMPO)

        self.combo_opcoes_foto.configure(values=rotulos, state="readonly")
        self.combo_opcoes_foto.current(indice_atual)
        self.botao_trocar_opcao.configure(state="normal")
        self._atualizar_info()

    def _trocar_por_opcao_indice(self, idx):
        """Troca em 1 clique diretamente pelo cartão da Galeria Fixa das 7 Fotos Sugeridas."""
        if not self._opcoes_atuais_lista or idx < 0 or idx >= len(self._opcoes_atuais_lista):
            return
        self.combo_opcoes_foto.current(idx)
        self._trocar_por_opcao_selecionada()

    def _trocar_por_opcao_selecionada(self):
        """Substitui a foto atual do formando pela foto candidata escolhida sem perder o foco do aluno."""

        if not self.arquivo or not self._opcoes_atuais_lista or self.ocupado:
            return
        idx = self.combo_opcoes_foto.current()
        if idx < 0 or idx >= len(self._opcoes_atuais_lista):
            return
        cand = self._opcoes_atuais_lista[idx]
        caminho_origem = cand.get("caminho")
        if not caminho_origem:
            return

        rf = self._modulo()
        nome_aluno = rf.nome_do_aluno(self.arquivo)
        tam = self.tam.get()
        try:
            destino = rf.trocar_foto_por_opcao(nome_aluno, tam, caminho_origem, self.cfg)
            self._conhecer_arquivo("fotos", self.cfg["fotos"])
            self._cache_recorte_revisao.clear()
        except Exception as e:
            self.status_revisao.set("Não foi possível trocar pela foto candidata: " + str(e))
            return

        c_pct = int(round(float(cand.get("contraste_score", 75))))
        n_pct = int(round(float(cand.get("nitidez_score") or min(100.0, float(cand.get("nitidez", 120)) / 1.8))))
        b_pct = int(round(float(cand.get("brilho_score", 85))))
        self.status_revisao.set(
            f"Foto de {nome_aluno} ({tam}) trocada em 1 clique para {cand.get('nome_arquivo')} "
            f"({cand.get('nota')} pts | Contraste {c_pct}% · Nitidez {n_pct}% · Brilho {b_pct}%)!"
        )
        self.abrir(destino)
        if self._aluno_nome:
            self._selecionar_aluno_revisao()

    # --------------------------------------------------------
    # Geometria
    # --------------------------------------------------------

    def _molde(self):
        return self.cfg["moldes"][self.tam.get()]

    def _tamanho_fonte(self):
        if getattr(self, "tam_fonte_orig", None):
            return self.tam_fonte_orig
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
        escala = min(1.0, LADO_PREVIA_REVISAO / max(mw, mh))
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

        # Atualiza ao vivo a Prévia Grande da Capa no centro enquanto arrasta o recorte!
        if self._aluno_nome and self._previa_aluno and self._aluno_nome[1].get("foto") and self.arquivo:
            try:
                if Path(self._aluno_nome[1]["foto"]).resolve() == self.arquivo.resolve():
                    self._previa_aluno = (recortada, self._previa_aluno[1], self._previa_aluno[2])
                    self._renderizar_nome_revisao()
                elif imagem is not recortada:
                    self._mostrar_previa_revisao(imagem)
            except Exception:
                pass
        elif imagem is not recortada:
            self._mostrar_previa_revisao(imagem)

        imagem_min = imagem.copy()
        imagem_min.thumbnail((380, 380), Image.LANCZOS)
        fundo = Image.new("RGB", (380, 380), COR_FUNDO)
        fundo.paste(imagem_min, ((380 - imagem_min.width) // 2, (380 - imagem_min.height) // 2))
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
        if getattr(self, "_aviso_opcoes_foto_atual", ""):
            alertas.append("Aviso da seleção: " + self._aviso_opcoes_foto_atual)
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

        arquivo, caixa = self.arquivo, self.caixa
        fonte = self.fonte if isinstance(self.fonte, tuple) else None
        tam_fonte = self._tamanho_fonte()
        self._ocupar(True, "Gravando " + arquivo.name + "...")

        def trabalho():
            try:
                destino = rf.aplicar_recorte(arquivo, caixa, tam, self.cfg, fonte)
                rf.gravar_manual(arquivo, tam, caixa, tam_fonte)
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
    multiprocessing.freeze_support()
    main()
