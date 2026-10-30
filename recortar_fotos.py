"""
Recorte automático das fotos das capas (fora do Photoshop).

Pega as fotos já definidas em CAPAS\\FOTO que ainda não estão no
tamanho do molde, detecta o rosto e corta no tamanho exato em pixels
de cada álbum. O enquadramento é aprendido (--calibrar) das fotos que
já estão no tamanho certo.

O corte escolhido à mão na aba Fotos da UI fica em recorte_manual.json
e tem prioridade sobre o automático.

Uso:
    python recortar_fotos.py --calibrar
    python recortar_fotos.py --simular
    python recortar_fotos.py
    python recortar_fotos.py --aluno "Maria" --tamanho 20X20
    python recortar_fotos.py --gerar-previa-capa
"""

import argparse
import ctypes
import csv
import datetime
import functools
import hashlib
import io
import json
import math
import multiprocessing
import os
import re
import shutil
import statistics
import sys
import tempfile
import threading
import concurrent.futures
import unicodedata
import urllib.request
from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont, ImageOps


PASTA = Path(__file__).resolve().parent
CONFIG_FILE = PASTA / "automacao_capas_config.json"
CALIBRACAO_FILE = PASTA / "recorte_calibracao.json"
RELATORIO_CSV = PASTA / "recorte_fotos.csv"
MANUAL_FILE = PASTA / "recorte_manual.json"
PREVIA_CAPAS_DIR = PASTA / "previa_capas"
PREVIEW_DIR = Path(tempfile.gettempdir()) / "recorte_preview"
PASTA_PROXIES_ANALISE = PASTA / ".cache_proxies_capaz"

# PSD e caminho da camada FOTO, iguais aos modelos de AUTOMAÇAO-PASTA.py
PSD_POR_TAMANHO = {
    "20X20": ("CAPA 20X20 (ESSENCIAL BOOK) - ZANGRAF.psd", ["FOTO", "FOTO"]),
    "30X20": ("CAPA 30X20 (PRIME BOOK) - ZANGRAF.psd", ["FOTO"]),
    "24X34": ("CAPA 20X20 24x34 (BERLIM) - ZANGRAF.psd", ["24X34", "FOTO"]),
    "30X22": ("CAPA FRENTE 30x22.jpg.psd", ["FOTO"]),
}
FRENTE_POR_TAMANHO = {
    "20X20": ["BRASÃO", "FRENTE (PRINCIPAL) copiar"],
    "30X20": ["BRAÇÕES", "FRENTE"],
    "24X34": ["24X34", "BRAÇÕES", "FRENTE (PRINCIPAL) copiar 2"],
    "30X22": ["FRENTE"],
}
# Todos os brasões visíveis na capa (frente pequeno + verso grande)
BRASOES_POR_TAMANHO = {
    "20X20": [["BRASÃO", "FRENTE (PRINCIPAL) copiar"], ["BRASÃO", "FRENTE (PRINCIPAL)"]],
    "30X20": [["BRAÇÕES", "FRENTE"], ["BRAÇÕES", "FRENTE copiar"]],
    "24X34": [["24X34", "BRAÇÕES", "FRENTE (PRINCIPAL) copiar 2"]],
    "30X22": [["FRENTE"]],
}
# 30X22: o verso é outro PSD; na prévia ele aparece à esquerda da frente
VERSO_SEPARADO = {
    "30X22": {
        "psd": "CAPA VERSO 30x22.jpg.psd",
        "preenchimento": ["BACKGROUND", "Preenchimento de Cor 1"],
        "brasao": ["FRENTE copiar"],
    },
}
PREENCHIMENTO_POR_TAMANHO = {
    "20X20": ["BACKGROUND", "Preenchimento de Cor 1"],
    "30X20": ["BACKGROUND", "Preenchimento de Cor 1"],
    "30X22": ["Preenchimento de Cor 1"],
}
# Capas que usam a foto de outro tamanho (o 30X22 usa a mesma foto/molde do 30X20)
CAPAS_DO_TAMANHO = {
    "30X20": ["30X20", "30X22"],
}
LADO_PREVIA_CAPA = 2000
VERSAO_PRESET_CURSO = 5
CAMADAS_SIMPLES = {"pixel", "type", "shape", "smartobject", "solidcolorfill", "patternfill"}

# Textos do nome do aluno em cada PSD (iguais aos de AUTOMAÇAO-PASTA.py)
NOME_POR_TAMANHO = {
    "20X20": {"grupo": ["NOME"], "lombada": ["LATERAL", "NOME"]},
    "30X20": {"grupo": ["NOME"], "lombada": ["LATERAL", "NOME"]},
    "30X22": {"grupo": ["NOME"], "lombada": ["NOME"]},
    "24X34": {"smart": ["24X34", "NOME copiar"], "grupo": ["NOME"]},
}
# x_min: até onde (px do PSD) o nome alinhado à direita pode crescer para a esquerda.
# lombada: comprimento máximo do nome da lombada, em fração da altura da capa.
LIMITES_NOME = {
    "20X20": {"x_min": 3145, "lombada": 0.80},
    "30X20": {"x_min": 4420, "lombada": 0.80},
    "30X22": {"x_min": 795, "lombada": 0.80},
    "24X34": {"x_min": 400},
}
ESCALA_NOME_MIN = 50
ESCALA_NOME_MAX = 130
VERSAO_MODELO_NOME = 3
AJUSTES_NOME_FILE = PASTA / "ajustes_nome.json"
PREPOSICOES_NOME = {"DA", "DE", "DO", "DAS", "DOS", "E", "DI", "DU", "DEL", "VAN", "VON"}

MODELO_YUNET = PASTA / "face_detection_yunet_2023mar.onnx"
URL_YUNET = (
    "https://github.com/opencv/opencv_zoo/raw/main/models/"
    "face_detection_yunet/face_detection_yunet_2023mar.onnx"
)

EXTENSOES = {".jpg", ".jpeg", ".png", ".tif", ".tiff"}

MOLDES_PADRAO = {
    "20X20": (2694, 3012),
    "30X20": (3874, 3012),
    "24X34": (4160, 6240),
    "30X22": (3874, 3012),
}

# Usado enquanto um tamanho não tiver fotos prontas para calibrar
CALIBRACAO_PADRAO = {
    "centro_x": 0.5,
    "olhos_y": 0.38,
    "altura_rosto": 0.16,
}

TOLERANCIA_MOLDE_PX = 2
LADO_DETECCAO = 960
LADO_PROXY_LOCAL = 1024
QUALIDADE_PROXY_LOCAL = 82
LADO_PREVIA_RAPIDA = 2000
MAX_FOTOS_POR_ONDA_SUBPASTA = 7
MAX_TRIAGEM_ORIENTACAO_CABECALHO = 18
LIMITE_CPU_PCT_PADRAO = 60
USAR_GPU_OPENCL_PADRAO = True
USAR_COPIA_LEVE_LOCAL_PADRAO = True


# ============================================================
# CONFIG
# ============================================================

def carregar_config():
    dados = {}
    if CONFIG_FILE.is_file():
        dados = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))

    moldes = dict(MOLDES_PADRAO)
    for k, v in (dados.get("recorte_moldes") or {}).items():
        moldes[k.upper()] = (int(v[0]), int(v[1]))

    fotos = Path(dados.get("photos_dir") or ".")
    backup = Path(dados.get("recorte_backup_dir") or (fotos / "_ORIGINAIS"))

    return {
        "fotos": fotos,
        "backup": backup,
        "psd_dir": Path(dados.get("psd_dir") or PASTA),
        "photos_link": (dados.get("photos_link") or "") if dados.get("fotos_origem") == "drive" else "",
        "moldes": moldes,
        "ampliacao_maxima": float(dados.get("recorte_ampliacao_maxima", 1.2)),
        "hardware_limite_cpu_pct": int(dados.get("hardware_limite_cpu_pct", LIMITE_CPU_PCT_PADRAO)),
        "hardware_usar_gpu_opencl": bool(dados.get("hardware_usar_gpu_opencl", USAR_GPU_OPENCL_PADRAO)),
        "usar_copia_leve_local": bool(dados.get("usar_copia_leve_local", USAR_COPIA_LEVE_LOCAL_PADRAO)),
    }


def ler_config_bruto():
    if CONFIG_FILE.is_file():
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    return {}


def gravar_config(alteracoes):
    """Atualiza só as chaves informadas do automacao_capas_config.json."""
    dados = ler_config_bruto()
    dados.update(alteracoes)
    temp = CONFIG_FILE.with_suffix(".tmp")
    temp.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(temp, CONFIG_FILE)


def gravar_molde(tam, largura, altura):
    """Grava o tamanho em pixels de um molde em recorte_moldes do config."""
    moldes = ler_config_bruto().get("recorte_moldes") or {}
    moldes[tam.upper()] = [int(largura), int(altura)]
    gravar_config({"recorte_moldes": moldes})


# ============================================================
# GOOGLE DRIVE / PLANILHAS (links públicos: "qualquer pessoa com o link")
# ============================================================

NAVEGADOR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def eh_link(texto):
    return str(texto or "").strip().lower().startswith(("http://", "https://"))


def id_google(link):
    """ID do arquivo/pasta/planilha em qualquer formato de link do Google."""
    for padrao in (r"/folders/([\w-]{10,})", r"/spreadsheets/d/([\w-]{10,})",
                   r"/file/d/([\w-]{10,})", r"/d/([\w-]{10,})", r"[?&]id=([\w-]{10,})"):
        m = re.search(padrao, link)
        if m:
            return m.group(1)
    raise ValueError("Não reconheci o link do Google:\n" + link)


def _baixar(url, destino, aceitar_html=False):
    req = urllib.request.Request(url, headers=NAVEGADOR)
    with urllib.request.urlopen(req, timeout=120) as r:
        tipo = r.headers.get("Content-Type", "")
        dados = r.read()
    if not aceitar_html and "text/html" in tipo:
        raise RuntimeError(
            "O Google devolveu uma página em vez do arquivo.\n"
            "Confira se o compartilhamento está como \"Qualquer pessoa com o link\"."
        )
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    temp = destino.with_name(destino.name + ".baixando")
    temp.write_bytes(dados)
    os.replace(temp, destino)
    return destino


def baixar_planilha_link(link, destino):
    """Planilha do Google (exporta xlsx) ou arquivo .xlsx guardado no Drive."""
    ident = id_google(link)
    if "/spreadsheets/" in link:
        gid = re.search(r"[#&?]gid=(\d+)", link)
        url = f"https://docs.google.com/spreadsheets/d/{ident}/export?format=xlsx"
        if gid:
            url += "&gid=" + gid.group(1)
    else:
        url = f"https://drive.usercontent.google.com/download?id={ident}&export=download&confirm=t"
    destino = _baixar(url, destino)
    if destino.read_bytes()[:2] != b"PK":
        raise RuntimeError("O link não devolveu uma planilha .xlsx válida.")
    return destino


def listar_pasta_drive(link):
    """[(id, nome)] das fotos de uma pasta pública do Drive (só o primeiro nível)."""
    import html

    ident = id_google(link)
    url = f"https://drive.google.com/embeddedfolderview?id={ident}"
    req = urllib.request.Request(url, headers=NAVEGADOR)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            pagina = r.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        if e.code in (401, 403, 404):
            raise RuntimeError(
                "Pasta do Drive não encontrada ou não pública.\n"
                "Confira o link e compartilhe como \"Qualquer pessoa com o link\" (leitor)."
            ) from None
        raise

    itens = re.findall(
        r'id="entry-([\w-]+)".*?href="([^"]*)".*?class="flip-entry-title">([^<]*)<',
        pagina, flags=re.S,
    )
    if not itens and ("ServiceLogin" in pagina or "accounts.google.com" in pagina):
        raise RuntimeError(
            "A pasta do Drive não é pública.\n"
            "Compartilhe como \"Qualquer pessoa com o link\" (leitor)."
        )

    fotos = []
    for fid, href, nome in itens:
        nome = html.unescape(nome).strip()
        if "/folders/" in href:
            continue
        if Path(nome).suffix.lower() in EXTENSOES:
            fotos.append((fid, nome))
    return fotos


def sincronizar_drive(link, destino, aviso=print):
    """
    Baixa para `destino` as fotos da pasta do Drive que ainda não existem lá.
    Fotos que já estão na pasta (inclusive as já recortadas) nunca são trocadas.
    """
    destino = Path(destino)
    destino.mkdir(parents=True, exist_ok=True)
    fotos = listar_pasta_drive(link)
    existentes = {f.stem.casefold() for f in destino.iterdir() if f.is_file()}
    novas = [(fid, nome) for fid, nome in fotos if Path(nome).stem.casefold() not in existentes]

    aviso(f"Pasta do Drive: {len(fotos)} foto(s) | já na pasta local: {len(fotos) - len(novas)} | novas: {len(novas)}")
    baixadas = erros = 0
    for i, (fid, nome) in enumerate(novas, 1):
        url = f"https://drive.usercontent.google.com/download?id={fid}&export=download&confirm=t"
        try:
            _baixar(url, destino / nome)
            baixadas += 1
            aviso(f"  [{i}/{len(novas)}] {nome}")
        except Exception as e:
            erros += 1
            aviso(f"  ERRO {nome}: {e}")
    aviso(f"Fotos baixadas: {baixadas} | erros: {erros} | pasta local: {destino}")
    return baixadas, erros


# ============================================================
# NOMES
# ============================================================

def chave(texto):
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip().casefold()


RE_EXTENSAO_IMAGEM = re.compile(r"(?i)\.(?:jpe?g|png|tiff?|psd)$")
RE_TOKEN_TAMANHO = re.compile(
    r"(?i)[\(\[\{\s_\-\.]*"
    r"(?<!\d)(?:20\s*[x×*]\s*20|24\s*[x×*]\s*34|34\s*[x×*]\s*24|30\s*[x×*]\s*20|20\s*[x×*]\s*30|30\s*[x×*]\s*22)(?:\s*cm)?"
    r"(?:\s*\(\d+\))?"
    r"[\)\]\}\s_\-\.]*(?!\d)"
)


def extrair_tamanho_do_texto(texto):
    """Detecta se o texto/nome contém indicação explícita de tamanho (24X34, 20X20, 30X20, 30X22)."""
    compacto = re.sub(r"[\s_\-\.]+", "", str(texto or "").upper()).replace("×", "X").replace("*", "X")
    if "24X34" in compacto or "34X24" in compacto:
        return "24X34"
    if "20X20" in compacto:
        return "20X20"
    if "30X22" in compacto:
        return "30X22"
    if "30X20" in compacto or "20X30" in compacto:
        return "30X20"
    return None


def limpar_tamanho_do_nome(nome):
    """Remove qualquer sufixo/token de tamanho (ex.: 24X34, 34X24, 20X20, 30X20, 30X22) e extensão do nome do aluno."""
    texto = str(nome or "").strip()
    texto = RE_EXTENSAO_IMAGEM.sub("", texto)
    texto = RE_TOKEN_TAMANHO.sub(" ", texto)
    texto = re.sub(r"[_\-]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def tamanho_do_arquivo(caminho):
    """20X20 / 24X34 pelo sufixo; sem sufixo é a foto do 30X20."""
    stem = caminho.stem if hasattr(caminho, "stem") else Path(str(caminho)).stem
    achado = extrair_tamanho_do_texto(stem)
    if achado in ("20X20", "24X34", "30X20"):
        return achado
    return "30X20"


PLANO_POR_TAMANHO = {
    "30X20": "Prime Book 30X20",
    "20X20": "Essencial Book 20x20",
    "24X34": "Studio Day Berlim 24x34",
    "30X22": "Prime Book 30X22",
}


def formatar_nome(nome):
    """Só a primeira letra de cada palavra maiúscula e sem o tamanho (24X34 / 20X20 / 30X20): "MARIA DA SILVA 24X34" -> "Maria Da Silva"."""
    texto = limpar_tamanho_do_nome(nome)
    return re.sub(r"[^\W\d_]+", lambda m: m.group(0)[0].upper() + m.group(0)[1:].lower(), texto)


def nome_do_aluno(caminho):
    """Nome do aluno a partir do arquivo ou pasta, sempre sem o tamanho (20X20 / 24X34 / 30X20 / 34X24)."""
    stem = caminho.stem if hasattr(caminho, "stem") else Path(str(caminho)).stem
    return formatar_nome(stem)


def _cursos_salvos_por_foto():
    """Lê cursos_por_foto.json e devolve mapas por caminho normalizado e por chave do nome do aluno."""
    arquivo_cursos = PASTA / "cursos_por_foto.json"
    por_caminho = {}
    por_nome = {}
    try:
        dados = json.loads(arquivo_cursos.read_text(encoding="utf-8"))
        mapa = dados.get("cursos_por_foto", {}) if isinstance(dados, dict) else {}
        if isinstance(mapa, dict):
            for cam, cur in mapa.items():
                cur_str = str(cur or "").strip()
                if not cur_str:
                    continue
                norm = os.path.normcase(os.path.abspath(os.path.normpath(str(cam))))
                por_caminho[norm] = cur_str
                try:
                    nm = nome_do_aluno(Path(cam))
                    if nm:
                        por_nome[chave(nm)] = cur_str
                except Exception:
                    pass
    except Exception:
        pass
    return por_caminho, por_nome


def alunos_das_fotos(pasta):
    """
    Monta a lista de capas só com as fotos da pasta (sem planilha):
    foto com 24X34 -> Berlim; com 20X20 (sem 24X34) -> Essencial; sem tamanho -> Prime 30X20.
    Recupera também o curso já salvo na revisão (cursos_por_foto.json) e verifica o cache
    persistente (cache_analise_subpastas.json) para evitar reprocessar fotos já analisadas.
    Retorna [{"nome", "curso", "plano"}] em ordem alfabética.
    """
    verificar_cache_analise_subpastas(pasta)
    por_caminho, por_nome = _cursos_salvos_por_foto()
    por_aluno = {}
    for f in listar_fotos(pasta):
        nome = nome_do_aluno(f)
        if not nome:
            continue
        ch = chave(nome)
        reg = por_aluno.setdefault(ch, {"nome": nome, "curso": "", "tamanhos": set()})
        reg["tamanhos"].add(tamanho_do_arquivo(f))
        if not reg["curso"]:
            norm = os.path.normcase(os.path.abspath(os.path.normpath(str(f))))
            reg["curso"] = por_caminho.get(norm) or por_nome.get(ch) or ""

    alunos = []
    for reg in sorted(por_aluno.values(), key=lambda r: chave(r["nome"])):
        tamanhos = reg["tamanhos"]
        curso = reg.get("curso") or ""
        if "24X34" in tamanhos:
            alunos.append({"nome": reg["nome"], "curso": curso, "plano": PLANO_POR_TAMANHO["24X34"]})
        elif "20X20" in tamanhos:
            alunos.append({"nome": reg["nome"], "curso": curso, "plano": PLANO_POR_TAMANHO["20X20"]})
        if "30X20" in tamanhos:
            alunos.append({"nome": reg["nome"], "curso": curso, "plano": PLANO_POR_TAMANHO["30X20"]})
    return alunos


def _norm_caminho_rapido(caminho):
    """Normaliza o caminho em memória (0ms I/O, sem chamar GetFinalPathNameByHandleW do .resolve())."""
    return os.path.normcase(os.path.abspath(os.path.normpath(str(caminho))))


def listar_fotos(pasta):
    pasta_p = Path(pasta)
    if not pasta_p.is_dir():
        return []
    arquivos = []
    try:
        with os.scandir(pasta_p) as it:
            for entry in it:
                if entry.is_file() and not entry.name.startswith("._"):
                    suf = os.path.splitext(entry.name)[1].lower()
                    if suf in EXTENSOES:
                        arquivos.append(Path(entry.path))
    except OSError:
        return []
    arquivos.sort()
    return arquivos


# ============================================================
# ROSTO E CACHE DE ANÁLISE DE SUBPASTAS (cache_analise_subpastas.json)
# ============================================================

_detector = None
_haar = None
_tls_detector = threading.local()
_cache_rostos_lock = threading.Lock()
CACHE_ROSTOS_FILE = PASTA / "cache_rostos_yunet.json"
_cache_rostos_mem = None


def _carregar_cache_rostos():
    global _cache_rostos_mem
    with _cache_rostos_lock:
        if _cache_rostos_mem is not None:
            return _cache_rostos_mem
        if CACHE_ROSTOS_FILE.is_file():
            try:
                _cache_rostos_mem = json.loads(CACHE_ROSTOS_FILE.read_text(encoding="utf-8"))
                if isinstance(_cache_rostos_mem, dict):
                    return _cache_rostos_mem
            except Exception:
                pass
        _cache_rostos_mem = {}
        return _cache_rostos_mem


def _salvar_cache_rostos():
    with _cache_rostos_lock:
        if _cache_rostos_mem is None:
            return
        try:
            temp = CACHE_ROSTOS_FILE.with_suffix(".tmp")
            temp.write_text(json.dumps(_cache_rostos_mem, ensure_ascii=False), encoding="utf-8")
            os.replace(temp, CACHE_ROSTOS_FILE)
        except Exception:
            pass


CACHE_ANALISE_SUBPASTAS_FILE = PASTA / "cache_analise_subpastas.json"
_cache_analise_mem = None
_cache_analise_mtime_ns = None


def _carregar_cache_analise_subpastas():
    global _cache_analise_mem, _cache_analise_mtime_ns
    with _cache_rostos_lock:
        mtime_atual = None
        if CACHE_ANALISE_SUBPASTAS_FILE.is_file():
            try:
                mtime_atual = CACHE_ANALISE_SUBPASTAS_FILE.stat().st_mtime_ns
            except Exception:
                mtime_atual = None
        if _cache_analise_mem is not None and (mtime_atual is None or mtime_atual == _cache_analise_mtime_ns):
            return _cache_analise_mem
        if CACHE_ANALISE_SUBPASTAS_FILE.is_file():
            try:
                lido = json.loads(CACHE_ANALISE_SUBPASTAS_FILE.read_text(encoding="utf-8"))
                if isinstance(lido, dict):
                    lido.setdefault("_subpastas", {})
                    lido.setdefault("_destinos_copiados", {})
                    _cache_analise_mem = lido
                    _cache_analise_mtime_ns = mtime_atual
                    return _cache_analise_mem
            except Exception:
                pass
        _cache_analise_mem = {"_subpastas": {}, "_destinos_copiados": {}}
        _cache_analise_mtime_ns = mtime_atual
        return _cache_analise_mem


def _salvar_cache_analise_subpastas():
    global _cache_analise_mtime_ns
    with _cache_rostos_lock:
        if _cache_analise_mem is None:
            return
        try:
            temp = CACHE_ANALISE_SUBPASTAS_FILE.with_suffix(".tmp")
            temp.write_text(json.dumps(_cache_analise_mem, ensure_ascii=False), encoding="utf-8")
            os.replace(temp, CACHE_ANALISE_SUBPASTAS_FILE)
            try:
                _cache_analise_mtime_ns = CACHE_ANALISE_SUBPASTAS_FILE.stat().st_mtime_ns
            except Exception:
                pass
        except Exception:
            pass


def calcular_recursos_hardware(limite_cpu_pct=None, usar_gpu_opencl=None, total_tarefas=None):
    """
    Calcula o orçamento seguro de hardware (CPU + GPU OpenCL + Prioridade Baixa do SO)
    para acelerar a seleção das fotos usando o hardware da máquina sem usar 100% da CPU
    e sem travar o computador.
    """
    bruto = {}
    try:
        bruto = ler_config_bruto()
    except Exception:
        bruto = {}

    if limite_cpu_pct is None:
        limite_cpu_pct = bruto.get("hardware_limite_cpu_pct", LIMITE_CPU_PCT_PADRAO)
    try:
        pct = max(20, min(90, int(round(float(limite_cpu_pct)))))
    except Exception:
        pct = LIMITE_CPU_PCT_PADRAO

    if usar_gpu_opencl is None:
        usar_gpu_opencl = bruto.get("hardware_usar_gpu_opencl", USAR_GPU_OPENCL_PADRAO)
    usar_gpu = bool(usar_gpu_opencl)

    nucleos_totais = max(1, os.cpu_count() or 4)
    # Garante pelo menos 1 núcleo livre em máquinas multi-core para o sistema não engasgar
    max_permitido = max(1, nucleos_totais - 1) if nucleos_totais > 1 else 1
    nucleos_calculados = max(1, int(round(nucleos_totais * (pct / 100.0))))
    nucleos_usados = min(max_permitido, nucleos_calculados)
    if total_tarefas is not None and total_tarefas > 0:
        nucleos_usados = max(1, min(nucleos_usados, int(total_tarefas)))
    nucleos_livres = max(0, nucleos_totais - nucleos_usados)

    gpu_disp = False
    try:
        gpu_disp = bool(hasattr(cv2, "ocl") and cv2.ocl.haveOpenCL())
    except Exception:
        gpu_disp = False

    return {
        "nucleos_totais": nucleos_totais,
        "nucleos_usados": nucleos_usados,
        "nucleos_livres": nucleos_livres,
        "limite_cpu_pct": pct,
        "usar_gpu_opencl": usar_gpu,
        "gpu_disponivel": gpu_disp,
        "gpu_ativa": bool(usar_gpu and gpu_disp),
        "prioridade": "BELOW_NORMAL",
    }


def verificar_cache_analise_subpastas(
    pasta_fotos=None,
    processar_pendentes_pool=False,
    limite_cpu_pct=None,
    usar_gpu_opencl=None,
):
    """
    Verifica e pré-carrega `cache_analise_subpastas.json` ao iniciar o programa ou carregar fotos.
    Sincroniza todos os rostos já analisados (nas subpastas e nos arquivos copiados para FOTO)
    para a memória, evitando reprocessar fotos já analisadas e garantindo carregamento instantâneo.
    Se `processar_pendentes_pool=True` e houver fotos ainda fora do cache na pasta `pasta_fotos`,
    distribui a análise entre os núcleos permitidos pelo limite de hardware via `multiprocessing.Pool`.
    """
    cache_an = _carregar_cache_analise_subpastas()
    cache_r = _carregar_cache_rostos()
    hw = calcular_recursos_hardware(limite_cpu_pct=limite_cpu_pct, usar_gpu_opencl=usar_gpu_opencl)

    fotos_em_cache = 0
    with _cache_rostos_lock:
        for k, val in cache_an.items():
            if k.startswith("_") or not isinstance(val, dict):
                continue
            fotos_em_cache += 1
            if k.endswith("|v3") or k.endswith("|v2"):
                chave_rosto = k[:-3]
                if chave_rosto not in cache_r and "rostos" in val:
                    rostos_lista = val.get("rostos") or []
                    cache_r[chave_rosto] = rostos_lista[0] if rostos_lista else None

        destinos_map = cache_an.get("_destinos_copiados")
        if isinstance(destinos_map, dict):
            for norm_dest, info_dest in destinos_map.items():
                if not isinstance(info_dest, dict):
                    continue
                chave_dest = info_dest.get("chave_dest")
                if chave_dest and chave_dest not in cache_r and "rosto" in info_dest:
                    cache_r[chave_dest] = info_dest.get("rosto")

    subpastas_map = cache_an.get("_subpastas")
    subpastas_em_cache = len(subpastas_map) if isinstance(subpastas_map, dict) else 0

    fotos_atuais_em_cache = 0
    total_fotos_pasta = 0
    pendentes_pasta = []
    if pasta_fotos is not None:
        try:
            p = Path(pasta_fotos)
            if p.is_dir():
                with os.scandir(p) as it:
                    for entry in it:
                        if not entry.is_file() or entry.name.startswith("._"):
                            continue
                        if os.path.splitext(entry.name)[1].lower() not in EXTENSOES:
                            continue
                        total_fotos_pasta += 1
                        try:
                            st = entry.stat()
                            norm_p = _norm_caminho_rapido(entry.path)
                            ch_r = f"{norm_p}|{st.st_mtime_ns}|{st.st_size}"
                            if ch_r in cache_r or f"{ch_r}|v3" in cache_an:
                                fotos_atuais_em_cache += 1
                            elif processar_pendentes_pool:
                                pendentes_pasta.append((str(entry.path), st.st_mtime_ns, st.st_size, False))
                        except Exception:
                            pass
        except Exception:
            pass

    if processar_pendentes_pool and len(pendentes_pasta) >= 2:
        try:
            num_nucleos = min(len(pendentes_pasta), hw["nucleos_usados"])
            with concurrent.futures.ThreadPoolExecutor(max_workers=num_nucleos) as pool:
                for ch_an, ch_r, m_base in pool.map(_extrair_metricas_foto_worker, pendentes_pasta):
                    if m_base is not None:
                        with _cache_rostos_lock:
                            cache_an[ch_an] = m_base
                            rostos_l = m_base.get("rostos") or []
                            cache_r[ch_r] = rostos_l[0] if rostos_l else None
                        fotos_atuais_em_cache += 1
                        fotos_em_cache += 1
            _salvar_cache_analise_subpastas()
            _salvar_cache_rostos()
        except Exception:
            pass

    proxies_count = 0
    try:
        if PASTA_PROXIES_ANALISE.is_dir():
            for root_d, _, files_d in os.walk(PASTA_PROXIES_ANALISE):
                proxies_count += sum(1 for f in files_d if f.endswith(".jpg"))
    except Exception:
        proxies_count = 0

    return {
        "arquivo": str(CACHE_ANALISE_SUBPASTAS_FILE),
        "existe": CACHE_ANALISE_SUBPASTAS_FILE.is_file(),
        "fotos_em_cache": fotos_em_cache,
        "subpastas_em_cache": subpastas_em_cache,
        "fotos_pasta_atual_em_cache": fotos_atuais_em_cache,
        "total_fotos_pasta": total_fotos_pasta,
        "nucleos_cpu": hw["nucleos_totais"],
        "nucleos_usados": hw["nucleos_usados"],
        "nucleos_livres": hw["nucleos_livres"],
        "limite_cpu_pct": hw["limite_cpu_pct"],
        "gpu_ativa": hw["gpu_ativa"],
        "proxies_locais": proxies_count,
    }


def detectar_rosto_cache(caminho, img=None):
    """
    Busca o rosto no cache persistente (`cache_analise_subpastas.json` e `cache_rostos_yunet.json`)
    por (caminho, mtime_ns, size) antes de rodar o YuNet ONNX.
    Se precisar detectar do zero, usa o pipeline C++ IMREAD_REDUCED_COLOR_8 em 640px (~6ms).
    """
    caminho_p = Path(caminho)
    chave_cache = None
    norm_path = None
    try:
        st = caminho_p.stat()
        norm_path = _norm_caminho_rapido(caminho_p)
        chave_cache = f"{norm_path}|{st.st_mtime_ns}|{st.st_size}"
        cache = _carregar_cache_rostos()
        if chave_cache in cache:
            return cache[chave_cache]
        # Verifica também o cache_analise_subpastas.json para nunca reprocessar fotos já analisadas
        cache_an = _carregar_cache_analise_subpastas()
        for sufixo_v in ("|v3", "|v2"):
            chave_an = f"{chave_cache}{sufixo_v}"
            if chave_an in cache_an and isinstance(cache_an[chave_an], dict):
                rostos_an = cache_an[chave_an].get("rostos") or []
                rosto_cached = rostos_an[0] if rostos_an else None
                with _cache_rostos_lock:
                    cache[chave_cache] = rosto_cached
                return rosto_cached
        destinos_map = cache_an.get("_destinos_copiados") or {}
        info_dest = destinos_map.get(norm_path)
        if isinstance(info_dest, dict) and info_dest.get("chave_dest") == chave_cache:
            rosto_cached = info_dest.get("rosto")
            with _cache_rostos_lock:
                cache[chave_cache] = rosto_cached
            return rosto_cached
    except Exception:
        pass

    # Mesmo que `img` completo tenha sido passado, decodificar ou redimensionar para 640px é 15x mais rápido
    if caminho_p.is_file():
        bgr_red, tam_orig = _decodificar_bgr_para_analise(caminho_p, LADO_DETECCAO)
        rostos = detectar_rostos_bgr(bgr_red, tamanho_original=tam_orig, usar_haar_fallback=True)
        rosto = rostos[0] if rostos else None
    elif img is not None:
        rosto = detectar_rosto(img)
    else:
        rosto = None

    if chave_cache is not None:
        cache = _carregar_cache_rostos()
        with _cache_rostos_lock:
            cache[chave_cache] = rosto
        _salvar_cache_rostos()
    return rosto


_yunet_falhou_sessao = False


def detector_yunet():
    global _yunet_falhou_sessao
    if _yunet_falhou_sessao or not hasattr(cv2, "FaceDetectorYN"):
        return None
    det = getattr(_tls_detector, "detector", None)
    if det is None:
        try:
            with _cache_rostos_lock:
                if not MODELO_YUNET.is_file():
                    req = urllib.request.Request(URL_YUNET, headers=NAVEGADOR)
                    with urllib.request.urlopen(req, timeout=8) as r:
                        dados_onnx = r.read()
                    if dados_onnx:
                        temp_m = MODELO_YUNET.with_suffix(".tmp")
                        temp_m.write_bytes(dados_onnx)
                        os.replace(temp_m, MODELO_YUNET)
            if not MODELO_YUNET.is_file():
                _yunet_falhou_sessao = True
                return None
            det = cv2.FaceDetectorYN.create(
                str(MODELO_YUNET), "", (320, 320), 0.55, 0.3, 5000
            )
            _tls_detector.detector = det
        except Exception:
            _yunet_falhou_sessao = True
            return None
    return det


def _detector_haar_tls():
    """Instância Haar Cascade isolada por thread (sem lock global) para detecção rápida ou fallback."""
    haar = getattr(_tls_detector, "haar", None)
    if haar is None and hasattr(cv2, "CascadeClassifier"):
        try:
            haar = cv2.CascadeClassifier(
                cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
            )
            _tls_detector.haar = haar
        except Exception:
            haar = None
    return haar


def detectar_rosto(img):
    """
    Retorna o maior rosto em coordenadas da imagem original:
    {x, y, w, h, cx, olhos_y, metodo} ou None.
    """
    rostos = detectar_todos_rostos(img, usar_haar_fallback=True)
    return rostos[0] if rostos else None


OPCOES_FOTO_FILE = PASTA / "opcoes_por_foto.json"


def ler_opcoes_foto():
    if OPCOES_FOTO_FILE.is_file():
        try:
            dados = json.loads(OPCOES_FOTO_FILE.read_text(encoding="utf-8"))
            if isinstance(dados, dict):
                return dados
        except Exception:
            pass
    return {}


def gravar_opcoes_foto(dados):
    try:
        temp = OPCOES_FOTO_FILE.with_suffix(".tmp")
        temp.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")
        os.replace(temp, OPCOES_FOTO_FILE)
    except Exception:
        pass


def nome_arquivo_destino_foto(nome_aluno, tam):
    """Nome padrão na pasta FOTO compatível com AUTOMAÇAO-PASTA.py e recortar_fotos.py (nunca duplica o tamanho no nome)."""
    limpo = formatar_nome(nome_aluno)
    if str(tam).upper() in ("20X20", "24X34"):
        return f"{limpo} {str(tam).upper()}.jpg"
    return f"{limpo}.jpg"


def tamanhos_do_plano(plano, berlim_modo="Ambos"):
    """Devolve a lista de tamanhos de foto exigidos pelo plano (ex.: ['30X20'], ['20X20'], ['24X34', '20X20'])."""
    p = chave(plano or "")
    if any(k in p for k in ("berlim", "berlin", "24x34", "34x24", "diamond", "diamante")):
        modo = chave(berlim_modo or "ambos")
        if "24x34" in modo:
            return ["24X34"]
        if "20x20" in modo:
            return ["20X20"]
        return ["24X34", "20X20"]
    if "20x20" in p or "essencial" in p:
        return ["20X20"]
    return ["30X20"]


def _aplicar_orientacao_exif_bgr(bgr, orientacao):
    """Aplica a orientação EXIF (1..8) diretamente na matriz BGR em C++ (< 0.05 ms)."""
    if orientacao == 2:
        return cv2.flip(bgr, 1)
    if orientacao == 3:
        return cv2.rotate(bgr, cv2.ROTATE_180)
    if orientacao == 4:
        return cv2.flip(bgr, 0)
    if orientacao == 5:
        return cv2.flip(cv2.rotate(bgr, cv2.ROTATE_90_CLOCKWISE), 1)
    if orientacao == 6:
        return cv2.rotate(bgr, cv2.ROTATE_90_CLOCKWISE)
    if orientacao == 7:
        return cv2.flip(cv2.rotate(bgr, cv2.ROTATE_90_COUNTERCLOCKWISE), 1)
    if orientacao == 8:
        return cv2.rotate(bgr, cv2.ROTATE_90_COUNTERCLOCKWISE)
    return bgr


def _decodificar_bgr_para_analise(caminho, lado_max=LADO_DETECCAO):
    """
    Decodifica a foto diretamente para matriz BGR reduzida usando OpenCV C++ `libjpeg-turbo`
    com fator DCT 1/8 (`IMREAD_REDUCED_COLOR_8`), 1/4 ou 1/2 liberando o GIL do Python.
    Uma foto de 24MP–45MP (6000x4000) é decodificada em ~2.5 ms (até 70x mais rápido que PIL completo).
    Retorna (bgr_reduzido, (W_original, H_original)).
    """
    caminho_p = Path(caminho)
    # 1) Lê apenas o cabeçalho (< 0.1 ms) para saber dimensões brutas e orientação EXIF
    with Image.open(caminho_p) as im_hdr:
        w_raw, h_raw = im_hdr.size
        fmt = im_hdr.format
        try:
            orientacao = im_hdr.getexif().get(0x0112, 1)
        except Exception:
            orientacao = 1

    W_orig, H_orig = (h_raw, w_raw) if orientacao in (5, 6, 7, 8) else (w_raw, h_raw)
    maior_raw = max(w_raw, h_raw)

    # 2) Fast path C++ libjpeg-turbo para JPEGs (libera o GIL nas threads!)
    if fmt == "JPEG" or caminho_p.suffix.lower() in (".jpg", ".jpeg"):
        try:
            if maior_raw >= lado_max * 5.2:
                flag_dct = cv2.IMREAD_REDUCED_COLOR_8
            elif maior_raw >= lado_max * 2.6:
                flag_dct = cv2.IMREAD_REDUCED_COLOR_4
            elif maior_raw >= lado_max * 1.4:
                flag_dct = cv2.IMREAD_REDUCED_COLOR_2
            else:
                flag_dct = cv2.IMREAD_COLOR

            flag = flag_dct | getattr(cv2, "IMREAD_IGNORE_ORIENTATION", 0)
            buf = np.fromfile(str(caminho_p), dtype=np.uint8)
            bgr = cv2.imdecode(buf, flag)
            if bgr is not None and bgr.size > 0:
                if orientacao != 1:
                    bgr = _aplicar_orientacao_exif_bgr(bgr, orientacao)
                ph, pw = bgr.shape[:2]
                if max(pw, ph) > lado_max:
                    escala = lado_max / float(max(pw, ph))
                    bgr = cv2.resize(
                        bgr,
                        (max(1, round(pw * escala)), max(1, round(ph * escala))),
                        interpolation=cv2.INTER_AREA,
                    )
                return bgr, (W_orig, H_orig)
        except Exception:
            pass

    # Fallback PIL com draft no tamanho 1x exato (permite DCT 1/8 do libjpeg)
    with Image.open(caminho_p) as img:
        try:
            if img.format == "JPEG" and maior_raw > lado_max * 1.4:
                img.draft("RGB", (lado_max, lado_max))
        except Exception:
            pass
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        if max(img.size) > lado_max:
            escala = lado_max / float(max(img.size))
            img = img.resize(
                (max(1, round(img.width * escala)), max(1, round(img.height * escala))),
                Image.BILINEAR,
            )
        bgr = cv2.cvtColor(np.asarray(img), cv2.COLOR_RGB2BGR)
    return bgr, (W_orig, H_orig)


def _garantir_copia_leve_local(
    caminho,
    mtime_ns=None,
    fsize=None,
    lado_proxy=LADO_PROXY_LOCAL,
    qualidade=QUALIDADE_PROXY_LOCAL,
):
    """
    Cria ou reutiliza uma cópia leve local (proxy JPEG com diminuição de qualidade e resolução, ~45 KB)
    em `.cache_proxies_capaz` no disco rápido da máquina do usuário, preservando no nome do arquivo
    as dimensões reais `(W_orig, H_orig)` da foto original no Google Drive / rede / HD externo.
    Retorna `(bgr_reduzido, (W_orig, H_orig), caminho_proxy_str, criado_agora)`.
    """
    caminho_p = Path(caminho)
    if mtime_ns is None or fsize is None:
        st = caminho_p.stat()
        mtime_ns, fsize = st.st_mtime_ns, st.st_size

    norm_p = _norm_caminho_rapido(caminho_p)
    norm_parent = _norm_caminho_rapido(caminho_p.parent)
    sub_hash = hashlib.md5(norm_parent.encode("utf-8", errors="ignore")).hexdigest()[:12]
    arq_hash = hashlib.md5(f"{norm_p}|{mtime_ns}|{fsize}|{lado_proxy}".encode("utf-8", errors="ignore")).hexdigest()[:16]

    pasta_sub_proxy = PASTA_PROXIES_ANALISE / sub_hash
    # Verifica se já existe proxy salvo com dimensões codificadas no nome (<arq_hash>_<W>x<H>.jpg)
    if pasta_sub_proxy.is_dir():
        try:
            with os.scandir(pasta_sub_proxy) as it:
                for entry in it:
                    if entry.is_file() and entry.name.startswith(arq_hash + "_") and entry.name.endswith(".jpg"):
                        dims_txt = entry.name[len(arq_hash) + 1 : -4]
                        if "x" in dims_txt:
                            w_s, h_s = dims_txt.split("x", 1)
                            W_orig, H_orig = int(w_s), int(h_s)
                            buf_p = np.fromfile(entry.path, dtype=np.uint8)
                            bgr_p = cv2.imdecode(buf_p, cv2.IMREAD_COLOR)
                            if bgr_p is not None and bgr_p.size > 0:
                                return bgr_p, (W_orig, H_orig), entry.path, False
        except Exception:
            pass

    # Decodifica a foto de origem usando fast DCT 1/8 (lê e descomprime em baixa resolução)
    bgr_red, (W_orig, H_orig) = _decodificar_bgr_para_analise(caminho_p, lado_max=lado_proxy)
    caminho_proxy = pasta_sub_proxy / f"{arq_hash}_{W_orig}x{H_orig}.jpg"
    try:
        pasta_sub_proxy.mkdir(parents=True, exist_ok=True)
        ok_enc, buf_jpg = cv2.imencode(
            ".jpg",
            bgr_red,
            [int(cv2.IMWRITE_JPEG_QUALITY), int(qualidade)],
        )
        if ok_enc:
            caminho_proxy.write_bytes(buf_jpg.tobytes())
    except Exception:
        pass

    return bgr_red, (W_orig, H_orig), str(caminho_proxy), True


def abrir_foto_para_analise(caminho, lado_max=LADO_DETECCAO):
    """
    Compatibilidade: retorna (PIL.Image RGB reduzida, (W_orig, H_orig)) usando o decodificador C++ rápido.
    """
    bgr, tam_orig = _decodificar_bgr_para_analise(caminho, lado_max)
    rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
    return Image.fromarray(rgb), tam_orig


def abrir_foto_previa_rapida(caminho, lado_max=LADO_PREVIA_RAPIDA):
    """
    Abre uma foto para prévia na UI (aba Revisão / Editor) usando DCT rápido no tamanho `lado_max`.
    Se o decodificador OpenCV falhar (ex.: JPEG CMYK ou TIFF especial), faz fallback seguro para PIL.
    Retorna (img_reduzida_rgb, (W_orig, H_orig), fator_reducao).
    """
    try:
        bgr, (W_orig, H_orig) = _decodificar_bgr_para_analise(caminho, lado_max)
        rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
        img_red = Image.fromarray(rgb)
    except Exception:
        img_orig, _, _ = abrir_foto(caminho)
        W_orig, H_orig = img_orig.size
        fator_p = min(1.0, float(lado_max) / float(max(1, W_orig, H_orig)))
        if fator_p < 1.0:
            img_red = img_orig.resize(
                (max(1, round(W_orig * fator_p)), max(1, round(H_orig * fator_p))), Image.BILINEAR
            )
        else:
            img_red = img_orig
    fator = img_red.width / float(max(1, W_orig))
    return img_red, (W_orig, H_orig), fator


def detectar_rostos_bgr(bgr, tamanho_original=None, usar_haar_fallback=True):
    """
    Detecta todos os rostos diretamente em uma matriz BGR usando YuNet ONNX em ~6ms (sem GIL).
    Se o YuNet não estiver disponível ou não encontrar rosto, faz fallback seguro para Haar Cascade.
    """
    if bgr is None or bgr.size == 0:
        return []
    ph, pw = bgr.shape[:2]
    W_orig, H_orig = tamanho_original if tamanho_original else (pw, ph)

    if max(pw, ph) > LADO_DETECCAO:
        escala_det = LADO_DETECCAO / float(max(pw, ph))
        pw_det = max(1, round(pw * escala_det))
        ph_det = max(1, round(ph * escala_det))
        bgr_det = cv2.resize(bgr, (pw_det, ph_det), interpolation=cv2.INTER_AREA)
    else:
        bgr_det = bgr
        pw_det, ph_det = pw, ph

    escala_x = pw_det / float(max(1, W_orig))
    escala_y = ph_det / float(max(1, H_orig))

    det = detector_yunet()
    faces = None
    if det is not None:
        try:
            det.setInputSize((pw_det, ph_det))
            _, faces = det.detect(bgr_det)
        except Exception:
            faces = None

    resultado = []
    if faces is not None and len(faces):
        for f in faces:
            score = float(f[14]) if len(f) > 14 else 0.85
            if score < 0.50:
                continue
            x = float(f[0]) / escala_x
            y = float(f[1]) / escala_y
            fw = float(f[2]) / escala_x
            fh = float(f[3]) / escala_y
            olho_d = (float(f[4]) / escala_x, float(f[5]) / escala_y)
            olho_e = (float(f[6]) / escala_x, float(f[7]) / escala_y)
            nariz = (float(f[8]) / escala_x, float(f[9]) / escala_y)
            boca_d = (float(f[10]) / escala_x, float(f[11]) / escala_y)
            boca_e = (float(f[12]) / escala_x, float(f[13]) / escala_y)
            resultado.append({
                "x": x, "y": y, "w": fw, "h": fh,
                "cx": x + fw / 2.0,
                "olhos_y": (olho_d[1] + olho_e[1]) / 2.0,
                "olho_d": olho_d,
                "olho_e": olho_e,
                "nariz": nariz,
                "boca_d": boca_d,
                "boca_e": boca_e,
                "score": score,
                "metodo": "yunet",
            })
        resultado.sort(key=lambda r: r["w"] * r["h"], reverse=True)
        # Garante que o rosto principal tenha confiança sólida (>= 0.54), mantendo rostos secundários (>= 0.50) para detectar fotos em dupla/grupo
        if resultado and resultado[0].get("score", 0.0) >= 0.54:
            return resultado
        resultado = [r for r in resultado if r.get("score", 0.0) >= 0.54]
        if resultado:
            return resultado

    if not usar_haar_fallback and det is not None:
        return []

    haar = _detector_haar_tls()
    if haar is None:
        return []
    try:
        cinza = cv2.cvtColor(bgr_det, cv2.COLOR_BGR2GRAY)
        caixas = haar.detectMultiScale(cinza, 1.15, 5, minSize=(28, 28))
    except Exception:
        caixas = ()
    if len(caixas):
        rostos_haar = []
        for r_box in caixas:
            x = float(r_box[0]) / escala_x
            y = float(r_box[1]) / escala_y
            fw = float(r_box[2]) / escala_x
            fh = (float(r_box[3]) / escala_y) * 0.9
            rostos_haar.append({
                "x": x, "y": y, "w": fw, "h": fh,
                "cx": x + fw / 2.0,
                "olhos_y": y + fh * 0.4,
                "score": 0.65,
                "metodo": "haar",
            })
        rostos_haar.sort(key=lambda r: r["w"] * r["h"], reverse=True)
        return rostos_haar
    return []


def detectar_todos_rostos(img, tamanho_original=None, usar_haar_fallback=False):
    """
    Detecta todos os rostos na imagem usando YuNet ONNX.
    Aceita tanto PIL.Image quanto matriz numpy BGR.
    """
    if isinstance(img, np.ndarray):
        return detectar_rostos_bgr(img, tamanho_original=tamanho_original, usar_haar_fallback=usar_haar_fallback)
    w, h = img.size
    W_orig, H_orig = tamanho_original if tamanho_original else (w, h)
    escala_det = min(1.0, LADO_DETECCAO / max(w, h))
    pequena = img.convert("RGB")
    if escala_det < 1.0:
        pequena = pequena.resize((max(1, round(w * escala_det)), max(1, round(h * escala_det))), Image.BILINEAR)
    bgr = cv2.cvtColor(np.asarray(pequena), cv2.COLOR_RGB2BGR)
    return detectar_rostos_bgr(bgr, tamanho_original=(W_orig, H_orig), usar_haar_fallback=usar_haar_fallback)


def _analisar_histograma_e_contraste(arr_rosto):
    """
    Analisa o histograma de luminância (256 níveis) da região do rosto em O(1) usando soma acumulada:
      1. Detecta 'estouro de luz' (picos nas altas luzes / clipping nos bins 242..255).
      2. Calcula pontuação de contraste (0 a 100) combinando desvio padrão RMS e faixa dinâmica P95 - P5.
    """
    if arr_rosto is None or arr_rosto.size == 0:
        return {
            "brilho_medio": 128.0,
            "contraste_score": 0.0,
            "baixo_contraste": True,
            "estouro_luz": False,
            "altas_luzes_pct": 0.0,
            "pico_clip_pct": 0.0,
            "sombra_pct": 0.0,
            "pts_luz": 0.0,
        }

    flat = arr_rosto.ravel()
    total = float(max(1, flat.size))
    hist = np.bincount(flat, minlength=256)
    cdf = np.cumsum(hist)

    brilho_medio = float(flat.mean())
    std_luma = float(flat.std())
    p5 = float(np.searchsorted(cdf, total * 0.05))
    p95 = float(np.searchsorted(cdf, total * 0.95))
    amplitude_dinamica = float(p95 - p5)

    # Pontuação de contraste (0 a 100): combina RMS contrast (meta >= 42) e faixa P95-P5 (meta >= 135)
    score_std = min(1.0, std_luma / 44.0)
    score_amp = min(1.0, amplitude_dinamica / 140.0)
    contraste_score = round((0.55 * score_std + 0.45 * score_amp) * 100.0, 1)
    baixo_contraste = bool(std_luma < 26.0 or amplitude_dinamica < 75.0 or contraste_score < 35.0)

    # Verificação de estouro de luz via histograma (altas luzes 242..255 e pico saturado 250..255)
    altas_luzes_pct = round(float(hist[242:256].sum()) * 100.0 / total, 2)
    pico_clip_pct = round(float(hist[250:256].sum()) * 100.0 / total, 2)
    sombra_pct = round(float(hist[0:28].sum()) * 100.0 / total, 2)

    # Identifica pico anômalo nas altas luzes do histograma (concentração abrupta no topo)
    media_meios_tons = float(hist[100:210].mean()) if hist[100:210].size else 1.0
    pico_topo = float(hist[245:256].max()) if hist[245:256].size else 0.0
    tem_pico_estouro = bool(
        pico_clip_pct > 4.2
        or altas_luzes_pct > 12.5
        or (altas_luzes_pct > 7.5 and pico_topo > max(total * 0.016, media_meios_tons * 3.8))
    )

    penalidade_luz = 0.0
    if brilho_medio < 92.0:
        penalidade_luz += (92.0 - brilho_medio) / 28.0
    elif brilho_medio > 204.0:
        penalidade_luz += (brilho_medio - 204.0) / 22.0

    if sombra_pct > 8.0:
        penalidade_luz += (sombra_pct - 8.0) * 0.14
    if altas_luzes_pct > 4.0:
        penalidade_luz += (altas_luzes_pct - 4.0) * 0.22
    if tem_pico_estouro:
        penalidade_luz += 3.5
    if baixo_contraste:
        penalidade_luz += 2.5

    pts_luz = round(max(0.0, 7.0 - penalidade_luz), 1)

    return {
        "brilho_medio": round(brilho_medio, 1),
        "contraste_score": contraste_score,
        "baixo_contraste": baixo_contraste,
        "estouro_luz": tem_pico_estouro,
        "altas_luzes_pct": altas_luzes_pct,
        "pico_clip_pct": pico_clip_pct,
        "sombra_pct": sombra_pct,
        "pts_luz": pts_luz,
    }


def _extrair_metricas_base_foto(caminho, mtime_ns=None, fsize=None, usar_copia_leve=True):
    """
    Extrai uma única vez por foto (com cache persistente em disco e pipeline 100% C++ OpenCV/NumPy)
    as métricas independentes de molde: dimensões originais, rostos detectados, nitidez Laplaciano,
    simetria frontal, histograma de altas luzes e pontuação de contraste (~6 ms por foto).
    Quando `usar_copia_leve=True`, gera/usa primeiro uma cópia leve local (`.cache_proxies_capaz`)
    com diminuição de qualidade no disco rápido do PC, mantendo a foto original só para a escolhida.
    """
    caminho_p = Path(caminho)
    chave_cache = None
    chave_rosto_cache = None
    try:
        if mtime_ns is None or fsize is None:
            st = caminho_p.stat()
            mtime_ns, fsize = st.st_mtime_ns, st.st_size
        norm_path = _norm_caminho_rapido(caminho_p)
        chave_rosto_cache = f"{norm_path}|{mtime_ns}|{fsize}"
        chave_cache = f"{chave_rosto_cache}|v3"
        cache_an = _carregar_cache_analise_subpastas()
        if chave_cache in cache_an:
            return cache_an[chave_cache]
    except Exception:
        pass

    caminho_proxy_str = ""
    try:
        if usar_copia_leve:
            bgr_red, (W, H), caminho_proxy_str, _ = _garantir_copia_leve_local(
                caminho_p, mtime_ns=mtime_ns, fsize=fsize
            )
        else:
            bgr_red, (W, H) = _decodificar_bgr_para_analise(caminho_p, LADO_DETECCAO)
        rostos = detectar_rostos_bgr(bgr_red, tamanho_original=(W, H), usar_haar_fallback=True)
    except Exception:
        try:
            W, H = tamanho_rapido(caminho_p)
        except Exception:
            W, H = 3000, 4000
        bgr_red = None
        rostos = []

    # Aproveita para alimentar também o cache de rosto da aba Revisão
    if chave_rosto_cache is not None:
        cache_r = _carregar_cache_rostos()
        with _cache_rostos_lock:
            cache_r[chave_rosto_cache] = rostos[0] if rostos else None

    if not rostos:
        dados_sem_rosto = {
            "W": W,
            "H": H,
            "rostos": [],
            "qtd_pessoas": 0,
            "var_lap": 0.0,
            "embacada": True,
            "pts_nitidez": 0.0,
            "pts_frente": 0.0,
            "rosto_de_frente": False,
            "hist_luz": _analisar_histograma_e_contraste(None),
            "proxy_local": caminho_proxy_str,
        }
        if chave_cache is not None:
            cache_an = _carregar_cache_analise_subpastas()
            with _cache_rostos_lock:
                cache_an[chave_cache] = dados_sem_rosto
        return dados_sem_rosto

    rosto = rostos[0]
    area_principal_base = max(1.0, rosto["w"] * rosto["h"])
    rostos_relevantes_base = [
        r for r in rostos if (r["w"] * r["h"]) >= area_principal_base * 0.12 and float(r.get("score", 0.8)) >= 0.50
    ]
    qtd_pessoas_base = max(1, len(rostos_relevantes_base))

    ph, pw = bgr_red.shape[:2]
    escala_x = pw / float(max(1, W))
    escala_y = ph / float(max(1, H))
    fx0 = int(max(0, round(rosto["x"] * escala_x)))
    fy0 = int(max(0, round(rosto["y"] * escala_y)))
    fx1 = int(min(pw, max(fx0 + 12, round((rosto["x"] + rosto["w"]) * escala_x))))
    fy1 = int(min(ph, max(fy0 + 12, round((rosto["y"] + rosto["h"]) * escala_y))))

    recorte_bgr = bgr_red[fy0:fy1, fx0:fx1]
    recorte_cinza = cv2.cvtColor(recorte_bgr, cv2.COLOR_BGR2GRAY)
    # Janela de 160x160 px (2.7x mais pixels na face) para cálculo de nitidez e contraste com alta precisão
    arr_rosto = cv2.resize(recorte_cinza, (160, 160), interpolation=cv2.INTER_LINEAR)

    # Nitidez rápida no rosto (Variância do Laplaciano em 160x160)
    var_lap = float(cv2.Laplacian(arr_rosto, cv2.CV_64F).var())
    embacada = var_lap < 32.0
    if var_lap >= 95.0:
        pts_nitidez = 6.0
    elif var_lap >= 32.0:
        pts_nitidez = round(2.5 + 3.5 * ((var_lap - 32.0) / (95.0 - 32.0)), 1)
    else:
        pts_nitidez = round(max(0.0, 2.0 * (var_lap / 32.0)), 1)

    # Rosto de frente olhando para a câmera (leve)
    pts_frente = 3.0
    rosto_de_frente = True
    if "olho_d" in rosto and "olho_e" in rosto and "nariz" in rosto:
        od, oe, nar = rosto["olho_d"], rosto["olho_e"], rosto["nariz"]
        dist_olhos = max(1.0, math.hypot(oe[0] - od[0], oe[1] - od[1]))
        meio_x = (od[0] + oe[0]) / 2.0
        assimetria_yaw = abs(nar[0] - meio_x) / dist_olhos
        inclinacao_roll = abs(oe[1] - od[1]) / dist_olhos
        fator_frente = max(
            0.0,
            1.0 - max(0.0, assimetria_yaw - 0.04) * 2.5 - max(0.0, inclinacao_roll - 0.04) * 2.0,
        )
        pts_frente = round(3.0 * min(1.0, fator_frente), 1)
        rosto_de_frente = fator_frente >= 0.65

    # Histograma rápido (estouro de luz) e Contraste
    hist_luz = _analisar_histograma_e_contraste(arr_rosto)

    dados_base = {
        "W": W,
        "H": H,
        "rostos": rostos,
        "qtd_pessoas": qtd_pessoas_base,
        "var_lap": round(var_lap, 1),
        "embacada": embacada,
        "pts_nitidez": pts_nitidez,
        "pts_frente": pts_frente,
        "rosto_de_frente": rosto_de_frente,
        "hist_luz": hist_luz,
        "proxy_local": caminho_proxy_str,
    }
    if chave_cache is not None:
        cache_an = _carregar_cache_analise_subpastas()
        with _cache_rostos_lock:
            cache_an[chave_cache] = dados_base
    return dados_base


def _orientacao_ideal_para_molde(W, H, tam, molde=None):
    """
    Verifica instantaneamente se as dimensões (W, H) estão na orientação ideal para o molde `tam`:
      - Para 24X34 e 30X20: prioriza fotos Verticais / Retrato (H >= W * 0.95).
      - Para 20X20: prioriza fotos Verticais ou Quadradas/Proporcionais (H >= W * 0.95 ou 0.78 <= W/H <= 1.32).
    """
    W_f, H_f = float(max(1, W)), float(max(1, H))
    if str(tam).upper() in ("24X34", "30X20"):
        return H_f >= W_f * 0.95
    return (H_f >= W_f * 0.95) or (0.78 <= (W_f / H_f) <= 1.32)


def _separar_por_orientacao_rapido(
    fotos_para_analisar,
    tamanhos_necessarios,
    moldes=None,
    max_orientacao_ideal=MAX_TRIAGEM_ORIENTACAO_CABECALHO,
):
    """
    Triagem ultra-rápida por orientação (< 0.08 ms por foto, lendo SOMENTE o cabeçalho EXIF/JPEG via
    `tamanho_rapido` sem decodificar pixels):
      1. Separa imediatamente as fotos que já estão na orientação ideal dos moldes (`grupo_ideal`)
         das fotos em outra orientação (`grupo_reserva`).
      2. Assim que reúne `max_orientacao_ideal` fotos na orientação certa (priorizando arquivos maiores/melhores),
         encerra a varredura de cabeçalhos em pastas gigantes para não perder tempo avaliando critérios em dezenas de fotos.
    Retorna `(grupo_ideal, grupo_reserva)`.
    """
    if not fotos_para_analisar:
        return [], []

    # Ordena primeiro pelas fotos de maior tamanho em bytes (retratos originais de estúdio têm mais dados que miniaturas)
    candidatas_ordenadas = sorted(fotos_para_analisar, key=lambda item: item[2], reverse=True)
    grupo_ideal = []
    grupo_reserva = []
    tams = list(tamanhos_necessarios or ["30X20"])

    for item in candidatas_ordenadas:
        arq = item[0]
        try:
            W_hdr, H_hdr = tamanho_rapido(arq)
            eh_ideal = any(_orientacao_ideal_para_molde(W_hdr, H_hdr, t) for t in tams)
        except Exception:
            eh_ideal = True

        if eh_ideal:
            grupo_ideal.append(item)
            if len(grupo_ideal) >= max_orientacao_ideal:
                break
        else:
            if len(grupo_reserva) < max(4, max_orientacao_ideal // 2):
                grupo_reserva.append(item)

    return grupo_ideal, grupo_reserva


def avaliar_foto_para_molde(caminho, tam, molde, calib, ampliacao_maxima=1.2, metricas_base=None):
    """
    Avalia uma foto candidata para um tamanho de capa ('20X20', '30X20', '24X34')
    identificando PRIMEIRO OS ROSTOS e bloqueando fotos com mais de 1 pessoa:
      1. Rosto detectado + Estritamente 1 pessoa na foto (`qtd_pessoas == 1`): critério #1 eliminatório.
         Fotos com 2 ou mais pessoas recebem bloqueio (`aprovada_individual = False`) e nunca são escolhidas
         quando há foto individual na subpasta.
      2. Orientação ideal + Encaixe pós-corte no molde.
      3. Densidade / Proporção do rosto no corte (em alta precisão de pixels).
      4. Nitidez, olhar de frente, brilho e contraste.
    """
    caminho_p = Path(caminho)
    base = metricas_base if metricas_base is not None else _extrair_metricas_base_foto(caminho_p)
    W, H = int(base["W"]), int(base["H"])
    mw, mh = molde
    orientacao_foto = "Vertical" if H > W * 1.05 else ("Horizontal" if W > H * 1.05 else "Quadrada")
    orientacao_ideal = _orientacao_ideal_para_molde(W, H, tam, molde)

    rostos = base.get("rostos") or []
    if not rostos:
        caixa_fallback = ajustar_caixa((0, 0, W, H), (W, H), molde)
        return {
            "caminho": str(caminho_p),
            "nome_arquivo": caminho_p.name,
            "tamanho": tam,
            "valida": False,
            "aprovada_individual": False,
            "nota": 5.0 if orientacao_ideal else 1.0,
            "qtd_pessoas": 0,
            "orientacao": orientacao_foto,
            "orientacao_ideal": orientacao_ideal,
            "encaixe_pct": 0,
            "nitidez": 0.0,
            "nitidez_score": 0,
            "embacada": True,
            "contraste_score": 0.0,
            "brilho_medio": 128.0,
            "brilho_score": 50,
            "baixo_contraste": True,
            "estouro_luz": False,
            "altas_luzes_pct": 0.0,
            "densidade_rosto_px": 0,
            "densidade_rosto_pct": 0,
            "aprovada_qualidade": False,
            "aviso_qualidade": "Sem rosto detectado",
            "ampliacao": round(ampliacao_da_caixa(caixa_fallback, molde), 2),
            "caixa": list(caixa_fallback),
            "resumo": "Sem rosto detectado",
        }

    rosto = rostos[0]
    area_principal = max(1.0, rosto["w"] * rosto["h"])
    # Sensibilidade reforçada (>= 12% da área do rosto principal) para identificar qualquer 2ª pessoa na foto
    rostos_relevantes = [
        r for r in rostos if (r["w"] * r["h"]) >= area_principal * 0.12 and float(r.get("score", 0.8)) >= 0.50
    ]
    qtd_pessoas = max(1, len(rostos_relevantes))
    aprovada_individual = (qtd_pessoas == 1)

    # Caixa ideal vs Caixa real pós-corte
    calib_uso = calib or CALIBRACAO_PADRAO
    altura_rosto_calib = max(0.05, float(calib_uso.get("altura_rosto", 0.16)))
    crop_h_ideal = rosto["h"] / altura_rosto_calib

    caixa = calcular_recorte((W, H), rosto, molde, calib_uso)
    cl, ct, cr, cb = caixa
    cw, ch = max(1.0, float(cr - cl)), max(1.0, float(cb - ct))

    # 1. Critério #1 Absoluto: Rosto Individual (1 única pessoa) vs Múltiplas Pessoas
    if aprovada_individual:
        pts_sozinha = 30.0
        tag_pessoas = "1 Rosto (Individual)"
    else:
        outros_no_corte = 0
        for r_sec in rostos_relevantes[1:]:
            if (r_sec["cx"] >= cl and r_sec["cx"] <= cr and r_sec["olhos_y"] >= ct and r_sec["olhos_y"] <= cb):
                outros_no_corte += 1
        pts_sozinha = -45.0 - (10.0 * outros_no_corte)
        tag_pessoas = f"{qtd_pessoas} pessoas (Bloqueada p/ capa individual)"

    # 2. Orientação Ideal + Encaixe Pós-Corte (até 40 pts: 20 orientação + 20 encaixe)
    pts_orientacao = 20.0 if orientacao_ideal else 0.0

    escala_preservada = min(1.0, ch / max(1.0, crop_h_ideal))
    olhos_y_pos_corte = (rosto["olhos_y"] - ct) / ch
    centro_x_pos_corte = (rosto["cx"] - cl) / cw
    erro_olhos = abs(olhos_y_pos_corte - calib_uso.get("olhos_y", 0.38))
    erro_centro = abs(centro_x_pos_corte - calib_uso.get("centro_x", 0.5))

    topo_cabeca_y = rosto["y"] - 0.32 * rosto["h"]
    base_ombros_y = rosto["y"] + 1.50 * rosto["h"]
    folga_cabeca_ok = 1.0 if topo_cabeca_y >= ct else max(0.0, 1.0 - (ct - topo_cabeca_y) / max(1.0, rosto["h"]))
    folga_ombros_ok = 1.0 if base_ombros_y <= cb else max(0.0, 1.0 - (base_ombros_y - cb) / max(1.0, rosto["h"]))

    alinhamento_olhos = max(0.0, 1.0 - erro_olhos * 3.2)
    alinhamento_centro = max(0.0, 1.0 - erro_centro * 2.5)
    encaixe_ratio = (
        0.40 * escala_preservada
        + 0.28 * alinhamento_olhos
        + 0.14 * alinhamento_centro
        + 0.10 * folga_cabeca_ok
        + 0.08 * folga_ombros_ok
    )
    encaixe_pct = int(round(max(0.0, min(1.0, encaixe_ratio)) * 100))
    pts_encaixe = round(20.0 * max(0.0, min(1.0, encaixe_ratio)), 1)

    # 3. Densidade de pixels no rosto após o crop simulado (até 18 pts)
    rx0 = max(float(cl), float(rosto["x"]))
    ry0 = max(float(ct), float(rosto["y"]))
    rx1 = min(float(cr), float(rosto["x"] + rosto["w"]))
    ry1 = min(float(cb), float(rosto["y"] + rosto["h"]))
    rosto_w_no_corte = max(0.0, rx1 - rx0)
    rosto_h_no_corte = max(0.0, ry1 - ry0)
    densidade_rosto_px = int(round(rosto_h_no_corte))
    rosto_area_px = int(round(rosto_w_no_corte * rosto_h_no_corte))

    alvo_rosto_h_px = max(220.0, mh * altura_rosto_calib)
    densidade_ratio = rosto_h_no_corte / alvo_rosto_h_px
    densidade_rosto_pct = int(round(max(0.0, min(2.0, densidade_ratio)) * 100))

    amp = ampliacao_da_caixa(caixa, molde)
    if densidade_ratio >= 1.0 and amp <= 1.0:
        bonus_reserva = min(2.0, (min(1.5, densidade_ratio) - 1.0) * 4.0)
        pts_densidade = round(16.0 + bonus_reserva, 1)
    elif amp <= ampliacao_maxima:
        fator_amp = max(0.0, 1.0 - ((amp - 1.0) / max(0.05, ampliacao_maxima - 1.0)))
        pts_densidade = round(10.0 + 6.0 * min(1.0, densidade_ratio) * (0.5 + 0.5 * fator_amp), 1)
    else:
        pts_densidade = round(max(0.0, 8.0 * min(1.0, densidade_ratio) - 10.0 * (amp - ampliacao_maxima)), 1)

    # 4. Nitidez / Luz em alta resolução (até 12 pts)
    var_lap = float(base.get("var_lap", 0.0))
    embacada = bool(base.get("embacada", False))
    pts_nitidez = min(6.0, float(base.get("pts_nitidez", 4.0)))
    pts_frente = min(3.0, float(base.get("pts_frente", 2.5)))
    hist_luz = base.get("hist_luz") or {}
    contraste_score = float(hist_luz.get("contraste_score", 50.0))
    brilho_medio = float(hist_luz.get("brilho_medio", 132.0))
    desvio_brilho = abs(brilho_medio - 134.0)
    brilho_score = int(round(max(15.0, min(100.0, 100.0 - desvio_brilho * 0.75))))
    nitidez_score = int(round(max(10.0, min(100.0, (var_lap / 150.0) * 100.0))))
    baixo_contraste = bool(hist_luz.get("baixo_contraste", False))
    estouro_luz = bool(hist_luz.get("estouro_luz", False))
    altas_luzes_pct = float(hist_luz.get("altas_luzes_pct", 0.0))
    pts_luz = min(3.0, float(hist_luz.get("pts_luz", 5.0)) * (3.0 / 7.0))

    penalidade_descarte = 0.0
    avisos_qualidade = []
    if not aprovada_individual:
        penalidade_descarte += 35.0
        avisos_qualidade.append(f"{qtd_pessoas} pessoas detectadas")
    if estouro_luz:
        penalidade_descarte += 8.0
        avisos_qualidade.append(f"Estouro de luz ({altas_luzes_pct:.1f}%)")
    if baixo_contraste:
        penalidade_descarte += 6.0
        avisos_qualidade.append(f"Baixo contraste ({contraste_score:.0f}/100)")
    if embacada:
        penalidade_descarte += 8.0
        avisos_qualidade.append(f"Foto embaçada (nitidez {var_lap:.0f})")

    aprovada_qualidade = aprovada_individual and (not estouro_luz) and (not baixo_contraste) and (not embacada)
    aviso_qualidade = " · ".join(avisos_qualidade)

    nota_bruta = (
        pts_sozinha
        + pts_orientacao
        + pts_encaixe
        + pts_densidade
        + pts_nitidez
        + pts_frente
        + pts_luz
        - penalidade_descarte
    )
    if not aprovada_individual:
        nota_total = round(max(1.0, min(25.0, nota_bruta)), 1)
    else:
        nota_total = round(max(1.0, min(100.0, nota_bruta)), 1)

    partes_resumo = [
        tag_pessoas,
        orientacao_foto + (" (Ideal)" if orientacao_ideal else " (Adaptada)"),
        f"Encaixe {encaixe_pct}%",
        f"Rosto {densidade_rosto_px}px",
    ]
    return {
        "caminho": str(caminho_p),
        "nome_arquivo": caminho_p.name,
        "tamanho": tam,
        "valida": True,
        "aprovada_individual": aprovada_individual,
        "nota": nota_total,
        "qtd_pessoas": qtd_pessoas,
        "orientacao": orientacao_foto,
        "orientacao_ideal": orientacao_ideal,
        "encaixe_pct": encaixe_pct,
        "nitidez": round(var_lap, 1),
        "nitidez_score": nitidez_score,
        "embacada": embacada,
        "contraste_score": contraste_score,
        "brilho_medio": round(brilho_medio, 1),
        "brilho_score": brilho_score,
        "baixo_contraste": baixo_contraste,
        "estouro_luz": estouro_luz,
        "altas_luzes_pct": altas_luzes_pct,
        "densidade_rosto_px": densidade_rosto_px,
        "densidade_rosto_pct": densidade_rosto_pct,
        "rosto_area_px": rosto_area_px,
        "aprovada_qualidade": aprovada_qualidade,
        "aviso_qualidade": aviso_qualidade,
        "ampliacao": round(amp, 2),
        "caixa": list(caixa),
        "resumo": " · ".join(partes_resumo),
    }


def _copiar_foto_campea_rapido(origem, destino, rosto_detectado=None):
    """
    Copia a foto campeã para a pasta FOTO sem recomprimir o JPEG quando ela já é .jpg/.jpeg
    com orientação EXIF normal (orientacao == 1), e pula a cópia em disco (0ms) caso o arquivo
    de destino já corresponda exatamente à mesma foto campeã registrada em `cache_analise_subpastas.json`.
    """
    origem_p = Path(origem)
    destino_p = Path(destino)
    norm_orig = ""
    norm_dest = ""
    chave_orig = ""
    try:
        st_orig = origem_p.stat()
        norm_orig = _norm_caminho_rapido(origem_p)
        chave_orig = f"{norm_orig}|{st_orig.st_mtime_ns}|{st_orig.st_size}"
        norm_dest = _norm_caminho_rapido(destino_p)
    except Exception:
        pass

    cache_an = _carregar_cache_analise_subpastas()
    destinos_map = cache_an.setdefault("_destinos_copiados", {})

    # Verifica se o arquivo de destino na pasta FOTO já foi copiado desta mesma foto campeã (0ms I/O)
    ja_copiado_identico = False
    if chave_orig and norm_dest and destino_p.is_file():
        try:
            st_dest_atual = destino_p.stat()
            chave_dest_atual = f"{norm_dest}|{st_dest_atual.st_mtime_ns}|{st_dest_atual.st_size}"
            reg_dest = destinos_map.get(norm_dest)
            if (
                isinstance(reg_dest, dict)
                and reg_dest.get("chave_orig") == chave_orig
                and reg_dest.get("chave_dest") == chave_dest_atual
            ):
                ja_copiado_identico = True
                if rosto_detectado is None:
                    rosto_detectado = reg_dest.get("rosto")
        except Exception:
            ja_copiado_identico = False

    if not ja_copiado_identico:
        copiou_direto = False
        if origem_p.suffix.lower() in (".jpg", ".jpeg"):
            try:
                with Image.open(origem_p) as im_check:
                    orientacao = im_check.getexif().get(0x0112, 1)
                    modo = im_check.mode
                if orientacao == 1 and modo == "RGB":
                    shutil.copy2(origem_p, destino_p)
                    copiou_direto = True
            except Exception:
                copiou_direto = False

        if not copiou_direto:
            img_campea, icc_c, dpi_c = abrir_foto(origem_p)
            salvar_jpg(img_campea.convert("RGB"), destino_p, icc_c, dpi_c or DPI_MOLDE)

    # Atualiza o cache de destino e o cache de rosto para o arquivo na pasta FOTO (evita rodar YuNet na Revisão)
    try:
        st = destino_p.stat()
        norm_d = norm_dest or _norm_caminho_rapido(destino_p)
        chave_dest = f"{norm_d}|{st.st_mtime_ns}|{st.st_size}"
        cache_r = _carregar_cache_rostos()
        with _cache_rostos_lock:
            if rosto_detectado is not None:
                cache_r[chave_dest] = rosto_detectado
            if chave_orig:
                destinos_map[norm_d] = {
                    "chave_orig": chave_orig,
                    "chave_dest": chave_dest,
                    "rosto": rosto_detectado,
                }
    except Exception:
        pass


def _assinatura_subpasta(itens_com_stat, tamanhos_necessarios, cfg, calib):
    """
    Gera assinatura determinística da subpasta em O(N) sem novas chamadas .stat() no disco,
    permitindo carregamento instantâneo (0ms) via `cache_analise_subpastas.json` quando nada mudou.
    """
    partes = ["v5_face1_1024px", ",".join(tamanhos_necessarios), str(cfg.get("ampliacao_maxima", 1.2))]
    for tam in tamanhos_necessarios:
        molde = cfg["moldes"].get(tam, MOLDES_PADRAO.get(tam, (3874, 3012)))
        c_tam = (calib or {}).get(tam, CALIBRACAO_PADRAO)
        partes.append(f"{tam}:{molde[0]}x{molde[1]}:{c_tam.get('centro_x')}:{c_tam.get('olhos_y')}:{c_tam.get('altura_rosto')}")
    for arq, mtime_ns, fsize in itens_com_stat:
        partes.append(f"{arq.name}|{mtime_ns}|{fsize}")
    return ";".join(partes)


def _listar_imagens_subpasta_com_stat(pasta_aluno):
    """
    Lista todas as fotos de uma subpasta de formando usando `os.scandir` recursivo em uma única passada.
    No Windows, `DirEntry.stat()` já vem preenchido pelo sistema operacional sem chamadas extras ao disco.
    Retorna lista ordenada de tuplas: `[(Path(caminho), mtime_ns, size)]`.
    """
    itens = []
    pilha = [str(pasta_aluno)]
    while pilha:
        atual = pilha.pop()
        try:
            with os.scandir(atual) as it:
                sub_dirs = []
                for entry in it:
                    nome = entry.name
                    if nome.startswith(("_", ".")):
                        continue
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            sub_dirs.append(entry.path)
                        elif entry.is_file(follow_symlinks=False):
                            suf = os.path.splitext(nome)[1].lower()
                            if suf in EXTENSOES:
                                st = entry.stat(follow_symlinks=False)
                                itens.append((Path(entry.path), st.st_mtime_ns, st.st_size))
                    except OSError:
                        continue
                for d in sorted(sub_dirs, reverse=True):
                    pilha.append(d)
        except OSError:
            continue
    itens.sort(key=lambda x: x[0].name.casefold())
    return itens


def _listar_imagens_subpasta(pasta_aluno):
    """Compatibilidade: retorna apenas `[Path(caminho), ...]`."""
    return [arq for arq, _, _ in _listar_imagens_subpasta_com_stat(pasta_aluno)]


def _ja_tem_candidatas_excelentes(metricas_por_arquivo, tamanhos_necessarios, cfg, calib, min_candidatas=MAX_FOTOS_POR_ONDA_SUBPASTA):
    """
    Parada inteligente que exige PRIMEIRO ROSTO ÚNICO (`qtd_pessoas == 1`):
    só encerra a busca antecipada na subpasta se já tiver reunido `min_candidatas` fotos
    com EXATAMENTE 1 PESSOA detectada e na orientação ideal. Se as fotos analisadas até aqui
    tiverem mais de 1 pessoa (ex.: fotos com família/amigos) ou não tiverem rosto, continua
    varrendo a subpasta em busca das fotos individuais do formando.
    """
    if not metricas_por_arquivo:
        return False
    alvo_qtd = max(1, int(min_candidatas or MAX_FOTOS_POR_ONDA_SUBPASTA))
    if len(metricas_por_arquivo) < alvo_qtd:
        return False
    amp_max = cfg.get("ampliacao_maxima", 1.2)
    for tam in tamanhos_necessarios:
        molde = cfg["moldes"].get(tam, MOLDES_PADRAO.get(tam, (3874, 3012)))
        calib_tam = calib.get(tam, CALIBRACAO_PADRAO)
        qtd_individual_ok = 0
        for cam_str, m_base in metricas_por_arquivo.items():
            rostos = (m_base or {}).get("rostos") or []
            if not rostos:
                continue
            W_m, H_m = int(m_base.get("W", 0)), int(m_base.get("H", 0))
            if not _orientacao_ideal_para_molde(W_m, H_m, tam, molde):
                continue
            av = avaliar_foto_para_molde(cam_str, tam, molde, calib_tam, amp_max, metricas_base=m_base)
            if (
                av.get("valida")
                and av.get("qtd_pessoas") == 1
                and av.get("orientacao_ideal")
                and int(av.get("encaixe_pct", 0)) >= 68
            ):
                qtd_individual_ok += 1
                if qtd_individual_ok >= min(alvo_qtd, 4):
                    break
        if qtd_individual_ok < min(alvo_qtd, 2):
            return False
    return True


def _init_worker_pool_subpastas(usar_gpu_opencl=True):
    """
    Inicializador de cada processo do `multiprocessing.Pool`:
      1. Reduz a prioridade do processo filho no sistema operacional (`BELOW_NORMAL_PRIORITY_CLASS` no Windows
         ou `os.nice(10)` no Linux/macOS) para nunca engasgar o mouse, navegador ou Photoshop.
      2. Ativa aceleração de hardware por placa de vídeo via OpenCL (`cv2.ocl.setUseOpenCL`) quando disponível.
      3. Limita threads internas do OpenCV (`cv2.setNumThreads(1)`) para respeitar o limite de núcleos escolhido.
      4. Carrega `cache_analise_subpastas.json` e `cache_rostos_yunet.json` em memória e inicializa YuNet ONNX.
    """
    try:
        if sys.platform == "win32":
            # BELOW_NORMAL_PRIORITY_CLASS = 0x00004000
            handle = ctypes.windll.kernel32.GetCurrentProcess()
            ctypes.windll.kernel32.SetPriorityClass(handle, 0x00004000)
        elif hasattr(os, "nice"):
            os.nice(10)
    except Exception:
        pass
    try:
        cv2.setNumThreads(1)
    except Exception:
        pass
    try:
        if hasattr(cv2, "ocl"):
            cv2.ocl.setUseOpenCL(bool(usar_gpu_opencl and cv2.ocl.haveOpenCL()))
    except Exception:
        pass
    try:
        _carregar_cache_analise_subpastas()
        _carregar_cache_rostos()
        if MODELO_YUNET.is_file():
            detector_yunet()
    except Exception:
        pass


def _extrair_metricas_foto_worker(item):
    """
    Worker top-level para `multiprocessing.Pool`: extrai as métricas base de uma única foto
    (`(caminho_str, mtime_ns, fsize, usar_copia_leve)`) em um processo de CPU dedicado.
    """
    if len(item) >= 4:
        cam_str, mtime_ns, fsize, usar_copia_leve = item[:4]
    else:
        cam_str, mtime_ns, fsize = item[:3]
        usar_copia_leve = True
    norm_p = _norm_caminho_rapido(cam_str)
    ch_r = f"{norm_p}|{mtime_ns}|{fsize}"
    ch_an = f"{ch_r}|v3"
    try:
        m_base = _extrair_metricas_base_foto(cam_str, mtime_ns, fsize, usar_copia_leve=bool(usar_copia_leve))
        return ch_an, ch_r, m_base
    except Exception:
        return ch_an, ch_r, None


def _processar_subpasta_worker(tarefa):
    """
    Worker top-level executado em paralelo pelo `multiprocessing.Pool` (respeitando o limite de CPU/GPU):
      1. Lista as imagens via `os.scandir` (`_listar_imagens_subpasta_com_stat`).
      2. Verifica a assinatura no `cache_analise_subpastas.json` (0ms se inalterada).
      3. Faz primeiro a SEPARAÇÃO RÁPIDA POR ORIENTAÇÃO lendo apenas o cabeçalho (< 0.08ms por foto),
         priorizando fotos na orientação ideal do molde e parando imediatamente ao encontrar a 1ª foto boa,
         sem demorar avaliando critérios no restante da pasta.
      4. Copia a foto ORIGINAL escolhida para a pasta FOTO.
    """
    idx_sub = tarefa["idx_sub"]
    sub = Path(tarefa["sub_str"])
    nome_aluno = tarefa["nome_aluno"]
    ch_aluno = tarefa["ch_aluno"]
    plano_aluno = tarefa["plano_aluno"]
    curso_aluno = tarefa["curso_aluno"]
    tamanhos_necessarios = tarefa["tamanhos_necessarios"]
    cfg_serial = tarefa["cfg_serial"]
    calib = tarefa["calib"]
    usar_copia_leve = bool(tarefa.get("usar_copia_leve_local", True))

    pasta_fotos = Path(cfg_serial["fotos"])
    cfg_local = {
        "fotos": pasta_fotos,
        "moldes": {k: tuple(v) for k, v in cfg_serial["moldes"].items()},
        "ampliacao_maxima": float(cfg_serial.get("ampliacao_maxima", 1.2)),
    }

    itens_sub = _listar_imagens_subpasta_com_stat(sub)
    if not itens_sub:
        return {
            "idx_sub": idx_sub,
            "nome_aluno": nome_aluno,
            "total_itens_sub": 0,
            "relatorio_item": {
                "nome": nome_aluno,
                "curso": curso_aluno,
                "plano": plano_aluno,
                "status": "SEM_FOTOS",
                "avisos": ["Nenhuma foto encontrada na subpasta"],
                "escolhidas": {},
                "do_cache": False,
                "fotos_em_cache": 0,
                "fotos_analisadas_agora": 0,
                "proxies_usados": 0,
            },
            "norm_sub": None,
            "sub_cache_entry": None,
            "opcoes_aluno": {},
            "novas_metricas": {},
            "novos_rostos": {},
            "novos_destinos": {},
            "destinos_apagar_manual": [],
        }

    cache_an = _carregar_cache_analise_subpastas()
    subpastas_cache = cache_an.get("_subpastas") or {}
    norm_sub = _norm_caminho_rapido(sub)
    assinatura_atual = _assinatura_subpasta(itens_sub, tamanhos_necessarios, cfg_local, calib)
    reg_sub_cache = subpastas_cache.get(norm_sub)

    opcoes_aluno = {}
    novas_metricas = {}
    novos_rostos = {}
    novos_destinos = {}
    destinos_apagar_manual = []

    # FAST PATH INSTANTÂNEO (0ms): Subpasta inalterada já está em cache_analise_subpastas.json
    if (
        isinstance(reg_sub_cache, dict)
        and reg_sub_cache.get("assinatura") == assinatura_atual
        and isinstance(reg_sub_cache.get("escolhidas"), dict)
    ):
        escolhidas_aluno = reg_sub_cache["escolhidas"]
        avisos_aluno = list(reg_sub_cache.get("avisos") or [])
        candidatas_por_tam = reg_sub_cache.get("candidatas_por_tam") or {}
        rostos_campeas = reg_sub_cache.get("rostos_campeas") or {}

        for tam, campea in escolhidas_aluno.items():
            nome_dest = nome_arquivo_destino_foto(nome_aluno, tam)
            destino = pasta_fotos / nome_dest
            origem_campea = Path(campea["caminho"])
            rosto_c = rostos_campeas.get(tam)
            if origem_campea.is_file():
                _copiar_foto_campea_rapido(origem_campea, destino, rosto_detectado=rosto_c)
                try:
                    st_o = origem_campea.stat()
                    st_d = destino.stat()
                    norm_o = _norm_caminho_rapido(origem_campea)
                    norm_d = _norm_caminho_rapido(destino)
                    ch_o = f"{norm_o}|{st_o.st_mtime_ns}|{st_o.st_size}"
                    ch_d = f"{norm_d}|{st_d.st_mtime_ns}|{st_d.st_size}"
                    novos_destinos[norm_d] = {"chave_orig": ch_o, "chave_dest": ch_d, "rosto": rosto_c}
                    if rosto_c is not None:
                        novos_rostos[ch_d] = rosto_c
                except Exception:
                    pass
            chave_op = f"{ch_aluno}|{tam}"
            opcoes_aluno[chave_op] = {
                "nome": nome_aluno,
                "tamanho": tam,
                "subpasta": str(sub),
                "arquivo_destino": str(destino),
                "escolhida": campea["caminho"],
                "aviso_revisao": campea.get("aviso_revisao", ""),
                "candidatas": candidatas_por_tam.get(tam, [campea]),
            }

        return {
            "idx_sub": idx_sub,
            "nome_aluno": nome_aluno,
            "total_itens_sub": len(itens_sub),
            "relatorio_item": {
                "nome": nome_aluno,
                "curso": curso_aluno,
                "plano": plano_aluno,
                "status": "AVISO_QUALIDADE" if avisos_aluno else ("OK" if escolhidas_aluno else "ERRO"),
                "avisos": avisos_aluno,
                "escolhidas": escolhidas_aluno,
                "do_cache": True,
                "fotos_em_cache": len(itens_sub),
                "fotos_analisadas_agora": 0,
                "proxies_usados": 0,
            },
            "norm_sub": norm_sub,
            "sub_cache_entry": reg_sub_cache,
            "opcoes_aluno": opcoes_aluno,
            "novas_metricas": novas_metricas,
            "novos_rostos": novos_rostos,
            "novos_destinos": novos_destinos,
            "destinos_apagar_manual": destinos_apagar_manual,
        }

    # 1) Verifica foto a foto no cache carregado pelo processo worker
    metricas_por_arquivo = {}
    fotos_para_analisar = []
    fotos_hit_cache = 0
    for arq, mtime_ns, fsize in itens_sub:
        norm_arq = _norm_caminho_rapido(arq)
        ch_rosto = f"{norm_arq}|{mtime_ns}|{fsize}"
        ch_an = f"{ch_rosto}|v3"
        if ch_an in cache_an and isinstance(cache_an[ch_an], dict):
            metricas_por_arquivo[str(arq)] = cache_an[ch_an]
            fotos_hit_cache += 1
        else:
            fotos_para_analisar.append((arq, mtime_ns, fsize, ch_an, ch_rosto))

    fotos_analisadas_sub = 0
    proxies_usados_sub = 0

    alvo_top_sub = min(MAX_FOTOS_POR_ONDA_SUBPASTA, len(itens_sub))
    # Só pula a análise de novas fotos se já tiver as candidatas com EXATAMENTE 1 ROSTO e orientação ideal
    precisa_analisar_novas = not _ja_tem_candidatas_excelentes(
        metricas_por_arquivo, tamanhos_necessarios, cfg_local, calib, min_candidatas=alvo_top_sub
    )

    if fotos_para_analisar and precisa_analisar_novas:
        # Separa primeiro as fotos na orientação ideal, mas mantém todas disponíveis caso as primeiras tenham >1 pessoa
        grupo_ideal, grupo_reserva = _separar_por_orientacao_rapido(
            fotos_para_analisar,
            tamanhos_necessarios,
            cfg_local["moldes"],
            max_orientacao_ideal=max(MAX_TRIAGEM_ORIENTACAO_CABECALHO, len(fotos_para_analisar)),
        )

        # Analisa identificando PRIMEIRO OS ROSTOS: só para quando reunir candidatas individuais (1 única pessoa)
        for idx_cand, (arq, mtime_ns, fsize, ch_an, ch_rosto) in enumerate(grupo_ideal, 1):
            try:
                m_base = _extrair_metricas_base_foto(
                    arq, mtime_ns, fsize, usar_copia_leve=usar_copia_leve
                )
                metricas_por_arquivo[str(arq)] = m_base
                novas_metricas[ch_an] = m_base
                rostos_l = (m_base or {}).get("rostos") or []
                novos_rostos[ch_rosto] = rostos_l[0] if rostos_l else None
                fotos_analisadas_sub += 1
                if usar_copia_leve and (m_base or {}).get("proxy_local"):
                    proxies_usados_sub += 1
            except Exception:
                pass

            qtd_individuais_ate_agora = sum(
                1 for m in metricas_por_arquivo.values()
                if bool((m or {}).get("rostos")) and int((m or {}).get("qtd_pessoas", len((m or {}).get("rostos") or []))) == 1
            )
            if qtd_individuais_ate_agora >= alvo_top_sub and _ja_tem_candidatas_excelentes(
                metricas_por_arquivo, tamanhos_necessarios, cfg_local, calib, min_candidatas=alvo_top_sub
            ):
                break

        # Se ainda faltam fotos individuais (1 única pessoa) para completar as 7 opções, analisa o grupo reserva
        qtd_individuais_total = sum(
            1 for m in metricas_por_arquivo.values()
            if bool((m or {}).get("rostos")) and int((m or {}).get("qtd_pessoas", len((m or {}).get("rostos") or []))) == 1
        )
        faltam_individuais = max(0, MAX_FOTOS_POR_ONDA_SUBPASTA - qtd_individuais_total)
        if faltam_individuais > 0 and grupo_reserva:
            for arq, mtime_ns, fsize, ch_an, ch_rosto in grupo_reserva:
                try:
                    m_base = _extrair_metricas_base_foto(
                        arq, mtime_ns, fsize, usar_copia_leve=usar_copia_leve
                    )
                    metricas_por_arquivo[str(arq)] = m_base
                    novas_metricas[ch_an] = m_base
                    rostos_l = (m_base or {}).get("rostos") or []
                    novos_rostos[ch_rosto] = rostos_l[0] if rostos_l else None
                    fotos_analisadas_sub += 1
                    if usar_copia_leve and (m_base or {}).get("proxy_local"):
                        proxies_usados_sub += 1
                    if len(rostos_l) == 1:
                        qtd_individuais_total += 1
                        if qtd_individuais_total >= MAX_FOTOS_POR_ONDA_SUBPASTA:
                            break
                except Exception:
                    pass

    # 2) Avalia cada tamanho exigido pelo plano identificando PRIMEIRO 1 ROSTO ÚNICO e bloqueando >1 pessoa
    escolhidas_aluno = {}
    avisos_aluno = []
    candidatas_por_tam = {}
    rostos_campeas = {}

    for tam in tamanhos_necessarios:
        molde = cfg_local["moldes"].get(tam, MOLDES_PADRAO.get(tam, (3874, 3012)))
        calib_tam = calib.get(tam, CALIBRACAO_PADRAO)
        amp_max = cfg_local["ampliacao_maxima"]

        avaliacoes = []
        for arq, _, _ in itens_sub:
            m_base = metricas_por_arquivo.get(str(arq))
            if m_base is None:
                continue
            try:
                avaliacoes.append(
                    avaliar_foto_para_molde(arq, tam, molde, calib_tam, amp_max, metricas_base=m_base)
                )
            except Exception:
                pass

        if not avaliacoes:
            continue

        # Hierarquia estrita:
        # 1º Rosto válido E EXATAMENTE 1 PESSOA (`qtd_pessoas == 1`) — nunca escolhe foto em grupo se houver individual!
        # 2º Rosto detectado (`valida == True`)
        # 3º Menor quantidade de pessoas extras
        # 4º Orientação ideal para o molde (`orientacao_ideal == True`)
        # 5º Encaixe pós-corte, densidade de pixels no rosto, qualidade e nota
        avaliacoes.sort(
            key=lambda a: (
                1 if (a.get("valida") and a.get("qtd_pessoas") == 1) else 0,
                1 if a.get("valida") else 0,
                -abs(int(a.get("qtd_pessoas") or 99) - 1),
                1 if a.get("orientacao_ideal") else 0,
                int(a.get("encaixe_pct", 0)),
                int(a.get("densidade_rosto_px", 0)),
                1 if a.get("aprovada_qualidade") else 0,
                float(a.get("nota", 0.0)),
            ),
            reverse=True,
        )

        # Filtra as 7 sugeridas priorizando apenas fotos de 1 única pessoa; só inclui >1 pessoa se não houver 7 individuais
        individuais_validas = [a for a in avaliacoes if a.get("valida") and a.get("qtd_pessoas") == 1]
        if individuais_validas:
            outras_reserva = [a for a in avaliacoes if not (a.get("valida") and a.get("qtd_pessoas") == 1)]
            avaliacoes_ordenadas = individuais_validas + outras_reserva
        else:
            avaliacoes_ordenadas = avaliacoes

        campea = dict(avaliacoes_ordenadas[0])
        if int(campea.get("qtd_pessoas") or 0) > 1:
            aviso_msg = f"Atenção: pasta sem foto individual ({campea.get('qtd_pessoas')} pessoas na foto p/ {tam})"
            campea["aviso_revisao"] = aviso_msg
            avisos_aluno.append(aviso_msg)
        elif not campea.get("valida"):
            aviso_msg = f"Atenção: nenhum rosto detectado na subpasta p/ {tam}"
            campea["aviso_revisao"] = aviso_msg
            avisos_aluno.append(aviso_msg)
        elif not campea.get("orientacao_ideal"):
            aviso_msg = f"Sem foto na orientação ideal p/ {tam} (adaptada de {campea.get('orientacao')})"
            campea["aviso_revisao"] = aviso_msg
            avisos_aluno.append(aviso_msg)
        elif not campea.get("aprovada_qualidade") and campea.get("aviso_qualidade"):
            aviso_msg = f"Escolhida p/ {tam} ({campea['aviso_qualidade']})"
            campea["aviso_revisao"] = aviso_msg
            avisos_aluno.append(aviso_msg)
        else:
            campea["aviso_revisao"] = ""

        nome_dest = nome_arquivo_destino_foto(nome_aluno, tam)
        destino = pasta_fotos / nome_dest
        origem_campea = Path(campea["caminho"])
        m_campea = metricas_por_arquivo.get(str(origem_campea)) or {}
        rosto_campea = (m_campea.get("rostos") or [None])[0]
        rostos_campeas[tam] = rosto_campea

        # Copia apenas a foto CAMPEÃ em qualidade ORIGINAL para a pasta FOTO
        _copiar_foto_campea_rapido(origem_campea, destino, rosto_detectado=rosto_campea)
        destinos_apagar_manual.append(str(destino))
        try:
            st_o = origem_campea.stat()
            st_d = destino.stat()
            norm_o = _norm_caminho_rapido(origem_campea)
            norm_d = _norm_caminho_rapido(destino)
            ch_o = f"{norm_o}|{st_o.st_mtime_ns}|{st_o.st_size}"
            ch_d = f"{norm_d}|{st_d.st_mtime_ns}|{st_d.st_size}"
            novos_destinos[norm_d] = {"chave_orig": ch_o, "chave_dest": ch_d, "rosto": rosto_campea}
            if rosto_campea is not None:
                novos_rostos[ch_d] = rosto_c if "rosto_c" in locals() else rosto_campea
        except Exception:
            pass

        top_candidatas = avaliacoes_ordenadas[:7]
        candidatas_por_tam[tam] = top_candidatas
        chave_op = f"{ch_aluno}|{tam}"
        opcoes_aluno[chave_op] = {
            "nome": nome_aluno,
            "tamanho": tam,
            "subpasta": str(sub),
            "arquivo_destino": str(destino),
            "escolhida": campea["caminho"],
            "aviso_revisao": campea.get("aviso_revisao", ""),
            "candidatas": top_candidatas,
        }
        escolhidas_aluno[tam] = campea

    sub_cache_entry = {
        "assinatura": assinatura_atual,
        "subpasta": str(sub),
        "escolhidas": escolhidas_aluno,
        "avisos": avisos_aluno,
        "candidatas_por_tam": candidatas_por_tam,
        "rostos_campeas": rostos_campeas,
    }

    return {
        "idx_sub": idx_sub,
        "nome_aluno": nome_aluno,
        "total_itens_sub": len(itens_sub),
        "relatorio_item": {
            "nome": nome_aluno,
            "curso": curso_aluno,
            "plano": plano_aluno,
            "status": "AVISO_QUALIDADE" if avisos_aluno else ("OK" if escolhidas_aluno else "ERRO"),
            "avisos": avisos_aluno,
            "escolhidas": escolhidas_aluno,
            "do_cache": fotos_analisadas_sub == 0,
            "fotos_em_cache": fotos_hit_cache,
            "fotos_analisadas_agora": fotos_analisadas_sub,
            "proxies_usados": proxies_usados_sub,
        },
        "norm_sub": norm_sub,
        "sub_cache_entry": sub_cache_entry,
        "opcoes_aluno": opcoes_aluno,
        "novas_metricas": novas_metricas,
        "novos_rostos": novos_rostos,
        "novos_destinos": novos_destinos,
        "destinos_apagar_manual": destinos_apagar_manual,
    }


PASTAS_IGNORADAS_SUBPASTAS = {
    "_originais",
    "_originais_fotos",
    ".cache_proxies_capaz",
    "previa_capas",
    "historico_sessoes",
    "__pycache__",
}


def _tem_imagens_diretas(pasta_p):
    try:
        with os.scandir(pasta_p) as it:
            for entry in it:
                if entry.is_file(follow_symlinks=False) and not entry.name.startswith(("._", ".")):
                    if os.path.splitext(entry.name)[1].lower() in EXTENSOES:
                        return True
    except OSError:
        pass
    return False


def _descobrir_subpastas_formandos(pasta_mae):
    """
    Localiza todas as subpastas reais de formandos dentro de `pasta_mae`, ignorando pastas técnicas
    (`_ORIGINAIS`, `.cache_proxies_capaz`, `previa_capas`, etc.) e descendo 1 nível caso as subpastas
    estejam agrupadas por turma/curso (ex.: `PASTA / DIREITO / MARIA DA SILVA`).
    """
    pasta_p = Path(pasta_mae)
    if not pasta_p.is_dir():
        return []
    subpastas = []
    try:
        with os.scandir(pasta_p) as it:
            nivel1 = []
            for entry in it:
                nome_low = entry.name.lower()
                if entry.name.startswith(("_", ".")) or nome_low in PASTAS_IGNORADAS_SUBPASTAS:
                    continue
                if entry.is_dir(follow_symlinks=False):
                    nivel1.append(Path(entry.path))
    except OSError:
        return []

    for d1 in nivel1:
        if _tem_imagens_diretas(d1):
            subpastas.append(d1)
            continue
        # Se não tem fotos diretas, verifica se é uma pasta de turma/curso com subpastas de alunos dentro
        filhas = []
        try:
            with os.scandir(d1) as it2:
                for e2 in it2:
                    if e2.name.startswith(("_", ".")) or e2.name.lower() in PASTAS_IGNORADAS_SUBPASTAS:
                        continue
                    if e2.is_dir(follow_symlinks=False):
                        filhas.append(Path(e2.path))
        except OSError:
            pass
        if len(filhas) >= 2:
            for d2 in filhas:
                if _listar_imagens_subpasta_com_stat(d2):
                    subpastas.append(d2)
        elif len(filhas) == 1 and _listar_imagens_subpasta_com_stat(d1):
            # Ex.: pasta do aluno contendo subpasta "FOTOS"
            subpastas.append(d1)

    subpastas.sort(key=lambda d: chave(d.name))
    return subpastas


def escolher_melhores_fotos_subpastas(
    pasta_subpastas,
    cfg,
    calib=None,
    alunos_planos=None,
    plano_padrao="Prime Book 30X20",
    curso_padrao="",
    berlim_modo="Ambos",
    progresso_cb=None,
    limite_cpu_pct=None,
    usar_gpu_opencl=None,
    usar_copia_leve_local=None,
):
    """
    Divide as subpastas dos formandos entre os núcleos da CPU permitidos pelo limite de hardware
    (`limite_cpu_pct`) usando `ThreadPoolExecutor` (com liberação de GIL C++ no OpenCV/YuNet) + GPU OpenCL:
      1. Faz primeiro uma SEPARAÇÃO RÁPIDA POR ORIENTAÇÃO via cabeçalho (< 0.08ms/foto, sem decodificar),
         filtrando apenas as fotos na orientação certa do molde.
      2. Identifica primeiro os rostos e prioriza estritamente fotos com 1 única pessoa (`qtd_pessoas == 1`).
      3. Copia em resolução ORIGINAL apenas a foto campeã escolhida para a pasta FOTO.
    """
    pasta_mae = Path(pasta_subpastas)
    if not pasta_mae.is_dir():
        raise FileNotFoundError("Pasta de subpastas dos formandos não encontrada: " + str(pasta_mae))

    if calib is None:
        calib = carregar_calibracao(cfg)

    if usar_copia_leve_local is None:
        usar_copia_leve_local = cfg.get("usar_copia_leve_local", USAR_COPIA_LEVE_LOCAL_PADRAO)
    usar_copia_leve = bool(usar_copia_leve_local)

    hw = calcular_recursos_hardware(
        limite_cpu_pct=limite_cpu_pct if limite_cpu_pct is not None else cfg.get("hardware_limite_cpu_pct"),
        usar_gpu_opencl=usar_gpu_opencl if usar_gpu_opencl is not None else cfg.get("hardware_usar_gpu_opencl"),
    )

    cfg["fotos"].mkdir(parents=True, exist_ok=True)
    if usar_copia_leve:
        try:
            PASTA_PROXIES_ANALISE.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

    try:
        if hasattr(cv2, "ocl"):
            cv2.ocl.setUseOpenCL(bool(hw["usar_gpu_opencl"] and cv2.ocl.haveOpenCL()))
    except Exception:
        pass

    # Inicializa detector uma vez antes das threads para aquecer o modelo em memória
    detector_yunet()

    verificar_cache_analise_subpastas(
        cfg["fotos"],
        limite_cpu_pct=hw["limite_cpu_pct"],
        usar_gpu_opencl=hw["usar_gpu_opencl"],
    )
    cache_an = _carregar_cache_analise_subpastas()
    cache_r = _carregar_cache_rostos()
    subpastas_cache = cache_an.setdefault("_subpastas", {})
    destinos_map = cache_an.setdefault("_destinos_copiados", {})
    _salvar_cache_analise_subpastas()

    subpastas = _descobrir_subpastas_formandos(pasta_mae)
    if subpastas:
        try:
            gravar_config({"subpastas_fotos_dir": str(pasta_mae)})
        except Exception:
            pass
    else:
        subpastas = [pasta_mae]

    mapa_planos = {}
    mapa_cursos = {}
    if alunos_planos:
        for item in alunos_planos:
            ch = chave(item.get("nome", ""))
            if ch:
                mapa_planos[ch] = item.get("plano") or plano_padrao
                if item.get("curso"):
                    mapa_cursos[ch] = item["curso"]

    _, cursos_salvos_nome = _cursos_salvos_por_foto()
    opcoes_salvas = ler_opcoes_foto()
    total_sub = len(subpastas)
    nucleos_usados = hw["nucleos_usados"]
    nucleos_totais = hw["nucleos_totais"]
    tag_hw = f"{nucleos_usados}/{nucleos_totais} núcleos ({hw['limite_cpu_pct']}%)" + (
        " + GPU OpenCL" if hw["gpu_ativa"] else ""
    )

    cfg_serial = {
        "fotos": str(cfg["fotos"]),
        "moldes": {k: list(v) for k, v in cfg["moldes"].items()},
        "ampliacao_maxima": float(cfg.get("ampliacao_maxima", 1.2)),
    }

    tarefas = []
    for idx_sub, sub in enumerate(subpastas, 1):
        bruto_pasta = sub.name if sub != pasta_mae else pasta_mae.name
        tam_na_pasta = extrair_tamanho_do_texto(bruto_pasta)
        nome_aluno = formatar_nome(bruto_pasta)
        ch_aluno = chave(nome_aluno)
        plano_inferido_pasta = PLANO_POR_TAMANHO.get(tam_na_pasta) if tam_na_pasta else None
        plano_aluno = mapa_planos.get(ch_aluno) or plano_inferido_pasta or plano_padrao
        curso_aluno = mapa_cursos.get(ch_aluno) or cursos_salvos_nome.get(ch_aluno) or curso_padrao
        tamanhos_necessarios = tamanhos_do_plano(plano_aluno, berlim_modo)
        tarefas.append({
            "idx_sub": idx_sub,
            "sub_str": str(sub),
            "nome_aluno": nome_aluno,
            "ch_aluno": ch_aluno,
            "plano_aluno": plano_aluno,
            "curso_aluno": curso_aluno,
            "tamanhos_necessarios": tamanhos_necessarios,
            "cfg_serial": cfg_serial,
            "calib": calib,
            "usar_copia_leve_local": usar_copia_leve,
        })

    resultados_ordenados = []
    novas_acumuladas_desde_save = 0
    destinos_para_limpar_manual = []

    def _integrar_resultado_worker(res, concluidas):
        nonlocal novas_acumuladas_desde_save
        if not res:
            return
        rel_item = res["relatorio_item"]
        rel_item["hardware_resumo"] = tag_hw
        resultados_ordenados.append((res["idx_sub"], rel_item))

        novas_m = res.get("novas_metricas") or {}
        novos_r = res.get("novos_rostos") or {}
        novos_d = res.get("novos_destinos") or {}
        norm_sub = res.get("norm_sub")
        sub_entry = res.get("sub_cache_entry")
        op_aluno = res.get("opcoes_aluno") or {}
        destinos_para_limpar_manual.extend(res.get("destinos_apagar_manual") or [])

        with _cache_rostos_lock:
            if novas_m:
                cache_an.update(novas_m)
            if norm_sub and sub_entry is not None:
                subpastas_cache[norm_sub] = sub_entry
            if novos_d:
                destinos_map.update(novos_d)
            if novos_r:
                cache_r.update(novos_r)
        if op_aluno:
            opcoes_salvas.update(op_aluno)

        novas_agora = int(rel_item.get("fotos_analisadas_agora", 0))
        novas_acumuladas_desde_save += novas_agora

        if progresso_cb:
            if rel_item.get("do_cache"):
                progresso_cb(
                    concluidas,
                    total_sub,
                    f"Hardware [{tag_hw}] · Cache 0ms: {res['nome_aluno']} ({res.get('total_itens_sub', 0)} fotos)...",
                )
            else:
                tag_proxy = "orientação rápida + cópia leve" if usar_copia_leve else "orientação rápida"
                progresso_cb(
                    concluidas,
                    total_sub,
                    f"Hardware [{tag_hw}] · {res['nome_aluno']} ({novas_agora} via {tag_proxy} / {rel_item.get('fotos_em_cache', 0)} cache)...",
                )

        if novas_acumuladas_desde_save >= 36:
            _salvar_cache_analise_subpastas()
            novas_acumuladas_desde_save = 0

    if total_sub >= 2:
        num_processos = max(1, min(total_sub, nucleos_usados))
        if progresso_cb:
            progresso_cb(
                0,
                total_sub,
                f"Separando rápido por Orientação com {tag_hw} ({hw['nucleos_livres']} núcleos livres)...",
            )
        with concurrent.futures.ThreadPoolExecutor(max_workers=num_processos) as tpool:
            concluidas = 0
            futures = [tpool.submit(_processar_subpasta_worker, t) for t in tarefas]
            for fut in concurrent.futures.as_completed(futures):
                concluidas += 1
                try:
                    res = fut.result()
                except Exception:
                    res = None
                _integrar_resultado_worker(res, concluidas)
    else:
        t_unica = tarefas[0]
        res_unico = _processar_subpasta_worker(t_unica)
        _integrar_resultado_worker(res_unico, 1)

    for dest_str in destinos_para_limpar_manual:
        try:
            apagar_manual(Path(dest_str))
        except Exception:
            pass

    _salvar_cache_analise_subpastas()
    _salvar_cache_rostos()
    gravar_opcoes_foto(opcoes_salvas)

    resultados_ordenados.sort(key=lambda par: par[0])
    return [item for _, item in resultados_ordenados]


RE_SUFIXO_NUMERACAO_FOTO = re.compile(r"(?i)[\s_\-\.]*(?:\(\d+\)|\[\d+\]|copy|\bcopia\b|\bimg[_\-\s]*\d+|\bdsc[_\-\s]*\d+|\bfoto[_\-\s]*\d+|\b\d{1,3})$")


def _nome_base_agrupamento(caminho_ou_nome):
    """Normaliza o nome do formando removendo também sufixos de numeração como 'Maria (1)', 'Maria_02'."""
    limpo = limpar_tamanho_do_nome(caminho_ou_nome)
    limpo = RE_SUFIXO_NUMERACAO_FOTO.sub("", limpo).strip()
    return formatar_nome(limpo) or formatar_nome(caminho_ou_nome)


def _ordenar_avaliacoes_candidatas(avaliacoes):
    avaliacoes.sort(
        key=lambda a: (
            1 if (a.get("valida") and a.get("qtd_pessoas") == 1) else 0,
            1 if a.get("valida") else 0,
            -abs(int(a.get("qtd_pessoas") or 99) - 1),
            1 if a.get("orientacao_ideal") else 0,
            int(a.get("encaixe_pct", 0)),
            int(a.get("densidade_rosto_px", 0)),
            1 if a.get("aprovada_qualidade") else 0,
            float(a.get("nota", 0.0)),
        ),
        reverse=True,
    )
    individuais_validas = [a for a in avaliacoes if a.get("valida") and a.get("qtd_pessoas") == 1]
    if individuais_validas:
        outras_reserva = [a for a in avaliacoes if not (a.get("valida") and a.get("qtd_pessoas") == 1)]
        return individuais_validas + outras_reserva
    return avaliacoes


def obter_opcoes_aluno_inteligente(nome_aluno, tam, cfg, calib=None, arquivo_atual=None):
    """
    Localiza ou constrói automaticamente as até 7 melhores fotos candidatas de um formando
    para exibir tanto na aba 'Revisão e Recorte' quanto na 'Tela: Escolher 1 entre as 7 Melhores':
      1. Verifica `opcoes_por_foto.json` pela chave exata `chave(nome)|tam` ou por qualquer tamanho do mesmo aluno.
      2. Se não estiver em `opcoes_por_foto.json`, procura automaticamente a subpasta do aluno em
         `subpastas_fotos_dir`, `cfg["fotos"]`, `cfg["backup"]` (`_ORIGINAIS`) ou `cfg["fotos"].parent`.
      3. Se não houver subpasta, agrupa fotos do aluno na pasta `FOTO` / `_ORIGINAIS` (incluindo a foto atual)
         e calcula os indicadores reais de Contraste, Nitidez, Brilho e Rosto.
    """
    if calib is None:
        calib = carregar_calibracao(cfg)

    tam_str = str(tam or "30X20").upper()
    tam_foto = "30X20" if tam_str == "30X22" else (tam_str if tam_str in ("20X20", "24X34", "30X20") else "30X20")
    molde = cfg["moldes"].get(tam_foto, MOLDES_PADRAO.get(tam_foto, (3874, 3012)))
    calib_tam = calib.get(tam_foto, CALIBRACAO_PADRAO)
    amp_max = float(cfg.get("ampliacao_maxima", 1.2))

    nome_limpo = formatar_nome(nome_aluno or (nome_do_aluno(arquivo_atual) if arquivo_atual else ""))
    ch_aluno = chave(nome_limpo)
    ch_base = chave(_nome_base_agrupamento(nome_limpo))
    norm_arq_atual = _norm_caminho_rapido(arquivo_atual) if arquivo_atual else ""

    opcoes_todas = ler_opcoes_foto()

    # 1) Busca direta em opcoes_por_foto.json (chave exata)
    for k_tentativa in (f"{ch_aluno}|{tam_foto}", f"{ch_aluno}|{tam_str}", f"{ch_base}|{tam_foto}"):
        dados_op = opcoes_todas.get(k_tentativa)
        if isinstance(dados_op, dict):
            cands_existentes = [
                c for c in (dados_op.get("candidatas") or [])
                if isinstance(c, dict) and c.get("caminho") and Path(c["caminho"]).is_file()
            ]
            if cands_existentes:
                dados_ret = dict(dados_op)
                dados_ret["candidatas"] = cands_existentes[:7]
                if not dados_ret.get("escolhida") or not Path(str(dados_ret.get("escolhida"))).is_file():
                    dados_ret["escolhida"] = cands_existentes[0]["caminho"]
                return dados_ret

    # 2) Busca em opcoes_por_foto.json por outro tamanho do mesmo aluno ou pelo arquivo_destino
    caminhos_candidatos = []
    subpasta_conhecida = None
    escolhida_anterior = None
    for k_op, v_op in opcoes_todas.items():
        if not isinstance(v_op, dict):
            continue
        ch_v = chave(limpar_tamanho_do_nome(v_op.get("nome") or k_op.split("|")[0]))
        dest_v = _norm_caminho_rapido(v_op.get("arquivo_destino")) if v_op.get("arquivo_destino") else ""
        if ch_v in (ch_aluno, ch_base) or (norm_arq_atual and dest_v == norm_arq_atual):
            if v_op.get("subpasta") and Path(v_op["subpasta"]).is_dir():
                subpasta_conhecida = Path(v_op["subpasta"])
            if v_op.get("escolhida") and Path(str(v_op["escolhida"])).is_file():
                escolhida_anterior = str(v_op["escolhida"])
            for c in (v_op.get("candidatas") or []):
                cam_c = c.get("caminho") if isinstance(c, dict) else None
                if cam_c and Path(cam_c).is_file():
                    caminhos_candidatos.append(Path(cam_c))

    # 3) Se não achou subpasta ainda, procura subpasta correspondente no disco
    if subpasta_conhecida is None and not caminhos_candidatos:
        bruto_cfg = ler_config_bruto()
        raizes_busca = []
        for p_cand in (
            bruto_cfg.get("subpastas_fotos_dir"),
            cfg.get("fotos"),
            cfg.get("backup"),
            Path(cfg["fotos"]).parent if cfg.get("fotos") else None,
        ):
            if p_cand:
                p_obj = Path(p_cand)
                if p_obj.is_dir() and p_obj not in raizes_busca:
                    raizes_busca.append(p_obj)

        for raiz in raizes_busca:
            for sub_d in _descobrir_subpastas_formandos(raiz):
                ch_sub = chave(limpar_tamanho_do_nome(sub_d.name))
                ch_sub_base = chave(_nome_base_agrupamento(sub_d.name))
                if ch_sub in (ch_aluno, ch_base) or ch_sub_base in (ch_aluno, ch_base):
                    subpasta_conhecida = sub_d
                    break
                if len(ch_aluno) >= 4 and (ch_aluno in ch_sub or ch_sub in ch_aluno):
                    subpasta_conhecida = sub_d
                    break
            if subpasta_conhecida is not None:
                break

    if subpasta_conhecida is not None and not caminhos_candidatos:
        itens_sub = _listar_imagens_subpasta_com_stat(subpasta_conhecida)
        if itens_sub:
            faltantes = []
            for arq, mtime_ns, fsize in itens_sub:
                ch_rosto = f"{_norm_caminho_rapido(arq)}|{mtime_ns}|{fsize}"
                faltantes.append((arq, mtime_ns, fsize, f"{ch_rosto}|v3", ch_rosto))
            g_ideal, g_res = _separar_por_orientacao_rapido(
                faltantes, [tam_foto], cfg["moldes"], max_orientacao_ideal=MAX_TRIAGEM_ORIENTACAO_CABECALHO
            )
            for arq, _, _, _, _ in (g_ideal + g_res)[:14]:
                caminhos_candidatos.append(arq)

    # 4) Também inclui fotos do mesmo aluno soltas em FOTO ou _ORIGINAIS
    if not caminhos_candidatos:
        vistos_norm = set()
        if arquivo_atual and Path(arquivo_atual).is_file():
            orig_at = fonte_original(Path(arquivo_atual), cfg)
            if orig_at.is_file():
                caminhos_candidatos.append(orig_at)
                vistos_norm.add(_norm_caminho_rapido(orig_at))
        for pasta_busca in (cfg.get("backup"), cfg.get("fotos")):
            if not pasta_busca or not Path(pasta_busca).is_dir():
                continue
            for f in listar_fotos(pasta_busca):
                norm_f = _norm_caminho_rapido(f)
                if norm_f in vistos_norm:
                    continue
                ch_f = chave(nome_do_aluno(f))
                ch_f_base = chave(_nome_base_agrupamento(f.stem))
                if ch_f in (ch_aluno, ch_base) or ch_f_base in (ch_aluno, ch_base):
                    orig_f = fonte_original(f, cfg)
                    norm_of = _norm_caminho_rapido(orig_f)
                    if norm_of not in vistos_norm and orig_f.is_file():
                        caminhos_candidatos.append(orig_f)
                        vistos_norm.add(norm_of)

    if not caminhos_candidatos:
        return None

    # Avalia e ordena as candidatas encontradas
    usar_copia_leve = bool(cfg.get("usar_copia_leve_local", USAR_COPIA_LEVE_LOCAL_PADRAO))
    avaliacoes = []
    vistos_c = set()
    for arq_c in caminhos_candidatos:
        norm_c = _norm_caminho_rapido(arq_c)
        if norm_c in vistos_c or not Path(arq_c).is_file():
            continue
        vistos_c.add(norm_c)
        try:
            m_base = _extrair_metricas_base_foto(arq_c, usar_copia_leve=usar_copia_leve)
            av = avaliar_foto_para_molde(arq_c, tam_foto, molde, calib_tam, amp_max, metricas_base=m_base)
            avaliacoes.append(av)
        except Exception:
            pass

    if not avaliacoes:
        return None

    avaliacoes_ordenadas = _ordenar_avaliacoes_candidatas(avaliacoes)[:7]
    campea = avaliacoes_ordenadas[0]
    destino_padrao = cfg["fotos"] / nome_arquivo_destino_foto(nome_limpo, tam_foto)
    escolhida_final = escolhida_anterior or (
        str(fonte_original(Path(arquivo_atual), cfg)) if (arquivo_atual and Path(arquivo_atual).is_file()) else campea["caminho"]
    )
    if not any(_norm_caminho_rapido(c["caminho"]) == _norm_caminho_rapido(escolhida_final) for c in avaliacoes_ordenadas):
        escolhida_final = campea["caminho"]

    registro_op = {
        "nome": nome_limpo,
        "tamanho": tam_foto,
        "subpasta": str(subpasta_conhecida) if subpasta_conhecida else "",
        "arquivo_destino": str(arquivo_atual or destino_padrao),
        "escolhida": escolhida_final,
        "aviso_revisao": campea.get("aviso_revisao", "") if not campea.get("aprovada_qualidade", True) else "",
        "candidatas": avaliacoes_ordenadas,
    }

    # Salva no opcoes_por_foto.json se veio de subpasta ou tem múltiplas opções
    if subpasta_conhecida is not None or len(avaliacoes_ordenadas) >= 1:
        chave_salvar = f"{ch_aluno}|{tam_foto}"
        opcoes_todas[chave_salvar] = registro_op
        gravar_opcoes_foto(opcoes_todas)
        _salvar_cache_analise_subpastas()
        _salvar_cache_rostos()

    return registro_op


def garantir_opcoes_foto_disponiveis(cfg, calib=None, alunos_atuais=None):
    """
    Garante que `opcoes_por_foto.json` contenha registros válidos para abrir a
    'Tela: Escolher 1 entre as 7 Melhores' mesmo quando o usuário clicou apenas em
    'Carregar Fotos da Pasta' ou colocou subpastas dentro de FOTO / subpastas_fotos_dir.
    """
    if calib is None:
        calib = carregar_calibracao(cfg)

    bruto_cfg = ler_config_bruto()
    # 1) Se houver subpastas em subpastas_fotos_dir ou dentro de cfg["fotos"] ainda não processadas, processa
    opcoes_atuais = ler_opcoes_foto()
    opcoes_validas = {}
    for k, v in opcoes_atuais.items():
        if not isinstance(v, dict):
            continue
        cands = [
            c for c in (v.get("candidatas") or [])
            if isinstance(c, dict) and c.get("caminho") and Path(c["caminho"]).is_file()
        ]
        if cands:
            v_copia = dict(v)
            v_copia["candidatas"] = cands[:7]
            opcoes_validas[k] = v_copia

    if not opcoes_validas:
        for pasta_cand in (bruto_cfg.get("subpastas_fotos_dir"), cfg.get("fotos")):
            if pasta_cand and Path(pasta_cand).is_dir():
                subs = _descobrir_subpastas_formandos(pasta_cand)
                if subs:
                    try:
                        escolher_melhores_fotos_subpastas(
                            pasta_cand,
                            cfg,
                            calib=calib,
                            alunos_planos=alunos_atuais,
                        )
                        opcoes_validas = ler_opcoes_foto()
                        if opcoes_validas:
                            break
                    except Exception:
                        pass

    # 2) Garante que os alunos atuais ou fotos da pasta FOTO também estejam presentes
    fotos_pasta = listar_fotos(cfg["fotos"]) if cfg.get("fotos") and Path(cfg["fotos"]).is_dir() else []
    for f in fotos_pasta:
        nm = nome_do_aluno(f)
        tm = tamanho_do_arquivo(f)
        ch_op = f"{chave(nm)}|{tm}"
        if ch_op not in opcoes_validas:
            reg = obter_opcoes_aluno_inteligente(nm, tm, cfg, calib=calib, arquivo_atual=f)
            if reg and reg.get("candidatas"):
                opcoes_validas[ch_op] = reg

    if opcoes_validas != opcoes_atuais and opcoes_validas:
        gravar_opcoes_foto(opcoes_validas)
    return opcoes_validas


def trocar_foto_por_opcao(nome_aluno, tam, caminho_origem, cfg):
    """Substitui na pasta FOTO a foto do aluno pela candidata escolhida na tela das 7 melhores ou na Revisão."""
    origem = Path(caminho_origem)
    if not origem.is_file():
        raise FileNotFoundError("Foto candidata não encontrada: " + str(origem))
    tam_real = "30X20" if str(tam).upper() == "30X22" else str(tam).upper()
    nome_dest = nome_arquivo_destino_foto(nome_aluno, tam_real)
    destino = cfg["fotos"] / nome_dest
    bkp = backup_de(destino, cfg)
    if bkp and bkp.is_file():
        try:
            bkp.unlink()
        except OSError:
            pass

    _copiar_foto_campea_rapido(origem, destino)
    apagar_manual(destino)

    opcoes = ler_opcoes_foto()
    for t_key in {tam_real, str(tam).upper()}:
        chave_op = f"{chave(nome_aluno)}|{t_key}"
        cand_match = None
        if chave_op in opcoes:
            opcoes[chave_op]["escolhida"] = str(origem)
            cand_match = next(
                (c for c in (opcoes[chave_op].get("candidatas") or []) if str(c.get("caminho")) == str(origem)),
                None,
            )
            opcoes[chave_op]["aviso_revisao"] = (
                (cand_match or {}).get("aviso_qualidade", "")
                if cand_match and not cand_match.get("aprovada_qualidade")
                else ""
            )
    gravar_opcoes_foto(opcoes)

    # Atualiza também o cache_analise_subpastas.json para lembrar a escolhida entre as 7 melhores
    try:
        cache_an = _carregar_cache_analise_subpastas()
        subpastas_map = cache_an.get("_subpastas") or {}
        norm_sub = _norm_caminho_rapido(origem.parent)
        reg_sub = subpastas_map.get(norm_sub)
        if isinstance(reg_sub, dict) and isinstance(reg_sub.get("escolhidas"), dict):
            cand_m = next(
                (
                    c for k_op, v_op in opcoes.items()
                    if isinstance(v_op, dict)
                    for c in (v_op.get("candidatas") or [])
                    if str(c.get("caminho")) == str(origem)
                ),
                None,
            )
            if cand_m:
                reg_sub["escolhidas"][tam_real] = dict(cand_m)
                _salvar_cache_analise_subpastas()
    except Exception:
        pass
    return destino


def obter_miniatura_candidata_rapida(
    caminho_foto,
    tam,
    cfg,
    calib=None,
    caixa_sugerida=None,
    lado_max=220,
    modo="recorte",
    tamanho_canvas=None,
):
    """
    Carrega em ~2-5ms a miniatura de uma das 7 melhores fotos candidatas (usando `.cache_proxies_capaz`
    via `_garantir_copia_leve_local` ou `abrir_foto_previa_rapida`) para exibir na Tela de Escolha e na Revisão:
      - `modo="recorte"`: retorna a foto já simulando o enquadramento no molde da capa (`20X20`, `30X20`, `24X34`).
      - `modo="inteira"`: retorna a foto inteira com o retângulo amarelo (`#FDE047`) do corte e linha azul dos olhos.
      - `tamanho_canvas=(larg, alt)`: centraliza a miniatura em um canvas fixo `#23272f` para estabilidade no Tkinter.
    """
    caminho_p = Path(caminho_foto)
    tam_molde = "30X20" if str(tam).upper() == "30X22" else str(tam).upper()
    molde = cfg["moldes"].get(tam_molde, MOLDES_PADRAO.get(tam_molde, (3874, 3012)))
    calib_tam = (calib or {}).get(tam_molde, CALIBRACAO_PADRAO)

    img_red = None
    tam_orig = None
    try:
        st = caminho_p.stat()
        bgr_p, tam_orig, _, _ = _garantir_copia_leve_local(caminho_p, mtime_ns=st.st_mtime_ns, fsize=st.st_size)
        if bgr_p is not None and bgr_p.size > 0:
            rgb_p = cv2.cvtColor(bgr_p, cv2.COLOR_BGR2RGB)
            img_red = Image.fromarray(rgb_p)
    except Exception:
        img_red = None

    if img_red is None:
        img_red, tam_orig, _ = abrir_foto_previa_rapida(caminho_p, max(512, lado_max * 2))

    W_orig, H_orig = tam_orig if tam_orig else img_red.size
    if caixa_sugerida and len(caixa_sugerida) == 4:
        caixa = tuple(int(round(v)) for v in caixa_sugerida)
    else:
        rosto = detectar_rosto_cache(caminho_p)
        caixa = (
            calcular_recorte((W_orig, H_orig), rosto, molde, calib_tam)
            if rosto
            else ajustar_caixa((0, 0, W_orig, H_orig), (W_orig, H_orig), molde)
        )

    fx = img_red.width / float(max(1, W_orig))
    fy = img_red.height / float(max(1, H_orig))
    cl = max(0, min(img_red.width - 1, int(round(caixa[0] * fx))))
    ct = max(0, min(img_red.height - 1, int(round(caixa[1] * fy))))
    cr = max(cl + 1, min(img_red.width, int(round(caixa[2] * fx))))
    cb = max(ct + 1, min(img_red.height, int(round(caixa[3] * fy))))

    if modo == "inteira":
        vis = img_red.copy()
        draw = ImageDraw.Draw(vis)
        draw.rectangle((cl, ct, cr, cb), outline="#FDE047", width=max(2, round(vis.width / 110)))
        oy = int(round(ct + (cb - ct) * calib_tam.get("olhos_y", 0.38)))
        draw.line((cl, oy, cr, oy), fill="#38BDF8", width=max(1, round(vis.width / 160)))
        resultado = vis
    else:
        resultado = img_red.crop((cl, ct, cr, cb))

    if tamanho_canvas and len(tamanho_canvas) == 2:
        cw, ch = max(24, int(tamanho_canvas[0])), max(24, int(tamanho_canvas[1]))
        resultado.thumbnail((cw - 4, ch - 4), Image.BILINEAR)
        canvas = Image.new("RGB", (cw, ch), (35, 39, 47))
        ox = (cw - resultado.width) // 2
        oy = (ch - resultado.height) // 2
        canvas.paste(resultado, (ox, oy))
        return canvas

    resultado.thumbnail((lado_max, lado_max), Image.BILINEAR)
    return resultado


def abrir_foto(caminho):
    img = Image.open(caminho)
    icc = img.info.get("icc_profile")
    dpi = img.info.get("dpi")
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    return img, icc, dpi



def no_molde(tam, molde):
    return (
        abs(tam[0] - molde[0]) <= TOLERANCIA_MOLDE_PX
        and abs(tam[1] - molde[1]) <= TOLERANCIA_MOLDE_PX
    )


def tamanho_rapido(caminho):
    """Largura x altura já girada pelo EXIF, lendo só o cabeçalho."""
    with Image.open(caminho) as img:
        w, h = img.size
        try:
            orientacao = img.getexif().get(0x0112, 1)
        except Exception:
            orientacao = 1
    return (h, w) if orientacao in (5, 6, 7, 8) else (w, h)


def backup_de(caminho, cfg):
    """Cópia em _ORIGINAIS do mesmo arquivo (o original pode ter outra extensão)."""
    pasta = cfg["backup"]
    exato = pasta / caminho.name
    if exato.is_file():
        return exato
    if pasta.is_dir():
        for f in pasta.iterdir():
            if f.is_file() and f.stem == caminho.stem and f.suffix.lower() in EXTENSOES:
                return f
    return None


def fonte_original(caminho, cfg):
    """Sempre corta a partir do original guardado, para não perder qualidade."""
    return backup_de(caminho, cfg) or caminho


# ============================================================
# CORTE MANUAL
# ============================================================

def ler_manual():
    if MANUAL_FILE.is_file():
        try:
            return json.loads(MANUAL_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _gravar_manual(dados):
    temp = MANUAL_FILE.with_suffix(".tmp")
    temp.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(temp, MANUAL_FILE)


def gravar_manual(caminho, tam, caixa, tamanho_fonte):
    """Guarda a caixa em frações da foto original (independe da resolução)."""
    W, H = tamanho_fonte
    dados = ler_manual()
    dados[caminho.stem] = {
        "tamanho": tam,
        "caixa": [round(caixa[0] / W, 6), round(caixa[1] / H, 6),
                  round(caixa[2] / W, 6), round(caixa[3] / H, 6)],
        "origem": [W, H],
        "data": datetime.datetime.now().isoformat(timespec="seconds"),
    }
    _gravar_manual(dados)


def apagar_manual(caminho):
    dados = ler_manual()
    if dados.pop(caminho.stem, None) is not None:
        _gravar_manual(dados)


def caixa_manual(caminho, tam, tamanho_fonte, molde):
    """Caixa manual em pixels da foto original, ou None."""
    reg = ler_manual().get(caminho.stem)
    if not reg or reg.get("tamanho") != tam:
        return None
    W, H = tamanho_fonte
    l, t, r, b = reg["caixa"]
    return ajustar_caixa((l * W, t * H, r * W, b * H), tamanho_fonte, molde)


def ajustar_caixa(caixa, tamanho_fonte, molde):
    """Força a proporção do molde e mantém a caixa dentro da foto."""
    W, H = tamanho_fonte
    mw, mh = molde
    l, t, r, b = caixa
    cx, cy = (l + r) / 2, (t + b) / 2
    h = max(10.0, b - t)
    w = h * mw / mh
    fator = min(1.0, W / w, H / h)
    w *= fator
    h *= fator
    l = min(max(0.0, cx - w / 2), W - w)
    t = min(max(0.0, cy - h / 2), H - h)
    return (int(round(l)), int(round(t)), int(round(l + w)), int(round(t + h)))


# ============================================================
# CALIBRAÇÃO
# ============================================================

def calibrar(cfg):
    resultado = {}
    amostras = {k: [] for k in cfg["moldes"]}

    for f in listar_fotos(cfg["fotos"]):
        tam = tamanho_do_arquivo(f)
        molde = cfg["moldes"].get(tam)
        if not molde:
            continue
        try:
            img, _, _ = abrir_foto(f)
        except Exception as e:
            print("ERRO ao abrir", f.name, e)
            continue
        if not no_molde(img.size, molde):
            continue
        rosto = detectar_rosto(img)
        if not rosto or rosto["metodo"] != "yunet":
            print(f"  sem rosto: {f.name}")
            continue
        W, H = img.size
        amostras[tam].append((rosto["cx"] / W, rosto["olhos_y"] / H, rosto["h"] / H))
        print(f"  {tam}  {f.name}")

    for tam, lista in amostras.items():
        if lista:
            resultado[tam] = {
                "centro_x": round(statistics.median(a[0] for a in lista), 4),
                "olhos_y": round(statistics.median(a[1] for a in lista), 4),
                "altura_rosto": round(statistics.median(a[2] for a in lista), 4),
                "amostras": len(lista),
            }

    # Tamanhos sem amostra herdam de um vertical calibrado
    for tam in cfg["moldes"]:
        if tam not in resultado:
            base = resultado.get("20X20") or resultado.get("24X34") or dict(CALIBRACAO_PADRAO)
            resultado[tam] = {**base, "amostras": 0, "herdado": True}

    CALIBRACAO_FILE.write_text(json.dumps(resultado, indent=2, ensure_ascii=False), encoding="utf-8")

    print()
    print("CALIBRAÇÃO GRAVADA:", CALIBRACAO_FILE)
    for tam, c in resultado.items():
        extra = " (herdado, sem fotos prontas)" if c.get("herdado") else ""
        print(
            f"  {tam}: centro_x={c['centro_x']}  olhos_y={c['olhos_y']}  "
            f"altura_rosto={c['altura_rosto']}  amostras={c['amostras']}{extra}"
        )
    return resultado


def carregar_calibracao(cfg):
    if not CALIBRACAO_FILE.is_file():
        print("Sem calibração ainda: calibrando com as fotos prontas...")
        return calibrar(cfg)
    return json.loads(CALIBRACAO_FILE.read_text(encoding="utf-8"))


# ============================================================
# RECORTE
# ============================================================

def calcular_recorte(img_size, rosto, molde, calib):
    """Retorna (left, top, right, bottom) na imagem original."""
    W, H = img_size
    mw, mh = molde
    if not rosto:
        return ajustar_caixa((0, 0, W, H), img_size, molde)

    crop_h = rosto["h"] / calib["altura_rosto"]
    crop_w = crop_h * mw / mh

    # Não cabe na foto: reduz mantendo a proporção do molde
    fator = min(1.0, W / crop_w, H / crop_h)
    crop_w *= fator
    crop_h *= fator

    left = rosto["cx"] - calib["centro_x"] * crop_w
    top = rosto["olhos_y"] - calib["olhos_y"] * crop_h

    left = min(max(0.0, left), W - crop_w)
    top = min(max(0.0, top), H - crop_h)

    return (
        int(round(left)),
        int(round(top)),
        int(round(left + crop_w)),
        int(round(top + crop_h)),
    )


# Os smart objects das capas são 300 dpi; foto com outro dpi entra ampliada no Photoshop.
DPI_MOLDE = (300, 300)


def salvar_jpg(img, destino, icc, dpi):
    opcoes = {"quality": 95, "subsampling": 0, "optimize": True}
    if icc:
        opcoes["icc_profile"] = icc
    if dpi:
        opcoes["dpi"] = dpi
    temp = destino.with_name(destino.stem + ".recorte_tmp.jpg")
    img.save(temp, "JPEG", **opcoes)
    os.replace(temp, destino)


def caixa_automatica(img, tam, cfg, calib, caminho=None):
    """(caixa, rosto) pelo rosto detectado (com cache persistente se caminho informado); (None, None) sem rosto."""
    rosto = detectar_rosto_cache(caminho, img) if caminho else detectar_rosto(img)
    if not rosto:
        return None, None
    return calcular_recorte(img.size, rosto, cfg["moldes"][tam], calib[tam]), rosto


def ampliacao_da_caixa(caixa, molde):
    return molde[0] / max(1, caixa[2] - caixa[0])


def recortar_imagem(img, caixa, molde):
    recortada = img.crop(caixa).resize(molde, Image.LANCZOS)
    if recortada.mode != "RGB":
        recortada = recortada.convert("RGB")
    return recortada


def aplicar_recorte(caminho, caixa, tam, cfg, fonte=None):
    """
    Grava a foto da pasta FOTO no tamanho exato do molde.
    `caixa` está em pixels da foto original (fonte_original).
    O original vai para _ORIGINAIS antes, se ainda não estiver lá.
    Retorna o caminho gravado.
    """
    origem = fonte_original(caminho, cfg)
    if fonte is None:
        fonte = abrir_foto(origem)
    img, icc, _ = fonte

    if backup_de(caminho, cfg) is None:
        cfg["backup"].mkdir(parents=True, exist_ok=True)
        shutil.copy2(caminho, cfg["backup"] / caminho.name)

    recortada = recortar_imagem(img, caixa, cfg["moldes"][tam])
    destino = caminho if caminho.suffix.lower() in (".jpg", ".jpeg") else caminho.with_suffix(".jpg")
    salvar_jpg(recortada, destino, icc, DPI_MOLDE)
    if destino != caminho and caminho.exists():
        caminho.unlink()
    return destino


def restaurar_original(caminho, cfg):
    """Volta o original de _ORIGINAIS para a pasta FOTO. Retorna o caminho restaurado."""
    backup = backup_de(caminho, cfg)
    if backup is None:
        raise FileNotFoundError("Sem original guardado em " + str(cfg["backup"]))
    destino = caminho.with_name(backup.name)
    temp = destino.with_name(destino.stem + ".restaura_tmp" + destino.suffix)
    shutil.copy2(backup, temp)
    os.replace(temp, destino)
    if destino != caminho and caminho.exists():
        caminho.unlink()
    apagar_manual(caminho)
    return destino


def salvar_preview(img, caixa, recortada, destino):
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    escala_a = min(1.0, 700.0 / max(1, max(img.size)))
    antes = img.resize(
        (max(1, round(img.width * escala_a)), max(1, round(img.height * escala_a))),
        Image.BILINEAR,
    ).convert("RGB")
    fx = antes.width / max(1, img.width)
    d = ImageDraw.Draw(antes)
    d.rectangle([c * fx for c in caixa], outline=(255, 0, 0), width=4)
    escala_d = min(1.0, 700.0 / max(1, max(recortada.size)))
    depois = recortada.resize(
        (max(1, round(recortada.width * escala_d)), max(1, round(recortada.height * escala_d))),
        Image.BILINEAR,
    ).convert("RGB")
    folha = Image.new("RGB", (antes.width + depois.width + 20, max(antes.height, depois.height)), "white")
    folha.paste(antes, (0, 0))
    folha.paste(depois, (antes.width + 20, 0))
    folha.save(destino, "JPEG", quality=85)


def processar(cfg, calib, simular, filtro_aluno, filtro_tamanho):
    linhas = []
    ok = skip = erro = 0

    if not simular:
        cfg["backup"].mkdir(parents=True, exist_ok=True)

    verificar_cache_analise_subpastas(cfg["fotos"])
    manual_map = ler_manual()
    fotos_lista = []
    for f in listar_fotos(cfg["fotos"]):
        tam = tamanho_do_arquivo(f)
        if filtro_tamanho and tam != filtro_tamanho:
            continue
        if filtro_aluno and chave(filtro_aluno) not in chave(f.stem):
            continue
        fotos_lista.append((f, tam))

    def _processar_uma(item):
        f, tam = item
        molde = cfg["moldes"].get(tam)
        registro = {"arquivo": f.name, "tamanho": tam, "antes": "", "depois": "", "status": "", "motivo": ""}
        msgs = []
        try:
            atual = tamanho_rapido(f)
            registro["antes"] = f"{atual[0]}x{atual[1]}"

            if no_molde(atual, molde):
                registro.update(status="SKIP", motivo="já no tamanho do molde")
                return registro, "SKIP", msgs

            orig_path = fonte_original(f, cfg)
            reg_man = manual_map.get(f.stem)
            rosto = None
            if not (reg_man and reg_man.get("tamanho") == tam):
                # Detecta rosto antes de abrir a imagem 24MP completa (usa cache 0ms ou C++ 640px ~6ms)
                rosto = detectar_rosto_cache(orig_path)
                if not rosto:
                    registro.update(status="ERRO", motivo="rosto não encontrado")
                    msgs.append(f"\n{tam}  {f.name}\n  ERRO: rosto não encontrado")
                    return registro, "ERRO", msgs

            fonte = abrir_foto(orig_path)
            img = fonte[0]
            msgs.append(f"\n{tam}  {f.name}  ({img.width}x{img.height})")

            if reg_man and reg_man.get("tamanho") == tam:
                W, H = img.size
                l, t, r, b = reg_man["caixa"]
                caixa = ajustar_caixa((l * W, t * H, r * W, b * H), img.size, molde)
                metodo = "corte manual"
            else:
                caixa = calcular_recorte(img.size, rosto, molde, calib[tam])
                metodo = "rosto " + rosto["metodo"]

            ampliacao = ampliacao_da_caixa(caixa, molde)
            if ampliacao > cfg["ampliacao_maxima"]:
                registro.update(
                    status="ERRO",
                    motivo=f"resolução baixa: precisaria ampliar {ampliacao:.0%}",
                )
                msgs.append("  ERRO: " + registro["motivo"])
                return registro, "ERRO", msgs

            recortada = recortar_imagem(img, caixa, molde)
            registro["depois"] = f"{molde[0]}x{molde[1]}"

            salvar_preview(img, caixa, recortada, PREVIEW_DIR / (f.stem + ".jpg"))

            if simular:
                registro.update(status="SIMULADO", motivo=f"{metodo}, escala {ampliacao:.0%}")
                msgs.append("  SIMULADO: " + registro["motivo"])
            else:
                aplicar_recorte(f, caixa, tam, cfg, fonte)
                registro.update(status="OK", motivo=f"{metodo}, escala {ampliacao:.0%}")
                msgs.append("  OK: " + registro["motivo"])
            return registro, "OK", msgs
        except Exception as e:
            registro.update(status="ERRO", motivo=f"{type(e).__name__}: {e}")
            msgs.append(f"\n{f.name}\n  ERRO: {registro['motivo']}")
            return registro, "ERRO", msgs

    num_workers = min(8, max(2, os.cpu_count() or 4))
    with concurrent.futures.ThreadPoolExecutor(max_workers=num_workers) as pool:
        for registro, tipo_res, msgs in pool.map(_processar_uma, fotos_lista):
            for m in msgs:
                print(m)
            if tipo_res == "SKIP":
                skip += 1
            elif tipo_res == "OK":
                ok += 1
            else:
                erro += 1
            linhas.append(registro)

    _salvar_cache_rostos()

    with open(RELATORIO_CSV, "w", newline="", encoding="utf-8-sig") as arq:
        w = csv.DictWriter(arq, fieldnames=list(linhas[0].keys()) if linhas else ["arquivo"], delimiter=";")
        w.writeheader()
        w.writerows(linhas)

    print()
    print("RELATORIO:", RELATORIO_CSV)
    if ok:
        print("MINIATURAS:", PREVIEW_DIR)
    print()
    print("=" * 75)
    print(f"RESUMO RECORTE — OK: {ok} | SKIP: {skip} | ERRO: {erro}")
    print("=" * 75)



# ============================================================
# PRÉVIA NA CAPA (lida do PSD com psd_tools)
# ============================================================

def _camada_foto(psd, caminho):
    """Mesma regra de encontrarCaminho no JSX: grupos pelo nome, depois a FOTO de cima."""
    alvo = chave(caminho[-1])
    container = psd
    for nome in caminho[:-1]:
        grupos = [l for l in container if l.is_group() and chave(l.name) == chave(nome)]
        if not grupos:
            break
        container = grupos[-1]
    candidatas = [
        l for l in container.descendants()
        if not l.is_group() and chave(l.name) == alvo and l.kind == "smartobject"
    ]
    if not candidatas:
        raise RuntimeError("camada FOTO não encontrada: " + " > ".join(caminho))
    return candidatas[-1]


def _tamanho_smart_object(camada):
    try:
        dados = camada.smart_object.data
        if dados[:4] == b"8BPS":
            altura = int.from_bytes(dados[14:18], "big")
            largura = int.from_bytes(dados[18:22], "big")
            return [largura, altura]
        with Image.open(io.BytesIO(dados)) as img:
            img = ImageOps.exif_transpose(img)
            return [img.width, img.height]
    except Exception:
        return None


def gerar_previa_capa(cfg, tamanhos=None):
    """
    Para cada tamanho grava em previa_capas:
      <TAM>_baixo.png  camadas abaixo da FOTO
      <TAM>_cima.png   camadas acima da FOTO (transparente)
      <TAM>.json       posição da FOTO na capa reduzida e tamanho do smart object
    """
    from psd_tools import PSDImage

    PREVIA_CAPAS_DIR.mkdir(parents=True, exist_ok=True)
    feitos = []

    for tam, (arquivo, caminho) in PSD_POR_TAMANHO.items():
        if tamanhos and tam not in tamanhos:
            continue
        psd_path = cfg["psd_dir"] / arquivo
        if not psd_path.is_file():
            print(f"  {tam}: PSD não encontrado: {psd_path}")
            continue

        print(f"\n{tam}: abrindo {arquivo} ...")
        psd = PSDImage.open(psd_path)
        foto = _camada_foto(psd, caminho)

        ordem = {id(l): i for i, l in enumerate(psd.descendants())}
        indice_foto = ordem[id(foto)]
        textos_nome = _camadas_do_nome(psd, tam)

        # Prancheta (24X34): a capa é só a área da prancheta
        area = (0, 0, psd.width, psd.height)
        pai = foto.parent
        while pai is not None and pai is not psd:
            if pai.kind == "artboard":
                area = pai.bbox
            pai = pai.parent

        def filtro(abaixo, so_simples=False):
            def f(camada):
                if not camada.is_visible():
                    return False
                if camada.is_group():
                    return True
                if id(camada) in textos_nome:
                    return False
                if so_simples and camada.kind not in CAMADAS_SIMPLES:
                    return False
                i = ordem.get(id(camada), -1)
                return i < indice_foto if abaixo else i > indice_foto
            return f

        def compor(abaixo, **extra):
            try:
                return psd.composite(viewport=area, force=True, layer_filter=filtro(abaixo, not abaixo), **extra)
            except ImportError:
                # ajustes/gradientes pedem scipy: monta sem eles
                return psd.composite(viewport=area, force=True, layer_filter=filtro(abaixo, True), **extra)

        print(f"  {tam}: montando a capa sem a foto (pode demorar)...")
        baixo = compor(True)
        cima = compor(False, color=0.0, alpha=0.0)

        largura, altura = area[2] - area[0], area[3] - area[1]
        escala = min(1.0, LADO_PREVIA_CAPA / max(largura, altura))
        tamanho = (max(1, round(largura * escala)), max(1, round(altura * escala)))
        baixo.convert("RGB").resize(tamanho, Image.LANCZOS).save(PREVIA_CAPAS_DIR / f"{tam}_baixo.png")
        cima.convert("RGBA").resize(tamanho, Image.LANCZOS).save(PREVIA_CAPAS_DIR / f"{tam}_cima.png")

        l, t, r, b = foto.bbox
        visivel = (max(l, area[0]), max(t, area[1]), min(r, area[2]), min(b, area[3]))
        origem = (area[0], area[1], area[0], area[1])

        def reduzir(caixa):
            return [round((c - o) * escala, 2) for c, o in zip(caixa, origem)]

        info = {
            "psd": arquivo,
            "capa": [largura, altura],
            "escala": escala,
            "foto": reduzir((l, t, r, b)),
            "foto_visivel": reduzir(visivel),
            "smart_object": _tamanho_smart_object(foto),
            "sem_nome": True,
        }
        (PREVIA_CAPAS_DIR / f"{tam}.json").write_text(json.dumps(info, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"  {tam}: OK  foto {foto.bbox}  smart object {info['smart_object']}")
        feitos.append(tam)

    print()
    print("PRÉVIAS DA CAPA:", PREVIA_CAPAS_DIR, "|", ", ".join(feitos) or "nenhuma")
    return feitos


def ler_previa_capa(tam):
    """(info, baixo, cima) da prévia gerada, ou None."""
    j = PREVIA_CAPAS_DIR / f"{tam}.json"
    b = PREVIA_CAPAS_DIR / f"{tam}_baixo.png"
    c = PREVIA_CAPAS_DIR / f"{tam}_cima.png"
    if not (j.is_file() and b.is_file() and c.is_file()):
        return None
    info = json.loads(j.read_text(encoding="utf-8"))
    if not info.get("sem_nome"):
        return None
    return info, Image.open(b).convert("RGB"), Image.open(c).convert("RGBA")


def montar_previa_capa(recortada, previa, nome=None, tamanho=None, ajuste=None, rapido=False):
    """Encaixa a foto já recortada no lugar da camada FOTO da capa reduzida."""
    info, baixo, cima = previa
    l, t, r, b = info["foto"]
    capa = baixo.copy().convert("RGBA")
    filtro_redim = Image.BILINEAR if rapido else Image.LANCZOS
    foto = recortada.convert("RGB").resize((max(1, round(r - l)), max(1, round(b - t))), filtro_redim)
    vl, vt, vr, vb = (round(v) for v in info["foto_visivel"])
    fl, ft = round(l), round(t)
    cl0 = max(vl, fl)
    ct0 = max(vt, ft)
    cr0 = min(vr, fl + foto.width)
    cb0 = min(vb, ft + foto.height)
    if cr0 > cl0 and cb0 > ct0:
        recorte_vis = foto.crop((cl0 - fl, ct0 - ft, cr0 - fl, cb0 - ft))
        capa.paste(recorte_vis, (cl0, ct0))
    capa.alpha_composite(cima.convert("RGBA"))
    if nome and tamanho:
        capa = desenhar_nome(capa, tamanho, nome, ajuste)
    return capa.convert("RGB")


def montar_previa_sintetica(recortada, tamanho, nome=None, curso="", cor=None, ajuste=None):
    """
    Gera uma prévia sintética rápida e elegante da capa quando o PSD base ainda não foi gerado.
    Garante que a aba Revisão e Recorte sempre mostre a capa do aluno com foto, cor do curso e nome.
    """
    tam = str(tamanho or "24X34").upper()
    if tam == "20X20":
        w, h = 920, 620
        box = (470, 55, 865, 450)
        area_nome = (470, 465, 865, 585)
    elif tam == "30X20":
        w, h = 980, 560
        box = (520, 42, 930, 395)
        area_nome = (520, 410, 930, 530)
    elif tam == "30X22":
        w, h = 980, 600
        box = (510, 45, 930, 420)
        area_nome = (510, 435, 930, 565)
    else:
        w, h = 760, 940
        box = (155, 110, 605, 650)
        area_nome = (90, 685, 670, 880)

    rgb_fundo = tuple(int(c) for c in cor) if (isinstance(cor, (list, tuple)) and len(cor) == 3) else (24, 34, 56)
    capa = Image.new("RGB", (w, h), rgb_fundo)
    draw = ImageDraw.Draw(capa)

    # Moldura dourada sutil e lombada simulada
    tom_escuro = tuple(max(0, c - 18) for c in rgb_fundo)
    if tam in ("20X20", "30X20", "30X22"):
        meio_x = w // 2 - 20
        draw.rectangle([0, 0, meio_x, h], fill=tom_escuro)
        draw.line([(meio_x, 0), (meio_x, h)], fill=(212, 175, 95), width=2)
    draw.rectangle([16, 16, w - 17, h - 17], outline=(212, 175, 95), width=2)

    # Foto recortada encaixada
    l, t, r, b = box
    fw, fh = max(1, r - l), max(1, b - t)
    foto_redim = recortada.convert("RGB").resize((fw, fh), Image.BILINEAR)
    capa.paste(foto_redim, (l, t))
    draw.rectangle([l - 2, t - 2, r + 1, b + 1], outline=(212, 175, 95), width=2)

    # Curso e formato no topo
    fonte_topo = _fonte_pil(None, 16)
    rotulo_topo = f"{tam}  •  {curso}" if curso else tam
    draw.text((32, 28), rotulo_topo, fill=(212, 175, 95), font=fonte_topo)

    # Nome do aluno em 2 linhas com corte inteligente
    if nome:
        layout = calcular_layout_nome(tam, nome, ajuste)
        nl1 = layout.get("linha1", "").upper()
        nl2 = layout.get("linha2", "").upper()
        nx0, ny0, nx1, ny1 = area_nome
        cx = (nx0 + nx1) // 2
        f1 = _fonte_pil(None, max(16, round(28 * layout.get("escala1", 100) / 100.0)))
        f2 = _fonte_pil(None, max(14, round(24 * layout.get("escala2", 100) / 100.0)))
        if nl1:
            draw.text((cx, ny0 + 28), nl1, fill=(232, 198, 116), font=f1, anchor="mm")
        if nl2:
            draw.text((cx, ny0 + 70), nl2, fill=(212, 175, 95), font=f2, anchor="mm")
    return capa


def _camada_por_caminho(psd, caminho):

    atual = psd
    for indice, nome in enumerate(caminho):
        correspondencias = [camada for camada in atual if camada.name == nome]
        if not correspondencias:
            raise ValueError(
                "Caminho de layer ausente: " + " > ".join(caminho)
            )
        if indice < len(caminho) - 1:
            grupos = [camada for camada in correspondencias if camada.is_group()]
            if len(grupos) != 1:
                raise ValueError(
                    "Grupo de layer ausente ou ambíguo: " + " > ".join(caminho[:indice + 1])
                )
            atual = grupos[0]
        else:
            objetos_foto = [camada for camada in correspondencias if camada.kind == "smartobject"]
            candidatas = objetos_foto or [camada for camada in correspondencias if not camada.is_group()]
            if not candidatas:
                candidatas = correspondencias
            atual = candidatas[-1]
    return atual


def _area_previa(psd, foto):

    area = (0, 0, psd.width, psd.height)
    pai = foto.parent
    while pai is not None and pai is not psd:
        if pai.kind == "artboard":
            area = pai.bbox
        pai = pai.parent
    return area


def _cor_atual_fill(camada):

    for chave, bloco in camada._record.tagged_blocks.items():
        if "SOLID_COLOR_SHEET_SETTING" not in str(chave):
            continue
        cor = bloco.data.get(b"Clr ")
        if cor is None:
            continue
        return tuple(round(cor[campo]) for campo in (b"Rd  ", b"Grn ", b"Bl  "))
    return None


def _recolorir_previa_fundo(imagem, cor_original, cor_nova):

    pixels = np.asarray(imagem.convert("RGBA"), dtype=np.int16).copy()
    diferenca = np.asarray(cor_nova, dtype=np.int16) - np.asarray(cor_original, dtype=np.int16)
    fundo = np.any(pixels[:, :, :3] < 245, axis=2)
    pixels[fundo, :3] = np.clip(pixels[fundo, :3] + diferenca, 0, 255)
    return Image.fromarray(pixels.astype(np.uint8), "RGBA")


def _compor_brasao_no_canvas(psd, caminho_frente, simbolo, area, escala, tamanho):

    camada_frente = _camada_por_caminho(psd, caminho_frente)
    if camada_frente.kind != "smartobject":
        raise ValueError("A layer FRENTE não é um Smart Object.")

    from psd_tools import PSDImage

    smart_object = PSDImage.open(io.BytesIO(camada_frente.smart_object.data))
    grupos_curso = {"ABC", "DEF", "GHJ", "MNOP", "RST"}
    grupos = [
        camada for camada in smart_object.descendants()
        if camada.is_group() and camada.name.upper() in grupos_curso
    ]
    for grupo in grupos:
        for filha in grupo:
            filha.visible = False

    chave = _chave_brasao(simbolo)
    correspondencias = [
        camada for camada in smart_object.descendants()
        if not camada.is_group() and _chave_brasao(camada.name) == chave
    ]
    if len(correspondencias) != 1:
        raise ValueError(
            "Brasão ausente ou duplicado no Smart Object: " + str(simbolo)
        )

    brasao = correspondencias[0]
    camada = brasao
    while camada is not None and camada is not smart_object:
        camada.visible = True
        camada = camada.parent

    interna = smart_object.composite(force=True)
    if interna is None:
        raise ValueError("Não foi possível rasterizar o brasão do curso.")

    caixa = tuple(float(valor) for valor in camada_frente.smart_object.transform_box)
    if len(caixa) != 8:
        raise ValueError("Transformação FRENTE inesperada no PSD.")
    x0, y0, x1, y1, x2, y2, x3, y3 = caixa
    tolerancia = 1.0
    if (abs(y0 - y1) > tolerancia or abs(x1 - x2) > tolerancia
            or abs(y2 - y3) > tolerancia or abs(x3 - x0) > tolerancia):
        raise ValueError("Transformação FRENTE inclinada não suportada pela prévia.")

    left, right = min(x0, x1, x2, x3), max(x0, x1, x2, x3)
    top, bottom = min(y0, y1, y2, y3), max(y0, y1, y2, y3)
    largura = max(1, round((right - left) * escala))
    altura = max(1, round((bottom - top) * escala))
    interna = interna.convert("RGBA").resize((largura, altura), Image.LANCZOS)

    x = round((left - area[0]) * escala)
    y = round((top - area[1]) * escala)
    frente = Image.new("RGBA", tamanho, (0, 0, 0, 0))
    frente.alpha_composite(interna, (x, y))
    return frente


def gerar_preset_curso(tamanho, simbolo, rgb=None, _psd=None):
    """Gera/abre o preset offline de um formato e brasão, sem salvar o PSD."""

    if tamanho not in PSD_POR_TAMANHO or tamanho not in FRENTE_POR_TAMANHO:
        raise ValueError("Formato de capa sem preset: " + str(tamanho))
    simbolo = str(simbolo or "").strip()
    if not simbolo:
        raise ValueError("O curso não tem símbolo de brasão.")
    if rgb is not None:
        rgb = tuple(int(valor) for valor in rgb)
        if len(rgb) != 3 or any(valor < 0 or valor > 255 for valor in rgb):
            raise ValueError("A cor do curso deve conter três canais RGB válidos.")

    try:
        config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        caminho_psd = Path(config["psd_dir"]) / PSD_POR_TAMANHO[tamanho][0]
        mtime_psd = caminho_psd.stat().st_mtime_ns if caminho_psd.is_file() else None
    except (OSError, KeyError, json.JSONDecodeError):
        caminho_psd, mtime_psd = None, None

    tamanho_cor = "semcor" if rgb is None else "%02x%02x%02x" % rgb
    slug = re.sub(r"[^a-z0-9_-]+", "_", _chave_brasao(simbolo)).strip("_")
    diretorio = PREVIA_CAPAS_DIR / "presets"
    diretorio_base = diretorio / "base"
    diretorio_brasoes = diretorio / "brasoes"
    chave_cache = tamanho + "_" + slug + "_" + tamanho_cor
    arquivo_demo = diretorio / (chave_cache + ".png")
    arquivo_info = diretorio / (chave_cache + ".json")
    chave_base = tamanho + "_" + tamanho_cor
    arquivos_base = {
        nome: diretorio_base / (chave_base + "_" + nome + ".png")
        for nome in ("baixo", "cima")
    }
    arquivo_base_info = diretorio_base / (chave_base + ".json")
    arquivo_frente = _arquivo_brasoes_preset(tamanho, slug)

    if (arquivo_info.is_file() and arquivo_demo.is_file() and arquivo_frente.is_file()
            and all(caminho.is_file() for caminho in arquivos_base.values())):
        try:
            info = json.loads(arquivo_info.read_text(encoding="utf-8"))
            if ((mtime_psd is None or info.get("psd_mtime_ns") == mtime_psd)
                    and info.get("versao") == VERSAO_PRESET_CURSO
                    and info.get("rgb") == (list(rgb) if rgb is not None else None)):
                resultado = dict(info)
                resultado["demo"] = Image.open(arquivo_demo).convert("RGBA")
                resultado["frente"] = Image.open(arquivo_frente).convert("RGBA")
                for nome, caminho in arquivos_base.items():
                    resultado[nome] = Image.open(caminho).convert("RGBA")
                resultado["meio"] = Image.new(
                    "RGBA", resultado["baixo"].size, (0, 0, 0, 0),
                )
                resultado["verso"] = _verso_preset(tamanho, simbolo, rgb, resultado["baixo"].height)
                return resultado
        except (OSError, json.JSONDecodeError):
            pass

    if caminho_psd is None or not caminho_psd.is_file():
        raise FileNotFoundError("PSD não encontrado: " + str(caminho_psd or tamanho))

    from psd_tools import PSDImage

    psd = _psd if _psd is not None else PSDImage.open(caminho_psd)
    foto = _camada_por_caminho(psd, PSD_POR_TAMANHO[tamanho][1])
    frente = _camada_por_caminho(psd, FRENTE_POR_TAMANHO[tamanho])
    ordem = {id(camada): indice for indice, camada in enumerate(psd.descendants())}
    indice_foto = ordem[id(foto)]
    indice_frente = ordem[id(frente)]
    area = _area_previa(psd, foto)
    largura, altura = area[2] - area[0], area[3] - area[1]
    escala = min(1.0, LADO_PREVIA_CAPA / max(largura, altura))
    tamanho_canvas = (max(1, round(largura * escala)), max(1, round(altura * escala)))

    if (arquivo_base_info.is_file()
            and all(caminho.is_file() for caminho in arquivos_base.values())):
        try:
            info_base = json.loads(arquivo_base_info.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            info_base = {}
    else:
        info_base = {}

    if (info_base.get("psd_mtime_ns") == mtime_psd
            and info_base.get("versao") == VERSAO_PRESET_CURSO
            and info_base.get("rgb") == (list(rgb) if rgb is not None else None)):
        baixo = Image.open(arquivos_base["baixo"]).convert("RGBA")
        cima = Image.open(arquivos_base["cima"]).convert("RGBA")
    else:
        previa_demo = ler_previa_capa(tamanho)
        if not previa_demo:
            raise FileNotFoundError("Gere primeiro a prévia-demo de " + tamanho + ".")
        info_demo, baixo, cima = previa_demo
        baixo = baixo.convert("RGBA")
        cima = cima.convert("RGBA")

        caminho_fill = PREENCHIMENTO_POR_TAMANHO.get(tamanho)
        if rgb is not None and caminho_fill:
            fill = _camada_por_caminho(psd, caminho_fill)
            cor_original = _cor_atual_fill(fill)
            if fill.kind != "solidcolorfill" or cor_original is None:
                raise ValueError("Não foi possível ler o preenchimento do preset.")
            baixo = _recolorir_previa_fundo(baixo, cor_original, rgb)

        info_base = {
            "tamanho": tamanho,
            "rgb": list(rgb) if rgb is not None else None,
            "foto": info_demo["foto"],
            "foto_visivel": info_demo["foto_visivel"],
            "frente_depois_foto": indice_frente > indice_foto,
            "escala": info_demo["escala"],
            "psd_mtime_ns": mtime_psd,
            "versao": VERSAO_PRESET_CURSO,
        }
        diretorio_base.mkdir(parents=True, exist_ok=True)
        for nome, imagem in (("baixo", baixo), ("cima", cima)):
            imagem.save(arquivos_base[nome], optimize=True)
        arquivo_base_info.write_text(
            json.dumps(info_base, indent=2, ensure_ascii=False), encoding="utf-8",
        )
    info = info_base
    escala = info["escala"]
    tamanho_canvas = baixo.size

    if arquivo_frente.is_file() and arquivo_frente.stat().st_mtime_ns >= mtime_psd:
        camada_frente = Image.open(arquivo_frente).convert("RGBA")
    else:
        camada_frente = Image.new("RGBA", tamanho_canvas, (0, 0, 0, 0))
        for caminho_brasao in BRASOES_POR_TAMANHO.get(tamanho, [FRENTE_POR_TAMANHO[tamanho]]):
            camada_frente.alpha_composite(_compor_brasao_no_canvas(
                psd, caminho_brasao, simbolo, area, escala, tamanho_canvas,
            ))
        diretorio_brasoes.mkdir(parents=True, exist_ok=True)
        camada_frente.save(arquivo_frente, optimize=True)

    meio = Image.new("RGBA", tamanho_canvas, (0, 0, 0, 0))
    demo = baixo.copy()
    if info["frente_depois_foto"]:
        for camada in (meio, cima, camada_frente):
            demo.alpha_composite(camada)
    else:
        for camada in (camada_frente, meio, cima):
            demo.alpha_composite(camada)

    resultado = {
        "tamanho": tamanho,
        "simbolo": simbolo,
        "rgb": list(rgb) if rgb is not None else None,
        "foto": info["foto"],
        "foto_visivel": info["foto_visivel"],
        "frente_depois_foto": info["frente_depois_foto"],
        "psd_mtime_ns": mtime_psd,
        "versao": VERSAO_PRESET_CURSO,
        "preset_png": str(arquivo_demo),
        "baixo": baixo,
        "meio": meio,
        "frente": camada_frente,
        "cima": cima,
        "demo": demo,
        "verso": _verso_preset(tamanho, simbolo, rgb, baixo.height),
    }
    diretorio.mkdir(parents=True, exist_ok=True)
    demo.save(arquivo_demo, optimize=True)
    arquivo_info.write_text(json.dumps({
        "tamanho": tamanho,
        "simbolo": simbolo,
        "rgb": list(rgb) if rgb is not None else None,
        "foto": info["foto"],
        "foto_visivel": info["foto_visivel"],
        "frente_depois_foto": info["frente_depois_foto"],
        "psd_mtime_ns": mtime_psd,
        "versao": VERSAO_PRESET_CURSO,
        "preset_png": str(arquivo_demo),
    }, indent=2, ensure_ascii=False), encoding="utf-8")
    return resultado


def _arquivo_brasoes_preset(tamanho, slug):
    return PREVIA_CAPAS_DIR / "presets" / "brasoes" / (
        tamanho + "_" + slug + "_v" + str(VERSAO_PRESET_CURSO) + ".png")


def _arquivo_verso_preset(tamanho, slug, rgb):
    cor = "semcor" if rgb is None else "%02x%02x%02x" % tuple(rgb)
    return PREVIA_CAPAS_DIR / "presets" / "verso" / (
        tamanho + "_" + slug + "_" + cor + "_v" + str(VERSAO_PRESET_CURSO) + ".png")


def _verso_preset(tamanho, simbolo, rgb, altura):
    """Verso em PSD separado (30X22): fundo na cor do curso + brasão, na altura da prévia."""

    dados = VERSO_SEPARADO.get(tamanho)
    if not dados:
        return None
    try:
        config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        caminho_psd = Path(config["psd_dir"]) / dados["psd"]
        mtime_psd = caminho_psd.stat().st_mtime_ns if caminho_psd.is_file() else None
    except (OSError, KeyError, json.JSONDecodeError):
        caminho_psd, mtime_psd = None, None

    slug = re.sub(r"[^a-z0-9_-]+", "_", _chave_brasao(simbolo)).strip("_")
    arquivo = _arquivo_verso_preset(tamanho, slug, rgb)
    if arquivo.is_file() and (mtime_psd is None or arquivo.stat().st_mtime_ns >= mtime_psd):
        imagem = Image.open(arquivo).convert("RGBA")
        if imagem.height == altura:
            return imagem
    if caminho_psd is None or not caminho_psd.is_file():
        return None

    from psd_tools import PSDImage

    psd = PSDImage.open(caminho_psd)
    area = (0, 0, psd.width, psd.height)
    escala = altura / psd.height
    tamanho_canvas = (max(1, round(psd.width * escala)), altura)

    def visivel(camada):
        return camada.is_visible()

    def visivel_simples(camada):
        return camada.is_visible() and (camada.is_group() or camada.kind in CAMADAS_SIMPLES)

    try:
        fundo = psd.composite(viewport=area, force=True, layer_filter=visivel)
    except ImportError:
        fundo = psd.composite(viewport=area, force=True, layer_filter=visivel_simples)
    fundo = fundo.convert("RGBA").resize(tamanho_canvas, Image.LANCZOS)

    if rgb is not None:
        fill = _camada_por_caminho(psd, dados["preenchimento"])
        cor_original = _cor_atual_fill(fill)
        if cor_original is not None:
            fundo = _recolorir_previa_fundo(fundo, cor_original, rgb)

    fundo.alpha_composite(_compor_brasao_no_canvas(
        psd, dados["brasao"], simbolo, area, escala, tamanho_canvas,
    ))
    arquivo.parent.mkdir(parents=True, exist_ok=True)
    fundo.save(arquivo, optimize=True)
    return fundo


def gerar_presets_curso(tamanho, cursos_rgb):
    """Gera/reusa presets em lote, abrindo cada PSD no máximo uma vez."""

    if not cursos_rgb:
        return {}
    try:
        config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        caminho_psd = Path(config["psd_dir"]) / PSD_POR_TAMANHO[tamanho][0]
        mtime_psd = caminho_psd.stat().st_mtime_ns if caminho_psd.is_file() else None
    except (OSError, KeyError, json.JSONDecodeError):
        caminho_psd, mtime_psd = None, None

    faltantes = {}
    resultado = {}
    diretorio = PREVIA_CAPAS_DIR / "presets"
    for simbolo, rgb in cursos_rgb.items():
        cor = tuple(rgb) if rgb is not None else None
        slug = re.sub(r"[^a-z0-9_-]+", "_", _chave_brasao(simbolo)).strip("_")
        cor_slug = "semcor" if cor is None else "%02x%02x%02x" % cor
        chave_cache = tamanho + "_" + slug + "_" + cor_slug
        arquivos = [
            diretorio / (chave_cache + ".png"),
            diretorio / (chave_cache + ".json"),
            _arquivo_brasoes_preset(tamanho, slug),
        ]
        base_arquivos = [
            diretorio / "base" / (tamanho + "_" + cor_slug + "_" + nome + ".png")
            for nome in ("baixo", "cima")
        ]
        try:
            info = json.loads(arquivos[1].read_text(encoding="utf-8"))
            cache_valido = (
                (mtime_psd is None or info.get("psd_mtime_ns") == mtime_psd)
                and info.get("versao") == VERSAO_PRESET_CURSO
                and info.get("rgb") == (list(cor) if cor is not None else None)
                and all(arquivo.is_file() for arquivo in arquivos + base_arquivos)
            )
        except (OSError, json.JSONDecodeError):
            cache_valido = False
        if cache_valido:
            resultado[simbolo] = gerar_preset_curso(tamanho, simbolo, cor)
        else:
            faltantes[simbolo] = cor

    if faltantes:
        if caminho_psd is None or not caminho_psd.is_file():
            raise FileNotFoundError("PSD não encontrado: " + str(caminho_psd or tamanho))
        from psd_tools import PSDImage

        psd = PSDImage.open(caminho_psd)
        for simbolo, cor in faltantes.items():
            resultado[simbolo] = gerar_preset_curso(tamanho, simbolo, cor, _psd=psd)
    return resultado


def montar_previa_preset(recortada, preset, nome=None, ajuste=None, rapido=False):
    """Insere o recorte do aluno no preset do curso, preservando a ordem das layers.

    Com nome, desenha o nome do aluno como o Photoshop (ajuste None = o salvo em ajustes_nome.json).
    """

    capa = preset["baixo"].copy().convert("RGBA")
    l, t, r, b = preset["foto"]
    filtro_redim = Image.BILINEAR if rapido else Image.LANCZOS
    foto = recortada.convert("RGB").resize(
        (max(1, round(r - l)), max(1, round(b - t))), filtro_redim,
    )
    vl, vt, vr, vb = (round(valor) for valor in preset["foto_visivel"])
    fl, ft = round(l), round(t)
    cl0 = max(vl, fl)
    ct0 = max(vt, ft)
    cr0 = min(vr, fl + foto.width)
    cb0 = min(vb, ft + foto.height)
    if cr0 > cl0 and cb0 > ct0:
        recorte_vis = foto.crop((cl0 - fl, ct0 - ft, cr0 - fl, cb0 - ft))
        capa.paste(recorte_vis, (cl0, ct0))

    sobre = preset.get("_sobre_foto")
    if sobre is None:
        sobre = Image.new("RGBA", capa.size, (0, 0, 0, 0))
        if not preset["frente_depois_foto"]:
            camadas = (preset["frente"], preset["meio"], preset["cima"])
        else:
            camadas = (preset["meio"], preset["cima"], preset["frente"])
        for imagem in camadas:
            sobre.alpha_composite(imagem.convert("RGBA"))
        preset["_sobre_foto"] = sobre
    capa.alpha_composite(sobre)

    if nome:
        capa = desenhar_nome(capa, preset["tamanho"], nome, ajuste)

    verso = preset.get("verso")
    if verso is not None:
        junta = Image.new("RGBA", (verso.width + capa.width, capa.height), (255, 255, 255, 255))
        junta.alpha_composite(verso.convert("RGBA"), (0, 0))
        junta.alpha_composite(capa, (verso.width, 0))
        capa = junta
    return capa.convert("RGB")


def preset_em_cache(tamanho, simbolo, rgb):
    """True se o preset deste curso já está pronto (abre na hora, sem ler o PSD)."""

    try:
        config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        caminho_psd = Path(config["psd_dir"]) / PSD_POR_TAMANHO[tamanho][0]
        mtime_psd = caminho_psd.stat().st_mtime_ns if caminho_psd.is_file() else None
    except (OSError, KeyError, json.JSONDecodeError):
        mtime_psd = None
    cor = tuple(rgb) if rgb is not None else None
    slug = re.sub(r"[^a-z0-9_-]+", "_", _chave_brasao(simbolo)).strip("_")
    cor_slug = "semcor" if cor is None else "%02x%02x%02x" % cor
    chave_cache = tamanho + "_" + slug + "_" + cor_slug
    diretorio = PREVIA_CAPAS_DIR / "presets"
    try:
        info = json.loads((diretorio / (chave_cache + ".json")).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False
    arquivos = [diretorio / (chave_cache + ".png"), _arquivo_brasoes_preset(tamanho, slug)]
    arquivos += [diretorio / "base" / (tamanho + "_" + cor_slug + "_" + n + ".png") for n in ("baixo", "cima")]
    if tamanho in VERSO_SEPARADO:
        arquivos.append(_arquivo_verso_preset(tamanho, slug, cor))
    return (
        (mtime_psd is None or info.get("psd_mtime_ns") == mtime_psd)
        and info.get("versao") == VERSAO_PRESET_CURSO
        and info.get("rgb") == (list(cor) if cor is not None else None)
        and all(a.is_file() for a in arquivos)
    )


# ============================================================
# NOME DO ALUNO NA PRÉVIA (mesma fonte, tamanho e âncora do PSD)
# ============================================================

def _achar_camada(container, caminho, tipo):
    """tipo: "group", "type" ou "smartobject" (na última parte do caminho)."""
    atual = container
    for indice, nome in enumerate(caminho):
        ultimo = indice == len(caminho) - 1
        candidatas = [
            camada for camada in atual
            if camada.name == nome and (
                camada.kind == tipo if ultimo and tipo != "group" else camada.is_group()
            )
        ]
        if not candidatas:
            raise ValueError("Camada do nome não encontrada: " + " > ".join(caminho))
        atual = candidatas[-1]
    return atual


def _camadas_do_nome(psd, tamanho):
    """ids das camadas com o nome de exemplo (ficam fora da prévia base)."""
    dados = NOME_POR_TAMANHO.get(tamanho) or {}
    ids = set()
    try:
        if "smart" in dados:
            ids.add(id(_achar_camada(psd, dados["smart"], "smartobject")))
        elif "grupo" in dados:
            grupo = _achar_camada(psd, dados["grupo"], "group")
            ids |= {id(c) for c in grupo.descendants() if c.kind == "type"}
        if "lombada" in dados:
            ids.add(id(_achar_camada(psd, dados["lombada"], "type")))
    except ValueError as e:
        print("  AVISO:", e)
    return ids


def _chave_fonte(texto):
    return re.sub(r"[^a-z0-9]", "", str(texto).lower())


_FONTES_POR_CHAVE = None


def _arquivo_fonte(nome_ps):
    """Arquivo .ttf/.otf instalado para o nome PostScript usado no PSD."""
    global _FONTES_POR_CHAVE
    alvo = _chave_fonte(nome_ps)
    if _FONTES_POR_CHAVE is None:
        _FONTES_POR_CHAVE = {}
        pastas = [
            Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts",
        ]
        arquivos = [a for p in pastas if p.is_dir() for a in p.iterdir() if a.suffix.lower() in (".ttf", ".otf")]
        for arquivo in arquivos:
            _FONTES_POR_CHAVE.setdefault(_chave_fonte(arquivo.stem), str(arquivo))
            try:
                familia, estilo = ImageFont.truetype(str(arquivo), 10).getname()
            except Exception:
                continue
            _FONTES_POR_CHAVE.setdefault(_chave_fonte(familia + estilo), str(arquivo))
            _FONTES_POR_CHAVE.setdefault(_chave_fonte(familia), str(arquivo))
    return _FONTES_POR_CHAVE.get(alvo) or _FONTES_POR_CHAVE.get(_chave_fonte(str(nome_ps).split("-")[0]))


@functools.lru_cache(maxsize=64)
def _fonte_pil(arquivo, px):
    if arquivo:
        try:
            return ImageFont.truetype(arquivo, max(1.0, float(px)))
        except OSError:
            pass
    return ImageFont.load_default(max(1.0, float(px)))


def _perfil_cor(imagem, rotacao, partes=24):
    """Cores de cima para baixo da letra (gradiente dourado do PSD)."""
    padrao = [[212, 175, 95]]
    if imagem is None:
        return padrao
    imagem = imagem.convert("RGBA")
    if abs(rotacao) > 0.5:
        imagem = imagem.rotate(rotacao, expand=True)
    arr = np.asarray(imagem, dtype=np.float32)
    alfa = arr[:, :, 3] / 255.0
    linhas = np.where(alfa.sum(axis=1) > 0.5)[0]
    if len(linhas) == 0:
        return padrao
    topo, base = linhas[0], linhas[-1] + 1
    cores = []
    for i in range(partes):
        a = topo + (base - topo) * i // partes
        b = max(a + 1, topo + (base - topo) * (i + 1) // partes)
        peso = alfa[a:b]
        if peso.sum() < 0.5:
            continue
        cor = (arr[a:b, :, :3] * peso[:, :, None]).sum(axis=(0, 1)) / peso.sum()
        cores.append([int(round(c)) for c in cor])
    return cores or padrao


def _medidas_texto(camada, desloc=(0.0, 0.0), escala=1.0):
    estilo = camada.engine_dict["StyleRun"]["RunArray"][0]["StyleSheet"]["StyleSheetData"]
    paragrafo = camada.engine_dict["ParagraphRun"]["RunArray"][0]["ParagraphSheet"]["Properties"]
    fontes = [str(f.get("Name")).strip("'\"") for f in camada.resource_dict.get("FontSet", [])]
    nome_fonte = fontes[int(estilo.get("Font", 0))] if fontes else ""
    xx, xy, _, _, tx, ty = (float(v) for v in camada.transform)
    escala_texto = math.hypot(xx, xy)
    rotacao = round(math.degrees(math.atan2(xy, xx)), 2)
    caps = int(estilo.get("FontCaps", 0)) == 2
    alinhamento = {0: "esquerda", 1: "direita", 2: "centro"}.get(int(paragrafo.get("Justification", 0)), "esquerda")
    arquivo = _arquivo_fonte(nome_fonte)
    px = float(estilo.get("FontSize", 12)) * escala_texto * escala
    exemplo = str(camada.text or "").strip()

    try:
        perfil = _perfil_cor(camada.composite(force=True), rotacao)
    except Exception:
        perfil = _perfil_cor(None, 0)

    return {
        "camada": camada.name,
        "fonte_ps": nome_fonte,
        "fonte": arquivo,
        "px": round(px, 3),
        "x": round(desloc[0] + tx * escala, 2),
        "y": round(desloc[1] + ty * escala, 2),
        "rotacao": rotacao,
        "caps": caps,
        "tracking": float(estilo.get("Tracking", 0)),
        "escala_h": float(estilo.get("HorizontalScale", 1.0)),
        "escala_v": float(estilo.get("VerticalScale", 1.0)),
        "alinhamento": alinhamento,
        "sombra": bool(camada.has_effects()),
        "cores": perfil,
        "exemplo": exemplo,
    }


_MODELOS_NOME = {}


def modelo_nome(tamanho):
    """Linhas do nome (1ª, 2ª) e lombada lidas do PSD, com cache por data do PSD."""
    dados = NOME_POR_TAMANHO.get(tamanho)
    if not dados or tamanho not in PSD_POR_TAMANHO:
        return None
    arquivo = PREVIA_CAPAS_DIR / "presets" / "nome" / (tamanho + ".json")
    try:
        config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        caminho_psd = Path(config["psd_dir"]) / PSD_POR_TAMANHO[tamanho][0]
        mtime_psd = caminho_psd.stat().st_mtime_ns
    except (OSError, KeyError, json.JSONDecodeError):
        caminho_psd, mtime_psd = None, None

    def valido(modelo):
        return (modelo and modelo.get("versao") == VERSAO_MODELO_NOME
                and (mtime_psd is None or modelo.get("psd_mtime_ns") == mtime_psd))

    if valido(_MODELOS_NOME.get(tamanho)):
        return _MODELOS_NOME[tamanho]
    try:
        modelo = json.loads(arquivo.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        modelo = None
    if valido(modelo):
        _MODELOS_NOME[tamanho] = modelo
        return modelo
    if caminho_psd is None or not caminho_psd.is_file():
        return modelo

    from psd_tools import PSDImage

    psd = PSDImage.open(caminho_psd)
    area = _area_previa(psd, _camada_por_caminho(psd, PSD_POR_TAMANHO[tamanho][1]))
    container, desloc, escala = psd, (0.0, 0.0), 1.0
    recorte = None
    if "smart" in dados:
        objeto = _achar_camada(psd, dados["smart"], "smartobject")
        caixa = [float(v) for v in objeto.smart_object.transform_box]
        interno = PSDImage.open(io.BytesIO(objeto.smart_object.data))
        desloc = (min(caixa[0::2]), min(caixa[1::2]))
        escala = (max(caixa[0::2]) - desloc[0]) / interno.width
        container = interno
        # o Photoshop corta o texto que sai da tela do smart object
        recorte = [
            round(desloc[0] - area[0], 2), round(desloc[1] - area[1], 2),
            round(max(caixa[0::2]) - area[0], 2), round(max(caixa[1::2]) - area[1], 2),
        ]
    grupo = _achar_camada(container, dados["grupo"], "group")
    textos = [c for c in grupo.descendants() if c.kind == "type"][::-1]
    if len(textos) < 2:
        raise ValueError("O grupo do nome precisa de 2 textos em " + tamanho)
    linhas = [_medidas_texto(t, desloc, escala) for t in textos[:2]]
    lombada = None
    if "lombada" in dados:
        lombada = _medidas_texto(_achar_camada(psd, dados["lombada"], "type"))
    for medida in linhas + ([lombada] if lombada else []):
        medida["x"] = round(medida["x"] - area[0], 2)
        medida["y"] = round(medida["y"] - area[1], 2)

    modelo = {
        "tamanho": tamanho,
        "versao": VERSAO_MODELO_NOME,
        "psd_mtime_ns": mtime_psd,
        "origem": [area[0], area[1]],
        "area": [area[2] - area[0], area[3] - area[1]],
        "linhas": linhas,
        "lombada": lombada,
        "recorte": recorte,
    }
    arquivo.parent.mkdir(parents=True, exist_ok=True)
    arquivo.write_text(json.dumps(modelo, indent=2, ensure_ascii=False), encoding="utf-8")
    _MODELOS_NOME[tamanho] = modelo
    return modelo


def palavras_nome(nome):
    limpo = limpar_tamanho_do_nome(nome)
    return [p for p in re.sub(r"\s+", " ", limpo).strip().split(" ") if p]


def _eh_preposicao(palavra):
    return palavra.upper() in PREPOSICOES_NOME


def corte_automatico_nome(palavras):
    """Quantas palavras vão na 1ª linha: mesma regra de dividirNomeDuasLinhas no JSX."""
    n = len(palavras)
    if n <= 1:
        return n
    if n == 2:
        return 1

    melhor_corte, melhor_score = -1, None
    for passada in (0, 1):
        if melhor_corte >= 1:
            break
        for corte in range(1, n):
            if passada == 0 and sum(1 for p in palavras[:corte] if not _eh_preposicao(p)) != 2:
                continue
            if _eh_preposicao(palavras[corte - 1]):
                continue
            if corte == n - 1 and _eh_preposicao(palavras[corte]):
                continue
            len1 = len(" ".join(palavras[:corte]))
            len2 = len(" ".join(palavras[corte:]))
            maior, menor = max(len1, len2), min(len1, len2)
            razao = maior / menor if menor > 0 else 99
            score = (
                abs(len1 - len2)
                + ((razao - 2.2) * 12 if razao > 2.2 else 0)
                + ((maior - 22) * 3 if maior > 22 else 0)
                + abs(corte - n / 2) * 1.5
                + (-6 if _eh_preposicao(palavras[corte]) else 0)
            )
            if melhor_score is None or score < melhor_score:
                melhor_score, melhor_corte = score, corte

    if melhor_corte < 1:
        melhor_corte = min(math.ceil(n / 2), n - 1)
        while melhor_corte > 1 and _eh_preposicao(palavras[melhor_corte - 1]):
            melhor_corte -= 1
    return melhor_corte


def _texto_da_linha(medida, texto):
    return texto.upper() if medida.get("caps") else texto


def _avanco_texto(medida, texto, px):
    """Largura do texto como o Photoshop: avanço da fonte + tracking, vezes a escala horizontal."""
    if not texto:
        return 0.0
    fonte = _fonte_pil(medida.get("fonte"), 400)
    tracking = medida.get("tracking", 0) / 1000.0 * 400 * (len(texto) - 1)
    return (fonte.getlength(texto) + tracking) * px / 400 * medida.get("escala_h", 1.0)


def _largura_texto(medida, texto, escala=100):
    """Largura (px do PSD) do texto nesta linha, na escala em %."""
    return _avanco_texto(medida, _texto_da_linha(medida, texto), medida["px"] * escala / 100)


def _limite_linha(modelo, medida, tamanho, lombada=False):
    limites = LIMITES_NOME.get(tamanho, {})
    if lombada:
        return limites.get("lombada", 0.8) * modelo["area"][1]
    if "x_min" not in limites:
        return float("inf")
    x_min = limites["x_min"] - modelo["origem"][0]
    x_max = modelo["area"][0]
    if medida["alinhamento"] == "direita":
        return medida["x"] - x_min
    if medida["alinhamento"] == "esquerda":
        return x_max - medida["x"]
    return 2 * min(medida["x"] - x_min, x_max - medida["x"])


def _escala_valida(valor):
    try:
        valor = int(round(float(valor)))
    except (TypeError, ValueError):
        return None
    return max(ESCALA_NOME_MIN, min(ESCALA_NOME_MAX, valor))


def calcular_layout_nome(tamanho, nome, ajuste=None, modelo=None):
    """
    Corte e escala (%) de cada linha do nome. Sem ajuste: corte automático e,
    se uma linha passar da área permitida, reduz só ela (reduzido_auto).
    Com ajuste {"corte", "escala1", "escala2"}: usa o que foi definido e só avisa (estoura).
    """
    palavras = palavras_nome(nome)
    n = len(palavras)
    ajuste = ajuste or {}
    corte = ajuste.get("corte")
    if not isinstance(corte, int) or not (1 <= corte < n):
        corte = corte_automatico_nome(palavras)
    linha1 = " ".join(palavras[:corte])
    linha2 = " ".join(palavras[corte:])
    layout = {
        "linha1": linha1, "linha2": linha2, "corte": corte,
        "escala1": 100, "escala2": 100, "escala_lombada": 100,
        "reduzido_auto": False, "estoura": False, "manual": bool(ajuste),
        "linhas_estouradas": [],
    }
    if modelo is None:
        try:
            modelo = modelo_nome(tamanho)
        except Exception:
            modelo = None
    if not modelo:
        return layout

    for indice, (medida, texto) in enumerate(zip(modelo["linhas"][:2], (linha1, linha2)), 1):
        largura = _largura_texto(medida, texto)
        limite = _limite_linha(modelo, medida, tamanho)
        manual = _escala_valida(ajuste.get("escala%d" % indice))
        if manual is not None:
            escala = manual
        elif largura > limite > 0:
            escala = max(ESCALA_NOME_MIN, int(limite / largura * 100))
            layout["reduzido_auto"] = True
        else:
            escala = 100
        layout["escala%d" % indice] = escala
        if largura * escala / 100 > limite + 1:
            layout["estoura"] = True
            layout["linhas_estouradas"].append(indice)

    lombada = modelo.get("lombada")
    if lombada:
        largura = _largura_texto(lombada, " ".join(palavras))
        limite = _limite_linha(modelo, lombada, tamanho, lombada=True)
        if largura > limite > 0:
            layout["escala_lombada"] = max(ESCALA_NOME_MIN, int(limite / largura * 100))
            layout["reduzido_auto"] = True
    return layout


@functools.lru_cache(maxsize=256)
def _imagem_texto_cached(fonte_arq, tracking_val, escala_h, escala_v, alinhamento, sombra, cores_tupla, rotacao_val, texto, px_round):
    medida = {
        "fonte": fonte_arq,
        "tracking": tracking_val,
        "escala_h": escala_h,
        "escala_v": escala_v,
        "alinhamento": alinhamento,
        "sombra": sombra,
        "cores": [list(c) for c in cores_tupla],
    }
    img, orig = _imagem_texto(medida, texto, px_round)
    return _girar_texto(img, orig, rotacao_val)


def _imagem_texto(medida, texto, px):
    """Texto renderizado em RGBA e o ponto de âncora (como no Photoshop) dentro da imagem."""
    fonte = _fonte_pil(medida.get("fonte"), round(px, 2))
    tracking = medida.get("tracking", 0) / 1000.0 * px
    avanco = fonte.getlength(texto) + tracking * (len(texto) - 1)
    subida, descida = fonte.getmetrics()
    margem = int(px * 0.4) + 4
    largura = int(avanco + 2 * margem + 1)
    altura = subida + descida + 2 * margem
    linha_base = margem + subida

    mascara = Image.new("L", (largura, altura), 0)
    desenho = ImageDraw.Draw(mascara)
    if abs(tracking) < 0.01:
        desenho.text((margem, linha_base), texto, font=fonte, fill=255, anchor="ls")
    else:
        for indice, caractere in enumerate(texto):
            x = fonte.getlength(texto[:indice]) + tracking * indice
            desenho.text((margem + x, linha_base), caractere, font=fonte, fill=255, anchor="ls")

    escala_h, escala_v = medida.get("escala_h", 1.0), medida.get("escala_v", 1.0)
    if abs(escala_h - 1) > 0.001 or abs(escala_v - 1) > 0.001:
        largura = max(1, round(largura * escala_h))
        altura = max(1, round(altura * escala_v))
        mascara = mascara.resize((largura, altura), Image.LANCZOS)
    margem_x = margem * escala_h
    avanco *= escala_h
    linha_base *= escala_v
    deslocamento = {"direita": avanco, "centro": avanco / 2}.get(medida["alinhamento"], 0.0)
    origem = (margem_x + deslocamento, linha_base)
    margem = int(margem * escala_v)

    cores = medida.get("cores") or [[212, 175, 95]]
    faixa = np.zeros((altura, 3), dtype=np.float32)
    topo, base = margem, max(margem + 1, altura - margem)
    for y in range(altura):
        pos = min(1.0, max(0.0, (y - topo) / max(1, base - topo))) * (len(cores) - 1)
        a = int(pos)
        b2 = min(len(cores) - 1, a + 1)
        f = pos - a
        faixa[y] = np.asarray(cores[a]) * (1 - f) + np.asarray(cores[b2]) * f
    preenchimento = Image.fromarray(
        np.repeat(faixa[:, None, :], largura, axis=1).astype(np.uint8), "RGB",
    ).convert("RGBA")

    camada = Image.new("RGBA", (largura, altura), (0, 0, 0, 0))
    if medida.get("sombra"):
        desloc = max(1, round(px * 0.03))
        sombra = Image.new("L", (largura, altura), 0)
        sombra.paste(mascara, (desloc, desloc))
        sombra = sombra.filter(ImageFilter.GaussianBlur(max(1, px * 0.03))).point(lambda v: v * 0.6)
        camada.paste((0, 0, 0, 255), (0, 0), sombra)
    camada.paste(preenchimento, (0, 0), mascara)
    return camada, origem


def _girar_texto(imagem, origem, graus):
    """Gira no sentido horário mantendo o ponto de âncora."""
    if abs(graus) < 0.5:
        return imagem, origem
    cx, cy = imagem.width / 2, imagem.height / 2
    nova = imagem.rotate(-graus, resample=Image.BICUBIC, expand=True)
    rad = math.radians(graus)
    dx, dy = origem[0] - cx, origem[1] - cy
    nx = dx * math.cos(rad) - dy * math.sin(rad)
    ny = dx * math.sin(rad) + dy * math.cos(rad)
    return nova, (nova.width / 2 + nx, nova.height / 2 + ny)


def desenhar_nome(capa, tamanho, nome, ajuste=None):
    """Desenha o nome do aluno na prévia (ajuste None = o salvo para este aluno/capa)."""
    try:
        modelo = modelo_nome(tamanho)
    except Exception as e:
        print("  AVISO: modelo do nome indisponível:", e)
        modelo = None
    if not modelo or not nome:
        return capa
    if ajuste is None:
        ajuste = ler_ajuste_nome(nome, tamanho)
    layout = calcular_layout_nome(tamanho, nome, ajuste, modelo)

    modo = capa.mode
    resultado = capa.convert("RGBA")
    escala = resultado.width / modelo["area"][0]
    itens = [
        (modelo["linhas"][0], layout["linha1"], layout["escala1"]),
        (modelo["linhas"][1], layout["linha2"], layout["escala2"]),
    ]
    caixa_recorte = None
    if modelo.get("recorte"):
        caixa_recorte = tuple(round(v * escala) for v in modelo["recorte"])
    if modelo.get("lombada"):
        itens.append((modelo["lombada"], " ".join(palavras_nome(nome)), layout["escala_lombada"]))
    for medida, texto, escala_linha in itens:
        if not texto:
            continue
        cores_tupla = tuple(tuple(c) for c in (medida.get("cores") or [[212, 175, 95]]))
        px_round = round(medida["px"] * escala * escala_linha / 100, 2)
        imagem, origem = _imagem_texto_cached(
            medida.get("fonte"),
            float(medida.get("tracking", 0)),
            float(medida.get("escala_h", 1.0)),
            float(medida.get("escala_v", 1.0)),
            medida.get("alinhamento", "esquerda"),
            bool(medida.get("sombra")),
            cores_tupla,
            float(medida.get("rotacao", 0.0)),
            _texto_da_linha(medida, texto),
            px_round,
        )
        dx = round(medida["x"] * escala - origem[0])
        dy = round(medida["y"] * escala - origem[1])
        img_comp = imagem
        if caixa_recorte is not None and medida is not modelo.get("lombada"):
            rl, rt, rr, rb = caixa_recorte
            ix0 = max(0, rl - dx, -dx)
            iy0 = max(0, rt - dy, -dy)
            ix1 = min(img_comp.width, rr - dx, resultado.width - dx)
            iy1 = min(img_comp.height, rb - dy, resultado.height - dy)
        else:
            ix0 = max(0, -dx)
            iy0 = max(0, -dy)
            ix1 = min(img_comp.width, resultado.width - dx)
            iy1 = min(img_comp.height, resultado.height - dy)
        if ix1 > ix0 and iy1 > iy0:
            sub = img_comp.crop((ix0, iy0, ix1, iy1))
            resultado.alpha_composite(sub, dest=(dx + ix0, dy + iy0))
    return resultado if modo == "RGBA" else resultado.convert(modo)


# ------------------------------------------------------------
# Instruções por capa (ajustes_nome.json)
# ------------------------------------------------------------

def chave_ajuste_nome(nome, tamanho):
    return chave(nome) + "|" + str(tamanho)


def ler_ajustes_nome():
    try:
        dados = json.loads(AJUSTES_NOME_FILE.read_text(encoding="utf-8"))
        return dados if isinstance(dados, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def ler_ajuste_nome(nome, tamanho):
    return ler_ajustes_nome().get(chave_ajuste_nome(nome, tamanho)) or {}


def gravar_ajuste_nome(nome, tamanho, ajuste):
    """Salva {"corte", "escala1", "escala2"}; ajuste vazio volta ao automático."""
    dados = ler_ajustes_nome()
    chave_nome = chave_ajuste_nome(nome, tamanho)
    limpo = {}
    if ajuste:
        if isinstance(ajuste.get("corte"), int):
            limpo["corte"] = ajuste["corte"]
        for campo in ("escala1", "escala2"):
            valor = _escala_valida(ajuste.get(campo))
            if valor is not None:
                limpo[campo] = valor
    if limpo:
        limpo["nome"] = str(nome)
        dados[chave_nome] = limpo
    else:
        dados.pop(chave_nome, None)
    temporario = AJUSTES_NOME_FILE.with_suffix(".tmp")
    temporario.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(temporario, AJUSTES_NOME_FILE)
    return limpo


def _apagar_arquivos(arquivos):
    apagados = []
    for arquivo in arquivos:
        try:
            if arquivo.is_file():
                arquivo.unlink()
                apagados.append(arquivo)
        except OSError:
            pass
    return apagados


def excluir_preset_curso(tamanho, simbolo, rgb=None):
    """Apaga a prévia em cache deste curso nesta capa para ser refeita; devolve os arquivos apagados."""
    cor = tuple(rgb) if rgb is not None else None
    slug = re.sub(r"[^a-z0-9_-]+", "_", _chave_brasao(simbolo)).strip("_")
    cor_slug = "semcor" if cor is None else "%02x%02x%02x" % cor
    diretorio = PREVIA_CAPAS_DIR / "presets"
    chave_cache = tamanho + "_" + slug + "_" + cor_slug
    arquivos = [
        diretorio / (chave_cache + ".png"),
        diretorio / (chave_cache + ".json"),
        _arquivo_brasoes_preset(tamanho, slug),
        _arquivo_verso_preset(tamanho, slug, cor),
    ]
    arquivos += [diretorio / "base" / (tamanho + "_" + cor_slug + fim) for fim in ("_baixo.png", "_cima.png", ".json")]
    return _apagar_arquivos(arquivos)


def excluir_previa_capa(tamanho):
    """Apaga a base da capa (lida do PSD) e o modelo do nome; a próxima prévia relê o PSD."""
    _MODELOS_NOME.pop(tamanho, None)
    return _apagar_arquivos([
        PREVIA_CAPAS_DIR / (tamanho + ".json"),
        PREVIA_CAPAS_DIR / (tamanho + "_baixo.png"),
        PREVIA_CAPAS_DIR / (tamanho + "_cima.png"),
        PREVIA_CAPAS_DIR / "presets" / "nome" / (tamanho + ".json"),
    ])


def _chave_brasao(simbolo):
    return "".join(
        caractere for caractere in unicodedata.normalize("NFKD", str(simbolo)).casefold()
        if unicodedata.category(caractere) != "Mn"
    )


# ============================================================
# MAIN
# ============================================================

def main(argv=None):
    p = argparse.ArgumentParser(description="Recorte automático das fotos das capas")
    p.add_argument("--calibrar", action="store_true", help="aprende o enquadramento das fotos prontas")
    p.add_argument("--simular", action="store_true", help="não grava nada, só relatório e miniaturas")
    p.add_argument("--aluno", help="só arquivos que contenham este nome")
    p.add_argument("--tamanho", choices=["20X20", "30X20", "24X34", "30X22"], type=str.upper)
    p.add_argument("--gerar-previa-capa", action="store_true",
                   help="lê os PSDs e grava a capa sem a foto para a prévia da UI")
    p.add_argument("--sincronizar-drive", action="store_true",
                   help="baixa as fotos novas da pasta do Drive configurada")
    args = p.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    cfg = carregar_config()

    if args.sincronizar_drive:
        if not cfg["photos_link"]:
            print("Nenhum link de pasta do Drive configurado para as fotos.")
            return 1
        try:
            sincronizar_drive(cfg["photos_link"], cfg["fotos"])
        except Exception as e:
            print("ERRO ao ler a pasta do Drive:", e)
            return 1
        return 0

    if args.gerar_previa_capa:
        feitos = gerar_previa_capa(cfg, [args.tamanho] if args.tamanho else None)
        return 0 if feitos else 1

    if not cfg["fotos"].is_dir():
        print("PASTA DE FOTOS NÃO ENCONTRADA:", cfg["fotos"])
        return 1

    if args.calibrar:
        calibrar(cfg)
        return 0

    if PREVIEW_DIR.exists():
        shutil.rmtree(PREVIEW_DIR, ignore_errors=True)

    calib = carregar_calibracao(cfg)
    processar(cfg, calib, args.simular, args.aluno, args.tamanho)
    return 0


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.exit(main())
