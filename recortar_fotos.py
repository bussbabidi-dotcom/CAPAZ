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
import csv
import datetime
import functools
import io
import json
import math
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
LADO_PREVIA_CAPA = 1400
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
}

# Usado enquanto um tamanho não tiver fotos prontas para calibrar
CALIBRACAO_PADRAO = {
    "centro_x": 0.5,
    "olhos_y": 0.38,
    "altura_rosto": 0.16,
}

TOLERANCIA_MOLDE_PX = 2
LADO_DETECCAO = 1600


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


def tamanho_do_arquivo(caminho):
    """20X20 / 24X34 pelo sufixo; sem sufixo é a foto do 30X20."""
    nome = caminho.stem.upper().replace(" ", "")
    for tam in ("20X20", "24X34", "30X20"):
        if tam in nome:
            return tam
    return "30X20"


PLANO_POR_TAMANHO = {
    "30X20": "Prime Book 30X20",
    "20X20": "Essencial Book 20x20",
    "24X34": "Studio Day Berlim 24x34",
}


def formatar_nome(nome):
    """Só a primeira letra de cada palavra maiúscula: "MARIA DA SILVA" -> "Maria Da Silva"."""
    texto = re.sub(r"\s+", " ", str(nome or "")).strip()
    return re.sub(r"[^\W\d_]+", lambda m: m.group(0)[0].upper() + m.group(0)[1:].lower(), texto)


def nome_do_aluno(caminho):
    """Nome do aluno a partir do arquivo, sem o tamanho (20X20 / 24X34 / 30X20)."""
    nome = re.sub(r"(?i)(?<![0-9a-z])(20|24|30)\s*x\s*(20|34)(?![0-9])", " ", caminho.stem)
    nome = re.sub(r"[_\-]+", " ", nome)
    return formatar_nome(nome)


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
    Recupera também o curso já salvo na revisão (cursos_por_foto.json).
    Retorna [{"nome", "curso", "plano"}] em ordem alfabética.
    """
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


def listar_fotos(pasta):
    return sorted(
        f for f in pasta.iterdir()
        if f.is_file() and f.suffix.lower() in EXTENSOES and not f.name.startswith("._")
    )


# ============================================================
# ROSTO
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


def detectar_rosto_cache(caminho, img=None):
    """Busca o rosto no cache persistente por (caminho, mtime_ns, size) antes de rodar o YuNet ONNX."""
    caminho_p = Path(caminho)
    chave_cache = None
    try:
        st = caminho_p.stat()
        chave_cache = f"{os.path.normcase(str(caminho_p.resolve()))}|{st.st_mtime_ns}|{st.st_size}"
        cache = _carregar_cache_rostos()
        if chave_cache in cache:
            return cache[chave_cache]
    except Exception:
        pass
    if img is None:
        img, _, _ = abrir_foto(caminho_p)
    rosto = detectar_rosto(img)
    if chave_cache is not None:
        cache = _carregar_cache_rostos()
        with _cache_rostos_lock:
            cache[chave_cache] = rosto
        _salvar_cache_rostos()
    return rosto


def detector_yunet():
    det = getattr(_tls_detector, "detector", None)
    if det is None:
        with _cache_rostos_lock:
            if not MODELO_YUNET.is_file():
                print("Baixando modelo de detecção de rosto...")
                urllib.request.urlretrieve(URL_YUNET, MODELO_YUNET)
        det = cv2.FaceDetectorYN.create(
            str(MODELO_YUNET), "", (320, 320), 0.7, 0.3, 5000
        )
        _tls_detector.detector = det
    return det


def detectar_rosto(img):
    """
    Retorna o maior rosto em coordenadas da imagem original:
    {x, y, w, h, cx, olhos_y, metodo} ou None.
    """
    w, h = img.size
    escala = min(1.0, LADO_DETECCAO / max(w, h))
    pequena = img.convert("RGB")
    if escala < 1.0:
        pequena = pequena.resize((round(w * escala), round(h * escala)), Image.BILINEAR)

    bgr = cv2.cvtColor(np.asarray(pequena), cv2.COLOR_RGB2BGR)
    ph, pw = bgr.shape[:2]

    det = detector_yunet()
    det.setInputSize((pw, ph))
    _, faces = det.detect(bgr)

    if faces is not None and len(faces):
        f = max(faces, key=lambda r: r[2] * r[3])
        x, y, fw, fh = (float(v) / escala for v in f[:4])
        olho_d_y = float(f[5]) / escala
        olho_e_y = float(f[7]) / escala
        return {
            "x": x, "y": y, "w": fw, "h": fh,
            "cx": x + fw / 2,
            "olhos_y": (olho_d_y + olho_e_y) / 2,
            "metodo": "yunet",
        }

    global _haar
    if not hasattr(cv2, "CascadeClassifier"):
        # OpenCV 5 não traz mais o Haar: sem rosto, fica o corte centralizado
        return None
    if _haar is None:
        _haar = cv2.CascadeClassifier(
            cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        )
    cinza = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    caixas = _haar.detectMultiScale(cinza, 1.1, 6, minSize=(40, 40))
    if len(caixas):
        x, y, fw, fh = (float(v) / escala for v in max(caixas, key=lambda r: r[2] * r[3]))
        # Haar pega um quadrado um pouco maior que o YuNet
        return {
            "x": x, "y": y, "w": fw, "h": fh * 0.9,
            "cx": x + fw / 2,
            "olhos_y": y + fh * 0.4,
            "metodo": "haar",
        }

    return None


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
    antes = img.convert("RGB").copy()
    antes.thumbnail((700, 700))
    fx = antes.width / img.width
    d = ImageDraw.Draw(antes)
    d.rectangle([c * fx for c in caixa], outline=(255, 0, 0), width=4)
    depois = recortada.copy()
    depois.thumbnail((700, 700))
    folha = Image.new("RGB", (antes.width + depois.width + 20, max(antes.height, depois.height)), "white")
    folha.paste(antes, (0, 0))
    folha.paste(depois, (antes.width + 20, 0))
    folha.save(destino, "JPEG", quality=85)


def processar(cfg, calib, simular, filtro_aluno, filtro_tamanho):
    linhas = []
    ok = skip = erro = 0

    if not simular:
        cfg["backup"].mkdir(parents=True, exist_ok=True)

    for f in listar_fotos(cfg["fotos"]):
        tam = tamanho_do_arquivo(f)

        if filtro_tamanho and tam != filtro_tamanho:
            continue
        if filtro_aluno and chave(filtro_aluno) not in chave(f.stem):
            continue

        molde = cfg["moldes"].get(tam)
        registro = {"arquivo": f.name, "tamanho": tam, "antes": "", "depois": "", "status": "", "motivo": ""}

        try:
            atual = tamanho_rapido(f)
            registro["antes"] = f"{atual[0]}x{atual[1]}"

            if no_molde(atual, molde):
                registro.update(status="SKIP", motivo="já no tamanho do molde")
                skip += 1
                linhas.append(registro)
                continue

            fonte = abrir_foto(fonte_original(f, cfg))
            img = fonte[0]
            print(f"\n{tam}  {f.name}  ({img.width}x{img.height})")

            caixa = caixa_manual(f, tam, img.size, molde)
            if caixa:
                metodo = "corte manual"
            else:
                caixa, rosto = caixa_automatica(img, tam, cfg, calib, fonte_original(f, cfg))
                if not caixa:
                    registro.update(status="ERRO", motivo="rosto não encontrado")
                    print("  ERRO: rosto não encontrado")
                    erro += 1
                    linhas.append(registro)
                    continue
                metodo = "rosto " + rosto["metodo"]

            ampliacao = ampliacao_da_caixa(caixa, molde)

            if ampliacao > cfg["ampliacao_maxima"]:
                registro.update(
                    status="ERRO",
                    motivo=f"resolução baixa: precisaria ampliar {ampliacao:.0%}",
                )
                print("  ERRO:", registro["motivo"])
                erro += 1
                linhas.append(registro)
                continue

            recortada = recortar_imagem(img, caixa, molde)
            registro["depois"] = f"{molde[0]}x{molde[1]}"

            salvar_preview(img, caixa, recortada, PREVIEW_DIR / (f.stem + ".jpg"))

            if simular:
                registro.update(status="SIMULADO", motivo=f"{metodo}, escala {ampliacao:.0%}")
                print("  SIMULADO:", registro["motivo"])
            else:
                aplicar_recorte(f, caixa, tam, cfg, fonte)
                registro.update(status="OK", motivo=f"{metodo}, escala {ampliacao:.0%}")
                print("  OK:", registro["motivo"])
            ok += 1

        except Exception as e:
            registro.update(status="ERRO", motivo=f"{type(e).__name__}: {e}")
            print(f"\n{f.name}\n  ERRO: {registro['motivo']}")
            erro += 1

        linhas.append(registro)

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


def montar_previa_capa(recortada, previa, nome=None, tamanho=None, ajuste=None):
    """Encaixa a foto já recortada no lugar da camada FOTO da capa reduzida."""
    info, baixo, cima = previa
    l, t, r, b = info["foto"]
    capa = baixo.copy()
    foto = recortada.convert("RGB").resize((max(1, round(r - l)), max(1, round(b - t))), Image.LANCZOS)
    camada = Image.new("RGB", capa.size)
    camada.paste(foto, (round(l), round(t)))
    vl, vt, vr, vb = (round(v) for v in info["foto_visivel"])
    mascara = Image.new("L", capa.size, 0)
    ImageDraw.Draw(mascara).rectangle([vl, vt, vr - 1, vb - 1], fill=255)
    capa.paste(camada, (0, 0), mascara)
    capa.paste(cima, (0, 0), cima)
    if nome and tamanho:
        capa = desenhar_nome(capa, tamanho, nome, ajuste)
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


def montar_previa_preset(recortada, preset, nome=None, ajuste=None):
    """Insere o recorte do aluno no preset do curso, preservando a ordem das layers.

    Com nome, desenha o nome do aluno como o Photoshop (ajuste None = o salvo em ajustes_nome.json).
    """

    capa = preset["baixo"].copy().convert("RGBA")
    l, t, r, b = preset["foto"]
    foto = recortada.convert("RGB").resize(
        (max(1, round(r - l)), max(1, round(b - t))), Image.LANCZOS,
    )
    camada_foto = Image.new("RGBA", capa.size, (0, 0, 0, 0))
    camada_foto.paste(foto, (round(l), round(t)))
    vl, vt, vr, vb = (round(valor) for valor in preset["foto_visivel"])
    mascara = Image.new("L", capa.size, 0)
    ImageDraw.Draw(mascara).rectangle((vl, vt, vr - 1, vb - 1), fill=255)
    capa.paste(camada_foto, (0, 0), mascara)

    camadas = (preset["meio"], preset["frente"], preset["cima"])
    if not preset["frente_depois_foto"]:
        camadas = (preset["frente"], preset["meio"], preset["cima"])
    else:
        camadas = (preset["meio"], preset["cima"], preset["frente"])
    for imagem in camadas:
        capa.alpha_composite(imagem.convert("RGBA"))
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
    pastas = [
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts",
    ]
    arquivos = [a for p in pastas if p.is_dir() for a in p.iterdir() if a.suffix.lower() in (".ttf", ".otf")]
    for arquivo in arquivos:
        if _chave_fonte(arquivo.stem) == alvo:
            return str(arquivo)
    if _FONTES_POR_CHAVE is None:
        _FONTES_POR_CHAVE = {}
        for arquivo in arquivos:
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
    return [p for p in re.sub(r"\s+", " ", str(nome or "")).strip().split(" ") if p]


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
    recorte = None
    if modelo.get("recorte"):
        l, t, r, b = (round(v * escala) for v in modelo["recorte"])
        recorte = Image.new("L", resultado.size, 0)
        ImageDraw.Draw(recorte).rectangle((l, t, r - 1, b - 1), fill=255)
    if modelo.get("lombada"):
        itens.append((modelo["lombada"], " ".join(palavras_nome(nome)), layout["escala_lombada"]))
    for medida, texto, escala_linha in itens:
        if not texto:
            continue
        imagem, origem = _imagem_texto(medida, _texto_da_linha(medida, texto), medida["px"] * escala * escala_linha / 100)
        imagem, origem = _girar_texto(imagem, origem, medida["rotacao"])
        camada = Image.new("RGBA", resultado.size, (0, 0, 0, 0))
        camada.paste(imagem, (round(medida["x"] * escala - origem[0]), round(medida["y"] * escala - origem[1])))
        if recorte is not None and medida is not modelo.get("lombada"):
            camada.putalpha(ImageChops.multiply(camada.getchannel("A"), recorte))
        resultado.alpha_composite(camada)
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
    sys.exit(main())
