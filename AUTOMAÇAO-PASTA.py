# -*- coding: utf-8 -*-

import argparse
import csv
import difflib
import os
import re
import json
import shutil
import sys
import tempfile
import threading
import concurrent.futures
import hashlib
import time
import uuid
from pathlib import Path

import openpyxl
import win32com.client


# ============================================================
# CONFIGURAÇÃO PADRÃO
# ============================================================

CONFIG_FILE = Path(__file__).parent / "automacao_capas_config.json"
CURSOS_POR_FOTO_FILE = Path(__file__).parent / "cursos_por_foto.json"
CURSOS_POR_FOTO = {}

PHOTOS_DIR = Path(
    r"D:\ÁLBUNS COLAÇÃO 2025.2\CAPAS\FOTO"
)

PLANILHA_DIR = Path(
    r"C:\Users\REIS\Documents\PAPA 2\capas\planilha"
)

PSD_DIR = Path(
    r"C:\Users\REIS\Documents\PAPA 2\capas\CAPA"
)


OUTPUT_DIRS = {

    "20X20": Path(
        r"D:\ÁLBUNS COLAÇÃO 2025.2\CAPAS\CAPAS 20X20"
    ),

    "30X20": Path(
        r"D:\ÁLBUNS COLAÇÃO 2025.2\CAPAS\CAPAS\CAPS 30X20"
    ),

    "24X34": Path(
        r"D:\ÁLBUNS COLAÇÃO 2025.2\CAPAS\CAPAS\CAPS 24X34"
    ),

    "30X22": Path(
        r"D:\ÁLBUNS COLAÇÃO 2025.2\CAPAS\CAPAS\CAPAS 30X22"
    )
}


# ============================================================
# CONTROLE
# ============================================================

DRY_RUN = False

EXPORT_JPG = True

EXPORT_PSD = True

SKIP_EXISTING = True

FOTO_MATCH_FLEXIVEL = True

BERLIM_GERAR_24 = True

BERLIM_GERAR_20 = True

PLANILHA_ESCOLHIDA = None

RELATORIO_CSV = Path(__file__).parent / "resultado_capas.csv"


def carregar_config(caminho=None):
    """Carrega automacao_capas_config.json e aplica nas variáveis globais."""

    global PHOTOS_DIR, PLANILHA_DIR, PSD_DIR
    global DRY_RUN, EXPORT_JPG, EXPORT_PSD
    global SKIP_EXISTING, FOTO_MATCH_FLEXIVEL
    global BERLIM_GERAR_24, BERLIM_GERAR_20, PLANILHA_ESCOLHIDA
    global CURSOS_POR_FOTO

    arquivo = Path(caminho) if caminho else CONFIG_FILE

    if not arquivo.is_file():
        return

    dados = json.loads(
        arquivo.read_text(encoding="utf-8")
    )

    if dados.get("photos_dir"):
        PHOTOS_DIR = Path(dados["photos_dir"])

    if dados.get("planilha_dir"):
        PLANILHA_DIR = Path(dados["planilha_dir"])

    if dados.get("psd_dir"):
        PSD_DIR = Path(dados["psd_dir"])

    if dados.get("output_dirs"):
        for chave_dir, valor in dados["output_dirs"].items():
            if valor:
                OUTPUT_DIRS[chave_dir] = Path(valor)
        if not dados["output_dirs"].get("30X22"):
            OUTPUT_DIRS["30X22"] = OUTPUT_DIRS["30X20"].parent / "CAPAS 30X22"

    if "dry_run" in dados:
        DRY_RUN = bool(dados["dry_run"])

    if "export_jpg" in dados:
        EXPORT_JPG = bool(dados["export_jpg"])

    if "export_psd" in dados:
        EXPORT_PSD = bool(dados["export_psd"])

    if "skip_existing" in dados:
        SKIP_EXISTING = bool(dados["skip_existing"])

    if "foto_match_flexivel" in dados:
        FOTO_MATCH_FLEXIVEL = bool(dados["foto_match_flexivel"])

    if "berlim_gerar_24" in dados:
        BERLIM_GERAR_24 = bool(dados["berlim_gerar_24"])

    if "berlim_gerar_20" in dados:
        BERLIM_GERAR_20 = bool(dados["berlim_gerar_20"])

    if dados.get("planilha"):
        PLANILHA_ESCOLHIDA = Path(dados["planilha"])

    CORES_PERSONALIZADAS.clear()
    for simbolo, valor in (dados.get("cores_cursos") or {}).items():
        rgb = cor_de_texto(valor)
        if rgb:
            CORES_PERSONALIZADAS[simbolo] = rgb

    CURSOS_POR_FOTO = {}
    try:
        mapa = json.loads(CURSOS_POR_FOTO_FILE.read_text(encoding="utf-8"))
        CURSOS_POR_FOTO = mapa.get("cursos_por_foto", {})
        if not isinstance(CURSOS_POR_FOTO, dict):
            CURSOS_POR_FOTO = {}
    except (OSError, json.JSONDecodeError):
        pass


# ============================================================
# MODELOS
# ============================================================

MODELOS = {

    # ========================================================
    # PRIME BOOK 30X20
    # ========================================================

    "PRIME BOOK 30X20": {

        "tamanho": "30X20",

        "psd":
            "CAPA 30X20 (PRIME BOOK) - ZANGRAF.psd",

        "foto": [
            "FOTO"
        ],

        "nome_principal": [
            "NOME"
        ],

        "background": [
            "BACKGROUND"
        ],

        "frente": [
            "BRAÇÕES",
            "FRENTE"
        ],

        "lateral": [
            "LATERAL"
        ],

        "nome_lateral": [
            "NOME"
        ]
    },


    # ========================================================
    # MOSCOU 30X20
    # ========================================================

    "STUDIO DAY MOSCOU 30X20": {

        "tamanho": "30X20",

        "psd":
            "CAPA 30X20 (PRIME BOOK) - ZANGRAF.psd",

        "foto": [
            "FOTO"
        ],

        "nome_principal": [
            "NOME"
        ],

        "background": [
            "BACKGROUND"
        ],

        "frente": [
            "BRAÇÕES",
            "FRENTE"
        ],

        "lateral": [
            "LATERAL"
        ],

        "nome_lateral": [
            "NOME"
        ]
    },


    # ========================================================
    # BERLIM 24X34
    #
    # SOMENTE O PSD 24X34.
    #
    # O 20X20 AGORA É FEITO EM OUTRO PSD.
    # ========================================================

    "STUDIO DAY BERLIM 24X34": {

        "tamanho": "24X34",

        "psd":
            "CAPA 20X20 24x34 (BERLIM) - ZANGRAF.psd",

        "foto34": [
            "24X34",
            "FOTO"
        ],

        "frente_24": [
            "24X34",
            "BRAÇÕES",
            "FRENTE (PRINCIPAL) copiar 2"
        ],

        "nome_24": [
            "24X34",
            "NOME copiar"
        ]
    },


    # ========================================================
    # PRIME BOOK 30X22
    #
    # DOIS PSDS: FRENTE (foto, nomes, cor, brasão) e
    # VERSO (cor do BACKGROUND, brasão). Foto igual à do 30X20.
    # ========================================================

    "PRIME BOOK 30X22": {

        "tamanho": "30X22",

        "frente": {

            "psd":
                "CAPA FRENTE 30x22.jpg.psd",

            "foto": [
                "FOTO"
            ],

            "nome_principal": [
                "NOME"
            ],

            "nome_lateral":
                "NOME",

            "cor": [
                "Preenchimento de Cor 1"
            ],

            "brasao": [
                "FRENTE"
            ]
        },

        "verso": {

            "psd":
                "CAPA VERSO 30x22.jpg.psd",

            "background": [
                "BACKGROUND"
            ],

            "brasao": [
                "FRENTE copiar"
            ]
        }
    }
}


# ============================================================
# MODELO ESSENCIAL BOOK 20X20
#
# ESTRUTURA CONFIRMADA PELO PRINT ENVIADO:
#
# NOME
# ├── NOME
# └── nome do aluno
#
# LATERAL
# └── NOME
#
# LOGOTIPO FINAL
#
# BRASÃO
# ├── FRENTE (PRINCIPAL) copiar
# └── FRENTE (PRINCIPAL)
#
# FOTO
# ├── Camada 1
# ├── FOTO
# ├── Preenchimento generativo 3
# └── Retângulo 5
#
# BACKGROUND
# ============================================================

MODELO_ESSENCIAL_20 = {

    "tamanho": "20X20",

    "psd":
        "CAPA 20X20 (ESSENCIAL BOOK) - ZANGRAF.psd",

    "foto": [
        "FOTO",
        "FOTO"
    ],

    "frente": [
        "BRASÃO",
        "FRENTE (PRINCIPAL) copiar"
    ],

    "nome_principal": [
        "NOME"
    ],

    "background": [
        "BACKGROUND"
    ],

    "lateral": [
        "LATERAL"
    ],

    "nome_lateral": [
        "NOME"
    ]
}


# ============================================================
# CURSOS
# ============================================================

CURSOS = {

    "ARTES VISUAIS":
        "ART VIS",

    "ARQUITETURA":
        "ARQ",

    "ARQUITETURA E URBANISMO":
        "ARQ",

    "AGRONOMIA":
        "AGRONOMIA",

    "ANALISE E DESENVOLVIMENTO DE SISTEMAS":
        "ADS",

    "ADMINISTRACAO":
        "ADM",

    "BIOMEDICINA":
        "BIOMED",

    "COMUNICACAO SOCIAL":
        "COM SOC",

    "CIENCIAS SOCIAIS":
        "CIEN SOC",

    "CIENCIAS DA COMPUTACAO":
        "CIEN COMP",

    "CIENCIAS BIOLOGICAS":
        "CIEN BIO",

    "CIENCIAS CONTABEIS":
        "CCONT",

    "DIREITO":
        "DIREITO",

    "DESIGN GRAFICO":
        "DESIGN GRAFICO",

    "DESIGN DE INTERIORES":
        "DESIGN DE INT",

    "ESTETICA":
        "ESTETICA",

    "ENGENHARIA DE PRODUCAO":
        "ENG PROD",

    "ENGENHARIA MECANICA":
        "ENG MEC",

    "ENGENHARIA ELETRICA":
        "ENG ELET",

    "ENGENHARIA CIVIL":
        "ENG CIV",

    "ENGENHARIA AMBIENTAL":
        "ENG AMB",

    "ENGENHARIA DA COMPUTACAO":
        "ENG DA COMP",

    "ENFERMAGEM":
        "ENF",

    "EDUCACAO FISICA":
        "ED FIS",

    "FILMAKER":
        "FILMAKER",

    "FONOAUDIOLOGIA":
        "FONO",

    "FISIOTERAPIA":
        "FISIO",

    "FISICA":
        "FÍSICA",

    "FARMACIA":
        "FARM",

    "GESTAO PUBLICA":
        "GEST PUBLI",

    "GESTAO COMERCIAL":
        "GEST COM",

    "GESTAO AMBIENTAL":
        "GEST AMB",

    "GASTRONOMIA":
        "GASTRONOMIA",

    "GESTAO PORTUARIA":
        "GESTÃO PORTUÁRIA",

    "HISTORIA":
        "HIST",

    "JORNALISMO":
        "JORNALISMO",

    "LOGISTICA":
        "LOG",

    "LETRAS":
        "LETRAS",

    "LETRAS LIBRAS":
        "LETRAS LIBRAS",

    "MODA":
        "MODA",

    "MEDICINA":
        "MEDICINA",

    "MEDICINA VETERINARIA":
        "MED VET",

    "MATEMATICA":
        "MAT",

    "MARKETING":
        "MARKETING",

    "NUTRICAO":
        "NUTRI",

    "NEGOCIOS IMOBILIARIOS":
        "NEGÓCIOS IMOBILIÁRIOS",

    "ODONTOLOGIA":
        "ODONTO",

    "PSICOLOGIA":
        "PSI",

    "PROCESSOS GERENCIAIS":
        "PROC GEN",

    "PODOLOGIA":
        "PODOLOGIA",

    "PERICIA CRIMINAL":
        "PERÍCIA CRIMINAL",

    "PEDAGOGIA":
        "PED",

    "PUBLICIDADE E PROPAGANDA":
        "PUBLI",

    "RECURSOS HUMANOS":
        "RH",

    "RELACOES INTERNACIONAIS":
        "RELAC INT",

    "REDES DE COMPUTADORES":
        "REDES DE COMP",

    "REDE DE COMPUTADORES":
        "REDES DE COMP",

    "RADIOLOGIA":
        "RADIO",

    "RADIO":
        "RADIO",

    "SERVICO SOCIAL":
        "SS",

    "SISTEMAS DE INFORMACAO":
        "SI",

    "SERVICOS JURIDICOS":
        "SER JUR",

    "SEGURANCA PUBLICA":
        "SEGURANÇA PUBLICA",

    "TECNOLOGIA DA INFORMACAO":
        "TI",

    "TERAPIA OCUPACIONAL":
        "TERAPIA OCUPACIONAL"
}


# Cores dos brasões (RGB), medidas nos arquivos de
# #TEMPLATES\BRASÕES (fundo do escudo, sem o dourado).
# Usadas em alterarCorBackground no Photoshop.
COR_AZUL = (33, 43, 74)
COR_VERDE = (28, 66, 35)
COR_VERMELHO = (122, 26, 24)
COR_ROXO = (79, 49, 108)
COR_VINHO = (123, 36, 91)
COR_CARMIM = (139, 24, 62)
COR_OLIVA = (157, 152, 47)

CORES_BACKGROUND = {

    # AZUL
    "ADM": COR_AZUL,
    "ADS": COR_AZUL,
    "AGRONOMIA": COR_AZUL,
    "ARQ": COR_AZUL,
    "ART VIS": COR_AZUL,
    "CIEN BIO": COR_AZUL,
    "CIEN COMP": COR_AZUL,
    "COM SOC": COR_AZUL,
    "DESIGN DE INT": COR_AZUL,
    "DESIGN GRAFICO": COR_AZUL,
    "ENG AMB": COR_AZUL,
    "ENG CIV": COR_AZUL,
    "ENG DA COMP": COR_AZUL,
    "ENG ELET": COR_AZUL,
    "ENG MEC": COR_AZUL,
    "ENG PROD": COR_AZUL,
    "FILMAKER": COR_AZUL,
    "FONO": COR_AZUL,
    "FÍSICA": COR_AZUL,
    "GEST AMB": COR_AZUL,
    "GEST COM": COR_AZUL,
    "GEST PUBLI": COR_AZUL,
    "GESTÃO PORTUÁRIA": COR_AZUL,
    "LETRAS": COR_AZUL,
    "LOG": COR_AZUL,
    "MARKETING": COR_AZUL,
    "MAT": COR_AZUL,
    "MODA": COR_AZUL,
    "NEGÓCIOS IMOBILIÁRIOS": COR_AZUL,
    "PROC GEN": COR_AZUL,
    "PSI": COR_AZUL,
    "PUBLI": COR_AZUL,
    "REDES DE COMP": COR_AZUL,
    "RH": COR_AZUL,
    "SI": COR_AZUL,
    "TI": COR_AZUL,

    # VERDE
    "BIOMED": COR_VERDE,
    "ED FIS": COR_VERDE,
    "ENF": COR_VERDE,
    "ESTETICA": COR_VERDE,
    "FISIO": COR_VERDE,
    "MED VET": COR_VERDE,
    "MEDICINA": COR_VERDE,
    "NUTRI": COR_VERDE,
    "PODOLOGIA": COR_VERDE,
    "SS": COR_VERDE,
    "TERAPIA OCUPACIONAL": COR_VERDE,

    # VERMELHO
    "DIREITO": COR_VERMELHO,
    "JORNALISMO": COR_VERMELHO,
    "PERÍCIA CRIMINAL": COR_VERMELHO,
    "RELAC INT": COR_VERMELHO,
    "SEGURANÇA PUBLICA": COR_VERMELHO,
    "SER JUR": COR_VERMELHO,

    # ROXO
    "CIEN SOC": COR_ROXO,
    "GASTRONOMIA": COR_ROXO,
    "HIST": COR_ROXO,
    "LETRAS LIBRAS": COR_ROXO,
    "PED": COR_ROXO,

    # VINHO
    "CCONT": COR_VINHO,

    # CARMIM
    "ODONTO": COR_CARMIM,

    # OLIVA
    "FARM": COR_OLIVA,
    "RADIO": COR_OLIVA,
}


# Cores trocadas pelo usuário (config "cores_cursos": {"DIREITO": "#7A1A18"})
CORES_PERSONALIZADAS = {}


def cor_de_texto(valor):
    """"#RRGGBB" ou [r, g, b] -> (r, g, b); None se inválido."""

    if isinstance(valor, str):
        m = re.fullmatch(r"\s*#?([0-9a-fA-F]{6})\s*", valor)
        if not m:
            return None
        h = m.group(1)
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

    if isinstance(valor, (list, tuple)) and len(valor) == 3:
        try:
            rgb = tuple(int(v) for v in valor)
        except (TypeError, ValueError):
            return None
        if all(0 <= v <= 255 for v in rgb):
            return rgb

    return None


def obter_cor_background(simbolo):
    """Retorna [r, g, b] ou None para o símbolo do brasão/curso."""

    if not simbolo:
        return None

    alvo_personalizado = chave(simbolo)
    for nome, rgb in CORES_PERSONALIZADAS.items():
        if chave(nome) == alvo_personalizado:
            return [rgb[0], rgb[1], rgb[2]]

    # Match direto
    if simbolo in CORES_BACKGROUND:
        r, g, b = CORES_BACKGROUND[simbolo]
        return [r, g, b]

    # Match por chave normalizada
    alvo = chave(simbolo)

    for nome, rgb in CORES_BACKGROUND.items():
        if chave(nome) == alvo:
            return [rgb[0], rgb[1], rgb[2]]

    return None


# Aliases e abreviações de curso (planilha -> chave canônica em CURSOS)
CURSO_ALIASES = {

    # Administração
    "ADM": "ADMINISTRACAO",
    "ADMINISTRACAO DE EMPRESAS": "ADMINISTRACAO",
    "ADMINISTRACAO EMPRESARIAL": "ADMINISTRACAO",
    "ADMNISTRACAO": "ADMINISTRACAO",

    # ADS
    "ADS": "ANALISE E DESENVOLVIMENTO DE SISTEMAS",
    "A.D.S": "ANALISE E DESENVOLVIMENTO DE SISTEMAS",
    "A.D.S.": "ANALISE E DESENVOLVIMENTO DE SISTEMAS",
    "ANALISE DE SISTEMAS": "ANALISE E DESENVOLVIMENTO DE SISTEMAS",
    "ANALISE E DESENVOLVIMENTO DE SISTEMA": "ANALISE E DESENVOLVIMENTO DE SISTEMAS",
    "ADS ANALISE E DESENVOLVIMENTO DE SISTEMAS": "ANALISE E DESENVOLVIMENTO DE SISTEMAS",
    "ADS- ANALISE E DESENVOLVIMENTO DE SISTEMA": "ANALISE E DESENVOLVIMENTO DE SISTEMAS",
    "ADS ANALISE E DESENVOLVIMENTO DE SISTEMA": "ANALISE E DESENVOLVIMENTO DE SISTEMAS",
    "DESENVOLVIMENTO DE SISTEMAS": "ANALISE E DESENVOLVIMENTO DE SISTEMAS",

    # Engenharia
    "ENG CIVIL": "ENGENHARIA CIVIL",
    "ENG. CIVIL": "ENGENHARIA CIVIL",
    "ENG.CIVIL": "ENGENHARIA CIVIL",
    "E CIVIL": "ENGENHARIA CIVIL",
    "ENG CIV": "ENGENHARIA CIVIL",
    "ENG MEC": "ENGENHARIA MECANICA",
    "ENG. MECANICA": "ENGENHARIA MECANICA",
    "ENG MECANICA": "ENGENHARIA MECANICA",
    "ENG ELET": "ENGENHARIA ELETRICA",
    "ENG. ELETRICA": "ENGENHARIA ELETRICA",
    "ENG ELETRICA": "ENGENHARIA ELETRICA",
    "ENG PROD": "ENGENHARIA DE PRODUCAO",
    "ENG. PRODUCAO": "ENGENHARIA DE PRODUCAO",
    "ENG PRODUCAO": "ENGENHARIA DE PRODUCAO",
    "ENG AMB": "ENGENHARIA AMBIENTAL",
    "ENG. AMBIENTAL": "ENGENHARIA AMBIENTAL",
    "ENG AMBIENTAL": "ENGENHARIA AMBIENTAL",
    "ENG COMP": "ENGENHARIA DA COMPUTACAO",
    "ENG DA COMP": "ENGENHARIA DA COMPUTACAO",
    "ENG. DA COMPUTACAO": "ENGENHARIA DA COMPUTACAO",
    "ENG DA COMPUTACAO": "ENGENHARIA DA COMPUTACAO",
    "ENG COMPUTACAO": "ENGENHARIA DA COMPUTACAO",
    "ENGENHARIA COMPUTACAO": "ENGENHARIA DA COMPUTACAO",
    "ENGENHARIA DE COMPUTACAO": "ENGENHARIA DA COMPUTACAO",
    "ENGENHARIA DA COMPUTACAO": "ENGENHARIA DA COMPUTACAO",

    # Computação / TI
    "CIENCIA DA COMPUTACAO": "CIENCIAS DA COMPUTACAO",
    "CIENCIAS DA COMPUTACAO": "CIENCIAS DA COMPUTACAO",
    "CIENCIA COMPUTACAO": "CIENCIAS DA COMPUTACAO",
    "C COMPUTACAO": "CIENCIAS DA COMPUTACAO",
    "CCOMP": "CIENCIAS DA COMPUTACAO",
    "TI": "TECNOLOGIA DA INFORMACAO",
    "T.I": "TECNOLOGIA DA INFORMACAO",
    "T.I.": "TECNOLOGIA DA INFORMACAO",
    "TEC INFORMACAO": "TECNOLOGIA DA INFORMACAO",
    "SI": "SISTEMAS DE INFORMACAO",
    "S.I": "SISTEMAS DE INFORMACAO",
    "S.I.": "SISTEMAS DE INFORMACAO",
    "SIS INFORMACAO": "SISTEMAS DE INFORMACAO",
    "REDES": "REDES DE COMPUTADORES",
    "REDES COMP": "REDES DE COMPUTADORES",

    # Contábeis / Gestão
    "CCONT": "CIENCIAS CONTABEIS",
    "C CONTABEIS": "CIENCIAS CONTABEIS",
    "CIENCIAS CONTABEIS": "CIENCIAS CONTABEIS",
    "CONTABEIS": "CIENCIAS CONTABEIS",
    "CONTABILIDADE": "CIENCIAS CONTABEIS",
    "RH": "RECURSOS HUMANOS",
    "R.H": "RECURSOS HUMANOS",
    "R.H.": "RECURSOS HUMANOS",
    "GEST COM": "GESTAO COMERCIAL",
    "GEST. COMERCIAL": "GESTAO COMERCIAL",
    "GEST AMB": "GESTAO AMBIENTAL",
    "GEST PUB": "GESTAO PUBLICA",
    "GEST PUBLICA": "GESTAO PUBLICA",
    "PROC GEN": "PROCESSOS GERENCIAIS",
    "PROCESSOS GERENCIAIS": "PROCESSOS GERENCIAIS",

    # Saúde
    "ENF": "ENFERMAGEM",
    "FISIO": "FISIOTERAPIA",
    "FONO": "FONOAUDIOLOGIA",
    "NUTRI": "NUTRICAO",
    "NUTRICAO": "NUTRICAO",
    "ODONTO": "ODONTOLOGIA",
    "BIOMED": "BIOMEDICINA",
    "FARM": "FARMACIA",
    "MED": "MEDICINA",
    "MED VET": "MEDICINA VETERINARIA",
    "MED. VET": "MEDICINA VETERINARIA",
    "MED. VETERINARIA": "MEDICINA VETERINARIA",
    "VET": "MEDICINA VETERINARIA",
    "VETERINARIA": "MEDICINA VETERINARIA",
    "PSI": "PSICOLOGIA",
    "PSICO": "PSICOLOGIA",
    "TO": "TERAPIA OCUPACIONAL",
    "T.O": "TERAPIA OCUPACIONAL",
    "TERAPIA OCUPACIONAL.": "TERAPIA OCUPACIONAL",

    # Educação / Humanas
    "PED": "PEDAGOGIA",
    "ED FISICA": "EDUCACAO FISICA",
    "ED. FISICA": "EDUCACAO FISICA",
    "ED FIS": "EDUCACAO FISICA",
    "ED. FIS": "EDUCACAO FISICA",
    "EDUCACAO FISICA": "EDUCACAO FISICA",
    "EDFISICA": "EDUCACAO FISICA",
    "HIST": "HISTORIA",
    "MAT": "MATEMATICA",
    "LETRAS LILAS": "LETRAS LIBRAS",
    "LIBRAS": "LETRAS LIBRAS",
    "LETRAS-LIBRAS": "LETRAS LIBRAS",

    # Comunicação / Design
    "PUBLI": "PUBLICIDADE E PROPAGANDA",
    "PP": "PUBLICIDADE E PROPAGANDA",
    "P&P": "PUBLICIDADE E PROPAGANDA",
    "PUBLICIDADE": "PUBLICIDADE E PROPAGANDA",
    "PROPAGANDA": "PUBLICIDADE E PROPAGANDA",
    "JORNAL": "JORNALISMO",
    "COM SOC": "COMUNICACAO SOCIAL",
    "DESIGN GRAF": "DESIGN GRAFICO",
    "DESIGN INT": "DESIGN DE INTERIORES",
    "DESIGN DE INT": "DESIGN DE INTERIORES",

    # Direito / Sociais
    "SS": "SERVICO SOCIAL",
    "S.S": "SERVICO SOCIAL",
    "SERVICO SOCIAL": "SERVICO SOCIAL",
    "REL INT": "RELACOES INTERNACIONAIS",
    "RI": "RELACOES INTERNACIONAIS",
    "R.I": "RELACOES INTERNACIONAIS",
    "SER JUR": "SERVICOS JURIDICOS",
    "SEG PUBLICA": "SEGURANCA PUBLICA",
    "SEG. PUBLICA": "SEGURANCA PUBLICA",

    # Outros
    "ARQ": "ARQUITETURA E URBANISMO",
    "ARQUITETURA": "ARQUITETURA E URBANISMO",
    "URBANISMO": "ARQUITETURA E URBANISMO",
    "ART VIS": "ARTES VISUAIS",
    "ARTES": "ARTES VISUAIS",
    "LOG": "LOGISTICA",
    "GASTRO": "GASTRONOMIA",
    "ESTETICA E COSMETICA": "ESTETICA",
    "COSMETICA": "ESTETICA",
}


# Expansão token a token (ENG -> ENGENHARIA, etc.)
ABREVIACOES_TOKEN_CURSO = {

    "ENG": "ENGENHARIA",
    "ENG.": "ENGENHARIA",
    "ADM": "ADMINISTRACAO",
    "ADM.": "ADMINISTRACAO",
    "ED": "EDUCACAO",
    "ED.": "EDUCACAO",
    "MED": "MEDICINA",
    "MED.": "MEDICINA",
    "VET": "VETERINARIA",
    "VET.": "VETERINARIA",
    "PSI": "PSICOLOGIA",
    "PSI.": "PSICOLOGIA",
    "PSICO": "PSICOLOGIA",
    "FISIO": "FISIOTERAPIA",
    "FONO": "FONOAUDIOLOGIA",
    "NUTRI": "NUTRICAO",
    "ODONTO": "ODONTOLOGIA",
    "BIOMED": "BIOMEDICINA",
    "FARM": "FARMACIA",
    "ENF": "ENFERMAGEM",
    "PED": "PEDAGOGIA",
    "HIST": "HISTORIA",
    "MAT": "MATEMATICA",
    "ARQ": "ARQUITETURA",
    "CIV": "CIVIL",
    "CIV.": "CIVIL",
    "MEC": "MECANICA",
    "MEC.": "MECANICA",
    "ELET": "ELETRICA",
    "ELET.": "ELETRICA",
    "PROD": "PRODUCAO",
    "PROD.": "PRODUCAO",
    "AMB": "AMBIENTAL",
    "AMB.": "AMBIENTAL",
    "COMP": "COMPUTACAO",
    "COMP.": "COMPUTACAO",
    "INFO": "INFORMACAO",
    "INFO.": "INFORMACAO",
    "GEST": "GESTAO",
    "GEST.": "GESTAO",
    "PUB": "PUBLICA",
    "PUB.": "PUBLICA",
    "COM": "COMERCIAL",
    "COM.": "COMERCIAL",
    "REL": "RELACOES",
    "REL.": "RELACOES",
    "INT": "INTERNACIONAIS",
    "INT.": "INTERNACIONAIS",
    "SER": "SERVICOS",
    "SER.": "SERVICOS",
    "JUR": "JURIDICOS",
    "JUR.": "JURIDICOS",
    "SEG": "SEGURANCA",
    "SEG.": "SEGURANCA",
    "PROC": "PROCESSOS",
    "PROC.": "PROCESSOS",
    "GEN": "GERENCIAIS",
    "GEN.": "GERENCIAIS",
    "ART": "ARTES",
    "ART.": "ARTES",
    "VIS": "VISUAIS",
    "VIS.": "VISUAIS",
    "LOG": "LOGISTICA",
    "LOG.": "LOGISTICA",
    "GASTRO": "GASTRONOMIA",
    "PUBLI": "PUBLICIDADE",
    "PROP": "PROPAGANDA",
    "GRAF": "GRAFICO",
    "SOC": "SOCIAL",
    "SOC.": "SOCIAL",
    "FIS": "FISICA",
    "FIS.": "FISICA",
}


def _cursos_lookup():

    tabela = {}

    for nome, simbolo in CURSOS.items():
        tabela[chave(nome)] = simbolo
        # Também indexa pelo próprio símbolo do brasão
        tabela[chave(simbolo)] = simbolo

    for alias, destino in CURSO_ALIASES.items():
        kd = chave(destino)
        if kd in tabela:
            tabela[chave(alias)] = tabela[kd]
        elif destino in CURSOS:
            tabela[chave(alias)] = CURSOS[destino]

    return tabela


def _expandir_abreviacoes_curso(texto):
    """Expande tokens abreviados e gera formas alternativas do curso."""

    bruto = chave(texto)
    # unifica pontuação / hífen
    bruto = re.sub(r"[.\-/]+", " ", bruto)
    bruto = re.sub(r"\s+", " ", bruto).strip()

    formas = [bruto, re.sub(r"\s+", "", bruto)]

    tokens = bruto.split()
    expandidos = []

    for tok in tokens:
        if tok in ABREVIACOES_TOKEN_CURSO:
            expandidos.append(ABREVIACOES_TOKEN_CURSO[tok])
        elif tok.rstrip(".") in ABREVIACOES_TOKEN_CURSO:
            expandidos.append(
                ABREVIACOES_TOKEN_CURSO[tok.rstrip(".")]
            )
        else:
            expandidos.append(tok)

    expandido = " ".join(expandidos)
    formas.append(expandido)
    formas.append(re.sub(r"\s+", "", expandido))

    # plural/singular simples
    for f in list(formas):
        if f.endswith("S") and len(f) > 4:
            formas.append(f[:-1])
        else:
            formas.append(f + "S")

    return list(dict.fromkeys(formas))


def resolver_curso(curso):
    """
    Resolve o símbolo do brasão a partir do curso da planilha.
    Aceita abreviações, pontuação, hífens e variações de escrita.
    Retorna (simbolo, aviso_ou_None).
    """

    curso_norm = normalizar(curso)

    if not curso_norm:
        return (
            None,
            "CURSO VAZIO — brasão/background podem ficar sem alteração"
        )

    if re.match(r"(?i)^(#|http \d{3}|curso vazio$|loading|carregando)", curso_norm):
        raise Exception(
            "CURSO COM ERRO NA PLANILHA: " + curso_norm
            + "\n\nA fórmula do curso não trouxe o nome. Corrija a célula na planilha."
        )

    lookup = _cursos_lookup()
    candidatos_chave = _expandir_abreviacoes_curso(curso_norm)

    # 1) Match direto / alias / forma expandida
    for k in candidatos_chave:
        if k in lookup:
            aviso = None
            if chave(k) != chave(curso_norm):
                aviso = (
                    "CURSO POR ABREVIACAO/VARIACAO: "
                    + curso_norm
                    + " -> "
                    + k
                )
            return lookup[k], aviso

    # 2) Match parcial entre formas e nomes cadastrados
    candidatos = []

    for k in candidatos_chave:
        for nome_chave, simbolo in lookup.items():
            if not k or len(k) < 3:
                continue
            if k == nome_chave:
                candidatos.append((nome_chave, simbolo))
            elif k in nome_chave or nome_chave in k:
                candidatos.append((nome_chave, simbolo))

    # únicos por símbolo
    por_simbolo = {}
    for nome_chave, simbolo in candidatos:
        por_simbolo.setdefault(simbolo, set()).add(nome_chave)

    if len(por_simbolo) == 1:
        simbolo = next(iter(por_simbolo))
        nome_hit = next(iter(por_simbolo[simbolo]))
        return (
            simbolo,
            "CURSO RESOLVIDO POR MATCH PARCIAL: "
            + curso_norm
            + " -> "
            + nome_hit
        )

    if len(por_simbolo) > 1:
        raise Exception(
            "CURSO AMBÍGUO:\n\n"
            + curso_norm
            + "\n\nCandidatos:\n"
            + "\n".join(sorted(por_simbolo.keys()))
        )

    # 3) Similaridade
    proximos = difflib.get_close_matches(
        chave(curso_norm),
        list(lookup.keys()),
        n=5,
        cutoff=0.65
    )

    # tenta também formas expandidas
    for k in candidatos_chave:
        proximos.extend(
            difflib.get_close_matches(
                k,
                list(lookup.keys()),
                n=3,
                cutoff=0.7
            )
        )

    proximos = list(dict.fromkeys(proximos))

    if len(proximos) == 1:
        return (
            lookup[proximos[0]],
            "CURSO POR SIMILARIDADE: "
            + curso_norm
            + " -> "
            + proximos[0]
        )

    sugestao = ""
    if proximos:
        sugestao = (
            "\n\nSugestões:\n"
            + "\n".join(proximos[:5])
        )

    raise Exception(
        "CURSO NÃO CONFIGURADO:\n\n"
        + curso_norm
        + sugestao
    )


# ============================================================
# NORMALIZAÇÃO
# ============================================================

def normalizar(valor):

    if valor is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(valor).strip()
    )


def chave(valor):

    import unicodedata

    texto = normalizar(
        valor
    ).upper()

    texto = "".join(
        c
        for c in unicodedata.normalize(
            "NFD",
            texto
        )
        if unicodedata.category(c) != "Mn"
    )

    return texto


def nome_seguro(valor):

    return re.sub(
        r'[<>:"/\\|?*]',
        "_",
        normalizar(valor)
    )


PARTICULAS_NOME = {
    "DE", "DA", "DO", "DOS", "DAS", "E", "DI", "DU", "DEL", "DELLA", "VAN", "VON"
}


def variantes_nome_pessoa(nome):
    """
    Gera variações do nome da pessoa para achar foto/arquivo:
    completo, sem partículas, primeiro+último, abreviações do meio, etc.
    """

    bruto = chave(nome)
    if not bruto:
        return []

    tokens = bruto.split()
    formas = [bruto, " ".join(tokens)]

    # sem partículas (de/da/dos...)
    sem_part = [t for t in tokens if t not in PARTICULAS_NOME]
    if sem_part and sem_part != tokens:
        formas.append(" ".join(sem_part))

    # primeiro + último
    if len(tokens) >= 2:
        formas.append(tokens[0] + " " + tokens[-1])
        if sem_part and len(sem_part) >= 2:
            formas.append(sem_part[0] + " " + sem_part[-1])

    # primeiro + penúltimo + último (quando há sobrenome composto)
    if len(tokens) >= 3:
        formas.append(
            tokens[0] + " " + tokens[-2] + " " + tokens[-1]
        )
    if len(sem_part) >= 3:
        formas.append(
            sem_part[0] + " " + sem_part[-2] + " " + sem_part[-1]
        )

    # primeiro + segundo (nome composto: Ana Paula ...)
    if len(tokens) >= 3:
        formas.append(tokens[0] + " " + tokens[1] + " " + tokens[-1])

    # iniciais do meio: Maria L Oliveira / M L Oliveira
    if len(sem_part) >= 3:
        meio_inicial = [
            sem_part[0]
        ] + [t[0] for t in sem_part[1:-1]] + [
            sem_part[-1]
        ]
        formas.append(" ".join(meio_inicial))
        formas.append("".join(meio_inicial))

        so_iniciais_meio = (
            [sem_part[0]]
            + [t[0] + "." for t in sem_part[1:-1]]
            + [sem_part[-1]]
        )
        formas.append(" ".join(so_iniciais_meio))

    # só primeiro nome (último recurso — usado com cuidado no match)
    if tokens:
        formas.append(tokens[0])

    # compactas
    resultado = []
    for f in formas:
        f = re.sub(r"\s+", " ", f).strip()
        if not f:
            continue
        resultado.append(f)
        resultado.append(re.sub(r"[\s_\-\.]+", "", f))

    return list(dict.fromkeys(resultado))


# ============================================================
# PLANILHA
# ============================================================

def encontrar_planilha(caminho=None):

    if caminho:
        arquivo = Path(caminho)

        if not arquivo.is_file():
            raise Exception(
                "PLANILHA NÃO ENCONTRADA:\n\n"
                + str(arquivo)
            )

        return arquivo

    if PLANILHA_ESCOLHIDA:
        arquivo = Path(PLANILHA_ESCOLHIDA)

        if not arquivo.is_file():
            raise Exception(
                "PLANILHA DO CONFIG NÃO ENCONTRADA:\n\n"
                + str(arquivo)
            )

        return arquivo

    arquivos = []

    arquivos.extend(
        sorted(PLANILHA_DIR.glob("*.xlsx"))
    )

    arquivos.extend(
        sorted(PLANILHA_DIR.glob("*.xlsm"))
    )

    # Remove temporários do Excel
    arquivos = [
        a
        for a in arquivos
        if not a.name.startswith("~$")
    ]

    if not arquivos:

        raise Exception(
            "NENHUMA PLANILHA ENCONTRADA:\n\n"
            + str(PLANILHA_DIR)
        )

    if len(arquivos) > 1:
        print()
        print(
            "VÁRIAS PLANILHAS ENCONTRADAS "
            "(usando a primeira; passe --planilha para escolher):"
        )
        for i, a in enumerate(arquivos):
            marca = " <<<" if i == 0 else ""
            print(
                " ",
                i + 1,
                "-",
                a.name + marca
        )

    return arquivos[0]


def encontrar_cabecalho(ws, padrao=True):
    """(linha, colunas) do cabeçalho; sem cabeçalho devolve o padrão A/C/D ou None (padrao=False)."""

    nomes = {
        chave("NOME"),
        chave("NOME DO ALUNO"),
        chave("NOME DO FORMANDO"),
        chave("NOME COMPLETO"),
        chave("ALUNO"),
        chave("FORMANDO"),
        chave("CLIENTE")
    }

    cursos = {
        chave("CURSO"),
        chave("FORMACAO"),
        chave("FORMAÇÃO")
    }

    planos = {
        chave("PLANO"),
        chave("ALBUM"),
        chave("ÁLBUM"),
        chave("PRODUTO"),
        chave("TIPO"),
        chave("TIPO DE ALBUM"),
        chave("TIPO DE ÁLBUM"),
        chave("TAMANHO")
    }

    extras = {
        chave("STATUS"): "status",
        chave("LOTE"): "lote",
    }

    for linha in range(
        1,
        min(
            ws.max_row,
            40
        ) + 1
    ):

        colunas = {}

        for coluna in range(
            1,
            ws.max_column + 1
        ):

            valor = chave(
                ws.cell(
                    linha,
                    coluna
                ).value
            )

            # Vale a primeira coluna de cada tipo (a planilha pode repetir "Cliente")
            if valor in nomes:
                colunas.setdefault("nome", coluna)

            if valor in cursos:
                colunas.setdefault("curso", coluna)

            if valor in planos:
                colunas.setdefault("plano", coluna)

            if valor in extras:
                colunas.setdefault(extras[valor], coluna)

        if (
            "nome" in colunas
            and
            "plano" in colunas
        ):

            return (
                linha,
                colunas
            )

    if not padrao:
        return None

    return (
        1,
        {
            "nome": 1,
            "curso": 3,
            "plano": 4
        }
    )


RE_EXTENSAO_IMAGEM = re.compile(r"(?i)\.(?:jpe?g|png|tiff?|psd)$")
RE_TOKEN_TAMANHO = re.compile(
    r"(?i)[\(\[\{\s_\-\.]*"
    r"(?<!\d)(?:20\s*[x×*]\s*20|24\s*[x×*]\s*34|34\s*[x×*]\s*24|30\s*[x×*]\s*20|20\s*[x×*]\s*30|30\s*[x×*]\s*22)(?:\s*cm)?"
    r"(?:\s*\(\d+\))?"
    r"[\)\]\}\s_\-\.]*(?!\d)"
)

PLANO_PADRAO_POR_TAMANHO = {
    "24X34": "STUDIO DAY BERLIM 24X34",
    "20X20": "ESSENCIAL BOOK 20X20",
    "30X22": "PRIME BOOK 30X22",
    "30X20": "PRIME BOOK 30X20",
}


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


def formatar_nome(nome):
    """Só a primeira letra de cada palavra maiúscula e sem o tamanho (24X34 / 20X20 / 30X20): "MARIA DA SILVA 24X34" -> "Maria Da Silva"."""
    texto = limpar_tamanho_do_nome(nome)
    return re.sub(r"[^\W\d_]+", lambda m: m.group(0)[0].upper() + m.group(0)[1:].lower(), texto)


def ler_planilha(caminho=None):

    arquivo = encontrar_planilha(caminho)

    print()
    print(
        "PLANILHA:",
        arquivo
    )

    workbook = openpyxl.load_workbook(
        arquivo,
        data_only=True
    )

    # Aba ativa primeiro; se ela não tiver cabeçalho de alunos, procura nas outras
    abas = [workbook.active] + [
        w for w in workbook.worksheets
        if w is not workbook.active
    ]

    ws = None

    for aba in abas:
        achado = encontrar_cabecalho(aba, padrao=False)
        if achado:
            ws = aba
            linha_cabecalho, colunas = achado
            break

    if ws is None:
        ws = workbook.active
        linha_cabecalho, colunas = encontrar_cabecalho(ws)

    print(
        "ABA:", ws.title,
        "| CABEÇALHO NA LINHA", linha_cabecalho
    )

    def celula(linha, campo):
        if not colunas.get(campo):
            return ""
        return normalizar(ws.cell(linha, colunas[campo]).value)

    registros = []

    for linha in range(
        linha_cabecalho + 1,
        ws.max_row + 1
    ):

        bruto_nome = normalizar(
            ws.cell(
                linha,
                colunas["nome"]
            ).value
        )
        tam_no_nome = extrair_tamanho_do_texto(bruto_nome)
        nome = formatar_nome(bruto_nome)

        curso = ""

        if colunas.get("curso"):

            curso = normalizar(
                ws.cell(
                    linha,
                    colunas["curso"]
                ).value
            )

        plano = normalizar(
            ws.cell(
                linha,
                colunas["plano"]
            ).value
        )
        if not plano and tam_no_nome:
            plano = PLANO_PADRAO_POR_TAMANHO.get(tam_no_nome, "")

        if not nome:
            continue

        registros.append({

            "linha":
                linha,

            "nome":
                nome,

            "curso":
                curso,

            "plano":
                plano,

            "status":
                celula(linha, "status"),

            "lote":
                celula(linha, "lote")
        })

    return registros


# ============================================================
# MODELO
# ============================================================

def encontrar_modelo(plano):

    plano_chave = chave(
        plano
    )

    if "30X22" in plano_chave.replace(" ", ""):
        return (
            "PRIME BOOK 30X22",
            MODELOS["PRIME BOOK 30X22"]
        )

    if "ESSENCIAL BOOK" in plano_chave:

        return (
            "ESSENCIAL BOOK 20X20",
            MODELO_ESSENCIAL_20
        )

    if "DIAMOND" in plano_chave and "24" in plano_chave:
        # Diamond 24x34 ainda nao tem PSD proprio — mapeia para Berlim
        return (
            "STUDIO DAY BERLIM 24X34",
            MODELOS["STUDIO DAY BERLIM 24X34"]
        )

    for nome, modelo in MODELOS.items():

        if chave(nome) == plano_chave:

            return (
                nome,
                modelo
            )

    for nome, modelo in MODELOS.items():

        nome_chave = chave(nome)

        if nome_chave in plano_chave or plano_chave in nome_chave:

            return (
                nome,
                modelo
            )

    # Atalhos parciais: "Moscou 30X20", "Berlim", "Prime Book"
    atalhos = [
        (("MOSCOU", "MOSKOW", "MOSCOVO"), "STUDIO DAY MOSCOU 30X20"),
        (("BERLIM", "BERLIN", "BERLINENSE"), "STUDIO DAY BERLIM 24X34"),
        (("PRIME",), "PRIME BOOK 30X20"),
        (("ESSENCIAL", "ESSENTIAL"), "ESSENCIAL BOOK 20X20"),
        (("DIAMOND", "DIAMANTE"), "STUDIO DAY BERLIM 24X34"),
    ]

    for palavras, modelo_nome in atalhos:
        if any(p in plano_chave for p in palavras):
            if modelo_nome == "ESSENCIAL BOOK 20X20":
                return modelo_nome, MODELO_ESSENCIAL_20
            return modelo_nome, MODELOS[modelo_nome]

    return (
        None,
        None
    )


# ============================================================
# FOTOS
# ============================================================

def _listar_fotos():

    if not PHOTOS_DIR.exists():

        raise Exception(
            "PASTA DE FOTOS NÃO ENCONTRADA:\n\n"
            + str(PHOTOS_DIR)
        )

    fotos = []

    for arquivo in PHOTOS_DIR.iterdir():

        if not arquivo.is_file():
            continue

        if arquivo.suffix.lower() not in [
            ".jpg",
            ".jpeg",
            ".png",
            ".tif",
            ".tiff"
        ]:
            continue

        fotos.append(arquivo)

    return fotos


def _chave_compacta(valor):
    """Remove espacos, _ e - para comparar Nome24X34 com Nome 24X34."""

    return re.sub(
        r"[\s_\-]+",
        "",
        chave(valor)
    )


def _variantes_busca_foto(nome, sufixo=""):
    """
    Gera chaves candidatas para achar a foto.
    Aceita abreviações/variações do nome + tamanhos:
    Nome24X34, Nome 24X34, Nome_24X34, Primeiro Ultimo 24x34, etc.
    """

    nomes = variantes_nome_pessoa(nome)
    # inclui também o nome original normalizado
    nomes = list(dict.fromkeys(
        [chave(nome), _chave_compacta(nome)] + nomes
    ))

    variantes = []

    if not sufixo:
        for n in nomes:
            variantes.append(chave(n))
            variantes.append(_chave_compacta(n))
        return list(dict.fromkeys(variantes))

    suf = normalizar(str(sufixo)).upper().replace(" ", "")

    tamanhos = list(dict.fromkeys([
        suf,
        suf.replace("X", "x"),
        suf.replace("x", "X"),
        # formas alternativas comuns
        suf.replace("24X34", "24X34"),
        suf.replace("20X20", "20X20"),
        # 34x24 invertido às vezes aparece
        "34X24" if suf in {"24X34", "24x34"} else suf,
        "34x24" if suf in {"24X34", "24x34"} else suf,
    ]))

    for n in nomes:
        nome_limpo = normalizar(n.replace(".", " "))
        for tam in tamanhos:
            for sep in ("", " ", "_", "-", "."):
                bruto = nome_limpo + sep + tam
                variantes.append(chave(bruto))
                variantes.append(_chave_compacta(bruto))
        # nome sozinho
        variantes.append(chave(nome_limpo))
        variantes.append(_chave_compacta(nome_limpo))

    return list(dict.fromkeys(v for v in variantes if v))


def _stem_chaves_foto(stem):
    """Chaves de comparacao do nome do arquivo."""

    return {
        chave(stem),
        _chave_compacta(stem),
    }


def _escolher_unica(encontrados, rotulo):

    if len(encontrados) == 1:
        return encontrados[0]

    if len(encontrados) > 1:
        raise Exception(
            "MAIS DE UMA FOTO ENCONTRADA"
            + ((" (" + rotulo + ")") if rotulo else "")
            + ":\n\n"
            + "\n".join(str(x) for x in encontrados)
        )

    return None


class FotoNaoEncontrada(Exception):
    """Nenhuma foto casou (diferente de foto ambígua)."""


# Tamanhos equivalentes (orientação invertida conta como o mesmo)
TAMANHOS_EQUIVALENTES = {
    "20X20": {"20X20"},
    "24X34": {"24X34", "34X24"},
    "34X24": {"24X34", "34X24"},
    "30X20": {"30X20", "20X30"},
    "20X30": {"30X20", "20X30"},
}


def _tamanhos_no_nome(stem):
    """Retorna os tamanhos (ex.: {'20X20'}) presentes no nome do arquivo."""

    compacto = _chave_compacta(stem)

    return set(
        a + "X" + b
        for a, b in re.findall(
            r"(?<!\d)(\d{2})X(\d{2})(?!\d)",
            compacto
        )
    )


def encontrar_foto(
    nome,
    sufixo="",
    flexivel=None
):
    """
    Com sufixo (24X34 / 20X20):
      - NUNCA aceita arquivo com outro tamanho no nome;
      - prefere arquivo com o tamanho pedido (ou invertido);
      - só usa foto sem tamanho no nome se não houver nenhuma
        com o tamanho pedido (e avisa).
    Sem sufixo: comportamento original.
    """

    if flexivel is None:
        flexivel = FOTO_MATCH_FLEXIVEL

    fotos = _listar_fotos()

    if not sufixo:
        return _buscar_foto_em(
            nome,
            "",
            flexivel,
            fotos
        )

    suf = _chave_compacta(sufixo)

    aceitos = TAMANHOS_EQUIVALENTES.get(
        suf,
        {suf}
    )

    com_tamanho = []
    sem_tamanho = []

    for arquivo in fotos:

        tamanhos = _tamanhos_no_nome(arquivo.stem)

        if not tamanhos:
            sem_tamanho.append(arquivo)
            continue

        # qualquer tamanho diferente do pedido descarta o arquivo
        if tamanhos - aceitos:
            continue

        com_tamanho.append(arquivo)

    # ----------------------------------------------------------
    # 1) Somente arquivos com o tamanho pedido
    # ----------------------------------------------------------

    erro_com_tamanho = None

    if com_tamanho:
        try:
            return _buscar_foto_em(
                nome,
                sufixo,
                flexivel,
                com_tamanho
            )
        except FotoNaoEncontrada as e:
            erro_com_tamanho = e

    # ----------------------------------------------------------
    # 2) Fallback: arquivo sem tamanho no nome
    # ----------------------------------------------------------

    try:
        escolhido = _buscar_foto_em(
            nome,
            sufixo,
            flexivel,
            sem_tamanho
        )
    except FotoNaoEncontrada:
        if erro_com_tamanho is not None:
            raise erro_com_tamanho
        raise FotoNaoEncontrada(
            "FOTO NÃO ENCONTRADA:\n\n"
            + nome
            + " "
            + suf
        )

    print(
        "  AVISO: SEM FOTO " + suf
        + " — USANDO FOTO SEM TAMANHO NO NOME:",
        escolhido.name
    )

    return escolhido


def _buscar_foto_em(
    nome,
    sufixo,
    flexivel,
    fotos
):

    variantes = _variantes_busca_foto(nome, sufixo)

    # ----------------------------------------------------------
    # 1) Match exato / compacto (Nome24X34 == Nome 24X34)
    # ----------------------------------------------------------

    encontrados = []

    for arquivo in fotos:
        stems = _stem_chaves_foto(arquivo.stem)
        if stems & set(variantes):
            encontrados.append(arquivo)

    # remove duplicatas preservando ordem
    vistos = set()
    unicos = []
    for a in encontrados:
        if a not in vistos:
            vistos.add(a)
            unicos.append(a)
    encontrados = unicos

    escolhido = _escolher_unica(
        encontrados,
        nome + (sufixo or "")
    )

    if escolhido is not None:
        if _chave_compacta(escolhido.stem) != _chave_compacta(
            nome + (sufixo or "")
        ) and chave(escolhido.stem) != chave(nome + (sufixo or "")):
            print(
                "  FOTO ENCONTRADA:",
                escolhido.name
            )
        return escolhido

    # ----------------------------------------------------------
    # 2) Com sufixo: tenta variações do nome (+ tamanho)
    # ----------------------------------------------------------

    nomes_var = variantes_nome_pessoa(nome)
    nome_chaves = set()
    for n in [nome] + nomes_var:
        nome_chaves.add(chave(n))
        nome_chaves.add(_chave_compacta(n))

    # remove chaves muito curtas (ex.: só "ANA") para não ambíguar
    nome_chaves = {
        k for k in nome_chaves
        if k and (len(k) >= 6 or " " in chave(k))
    } | {
        chave(nome),
        _chave_compacta(nome),
    }

    tamanhos_conhecidos = (
        "20X20", "24X34", "30X20", "34X24", "20X30"
    )

    def _stem_sem_tamanho(stem):
        c = _chave_compacta(stem)
        for t in tamanhos_conhecidos:
            if c.endswith(t):
                return c[:-len(t)]
        return c

    so_nome = []

    for arquivo in fotos:
        stems = _stem_chaves_foto(arquivo.stem)
        compact = _chave_compacta(arquivo.stem)
        base_stem = _stem_sem_tamanho(arquivo.stem)

        if stems & nome_chaves or base_stem in nome_chaves:
            so_nome.append(arquivo)
            continue

        for base in nome_chaves:
            if not base or len(base) < 4:
                continue
            if compact.startswith(base):
                resto = compact[len(base):]
                if (
                    not resto
                    or resto in tamanhos_conhecidos
                    or any(resto.startswith(t) for t in tamanhos_conhecidos)
                ):
                    so_nome.append(arquivo)
                    break
            # base (primeiro+ultimo) contido no stem
            if len(base) >= 8 and base in compact:
                so_nome.append(arquivo)
                break

    # remove duplicatas
    vistos = set()
    unicos = []
    for a in so_nome:
        if a not in vistos:
            vistos.add(a)
            unicos.append(a)
    so_nome = unicos

    if sufixo and so_nome:
        preferidos = [
            a for a in so_nome
            if _chave_compacta(sufixo) in _chave_compacta(a.stem)
            or (
                _chave_compacta(sufixo) == "24X34"
                and "34X24" in _chave_compacta(a.stem)
            )
        ]
        if len(preferidos) == 1:
            print(
                "  FOTO POR TAMANHO/VARIACAO:",
                preferidos[0].name
            )
            return preferidos[0]
        if len(preferidos) > 1:
            return _escolher_unica(
                preferidos,
                nome + sufixo
            )

    # Fallback: foto só com variação do nome (sem tamanho no arquivo)
    so_exato = [
        a for a in so_nome
        if _stem_sem_tamanho(a.stem) in nome_chaves
        or chave(a.stem) in nome_chaves
        or _chave_compacta(a.stem) in nome_chaves
    ]
    if len(so_exato) == 1:
        print(
            "  FOTO POR VARIACAO DO NOME:",
            so_exato[0].name
        )
        return so_exato[0]

    if not flexivel:

        raise FotoNaoEncontrada(
            "FOTO NÃO ENCONTRADA:\n\n"
            + nome
            + (sufixo or "")
        )

    # ----------------------------------------------------------
    # 3) Match flexivel: contem / similaridade / abreviações
    # ----------------------------------------------------------

    suf_c = _chave_compacta(sufixo) if sufixo else ""
    aproximados = []

    for arquivo in fotos:
        stem_c = _chave_compacta(arquivo.stem)
        stem_k = chave(arquivo.stem)
        base_stem = _stem_sem_tamanho(arquivo.stem)

        hit = False

        for base in nome_chaves:
            if not base or len(base) < 5:
                continue
            if (
                base == base_stem
                or base in stem_c
                or base_stem in base
                or stem_c.startswith(base)
            ):
                hit = True
                break

        if not hit:
            palavras_stem = " " + stem_k + " "
            palavras_nome = set(chave(nome).split())
            for var in variantes:
                if len(var) < 5:
                    continue
                if " " in var:
                    hit = (" " + var + " ") in palavras_stem or stem_c == _chave_compacta(var)
                elif var in palavras_nome:
                    # só o primeiro nome: "TIAGO" não pode casar com "SANTIAGO" nem com outro Tiago
                    hit = stem_c == var
                else:
                    hit = len(var) >= 8 and var in stem_c
                if hit:
                    break

        if hit:
            aproximados.append(arquivo)

    vistos = set()
    unicos = []
    for a in aproximados:
        if a not in vistos:
            vistos.add(a)
            unicos.append(a)
    aproximados = unicos

    if sufixo and aproximados:
        preferidos = [
            a for a in aproximados
            if suf_c in _chave_compacta(a.stem)
            or (
                suf_c == "24X34"
                and "34X24" in _chave_compacta(a.stem)
            )
        ]
        if len(preferidos) == 1:
            print(
                "  FOTO POR MATCH FLEXÍVEL:",
                preferidos[0].name
            )
            return preferidos[0]
        if preferidos:
            aproximados = preferidos

    if len(aproximados) == 1:
        print(
            "  FOTO POR MATCH FLEXÍVEL:",
            aproximados[0].name
        )
        return aproximados[0]

    if len(aproximados) > 1:

        raise Exception(
            "FOTO AMBÍGUA (match flexível):\n\n"
            + nome
            + (sufixo or "")
            + "\n\nCandidatos:\n"
            + "\n".join(str(x) for x in aproximados)
        )

    stems = {}
    for a in fotos:
        stems[chave(a.stem)] = a
        stems[_chave_compacta(a.stem)] = a
        stems[_stem_sem_tamanho(a.stem)] = a

    alvos = [
        _chave_compacta(nome + (sufixo or "")),
        _chave_compacta(nome),
    ] + [_chave_compacta(n) for n in nomes_var[:8]]

    proximos = []
    for alvo in alvos:
        if not alvo:
            continue
        proximos.extend(
            difflib.get_close_matches(
                alvo,
                list(stems.keys()),
                n=3,
                cutoff=0.72
            )
        )

    proximos = list(dict.fromkeys(proximos))

    if len(proximos) == 1:
        escolhido = stems[proximos[0]]
        print(
            "  FOTO POR SIMILARIDADE:",
            escolhido.name
        )
        return escolhido

    # se vários próximos apontam para o mesmo arquivo
    arquivos_prox = list(dict.fromkeys(
        stems[p] for p in proximos if p in stems
    ))
    if len(arquivos_prox) == 1:
        print(
            "  FOTO POR SIMILARIDADE:",
            arquivos_prox[0].name
        )
        return arquivos_prox[0]

    sugestao = ""

    if proximos:
        sugestao = (
            "\n\nSugestões:\n"
            + "\n".join(
                stems[p].name for p in proximos if p in stems
            )
        )

    raise FotoNaoEncontrada(
        "FOTO NÃO ENCONTRADA:\n\n"
        + nome
        + (sufixo or "")
        + sugestao
    )


def listar_fotos_orfas(registros):
    """Fotos na pasta que não casam com nenhum aluno+sufixo esperado."""

    try:
        fotos = _listar_fotos()
    except Exception:
        return []

    usados = set()

    for registro in registros:

        nome = registro["nome"]

        for n in [nome] + variantes_nome_pessoa(nome):
            for bruto in (
                n,
                n + "20X20",
                n + " 20X20",
                n + "_20X20",
                n + "24X34",
                n + " 24X34",
                n + "_24X34",
                n + "30X20",
                n + " 30X20",
                n + "34X24",
            ):
                usados.add(chave(bruto))
                usados.add(_chave_compacta(bruto))

    orfas = []

    for foto in fotos:
        stem_k = chave(foto.stem)
        stem_c = _chave_compacta(foto.stem)

        if stem_k in usados or stem_c in usados:
            continue

        ligado = False
        for u in usados:
            if not u or len(u) < 6:
                continue
            if u in stem_c or stem_c in u or u in stem_k or stem_k in u:
                ligado = True
                break

        if not ligado:
            orfas.append(foto)

    return orfas


# ============================================================
# CRIAR JOB 24X34
# ============================================================

# ============================================================
# LAYOUT DO NOME (mesmo cálculo da prévia da revisão)
# ============================================================

def layout_nome(nome, tamanho):
    """Linhas e escala (%) de cada linha do nome; None = corte automático do JSX."""

    try:
        pasta = str(Path(__file__).resolve().parent)
        if pasta not in sys.path:
            sys.path.insert(0, pasta)
        import recortar_fotos
        layout = recortar_fotos.calcular_layout_nome(
            tamanho, nome, recortar_fotos.ler_ajuste_nome(nome, tamanho)
        )
    except Exception as e:
        print("  AVISO: layout do nome indisponível (" + str(e) + "); usando o corte automático.")
        return None

    if layout.get("estoura"):
        print("  AVISO: NOME ESTOURA A ÁREA NA CAPA " + tamanho + " (ajuste na aba Revisão): " + nome)
    elif layout.get("reduzido_auto"):
        print("  NOME REDUZIDO PARA CABER (" + str(layout["escala1"]) + "% / "
              + str(layout["escala2"]) + "%): " + nome)
    return layout


def criar_job_berlim_24(
    registro
):

    modelo_nome = (
        "STUDIO DAY BERLIM 24X34"
    )

    modelo = MODELOS[
        modelo_nome
    ]

    nome = formatar_nome(registro["nome"])

    psd = (
        PSD_DIR /
        modelo["psd"]
    )

    if not psd.exists():

        raise Exception(
            "PSD BERLIM 24X34 NÃO ENCONTRADO:\n\n"
            + str(psd)
        )

    foto34 = encontrar_foto(
        nome,
        "24X34"
    )

    simbolo, aviso_curso = resolver_curso(
        registro.get("curso", "")
    )

    if aviso_curso:
        print(
            "  AVISO:",
            aviso_curso
            )

    pasta = OUTPUT_DIRS[
        "24X34"
    ]

    pasta.mkdir(
        parents=True,
        exist_ok=True
    )

    nome_saida = nome_seguro(
        nome
    )

    saida_jpg = ""

    if EXPORT_JPG:
        saida_jpg = str(
            pasta /
            (
                nome_saida.replace(" ", "-")
                + "_24X34.jpg"
            )
    )

    return {

        "nome":
            nome,

        "curso":
            registro["curso"],

        "curso_simbolo":
            simbolo,

        "cor_background":
            obter_cor_background(simbolo),

        "modelo":
            modelo_nome,

        "template":
            str(psd),

        "config":
            modelo,

        "nome_layout":
            layout_nome(nome, "24X34"),

        "fotos": {

            "foto34":
                str(foto34)
        },

        "saida_jpg":
            saida_jpg,

        "saida_psd":
            str(
                pasta /
                (
                    nome_saida
                    + "_24X34.psd"
                )
            ),

        "exportar_psd":
            EXPORT_PSD
    }


# ============================================================
# CRIAR JOB 20X20 ESSENCIAL BOOK
# ============================================================

def criar_job_essencial_20(
    registro
):

    nome = formatar_nome(registro["nome"])

    modelo = MODELO_ESSENCIAL_20

    psd = (
        PSD_DIR /
        modelo["psd"]
    )

    if not psd.exists():

        raise Exception(
            "PSD ESSENCIAL BOOK 20X20 NÃO ENCONTRADO:\n\n"
            + str(psd)
        )

    foto = encontrar_foto(
        nome,
        "20X20"
    )

    simbolo, aviso_curso = resolver_curso(
        registro.get("curso", "")
    )

    if aviso_curso:
        print(
            "  AVISO:",
            aviso_curso
            )

    pasta = OUTPUT_DIRS[
        "20X20"
    ]

    pasta.mkdir(
        parents=True,
        exist_ok=True
    )

    nome_saida = nome_seguro(
        nome
    )

    saida_jpg = ""

    if EXPORT_JPG:
        saida_jpg = str(
            pasta /
            (
                nome_saida.replace(" ", "-")
                + "_20X20.jpg"
            )
    )

    return {

        "nome":
            nome,

        "curso":
            registro["curso"],

        "curso_simbolo":
            simbolo,

        "cor_background":
            obter_cor_background(simbolo),

        "modelo":
            "ESSENCIAL BOOK 20X20",

        "template":
            str(psd),

        "config":
            modelo,

        "nome_layout":
            layout_nome(nome, "20X20"),

        "fotos": {

            "foto":
                str(foto)
        },

        "saida_jpg":
            saida_jpg,

        "saida_psd":
            str(
                pasta /
                (
                    nome_saida
                    + "_20X20.psd"
                )
            ),

        "exportar_psd":
            EXPORT_PSD
    }


# ============================================================
# CRIAR JOB 30X20
# ============================================================

def criar_job_30(
    registro,
    modelo_nome,
    modelo
):

    nome = formatar_nome(registro["nome"])

    psd = (
        PSD_DIR /
        modelo["psd"]
    )

    if not psd.exists():

        raise Exception(
            "PSD NÃO ENCONTRADO:\n\n"
            + str(psd)
        )

    foto = encontrar_foto(
        nome
    )

    simbolo, aviso_curso = resolver_curso(
        registro.get("curso", "")
    )

    if aviso_curso:
        print(
            "  AVISO:",
            aviso_curso
            )

    pasta = OUTPUT_DIRS[
        "30X20"
    ]

    pasta.mkdir(
        parents=True,
        exist_ok=True
    )

    nome_saida = nome_seguro(
        nome
    )

    saida_jpg = ""

    if EXPORT_JPG:
        saida_jpg = str(
            pasta /
            (
                nome_saida.replace(" ", "-")
                + "_30X20.jpg"
            )
    )

    return {

        "nome":
            nome,

        "curso":
            registro["curso"],

        "curso_simbolo":
            simbolo,

        "cor_background":
            obter_cor_background(simbolo),

        "modelo":
            modelo_nome,

        "template":
            str(psd),

        "config":
            modelo,

        "nome_layout":
            layout_nome(nome, "30X20"),

        "fotos": {

            "foto":
                str(foto)
        },

        "saida_jpg":
            saida_jpg,

        "saida_psd":
            str(
                pasta /
                (
                    nome_saida
                    + "_30X20.psd"
                )
            ),

        "exportar_psd":
            EXPORT_PSD
    }


def criar_job_30x22(
    registro,
    lado
):
    """Job de um lado do 30X22: lado = "FRENTE" ou "VERSO"."""

    modelo = MODELOS["PRIME BOOK 30X22"]
    config = modelo["frente" if lado == "FRENTE" else "verso"]

    nome = formatar_nome(registro["nome"])

    psd = PSD_DIR / config["psd"]

    if not psd.exists():
        raise Exception(
            "PSD NÃO ENCONTRADO:\n\n"
            + str(psd)
        )

    fotos = {}
    if lado == "FRENTE":
        fotos["foto"] = str(encontrar_foto(nome))

    simbolo, aviso_curso = resolver_curso(
        registro.get("curso", "")
    )

    if aviso_curso and lado == "FRENTE":
        print("  AVISO:", aviso_curso)

    pasta = OUTPUT_DIRS["30X22"]
    pasta.mkdir(parents=True, exist_ok=True)

    nome_saida = nome_seguro(nome)
    sufixo = "_30X22_" + lado

    saida_jpg = ""
    if EXPORT_JPG:
        saida_jpg = str(pasta / (nome_saida.replace(" ", "-") + sufixo + ".jpg"))

    return {
        "nome": nome,
        "curso": registro["curso"],
        "curso_simbolo": simbolo,
        "cor_background": obter_cor_background(simbolo),
        "modelo": "PRIME BOOK 30X22 " + lado,
        "template": str(psd),
        "config": config,
        "nome_layout": layout_nome(nome, "30X22") if lado == "FRENTE" else None,
        "fotos": fotos,
        "saida_jpg": saida_jpg,
        "saida_psd": str(pasta / (nome_saida + sufixo + ".psd")),
        "exportar_psd": EXPORT_PSD
    }


# ============================================================
# JSX DO PHOTOSHOP
# ============================================================

JSX = r'''
// ==========================================================
// UTILITÁRIOS
// ==========================================================

// Etapa atual, usada nas mensagens de erro devolvidas ao Python.
var ETAPA_ATUAL =
    "INÍCIO";


function etapa(nome) {

    ETAPA_ATUAL =
        nome;
}


function erro(mensagem) {

    throw new Error(
        mensagem
    );
}


// Grava o resultado num arquivo de status (pasta temp local),
// para o Python saber a mensagem exata mesmo se o
// DoJavaScriptFile não devolver o erro.
function gravarStatus(texto) {

    try {

        if (
            typeof JOB === "undefined"
            ||
            !JOB
            ||
            !JOB.arquivo_status
        ) {

            return;
        }

        var arquivoStatus =
            new File(
                JOB.arquivo_status
            );

        arquivoStatus.encoding =
            "UTF-8";

        if (
            arquivoStatus.open("w")
        ) {

            arquivoStatus.write(
                texto
            );

            arquivoStatus.close();
        }

    }

    catch (_) {}
}


function normalizar(valor) {

    if (valor === undefined || valor === null) {
        return "";
    }

    return String(valor)
        .replace(/^\s+|\s+$/g, "")
        .replace(/\s+/g, " ");
}


// ==========================================================
// BUSCA RECURSIVA
// ==========================================================

function buscarRecursivo(
    container,
    nome,
    resultado
) {

    if (!resultado) {
        resultado = [];
    }

    if (!container.layers) {
        return resultado;
    }

    for (
        var i = 0;
        i < container.layers.length;
        i++
    ) {

        var camada =
            container.layers[i];

        if (
            camada.name ===
            nome
        ) {

            resultado.push(
                camada
            );
        }

        if (
            camada.typename ===
            "LayerSet"
        ) {

            buscarRecursivo(
                camada,
                nome,
                resultado
            );
        }
    }

    return resultado;
}


// ==========================================================
// ENCONTRAR CAMINHO
// ==========================================================

function encontrarCaminho(
    documento,
    caminho
) {

    var atual =
        documento;

    for (
        var i = 0;
        i < caminho.length;
        i++
    ) {

        var nome =
            caminho[i];

        var encontrados =
            buscarRecursivo(
                atual,
                nome,
                []
            );

        if (
            encontrados.length === 0
        ) {

            erro(
                "CAMADA/GRUPO NÃO ENCONTRADO:\n\n"
                + caminho.join(" > ")
            );
        }

        if (
            encontrados.length > 1
        ) {

            // Se há duplicidade, tenta usar
            // o primeiro dentro do caminho atual.
            atual =
                encontrados[0];

        } else {

            atual =
                encontrados[0];
        }
    }

    return atual;
}


// ==========================================================
// BUSCAR ÚNICO
// ==========================================================

function buscarUnico(
    container,
    nomes,
    descricao
) {

    var encontrados = [];

    for (
        var i = 0;
        i < nomes.length;
        i++
    ) {

        buscarRecursivo(
            container,
            nomes[i],
            encontrados
        );
    }

    if (
        encontrados.length === 0
    ) {

        erro(
            descricao
            + " NÃO ENCONTRADO:\n\n"
            + nomes.join("\n")
        );
    }

    return encontrados[0];
}


// ==========================================================
// BUSCAR TEXTOS
// ==========================================================

function buscarTextos(
    container,
    resultado
) {

    if (!resultado) {
        resultado = [];
    }

    if (!container.layers) {
        return resultado;
    }

    for (
        var i = 0;
        i < container.layers.length;
        i++
    ) {

        var camada =
            container.layers[i];

        if (
            camada.typename ===
            "ArtLayer"
        ) {

            if (
                camada.kind ===
                LayerKind.TEXT
            ) {

                resultado.push(
                    camada
                );
            }
        }

        if (
            camada.typename ===
            "LayerSet"
        ) {

            buscarTextos(
                camada,
                resultado
            );
        }
    }

    return resultado;
}


// ==========================================================
// FOTO
// ==========================================================

function substituirFoto(
    documento,
    caminhoFoto,
    caminhoCamada,
    redimensionarFoto,
    ancoragem
) {

    app.activeDocument =
        documento;

    var camada =
        encontrarCaminho(
            documento,
            caminhoCamada
        );

    if (
        camada.kind !==
        LayerKind.SMARTOBJECT
    ) {

        erro(
            "A CAMADA FOTO NÃO É SMART OBJECT:\n\n"
            + camada.name
        );
    }

    var arquivo =
        new File(
            caminhoFoto
        );

    if (
        !arquivo.exists
    ) {

        erro(
            "ARQUIVO DA FOTO NÃO EXISTE:\n\n"
            + caminhoFoto
        );
    }

    var alvoLeft =
        camada.bounds[0].as("px");

    var alvoTop =
        camada.bounds[1].as("px");

    var alvoRight =
        camada.bounds[2].as("px");

    var alvoBottom =
        camada.bounds[3].as("px");

    var alvoWidth =
        alvoRight -
        alvoLeft;

    var alvoHeight =
        alvoBottom -
        alvoTop;

    var alvoCenterX =
        (alvoLeft + alvoRight) / 2;

    var alvoCenterY =
        (alvoTop + alvoBottom) / 2;

    var molde =
        redimensionarFoto
            ? areaDoMolde(camada)
            : null;

    documento.activeLayer =
        camada;

    etapa(
        "SUBSTITUIR FOTO NO SMART OBJECT (placedLayerReplaceContents): "
        + caminhoFoto
    );

    var desc =
        new ActionDescriptor();

    desc.putPath(
        charIDToTypeID("null"),
        arquivo
    );

    executeAction(
        stringIDToTypeID(
            "placedLayerReplaceContents"
        ),
        desc,
        DialogModes.NO
    );


    // ======================================================
    // 20X20
    //
    // NÃO REDIMENSIONAR.
    // As fotos do 20x20 já vêm no molde certo:
    // só troca o conteúdo do smart object.
    // ======================================================

    if (
        !redimensionarFoto
    ) {

        ajustarFotoAoAlvo(
            documento.activeLayer,
            alvoLeft,
            alvoTop,
            alvoWidth,
            alvoHeight,
            caminhoFoto
        );

        camada.visible =
            true;

        return;
    }


    etapa(
        "ENCAIXAR FOTO NO MOLDE (base centralizada, aumenta até o topo)"
    );

    encaixarFotoNoMolde(
        documento,
        camada,
        molde,
        caminhoFoto
    );
}


// ==========================================================
// FOTO 24X34 (BERLIM)
//
// Mesmo sistema do 20x20: só troca o conteúdo, sem redimensionar.
// ==========================================================

function substituirFotoBerlim24(
    documento,
    caminhoFoto,
    caminhoCamada
) {

    app.activeDocument =
        documento;

    var camada =
            encontrarCaminho(
                documento,
            caminhoCamada
        );

    if (
        camada.kind !==
        LayerKind.SMARTOBJECT
    ) {

        erro(
            "A CAMADA FOTO 24X34 NÃO É SMART OBJECT:\n\n"
            + camada.name
        );
    }

    var arquivo =
        new File(
            caminhoFoto
        );

    if (
        !arquivo.exists
    ) {

        erro(
            "ARQUIVO DA FOTO 24X34 NÃO EXISTE:\n\n"
            + caminhoFoto
        );
    }


    // ======================================================
    // ÁREA ORIGINAL DA CAMADA FOTO
    // ======================================================

    var alvoLeft =
        camada.bounds[0].as("px");

    var alvoTop =
        camada.bounds[1].as("px");

    var alvoRight =
        camada.bounds[2].as("px");

    var alvoBottom =
        camada.bounds[3].as("px");

    var alvoWidth =
        alvoRight -
        alvoLeft;

    var alvoHeight =
        alvoBottom -
        alvoTop;

    var alvoCenterX =
        (alvoLeft + alvoRight) / 2;

    var alvoCenterY =
        (alvoTop + alvoBottom) / 2;


    // ======================================================
    // TROCAR CONTEÚDO
    // ======================================================

    documento.activeLayer =
        camada;

    etapa(
        "SUBSTITUIR FOTO 24X34 NO SMART OBJECT (placedLayerReplaceContents): "
        + caminhoFoto
    );

    var desc =
        new ActionDescriptor();

    desc.putPath(
        charIDToTypeID("null"),
        arquivo
    );

    executeAction(
        stringIDToTypeID(
            "placedLayerReplaceContents"
        ),
        desc,
        DialogModes.NO
    );


    // Igual ao 20x20: a foto já vem no molde certo,
    // só troca o conteúdo do smart object.
    ajustarFotoAoAlvo(
        documento.activeLayer,
        alvoLeft,
        alvoTop,
        alvoWidth,
        alvoHeight,
        caminhoFoto
    );

    camada.visible =
        true;
}


// ==========================================================
// A foto nova com outro dpi (ex.: 72 dpi num smart object de
// 300 dpi) entra ampliada pelo Photoshop. Devolve a camada ao
// tamanho e à posição da foto original do modelo.
// ==========================================================

function ajustarFotoAoAlvo(
    camada,
    alvoLeft,
    alvoTop,
    alvoWidth,
    alvoHeight,
    caminhoFoto
) {

    var largura =
        camada.bounds[2].as("px") - camada.bounds[0].as("px");

    var altura =
        camada.bounds[3].as("px") - camada.bounds[1].as("px");

    if (
        largura <= 0
        || altura <= 0
    ) {

        erro(
            "A FOTO FICOU COM DIMENSÕES INVÁLIDAS:\n\n"
            + caminhoFoto
        );
    }

    if (
        Math.abs(largura - alvoWidth) / alvoWidth > 0.005
        || Math.abs(altura - alvoHeight) / alvoHeight > 0.005
    ) {

        var escala =
            Math.max(
                alvoWidth / largura,
                alvoHeight / altura
            ) * 100;

        etapa(
            "AJUSTAR FOTO AO TAMANHO DO MODELO (" + Math.round(escala) + "%)"
        );

        camada.resize(
            escala,
            escala,
            AnchorPosition.MIDDLECENTER
        );
    }

    var centroX =
        (camada.bounds[0].as("px") + camada.bounds[2].as("px")) / 2;

    var centroY =
        (camada.bounds[1].as("px") + camada.bounds[3].as("px")) / 2;

    camada.translate(
        alvoLeft + alvoWidth / 2 - centroX,
        alvoTop + alvoHeight / 2 - centroY
    );
}


// ==========================================================
// MOLDE DA FOTO
//
// O molde é a área onde a foto APARECE (linha do visível),
// nunca a borda do projeto nem o tamanho da foto antiga.
// É a interseção de:
//   - a área do documento (canvas);
//   - a camada base, se a FOTO estiver em máscara de corte;
//   - a máscara (pixel ou vetorial) da própria camada FOTO;
//   - as máscaras dos grupos onde a FOTO está.
// Sem nenhuma máscara, usa a camada FOTO original
// recortada pelo documento.
// ==========================================================

function intersecaoArea(
    area,
    l,
    t,
    r,
    b
) {

    if (
        r - l <= 0
        || b - t <= 0
    ) {
        return false;
    }

    var nl = Math.max(area.left, l);
    var nt = Math.max(area.top, t);
    var nr = Math.min(area.right, r);
    var nb = Math.min(area.bottom, b);

    if (
        nr - nl <= 0
        || nb - nt <= 0
    ) {
        return false;
    }

    area.left = nl;
    area.top = nt;
    area.right = nr;
    area.bottom = nb;

    return true;
}


function camadaAtivaTem(
    propriedade
) {

    try {

        var ref =
            new ActionReference();

        ref.putEnumerated(
            charIDToTypeID("Lyr "),
            charIDToTypeID("Ordn"),
            charIDToTypeID("Trgt")
        );

        var d =
            executeActionGet(ref);

        var chave =
            stringIDToTypeID(propriedade);

        return d.hasKey(chave)
            && d.getBoolean(chave);

    } catch (e) {}

    return false;
}


function selecaoDaMascaraAtiva(
    documento,
    vetorial
) {

    var desc =
        new ActionDescriptor();

    var refSel =
        new ActionReference();

    refSel.putProperty(
        charIDToTypeID("Chnl"),
        charIDToTypeID("fsel")
    );

    desc.putReference(
        charIDToTypeID("null"),
        refSel
    );

    var refMask =
        new ActionReference();

    if (
        vetorial
    ) {

        refMask.putEnumerated(
            charIDToTypeID("Path"),
            charIDToTypeID("Path"),
            stringIDToTypeID("vectorMask")
        );

        desc.putReference(
            charIDToTypeID("T   "),
            refMask
        );

        desc.putInteger(
            charIDToTypeID("Vrsn"),
            1
        );

        desc.putBoolean(
            stringIDToTypeID("vectorMaskParams"),
            true
        );

    } else {

        refMask.putEnumerated(
            charIDToTypeID("Chnl"),
            charIDToTypeID("Chnl"),
            charIDToTypeID("Msk ")
        );

        desc.putReference(
            charIDToTypeID("T   "),
            refMask
        );
    }

    executeAction(
        charIDToTypeID("setd"),
        desc,
        DialogModes.NO
    );

    var b =
        documento.selection.bounds;

    var limites = [
        b[0].as("px"),
        b[1].as("px"),
        b[2].as("px"),
        b[3].as("px")
    ];

    documento.selection.deselect();

    return limites;
}


function limitarPelasMascaras(
    documento,
    camada,
    area
) {

    var usou =
        false;

    try {

        documento.activeLayer =
            camada;

    } catch (e0) {
        return false;
    }

    var tipos = [
        ["hasUserMask", false],
        ["hasVectorMask", true]
    ];

    for (
        var i = 0;
        i < tipos.length;
        i++
    ) {

        if (
            !camadaAtivaTem(tipos[i][0])
        ) {
            continue;
        }

        try {

            var l =
                selecaoDaMascaraAtiva(
                    documento,
                    tipos[i][1]
                );

            if (
                intersecaoArea(
                    area,
                    l[0],
                    l[1],
                    l[2],
                    l[3]
                )
            ) {
                usou = true;
            }

        } catch (e1) {

            try {
                documento.selection.deselect();
            } catch (e2) {}
        }
    }

    return usou;
}


function areaDoMolde(
    camada
) {

    var documento =
        app.activeDocument;

    var area = {
        left: 0,
        top: 0,
        right: documento.width.as("px"),
        bottom: documento.height.as("px")
    };

    var achouLimite =
        false;


    // 1) Base da máscara de corte
    var base =
        baseDoClipping(
            camada
        );

    if (
        base
    ) {

        try {

            if (
                intersecaoArea(
                    area,
                    base.bounds[0].as("px"),
                    base.bounds[1].as("px"),
                    base.bounds[2].as("px"),
                    base.bounds[3].as("px")
                )
            ) {
                achouLimite = true;
            }

        } catch (e1) {}

        if (
            limitarPelasMascaras(
                documento,
                base,
                area
            )
        ) {
            achouLimite = true;
        }
    }


    // 2) Máscaras da própria FOTO
    if (
        limitarPelasMascaras(
            documento,
            camada,
            area
        )
    ) {
        achouLimite = true;
    }


    // 3) Máscaras dos grupos acima da FOTO
    var pai =
        camada.parent;

    while (
        pai
        && pai.typename === "LayerSet"
    ) {

        if (
            limitarPelasMascaras(
                documento,
                pai,
                area
            )
        ) {
            achouLimite = true;
        }

        pai =
            pai.parent;
    }


    // 4) Sem máscara: a própria FOTO original, só a parte no documento
    if (
        !achouLimite
    ) {

        try {

            intersecaoArea(
                area,
                camada.bounds[0].as("px"),
                camada.bounds[1].as("px"),
                camada.bounds[2].as("px"),
                camada.bounds[3].as("px")
            );

        } catch (e3) {}
    }

    documento.activeLayer =
        camada;

    etapa(
        "MOLDE VISIVEL: "
        + Math.round(area.left) + ", "
        + Math.round(area.top) + " ate "
        + Math.round(area.right) + ", "
        + Math.round(area.bottom)
        + " (" + Math.round(area.right - area.left)
        + " x " + Math.round(area.bottom - area.top) + " px)"
    );

    return area;
}


function baseDoClipping(
    camada
) {

    try {

        if (
            !camada.grouped
        ) {
            return null;
        }

        var irmaos =
            camada.parent.layers;

        var indice = -1;

        for (
            var i = 0;
            i < irmaos.length;
            i++
        ) {

            if (
                irmaos[i] === camada
            ) {
                indice = i;
                break;
            }
        }

        if (
            indice < 0
        ) {
            return null;
        }

        // Camadas abaixo têm índice maior.
        // A base é a primeira abaixo que NÃO está em clipping.
        for (
            var j = indice + 1;
            j < irmaos.length;
            j++
        ) {

            var candidata =
                irmaos[j];

            if (
                candidata.typename === "ArtLayer"
                && candidata.grouped
            ) {
                continue;
            }

            return candidata;
        }

    } catch (e) {}

    return null;
}


// ==========================================================
// ENCAIXE PADRÃO DA FOTO NO MOLDE
//
// Igual para todos os modelos:
// 1) nunca reduz e nunca distorce (mesma escala em X e Y);
// 2) aumenta até cobrir a largura E a altura do molde,
//    com uma pequena sobra de segurança;
// 3) alinha a BASE da foto com a base do molde,
//    centralizada na horizontal;
// 4) a foto sobe até passar do topo do molde,
//    sem deixar borda em nenhum lado.
// ==========================================================

var SOBRA_MOLDE_PX = 4;

var ESCALA_MAXIMA_FOTO = 400;

function encaixarFotoNoMolde(
    documento,
    camada,
    molde,
    caminhoFoto
) {

    app.activeDocument =
        documento;

    documento.activeLayer =
        camada;

    var moldeLeft = molde.left - SOBRA_MOLDE_PX;
    var moldeTop = molde.top - SOBRA_MOLDE_PX;
    var moldeRight = molde.right + SOBRA_MOLDE_PX;
    var moldeBottom = molde.bottom + SOBRA_MOLDE_PX;

    var moldeWidth =
        moldeRight - moldeLeft;

    var moldeHeight =
        moldeBottom - moldeTop;

    var moldeCenterX =
        (moldeLeft + moldeRight) / 2;


    var fotoWidth =
        camada.bounds[2].as("px") -
        camada.bounds[0].as("px");

    var fotoHeight =
        camada.bounds[3].as("px") -
        camada.bounds[1].as("px");

    if (
        fotoWidth <= 0
        || fotoHeight <= 0
    ) {

        erro(
            "A FOTO FICOU COM DIMENSÕES INVÁLIDAS:\n\n"
            + caminhoFoto
        );
    }


    // Escala para cobrir o molde inteiro; nunca abaixo de 100%
    var escala =
        Math.max(
            100,
            (moldeWidth / fotoWidth) * 100,
            (moldeHeight / fotoHeight) * 100
        );

    if (
        !isFinite(escala)
        || escala <= 0
    ) {

        erro(
            "NÃO FOI POSSÍVEL CALCULAR A ESCALA DA FOTO:\n\n"
            + caminhoFoto
        );
    }

    if (
        escala > ESCALA_MAXIMA_FOTO
    ) {

        erro(
            "ESCALA DA FOTO MUITO ALTA (" + Math.round(escala) + "%) — CANCELADO PARA NÃO ENCHER O DISCO DE TRABALHO.\n\n"
            + "MOLDE: " + Math.round(moldeWidth) + " x " + Math.round(moldeHeight) + " px\n"
            + "FOTO: " + Math.round(fotoWidth) + " x " + Math.round(fotoHeight) + " px\n\n"
            + caminhoFoto
        );
    }

    if (
        escala > 100.01
    ) {

        camada.resize(
            escala,
            escala,
            AnchorPosition.BOTTOMCENTER
        );
    }


    // Base da foto na base do molde, centralizada na horizontal
    var fotoLeft = camada.bounds[0].as("px");
    var fotoRight = camada.bounds[2].as("px");
    var fotoBottom = camada.bounds[3].as("px");

    camada.translate(
        moldeCenterX - (fotoLeft + fotoRight) / 2,
        moldeBottom - fotoBottom
    );


    // Garantia final: se por arredondamento não chegou ao topo,
    // aumenta mais um pouco a partir da base.
    var topoAtual =
        camada.bounds[1].as("px");

    if (
        topoAtual > moldeTop
    ) {

        var baseAtual =
        camada.bounds[3].as("px");

        var ajuste =
            ((baseAtual - moldeTop) / (baseAtual - topoAtual)) * 100;

        if (
            !isFinite(ajuste)
            || ajuste > 105
        ) {

            erro(
                "A FOTO NÃO FICOU ALINHADA AO MOLDE (ajuste de " + Math.round(ajuste) + "%).\n\n"
                + caminhoFoto
            );
        }

        camada.resize(
            ajuste,
            ajuste,
            AnchorPosition.BOTTOMCENTER
        );
    }

    camada.visible =
        true;
}


// ==========================================================
// FOTO DENTRO DO SMART OBJECT
// ==========================================================

function colarFotoNoSmartObject(
    documento,
    caminhoFoto,
    caminhoCamada
) {

    var camadaSmart =
        encontrarCaminho(
            documento,
            caminhoCamada
        );

    var documentoSmart =
        abrirSmartObject(
            camadaSmart
        );

    var arquivo =
        new File(
            caminhoFoto
        );

    app.activeDocument =
        documentoSmart;

    etapa(
        "COLOCAR FOTO DENTRO DO SMART OBJECT: "
        + caminhoFoto
    );

    var descColocar =
        new ActionDescriptor();

    descColocar.putPath(
        charIDToTypeID("null"),
        arquivo
    );

    descColocar.putEnumerated(
        charIDToTypeID("FTcs"),
        charIDToTypeID("QCSt"),
        charIDToTypeID("Qcsa")
    );

    descColocar.putUnitDouble(
        charIDToTypeID("Hrzn"),
        charIDToTypeID("#Pxl"),
        0
    );

    descColocar.putUnitDouble(
        charIDToTypeID("Vrtc"),
        charIDToTypeID("#Pxl"),
        0
    );

    executeAction(
        charIDToTypeID("Plc "),
        descColocar,
        DialogModes.NO
    );

    var camadaFoto =
        documentoSmart.activeLayer;

    camadaFoto.name =
        "FOTO NOVA";

    camadaFoto.visible =
        true;

    camadaFoto.move(
        documentoSmart,
        ElementPlacement.PLACEATBEGINNING
    );

    var fotoWidth =
        camadaFoto.bounds[2].as("px") -
        camadaFoto.bounds[0].as("px");

    var fotoHeight =
        camadaFoto.bounds[3].as("px") -
        camadaFoto.bounds[1].as("px");

    var alvoWidth =
        documentoSmart.width.as("px");

    var alvoHeight =
        documentoSmart.height.as("px");

    var escala =
        Math.max(
            (alvoWidth / fotoWidth) * 100,
            (alvoHeight / fotoHeight) * 100
        );

    camadaFoto.resize(
        escala,
        escala,
        AnchorPosition.TOPCENTER
    );

    var fotoLeft =
        camadaFoto.bounds[0].as("px");

    var fotoTop =
        camadaFoto.bounds[1].as("px");

    var fotoRight =
        camadaFoto.bounds[2].as("px");

    var fotoBottom =
        camadaFoto.bounds[3].as("px");

    camadaFoto.translate(
        (alvoWidth - (fotoLeft + fotoRight)) / 2,
        -fotoTop
    );

    app.activeDocument =
        documentoSmart;

    etapa(
        "SALVAR SMART OBJECT (.psb temporário do Photoshop)"
    );

    var descSalvar =
        new ActionDescriptor();

    descSalvar.putBoolean(
        stringIDToTypeID("asCopy"),
        false
    );

    executeAction(
        stringIDToTypeID("save"),
        descSalvar,
        DialogModes.NO
    );

    documentoSmart.close(
        SaveOptions.DONOTSAVECHANGES
    );

    app.activeDocument =
        documento;

    etapa(
        "VERIFICAR FOTO DENTRO DO SMART OBJECT"
    );

    var documentoVerificacao =
        abrirSmartObject(
            camadaSmart
        );

    try {

        validarFotoNoSmartObject(
            documentoVerificacao
        );

    } finally {

        documentoVerificacao.close(
            SaveOptions.DONOTSAVECHANGES
        );

        app.activeDocument =
            documento;
    }

    documento.activeLayer =
        camadaSmart;

    camadaSmart.visible =
        true;
}


function validarFotoNoSmartObject(
    documentoSmart
) {

    var camadaFoto;

    try {

        camadaFoto =
            documentoSmart.artLayers.getByName(
                "FOTO NOVA"
            );

    } catch (e) {

        erro(
            "A FOTO NOVA NÃO FOI ENCONTRADA DENTRO DO SMART OBJECT."
        );
    }

    if (
        !camadaFoto.visible
    ) {

        erro(
            "A FOTO NOVA ESTÁ OCULTA DENTRO DO SMART OBJECT."
        );
    }

    var largura =
        camadaFoto.bounds[2].as("px") -
        camadaFoto.bounds[0].as("px");

    var altura =
        camadaFoto.bounds[3].as("px") -
        camadaFoto.bounds[1].as("px");

    if (
        largura <= 0
        || altura <= 0
    ) {

        erro(
            "A FOTO NOVA FICOU COM DIMENSÕES INVÁLIDAS DENTRO DO SMART OBJECT."
        );
    }
}


// ==========================================================
// REDIMENSIONAMENTO EM ETAPAS
// ==========================================================

function redimensionarEmEtapas(
    camada,
    percentual,
    ancora
) {

    if (
        !isFinite(percentual)
        ||
        percentual <= 0
    ) {

        erro(
            "ESCALA INVÁLIDA PARA A FOTO."
        );
    }

    if (!ancora) {

        ancora =
            AnchorPosition.MIDDLECENTER;
    }

    if (
        percentual <= 100.01
    ) {

        return;
    }

    app.activeDocument.activeLayer =
        camada;

    var fatorTotal =
        percentual / 100.0;

    var fatorAtual =
        1.0;

    var limiteEtapa =
        1.20;

    var contador =
        0;

    var maxEtapas =
        100;


    while (
        fatorAtual <
            fatorTotal - 0.0001
        &&
        contador <
            maxEtapas
    ) {

        var fatorRestante =
            fatorTotal /
            fatorAtual;

        var fatorEtapa =
            Math.min(
                limiteEtapa,
                fatorRestante
            );

        var percentualEtapa =
            fatorEtapa * 100.0;

        if (
            percentualEtapa <=
            100.0001
        ) {

            break;
        }

        app.activeDocument.activeLayer =
            camada;

        camada.resize(
            percentualEtapa,
            percentualEtapa,
            ancora
        );

        fatorAtual *=
            fatorEtapa;

        contador++;
    }


    if (
        contador >= maxEtapas
        &&
        fatorAtual <
            fatorTotal - 0.001
    ) {

        erro(
            "NÃO FOI POSSÍVEL CONCLUIR O REDIMENSIONAMENTO EM ETAPAS."
        );
    }
}


// ==========================================================
// DIVIDIR NOME EM DUAS LINHAS (EQUILIBRADO)
// ==========================================================
//
// Evita o erro antigo: ao achar "de/da/do", colocava só o
// último sobrenome na 2ª linha e todo o resto na 1ª.
//
// Agora escolhe o ponto de corte que:
// 1) não deixa preposição no fim da 1ª linha
// 2) mantém preposição junto do sobrenome seguinte
// 3) equilibra o comprimento das duas linhas
//

function ehPreposicaoNome(palavra) {

    var p =
        normalizar(palavra).toUpperCase();

    return (
        p === "DA"
        || p === "DE"
        || p === "DO"
        || p === "DAS"
        || p === "DOS"
        || p === "E"
        || p === "DI"
        || p === "DU"
        || p === "DEL"
        || p === "VAN"
        || p === "VON"
    );
}


function juntarPalavras(lista, inicio, fim) {

    var texto = "";

    for (
        var i = inicio;
        i < fim;
        i++
    ) {

        if (texto !== "") {
            texto += " ";
        }

        texto += lista[i];
    }

    return texto;
}


// Quantos nomes de verdade (sem preposição) há no trecho
function contarNomes(lista, inicio, fim) {

    var total = 0;

    for (
        var i = inicio;
        i < fim;
        i++
    ) {

        if (
            !ehPreposicaoNome(lista[i])
        ) {
            total++;
        }
    }

    return total;
}


function limparTamanhoDoNomeJSX(texto) {
    if (!texto) {
        return "";
    }
    var limpo = String(texto).replace(/\.(?:jpe?g|png|tiff?|psd)$/i, "");
    limpo = limpo.replace(/[\(\[\{\s_\-\.]*(?:20\s*[xX×\*]\s*20|24\s*[xX×\*]\s*34|34\s*[xX×\*]\s*24|30\s*[xX×\*]\s*20|20\s*[xX×\*]\s*30|30\s*[xX×\*]\s*22)(?:\s*cm)?(?:\s*\(\d+\))?[\)\]\}\s_\-\.]*/gi, " ");
    return normalizar(limpo);
}


function dividirNomeDuasLinhas(nomeCompleto) {

    var bruto =
        limparTamanhoDoNomeJSX(nomeCompleto).split(" ");

    var palavras = [];

    for (
        var p = 0;
        p < bruto.length;
        p++
    ) {

        if (
            bruto[p] !== ""
        ) {
            palavras.push(
                bruto[p]
            );
        }
    }

    var n = palavras.length;

    if (n === 0) {
        return {
            linha1: "",
            linha2: ""
        };
    }

    if (n === 1) {
        return {
            linha1: palavras[0],
            linha2: ""
        };
    }

    if (n === 2) {
        return {
            linha1: palavras[0],
            linha2: palavras[1]
        };
    }


    // corte = quantidade de palavras na 1ª linha (1 .. n-1)
    var melhorCorte = -1;
    var melhorScore = null;

    // 1ª passada: a 1ª linha tem EXATAMENTE 2 nomes
    // (nunca só 1, nunca 3 ou mais — preposições não contam).
    // 2ª passada (só se a 1ª não achar corte): regra livre.
    for (
        var passada = 0;
        passada < 2 && melhorCorte < 1;
        passada++
    )
    for (
        var corte = 1;
        corte < n;
        corte++
            ) {

                if (
            passada === 0
            && contarNomes(palavras, 0, corte) !== 2
        ) {
            continue;
        }

        // Não terminar a 1ª linha com preposição
        if (
            ehPreposicaoNome(
                palavras[corte - 1]
            )
        ) {
            continue;
        }

        // Não deixar a 2ª linha só com preposição
        if (
            corte === n - 1
            && ehPreposicaoNome(
                palavras[corte]
            )
        ) {
            continue;
        }

        var l1 =
            juntarPalavras(
                palavras,
                0,
                corte
            );

        var l2 =
            juntarPalavras(
                palavras,
                corte,
                n
            );

        var len1 = l1.length;
        var len2 = l2.length;

        // Desequilíbrio absoluto
        var desequilibrio =
            Math.abs(len1 - len2);

        // Penaliza linha "muito grande" em relação à outra
        var maior =
            Math.max(len1, len2);

        var menor =
            Math.min(len1, len2);

        var razao =
            menor > 0
                ? (maior / menor)
                : 99;

        var penalidadeRazao =
            razao > 2.2
                ? (razao - 2.2) * 12
                : 0;

        var penalidadeAbsoluta =
            maior > 22
                ? (maior - 22) * 3
                : 0;

        // Prefere preposição no INÍCIO da 2ª linha
        // ("Maria Luzia" | "de Lima Oliveira")
        var bonusPreposicao = 0;

        if (
            ehPreposicaoNome(
                palavras[corte]
            )
        ) {
            bonusPreposicao = -6;
        }

        // Leve preferência pelo meio das palavras
        var meioIdeal =
            n / 2;

        var distanciaMeio =
            Math.abs(
                corte - meioIdeal
            ) * 1.5;

        var score =
            desequilibrio
            + penalidadeRazao
            + penalidadeAbsoluta
            + distanciaMeio
            + bonusPreposicao;

        if (
            melhorScore === null
            || score < melhorScore
        ) {

            melhorScore =
                score;

            melhorCorte =
                corte;
        }
    }


    // Fallback: meio matemático sem preposição no fim
    if (melhorCorte < 1) {

        melhorCorte =
            Math.ceil(n / 2);

        if (
            melhorCorte >= n
        ) {
            melhorCorte =
                n - 1;
        }

        while (
            melhorCorte > 1
            && ehPreposicaoNome(
                palavras[melhorCorte - 1]
            )
        ) {
            melhorCorte--;
        }
    }


    return {
        linha1: juntarPalavras(
            palavras,
            0,
            melhorCorte
        ),
        linha2: juntarPalavras(
            palavras,
            melhorCorte,
            n
        )
    };
}


// ==========================================================
// NOME
// ==========================================================

// Escala o texto em % mantendo o ponto de âncora (fim da linha
// no texto alinhado à direita), igual à prévia da revisão.
function escalarTextoNaAncora(
    camada,
    escala
) {

    if (
        !camada
        || escala === undefined
        || escala === null
        || Math.abs(Number(escala) - 100) < 0.5
    ) {
        return;
    }

    var ancora =
        camada.textItem.position;

    camada.resize(
        Number(escala),
        Number(escala),
        AnchorPosition.MIDDLECENTER
    );

    camada.textItem.position =
        ancora;
}


function alterarNomePrincipal(
    container,
    nomeCompleto,
    layout
) {

    var textos =
        buscarTextos(
            container,
            []
        );

    if (
        textos.length < 2
    ) {

        erro(
            "O GRUPO NOME PRECISA TER PELO MENOS 2 TEXTOS.\n\n"
            + "TEXTOS ENCONTRADOS: "
            + textos.length
        );
    }


    var primeiraLinha;
    var segundaLinha;

    if (
        layout
        && layout.linha1 !== undefined
        && layout.linha1 !== null
    ) {

        primeiraLinha =
            layout.linha1;

        segundaLinha =
            layout.linha2 || "";

    } else {

        var dividido =
            dividirNomeDuasLinhas(
                nomeCompleto
            );

        primeiraLinha =
            dividido.linha1;

        segundaLinha =
            dividido.linha2;
    }


    textos[0].textItem.contents =
        primeiraLinha;

    textos[0].visible =
        true;

    textos[0].opacity =
        100;


    textos[1].textItem.contents =
        segundaLinha;

    textos[1].visible =
        true;

    textos[1].opacity =
        100;


    if (layout) {

        escalarTextoNaAncora(
            textos[0],
            layout.escala1
        );

        escalarTextoNaAncora(
            textos[1],
            layout.escala2
        );
    }
}


// ==========================================================
// ABRIR SMART OBJECT
// ==========================================================

function abrirSmartObject(
    camada
) {

    if (
        !camada
    ) {

        erro(
            "SMART OBJECT NÃO ENCONTRADO."
        );
    }

    if (
        camada.kind !==
        LayerKind.SMARTOBJECT
    ) {

        erro(
            "A CAMADA NÃO É SMART OBJECT:\n\n"
            + camada.name
        );
    }


    var documentoPai =
        camada;

    while (
        documentoPai
        &&
        documentoPai.typename !==
            "Document"
    ) {

        documentoPai =
            documentoPai.parent;
    }


    if (
        !documentoPai
    ) {

        erro(
            "NÃO FOI POSSÍVEL IDENTIFICAR O DOCUMENTO PAI."
        );
    }


    app.activeDocument =
        documentoPai;

    documentoPai.activeLayer =
        camada;


    etapa(
        "ABRIR SMART OBJECT (placedLayerEditContents): "
        + camada.name
    );

    executeAction(
        stringIDToTypeID(
            "placedLayerEditContents"
        ),
        undefined,
        DialogModes.NO
    );


    return app.activeDocument;
}


// ==========================================================
// FECHAR SMART OBJECT
// ==========================================================

function fecharSmartObjectSalvando(
    documentoSmart,
    documentoPai
) {

    if (
        !documentoSmart
    ) {

        return;
    }

    try {

        app.activeDocument =
            documentoSmart;

        etapa(
            "SALVAR SMART OBJECT (.psb temporário do Photoshop)"
        );

        // Salva pelo Action Manager, que aceita o documento temporário
        // do Smart Object mesmo quando o método DOM falha.
        var descSalvar =
            new ActionDescriptor();

        descSalvar.putBoolean(
            stringIDToTypeID("asCopy"),
            false
        );

        executeAction(
            stringIDToTypeID("save"),
            descSalvar,
            DialogModes.NO
        );

        documentoSmart.close(
            SaveOptions.DONOTSAVECHANGES
        );

    }

    catch (e) {

        erro(
            "NÃO FOI POSSÍVEL SALVAR O SMART OBJECT:\n\n"
            + e.message
        );
    }


    if (
        documentoPai
    ) {

        app.activeDocument =
            documentoPai;
    }
}


// ==========================================================
// DESATIVAR BRASÕES
// ==========================================================

function desativarBrasoes(
    documento
) {

    var grupos = [

        "ABC",
        "DEF",
        "GHJ",
        "MNOP",
        "RST"
    ];


    for (
        var g = 0;
        g < grupos.length;
        g++
    ) {

        var encontrados =
            buscarRecursivo(
                documento,
                grupos[g],
                []
            );


        for (
            var i = 0;
            i < encontrados.length;
            i++
        ) {

            var grupo =
                encontrados[i];

            if (
                grupo.typename !==
                "LayerSet"
            ) {

                continue;
            }


            for (
                var j = 0;
                j < grupo.layers.length;
                j++
            ) {

                grupo.layers[j].visible =
                    false;
            }
        }
    }
}


// ==========================================================
// ATIVAR BRASÃO
// ==========================================================

function ativarBrasao(
    documento,
    nomeCurso
) {

    if (
        !nomeCurso
    ) {

        return;
    }


    desativarBrasoes(
        documento
    );


    var encontrados =
        buscarRecursivo(
            documento,
            nomeCurso,
            []
        );


    if (
        encontrados.length === 0
    ) {

        erro(
            "BRASÃO NÃO ENCONTRADO:\n\n"
            + nomeCurso
        );
    }


    if (
        encontrados.length > 1
    ) {

        erro(
            "BRASÃO DUPLICADO:\n\n"
            + nomeCurso
        );
    }


    encontrados[0].visible =
        true;
}


// ==========================================================
// COR DE BACKGROUND
// ==========================================================

function rgbParaCor(
    r,
    g,
    b
) {

    var cor =
        new SolidColor();

    cor.rgb.red =
        r;

    cor.rgb.green =
        g;

    cor.rgb.blue =
        b;

    return cor;
}


function obterPresetCor(
    nomeBrasao
) {

    if (
        !nomeBrasao
    ) {
        return null;
    }

    var valor =
        normalizar(
            nomeBrasao
        ).toUpperCase();

    // Presets locais (fallback se o Python nao enviar RGB)
    var tabela = {

        "ARQ": [58, 32, 75],
        "DIREITO": [105, 0, 0],
        "ADM": [180, 130, 20],
        "ADS": [30, 60, 120],
        "ENF": [0, 110, 70],
        "PSI": [90, 40, 120],
        "MEDICINA": [0, 100, 40],
        "PED": [180, 100, 20],
        "ODONTO": [0, 120, 130],
        "NUTRI": [40, 130, 40],
        "FISIO": [0, 120, 140],
        "FARM": [0, 100, 80],
        "BIOMED": [0, 90, 90],
        "ENG CIV": [180, 90, 20],
        "ENG MEC": [70, 70, 70],
        "ENG ELET": [200, 160, 0],
        "ENG PROD": [40, 80, 120],
        "ENG AMB": [40, 120, 40],
        "CIEN COMP": [20, 50, 120],
        "TI": [20, 50, 110],
        "SI": [30, 70, 120],
        "RH": [140, 80, 20],
        "PUBLI": [160, 30, 90],
        "ED FIS": [0, 100, 60],
        "MED VET": [0, 90, 50],
        "SS": [140, 60, 20],
        "CCONT": [80, 80, 80]
    };

    if (
        tabela[valor]
    ) {

        var rgb =
            tabela[valor];

        return rgbParaCor(
            rgb[0],
            rgb[1],
            rgb[2]
        );
    }

    return null;
}


function listarCamadasPintaveis(
    container,
    lista
) {

    if (
        !lista
    ) {
        lista = [];
    }

    for (
        var i = 0;
        i < container.layers.length;
        i++
    ) {

        var camada =
            container.layers[i];

        if (
            camada.typename ===
            "LayerSet"
        ) {

            listarCamadasPintaveis(
                camada,
                lista
            );

        }

        else if (
            camada.typename ===
            "ArtLayer"
        ) {

            if (
                camada.kind !==
                LayerKind.TEXT
            ) {

                lista.push(
                    camada
                );
            }
        }
    }

    return lista;
}


function aplicarCorShapeOuFill(
    documento,
    camada,
    cor,
    modo
) {

    app.activeDocument =
        documento;

    documento.activeLayer =
        camada;

    app.foregroundColor =
        cor;


    // 1) Camada de preenchimento / shape (content layer)
    if (
        modo !== "pixels"
    ) try {

        var desc =
            new ActionDescriptor();

        var ref =
            new ActionReference();

        ref.putEnumerated(
            stringIDToTypeID(
                "contentLayer"
            ),
            charIDToTypeID(
                "Ordn"
            ),
            charIDToTypeID(
                "Trgt"
            )
        );

        desc.putReference(
            charIDToTypeID(
                "null"
            ),
            ref
        );

        var fillDesc =
            new ActionDescriptor();

        var colorDesc =
            new ActionDescriptor();

        colorDesc.putDouble(
            charIDToTypeID(
                "Rd  "
            ),
            cor.rgb.red
        );

        colorDesc.putDouble(
            charIDToTypeID(
                "Grn "
            ),
            cor.rgb.green
        );

        colorDesc.putDouble(
            charIDToTypeID(
                "Bl  "
            ),
            cor.rgb.blue
        );

        fillDesc.putObject(
            charIDToTypeID(
                "Clr "
            ),
            charIDToTypeID(
                "RGBC"
            ),
            colorDesc
        );

        desc.putObject(
            charIDToTypeID(
                "T   "
            ),
            stringIDToTypeID(
                "solidColorLayer"
            ),
            fillDesc
        );

        executeAction(
            charIDToTypeID(
                "setd"
            ),
            desc,
            DialogModes.NO
        );

        return true;

    }

    catch (e1) {}


    if (
        modo === "fill"
    ) {
        return false;
    }


    // 2) Camada normal: preenche pixels (respeitando transparencia)
    try {

        var trava =
            camada.transparentPixelsLocked;

        camada.transparentPixelsLocked =
            true;

        documento.selection.selectAll();

        documento.selection.fill(
            cor,
            ColorBlendMode.NORMAL,
            100,
            true
        );

        documento.selection.deselect();

        camada.transparentPixelsLocked =
            trava;

        return true;

    }

    catch (e2) {}


    return false;
}


function ativarCamadaCorPorNome(
    background,
    nomeBrasao
) {

    if (
        !nomeBrasao
    ) {
        return false;
    }

    var encontrados =
        buscarRecursivo(
            background,
            nomeBrasao,
            []
        );

    if (
        encontrados.length === 0
    ) {
        return false;
    }


    // Desliga irmãos do mesmo grupo pai (swatches de cor)
    var alvo =
        encontrados[0];

    var pai =
        alvo.parent;

    if (
        pai
        && pai.typename ===
            "LayerSet"
    ) {

        for (
            var i = 0;
            i < pai.layers.length;
            i++
        ) {

            pai.layers[i].visible =
                false;
        }
    }

    else if (
        background.typename ===
        "LayerSet"
    ) {

        for (
            var j = 0;
            j < background.layers.length;
            j++
        ) {

            background.layers[j].visible =
                false;
        }
    }


    alvo.visible =
        true;

    return true;
}


// ==========================================================
// BACKGROUND
// ==========================================================

function alterarCorBackground(
    documento,
    caminhoBackground,
    nomeBrasao,
    rgbArray,
    corAbaixoDaTextura
) {

    if (
        !caminhoBackground
        || !caminhoBackground.length
    ) {
        return;
    }


    var background;

    try {

        background =
        encontrarCaminho(
            documento,
            caminhoBackground
        );

    }

    catch (e) {

        // Sem grupo BACKGROUND neste PSD — ignora
        return;
    }


    if (
        background.typename !==
        "LayerSet"
    ) {

        erro(
            "BACKGROUND DEVERIA SER UM GRUPO:\n\n"
            + background.name
        );
    }


    // ----------------------------------------------------------
    // Modo A: camada/grupo com o nome do curso dentro do BACKGROUND
    // (mesma ideia do brasão — liga a cor certa)
    // ----------------------------------------------------------

    if (
        ativarCamadaCorPorNome(
            background,
            nomeBrasao
        )
    ) {

        return;
    }


    // ----------------------------------------------------------
    // Modo B: pintar shape / fill com RGB do curso
    // ----------------------------------------------------------

    var cor =
        null;

    if (
        rgbArray
        && rgbArray.length === 3
    ) {

        cor =
            rgbParaCor(
                rgbArray[0],
                rgbArray[1],
                rgbArray[2]
            );

    }

    else {

        cor =
        obterPresetCor(
            nomeBrasao
        );
    }


    if (
        !cor
    ) {

        // Sem cor mapeada — nao quebra o lote
        return;
    }


        app.foregroundColor =
            cor;


    var pintaveis =
        listarCamadasPintaveis(
            background,
            []
        );


    // 1ª tentativa: só camadas de preenchimento de cor (fill/shape).
    // Assim a textura (camada de pixels) não é pintada por cima.
    var fillsPintados = [];

    for (
        var k = 0;
        k < pintaveis.length;
        k++
    ) {

        if (
            !pintaveis[k].visible
        ) {
            continue;
        }

        if (
            aplicarCorShapeOuFill(
                documento,
                pintaveis[k],
                cor,
                "fill"
            )
        ) {

            fillsPintados.push(
                pintaveis[k]
            );
        }
    }


    if (
        fillsPintados.length > 0
    ) {

        if (
            corAbaixoDaTextura
        ) {

            for (
                var f = 0;
                f < fillsPintados.length;
                f++
            ) {

                moverFillAbaixoDaTextura(
                    background,
                    fillsPintados[f]
                );
            }
        }

        return;
    }


    // 2ª tentativa (PSD sem camada de preenchimento):
    // pinta os pixels das camadas visíveis, como antes.
    var pintou =
        false;

    for (
        var m = 0;
        m < pintaveis.length;
        m++
    ) {

        if (
            !pintaveis[m].visible
        ) {
            continue;
        }

        if (
            aplicarCorShapeOuFill(
                documento,
                pintaveis[m],
                cor,
                "pixels"
            )
        ) {

            pintou =
                true;
        }
    }


    // Se nada visivel foi pintado, tenta a primeira camada pintavel
    if (
        !pintou
        && pintaveis.length > 0
    ) {

        pintaveis[0].visible =
            true;

        aplicarCorShapeOuFill(
            documento,
            pintaveis[0],
            cor
        );
    }
}


// ==========================================================
// PREENCHIMENTO DE COR ABAIXO DA TEXTURA (20X20)
//
// A textura é a camada do BACKGROUND com "TEXTURA" no nome.
// Se não houver, é a camada logo abaixo do preenchimento
// (desce o preenchimento uma posição).
// Só move se o preenchimento ainda estiver acima da textura.
// ==========================================================

function indiceNoPai(
    camada
) {

    var irmaos =
        camada.parent.layers;

    for (
        var i = 0;
        i < irmaos.length;
        i++
    ) {

        if (
            irmaos[i] === camada
        ) {
            return i;
        }
    }

    return -1;
}


function acharTextura(
    container,
    ignorar
) {

    for (
        var i = 0;
        i < container.layers.length;
        i++
    ) {

        var camada =
            container.layers[i];

        if (
            camada !== ignorar
            && normalizar(camada.name).toUpperCase().search("TEXTURA") >= 0
        ) {
            return camada;
        }

        if (
            camada.typename === "LayerSet"
        ) {

            var dentro =
                acharTextura(
                    camada,
                    ignorar
                );

            if (
                dentro
            ) {
                return dentro;
            }
        }
    }

    return null;
}


function moverFillAbaixoDaTextura(
    background,
    fill
) {

    try {

        var textura =
            acharTextura(
                background,
                fill
            );

        if (
            !textura
        ) {

            var indice =
                indiceNoPai(
                    fill
                );

            var irmaos =
                fill.parent.layers;

            if (
                indice < 0
                || indice + 1 >= irmaos.length
            ) {
                return;
            }

            textura =
                irmaos[indice + 1];
        }


        // Já está abaixo da textura no mesmo grupo: nada a fazer
        if (
            textura.parent === fill.parent
            && indiceNoPai(fill) > indiceNoPai(textura)
        ) {
            return;
        }

        etapa(
            "MOVER PREENCHIMENTO DE COR PARA BAIXO DA TEXTURA: "
            + textura.name
        );

        fill.move(
            textura,
            ElementPlacement.PLACEAFTER
        );

    } catch (e) {

        erro(
            "NÃO FOI POSSÍVEL COLOCAR O PREENCHIMENTO DE COR ABAIXO DA TEXTURA:\n\n"
            + e.message
        );
    }
}


// ==========================================================
// TEXTO SIMPLES
// ==========================================================

function alterarTexto(
    camada,
    texto,
    descricao
) {

    if (
        camada.kind !==
        LayerKind.TEXT
    ) {

        erro(
            descricao
            + " NÃO É TEXTO:\n\n"
            + camada.name
        );
    }


    camada.textItem.contents =
        limparTamanhoDoNomeJSX(texto);
}


// ==========================================================
// PROCESSAR PRIME BOOK / MOSCOU
// ==========================================================

function processarPrimeBook(
    documento,
    job
) {

    var config =
        job.config;


    // FOTO (igual ao 20x20: só troca o conteúdo, sem redimensionar)

    substituirFoto(
        documento,
        job.fotos.foto,
        config.foto,
        false,
        AnchorPosition.MIDDLECENTER
    );


    // NOME

    var grupoNome =
        encontrarCaminho(
            documento,
            config.nome_principal
        );

    alterarNomePrincipal(
        grupoNome,
        job.nome,
        job.nome_layout
    );


    // BACKGROUND

    alterarCorBackground(
        documento,
        config.background,
        job.curso_simbolo,
        job.cor_background
    );


    // BRASÃO

    var frente =
        encontrarCaminho(
            documento,
            config.frente
        );

    var objeto =
        abrirSmartObject(
            frente
        );


    ativarBrasao(
        objeto,
        job.curso_simbolo
    );


    fecharSmartObjectSalvando(
        objeto,
        documento
    );


    app.activeDocument =
        documento;


    // LATERAL

    var lateral =
        encontrarCaminho(
            documento,
            config.lateral
        );

    var nomeLateral =
        buscarUnico(
            lateral,
            config.nome_lateral,
            "NOME DA LATERAL"
        );


    alterarTexto(
        nomeLateral,
        job.nome,
        "NOME DA LATERAL"
    );

    if (job.nome_layout) {
        escalarTextoNaAncora(
            nomeLateral,
            job.nome_layout.escala_lombada
        );
    }
}


// ==========================================================
// PROCESSAR 30X22 (FRENTE E VERSO EM PSDS SEPARADOS)
// ==========================================================

function acharTextoNaRaiz(
    documento,
    nome
) {

    for (
        var i = 0;
        i < documento.artLayers.length;
        i++
    ) {

        var camada =
            documento.artLayers[i];

        if (
            camada.name === nome
            && camada.kind === LayerKind.TEXT
        ) {
            return camada;
        }
    }

    erro(
        "TEXTO NÃO ENCONTRADO NA RAIZ DO PSD:\n\n"
        + nome
    );
}


function corDoCurso(
    job
) {

    var rgb =
        job.cor_background;

    if (
        rgb
        && rgb.length === 3
    ) {
        return rgbParaCor(
            rgb[0],
            rgb[1],
            rgb[2]
        );
    }

    return obterPresetCor(
        job.curso_simbolo
    );
}


function pintarFillPorCaminho(
    documento,
    caminho,
    job
) {

    var cor =
        corDoCurso(
            job
        );

    if (
        !cor
    ) {
        // Sem cor mapeada para o curso: mantém a do PSD
        return;
    }

    var camada =
        encontrarCaminho(
            documento,
            caminho
        );

    if (
        !aplicarCorShapeOuFill(
            documento,
            camada,
            cor,
            "fill"
        )
    ) {
        erro(
            "NÃO CONSEGUI PINTAR O PREENCHIMENTO:\n\n"
            + caminho.join(" > ")
        );
    }
}


function trocarBrasao(
    documento,
    caminho,
    simbolo
) {

    var camada =
        encontrarCaminho(
            documento,
            caminho
        );

    var objeto =
        abrirSmartObject(
            camada
        );

    ativarBrasao(
        objeto,
        simbolo
    );

    fecharSmartObjectSalvando(
        objeto,
        documento
    );

    app.activeDocument =
        documento;
}


function processarFrente30x22(
    documento,
    job
) {

    var config =
        job.config;

    // FOTO (mesmo tamanho do 30X20: só troca o conteúdo)

    substituirFoto(
        documento,
        job.fotos.foto,
        config.foto,
        false,
        AnchorPosition.MIDDLECENTER
    );

    // NOME (grupo NOME com 2 textos)

    alterarNomePrincipal(
        encontrarCaminho(
            documento,
            config.nome_principal
        ),
        job.nome,
        job.nome_layout
    );

    // NOME DA LOMBADA (texto NOME solto na raiz)

    var nomeLombada =
        acharTextoNaRaiz(
            documento,
            config.nome_lateral
        );

    alterarTexto(
        nomeLombada,
        job.nome,
        "NOME DA LOMBADA"
    );

    if (job.nome_layout) {
        escalarTextoNaAncora(
            nomeLombada,
            job.nome_layout.escala_lombada
        );
    }

    // COR (preenchimento solto na raiz)

    pintarFillPorCaminho(
        documento,
        config.cor,
        job
    );

    // BRASÃO

    trocarBrasao(
        documento,
        config.brasao,
        job.curso_simbolo
    );
}


function processarVerso30x22(
    documento,
    job
) {

    var config =
        job.config;

    alterarCorBackground(
        documento,
        config.background,
        job.curso_simbolo,
        job.cor_background
    );

    trocarBrasao(
        documento,
        config.brasao,
        job.curso_simbolo
    );
}


// ==========================================================
// PROCESSAR BERLIM 24X34
// ==========================================================

function processarBerlim24(
    documento,
    job
) {

    var config =
        job.config;


    // ======================================================
    // FOTO 24X34
    //
    // Só aumenta (proporcional, centralizado) até cobrir a
    // área original da camada FOTO. Nunca reduz.
    // ======================================================

    substituirFotoBerlim24(
        documento,
        job.fotos.foto34,
        config.foto34
    );


    // ======================================================
    // BRASÃO 24X34
    // ======================================================

    var frente =
        encontrarCaminho(
            documento,
            config.frente_24
        );


    var objeto =
        abrirSmartObject(
            frente
        );


    ativarBrasao(
        objeto,
        job.curso_simbolo
    );


    fecharSmartObjectSalvando(
        objeto,
        documento
    );


    app.activeDocument =
        documento;


    // ======================================================
    // NOME 24X34
    // ======================================================

    var camadaNome =
        encontrarCaminho(
            documento,
            config.nome_24
        );


    if (
        camadaNome.kind !==
        LayerKind.SMARTOBJECT
    ) {

        erro(
            "NOME 24X34 NÃO É SMART OBJECT:\n\n"
            + camadaNome.name
        );
    }


    var objetoNome =
        abrirSmartObject(
            camadaNome
        );


    alterarNomePrincipal(
        objetoNome,
        job.nome,
        job.nome_layout
    );


    fecharSmartObjectSalvando(
        objetoNome,
        documento
    );


    app.activeDocument =
        documento;
}


// ==========================================================
// PROCESSAR ESSENCIAL BOOK 20X20
// ==========================================================

function processarEssencial20(
    documento,
    job
) {

    var config =
        job.config;


    // ======================================================
    // 1. FOTO
    //
    // Ajusta a nova imagem aos limites originais da camada FOTO.
    // ======================================================

    substituirFoto(
        documento,
        job.fotos.foto,
        config.foto,
        false,
        AnchorPosition.MIDDLECENTER
    );


    // ======================================================
    // 2. BRASÃO
    //
    // BRASÃO
    // └── FRENTE (PRINCIPAL)
    //
    // O "copiar" não será utilizado.
    // ======================================================

    var frente =
        encontrarCaminho(
            documento,
            config.frente
        );


    var objetoFrente =
        abrirSmartObject(
            frente
        );


    ativarBrasao(
        objetoFrente,
        job.curso_simbolo
    );


    fecharSmartObjectSalvando(
        objetoFrente,
        documento
    );


    app.activeDocument =
        documento;


    // ======================================================
    // 3. NOME
    //
    // NOME
    // ├── NOME
    // └── nome do aluno
    // ======================================================

    var grupoNome =
        encontrarCaminho(
            documento,
            config.nome_principal
        );


    alterarNomePrincipal(
        grupoNome,
        job.nome,
        job.nome_layout
    );


    // ======================================================
    // 4. BACKGROUND
    //
    // No 20x20 o preenchimento de cor fica abaixo da textura.
    // ======================================================

    alterarCorBackground(
        documento,
        config.background,
        job.curso_simbolo,
        job.cor_background,
        true
    );


    // ======================================================
    // 5. LATERAL
    // ======================================================

    var lateral =
        encontrarCaminho(
            documento,
            config.lateral
        );


    var nomeLateral =
        buscarUnico(
            lateral,
            config.nome_lateral,
            "NOME DA LATERAL"
        );


    alterarTexto(
        nomeLateral,
        job.nome,
        "NOME DA LATERAL"
    );

    if (job.nome_layout) {
        escalarTextoNaAncora(
            nomeLateral,
            job.nome_layout.escala_lombada
        );
    }


    app.activeDocument =
        documento;
}


// ==========================================================
// EXPORTAR JPG
// ==========================================================

function exportarJPG(
    documento,
    caminho
) {

    var arquivo =
        new File(
            caminho
        );


    var pasta =
        arquivo.parent;


    if (
        !pasta.exists
    ) {

        pasta.create();
    }


    if (
        !pasta.exists
    ) {

        erro(
            "PASTA DE DESTINO NÃO EXISTE:\n\n"
            + pasta.fsName
        );
    }


    app.activeDocument =
        documento;

    var opcoes =
        new ExportOptionsSaveForWeb();

    opcoes.format =
        SaveDocumentType.JPEG;

    opcoes.includeProfile =
        true;

    opcoes.optimized =
        true;

    opcoes.quality =
        100;

    etapa(
        "EXPORTAR JPG: "
        + arquivo.fsName
    );

    documento.exportDocument(
        arquivo,
        ExportType.SAVEFORWEB,
        opcoes
    );
}


// ==========================================================
// EXPORTAR PSD
// ==========================================================

function exportarPSD(
    documento,
    caminho
) {

    var arquivo =
        new File(
            caminho
        );


    var pasta =
        arquivo.parent;


    if (
        !pasta.exists
    ) {

        pasta.create();
    }


    if (
        !pasta.exists
    ) {

        erro(
            "PASTA DE DESTINO NÃO EXISTE:\n\n"
            + pasta.fsName
        );
    }


    var opcoes =
        new PhotoshopSaveOptions();


    opcoes.embedColorProfile =
        true;

    opcoes.layers =
        true;

    opcoes.maximizeCompatibility =
        true;


    etapa(
        "EXPORTAR PSD: "
        + arquivo.fsName
    );

    documento.saveAs(
        arquivo,
        opcoes,
        true,
        Extension.LOWERCASE
    );
}


// ==========================================================
// EXECUTAR
// ==========================================================

function executar() {

    app.displayDialogs = DialogModes.NO;

    var job =
        JOB;


    if (
        !job
    ) {

        erro(
            "JOB NÃO CARREGADO."
        );
    }


    var template =
        new File(
            job.template
        );


    etapa(
        "ABRIR PSD: "
        + job.template
    );

    if (
        !template.exists
    ) {

        erro(
            "PSD NÃO ENCONTRADO:\n\n"
            + job.template
        );
    }


    var documento =
        app.open(
            template
        );


    try {

        // ==================================================
        // PROCESSAMENTO
        // ==================================================

        etapa(
            "PROCESSAR CAMADAS ("
            + job.modelo
            + ")"
        );

        if (
            job.modelo ===
            "PRIME BOOK 30X20"
            ||
            job.modelo ===
            "STUDIO DAY MOSCOU 30X20"
        ) {

            processarPrimeBook(
                documento,
                job
            );
        }


        else if (
            job.modelo ===
            "STUDIO DAY BERLIM 24X34"
        ) {

            processarBerlim24(
                documento,
                job
            );
        }


        else if (
            job.modelo ===
            "ESSENCIAL BOOK 20X20"
        ) {

            processarEssencial20(
                documento,
                job
            );
        }


        else if (
            job.modelo ===
            "PRIME BOOK 30X22 FRENTE"
        ) {

            processarFrente30x22(
                documento,
                job
            );
        }


        else if (
            job.modelo ===
            "PRIME BOOK 30X22 VERSO"
        ) {

            processarVerso30x22(
                documento,
                job
            );
        }


        else {

            erro(
                "MODELO NÃO CONFIGURADO:\n\n"
                + job.modelo
            );
        }


        // ==================================================
        // JPG
        // ==================================================

        if (
            job.saida_jpg
        ) {

            exportarJPG(
                documento,
                job.saida_jpg
            );
        }


        // ==================================================
        // PSD
        // ==================================================

        if (
            job.saida_psd
            &&
            job.exportar_psd
        ) {

            app.activeDocument =
                documento;

            exportarPSD(
                documento,
                job.saida_psd
            );
        }


        // ==================================================
        // FECHAR
        // ==================================================

        etapa(
            "FECHAR DOCUMENTO"
        );

        app.activeDocument =
            documento;

        documento.close(
            SaveOptions.DONOTSAVECHANGES
        );


        $.writeln(
            "CAPA CONCLUÍDA! "
            + job.nome
            + " - "
            + job.modelo
        );

    }

    catch (e) {

        try {

            app.activeDocument =
                documento;

            documento.close(
                SaveOptions.DONOTSAVECHANGES
            );

        }

        catch (_) {}


        $.writeln(
            "CAPA NÃO GERADA. "
            + job.nome
            + " - "
            + job.modelo
            + " - "
            + e.message
        );


        throw e;
    }
}


// ==========================================================
// PONTO DE ENTRADA
//
// Não deixa o erro "escapar" para um alerta do Photoshop:
// devolve "OK" ou "ERRO|etapa|mensagem" ao Python
// (retorno do DoJavaScriptFile + arquivo de status).
// ==========================================================

var RESULTADO_AUTOMACAO =
    "";

var REGUA_ORIGINAL =
    app.preferences.rulerUnits;

var TIPO_ORIGINAL =
    app.preferences.typeUnits;

try {

    // translate() usa a unidade da régua: sem isso, valores
    // em px viram cm/mm e a foto é jogada e aumentada demais.
    app.preferences.rulerUnits =
        Units.PIXELS;

executar();

    RESULTADO_AUTOMACAO =
        "OK";

}

catch (e) {

    RESULTADO_AUTOMACAO =
        "ERRO|"
        + ETAPA_ATUAL
        + "|"
        + (e && e.message ? e.message : String(e))
        + (e && e.line ? " (linha JSX " + e.line + ")" : "");
}

try {
    app.preferences.rulerUnits = REGUA_ORIGINAL;
    app.preferences.typeUnits = TIPO_ORIGINAL;
} catch (eRegua) {}

gravarStatus(
    RESULTADO_AUTOMACAO
);

RESULTADO_AUTOMACAO;
'''


# ============================================================
# EXECUTAR PHOTOSHOP
# ============================================================

CACHE_PSD_DIR = Path(tempfile.gettempdir()) / "capaz_psd_cache"
_PHOTOSHOP_APP = None


def obter_photoshop_app():
    """Reutiliza a instância COM do Photoshop evitando overhead de Dispatch repetido."""
    global _PHOTOSHOP_APP
    if _PHOTOSHOP_APP is not None:
        try:
            _ = _PHOTOSHOP_APP.Visible
            return _PHOTOSHOP_APP
        except Exception:
            _PHOTOSHOP_APP = None
    _PHOTOSHOP_APP = win32com.client.Dispatch("Photoshop.Application")
    _PHOTOSHOP_APP.Visible = True
    return _PHOTOSHOP_APP


def obter_template_psd_cache(template_original):
    """Mantém cópia local persistente do PSD validada por (mtime_ns, size) para eliminar cópias repetidas de centenas de MB."""
    template_original = Path(template_original)
    st = template_original.stat()
    chave_hash = hashlib.sha1(os.path.normcase(str(template_original.resolve())).encode("utf-8")).hexdigest()[:12]
    CACHE_PSD_DIR.mkdir(parents=True, exist_ok=True)
    meta_file = CACHE_PSD_DIR / f"template_{chave_hash}.json"
    cached_psd = CACHE_PSD_DIR / f"template_{chave_hash}{template_original.suffix.lower()}"
    assinatura = {"mtime_ns": st.st_mtime_ns, "size": st.st_size, "origem": str(template_original)}
    if cached_psd.is_file() and meta_file.is_file():
        try:
            meta = json.loads(meta_file.read_text(encoding="utf-8"))
            if meta.get("mtime_ns") == st.st_mtime_ns and meta.get("size") == st.st_size and cached_psd.stat().st_size == st.st_size:
                print("  [CACHE PSD HIT] Usando template local em cache:", cached_psd.name)
                return cached_psd
        except Exception:
            pass
    print("  [CACHE PSD] Atualizando cache local do PSD:", template_original.name)
    copiar_com_retry(template_original, cached_psd, "PSD TEMPLATE (cache persistente local)")
    meta_file.write_text(json.dumps(assinatura, ensure_ascii=False), encoding="utf-8")
    return cached_psd


def executar_photoshop(job):

    print()
    print(
        "ABRINDO PHOTOSHOP:",
        job["modelo"]
    )

    print(
        "ALUNO:",
        job["nome"]
    )

    print(
        "TEMPLATE:",
        job["template"]
    )

    template_original = Path(
        job["template"]
    )

    if not template_original.exists():

        raise Exception(
            "PSD NÃO ENCONTRADO:\n\n"
            + str(template_original)
        )

    # ========================================================
    # SAÍDAS FINAIS (REDE) E TESTE DE ESCRITA ANTES DO PHOTOSHOP
    # ========================================================

    # Lista de (chave no job, caminho final no destino).
    saidas_finais = []

    if job.get("saida_jpg"):
        saidas_finais.append(
            ("saida_jpg", Path(job["saida_jpg"]))
        )

    if job.get("exportar_psd") and job.get("saida_psd"):
        saidas_finais.append(
            ("saida_psd", Path(job["saida_psd"]))
        )

    if not saidas_finais:

        raise Exception(
            "NENHUMA SAÍDA CONFIGURADA (export_jpg e export_psd desligados)."
        )

    for pasta_destino in {
        destino.parent
        for _, destino in saidas_finais
    }:
        testar_escrita_pasta(
            pasta_destino
        )

    photoshop = obter_photoshop_app()

    # ========================================================
    # PASTA TEMPORÁRIA LOCAL
    #
    # Tudo que o Photoshop lê ou grava fica aqui (caminho
    # local, sem UNC): PSD, fotos, JSX, JPG/PSD de saída e
    # arquivo de status. O Python copia a saída para a rede.
    # ========================================================

    pasta_temp = Path(
        tempfile.mkdtemp(
            prefix="automacao_photoshop_"
        )
    )

    template_local = (
        pasta_temp /
        ("template" + template_original.suffix.lower())
    )

    arquivo_jsx = (
        pasta_temp /
        "_automacao_temp.jsx"
    )

    arquivo_status = (
        pasta_temp /
        "_status.txt"
    )

    pasta_fotos_local = (
        pasta_temp /
        "fotos"
    )

    pasta_saida_local = (
        pasta_temp /
        "saida"
    )

    try:

        # ----------------------------------------------------
        # COPIAR PSD
        # ----------------------------------------------------

        template_cached = obter_template_psd_cache(template_original)

        job_exec = dict(
            job
        )

        job_exec["template"] = str(
            template_cached
        )

        # ----------------------------------------------------
        # COPIAR FOTOS (nomes ASCII, sem acentos/UNC)
        # ----------------------------------------------------

        pasta_fotos_local.mkdir(
            parents=True,
            exist_ok=True
        )

        fotos_locais = {}

        for chave_foto, caminho_foto in job.get("fotos", {}).items():

            foto_origem = Path(
                caminho_foto
            )

            if not foto_origem.is_file():

                raise Exception(
                    "FOTO NÃO ENCONTRADA OU SEM ACESSO DE LEITURA:\n\n"
                    + str(foto_origem)
                )

            foto_local = (
                pasta_fotos_local /
                (chave_foto + foto_origem.suffix.lower())
            )

            copiar_com_retry(
                foto_origem,
                foto_local,
                "FOTO (cópia para temporário local)"
            )

            fotos_locais[chave_foto] = str(
                foto_local
            )

        job_exec["fotos"] = fotos_locais

        # ----------------------------------------------------
        # SAÍDAS LOCAIS
        # ----------------------------------------------------

        pasta_saida_local.mkdir(
            parents=True,
            exist_ok=True
        )

        saidas_locais = []

        for chave_saida, destino in saidas_finais:

            saida_local = (
                pasta_saida_local /
                ("saida" + destino.suffix.lower())
            )

            job_exec[chave_saida] = str(
                saida_local
            )

            saidas_locais.append(
                (saida_local, destino)
            )

        if not job.get("saida_jpg"):
            job_exec["saida_jpg"] = ""

        job_exec["arquivo_status"] = str(
            arquivo_status
        )

        # ----------------------------------------------------
        # ESCREVER E RODAR JSX
        # ----------------------------------------------------

        job_json = json.dumps(
            job_exec,
            ensure_ascii=False
        )

        codigo = (
            "#target photoshop\n"
            "var JOB = "
            + job_json
            + ";\n"
            + JSX
        )

        arquivo_jsx.write_text(
            codigo,
            encoding="utf-8"
        )

        retorno_jsx = None

        try:

            retorno_jsx = photoshop.DoJavaScriptFile(
            str(arquivo_jsx),
            [],
            1
        )

        except Exception as e:

            status = _ler_status_jsx(
                arquivo_status
            )

            raise Exception(
                "ERRO NA CHAMADA AO PHOTOSHOP (DoJavaScriptFile / COM):\n\n"
                + str(e)
                + (("\n\nSTATUS DO JSX: " + status) if status else "")
                + "\n\nSe aparecer 'acesso negado' aqui, verifique se o "
                "Photoshop e o Python rodam com o MESMO nível de permissão "
                "(os dois como administrador, ou nenhum)."
            )

        status = (
            _ler_status_jsx(arquivo_status)
            or (str(retorno_jsx) if retorno_jsx else "")
        )

        if status.startswith("ERRO|"):

            partes = status.split("|", 2)

            etapa_erro = partes[1] if len(partes) > 1 else "?"
            mensagem_erro = partes[2] if len(partes) > 2 else status

            raise Exception(
                "ERRO NO PHOTOSHOP\n\n"
                + "ETAPA: " + etapa_erro
                + "\n\nMENSAGEM: " + mensagem_erro
            )

        # ----------------------------------------------------
        # CONFERIR SAÍDAS LOCAIS
        # ----------------------------------------------------

        for saida_local, destino in saidas_locais:

            if (
                not saida_local.is_file()
                or saida_local.stat().st_size == 0
            ):
                raise Exception(
                    "O PHOTOSHOP NÃO GEROU O ARQUIVO LOCAL (ou está vazio):\n\n"
                    + str(saida_local)
                    + "\n\nDESTINO PREVISTO: " + str(destino)
                    + (("\n\nSTATUS DO JSX: " + status) if status else
                       "\n\nO JSX NÃO DEVOLVEU STATUS (veja se o Photoshop "
                       "mostrou algum alerta).")
                )

        # ----------------------------------------------------
        # COPIAR PARA O DESTINO FINAL (REDE) COM RETRY
        # ----------------------------------------------------

        for saida_local, destino in saidas_locais:

            copiar_com_retry(
                saida_local,
                destino,
                "ARQUIVO DE SAÍDA NO DESTINO"
            )

        for _, destino in saidas_locais:

            if (
                not destino.is_file()
                or destino.stat().st_size == 0
            ):
                raise Exception(
                    "ARQUIVO DE SAÍDA NÃO FOI CRIADO OU ESTÁ VAZIO:\n\n"
                    + str(destino)
                )

    finally:

        _fechar_documentos_temp(
            photoshop,
            pasta_temp
        )

        _apagar_pasta_temp(
            pasta_temp
        )


# ============================================================
# AUXILIARES DE ARQUIVO (TEMP LOCAL / REDE)
# ============================================================

DICA_ACESSO_NEGADO = (
    "Possíveis causas:\n"
    "- arquivo aberto em outro programa (Photoshop, visualizador de imagens);\n"
    "- arquivo marcado como somente leitura;\n"
    "- pasta/compartilhamento de rede sem permissão de escrita "
    "(ou disco do servidor em modo somente leitura)."
)


def testar_escrita_pasta(pasta):
    """Cria e apaga um arquivo de teste; erro claro se não puder gravar."""

    pasta = Path(pasta)

    try:
        pasta.mkdir(
            parents=True,
            exist_ok=True
        )
    except Exception as e:
        raise Exception(
            "PASTA DE DESTINO INACESSÍVEL:\n\n"
            + str(pasta)
            + "\n\n"
            + str(e)
        )

    teste = pasta / (
        "_teste_escrita_" + uuid.uuid4().hex + ".tmp"
    )

    try:
        teste.write_bytes(b"")
    except PermissionError as e:
        raise Exception(
            "ACESSO NEGADO: SEM PERMISSÃO DE ESCRITA NA PASTA DE DESTINO:\n\n"
            + str(pasta)
            + "\n\n"
            + str(e)
            + "\n\n"
            + DICA_ACESSO_NEGADO
        )
    except OSError as e:
        raise Exception(
            "ERRO AO GRAVAR NA PASTA DE DESTINO:\n\n"
            + str(pasta)
            + "\n\n"
            + str(e)
        )
    finally:
        try:
            if teste.exists():
                teste.unlink()
        except Exception:
            pass


def copiar_com_retry(origem, destino, descricao, tentativas=5, pausa=2.0):
    """
    Copia origem -> destino via arquivo '.parcial' + os.replace,
    tentando várias vezes (arquivo travado / rede instável).
    """

    origem = Path(origem)
    destino = Path(destino)

    parcial = destino.with_name(
        destino.name + ".parcial"
    )

    ultimo_erro = None

    for tentativa in range(1, tentativas + 1):

        try:

            destino.parent.mkdir(
                parents=True,
                exist_ok=True
            )

            shutil.copy2(
                origem,
                parcial
            )

            if parcial.stat().st_size != origem.stat().st_size:
                raise OSError(
                    "cópia incompleta (tamanho diferente)"
                )

            os.replace(
                parcial,
                destino
            )

            return

        except Exception as e:

            ultimo_erro = e

            try:
                if parcial.exists():
                    parcial.unlink()
            except Exception:
                pass

            if tentativa < tentativas:
                print(
                    "  AVISO: falha ao copiar",
                    descricao,
                    "(tentativa", tentativa, "de", str(tentativas) + "):",
                    e
                )
                time.sleep(pausa)

    arquivo_erro = getattr(ultimo_erro, "filename", None)

    if isinstance(ultimo_erro, PermissionError):
        raise Exception(
            "ACESSO NEGADO AO COPIAR " + descricao + ":\n\n"
            + "ORIGEM: " + str(origem)
            + "\nDESTINO: " + str(destino)
            + (("\nARQUIVO COM ACESSO NEGADO: " + str(arquivo_erro))
               if arquivo_erro else "")
            + "\n\n" + DICA_ACESSO_NEGADO
        )

    raise Exception(
        "FALHA AO COPIAR " + descricao + " APÓS "
        + str(tentativas) + " TENTATIVAS:\n\n"
        + "ORIGEM: " + str(origem)
        + "\nDESTINO: " + str(destino)
        + "\n\n" + str(ultimo_erro)
    )


def _ler_status_jsx(arquivo_status):

    try:
        if Path(arquivo_status).is_file():
            return Path(arquivo_status).read_text(
                encoding="utf-8-sig",
                errors="replace"
            ).strip()
    except Exception:
        pass

    return ""


def _fechar_documentos_temp(photoshop, pasta_temp):
    """Fecha (sem salvar) documentos do Photoshop abertos a partir do temp."""

    try:
        prefixo = os.path.normcase(
            str(pasta_temp)
        )

        documentos = photoshop.Documents

        for i in range(documentos.Count, 0, -1):

            try:
                doc = documentos.Item(i)
                caminho = os.path.normcase(
                    str(doc.FullName)
                )
            except Exception:
                continue

            prefixo_cache = os.path.normcase(str(CACHE_PSD_DIR))
            if caminho.startswith(prefixo) or caminho.startswith(prefixo_cache):
                try:
                    # 2 = psDoNotSaveChanges
                    doc.Close(2)
                except Exception:
                    pass

    except Exception:
        pass


def _apagar_pasta_temp(pasta_temp, tentativas=5, pausa=1.0):

    for _ in range(tentativas):

        shutil.rmtree(
            pasta_temp,
            ignore_errors=True
        )

        if not Path(pasta_temp).exists():
            return

        time.sleep(pausa)

    print(
        "  AVISO: não foi possível apagar a pasta temporária:",
        pasta_temp
    )


# ============================================================
# RELATÓRIO
# ============================================================
# RELATORIO / FILTROS / PREFLIGHT / CLI
# ============================================================

def salvar_relatorio(resultados, arquivo=None):

    if arquivo is None:
        arquivo = RELATORIO_CSV
    else:
        arquivo = Path(arquivo)

    linhas = [
        "STATUS;LINHA;NOME;CURSO;PLANO;MENSAGEM"
    ]

    for resultado in resultados:

        valores = [
            resultado.get("status", ""),
            resultado.get("linha", ""),
            resultado.get("nome", ""),
            resultado.get("curso", ""),
            resultado.get("plano", ""),
            resultado.get("mensagem", ""),
        ]

        linhas.append(
            ";".join(
                str(x).replace(";", ",").replace("\n", " ")
                for x in valores
            )
        )

    arquivo.write_text(
        "\n".join(linhas),
        encoding="utf-8-sig"
    )

    print()
    print("RELATORIO:", arquivo)
    return arquivo


def carregar_erros_csv(arquivo=None):
    """{linha: nome} com STATUS=ERRO no CSV anterior, para --so-erros / retomar."""

    arquivo = Path(arquivo) if arquivo else RELATORIO_CSV

    if not arquivo.is_file():
        return {}

    erros = {}

    with arquivo.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            status = (row.get("STATUS") or "").strip().upper()
            if status != "ERRO":
                continue
            try:
                erros[int(row.get("LINHA") or 0)] = (row.get("NOME") or "").strip()
            except ValueError:
                continue

    return erros


def parse_faixa_linhas(texto):
    """
    Aceita: '10', '10-20', '10,12,15-18'
    Retorna set de ints.
    """

    if not texto:
        return None

    linhas = set()

    for parte in str(texto).split(","):
        parte = parte.strip()
        if not parte:
            continue
        if "-" in parte:
            a, b = parte.split("-", 1)
            inicio = int(a.strip())
            fim = int(b.strip())
            if fim < inicio:
                inicio, fim = fim, inicio
            linhas.update(range(inicio, fim + 1))
        else:
            linhas.add(int(parte))

    return linhas


def filtrar_registros(
    registros,
    aluno=None,
    plano=None,
    linhas=None,
    so_erros=False,
    csv_erros=None,
    lote=None,
    status=None,
):
    filtrados = list(registros)

    if lote:
        # "12" ou "Lote 12" -> só o lote 12 (não pega 1, 2, 120...); vários com vírgula
        def numero(texto):
            m = re.search(r"\d+", texto or "")
            return m.group(0).lstrip("0") if m else chave(texto)
        alvos = {numero(p) for p in lote.split(",") if p.strip()}
        filtrados = [r for r in filtrados if r.get("lote") and numero(r["lote"]) in alvos]

    if status:
        alvos = [chave(p) for p in status.split(",") if p.strip()]
        filtrados = [
            r for r in filtrados
            if any(a in chave(r.get("status", "")) for a in alvos)
        ]

    if aluno:
        alvo = chave(aluno)
        filtrados = [
            r for r in filtrados
            if alvo in chave(r["nome"])
            or chave(r["nome"]) in alvo
        ]

    if plano:
        alvo = chave(plano)
        filtrados = [
            r for r in filtrados
            if alvo in chave(r["plano"])
            or chave(r["plano"]) in alvo
        ]

    if linhas is not None:
        filtrados = [
            r for r in filtrados
            if r["linha"] in linhas
        ]

    if so_erros:
        erros = carregar_erros_csv(csv_erros)
        if not erros:
            print(
                "AVISO: --so-erros sem erros no CSV "
                "(ou CSV ausente). Nada a processar."
            )
            return []
        # pelo nome: a lista colada pode mudar de ordem entre uma rodada e outra
        nomes_erro = {chave(nome) for nome in erros.values() if nome}
        if nomes_erro:
            filtrados = [r for r in filtrados if chave(r["nome"]) in nomes_erro]
        else:
            filtrados = [r for r in filtrados if r["linha"] in erros]

    return filtrados


def _chave_caminho_foto(caminho):
    return os.path.normcase(os.path.abspath(os.path.normpath(str(caminho))))


def aplicar_curso_manual_foto(job):
    """Aplica ao job um curso salvo para qualquer uma das fotos que ele usa."""

    try:
        dados = json.loads(CURSOS_POR_FOTO_FILE.read_text(encoding="utf-8"))
        cursos_por_foto = dados.get("cursos_por_foto", {}) if isinstance(dados, dict) else {}
        if not isinstance(cursos_por_foto, dict):
            cursos_por_foto = {}
    except (OSError, json.JSONDecodeError):
        cursos_por_foto = CURSOS_POR_FOTO

    cursos = {
        str(cursos_por_foto.get(_chave_caminho_foto(caminho), "")).strip()
        for caminho in job.get("fotos", {}).values()
        if str(cursos_por_foto.get(_chave_caminho_foto(caminho), "")).strip()
    }
    if not cursos:
        return job
    if len(cursos) > 1:
        raise Exception("FOTOS DO MESMO JOB ESTÃO ATRIBUÍDAS A CURSOS DIFERENTES.")

    curso = cursos.pop()
    simbolo, aviso = resolver_curso(curso)
    if not simbolo:
        raise Exception("CURSO ATRIBUÍDO SEM BRASÃO: " + curso)
    job["curso"] = curso
    job["curso_simbolo"] = simbolo
    job["cor_background"] = obter_cor_background(simbolo)
    return job


def montar_jobs_validacao(registro, modelo_nome, modelo):
    """Monta jobs sem executar Photoshop (para preflight / skip)."""

    jobs = []

    if chave(modelo_nome) == chave("ESSENCIAL BOOK 20X20"):
        jobs.append(criar_job_essencial_20(registro))

    elif chave(modelo_nome) == chave("STUDIO DAY BERLIM 24X34"):
        if BERLIM_GERAR_24:
            jobs.append(criar_job_berlim_24(registro))
        if BERLIM_GERAR_20:
            jobs.append(criar_job_essencial_20(registro))
        if not jobs:
            raise Exception(
                "BERLIM: berlim_gerar_24 e berlim_gerar_20 estao desligados."
            )

    elif chave(modelo_nome) == chave("PRIME BOOK 30X22"):
        jobs.append(criar_job_30x22(registro, "FRENTE"))
        jobs.append(criar_job_30x22(registro, "VERSO"))

    else:
        jobs.append(criar_job_30(registro, modelo_nome, modelo))

    return [aplicar_curso_manual_foto(job) for job in jobs]


def validar_registro(registro):
    """
    Valida uma linha sem abrir Photoshop.
    Retorna dict status OK/ERRO/AVISO.
    """

    avisos = []

    try:
        modelo_nome, modelo = encontrar_modelo(registro["plano"])

        if not modelo:
            raise Exception(
                "PLANO NAO CONFIGURADO:\n\n" + registro["plano"]
            )

        # Destinos acessiveis
        for pasta in OUTPUT_DIRS.values():
            try:
                pasta.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                raise Exception(
                    "PASTA DE DESTINO INACESSIVEL:\n\n"
                    + str(pasta)
                    + "\n\n"
                    + str(e)
                )

        jobs = montar_jobs_validacao(registro, modelo_nome, modelo)

        for job in jobs:
            template = Path(job["template"])
            if not template.is_file():
                raise Exception(
                    "PSD NAO ENCONTRADO:\n\n" + str(template)
                )

            for caminho_foto in job.get("fotos", {}).values():
                if not Path(caminho_foto).is_file():
                    raise Exception(
                        "FOTO INVALIDA:\n\n" + str(caminho_foto)
                    )

            if not job.get("curso_simbolo") and not registro.get("curso"):
                avisos.append("curso vazio")

        msg = "OK — " + modelo_nome
        if avisos:
            msg += " | AVISOS: " + "; ".join(avisos)

        return {
            "status": "OK" if not avisos else "AVISO",
            "linha": registro["linha"],
            "nome": registro["nome"],
            "curso": registro["curso"],
            "plano": registro["plano"],
            "mensagem": msg,
            "modelo": modelo_nome,
        }

    except Exception as e:
        return {
            "status": "ERRO",
            "linha": registro["linha"],
            "nome": registro["nome"],
            "curso": registro["curso"],
            "plano": registro["plano"],
            "mensagem": str(e),
            "modelo": None,
        }


def rodar_preflight(registros):
    print()
    print("=" * 75)
    print("PREFLIGHT — validacao sem Photoshop")
    print("=" * 75)

    resultados = []
    ok = aviso = erro = 0

    for registro in registros:
        print()
        print("-" * 75)
        print("LINHA:", registro["linha"], "|", registro["nome"])

        r = validar_registro(registro)
        resultados.append(r)

        print(r["status"] + ":", r["mensagem"][:200])

        if r["status"] == "OK":
            ok += 1
        elif r["status"] == "AVISO":
            aviso += 1
        else:
            erro += 1

    orfas = listar_fotos_orfas(registros)
    if orfas:
        print()
        print("FOTOS ORFAS (nao casam com alunos da planilha):", len(orfas))
        for f in orfas[:30]:
            print(" ", f.name)
        if len(orfas) > 30:
            print("  ... e mais", len(orfas) - 30)

    arquivo = Path(__file__).parent / "preflight_capas.csv"
    salvar_relatorio(resultados, arquivo)

    print()
    print("RESUMO PREFLIGHT — OK:", ok, "| AVISO:", aviso, "| ERRO:", erro)

    return resultados, erro == 0


def _arquivo_pronto(caminho):
    try:
        c = Path(caminho)
        return c.is_file() and c.stat().st_size > 0
    except Exception:
        return False


def ajustar_saidas_existentes(job):
    """
    Com skip_existing ligado, tira do job o que já existe no destino:
    - JPG e PSD prontos  -> "SKIP" (nem abre o Photoshop)
    - só JPG pronto      -> "SO_PSD" (não regrava o JPG)
    - só PSD pronto      -> "SO_JPG"
    - nada pronto        -> "TUDO"
    """

    quer_jpg = bool(job.get("saida_jpg"))
    quer_psd = bool(job.get("exportar_psd") and job.get("saida_psd"))

    if not SKIP_EXISTING:
        return "TUDO"

    jpg_pronto = quer_jpg and _arquivo_pronto(job["saida_jpg"])
    psd_pronto = quer_psd and _arquivo_pronto(job["saida_psd"])

    falta_jpg = quer_jpg and not jpg_pronto
    falta_psd = quer_psd and not psd_pronto

    if not falta_jpg and not falta_psd:
        return "SKIP"

    if not falta_jpg:
        job["saida_jpg"] = ""
        return "SO_PSD"

    if not falta_psd:
        job["exportar_psd"] = False
        return "SO_JPG"

    return "TUDO"


def _molde_da_foto(job, chave_foto):
    if chave_foto == "foto34":
        return "24X34"
    if "20X20" in str(job.get("modelo", "")).upper():
        return "20X20"
    return "30X20"


def preparar_foto_no_molde(caminho, tamanho_foto):
    """
    Os PSDs só trocam o conteúdo da foto, sem redimensionar: uma foto fora do
    tamanho do molde sai gigante na capa. Nesse caso recorta uma cópia temporária
    igual ao "Recortar fotos" (corte manual salvo ou rosto detectado), sem mexer
    na foto da pasta.
    """

    import recortar_fotos as rf

    caminho = Path(caminho)
    cfg = rf.carregar_config()
    molde = cfg["moldes"].get(tamanho_foto)
    if not molde or rf.no_molde(rf.tamanho_rapido(caminho), molde):
        return caminho

    img, icc, _ = rf.abrir_foto(rf.fonte_original(caminho, cfg))
    caixa = rf.caixa_manual(caminho, tamanho_foto, img.size, molde)
    modo = "corte manual salvo"
    if caixa is None:
        calib = {t: dict(rf.CALIBRACAO_PADRAO) for t in cfg["moldes"]}
        if rf.CALIBRACAO_FILE.is_file():
            calib.update(json.loads(rf.CALIBRACAO_FILE.read_text(encoding="utf-8")))
        caixa, _ = rf.caixa_automatica(img, tamanho_foto, cfg, calib)
        modo = "pelo rosto"
    if caixa is None:
        raise Exception(
            "FOTO FORA DO TAMANHO DO MOLDE E SEM ROSTO DETECTADO:\n\n"
            + caminho.name
            + "\n\nRecorte essa foto à mão na aba Fotos (recorte) antes de gerar."
        )

    ampliacao = rf.ampliacao_da_caixa(caixa, molde)
    if ampliacao > cfg["ampliacao_maxima"]:
        print("  AVISO: foto ampliada", format(ampliacao, ".0%"), "- pode perder nitidez:", caminho.name)

    pasta = Path(tempfile.gettempdir()) / "capas_fotos_no_molde"
    pasta.mkdir(parents=True, exist_ok=True)
    destino = pasta / (nome_seguro(caminho.stem) + "_" + tamanho_foto + ".jpg")
    rf.salvar_jpg(rf.recortar_imagem(img, caixa, molde), destino, icc, (300, 300))
    print("  FOTO FORA DO MOLDE: recortada automaticamente (" + modo + "):", caminho.name)
    return destino


def preparar_fotos_do_job(job):

    for chave_foto, caminho in list(job.get("fotos", {}).items()):
        if caminho:
            job["fotos"][chave_foto] = str(preparar_foto_no_molde(caminho, _molde_da_foto(job, chave_foto)))


def rodar_job(job, rotulo):
    """Executa um job respeitando o que já existe. Retorna a mensagem."""

    aplicar_curso_manual_foto(job)
    modo = ajustar_saidas_existentes(job)
    if modo != "SKIP":
        preparar_fotos_do_job(job)

    if modo == "SKIP":
        print(">>> PULANDO", rotulo, "(JPG e PSD ja existem)" if job.get("exportar_psd") else "(JPG ja existe)")
        return None

    if modo == "SO_PSD":
        print(">>> JPG JA EXISTE — GERANDO SO O PSD")
    elif modo == "SO_JPG":
        print(">>> PSD JA EXISTE — GERANDO SO O JPG")

    if not DRY_RUN:
        executar_photoshop(job)

    if modo == "SO_PSD":
        return rotulo + " PSD GERADO"
    if modo == "SO_JPG":
        return rotulo + " JPG GERADO"
    return rotulo + " GERADO"


def processar_registro(registro):
    """Processa um registro (Photoshop ou dry-run). Retorna dict resultado."""

    modelo_nome, modelo = encontrar_modelo(registro["plano"])

    if not modelo:
        raise Exception(
            "PLANO NAO CONFIGURADO:\n\n" + registro["plano"]
        )

    mensagens = []
    pulados = 0

    if chave(modelo_nome) == chave("ESSENCIAL BOOK 20X20"):

        job20 = criar_job_essencial_20(registro)

        print()
        print(">>> PROCESSANDO SOMENTE ESSENCIAL BOOK 20X20")
        print("PSD:", job20["template"])
        print("FOTO:", job20["fotos"]["foto"])
        print("JPG:", job20["saida_jpg"])

        msg = rodar_job(job20, "ESSENCIAL BOOK 20X20")
        if msg:
            mensagens.append(msg)
        else:
            pulados += 1

    elif chave(modelo_nome) == chave("STUDIO DAY BERLIM 24X34"):

        print()
        print("MODELO BERLIM DETECTADO.")
        print(
            "  berlim_gerar_24:", BERLIM_GERAR_24,
            "| berlim_gerar_20:", BERLIM_GERAR_20,
        )

        if BERLIM_GERAR_24:
            job24 = criar_job_berlim_24(registro)

            print()
            print(">>> PROCESSANDO 24X34")
            print("FOTO:", job24["fotos"]["foto34"])
            print("JPG:", job24["saida_jpg"])

            msg = rodar_job(job24, "24X34")
            if msg:
                mensagens.append(msg)
            else:
                pulados += 1

        if BERLIM_GERAR_20:
            job20 = criar_job_essencial_20(registro)

            print()
            print(">>> PROCESSANDO 20X20")
            print("PSD:", job20["template"])
            print("FOTO:", job20["fotos"]["foto"])
            print("JPG:", job20["saida_jpg"])

            msg = rodar_job(job20, "20X20")
            if msg:
                mensagens.append(msg)
            else:
                pulados += 1

        if not mensagens and not pulados:
            raise Exception(
                "BERLIM: nenhum tamanho habilitado "
                "(berlim_gerar_24 / berlim_gerar_20)."
            )

    elif chave(modelo_nome) == chave("PRIME BOOK 30X22"):

        for lado in ("FRENTE", "VERSO"):
            job = criar_job_30x22(registro, lado)

            print()
            print(">>> PROCESSANDO 30X22 " + lado)
            print("PSD:", job["template"])
            if job["fotos"].get("foto"):
                print("FOTO:", job["fotos"]["foto"])
            print("JPG:", job["saida_jpg"])

            msg = rodar_job(job, "30X22 " + lado)
            if msg:
                mensagens.append(msg)
            else:
                pulados += 1

    else:

        job = criar_job_30(registro, modelo_nome, modelo)

        print()
        print("MODELO:", modelo_nome)
        print("JPG:", job["saida_jpg"])
        print("PSD:", job["saida_psd"])

        msg = rodar_job(job, "CAPA")
        if msg:
            mensagens.append(msg)
        else:
            pulados += 1

    if not mensagens:
        return {
            "status": "SKIP",
            "linha": registro["linha"],
            "nome": registro["nome"],
            "curso": registro["curso"],
            "plano": registro["plano"],
            "mensagem": "JPG e PSD ja existentes — pulado" if EXPORT_PSD else "JPG ja existente — pulado",
        }

    return {
        "status": "OK",
        "linha": registro["linha"],
        "nome": registro["nome"],
        "curso": registro["curso"],
        "plano": registro["plano"],
        "mensagem": " + ".join(mensagens),
    }


def parse_args(argv=None):

    parser = argparse.ArgumentParser(
        description=(
            "Automacao de capas no Photoshop. "
            "Config: automacao_capas_config.json"
        )
    )

    parser.add_argument(
        "--config",
        default=None,
        help="Caminho do JSON de configuracao",
    )
    parser.add_argument(
        "--planilha",
        default=None,
        help="Arquivo .xlsx/.xlsm especifico",
    )
    parser.add_argument(
        "--aluno",
        default=None,
        help="Filtrar por nome (parcial)",
    )
    parser.add_argument(
        "--plano",
        default=None,
        help="Filtrar por plano/album (parcial)",
    )
    parser.add_argument(
        "--lote",
        default=None,
        help="Filtrar pela coluna LOTE (ex: 12 ou 12,13)",
    )
    parser.add_argument(
        "--status",
        default=None,
        help="Filtrar pela coluna STATUS (parcial, ex: Aprovado)",
    )
    parser.add_argument(
        "--linhas",
        default=None,
        help="Linhas da planilha: 10 ou 10-20 ou 10,12,15-18",
    )
    parser.add_argument(
        "--so-erros",
        action="store_true",
        help="So processa linhas com ERRO no resultado_capas.csv",
    )
    parser.add_argument(
        "--preflight",
        action="store_true",
        help="So valida (foto/curso/PSD/destino), sem Photoshop",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Monta jobs e valida, nao chama Photoshop",
    )
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        default=None,
        help="Pula se JPG de saida ja existe",
    )
    parser.add_argument(
        "--no-skip-existing",
        action="store_true",
        help="Reprocessa mesmo se JPG ja existe",
    )
    parser.add_argument(
        "--berlim-24",
        action="store_true",
        default=None,
        help="Berlim: gerar so 24x34 (desliga 20x20)",
    )
    parser.add_argument(
        "--berlim-20",
        action="store_true",
        default=None,
        help="Berlim: gerar so 20x20 (desliga 24x34)",
    )
    parser.add_argument(
        "--berlim-ambos",
        action="store_true",
        help="Berlim: gerar 24x34 e 20x20",
    )
    parser.add_argument(
        "--export-psd",
        action="store_true",
        help="Tambem exporta PSD",
    )
    parser.add_argument(
        "--no-export-jpg",
        action="store_true",
        help="Nao exporta JPG",
    )
    parser.add_argument(
        "--no-export-psd",
        action="store_true",
        help="Nao exporta PSD",
    )
    parser.add_argument(
        "--parar-arquivo",
        default=None,
        help="Se este arquivo aparecer, para antes da proxima capa (botao Parar da UI)",
    )
    parser.add_argument(
        "--ui",
        action="store_true",
        help="Abre UI minima (tkinter) para montar o comando",
    )

    return parser.parse_args(argv)


def aplicar_args(args):

    global DRY_RUN, EXPORT_JPG, EXPORT_PSD, SKIP_EXISTING
    global BERLIM_GERAR_24, BERLIM_GERAR_20, PLANILHA_ESCOLHIDA

    carregar_config(args.config)

    if args.planilha:
        PLANILHA_ESCOLHIDA = Path(args.planilha)

    if args.dry_run:
        DRY_RUN = True

    if args.export_psd:
        EXPORT_PSD = True

    if args.no_export_jpg:
        EXPORT_JPG = False

    if args.no_export_psd:
        EXPORT_PSD = False

    if args.skip_existing is True:
        SKIP_EXISTING = True

    if args.no_skip_existing:
        SKIP_EXISTING = False

    if args.berlim_ambos:
        BERLIM_GERAR_24 = True
        BERLIM_GERAR_20 = True
    elif args.berlim_24:
        BERLIM_GERAR_24 = True
        BERLIM_GERAR_20 = False
    elif args.berlim_20:
        BERLIM_GERAR_24 = False
        BERLIM_GERAR_20 = True


def abrir_ui():
    """UI minima para escolher filtros e rodar preflight/processamento."""

    import tkinter as tk
    from tkinter import filedialog, ttk

    root = tk.Tk()
    root.title("Automacao de Capas")
    root.geometry("520x420")

    vars_ = {
        "planilha": tk.StringVar(),
        "aluno": tk.StringVar(),
        "plano": tk.StringVar(),
        "linhas": tk.StringVar(),
        "so_erros": tk.BooleanVar(value=False),
        "preflight": tk.BooleanVar(value=False),
        "dry_run": tk.BooleanVar(value=False),
        "skip_existing": tk.BooleanVar(value=SKIP_EXISTING),
        "berlim": tk.StringVar(value="ambos"),
    }

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)

    def escolher_planilha():
        path = filedialog.askopenfilename(
            title="Planilha",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
            initialdir=str(PLANILHA_DIR),
        )
        if path:
            vars_["planilha"].set(path)

    row = 0
    ttk.Label(frm, text="Planilha").grid(row=row, column=0, sticky="w")
    ttk.Entry(frm, textvariable=vars_["planilha"], width=48).grid(
        row=row, column=1, sticky="we"
    )
    ttk.Button(frm, text="...", width=3, command=escolher_planilha).grid(
        row=row, column=2
    )

    row += 1
    ttk.Label(frm, text="Aluno").grid(row=row, column=0, sticky="w")
    ttk.Entry(frm, textvariable=vars_["aluno"], width=48).grid(
        row=row, column=1, columnspan=2, sticky="we"
    )

    row += 1
    ttk.Label(frm, text="Plano").grid(row=row, column=0, sticky="w")
    ttk.Entry(frm, textvariable=vars_["plano"], width=48).grid(
        row=row, column=1, columnspan=2, sticky="we"
    )

    row += 1
    ttk.Label(frm, text="Linhas").grid(row=row, column=0, sticky="w")
    ttk.Entry(frm, textvariable=vars_["linhas"], width=48).grid(
        row=row, column=1, columnspan=2, sticky="we"
    )

    row += 1
    ttk.Checkbutton(
        frm, text="So erros do CSV", variable=vars_["so_erros"]
    ).grid(row=row, column=1, sticky="w")

    row += 1
    ttk.Checkbutton(
        frm, text="Preflight (sem Photoshop)", variable=vars_["preflight"]
    ).grid(row=row, column=1, sticky="w")

    row += 1
    ttk.Checkbutton(
        frm, text="Dry-run", variable=vars_["dry_run"]
    ).grid(row=row, column=1, sticky="w")

    row += 1
    ttk.Checkbutton(
        frm, text="Pular JPG existente", variable=vars_["skip_existing"]
    ).grid(row=row, column=1, sticky="w")

    row += 1
    ttk.Label(frm, text="Berlim").grid(row=row, column=0, sticky="w")
    berlim_box = ttk.Combobox(
        frm,
        textvariable=vars_["berlim"],
        values=["ambos", "so 24x34", "so 20x20"],
        state="readonly",
        width=20,
    )
    berlim_box.grid(row=row, column=1, sticky="w")

    resultado_ui = {"argv": None}

    def montar_e_fechar():
        argv = []
        if vars_["planilha"].get().strip():
            argv += ["--planilha", vars_["planilha"].get().strip()]
        if vars_["aluno"].get().strip():
            argv += ["--aluno", vars_["aluno"].get().strip()]
        if vars_["plano"].get().strip():
            argv += ["--plano", vars_["plano"].get().strip()]
        if vars_["linhas"].get().strip():
            argv += ["--linhas", vars_["linhas"].get().strip()]
        if vars_["so_erros"].get():
            argv.append("--so-erros")
        if vars_["preflight"].get():
            argv.append("--preflight")
        if vars_["dry_run"].get():
            argv.append("--dry-run")
        if vars_["skip_existing"].get():
            argv.append("--skip-existing")
        else:
            argv.append("--no-skip-existing")
        b = vars_["berlim"].get()
        if b == "so 24x34":
            argv.append("--berlim-24")
        elif b == "so 20x20":
            argv.append("--berlim-20")
        else:
            argv.append("--berlim-ambos")
        resultado_ui["argv"] = argv
        root.destroy()

    row += 1
    ttk.Button(frm, text="Executar", command=montar_e_fechar).grid(
        row=row, column=1, pady=16, sticky="w"
    )

    root.mainloop()
    return resultado_ui["argv"]


def main(argv=None):

    args = parse_args(argv)

    if args.ui:
        ui_argv = abrir_ui()
        if ui_argv is None:
            print("UI cancelada.")
            return
        # Reparse com opcoes da UI (mantem --config se veio na CLI)
        base = []
        if args.config:
            base += ["--config", args.config]
        args = parse_args(base + ui_argv)

    aplicar_args(args)

    print()
    print("=" * 75)
    print("AUTOMACAO DE CAPAS")
    print("=" * 75)
    print("DRY_RUN:", DRY_RUN)
    print("EXPORT_JPG:", EXPORT_JPG)
    print("EXPORT_PSD:", EXPORT_PSD)
    print("SKIP_EXISTING:", SKIP_EXISTING)
    print("FOTO_MATCH_FLEXIVEL:", FOTO_MATCH_FLEXIVEL)
    print("BERLIM_GERAR_24:", BERLIM_GERAR_24)
    print("BERLIM_GERAR_20:", BERLIM_GERAR_20)
    print("PHOTOS_DIR:", PHOTOS_DIR)
    print("PSD_DIR:", PSD_DIR)

    registros = ler_planilha(
        args.planilha or PLANILHA_ESCOLHIDA
    )

    print()
    print("TOTAL NA PLANILHA:", len(registros))

    linhas = parse_faixa_linhas(args.linhas)

    registros = filtrar_registros(
        registros,
        aluno=args.aluno,
        plano=args.plano,
        linhas=linhas,
        so_erros=args.so_erros,
        lote=args.lote,
        status=args.status,
    )

    print("APOS FILTROS:", len(registros))

    if not registros:
        print("Nenhum registro para processar.")
        return

    if args.preflight:
        rodar_preflight(registros)
        return

    # Preflight automatico antes do lote (lista problemas, nao bloqueia)
    print()
    print("Checagem rapida antes do processamento...")
    pre_ok = 0
    pre_erro = 0
    for registro in registros:
        r = validar_registro(registro)
        if r["status"] == "ERRO":
            pre_erro += 1
            print(
                "  FALHA LINHA",
                    registro["linha"],
                    registro["nome"],
                "->",
                r["mensagem"][:120].replace("\n", " "),
            )
        else:
            pre_ok += 1

    print("  Pre-check OK:", pre_ok, "| com erro:", pre_erro)

    resultados = []
    total = len(registros)

    parar = Path(args.parar_arquivo) if args.parar_arquivo else None

    for i, registro in enumerate(registros, start=1):

        if parar and parar.exists():
            print()
            print("PARADO PELO USUARIO antes da capa", i, "/", total,
                  "- as capas ja feitas foram mantidas.")
            break

        print()
        print("-" * 75)
        print("PROGRESSO:", i, "/", total)
        print("LINHA:", registro["linha"])
        print("NOME:", registro["nome"])
        print("CURSO:", registro["curso"])
        print("PLANO:", registro["plano"])

        try:
            resultado = processar_registro(registro)
            resultados.append(resultado)
            print("STATUS:", resultado["status"], "-", resultado["mensagem"])

        except Exception as e:
            print()
            print("ERRO:")
            print(str(e))

            resultados.append({
                "status": "ERRO",
                "linha": registro["linha"],
                "nome": registro["nome"],
                "curso": registro["curso"],
                "plano": registro["plano"],
                "mensagem": str(e),
            })

    salvar_relatorio(resultados)

    ok = sum(1 for r in resultados if r["status"] == "OK")
    skip = sum(1 for r in resultados if r["status"] == "SKIP")
    erro = sum(1 for r in resultados if r["status"] == "ERRO")

    print()
    print("=" * 75)
    print("PROCESSAMENTO FINALIZADO.")
    print("OK:", ok, "| SKIP:", skip, "| ERRO:", erro)
    print("=" * 75)


# ============================================================
# INICIO
# ============================================================

if __name__ == "__main__":

    # Carrega config padrao mesmo sem CLI explicita
    if CONFIG_FILE.is_file():
        try:
            carregar_config()
        except Exception as e:
            print("AVISO ao carregar config:", e)

    main()
