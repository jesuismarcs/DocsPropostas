from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parents[2] / ".env")

import re
import urllib.request
import sys
import os
from pathlib import Path


def resource_path(relative_path):
    """Obtém o caminho absoluto para um recurso, funciona para desenvolvimento e para o PyInstaller"""
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # _MEIPASS não está definido, estamos a correr em modo de desenvolvimento
        base_path = str(Path(__file__).resolve().parents[2] / "resources")

    return os.path.join(base_path, relative_path)


def verificar_chave_acesso():
    """Verifica a chave de acesso quando uma fonte de licenca esta configurada."""
    chave_esperada = os.getenv("DOCSPROPOSTAS_LICENSE_KEY") or os.getenv(
        "LICENSE_KEY_EXPECTED"
    )
    ficheiro_chave = os.getenv("DOCSPROPOSTAS_LICENSE_FILE")
    url_chave = os.getenv("DOCSPROPOSTAS_LICENSE_URL") or os.getenv("LICENSE_KEY_URL")

    if not chave_esperada:
        return os.getenv("APP_ENV", "development").lower() != "production"

    try:
        if ficheiro_chave:
            chave_online = Path(ficheiro_chave).read_text(encoding="utf-8").strip()
        elif url_chave:
            req = urllib.request.Request(
                url_chave,
                headers={
                    "Cache-Control": "no-cache",
                    "Pragma": "no-cache",
                    "User-Agent": "Mozilla/5.0",
                },
            )
            with urllib.request.urlopen(req, timeout=10) as resposta:
                chave_online = resposta.read().decode("utf-8").strip()
        else:
            return False
        return chave_online == chave_esperada
    except Exception:
        return False


def sanitizar_nome_ficheiro(nome):
    """Remove caracteres inválidos de um nome de ficheiro."""
    if nome is None:
        return ""
    return re.sub(r'[<>:"/\\|?*]', "_", str(nome))
