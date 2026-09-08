import importlib.util
from pathlib import Path
from unittest.mock import Mock

import pandas as pd
from docx import Document

from docspropostas import data_logic
from docspropostas.constants import ALVARA_DATA


def demo(tmp_path):
    path = Path(__file__).resolve().parents[1] / "examples" / "create_demo.py"
    spec = importlib.util.spec_from_file_location("demo", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.create_demo(tmp_path)


def context():
    return {
        "ID_INTERNO": "DEMO-001",
        "NOME_DA_EMPREITADA": "A & B <exemplo>",
        "LOCAL_DA_OBRA": "Local",
        "DURACAO_DA_EMPREITADA": "90",
        "DOCUMENTOS_DA_PROPOSTA": "Documento A\nDocumento B",
        "DATA_HOJE": "01 de janeiro",
        "ALVARA_SELECOES": [],
    }


def test_demo_contract_and_render(tmp_path, monkeypatch):
    root = demo(tmp_path)
    monkeypatch.setattr(data_logic, "messagebox", Mock())
    df = pd.read_excel(
        root / "proposta-demo.xlsx", sheet_name="Doc Proposta", header=None
    )
    assert df.iloc[2, 10] == "DEMO-001"
    assert "Plano de trabalhos" in data_logic.extrair_lista_de_documentos_pandas(
        df, 9, 24, 44
    )
    assert data_logic.gerar_documentos(
        context(), str(root / "saida"), str(root / "modelos"), Mock(), Mock()
    )
    output = next((root / "saida").glob("*.docx"))
    text = "\n".join(p.text for p in Document(output).paragraphs)
    assert "A & B <exemplo>" in text
    assert "{{" not in text


def test_missing_template_variable_is_reported(tmp_path, monkeypatch):
    root = demo(tmp_path)
    monkeypatch.setattr(data_logic, "messagebox", Mock())
    assert not data_logic.gerar_documentos(
        {}, str(root / "saida"), str(root / "modelos"), Mock(), Mock()
    )
    assert not list((root / "saida").glob("*.docx"))


def test_existing_document_not_overwritten(tmp_path, monkeypatch):
    root = demo(tmp_path)
    monkeypatch.setattr(data_logic, "messagebox", Mock())
    args = (context(), str(root / "saida"), str(root / "modelos"), Mock(), Mock())
    assert data_logic.gerar_documentos(*args)
    output = next((root / "saida").glob("*.docx"))
    original = output.read_bytes()
    assert not data_logic.gerar_documentos(*args)
    assert output.read_bytes() == original


def test_no_company_alvara_classes_distributed():
    assert all(item["class"] == "" for items in ALVARA_DATA.values() for item in items)
