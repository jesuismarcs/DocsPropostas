import os
import tempfile
import shutil
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np
from docxtpl import DocxTemplate
from jinja2 import Environment, StrictUndefined
from tkinter import messagebox
from .utils import sanitizar_nome_ficheiro


def extrair_lista_de_documentos_pandas(df, col_index, start_row, end_row):
    """Extrai uma lista de uma coluna específica de um DataFrame."""
    return "\n".join(
        df.iloc[start_row : end_row + 1, col_index].dropna().astype(str).tolist()
    )


def gerar_documentos(
    contexto, output_path, template_path, log_callback, progress_callback, options=None
):
    """Gera os documentos .docx a partir dos modelos."""
    try:
        log_callback("--- INÍCIO DA GERAÇÃO ---")

        if "Nenhuma" in output_path or not os.path.isdir(output_path):
            messagebox.showerror(
                "Erro", "Por favor, selecione uma Pasta de Saída válida."
            )
            return False

        if not os.path.isdir(template_path):
            messagebox.showerror(
                "Erro", f"A pasta de modelos não foi encontrada em:\n{template_path}"
            )
            return False

        modelos_encontrados = sorted(
            f
            for f in os.listdir(template_path)
            if f.lower().endswith(".docx") and not f.startswith("~$")
        )
        if not modelos_encontrados:
            messagebox.showwarning(
                "Aviso",
                f"Nenhum modelo (.docx) válido encontrado em '{template_path}'.",
            )
            log_callback("AVISO: Nenhum modelo (.docx) válido foi encontrado.")
            return False

        options = options or {}
        selected = [
            x.strip()
            for x in options.get("selected_models", "").split(";")
            if x.strip()
        ]
        if selected:
            missing = set(selected) - set(modelos_encontrados)
            if missing:
                raise ValueError("Modelos inexistentes: " + ", ".join(sorted(missing)))
            modelos_encontrados = [x for x in modelos_encontrados if x in selected]
        total_files = len(modelos_encontrados)
        progress_callback(0, total_files)

        destinations = []
        published = []
        with tempfile.TemporaryDirectory(dir=output_path, prefix=".docs-") as staging:
            for i, nome_template in enumerate(modelos_encontrados):
                doc = DocxTemplate(os.path.join(template_path, nome_template))
                doc.render(
                    contexto,
                    jinja_env=Environment(undefined=StrictUndefined),
                    autoescape=True,
                )
                fields = {
                    **contexto,
                    "ID_INTERNO": contexto.get("ID_INTERNO") or "SEM-ID",
                    "LOCAL_DA_OBRA": contexto.get("LOCAL_DA_OBRA") or "SEM-LOCAL",
                    "modelo": Path(nome_template).stem,
                }
                name = sanitizar_nome_ficheiro(
                    options.get(
                        "filename", "{ID_INTERNO}_{modelo}_{LOCAL_DA_OBRA}.docx"
                    ).format_map(fields)
                )
                if not name.lower().endswith(".docx"):
                    name += ".docx"
                target = Path(output_path) / name
                if target.exists() or any(
                    t.name.casefold() == name.casefold() for _, t in destinations
                ):
                    raise FileExistsError(f"Já existe ou está duplicado: {name}")
                temp = Path(staging) / name
                doc.save(temp)
                destinations.append((temp, target))
                progress_callback(i + 1, total_files)
            try:
                for temp, target in destinations:
                    with target.open("xb") as destination:
                        published.append(target)
                        with temp.open("rb") as source:
                            shutil.copyfileobj(source, destination)
                    log_callback("Gerado: " + target.name)
            except Exception:
                for target in published:
                    target.unlink(missing_ok=True)
                raise

        log_callback("\n--- PROCESSO CONCLUÍDO COM SUCESSO ---")
        messagebox.showinfo(
            "Sucesso", f"Documentos gerados com sucesso na pasta:\n{output_path}"
        )
        return True

    except Exception as e:
        log_callback(f"ERRO: {e}")
        messagebox.showerror("Erro Inesperado", f"Ocorreu um erro:\n{e}")
        return False
