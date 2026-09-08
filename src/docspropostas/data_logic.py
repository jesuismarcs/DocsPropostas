import os
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
    contexto, output_path, template_path, log_callback, progress_callback
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

        total_files = len(modelos_encontrados)
        progress_callback(0, total_files)

        log_callback("\nA processar modelos:")
        for i, nome_template in enumerate(modelos_encontrados):
            log_callback(f"- {nome_template}")
            doc = DocxTemplate(os.path.join(template_path, nome_template))
            doc.render(
                contexto,
                jinja_env=Environment(undefined=StrictUndefined),
                autoescape=True,
            )

            id_interno = sanitizar_nome_ficheiro(contexto.get("ID_INTERNO") or "SEM-ID")
            nome_base = sanitizar_nome_ficheiro(os.path.splitext(nome_template)[0])
            local_obra = sanitizar_nome_ficheiro(
                contexto.get("LOCAL_DA_OBRA") or "SEM-LOCAL"
            )
            nome_final = f"{id_interno}_{nome_base}_{local_obra}.docx"

            destino = os.path.join(output_path, nome_final)
            if os.path.exists(destino):
                raise FileExistsError(
                    f"Já existe: {nome_final}. Escolha outra pasta de saída."
                )
            doc.save(destino)
            log_callback(f"   => Gerado: {nome_final}")
            progress_callback(i + 1, total_files)

        log_callback("\n--- PROCESSO CONCLUÍDO COM SUCESSO ---")
        messagebox.showinfo(
            "Sucesso", f"Documentos gerados com sucesso na pasta:\n{output_path}"
        )
        return True

    except Exception as e:
        log_callback(f"ERRO: {e}")
        messagebox.showerror("Erro Inesperado", f"Ocorreu um erro:\n{e}")
        return False
