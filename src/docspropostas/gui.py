import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re  # <- Adicionado
import pandas as pd
import numpy as np
from datetime import datetime
from PIL import Image, ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD

import sv_ttk
from . import config_manager
from toolkit.profiles import ProfileStore
from toolkit.mapping import Mapper
from toolkit.settings_ui import SettingsWindow
from .alvara_gui import AlvaraApp
from .data_logic import extrair_lista_de_documentos_pandas, gerar_documentos
from .utils import resource_path
from .constants import ALVARA_DATA  # <- Adicionado


class ToolTip:
    def __init__(self, widget, text):
        self.widget, self.text, self.tooltip_window = widget, text, None
        widget.bind("<Enter>", self.show_tooltip)
        widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        x, y = 0, self.widget.winfo_height()
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(
            self.tooltip_window,
            text=self.text,
            justify="left",
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            font=("tahoma", "8", "normal"),
            padding=(2, 2, 2, 2),
        )
        label.pack(ipadx=1)

    def hide_tooltip(self, event):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None


class MainApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.settings = config_manager.load_config()
        self.profile_store = ProfileStore("DocsPropostas")
        self.profile_store.migrate(self.settings)
        self.profile = self.profile_store.active()
        self.extra_context = {}
        self.settings.update(
            theme=self.profile["preferences"]["theme"],
            template_path=self.profile["tools"]["docs"]["templates"],
            output_path=self.profile["tools"]["docs"]["output"],
        )
        sv_ttk.set_theme(self.settings.get("theme", "light"))
        self.title("Gerador de Documentos de Proposta v2.3")
        self.geometry("900x800")  # Versão incrementada
        self.icons = self._load_icons()
        self.caminho_excel = tk.StringVar(
            value="Arraste um ficheiro Excel para aqui ou selecione..."
        )
        self.caminho_saida = tk.StringVar(
            value=self.settings.get("output_path") or "Nenhuma pasta selecionada"
        )
        self.alvara_selecoes, self.alvara_status = (
            [],
            tk.StringVar(value="0 itens selecionados"),
        )
        self.id_interno_var, self.nome_empreitada_var = tk.StringVar(), tk.StringVar()
        self.local_obra_var, self.duracao_empreitada_var = (
            tk.StringVar(),
            tk.StringVar(),
        )
        ttk.Button(
            self,
            text="Definições · " + self.profile["name"],
            command=self.open_settings,
        ).pack(anchor="e", padx=12, pady=4)
        self._criar_widgets()
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self.handle_drop)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _load_icons(self):
        if not os.path.isfile(resource_path("assets/excel.png")):
            return {}
        try:
            return {
                "excel": ImageTk.PhotoImage(
                    Image.open(resource_path("assets/excel.png")).resize(
                        (20, 20), Image.Resampling.LANCZOS
                    )
                ),
                "folder": ImageTk.PhotoImage(
                    Image.open(resource_path("assets/folder.png")).resize(
                        (20, 20), Image.Resampling.LANCZOS
                    )
                ),
                "clear": ImageTk.PhotoImage(
                    Image.open(resource_path("assets/clear.png")).resize(
                        (20, 20), Image.Resampling.LANCZOS
                    )
                ),
                "alvara": ImageTk.PhotoImage(
                    Image.open(resource_path("assets/alvara.png")).resize(
                        (20, 20), Image.Resampling.LANCZOS
                    )
                ),
                "generate": ImageTk.PhotoImage(
                    Image.open(resource_path("assets/generate.png")).resize(
                        (20, 20), Image.Resampling.LANCZOS
                    )
                ),
            }
        except Exception as e:
            primeiro_caminho_tentado = resource_path("assets/excel.png")
            messagebox.showerror(
                "Erro de Diagnóstico ao Carregar Ícones",
                f"A aplicação não conseguiu carregar os ícones.\n\nO caminho completo que foi tentado é:\n'{primeiro_caminho_tentado}'\n\nO erro específico foi:\n{type(e).__name__}: {e}",
            )
            return {}

    def _criar_widgets(self):
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(
            header_frame, text="Gerador de Propostas", font=("Arial", 18, "bold")
        ).pack(side=tk.LEFT)
        self.theme_switch = ttk.Checkbutton(
            header_frame,
            text="Modo Noturno",
            style="Switch.TCheckbutton",
            command=self.toggle_theme,
        )
        if sv_ttk.get_theme() == "dark":
            self.theme_switch.state(["selected"])
        self.theme_switch.pack(side=tk.RIGHT)
        ToolTip(self.theme_switch, "Alternar entre modo claro e escuro")
        file_frame = ttk.LabelFrame(
            main_frame, text="1. Ficheiros e Pastas", padding=(10, 5)
        )
        file_frame.pack(fill=tk.X, pady=5)
        excel_label = ttk.Label(
            file_frame, textvariable=self.caminho_excel, style="Card.TLabel", padding=10
        )
        excel_label.grid(row=0, column=0, sticky="ew", padx=5, pady=5, columnspan=2)
        ToolTip(excel_label, "Arraste e largue o ficheiro Excel aqui")
        btn_select_excel = ttk.Button(
            file_frame,
            text="Selecionar Excel...",
            image=self.icons.get("excel"),
            compound="left",
            command=self.selecionar_e_carregar_excel,
        )
        btn_select_excel.grid(row=0, column=2, padx=5)
        ToolTip(
            btn_select_excel, "Selecionar o ficheiro Excel com os dados da proposta"
        )
        ttk.Button(
            file_frame, text="Pasta de Modelos...", command=self.selecionar_modelos
        ).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Label(
            file_frame, textvariable=self.caminho_saida, relief="sunken", padding=2
        ).grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        btn_select_output = ttk.Button(
            file_frame,
            text="Alterar Saída...",
            image=self.icons.get("folder"),
            compound="left",
            command=self.selecionar_saida,
        )
        btn_select_output.grid(row=1, column=1, columnspan=2, padx=5)
        ToolTip(
            btn_select_output, "Selecionar a pasta onde os documentos serão guardados"
        )
        file_frame.columnconfigure(0, weight=1)
        data_frame = ttk.LabelFrame(
            main_frame, text="2. Dados da Proposta", padding=(10, 5)
        )
        data_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        ttk.Label(data_frame, text="ID Interno:").grid(
            row=0, column=0, sticky="w", padx=5, pady=3
        )
        ttk.Entry(data_frame, textvariable=self.id_interno_var).grid(
            row=0, column=1, sticky="ew"
        )
        ttk.Label(data_frame, text="Nome da Empreitada:").grid(
            row=1, column=0, sticky="w", padx=5, pady=3
        )
        ttk.Entry(data_frame, textvariable=self.nome_empreitada_var).grid(
            row=1, column=1, sticky="ew"
        )
        ttk.Label(data_frame, text="Local da Obra:").grid(
            row=2, column=0, sticky="w", padx=5, pady=3
        )
        ttk.Entry(data_frame, textvariable=self.local_obra_var).grid(
            row=2, column=1, sticky="ew"
        )
        ttk.Label(data_frame, text="Duração da Empreitada:").grid(
            row=3, column=0, sticky="w", padx=5, pady=3
        )
        ttk.Entry(data_frame, textvariable=self.duracao_empreitada_var).grid(
            row=3, column=1, sticky="ew"
        )
        btn_clear = ttk.Button(
            data_frame,
            text="Limpar Dados",
            image=self.icons.get("clear"),
            compound="left",
            command=self.limpar_campos_dados,
        )
        btn_clear.grid(row=0, column=2, rowspan=4, padx=10, sticky="ns")
        ToolTip(btn_clear, "Limpa todos os campos do formulário")
        data_frame.columnconfigure(1, weight=1)
        docs_frame = ttk.LabelFrame(
            data_frame, text="Documentos da Proposta (um por linha)", padding=(10, 5)
        )
        docs_frame.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=5)
        self.docs_text = scrolledtext.ScrolledText(
            docs_frame, wrap=tk.WORD, height=8, relief="flat"
        )
        self.docs_text.pack(fill=tk.BOTH, expand=True)
        data_frame.rowconfigure(4, weight=1)
        alvara_frame = ttk.LabelFrame(
            main_frame, text="3. Habilitações de Alvará", padding=(10, 5)
        )
        alvara_frame.pack(fill=tk.X, pady=5)
        btn_alvara = ttk.Button(
            alvara_frame,
            text="Inserir/Editar Habilitações...",
            image=self.icons.get("alvara"),
            compound="left",
            command=self.abrir_janela_alvara,
        )
        btn_alvara.pack(side=tk.LEFT, padx=5)
        ToolTip(
            btn_alvara, "Adicionar ou editar as habilitações do alvará para a proposta"
        )
        ttk.Label(alvara_frame, textvariable=self.alvara_status, wraplength=600).pack(
            side=tk.LEFT, padx=5, pady=5
        )
        run_frame = ttk.LabelFrame(
            main_frame, text="4. Geração e Progresso", padding=(10, 5)
        )
        run_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        btn_generate = ttk.Button(
            run_frame,
            text="Gerar Documentos",
            style="Accent.TButton",
            image=self.icons.get("generate"),
            compound="left",
            command=self.processar_e_gerar,
        )
        btn_generate.pack(pady=10)
        ToolTip(btn_generate, "Iniciar o processo de geração dos documentos")
        self.progress_bar = ttk.Progressbar(
            run_frame, orient="horizontal", mode="determinate"
        )
        self.progress_bar.pack(fill=tk.X, padx=5, pady=(5, 0))
        self.log_area = scrolledtext.ScrolledText(
            run_frame, state="disabled", height=6, wrap=tk.WORD, relief="flat"
        )
        self.log_area.pack(fill=tk.BOTH, expand=True, pady=5)

    # ### INÍCIO DA NOVA FUNCIONALIDADE ###
    def _parsear_alvara_de_excel(self, alvara_strings):
        parsed_data = []
        # Mapa para encontrar rapidamente a categoria pelo seu número
        catalog = self.profile["tools"]["docs"]["catalog"] or ALVARA_DATA
        cat_map = {
            re.match(r"\d+", key).group(): key
            for key in catalog
            if re.match(r"\d+", key)
        }

        for s in alvara_strings:
            s_lower = s.lower()
            # Expressão regular para capturar os números das subcategorias e o número da categoria
            match = re.search(r"([\d,\s]+)\s*sub,\s*(\d+)\s*cat", s_lower)
            if not match:
                continue

            sub_nums_str, cat_num = match.groups()
            sub_nums = [num.strip() for num in sub_nums_str.split(",")]
            valor = "Valor Global" if "valor global" in s_lower else "1"

            full_cat_name = cat_map.get(cat_num)
            if not full_cat_name:
                self.log(
                    f"AVISO: Categoria de alvará nº '{cat_num}' não encontrada. A ignorar linha: '{s}'"
                )
                continue

            # Mapa para encontrar subcategorias da categoria atual
            sub_map = {
                re.match(r"\d+", sub["name"]).group(): sub
                for sub in catalog[full_cat_name]
                if re.match(r"\d+", sub["name"])
            }

            for sub_num in sub_nums:
                sub_info = sub_map.get(sub_num)
                if not sub_info:
                    self.log(
                        f"AVISO: Subcategoria nº '{sub_num}' na categoria '{cat_num}' não encontrada. A ignorar."
                    )
                    continue

                cat_num_part = full_cat_name.split(" ")[0]
                sub_prefix_display = sub_info["name"].split(" ")[0]
                sub_desc_text = " ".join(sub_info["name"].split(" ")[1:]).lstrip("- ")

                parsed_data.append(
                    {
                        "cat_num_part": cat_num_part,
                        "full_cat_name": full_cat_name,  # Adicionado para facilitar o preenchimento
                        "sub_prefix_display": sub_prefix_display,
                        "full_sub_name": sub_info[
                            "name"
                        ],  # Adicionado para facilitar o preenchimento
                        "class_val": f"Classe {sub_info['class']}"
                        if sub_info["class"]
                        else "A preencher",
                        "sub_desc": sub_desc_text,
                        "valor_executar": f"{valor} €",
                    }
                )
        return parsed_data

    # ### FIM DA NOVA FUNCIONALIDADE ###

    def toggle_theme(self):
        if self.theme_switch.instate(["selected"]):
            sv_ttk.set_theme("dark")
            self.settings["theme"] = "dark"
        else:
            sv_ttk.set_theme("light")
            self.settings["theme"] = "light"

    def log(self, message):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")
        self.update_idletasks()

    def handle_drop(self, event):
        filepath = event.data.strip()
        if filepath.startswith("{") and filepath.endswith("}"):
            filepath = filepath[1:-1]
        if filepath.lower().endswith((".xlsx", ".xlsm", ".xls")):
            self._carregar_dados_excel(filepath)
        else:
            self.log("ERRO: O ficheiro arrastado não é um ficheiro Excel válido.")
            messagebox.showerror(
                "Ficheiro Inválido",
                "Por favor, arraste um ficheiro Excel (.xlsx, .xlsm, .xls).",
            )

    def limpar_campos_dados(self):
        (
            self.id_interno_var.set(""),
            self.nome_empreitada_var.set(""),
            self.local_obra_var.set(""),
            self.duracao_empreitada_var.set(""),
        )
        self.docs_text.delete("1.0", tk.END)
        self.caminho_excel.set("Arraste um ficheiro Excel para aqui ou selecione...")
        self.alvara_selecoes = []  # Limpar alvará também
        self._update_alvara_display()
        self.log("Campos de dados limpos.")

    def _carregar_dados_excel(self, path):
        if not path:
            return
        self.caminho_excel.set(path)
        if (
            not self.settings.get("output_path")
            or "Nenhuma" in self.caminho_saida.get()
        ):
            self.caminho_saida.set(os.path.dirname(path))
            self.settings["output_path"] = os.path.dirname(path)
        try:
            self.log(f"A carregar dados de: {os.path.basename(path)}")
            mapper = Mapper(self.profile, path)
            try:
                data = mapper.proposal()
            finally:
                mapper.close()
            self.extra_context = data
            self.id_interno_var.set(data.get("ID_INTERNO", ""))
            self.nome_empreitada_var.set(data.get("NOME_DA_EMPREITADA", ""))
            self.local_obra_var.set(data.get("LOCAL_DA_OBRA", ""))
            self.duracao_empreitada_var.set(data.get("DURACAO_DA_EMPREITADA", ""))
            documents = data.get("DOCUMENTOS_DA_PROPOSTA", [])
            self.docs_text.delete("1.0", tk.END)
            self.docs_text.insert(
                "1.0",
                "\n".join(documents) if isinstance(documents, list) else documents,
            )
            alvara = data.get("HABILITACOES", [])
            self.alvara_selecoes = []
            if self.profile["tools"]["docs"]["alvara"] and alvara:
                self.alvara_selecoes = (
                    alvara
                    if isinstance(alvara[0], dict)
                    else self._parsear_alvara_de_excel(alvara)
                )
            self._update_alvara_display()
            self.log("Dados carregados pelo perfil " + self.profile["name"])

        except Exception as e:
            self.log(f"ERRO ao ler o Excel: {e}")
            messagebox.showerror(
                "Erro ao ler Ficheiro",
                f"Não foi possível ler os dados do ficheiro Excel.\nVerifique se o ficheiro e a folha 'Doc Proposta' estão corretos.\n\nDetalhe: {e}",
            )
            self.limpar_campos_dados()

    def selecionar_e_carregar_excel(self):
        path = filedialog.askopenfilename(
            title="Selecione o ficheiro Excel",
            filetypes=[
                ("Ficheiros Excel", "*.xlsx *.xlsm *.xls"),
                ("Todos os ficheiros", "*.*"),
            ],
        )
        self._carregar_dados_excel(path)

    def selecionar_modelos(self):
        path = filedialog.askdirectory(title="Selecione a pasta dos modelos Word")
        if path:
            self.settings["template_path"] = path
            self.log("Pasta de modelos selecionada.")

    def selecionar_saida(self):
        path = filedialog.askdirectory(title="Selecione a pasta de saída")
        if path:
            self.caminho_saida.set(path)
            self.settings["output_path"] = path

    def abrir_janela_alvara(self):
        # ### ALTERAÇÃO: Passar dados existentes para a janela de alvará ###
        app = AlvaraApp(
            self,
            sv_ttk.get_theme(),
            initial_data=self.alvara_selecoes,
            catalog=self.profile["tools"]["docs"]["catalog"] or ALVARA_DATA,
        )
        self.alvara_selecoes = app.selecoes
        self._update_alvara_display()

    def _update_alvara_display(self):
        if not self.alvara_selecoes:
            self.alvara_status.set("0 itens selecionados")
            return
        grouped_by_cat = {}
        for selecao in self.alvara_selecoes:
            cat_num_match, sub_num_match = (
                re.search(r"\d+", selecao["cat_num_part"]),
                re.search(r"\d+", selecao["sub_prefix_display"]),
            )
            if cat_num_match and sub_num_match:
                cat_num, sub_num = cat_num_match.group(), sub_num_match.group()
                if cat_num not in grouped_by_cat:
                    grouped_by_cat[cat_num] = []
                if sub_num not in grouped_by_cat[cat_num]:
                    grouped_by_cat[cat_num].append(sub_num)
        category_parts = []
        for cat_num in sorted(grouped_by_cat.keys(), key=int):
            sub_nums = sorted(grouped_by_cat[cat_num], key=int)
            if len(sub_nums) == 1:
                sub_str = f"Sub. {sub_nums[0]}"
            else:
                sub_str = f"Subcategoria {', '.join(sub_nums[:-1])} e {sub_nums[-1]}"
            category_parts.append(f"Categoria {cat_num}; {sub_str}")

        display_text = ". ".join(category_parts)
        self.alvara_status.set(f"{len(self.alvara_selecoes)} itens: {display_text}")

    def update_progress(self, current, total):
        self.progress_bar["maximum"] = total
        self.progress_bar["value"] = current
        self.update_idletasks()

    def processar_e_gerar(self):
        self.log_area.config(state="normal")
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state="disabled")
        template_path = self.settings.get("template_path", "")
        if not template_path or not os.path.isdir(template_path):
            template_path = filedialog.askdirectory(
                title="Selecione a pasta dos modelos Word"
            )
            if not template_path:
                return
            self.settings["template_path"] = template_path
        manual = Mapper(self.profile, "")
        for name, mapping in self.profile["proposal"].items():
            if mapping["kind"] in ("constant", "manual"):
                self.extra_context[name] = manual.field(mapping)
        contexto = {
            **self.extra_context,
            "ID_INTERNO": self.id_interno_var.get(),
            "NOME_DA_EMPREITADA": self.nome_empreitada_var.get(),
            "LOCAL_DA_OBRA": self.local_obra_var.get(),
            "DURACAO_DA_EMPREITADA": self.duracao_empreitada_var.get(),
            "DOCUMENTOS_DA_PROPOSTA": self.docs_text.get("1.0", tk.END).strip(),
            "DATA_HOJE": datetime.now().strftime(
                self.profile["tools"]["docs"]["date_format"]
            ),
            "ALVARA_SELECOES": self.alvara_selecoes,
        }
        self.log("Dados recolhidos da interface.")
        gerar_documentos(
            contexto,
            self.caminho_saida.get(),
            template_path,
            self.log,
            self.update_progress,
            options=self.profile["tools"]["docs"],
        )

    def open_settings(self):
        self.profile["tools"]["docs"].update(
            templates=self.settings.get("template_path", ""),
            output=self.settings.get("output_path", ""),
        )
        self.profile_store.save(self.profile)
        SettingsWindow(self, self.profile_store, "docs", self.apply_profile)

    def apply_profile(self, profile):
        self.profile = profile
        options = profile["tools"]["docs"]
        self.settings.update(
            template_path=options["templates"], output_path=options["output"]
        )
        self.caminho_saida.set(options["output"] or "Nenhuma pasta selecionada")
        sv_ttk.set_theme(profile["preferences"]["theme"])
        self.extra_context = {}
        self.limpar_campos_dados()
        self.log("Perfil aplicado. Carregue novamente o Excel.")

    def on_closing(self):
        self.profile["preferences"]["theme"] = self.settings.get("theme", "light")
        self.profile["tools"]["docs"].update(
            templates=self.settings.get("template_path", ""),
            output=self.settings.get("output_path", ""),
        )
        self.profile_store.save(self.profile)
        config_manager.save_config(self.settings)
        self.destroy()
