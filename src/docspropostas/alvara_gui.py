import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import sv_ttk

from .constants import ALVARA_DATA
from .utils import resource_path
import os


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
            padding=(2, 2, 2, 2),
        )
        label.pack(ipadx=1)

    def hide_tooltip(self, event):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None


class HabilitacaoRow(ttk.Frame):
    def __init__(self, parent, app_parent, row_number, delete_callback):
        super().__init__(parent, style="Card.TFrame", padding=5)
        self.app_parent, self.row_number, self.delete_callback = (
            app_parent,
            row_number,
            delete_callback,
        )
        self.number_label = ttk.Label(
            self, text=str(row_number), font=("Arial", 10, "bold")
        )
        self.number_label.grid(row=0, column=0, padx=(0, 10), rowspan=2)
        ttk.Label(self, text="Categoria").grid(row=0, column=1, sticky="w")
        self.categoria_combo = ttk.Combobox(
            self, values=list(ALVARA_DATA.keys()), state="readonly", width=35
        )
        self.categoria_combo.grid(row=1, column=1, padx=5)
        self.categoria_combo.bind("<<ComboboxSelected>>", self.on_category_change)
        ttk.Label(self, text="Subcategoria").grid(row=0, column=2, sticky="w")
        self.subcategoria_combo = ttk.Combobox(self, state="readonly", width=45)
        self.subcategoria_combo.grid(row=1, column=2, padx=5)
        self.subcategoria_combo.bind("<<ComboboxSelected>>", self.on_subcategory_change)
        ttk.Label(self, text="Classe").grid(row=0, column=3, sticky="w")
        self.class_var = tk.StringVar(value="")
        self.class_entry = ttk.Entry(self, textvariable=self.class_var, width=7)
        self.class_entry.grid(row=1, column=3, padx=5)
        ttk.Label(self, text="Valor (€)").grid(row=0, column=4, sticky="w")
        self.valor_entry = ttk.Entry(self, width=15)
        self.valor_entry.grid(row=1, column=4, padx=5)
        self.valor_entry.insert(0, "1")
        self.delete_button = ttk.Button(
            self,
            text="🗑️",
            command=lambda: self.delete_callback(self),
            style="Toolbutton.TButton",
        )
        self.delete_button.grid(row=1, column=5, padx=5)
        ToolTip(self.delete_button, "Remover esta habilitação")

    def on_category_change(self, event=None):
        self.subcategoria_combo.set("")
        self.class_var.set("")
        self.app_parent.update_all_dropdowns()

    def on_subcategory_change(self, event=None):
        categoria, subcategoria_name = (
            self.categoria_combo.get(),
            self.subcategoria_combo.get(),
        )
        if categoria and subcategoria_name:
            sub_info = next(
                (s for s in ALVARA_DATA[categoria] if s["name"] == subcategoria_name),
                None,
            )
            if sub_info:
                self.class_var.set(sub_info["class"])
        self.app_parent.update_all_dropdowns()

    def refresh_options(self, selected_set):
        current_cat, current_sub = (
            self.categoria_combo.get(),
            self.subcategoria_combo.get(),
        )
        if not current_cat:
            self.subcategoria_combo["values"] = []
            return
        all_subcats = [sub["name"] for sub in ALVARA_DATA.get(current_cat, [])]
        available_subcats = [
            s
            for s in all_subcats
            if (current_cat, s) not in selected_set or s == current_sub
        ]
        self.subcategoria_combo["values"] = available_subcats

    def set_row_number(self, number):
        self.row_number = number
        self.number_label.config(text=str(number))

    def get_data(self):
        categoria_full, subcategoria_full, classe, valor = (
            self.categoria_combo.get(),
            self.subcategoria_combo.get(),
            self.class_var.get(),
            self.valor_entry.get().strip(),
        )
        if not all([categoria_full, subcategoria_full, classe, valor]):
            return None
        cat_num_part = categoria_full.split(" ")[0]
        sub_prefix_display = subcategoria_full.split(" ")[0]
        sub_desc_text = " ".join(subcategoria_full.split(" ")[1:]).lstrip("- ")
        return {
            "cat_num_part": cat_num_part,
            "sub_prefix_display": sub_prefix_display,
            "class_val": f"Classe {classe}",
            "sub_desc": sub_desc_text,
            "valor_executar": f"{valor} €",
            "full_cat_name": categoria_full,
            "full_sub_name": subcategoria_full,
        }


class AlvaraApp(tk.Toplevel):
    # ### ALTERAÇÃO: Adicionado `initial_data` ao construtor ###
    def __init__(self, parent, theme, initial_data=None):
        super().__init__(parent)
        sv_ttk.set_theme(theme)
        self.title("Habilitações de Alvará")
        self.geometry("1000x650")
        self.selecoes, self.rows, self.icons, self.selected_habilitacoes = (
            list(initial_data or []),
            [],
            self._load_icons(),
            set(),
        )
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.update_all_dropdowns())
        self.fix_category_var, self.total_var = (
            tk.BooleanVar(value=False),
            tk.StringVar(value="Total de Habilitações: 0"),
        )

        header_frame = ttk.Frame(self, padding=(10, 10))
        header_frame.pack(fill=tk.X)
        search_entry = ttk.Entry(header_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side=tk.LEFT, padx=(0, 10))
        ToolTip(search_entry, "Escreva aqui para pesquisar nas subcategorias")
        fix_category_check = ttk.Checkbutton(
            header_frame,
            text="Fixar Categoria",
            variable=self.fix_category_var,
            style="Switch.TCheckbutton",
        )
        fix_category_check.pack(side=tk.LEFT, padx=10)
        ToolTip(
            fix_category_check,
            "Se ativo, as novas linhas mantêm a categoria da linha anterior",
        )
        add_button = ttk.Button(
            header_frame,
            text="Adicionar Habilitação",
            image=self.icons.get("add"),
            compound="left",
            command=self.add_row,
        )
        add_button.pack(side=tk.RIGHT)
        ToolTip(add_button, "Adiciona uma nova linha para preenchimento")

        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=5)

        main_frame = ttk.Frame(self)
        main_frame.pack(fill="both", expand=True, padx=10)
        self.canvas = tk.Canvas(main_frame, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(
            main_frame, orient="vertical", command=self.canvas.yview
        )
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=5)

        footer = ttk.Frame(self, padding=(10, 10))
        footer.pack(fill=tk.X, side="bottom")
        ttk.Label(footer, textvariable=self.total_var).pack(side=tk.LEFT)
        confirm_button = ttk.Button(
            footer,
            text="Confirmar e Fechar",
            style="Accent.TButton",
            image=self.icons.get("confirm"),
            compound="left",
            command=self.confirmar,
        )
        confirm_button.pack(side=tk.RIGHT)
        ToolTip(confirm_button, "Guarda as seleções e fecha esta janela")

        # ### ALTERAÇÃO: Preencher com dados iniciais ###
        if initial_data:
            self._populate_from_initial_data(initial_data)

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.wait_window()

    # ### INÍCIO DA NOVA FUNCIONALIDADE ###
    def _populate_from_initial_data(self, initial_data):
        for data in initial_data:
            self.add_row(pre_fill_data=data)

        self.update_all_dropdowns()
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0.0)  # Ir para o topo depois de popular

    # ### FIM DA NOVA FUNCIONALIDADE ###

    def _load_icons(self):
        if not os.path.isfile(resource_path("assets/add.png")):
            return {}
        try:
            return {
                "add": ImageTk.PhotoImage(
                    Image.open(resource_path("assets/add.png")).resize(
                        (16, 16), Image.Resampling.LANCZOS
                    )
                ),
                "confirm": ImageTk.PhotoImage(
                    Image.open(resource_path("assets/confirm.png")).resize(
                        (16, 16), Image.Resampling.LANCZOS
                    )
                ),
            }
        except Exception as e:
            primeiro_caminho_tentado = resource_path("assets/add.png")
            messagebox.showerror(
                "Erro de Diagnóstico ao Carregar Ícones",
                f"A janela de Alvarás não conseguiu carregar os ícones.\n\nO caminho completo que foi tentado é:\n'{primeiro_caminho_tentado}'\n\nO erro específico foi:\n{type(e).__name__}: {e}",
            )
            return {}

    def update_all_dropdowns(self):
        self.selected_habilitacoes.clear()
        for row in self.rows:
            cat, sub = row.categoria_combo.get(), row.subcategoria_combo.get()
            if cat and sub:
                self.selected_habilitacoes.add((cat, sub))
        query = self.search_var.get().lower()
        for row in self.rows:
            row.refresh_options(self.selected_habilitacoes)
            if query and row.categoria_combo.get():
                current_values = list(row.subcategoria_combo["values"])
                filtered_values = [v for v in current_values if query in v.lower()]
                row.subcategoria_combo["values"] = filtered_values

    def add_row(self, pre_fill_data=None):
        row_number = len(self.rows) + 1
        row = HabilitacaoRow(self.scrollable_frame, self, row_number, self.remove_row)
        self.rows.append(row)

        if self.fix_category_var.get() and len(self.rows) > 1:
            last_category = self.rows[-2].categoria_combo.get()
            if last_category:
                row.categoria_combo.set(last_category)

        # ### ALTERAÇÃO: Lógica para preencher a linha com dados ###
        if pre_fill_data:
            row.categoria_combo.set(pre_fill_data.get("full_cat_name", ""))
            row.on_category_change()  # Simula evento para popular subcategorias
            row.subcategoria_combo.set(pre_fill_data.get("full_sub_name", ""))
            row.on_subcategory_change()
            row.class_var.set(
                pre_fill_data.get("class_val", "")
                .replace("Classe ", "")
                .replace("A preencher", "")
            )
            row.valor_entry.delete(0, tk.END)
            valor = pre_fill_data.get("valor_executar", "1 €").replace(" €", "")
            row.valor_entry.insert(0, valor)

        row.pack(pady=5, padx=5, fill=tk.X)
        self.update_total_count()
        self.update_all_dropdowns()
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(1.0)

    def remove_row(self, row_to_remove):
        row_to_remove.destroy()
        self.rows.remove(row_to_remove)
        self.renumber_rows()
        self.update_total_count()
        self.update_all_dropdowns()

    def renumber_rows(self):
        for i, row in enumerate(self.rows, 1):
            row.set_row_number(i)

    def update_total_count(self):
        self.total_var.set(f"Total de Habilitações: {len(self.rows)}")

    def confirmar(self):
        self.selecoes = []
        for i, row in enumerate(self.rows, 1):
            data = row.get_data()
            if data is None:
                # Apenas mostra erro se a linha não estiver completamente vazia
                if (
                    row.categoria_combo.get()
                    or row.subcategoria_combo.get()
                    or row.valor_entry.get().strip() != "1"
                ):
                    messagebox.showerror(
                        "Erro", f"A linha {i} está incompleta...", parent=self
                    )
                    return
            else:
                self.selecoes.append(data)

        if not self.selecoes and len(self.rows) > 0:
            if not messagebox.askyesno(
                "Confirmação",
                "A lista final de habilitações está vazia. Deseja continuar e limpar todas as habilitações?",
                parent=self,
            ):
                return

        self.destroy()

    def on_closing(self):
        self.destroy()
