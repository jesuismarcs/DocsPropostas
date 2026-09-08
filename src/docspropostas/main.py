import tkinter as tk
from tkinter import messagebox
import sys
import locale
from .gui import MainApp
from .utils import verificar_chave_acesso


def main():
    """Função principal para configurar e correr a aplicação."""
    if verificar_chave_acesso():
        try:
            # Tenta definir a localização para Português para formatar a data corretamente
            locale.setlocale(locale.LC_ALL, "pt_PT.UTF-8")
        except locale.Error:
            try:
                # Tenta uma alternativa comum no Windows
                locale.setlocale(locale.LC_ALL, "Portuguese_Portugal.1252")
            except locale.Error:
                # Se falhar, o programa continua com a localização padrão
                pass

        app = MainApp()
        app.mainloop()
    else:
        # Esconde a janela principal do Tkinter se a autenticação falhar
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Acesso Negado", "Acesso negado. Contacte o administrador."
        )
        sys.exit()


if __name__ == "__main__":
    main()
