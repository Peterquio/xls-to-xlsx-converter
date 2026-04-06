import os
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    win32 = None
    pythoncom = None


APP_TITLE = "Conversor XLS para XLSX"
WINDOW_SIZE = "840x520"


class ConversorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.minsize(700, 420)
        self.root.configure(bg="#0f172a")

        self.arquivo_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Selecione um arquivo .xls para começar.")

        self._configurar_estilo()
        self._montar_interface()

    def _configurar_estilo(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Card.TFrame",
            background="#111827",
            relief="flat"
        )
        style.configure(
            "Title.TLabel",
            background="#0f172a",
            foreground="#f8fafc",
            font=("Segoe UI", 22, "bold")
        )
        style.configure(
            "Subtitle.TLabel",
            background="#0f172a",
            foreground="#cbd5e1",
            font=("Segoe UI", 10)
        )
        style.configure(
            "Label.TLabel",
            background="#111827",
            foreground="#e5e7eb",
            font=("Segoe UI", 10, "bold")
        )
        style.configure(
            "Info.TLabel",
            background="#111827",
            foreground="#cbd5e1",
            font=("Segoe UI", 10)
        )
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=10,
            background="#2563eb",
            foreground="#ffffff",
            borderwidth=0
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#1d4ed8"), ("disabled", "#334155")],
            foreground=[("disabled", "#cbd5e1")]
        )
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=10,
            background="#334155",
            foreground="#ffffff",
            borderwidth=0
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#475569")]
        )
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor="#1e293b",
            background="#22c55e",
            bordercolor="#1e293b",
            lightcolor="#22c55e",
            darkcolor="#22c55e"
        )

    def _montar_interface(self):
        container = tk.Frame(self.root, bg="#0f172a")
        container.pack(fill="both", expand=True, padx=24, pady=24)

        topo = tk.Frame(container, bg="#0f172a")
        topo.pack(fill="x", pady=(0, 18))

        ttk.Label(topo, text="Conversor XLS → XLSX", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            topo,
            text="Selecione um arquivo .xls e o programa salvará a versão .xlsx na mesma pasta do arquivo original.",
            style="Subtitle.TLabel"
        ).pack(anchor="w", pady=(6, 0))

        card = ttk.Frame(container, style="Card.TFrame", padding=20)
        card.pack(fill="both", expand=True)

        ttk.Label(card, text="Arquivo de origem", style="Label.TLabel").pack(anchor="w")

        linha_arquivo = tk.Frame(card, bg="#111827")
        linha_arquivo.pack(fill="x", pady=(10, 8))

        self.entry_arquivo = tk.Entry(
            linha_arquivo,
            textvariable=self.arquivo_var,
            font=("Segoe UI", 11),
            bg="#0b1220",
            fg="#f8fafc",
            insertbackground="#f8fafc",
            relief="flat",
            bd=0
        )
        self.entry_arquivo.pack(side="left", fill="x", expand=True, ipady=12, padx=(0, 12))

        self.btn_buscar = ttk.Button(
            linha_arquivo,
            text="Selecionar arquivo",
            style="Secondary.TButton",
            command=self.selecionar_arquivo
        )
        self.btn_buscar.pack(side="right")

        ttk.Label(
            card,
            text="Aceita arquivos .xls e cria um .xlsx com o mesmo nome na mesma pasta.",
            style="Info.TLabel"
        ).pack(anchor="w", pady=(0, 16))

        bloco_prev = tk.Frame(card, bg="#0b1220", highlightbackground="#1f2937", highlightthickness=1)
        bloco_prev.pack(fill="x", pady=(0, 16))

        self.label_destino = tk.Label(
            bloco_prev,
            text="Destino: aguardando seleção de arquivo...",
            font=("Segoe UI", 10),
            bg="#0b1220",
            fg="#93c5fd",
            anchor="w",
            justify="left",
            padx=12,
            pady=12,
            wraplength=650
        )
        self.label_destino.pack(fill="x")

        self.progress = ttk.Progressbar(
            card,
            mode="indeterminate",
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x", pady=(0, 12))

        self.label_status = tk.Label(
            card,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            bg="#111827",
            fg="#e5e7eb",
            anchor="w",
            justify="left",
            wraplength=650
        )
        self.label_status.pack(fill="x", pady=(0, 20))

        self.btn_converter = ttk.Button(
            card,
            text="Converter para XLSX",
            style="Primary.TButton",
            command=self.iniciar_conversao
        )
        self.btn_converter.pack(anchor="e")

        rodape = tk.Frame(container, bg="#0f172a")
        rodape.pack(fill="x", pady=(12, 0))

        aviso = (
            "Observação: este conversor usa o Microsoft Excel via automação do Windows. "
            "Por isso, o Excel precisa estar instalado nesta máquina."
        )
        ttk.Label(rodape, text=aviso, style="Subtitle.TLabel").pack(anchor="w")

        self.arquivo_var.trace_add("write", self.atualizar_previsao_destino)

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha XLS",
            filetypes=[("Planilhas Excel 97-2003", "*.xls"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            self.arquivo_var.set(caminho)
            self.status_var.set("Arquivo selecionado. Clique em 'Converter para XLSX'.")

    def atualizar_previsao_destino(self, *_):
        caminho = self.arquivo_var.get().strip()
        if not caminho:
            self.label_destino.config(text="Destino: aguardando seleção de arquivo...")
            return

        try:
            destino = self.gerar_caminho_destino(Path(caminho))
            self.label_destino.config(text=f"Destino: {destino}")
        except Exception:
            self.label_destino.config(text="Destino: não foi possível calcular o arquivo de saída.")

    def iniciar_conversao(self):
        if win32 is None:
            messagebox.showerror(
                "Dependência ausente",
                "O módulo 'pywin32' não está instalado.\n\n"
                "Instale com:\n"
                "pip install pywin32"
            )
            return

        caminho = self.arquivo_var.get().strip()
        if not caminho:
            messagebox.showwarning("Arquivo não selecionado", "Selecione um arquivo .xls antes de converter.")
            return

        origem = Path(caminho)

        if not origem.exists():
            messagebox.showerror("Arquivo não encontrado", "O arquivo selecionado não existe mais.")
            return

        if origem.suffix.lower() != ".xls":
            messagebox.showwarning("Formato inválido", "Selecione um arquivo com extensão .xls.")
            return

        self._set_botoes_habilitados(False)
        self.progress.start(10)
        self.status_var.set("Convertendo arquivo... aguarde.")

        thread = threading.Thread(target=self.converter_arquivo, args=(origem,), daemon=True)
        thread.start()

    def converter_arquivo(self, origem: Path):
        try:
            destino = self.gerar_caminho_destino(origem)
            self._converter_xls_para_xlsx_excel(origem, destino)
            self.root.after(0, lambda: self._finalizar_sucesso(destino))
        except Exception as e:
            self.root.after(0, lambda: self._finalizar_erro(str(e)))

    def gerar_caminho_destino(self, origem: Path) -> Path:
        pasta = origem.parent
        nome_base = origem.stem
        destino = pasta / f"{nome_base}.xlsx"

        contador = 1
        while destino.exists() and destino.resolve() != origem.resolve():
            destino = pasta / f"{nome_base}_{contador}.xlsx"
            contador += 1

        return destino

    def _converter_xls_para_xlsx_excel(self, origem: Path, destino: Path):
        if pythoncom is None:
            raise RuntimeError("pythoncom não está disponível. Instale o pacote pywin32.")

        pythoncom.CoInitialize()
        excel = None
        workbook = None

        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(str(origem.resolve()))
            formato_xlsx = 51  # xlOpenXMLWorkbook
            workbook.SaveAs(str(destino.resolve()), FileFormat=formato_xlsx)
        except Exception as e:
            raise RuntimeError(f"Não foi possível converter a planilha: {e}")
        finally:
            try:
                if workbook is not None:
                    workbook.Close(False)
            except Exception:
                pass

            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass

            pythoncom.CoUninitialize()

    def _finalizar_sucesso(self, destino: Path):
        self.progress.stop()
        self._set_botoes_habilitados(True)
        self.status_var.set("Conversão concluída com sucesso.")
        self.label_destino.config(text=f"Arquivo gerado: {destino}")
        messagebox.showinfo("Sucesso", f"Planilha convertida com sucesso.\n\nArquivo salvo em:\n{destino}")

    def _finalizar_erro(self, erro: str):
        self.progress.stop()
        self._set_botoes_habilitados(True)
        self.status_var.set("Ocorreu um erro durante a conversão.")
        messagebox.showerror("Erro na conversão", erro)

    def _set_botoes_habilitados(self, habilitado: bool):
        estado = "normal" if habilitado else "disabled"
        self.btn_buscar.config(state=estado)
        self.btn_converter.config(state=estado)


def verificar_excel_instalado() -> bool:
    if win32 is None:
        return False

    try:
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Quit()
        pythoncom.CoUninitialize()
        return True
    except Exception:
        return False


def main():
    root = tk.Tk()
    app = ConversorApp(root)

    if win32 is None:
        app.status_var.set("pywin32 não encontrado. Instale com: pip install pywin32")
    elif not verificar_excel_instalado():
        app.status_var.set("Microsoft Excel não foi detectado nesta máquina.")

    root.mainloop()


if __name__ == "__main__":
    main()
