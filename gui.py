"""
Interface gráfica principal do EmailBatchPro
"""

import os
import sys
import tkinter as tk
import webbrowser
import pandas as pd

from tkinter import ttk
from tkinter import messagebox
from tkinter import simpledialog
from tkinter import filedialog

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap import Style

from config_manager import ConfigManager
from excel_manager import ExcelManager
from outlook_manager import OutlookManager
from logger_manager import LoggerManager
from template_manager import TemplateManager
from utils import column_name_to_number
from utils import current_timestamp


APP_VERSION = "1.1.0"
APP_NAME = "EmailBatchPro"

def resource_path(relative_path):

    try:

        base_path = sys._MEIPASS

    except Exception:

        base_path = os.path.abspath(".")

    return os.path.join(
        base_path,
        relative_path
    )

class EmailBatchGUI:

    def validate_excel(self):

        try:

            self.excel.load()

            result = (
                self.excel.validate_data(
                    self.outlook.is_valid_email
                )
            )

            report_file = None

            if result["validation_records"]:

                report_df = pd.DataFrame(
                    result["validation_records"]
                )

                report_folder = os.path.dirname(
                    self.excel.excel_file
                )

                report_file = os.path.join(
                    report_folder,
                    f"validacao_{current_timestamp()}.xlsx"
                )

                report_df.to_excel(
                    report_file,
                    index=False
                )

            msg = (
                "Validação concluída\n\n"

                f"Registos: {result['rows']}\n\n"

                f"✓ Emails válidos: "
                f"{result['valid_emails']}\n"

                f"⚠ Emails vazios: "
                f"{len(result['empty_email_rows'])}\n"

                f"✗ Emails inválidos: "
                f"{len(result['invalid_email_rows'])}\n"

                f"⚠ Anexos não encontrados: "
                f"{len(result['missing_attachment_rows'])}"
            )

            if result["empty_email_rows"]:

                msg += "\n\nEmails vazios:\n"

                msg += "\n".join(
                    result["empty_email_rows"]
                )

            if result["invalid_email_rows"]:

                msg += "\n\nEmails inválidos:\n"

                msg += "\n".join(
                    result["invalid_email_rows"]
                )

            if result["missing_attachment_rows"]:

                msg += (
                    "\n\nAnexos não encontrados:\n"
                )

                msg += "\n".join(
                    result[
                        "missing_attachment_rows"
                    ]
                )

            if report_file:

                msg += (
                    "\n\nRelatório guardado em:\n"
                    f"{report_file}"
                )

            result_win = tk.Toplevel(
                self.root
            )

            self.set_window_icon(
                result_win
            )

            try:

                self.set_window_icon(
                    result_win
                )

            except Exception:

                pass

            result_win.title(
                "Validação Excel"
            )

            result_win.geometry(
                "750x550"
            )

            result_win.transient(
                self.root
            )

            result_win.grab_set()

            txt = tk.Text(
                result_win,
                wrap="word"
            )

            txt.pack(
                fill="both",
                expand=True,
                padx=10,
                pady=10
            )

            txt.insert(
                "1.0",
                msg
            )

            txt.config(
                state="disabled"
            )

            buttons = ttk.Frame(
                result_win
            )

            buttons.pack(
                fill="x",
                pady=10
            )

            buttons = ttk.Frame(
                result_win
            )

            buttons.pack(
                pady=10
            )

            if report_file:

                ttk.Button(
                    buttons,
                    text="Abrir relatório",
                    command=lambda:
                    (
                        os.startfile(report_file),
                        result_win.destroy()
                    )
                ).pack(
                    side="left",
                    padx=5
                )

            ttk.Button(
                buttons,
                text="Fechar",
                command=result_win.destroy
            ).pack(
                side="left",
                padx=5
            )

        except Exception as e:

            messagebox.showerror(
                "Validação Excel",
                str(e)
            )

    def open_paypal(self):
        webbrowser.open(
            "https://paypal.me/mmartins73"
        )

    def __init__(self):

        self.config = ConfigManager()
        self.excel = ExcelManager()
        self.outlook = OutlookManager()
        self.logger = LoggerManager()
        self.templates = TemplateManager()

        self.root = tb.Window(
            themename="litera"
        )

        self.style = Style()

        self.style.configure(
            "Success.Horizontal.TProgressbar",
            background="#198754"
        )

        try:

            self.root.iconbitmap(
                resource_path(
                    "EmailBatchPro.ico"
                )
            )

        except Exception as e:

            print(
                "Erro:",
                e
            )

        self.root.title(
            f"{APP_NAME} v{APP_VERSION}"
        )

        self.root.geometry("1100x850")

        self.create_variables()

        self.last_text_widget = None

        self.create_widgets()

        self.load_configuration()

        self.load_placeholders()

        self.font_name_var.set("Aptos")

    def select_rows(self):
        try:
            self.excel.load()
        except Exception as e:
            messagebox.showerror(
                "Selecionar registos",
                str(e)
            )
            return

        # Identifica o ficheiro e o número de registos. Se mudar,
        # a janela começa novamente com todos selecionados.
        selection_source = (
            self.excel_file_var.get(),
            len(self.excel.df)
        )

        previous_source = getattr(
            self,
            "_row_selection_source",
            None
        )

        if previous_source == selection_source:
            selected = set(self.selected_rows)
        else:
            selected = set(range(len(self.excel.df)))

        win = tk.Toplevel(self.root)
        self.set_window_icon(win)
        win.title("Selecionar registos")
        win.geometry("700x600")
        win.transient(self.root)
        win.grab_set()

        top_frame = ttk.Frame(win)
        top_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(top_frame)

        scrollbar = ttk.Scrollbar(
            top_frame,
            orient="vertical",
            command=canvas.yview
        )

        scroll_frame = ttk.Frame(canvas)

        scroll_frame.bind(
            "<Configure>",
            lambda event: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window(
            (0, 0),
            window=scroll_frame,
            anchor="nw"
        )

        canvas.configure(
            yscrollcommand=scrollbar.set
        )

        canvas.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        check_vars = []

        ttk.Label(
            scroll_frame,
            text="Selecione os registos a processar",
            font=("Segoe UI", 10, "bold")
        ).pack(
            anchor="w",
            padx=10,
            pady=10
        )

        cols = list(self.excel.df.columns[:3])

        # A posição é a mesma usada em self.selected_rows.
        for position, (_, row) in enumerate(
            self.excel.df.iterrows()
        ):
            var = tk.BooleanVar(
                master=win,
                value=(position in selected)
            )
            check_vars.append(var)

            parts = [
                f"{col}: {row[col]}"
                for col in cols
            ]

            text = (
                f"Linha {position + 1} - "
                + " | ".join(parts)
            )

            ttk.Checkbutton(
                scroll_frame,
                text=text,
                variable=var
            ).pack(
                anchor="w",
                padx=20,
                pady=1
            )

        buttons_frame = ttk.Frame(win)
        buttons_frame.pack(
            fill="x",
            padx=10,
            pady=10
        )

        def select_all():
            for var in check_vars:
                var.set(True)

        def clear_all():
            for var in check_vars:
                var.set(False)

        def save_selection():
            self.selected_rows = [
                position
                for position, var in enumerate(check_vars)
                if var.get()
            ]

            self._row_selection_source = selection_source
            win.destroy()

        ttk.Button(
            buttons_frame,
            text="Selecionar todos",
            command=select_all
        ).pack(side="left", padx=5)

        ttk.Button(
            buttons_frame,
            text="Limpar todos",
            command=clear_all
        ).pack(side="left", padx=5)

        ttk.Button(
            buttons_frame,
            text="Cancelar",
            command=win.destroy
        ).pack(side="right", padx=5)

        ttk.Button(
            buttons_frame,
            text="OK",
            command=save_selection
        ).pack(side="right", padx=5)

        # Fechar pelo X tem o mesmo efeito que premir OK.
        win.protocol("WM_DELETE_WINDOW", save_selection)

    def set_window_icon(
        self,
        window
    ):

        try:

            window.iconbitmap(
                resource_path(
                    "EmailBatchPro.ico"
                )
            )

        except Exception:

            pass

    def show_about(self):

        about = tk.Toplevel(self.root)

        try:

            self.set_window_icon(
                about
            )

        except Exception:

            pass

        about.transient(self.root)
        about.grab_set()
        about.focus_set()
        about.title("Sobre o EmailBatchPro")
        about.geometry("550x550")
        about.resizable(True, True)

        about.minsize(
            550,
            550
        )

        ttk.Label(
            about,
            text="EmailBatchPro",
            font=("Segoe UI", 12, "bold")
        ).pack(pady=(10,5))

        ttk.Label(
            about,
            text=f"Versão {APP_VERSION}"
        ).pack()

        texto = """
Aplicação para envio em lote de emails
personalizados através do Microsoft Outlook.

Desenvolvido por:
Miguel Martins

Funcionalidades principais:
• Placeholders dinâmicos
• Anexos automáticos
• Assinatura Outlook
• Modelos reutilizáveis
• Logs de processamento
• Integração com Outlook

Este software é disponibilizado gratuitamente.

Se o EmailBatchPro lhe poupar tempo e quiser
apoiar o desenvolvimento de futuras versões,
considere oferecer um café ao programador ☕
    """

        ttk.Label(
            about,
            text=texto,
            justify="left",
            anchor="w"
        ).pack(
            fill="both",
            expand=True,
            padx=20,
            pady=10
        )

        buttons_frame = ttk.Frame(about)

        buttons_frame.pack(
            side="bottom",
            fill="x",
            pady=10
        )

        ttk.Button(
            buttons_frame,
            text="☕ Oferecer um café",
            command=self.open_paypal
        ).pack(
            side="left",
            padx=10
        )

        ttk.Button(
            buttons_frame,
            text="Fechar",
            command=about.destroy
        ).pack(
            side="right",
            padx=10
        )

    def show_help(self):

        win = tk.Toplevel(self.root)
        try:

            self.set_window_icon(
                win
            )

        except Exception:

            pass
        win.title("Ajuda")

        win.geometry("700x500")

        win.transient(self.root)

        win.grab_set()

        txt = tk.Text(
            win,
            wrap="word"
        )

        txt.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        txt.insert(
            "1.0",
            """
EMAILBATCHPRO

1. Selecione o ficheiro Excel.

2. O ficheiro deve conter uma coluna obrigatória chamada EMAIL.

3. Pode existir uma coluna opcional chamada ANEXOS.

4. Os restantes cabeçalhos do Excel ficam disponíveis como placeholders.

Exemplo:

NOME | EMAIL | PROCESSO | ANEXOS

Permite utilizar:

{NOME}
{EMAIL}
{PROCESSO}

5. Utilize os botões dos campos para inserir placeholders no assunto ou no corpo da mensagem.

6. Pré-visualizar
Mostra o resultado da primeira linha do Excel sem criar emails.

7. Testar 1ª linha
Abre no Outlook um email de teste utilizando a primeira linha do Excel.

8. Criar todas
Cria ou envia um email por cada linha do ficheiro Excel.

9. Incluir assinatura Outlook
Acrescenta a assinatura atualmente configurada no Outlook.

10. Enviar diretamente
Quando ativado, os emails são enviados automaticamente.

Quando desativado, os emails são apenas criados e abertos para revisão antes do envio.

Placeholders do sistema:

{DATA}
{HORA}
{LINHA}

A coluna EMAIL é obrigatória.
A coluna ANEXOS é opcional.

Recomendação:

Utilize sempre "Pré-visualizar" ou "Testar 1ª linha" antes de criar ou enviar todos os emails.
    """
        )

        txt.config(
            state="disabled"
        )

        ttk.Button(
            win,
            text="Fechar",
            command=win.destroy
        ).pack(
            pady=10
        )

    def load_placeholders(self):

        try:

            if not self.excel_file_var.get():

                return

            self.excel.load()
            
            self.excel.validate_required_columns()

            for widget in (
                self.placeholders_frame
                .winfo_children()
            ):

                widget.destroy()

            ttk.Label(
                self.placeholders_frame,
                text="Campos:"
            ).grid(
                row=0,
                column=0,
                sticky="ew",
                padx=(0, 10)
            )

            row_num = 0
            col_num = 1

            MAX_COLS = 5

            for name in (
                self.excel.get_placeholders()
            ):

                if name.upper() in [
                    "EMAIL",
                    "ANEXOS"
                ]:
                    continue

                ttk.Button(
                    self.placeholders_frame,
                    text=name,
                    command=lambda n=name:
                    self.insert_placeholder(n)
                ).grid(
                    row=row_num,
                    column=col_num,
                    padx=2,
                    pady=2,
                    sticky="ew"
                )

                col_num += 1

                if col_num > MAX_COLS:

                    row_num += 1
                    col_num = 0

            ttk.Label(
                self.placeholders_frame,
                text="Sistema:"
            ).grid(
                row=row_num + 1,
                column=0,
                sticky="ew",
                padx=(0, 10),
                pady=(5, 0)
            )

            col_system = 1

            for name in [
                "DATA",
                "HORA",
                "LINHA"
            ]:

                ttk.Button(
                    self.placeholders_frame,
                    text=name,
                    command=lambda n=name:
                    self.insert_placeholder(n)
                ).grid(
                    row=row_num + 1,
                    column=col_system,
                    padx=2,
                    pady=(5, 0),
                    sticky="ew"
                )

                col_system += 1

        except Exception as e:

            for widget in (
                self.placeholders_frame
                .winfo_children()
            ):

                widget.destroy()

            messagebox.showerror(
                "Erro no ficheiro Excel",
                str(e)
            )

    def create_variables(self):

        self.selection_info_var = tk.StringVar(
            value="Todos os registos"
        )

        self.selected_rows = []

        self.account_var = tk.StringVar()

        self.cc_var = tk.StringVar()

        self.bcc_var = tk.StringVar()

        self.subject_var = tk.StringVar()

        self.excel_file_var = tk.StringVar()

        self.include_signature_var = tk.BooleanVar(
            value=True
        )

        self.send_directly_var = tk.BooleanVar(
            value=False
        )

        self.font_name_var = tk.StringVar(
        value="Aptos"
        )

        self.font_size_var = tk.StringVar(
        value="10"
        )

    def create_widgets(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="Conta Outlook").grid(
            row=0, column=0, sticky="ew"
        )

        self.account_combo = ttk.Combobox(
            main,
            textvariable=self.account_var,
            width=60
        )

        try:
            self.account_combo["values"] = self.outlook.get_accounts()
        except Exception:
            pass

        self.account_combo.grid(
            row=0, column=1, sticky="ew", pady=3
        )

        ttk.Label(main, text="Ficheiro Excel").grid(
            row=1, column=0, sticky="ew"
        )

        ttk.Entry(
            main,
            textvariable=self.excel_file_var
        ).grid(
            row=1, column=1, sticky="ew", pady=3
        )

        ttk.Label(main, text="CC").grid(
            row=2, column=0, sticky="ew"
        )

        ttk.Entry(
            main,
            textvariable=self.cc_var
        ).grid(
            row=2, column=1, sticky="ew", pady=3
        )

        ttk.Label(main, text="BCC").grid(
            row=3, column=0, sticky="ew"
        )

        ttk.Entry(
            main,
            textvariable=self.bcc_var
        ).grid(
            row=3, column=1, sticky="ew", pady=3
        )

        ttk.Label(main, text="Assunto").grid(
            row=4, column=0, sticky="ew"
        )

        self.subject_entry = ttk.Entry(
            main,
            textvariable=self.subject_var
        )
        self.subject_entry.grid(
            row=4, column=1, sticky="ew", pady=3
        )
        self.subject_entry.bind(
            "<FocusIn>",
            self.remember_widget
        )

        ttk.Label(main, text="Tipo de letra").grid(
            row=5, column=0, sticky="ew"
        )

        font_frame = ttk.Frame(main)
        font_frame.grid(
            row=5, column=1, sticky="ew", pady=3
        )

        self.font_combo = ttk.Combobox(
            font_frame,
            textvariable=self.font_name_var,
            width=20,
            values=[
                "Aptos",
                "Calibri",
                "Arial",
                "Segoe UI",
                "Times New Roman"
            ]
        )
        self.font_combo.pack(side="left")

        self.font_size_combo = ttk.Combobox(
            font_frame,
            textvariable=self.font_size_var,
            width=5,
            values=["10", "11", "12", "13", "14", "16"]
        )
        self.font_size_combo.pack(
            side="left", padx=(5, 0)
        )

        self.placeholders_frame = ttk.LabelFrame(
            main,
            text="Campos disponíveis",
            padding=5
        )
        self.placeholders_frame.grid(
            row=6, column=1, sticky="w", pady=3
        )

        text_frame = tk.LabelFrame(
            main,
            text="Corpo da mensagem",
            padx=5,
            pady=5,
            relief="flat",
            bd=1
        )
        text_frame.grid(
            row=7, column=1, sticky="nsew", pady=5
        )

        self.body_text = tk.Text(
            text_frame,
            wrap="word",
            height=20,
            bd=0
        )
        self.body_text.bind(
            "<FocusIn>",
            self.remember_widget
        )
        self.body_text.pack(
            fill="both", expand=True
        )

        ttk.Checkbutton(
            main,
            text="Incluir assinatura Outlook",
            variable=self.include_signature_var
        ).grid(
            row=8, column=1, sticky="ew"
        )

        ttk.Checkbutton(
            main,
            text="Enviar diretamente",
            variable=self.send_directly_var
        ).grid(
            row=9, column=1, sticky="ew"
        )

        self.progress = ttk.Progressbar(
            main,
            orient="horizontal",
            style="Success.Horizontal.TProgressbar"
        )
        self.progress.grid(
            row=10,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=10
        )

        # Quatro grupos de botões, em duas colunas.
        buttons = ttk.Frame(main)
        buttons.grid(
            row=11,
            column=0,
            columnspan=2,
            sticky="ew"
        )
        buttons.columnconfigure(0, weight=1, uniform="button_groups")
        buttons.columnconfigure(1, weight=1, uniform="button_groups")

        excel_group = ttk.LabelFrame(
            buttons, text="Excel", padding=8
        )
        excel_group.grid(
            row=0, column=0, sticky="nsew", padx=5, pady=5
        )
        excel_group.columnconfigure(0, weight=1)
        excel_group.columnconfigure(1, weight=1)

        ttk.Button(
            excel_group,
            text="Selecionar Excel",
            command=self.select_excel_file
        ).grid(
            row=0, column=0, sticky="ew", padx=3, pady=3
        )

        ttk.Button(
            excel_group,
            text="Atualizar campos",
            command=self.refresh_placeholders
        ).grid(
            row=0, column=1, sticky="ew", padx=3, pady=3
        )

        ttk.Button(
            excel_group,
            text="Validar Excel",
            command=self.validate_excel
        ).grid(
            row=1, column=0, sticky="ew", padx=3, pady=3
        )

        ttk.Button(
            excel_group,
            text="Selecionar registos",
            command=self.select_rows
        ).grid(
            row=1, column=1, sticky="ew", padx=3, pady=3
        )

        tests_group = ttk.LabelFrame(
            buttons, text="Testes", padding=8
        )
        tests_group.grid(
            row=0, column=1, sticky="nsew", padx=5, pady=5
        )
        tests_group.columnconfigure(0, weight=1)
        tests_group.columnconfigure(1, weight=1)

        ttk.Button(
            tests_group,
            text="Pré-visualizar",
            command=self.preview_first
        ).grid(
            row=0, column=0, sticky="ew", padx=3, pady=3
        )

        ttk.Button(
            tests_group,
            text="Testar 1ª linha",
            command=self.test_first_row
        ).grid(
            row=0, column=1, sticky="ew", padx=3, pady=3
        )

        production_group = ttk.LabelFrame(
            buttons, text="Produção", padding=8
        )
        production_group.grid(
            row=1, column=0, sticky="nsew", padx=5, pady=5
        )

        ttk.Button(
            production_group,
            text="Criar todas",
            command=self.create_all_messages
        ).grid(
            row=0, column=0, sticky="w", padx=3, pady=3
        )

        tools_group = ttk.LabelFrame(
            buttons, text="Ferramentas", padding=8
        )
        tools_group.grid(
            row=1, column=1, sticky="nsew", padx=5, pady=5
        )
        tools_group.columnconfigure(0, weight=1)
        tools_group.columnconfigure(1, weight=1)

        ttk.Button(
            tools_group,
            text="Guardar modelo",
            command=self.save_template
        ).grid(
            row=0, column=0, sticky="ew", padx=3, pady=3
        )

        ttk.Button(
            tools_group,
            text="Carregar modelo",
            command=self.load_template
        ).grid(
            row=0, column=1, sticky="ew", padx=3, pady=3
        )

        ttk.Button(
            tools_group,
            text="Ajuda",
            command=self.show_help
        ).grid(
            row=1, column=0, sticky="ew", padx=3, pady=3
        )

        ttk.Button(
            tools_group,
            text="Sobre",
            command=self.show_about
        ).grid(
            row=1, column=1, sticky="ew", padx=3, pady=3
        )

        main.rowconfigure(7, weight=1)
        main.columnconfigure(1, weight=1)


    def select_excel_file(self):
        filename = filedialog.askopenfilename(
            title="Selecionar ficheiro Excel",
            filetypes=[
                ("Ficheiros Excel", "*.xlsx *.xls")
            ]
        )

        if filename:
            self.excel_file_var.set(filename)
            self.save_configuration()
            self.excel.excel_file = filename
            self.load_placeholders()

    def remember_widget(
        self,
        event
    ):

        self.last_text_widget = event.widget

    def insert_placeholder(
        self,
        placeholder
    ):

        text = f"{{{placeholder}}}"

        widget = self.last_text_widget

        try:

            if isinstance(widget, tk.Text):

                widget.insert(
                    tk.INSERT,
                    text
                )

                widget.focus_set()

            elif widget is self.subject_entry:

                pos = widget.index(
                    tk.INSERT
                )

                widget.insert(
                    pos,
                    text
                )

                widget.focus_set()

            else:

                self.body_text.insert(
                    tk.INSERT,
                    text
                )

                self.body_text.focus_set()

        except Exception:

            self.body_text.insert(
                tk.INSERT,
                text
            )

            self.body_text.focus_set()

    def load_configuration(self):

        cfg = self.config.load()

        self.account_var.set(
            cfg.get("account", "")
        )

        self.cc_var.set(
            cfg.get("cc", "")
        )

        self.bcc_var.set(
            cfg.get("bcc", "")
        )

        self.subject_var.set(
            cfg.get("subject", "")
        )

        self.include_signature_var.set(
            cfg.get(
                "include_signature",
                True
            )
        )

        self.send_directly_var.set(
            cfg.get(
                "send_directly",
                False
            )
        )

        self.body_text.delete(
            "1.0",
            tk.END
        )

        self.body_text.insert(
            "1.0",
            cfg.get(
                "body",
                ""
            )
        )

        self.font_name_var.set(
            cfg.get(
                "font_name",
                "Aptos"
            )
        )

        self.font_size_var.set(
            cfg.get(
                "font_size",
                "10"
            )
        )

        self.excel_file_var.set(
            cfg.get(
                "excel_file",
                ""
            )
        )

        if self.excel_file_var.get():

            self.excel.excel_file = (
                self.excel_file_var.get()
            )

    def save_configuration(self):

        cfg = {

            "account":
                self.account_var.get(),

            "cc":
                self.cc_var.get(),

            "bcc":
                self.bcc_var.get(),

            "subject":
                self.subject_var.get(),

            "body":
                self.body_text.get(
                    "1.0",
                    tk.END
                ),

            "include_signature":
                self.include_signature_var.get(),

            "send_directly":
                self.send_directly_var.get(),

            "font_name":
                self.font_name_var.get(),

            "font_size":
                self.font_size_var.get(),

            "excel_file":
                self.excel_file_var.get(),
        }

        self.config.save(
            cfg
        )

    def preview_first(self):

        try:

            self.excel.load()

            preview = self.excel.build_preview(
                self.subject_var.get(),
                self.body_text.get(
                    "1.0",
                    tk.END
                ),
            )

            text = (
                f"Destinatário: {preview['to']}\n\n"
                f"Assunto:\n{preview['subject']}\n\n"
                f"Corpo:\n{preview['body']}"
            )

            messagebox.showinfo(
                "Pré-visualização",
                text
            )

        except Exception as e:

            messagebox.showerror(
                "Erro",
                str(e)
            )

    def save_template(self):

        try:

            name = simpledialog.askstring(
                "Guardar modelo",
                "Nome do modelo:"
            )

            if not name:
                return

            self.templates.save_template(
                name,
                self.subject_var.get(),
                self.body_text.get(
                    "1.0",
                    tk.END
                )
            )

            messagebox.showinfo(
                "Sucesso",
                "Modelo guardado."
            )

        except Exception as e:

            messagebox.showerror(
                "Erro",
                str(e)
            )

    def load_template(self):

        try:

            templates = (
                self.templates.list_templates()
            )

            if not templates:

                messagebox.showwarning(
                    "Modelos",
                    "Não existem modelos."
                )

                return

            name = simpledialog.askstring(
                "Carregar modelo",
                "Nome do modelo:\n\n"
                + "\n".join(templates)
            )

            if not name:
                return

            data = (
                self.templates.load_template(
                    name
                )
            )

            self.subject_var.set(
                data.get(
                    "subject",
                    ""
                )
            )

            self.body_text.delete(
                "1.0",
                tk.END
            )

            self.body_text.insert(
                "1.0",
                data.get(
                    "body",
                    ""
                )
            )

        except Exception as e:

            messagebox.showerror(
                "Erro",
                str(e)
            )

    def test_first_row(self):

        try:

            self.save_configuration()

            self.excel.load()

            row = self.excel.get_row(0)

            recipient = str(
                row["EMAIL"]
            ).strip()
            
            to_email = str(
                row["EMAIL"]
            ).strip()

            subject = (
                self.excel.replace_placeholders(
                    self.subject_var.get(),
                    row,
                    1,
                )
            )

            body = (
                self.excel.replace_placeholders(
                    self.body_text.get(
                        "1.0",
                        tk.END
                    ),
                    row,
                    1,
                )
            )

            html_body = (
                "<div style='"
                f"font-family:{self.font_name_var.get()};"
                f"font-size:{self.font_size_var.get()}pt;"
                "text-align:justify;'>"
                + body.replace(
                    "\n",
                    "<br>"
                )
                + "</div>"
            )

            attachments = []

            if "ANEXOS" in row.index:

                attachments = (
                    self.outlook.parse_attachment_field(
                        row["ANEXOS"]
                    )
                )

            valid_files, _ = (
                self.outlook
                .validate_attachments(
                    attachments
                )
            )

            mail = (
                self.outlook.create_email(
                    account_email=
                    self.account_var.get(),

                    to_email = recipient,

                    cc=
                    self.cc_var.get(),

                    bcc=
                    self.bcc_var.get(),

                    subject=
                    subject,

                    html_body=
                    html_body,

                    attachments=
                    valid_files,

                    include_signature=
                    self.include_signature_var
                    .get()
                )
            )

            self.outlook.display_email(
                mail
            )

        except Exception as e:

            messagebox.showerror(
                "Erro",
                str(e)
            )

    def refresh_placeholders(self):

        try:

            self.excel.load()

            self.load_placeholders()

            messagebox.showinfo(
                "Campos atualizados",
                "Os placeholders foram atualizados a partir do ficheiro Excel."
            )

        except Exception as e:

            messagebox.showerror(
                "Erro",
                str(e)
            )

    def create_all_messages(self):
        try:
            self.save_configuration()
            self.excel.load()

            selection_source = (
                self.excel_file_var.get(),
                len(self.excel.df)
            )

            # Usa a seleção guardada para este ficheiro.
            # Se ainda não houve seleção, processa todos os registos.
            if getattr(self, "_row_selection_source", None) == selection_source:
                rows_to_process = list(self.selected_rows)
            else:
                rows_to_process = list(range(self.excel.row_count()))

            total = len(rows_to_process)

            if total == 0:
                messagebox.showwarning(
                    "Aviso",
                    "Não existem registos selecionados para processar."
                )
                return

            answer = messagebox.askyesno(
                "Confirmação",
                f"Processar {total} registos?"
            )

            if not answer:
                return

            self.logger.clear()

            sent_count = 0
            draft_count = 0
            error_count = 0

            self.progress["value"] = 0
            self.progress["maximum"] = total

            for processed, i in enumerate(rows_to_process):
                recipient = ""

                try:
                    row = self.excel.get_row(i)

                    recipient = str(row["EMAIL"]).strip()

                    if not recipient:
                        self.logger.add_record(
                            i + 1,
                            "",
                            "Ignorado",
                            "Sem destinatário"
                        )
                        continue

                    if not self.outlook.is_valid_email(recipient):
                        error_count += 1

                        self.logger.add_record(
                            i + 1,
                            recipient,
                            "Erro",
                            "Email inválido"
                        )
                        continue

                    subject = self.excel.replace_placeholders(
                        self.subject_var.get(),
                        row,
                        i + 1
                    )

                    body = self.excel.replace_placeholders(
                        self.body_text.get("1.0", tk.END),
                        row,
                        i + 1
                    )

                    html_body = (
                        "<div style='"
                        f"font-family:{self.font_name_var.get()};"
                        f"font-size:{self.font_size_var.get()}pt;"
                        "text-align:justify;'>"
                        + body.replace("\n", "<br>")
                        + "</div>"
                    )

                    attachments = []

                    if "ANEXOS" in row.index:
                        attachments = self.outlook.parse_attachment_field(
                            row["ANEXOS"]
                        )

                    valid_files, invalid_files = (
                        self.outlook.validate_attachments(attachments)
                    )

                    mail = self.outlook.create_email(
                        account_email=self.account_var.get(),
                        to_email=recipient,
                        cc=self.cc_var.get(),
                        bcc=self.bcc_var.get(),
                        subject=subject,
                        html_body=html_body,
                        attachments=valid_files,
                        include_signature=self.include_signature_var.get()
                    )

                    if self.send_directly_var.get():
                        self.outlook.send_email(mail)
                        sent_count += 1
                        details = "Email enviado"
                    else:
                        self.outlook.display_email(mail)
                        draft_count += 1
                        details = "Email criado para revisão"

                    if invalid_files:
                        details += (
                            " | Anexos não encontrados: "
                            + "; ".join(invalid_files)
                        )

                    self.logger.add_record(
                        i + 1,
                        recipient,
                        "OK",
                        details
                    )

                except Exception as err:
                    error_count += 1

                    self.logger.add_record(
                        i + 1,
                        recipient,
                        "Erro",
                        str(err)
                    )

                finally:
                    # Atualiza a barra mesmo quando um registo é ignorado.
                    self.progress["value"] = processed + 1
                    self.root.update_idletasks()

            log_file = self.logger.save_log()

            summary = tk.Toplevel(self.root)
            self.set_window_icon(summary)

            summary.title("Processamento concluído")
            summary.geometry("500x250")
            summary.transient(self.root)
            summary.grab_set()

            msg = (
                "Processamento concluído\n\n"
                f"✓ Emails enviados: {sent_count}\n"
                f"✓ Emails criados para revisão: {draft_count}\n"
                f"✗ Erros: {error_count}"
            )

            ttk.Label(
                summary,
                text=msg,
                justify="left"
            ).pack(
                padx=20,
                pady=20
            )

            buttons = ttk.Frame(summary)
            buttons.pack(pady=10)

            ttk.Button(
                buttons,
                text="Abrir log",
                command=lambda: os.startfile(log_file)
            ).pack(
                side="left",
                padx=5
            )

            ttk.Button(
                buttons,
                text="Fechar",
                command=summary.destroy
            ).pack(
                side="left",
                padx=5
            )

        except Exception as e:
            messagebox.showerror(
                "Erro",
                str(e)
            )

    def run(self):

        self.root.mainloop()
