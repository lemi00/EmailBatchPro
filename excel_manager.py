"""
Gestão do ficheiro Excel e substituição de placeholders.
"""

import pandas as pd
import os

from utils import (
    column_number_to_name,
    current_date,
    current_time
)


class ExcelManager:

    def __init__(self, excel_file="dados.xlsx"):

        self.excel_file = excel_file
        self.df = None

    # --------------------------------------------------
    # Leitura do Excel
    # --------------------------------------------------

    def load(self):

        self.df = pd.read_excel(
            self.excel_file,
            header=0,
            dtype=str
        )

        self.df = self.df.fillna("")

        return self.df
        
    # --------------------------------------------------
    # Leitura da primeira linha
    # --------------------------------------------------

    def get_placeholders(self):

        if self.df is None:
            self.load()

        return [
            str(col).strip()
            for col in self.df.columns
        ]

    # --------------------------------------------------
    # Validação do placeholder EMAIL
    # --------------------------------------------------
        
    def validate_required_columns(self):

        if self.df is None:
            self.load()

        columns = [
            str(col).strip().upper()
            for col in self.df.columns
        ]

        if "EMAIL" not in columns:

            raise ValueError(
                "Não foi encontrada a coluna EMAIL."
            )

        return True

    # --------------------------------------------------
    # Número de linhas
    # --------------------------------------------------

    def row_count(self):

        if self.df is None:
            return 0

        return len(self.df)

    # --------------------------------------------------
    # Obter linha
    # --------------------------------------------------

    def get_row(self, index):

        if self.df is None:
            self.load()

        return self.df.iloc[index]

    # --------------------------------------------------
    # Placeholders
    # --------------------------------------------------

    def replace_placeholders(
        self,
        text,
        row,
        row_number,
    ):

        result = text

        # Colunas do Excel

        for column_name in self.df.columns:

            placeholder = (
                "{"
                + str(column_name).strip()
                + "}"
            )

            result = result.replace(
                placeholder,
                str(row[column_name])
            )
        
        # Placeholders especiais

        result = result.replace(
            "{LINHA}",
            str(row_number)
        )

        result = result.replace(
            "{DATA}",
            current_date()
        )

        result = result.replace(
            "{HORA}",
            current_time()
        )

        return result

    # --------------------------------------------------
    # Preview da primeira linha
    # --------------------------------------------------

    def build_preview(
        self,
        subject_template,
        body_template,
    ):

        if self.df is None:
            self.load()

        if len(self.df) == 0:
            raise ValueError(
                "O ficheiro dados.xlsx está vazio."
            )

        row = self.df.iloc[0]

        return {

            "to": str(row["EMAIL"]),

            "subject": self.replace_placeholders(
                subject_template,
                row,
                1,
            ),

            "body": self.replace_placeholders(
                body_template,
                row,
                1,
            )
        }

    # --------------------------------------------------
    # Verificar ficheiro carregado
    # --------------------------------------------------

    def is_loaded(self):

        return self.df is not None

    # --------------------------------------------------
    # Número de colunas
    # --------------------------------------------------

    def column_count(self):

        if self.df is None:
            self.load()

        return len(self.df.columns)

    # --------------------------------------------------
    # Validação do Excel
    # --------------------------------------------------

    def validate_data(
        self,
        email_validator=None
    ):

        if self.df is None:
            self.load()

        self.validate_required_columns()

        total_rows = len(self.df)

        valid_emails = 0

        empty_email_rows = []

        invalid_email_rows = []

        missing_attachment_rows = []

        validation_records = []

        for idx, row in self.df.iterrows():

            line = idx + 2

            email = str(
                row["EMAIL"]
            ).strip()

            if not email:

                empty_email_rows.append(
                    f"Linha {line}"
                )

                validation_records.append(
                    {
                        "Linha": line,
                        "Tipo": "Email vazio",
                        "Valor": ""
                    }
                )

                continue

            if (
                email_validator
                and
                email_validator(email)
            ):

                valid_emails += 1

            else:
                invalid_email_rows.append(
                    f"Linha {line}: {email}"
                )

                validation_records.append(
                    {
                        "Linha": line,
                        "Tipo": "Email inválido",
                        "Valor": email
                    }
                )

            if "ANEXOS" in row.index:

                attachments = str(
                    row["ANEXOS"]
                ).strip()

                if attachments:

                    files = [
                        f.strip()
                        for f in attachments.split(";")
                    ]

                    for file in files:

                        if (
                            file
                            and
                            not os.path.exists(file)
                        ):

                            missing_attachment_rows.append(
                                f"Linha {line}: {file}"
                            )

                            validation_records.append(
                                {
                                    "Linha": line,
                                    "Tipo": "Anexo inexistente",
                                    "Valor": file
                                }
                            )

        return {

            "rows":
                total_rows,

            "valid_emails":
                valid_emails,

            "empty_email_rows":
                empty_email_rows,

            "invalid_email_rows":
                invalid_email_rows,

            "missing_attachment_rows":
                missing_attachment_rows,

            "validation_records":
                validation_records

        }