"""
Gestão dos logs de processamento.
"""

import os
import pandas as pd

from utils import current_timestamp


class LoggerManager:

    def __init__(self):

        self.records = []

        self.log_dir = os.path.join(
            os.getenv("APPDATA"),
            "EmailBatchPro",
            "logs"
        )

        os.makedirs(
            self.log_dir,
            exist_ok=True
        )

    # --------------------------------------
    # Limpar registos
    # --------------------------------------

    def clear(self):

        self.records = []

    # --------------------------------------
    # Adicionar registo
    # --------------------------------------

    def add_record(
        self,
        line,
        recipient,
        status,
        details=""
    ):

        self.records.append(
            {
                "Linha": line,
                "Destinatário": recipient,
                "Estado": status,
                "Detalhes": details
            }
        )

    # --------------------------------------
    # Guardar log Excel
    # --------------------------------------

    def save_log(self):

        if len(self.records) == 0:
            return None

        filename = os.path.join(
            self.log_dir,
            f"envio_{current_timestamp()}.xlsx"
        )

        df = pd.DataFrame(
            self.records
        )

        df.to_excel(
            filename,
            index=False
        )

        return filename

    # --------------------------------------
    # Número de registos
    # --------------------------------------

    def count(self):

        return len(self.records)