"""
Gestão de modelos de email.
"""

import json
import os


class TemplateManager:

    def __init__(self):

        self.templates_folder = os.path.join(
            os.getenv("APPDATA"),
            "EmailBatchPro",
            "modelos"
        )

        os.makedirs(
            self.templates_folder,
            exist_ok=True
        )

    # --------------------------------------
    # Caminho do modelo
    # --------------------------------------

    def get_template_path(
        self,
        template_name
    ):

        return os.path.join(
            self.templates_folder,
            f"{template_name}.json"
        )

    # --------------------------------------
    # Guardar modelo
    # --------------------------------------

    def save_template(
        self,
        template_name,
        subject,
        body
    ):

        data = {
            "subject": subject,
            "body": body
        }

        path = self.get_template_path(
            template_name
        )

        with open(
            path,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                data,
                f,
                indent=4,
                ensure_ascii=False
            )

    # --------------------------------------
    # Carregar modelo
    # --------------------------------------

    def load_template(
        self,
        template_name
    ):

        path = self.get_template_path(
            template_name
        )

        if not os.path.exists(path):

            raise FileNotFoundError(
                f"Modelo '{template_name}' não encontrado."
            )

        with open(
            path,
            "r",
            encoding="utf-8"
        ) as f:

            return json.load(f)

    # --------------------------------------
    # Listar modelos
    # --------------------------------------

    def list_templates(self):

        templates = []

        for file_name in os.listdir(
            self.templates_folder
        ):

            if file_name.endswith(".json"):

                templates.append(
                    file_name[:-5]
                )

        return sorted(templates)

    # --------------------------------------
    # Apagar modelo
    # --------------------------------------

    def delete_template(
        self,
        template_name
    ):

        path = self.get_template_path(
            template_name
        )

        if os.path.exists(path):

            os.remove(path)