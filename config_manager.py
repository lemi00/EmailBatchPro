"""
Gestão da configuração da aplicação
"""

import json
import os


class ConfigManager:

    def __init__(self):

        self.config_file = os.path.join(
            os.getenv("APPDATA"),
            "EmailBatchPro",
            "config.json"
        )

        os.makedirs(
            os.path.dirname(self.config_file),
            exist_ok=True
        )

        self.default_config = {
            "account": "",
            "cc": "",
            "bcc": "",
            "subject": "",
            "body": "",
            "attachments_column": "D",
            "include_signature": True,
            "send_directly": False
        }

    def load(self):

        if not os.path.exists(self.config_file):

            self.save(self.default_config)

            return self.default_config.copy()

        try:

            with open(
                self.config_file,
                "r",
                encoding="utf-8"
            ) as f:

                data = json.load(f)

            config = self.default_config.copy()

            config.update(data)

            return config

        except Exception:

            return self.default_config.copy()

    def save(self, data):

        with open(
            self.config_file,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                data,
                f,
                indent=4,
                ensure_ascii=False
            )

    def reset(self):

        self.save(
            self.default_config
        )