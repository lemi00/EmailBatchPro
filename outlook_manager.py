"""
Integração com Microsoft Outlook.
"""

import os
import re

import pythoncom
import win32com.client


class OutlookManager:

    EMAIL_REGEX = re.compile(
        r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    )

    def __init__(self):

        pythoncom.CoInitialize()

        self.outlook = win32com.client.Dispatch(
            "Outlook.Application"
        )

        self.namespace = self.outlook.GetNamespace(
            "MAPI"
        )

    # --------------------------------------------------
    # Contas configuradas
    # --------------------------------------------------

    def get_accounts(self):

        accounts = []

        try:

            for account in self.namespace.Accounts:

                try:

                    accounts.append(
                        account.SmtpAddress
                    )

                except Exception:

                    pass

        except Exception:

            pass

        return accounts

    # --------------------------------------------------
    # Validação email
    # --------------------------------------------------

    def is_valid_email(
        self,
        email
    ):

        if not email:

            return False

        return bool(
            self.EMAIL_REGEX.match(
                email.strip()
            )
        )

    # --------------------------------------------------
    # Assinatura Outlook
    # --------------------------------------------------

    def get_default_signature(self):

        try:

            mail = self.outlook.CreateItem(0)

            return mail.HTMLBody

        except Exception:

            return ""

    # --------------------------------------------------
    # Processar anexos
    # --------------------------------------------------

    def parse_attachment_field(
        self,
        value
    ):

        attachments = []

        if value is None:

            return attachments

        value = str(value).strip()

        if not value:

            return attachments

        for item in value.split(";"):

            item = item.strip()

            if item:

                attachments.append(item)

        return attachments

    # --------------------------------------------------
    # Validar anexos
    # --------------------------------------------------

    def validate_attachments(
        self,
        attachments
    ):

        valid = []

        invalid = []

        for file_path in attachments:

            if os.path.exists(
                file_path
            ):

                valid.append(
                    file_path
                )

            else:

                invalid.append(
                    file_path
                )

        return valid, invalid

    # --------------------------------------------------
    # Selecionar conta Outlook
    # --------------------------------------------------

    def set_account(
        self,
        mail,
        account_email
    ):

        if not account_email:

            return

        try:

            for account in self.namespace.Accounts:

                if (
                    account.SmtpAddress.lower()
                    ==
                    account_email.lower()
                ):

                    mail._oleobj_.Invoke(
                        *(64209, 0, 8, 0, account)
                    )

                    break

        except Exception:

            pass

    # --------------------------------------------------
    # Criar email
    # --------------------------------------------------

    def create_email(
        self,
        account_email,
        to_email,
        cc,
        bcc,
        subject,
        html_body,
        attachments=None,
        include_signature=True
    ):

        if attachments is None:

            attachments = []

        mail = self.outlook.CreateItem(
            0
        )

        self.set_account(
            mail,
            account_email
        )

        mail.To = to_email

        mail.CC = cc

        mail.BCC = bcc

        mail.Subject = subject

        if include_signature:

            mail.Display()

            signature = mail.HTMLBody

            # Elimina linhas vazias do início
            while True:

                old = signature

                signature = re.sub(
                    r'^\s*<p[^>]*>(&nbsp;|\s|<br\s*/?>)*</p>',
                    '',
                    signature,
                    flags=re.IGNORECASE
                )

                signature = re.sub(
                    r'^\s*<div[^>]*>(&nbsp;|\s|<br\s*/?>)*</div>',
                    '',
                    signature,
                    flags=re.IGNORECASE
                )

                if signature == old:
                    break

            mail.HTMLBody = (
                html_body
                + signature
            )

        else:

            mail.HTMLBody = html_body

        for file_path in attachments:

            try:

                mail.Attachments.Add(
                    file_path
                )

            except Exception:

                pass

        return mail

    # --------------------------------------------------
    # Abrir email
    # --------------------------------------------------

    def display_email(
        self,
        mail
    ):

        mail.Display()

    # --------------------------------------------------
    # Enviar email
    # --------------------------------------------------

    def send_email(
        self,
        mail
    ):

        mail.Send()