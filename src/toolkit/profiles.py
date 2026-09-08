"""Versioned, portable configuration; no workbook contents or secrets in exports."""

import copy
import json
import os
import re
import uuid
from pathlib import Path


def default_profile():
    resource = Path(__file__).with_name("default-profile.json")
    if resource.exists():
        return json.loads(resource.read_text(encoding="utf-8"))

    def cell(address, kind="text", required=False):
        return {
            "source": "main",
            "sheet": "Doc Proposta",
            "kind": "cell",
            "address": address,
            "type": kind,
            "required": required,
        }

    return {
        "version": 1,
        "id": "marcos",
        "name": "Padrão Marcos",
        "sources": {"main": "", "contacts": ""},
        "rules": {
            "decimal": ",",
            "thousands": ".",
            "date_format": "%d/%m/%Y",
            "separator": ";",
            "hidden": "include",
            "merged": "anchor",
            "blank_rows": "skip",
        },
        "proposal": {
            "ID_INTERNO": cell("K3"),
            "NOME_DA_EMPREITADA": cell("K4"),
            "LOCAL_DA_OBRA": cell("L6"),
            "DURACAO_DA_EMPREITADA": cell("N14"),
            "PRAZO": cell("N16", "date"),
            "DOCUMENTOS_DA_PROPOSTA": {
                **cell("J25:J45"),
                "kind": "range",
                "type": "list",
            },
            "HABILITACOES": {**cell("M25:M36"), "kind": "range", "type": "list"},
        },
        "items": {
            "source": "main",
            "sheet": "Lista de Items",
            "kind": "range",
            "address": "",
            "header_row": 1,
            "first_row": 2,
            "last_row": 0,
            "columns": {
                "ID": "ID",
                "Código": "Código",
                "Descrição": "Descrição",
                "Unidade": "Unidade",
                "Quantidade": "Quantidade",
            },
        },
        "contacts": {
            "source": "contacts",
            "sheet": "Contactos",
            "kind": "range",
            "address": "",
            "header_row": 1,
            "first_row": 2,
            "last_row": 0,
            "columns": {"ID": "ID", "Email": "Email", "Especialidade": "Especialidade"},
        },
        "tools": {
            "docs": {
                "templates": "",
                "output": "",
                "selected_models": "",
                "filename": "{ID_INTERNO}_{modelo}_{LOCAL_DA_OBRA}.docx",
                "date_format": "%d/%m/%Y",
                "alvara": True,
                "catalog": {},
            },
            "email": {
                "output": "",
                "base_folder": "",
                "deadline_days": 2,
                "subject": '{id_interno} - Pedido de preços - "{nome_empreitada}" - "{id_cliente}"',
                "signature": "Equipa de Orçamentação",
                "body": "",
                "reply_to": "",
                "cc": "",
                "bcc": "",
                "recipient_mode": "bcc",
                "excel_template": "",
                "report_template": "",
                "output_sheet": "Pedido de Preços",
                "output_cell": "B8",
                "output_columns": "Código;Descrição;Unidade;Quantidade",
                "filename": "{id_interno}_{id_cliente}_{especialidade}.xlsx",
                "received_folder": "2/Recebidas/{especialidade}",
                "smtp": {
                    "host": "",
                    "port": 465,
                    "security": "ssl",
                    "username": "",
                    "from": "",
                    "auth": True,
                },
                "imap": {
                    "host": "",
                    "port": 993,
                    "security": "ssl",
                    "username": "",
                    "folder": "INBOX",
                    "auth": True,
                },
            },
            "excel": {"preserve_styles": True},
        },
        "preferences": {"language": "pt-PT", "theme": "light"},
    }


def validate(profile):
    if not isinstance(profile, dict) or profile.get("version") != 1:
        raise ValueError("Perfil incompatível: é necessária a versão 1.")
    for key in (
        "id",
        "name",
        "sources",
        "rules",
        "proposal",
        "items",
        "contacts",
        "tools",
        "preferences",
    ):
        if key not in profile:
            raise ValueError(f"Perfil: falta {key}.")
    if not str(profile["name"]).strip():
        raise ValueError("O nome do perfil é obrigatório.")
    if not isinstance(profile["sources"], dict) or not all(
        isinstance(v, str) for v in profile["sources"].values()
    ):
        raise ValueError("As origens devem associar nomes a caminhos de ficheiros.")
    rules = profile["rules"]
    if (
        not rules.get("separator")
        or rules.get("hidden") not in ("include", "skip")
        or rules.get("merged") not in ("anchor", "fill", "error")
        or rules.get("blank_rows") not in ("skip", "stop")
    ):
        raise ValueError("Regras de importação inválidas.")
    if rules.get("decimal") == rules.get("thousands"):
        raise ValueError("Os separadores decimal e de milhares devem ser diferentes.")
    if profile["preferences"].get("theme") not in ("light", "dark"):
        raise ValueError("Tema inválido.")
    for name, mapping in profile["proposal"].items():
        if mapping.get("kind") not in (
            "cell",
            "range",
            "name",
            "table",
            "constant",
            "manual",
        ):
            raise ValueError(f"{name}: origem inválida.")
        if mapping.get("type") not in ("text", "number", "date", "list", "table"):
            raise ValueError(f"{name}: tipo inválido.")
    for name in ("items", "contacts"):
        m = profile[name]
        if not isinstance(m.get("columns"), dict) or not m["columns"]:
            raise ValueError(f"{name}: associe as colunas.")
        if int(m.get("header_row", 1)) < 1 or int(m.get("first_row", 2)) <= int(
            m.get("header_row", 1)
        ):
            raise ValueError(
                f"{name}: a primeira linha de dados deve seguir os cabeçalhos."
            )
        if m.get("kind") not in ("range", "name", "table") or not all(
            isinstance(v, str) and v for v in m["columns"].values()
        ):
            raise ValueError(f"{name}: origem ou associação inválida.")
    for name in ("smtp", "imap"):
        a = profile["tools"]["email"][name]
        if (
            a.get("security") not in ("ssl", "starttls", "plain")
            or not 1 <= int(a["port"]) <= 65535
        ):
            raise ValueError(f"{name}: porta ou segurança inválida.")
        if a["security"] == "plain" and a.get("auth"):
            raise ValueError(
                "Ligação sem TLS apenas disponível para relay sem autenticação."
            )
    return profile


def portable(profile):
    p = copy.deepcopy(validate(profile))
    p["sources"] = {key: "" for key in p["sources"]}
    private_keys = {
        "password",
        "secret",
        "token",
        "credential",
        "templates",
        "output",
        "base_folder",
        "excel_template",
        "report_template",
        "username",
        "from",
        "reply_to",
        "cc",
        "bcc",
        "host",
        "signature",
        "body",
    }

    def clean(value):
        if isinstance(value, dict):
            return {
                k: ("" if k.lower() in private_keys else clean(v))
                for k, v in value.items()
                if not any(
                    x in k.lower()
                    for x in ("password", "secret", "token", "credential")
                )
            }
        if isinstance(value, list):
            return [clean(v) for v in value]
        if isinstance(value, str) and (
            re.match(r"^[A-Za-z]:[\\/]", value) or value.startswith(("/", "\\\\"))
        ):
            return ""
        return value

    # Constant/manual values may contain project data and must be entered again.
    for m in p["proposal"].values():
        if m.get("kind") in ("constant", "manual"):
            m["value"] = ""
    for tool in p["tools"].get("excel", {}).values():
        if isinstance(tool, dict):
            for field in tool.get("fields", []):
                field["value"] = ""
    return clean(p)


class ProfileStore:
    def __init__(self, app, directory=None):
        self.app = app
        self.directory = Path(
            directory
            or Path(os.getenv("LOCALAPPDATA", Path.home() / ".config"))
            / "MarcosTools"
            / app
        )
        self.path = self.directory / "profiles.json"
        if self.path.exists():
            self.data = json.loads(self.path.read_text(encoding="utf-8"))
            for p in self.data["profiles"]:
                validate(p)
        else:
            self.data = {
                "active": "marcos",
                "profiles": [default_profile()],
                "migrated": False,
            }

    def active(self):
        return copy.deepcopy(
            next(p for p in self.data["profiles"] if p["id"] == self.data["active"])
        )

    def save(self, profile):
        validate(profile)
        updated = copy.deepcopy(self.data)
        updated["profiles"] = [
            p for p in updated["profiles"] if p["id"] != profile["id"]
        ] + [copy.deepcopy(profile)]
        updated["active"] = profile["id"]
        self.directory.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(".tmp")
        temporary.write_text(
            json.dumps(updated, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        os.replace(temporary, self.path)
        self.data = updated

    def duplicate(self, profile, name):
        result = copy.deepcopy(profile)
        result.update(id=str(uuid.uuid4()), name=name)
        return result

    def migrate(self, legacy=None):
        if self.data.get("migrated"):
            return
        p = self.active()
        if legacy:
            p["preferences"]["theme"] = legacy.get("theme", "light")
            p["tools"]["docs"].update(
                templates=legacy.get("template_path", ""),
                output=legacy.get("output_path", ""),
            )
        p["tools"]["docs"]["templates"] = os.getenv(
            "DOCSPROPOSTAS_TEMPLATES", p["tools"]["docs"]["templates"]
        )
        e = p["tools"]["email"]
        for protocol in ("smtp", "imap"):
            a = e[protocol]
            for key in ("host", "port", "username"):
                value = os.getenv(f"{protocol.upper()}_{key.upper()}")
                if value:
                    a[key] = int(value) if key == "port" else value
            a["username"] = a["username"] or os.getenv("SMTP_FROM", "")
            password = os.getenv(f"{protocol.upper()}_PASSWORD") or (
                os.getenv("SMTP_PASSWORD") if protocol == "imap" else None
            )
            if password:
                self.set_password(p["id"], protocol, password)
        e["smtp"]["from"] = os.getenv("SMTP_FROM", "")
        e["base_folder"] = os.getenv("PASTA_BASE_OBRAS", "")
        e["signature"] = os.getenv("ORCAMENTISTA_NOME", e["signature"])
        self.data["migrated"] = True
        self.save(p)

    def set_password(self, profile_id, protocol, value):
        from keyring.backends.Windows import WinVaultKeyring

        WinVaultKeyring().set_password(
            f"MarcosTools/{self.app}/{profile_id}", protocol, value
        )

    def password(self, profile_id, protocol):
        from keyring.backends.Windows import WinVaultKeyring

        return (
            WinVaultKeyring().get_password(
                f"MarcosTools/{self.app}/{profile_id}", protocol
            )
            or ""
        )
