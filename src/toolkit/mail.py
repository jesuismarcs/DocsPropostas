"""Configurable transports and local message journal. No automatic retries."""

import hashlib
import imaplib
import json
import mimetypes
import re
import smtplib
import sqlite3
import ssl
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import make_msgid, getaddresses
from pathlib import Path


def connection(account, password, protocol="smtp"):
    if not account["host"]:
        raise ValueError("Configure o servidor.")
    context = ssl.create_default_context()
    security = account["security"]
    if security == "plain" and account.get("auth", True):
        raise ValueError("Autenticação exige TLS.")
    if protocol == "smtp":
        cls = smtplib.SMTP_SSL if security == "ssl" else smtplib.SMTP
        kwargs = {"context": context} if security == "ssl" else {}
        client = cls(account["host"], int(account["port"]), timeout=30, **kwargs)
        try:
            client.ehlo()
            if security == "starttls":
                client.starttls(context=context)
                client.ehlo()
            if account.get("auth", True):
                if not account["username"] or not password:
                    raise ValueError("Configure utilizador e palavra-passe.")
                client.login(account["username"], password)
            return client
        except Exception:
            client.close()
            raise
    cls = imaplib.IMAP4_SSL if security == "ssl" else imaplib.IMAP4
    kwargs = {"ssl_context": context} if security == "ssl" else {}
    client = cls(account["host"], int(account["port"]), timeout=30, **kwargs)
    try:
        if security == "starttls":
            client.starttls(ssl_context=context)
        client.login(account["username"], password)
        status, _ = client.select(account.get("folder", "INBOX"), readonly=True)
        if status != "OK":
            raise ValueError("Pasta IMAP indisponível.")
        return client
    except Exception:
        client.logout()
        raise


def addresses(text):
    if "\r" in text or "\n" in text:
        raise ValueError("Endereço com quebra de linha.")
    if not text.strip():
        return []
    result = []
    for _, address in getaddresses([text.replace(";", ",")]):
        if not re.fullmatch(r"[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+", address):
            raise ValueError("Endereço de email inválido.")
        if address.casefold() not in [a.casefold() for a in result]:
            result.append(address)
    return result


class Journal:
    def __init__(self, path):
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(path)
        self.db.executescript("""
            CREATE TABLE IF NOT EXISTS sent (id TEXT PRIMARY KEY, project TEXT, specialty TEXT, recipients TEXT, status TEXT, created TEXT);
            CREATE TABLE IF NOT EXISTS received (id TEXT PRIMARY KEY, files TEXT);
        """)

    def close(self):
        self.db.close()


def compose(settings, recipients, files, fields):
    a = settings["smtp"]
    sender = addresses(a["from"])
    if len(sender) != 1:
        raise ValueError("Indique um único remetente.")
    to = addresses(recipients)
    cc, bcc = addresses(settings.get("cc", "")), addresses(settings.get("bcc", ""))
    envelope = list(
        dict.fromkeys(
            to + cc + bcc + (sender if settings.get("recipient_mode") == "bcc" else [])
        )
    )
    if not to:
        raise ValueError("Nenhum destinatário.")
    msg = EmailMessage()
    msg["From"] = sender[0]
    msg["To"] = sender[0] if settings.get("recipient_mode") == "bcc" else ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    if settings.get("reply_to"):
        msg["Reply-To"] = ", ".join(addresses(settings["reply_to"]))
    msg["Subject"] = settings["subject"].format_map(fields)
    msg["Message-ID"] = make_msgid()
    body = settings["body"].format_map(fields)
    if settings.get("signature"):
        body += "\n\n" + settings["signature"]
    msg.set_content(body)
    for file in files:
        path = Path(file)
        main, sub = (
            mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        ).split("/", 1)
        msg.add_attachment(
            path.read_bytes(), maintype=main, subtype=sub, filename=path.name
        )
    return msg, envelope


def send(settings, password, recipients, files, fields, journal_path, simulate=False):
    message, envelope = compose(settings, recipients, files, fields)
    if simulate:
        return {
            "status": "simulated",
            "recipients": {x: "simulated" for x in envelope},
            "message_id": message["Message-ID"],
        }
    journal = Journal(journal_path)
    ident = message["Message-ID"]
    journal.db.execute(
        "INSERT INTO sent VALUES (?,?,?,?,?,?)",
        (
            ident,
            fields["id_interno"],
            fields["id_cliente"],
            json.dumps({x: "pending" for x in envelope}),
            "pending",
            datetime.now().isoformat(),
        ),
    )
    journal.db.commit()
    client = None
    try:
        client = connection(settings["smtp"], password)
        try:
            refused = client.send_message(
                message, from_addr=settings["smtp"]["from"], to_addrs=envelope
            )
        except smtplib.SMTPRecipientsRefused as exc:
            refused = exc.recipients
        statuses = {x: "refused" if x in refused else "accepted" for x in envelope}
        status = "partial" if refused else "accepted"
        journal.db.execute(
            "UPDATE sent SET status=?, recipients=? WHERE id=?",
            (status, json.dumps(statuses), ident),
        )
        journal.db.commit()
        return {"status": status, "recipients": statuses, "message_id": ident}
    except Exception:
        journal.db.execute(
            "UPDATE sent SET status='unknown', recipients=? WHERE id=?",
            (json.dumps({x: "unknown" for x in envelope}), ident),
        )
        journal.db.commit()
        raise
    finally:
        if client:
            try:
                client.quit()
            except Exception:
                client.close()
        journal.close()


def receive(settings, password, journal_path, start, end, find_folder, parse_subject):
    import email
    from email.header import decode_header, make_header

    journal = Journal(journal_path)
    client = connection(settings["imap"], password, "imap")
    result = []
    try:
        first = datetime.strptime(start, "%d-%m-%Y")
        last = datetime.strptime(end, "%d-%m-%Y") + timedelta(days=1)
        if last <= first:
            raise ValueError("Intervalo de datas inválido.")
        months = [
            "Jan",
            "Feb",
            "Mar",
            "Apr",
            "May",
            "Jun",
            "Jul",
            "Aug",
            "Sep",
            "Oct",
            "Nov",
            "Dec",
        ]

        def imap_date(d):
            return f"{d.day:02d}-{months[d.month - 1]}-{d.year}"

        status, data = client.uid(
            "search", None, f'(SINCE "{imap_date(first)}" BEFORE "{imap_date(last)}")'
        )
        if status != "OK":
            raise ValueError("Pesquisa IMAP falhou.")
        for uid in data[0].split():
            status, parts = client.uid("fetch", uid, "(BODY.PEEK[])")
            if status != "OK":
                continue
            raw = next((p[1] for p in parts if isinstance(p, tuple)), None)
            if raw is None:
                continue
            message = email.message_from_bytes(raw)
            account_key = settings["imap"]["host"] + "/" + settings["imap"]["username"]
            message_key = message.get("Message-ID") or hashlib.sha256(raw).hexdigest()
            ident = hashlib.sha256(
                (account_key + "/" + message_key).encode()
            ).hexdigest()
            if journal.db.execute(
                "SELECT 1 FROM received WHERE id=?", (ident,)
            ).fetchone():
                continue
            subject = str(make_header(decode_header(message.get("Subject", ""))))
            match = None
            references = re.findall(
                r"<[^>]+>",
                message.get("In-Reply-To", "") + " " + message.get("References", ""),
            )
            for reference in reversed(references):
                match = journal.db.execute(
                    "SELECT project,specialty FROM sent WHERE id=?", (reference,)
                ).fetchone()
                if match:
                    break
            match = match or parse_subject(subject)
            if not match:
                continue
            folder = find_folder(settings["base_folder"], match[0])
            if not folder:
                continue
            safe_specialty = (
                re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", match[1]).strip(" .")
                or "especialidade"
            )
            destination = (
                Path(folder)
                / settings["received_folder"].format(especialidade=safe_specialty)
            ).resolve()
            if not destination.is_relative_to(Path(folder).resolve()):
                raise ValueError("Pasta de respostas fora da obra.")
            destination.mkdir(parents=True, exist_ok=True)
            saved = []
            for part in message.walk():
                if part.get_filename():
                    name = str(make_header(decode_header(part.get_filename())))
                    name = (
                        re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
                        or "anexo"
                    )
                    content = part.get_payload(decode=True)
                    if content is None:
                        continue
                    target = destination / (ident[:12] + "_" + name[:120])
                    if target.exists() and target.read_bytes() != content:
                        raise ValueError(
                            "Colisão de anexo; ficheiro existente preservado."
                        )
                    if not target.exists():
                        with target.open("xb") as f:
                            f.write(content)
                    saved.append(str(target))
            journal.db.execute(
                "INSERT INTO received VALUES (?,?)", (ident, json.dumps(saved))
            )
            journal.db.commit()
            result.append({"project": match[0], "specialty": match[1], "files": saved})
        return result
    finally:
        client.logout()
        journal.close()
