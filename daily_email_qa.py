#!/usr/bin/env python3
"""
Envia o e-mail diário da equipe Testes e Qualidade no Gmail.

Regras principais:
- Usa a data atual em America/Sao_Paulo.
- Localiza exatamente a planilha do mês/ano atual:
  @Testes e Qualidade - Avaliação Diária <Mês>/<Ano>
- Descobre as abas dinamicamente e ignora apenas "capa" e "Resumo".
- Lê somente Data e Objetivo do dia para a data atual.
- Se encontrar "SEM EXPEDIENTE", encerra sem enviar o e-mail.
- Evita envios duplicados pelo assunto exato.
- Envia o e-mail diretamente pelo Gmail API.

Autenticação no GitHub Actions:
- Crie um Secret chamado GOOGLE_TOKEN_JSON contendo o JSON OAuth de usuário.

Para gerar esse JSON uma única vez no computador local:
  pip install -r requirements.txt
  python daily_email_qa.py --authorize client_secret.json --token-output token.json

Depois copie TODO o conteúdo de token.json para o Secret GOOGLE_TOKEN_JSON.
Nunca versione client_secret.json nem token.json.
"""

from __future__ import annotations

import argparse
import base64
import html
import json
import os
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


TIMEZONE_NAME = "America/Sao_Paulo"
TIMEZONE = ZoneInfo(TIMEZONE_NAME)
RECIPIENT = "testequali@nasajon.com.br"
DIARY_PREFIX = "@Testes e Qualidade - Avaliação Diária"
AUXILIARY_TABS = {"capa", "resumo"}

SCOPES = [
    "https://www.googleapis.com/auth/drive.metadata.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/gmail.send",
]

MONTHS_PT = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}

# Correções propositalmente conservadoras. O script não tenta "adivinhar"
# nomes de produtos, sistemas, pessoas, siglas ou termos técnicos.
CLEAR_WORD_CORRECTIONS = {
    "analise": "análise",
    "alteracao": "alteração",
    "atualizacao": "atualização",
    "automatizacao": "automatização",
    "configuracao": "configuração",
    "correcao": "correção",
    "criacao": "criação",
    "exclusao": "exclusão",
    "execucao": "execução",
    "exploratorio": "exploratório",
    "exploratorios": "exploratórios",
    "homologacao": "homologação",
    "integracao": "integração",
    "preparacao": "preparação",
    "regressao": "regressão",
    "validacao": "validação",
    "verificacao": "verificação",
    "versao": "versão",
}


class WorkflowError(RuntimeError):
    pass


class PermissionBlocked(WorkflowError):
    pass


@dataclass
class PersonGoals:
    name: str
    goals: list[str]


def log(message: str) -> None:
    print(message, flush=True)


def google_execute(request: Any, app: str, action: str) -> dict[str, Any]:
    try:
        return request.execute()
    except HttpError as exc:
        status = getattr(exc.resp, "status", None)
        if status in (401, 403):
            raise PermissionBlocked(
                f"BLOQUEADO: {app} / {action} (HTTP {status}). "
                "Verifique as permissões OAuth e os escopos autorizados."
            ) from exc
        raise WorkflowError(
            f"Falha em {app} / {action}: HTTP {status or 'desconhecido'} - {exc}"
        ) from exc


def normalize_header(value: str) -> str:
    return " ".join(value.strip().casefold().split())


def a1_sheet_name(title: str) -> str:
    return "'" + title.replace("'", "''") + "'"


def column_letter(index_zero_based: int) -> str:
    n = index_zero_based + 1
    letters = []
    while n:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


def drive_query_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def get_credentials() -> Credentials:
    raw = os.getenv("GOOGLE_TOKEN_JSON", "").strip()
    if not raw:
        raise WorkflowError(
            "Configuração ausente: defina o Secret/variável GOOGLE_TOKEN_JSON "
            "com o JSON OAuth autorizado."
        )

    try:
        info = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise WorkflowError("GOOGLE_TOKEN_JSON não contém um JSON válido.") from exc

    creds = Credentials.from_authorized_user_info(info, SCOPES)
    if creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as exc:
            raise WorkflowError(
                "Não foi possível renovar o token OAuth do Google. "
                "Gere novamente o GOOGLE_TOKEN_JSON."
            ) from exc

    if not creds.valid:
        raise WorkflowError("Credencial OAuth do Google inválida ou expirada.")
    return creds


def authorize(client_secret_path: str, token_output: str) -> int:
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        log("Instale google-auth-oauthlib: pip install -r requirements.txt")
        return 2

    path = Path(client_secret_path)
    if not path.exists():
        log(f"Arquivo não encontrado: {path}")
        return 2

    flow = InstalledAppFlow.from_client_secrets_file(str(path), SCOPES)
    creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")

    output = Path(token_output)
    output.write_text(creds.to_json(), encoding="utf-8")
    log(f"Token salvo em: {output}")
    log("Copie o conteúdo desse arquivo para o Secret GOOGLE_TOKEN_JSON.")
    log("Não faça commit desse arquivo.")
    return 0


def current_context() -> tuple[datetime, str, str]:
    now = datetime.now(TIMEZONE)
    month = MONTHS_PT[now.month]
    title = f"{DIARY_PREFIX} {month}/{now.year}"
    subject = f"Reunião diária {now:%d/%m} - Equipe de Testes e Qualidade"
    return now, title, subject


def find_diary(drive: Any, expected_title: str) -> dict[str, Any]:
    q = (
        f"name = '{drive_query_escape(expected_title)}' and "
        "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    )
    response = google_execute(
        drive.files().list(
            q=q,
            spaces="drive",
            pageSize=100,
            fields="files(id,name,mimeType,modifiedTime,webViewLink)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ),
        "Google Drive",
        "localizar diário do mês atual",
    )

    matches = [f for f in response.get("files", []) if f.get("name") == expected_title]
    if not matches:
        raise WorkflowError(
            f'Diário do mês atual não encontrado: "{expected_title}". Nenhum e-mail enviado.'
        )
    if len(matches) > 1:
        raise WorkflowError(
            f'Foram encontrados {len(matches)} arquivos com o título exato "{expected_title}". '
            "Diário ambíguo; nenhum e-mail enviado."
        )
    return matches[0]


def discover_person_sheets(sheets_api: Any, spreadsheet_id: str) -> list[dict[str, Any]]:
    metadata = google_execute(
        sheets_api.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields=(
                "properties(title,locale,timeZone),"
                "sheets(properties(sheetId,title,index,hidden,gridProperties(rowCount,columnCount)))"
            ),
        ),
        "Google Sheets",
        "ler metadados e abas do diário",
    )

    tabs: list[dict[str, Any]] = []
    for item in metadata.get("sheets", []):
        props = item.get("properties", {})
        title = str(props.get("title", ""))
        if props.get("hidden", False):
            continue
        if title.casefold() in AUXILIARY_TABS:
            continue
        tabs.append(props)

    tabs.sort(key=lambda p: int(p.get("index", 0)))
    if not tabs:
        raise WorkflowError("Nenhuma aba de pessoa foi encontrada no diário atual.")
    return tabs


def parse_sheet_date(value: str) -> date | None:
    text = value.strip()
    if not text:
        return None

    candidates = [text]
    if " " in text:
        candidates.append(text.split(" ", 1)[0])

    for candidate in candidates:
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
            try:
                return datetime.strptime(candidate, fmt).date()
            except ValueError:
                pass
    return None


def first_cell(row: list[Any] | None) -> str:
    if not row:
        return ""
    return str(row[0]) if row[0] is not None else ""


def correct_goal_text(text: str) -> str:
    """Correções conservadoras, sem tocar em IDs numéricos ou termos técnicos incertos."""
    value = text.strip()
    if not value:
        return value

    token_pattern = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ]+")

    def replace_token(match: re.Match[str]) -> str:
        token = match.group(0)
        lower = token.casefold()
        replacement = CLEAR_WORD_CORRECTIONS.get(lower)
        if not replacement:
            return token

        if token.isupper():
            # Acrônimos e tokens em caixa alta são preservados.
            return token
        if token[:1].isupper():
            return replacement[:1].upper() + replacement[1:]
        return replacement

    value = token_pattern.sub(replace_token, value)

    # Capitalização comum de início de frase, sem alterar números ou siglas.
    chars = list(value)
    for i, ch in enumerate(chars):
        if ch.isalpha():
            if ch.islower():
                chars[i] = ch.upper()
            break
    return "".join(chars)


def split_goals(objective_cells: Iterable[str]) -> list[str]:
    goals: list[str] = []
    for cell in objective_cells:
        for part in cell.split(","):
            part = part.strip()
            if part:
                goals.append(correct_goal_text(part))
    return goals


def read_person_goals(
    sheets_api: Any,
    spreadsheet_id: str,
    sheet_props: dict[str, Any],
    today: date,
) -> PersonGoals:
    sheet_name = str(sheet_props.get("title", ""))
    grid = sheet_props.get("gridProperties", {}) or {}
    row_count = max(int(grid.get("rowCount", 1000)), 2)
    col_count = max(int(grid.get("columnCount", 2)), 2)
    last_col = column_letter(col_count - 1)
    quoted = a1_sheet_name(sheet_name)

    header_response = google_execute(
        sheets_api.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{quoted}!A1:{last_col}1",
            valueRenderOption="FORMATTED_VALUE",
        ),
        "Google Sheets",
        f"ler cabeçalhos da aba {sheet_name}",
    )
    header_values = (header_response.get("values") or [[]])[0]
    header_values = [str(v) for v in header_values]

    data_idx = None
    objective_idx = None
    for idx, header in enumerate(header_values):
        normalized = normalize_header(header)
        if normalized == "data":
            data_idx = idx
        elif normalized == "objetivo do dia":
            objective_idx = idx

    if data_idx is None:
        raise WorkflowError(f'A aba "{sheet_name}" não possui o cabeçalho "Data".')

    if objective_idx is None:
        # Compatibilidade conservadora com abas antigas em que o cabeçalho do
        # objetivo ficou vazio, mas a coluna imediatamente após Data contém os objetivos.
        fallback_idx = data_idx + 1
        fallback_header = header_values[fallback_idx].strip() if fallback_idx < len(header_values) else ""
        if fallback_idx < col_count and not fallback_header:
            objective_idx = fallback_idx
            log(
                f'AVISO: aba "{sheet_name}" sem cabeçalho "Objetivo do dia"; '
                "usando a coluna imediatamente após Data por compatibilidade."
            )
        else:
            raise WorkflowError(
                f'A aba "{sheet_name}" não possui o cabeçalho "Objetivo do dia".'
            )

    data_col = column_letter(data_idx)
    objective_col = column_letter(objective_idx)
    ranges = [
        f"{quoted}!{data_col}2:{data_col}{row_count}",
        f"{quoted}!{objective_col}2:{objective_col}{row_count}",
    ]

    values_response = google_execute(
        sheets_api.spreadsheets().values().batchGet(
            spreadsheetId=spreadsheet_id,
            ranges=ranges,
            valueRenderOption="FORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        ),
        "Google Sheets",
        f"ler Data e Objetivo do dia da aba {sheet_name}",
    )

    value_ranges = values_response.get("valueRanges", [])
    data_values = value_ranges[0].get("values", []) if len(value_ranges) > 0 else []
    objective_values = value_ranges[1].get("values", []) if len(value_ranges) > 1 else []

    objective_cells: list[str] = []
    for row_idx, row in enumerate(data_values):
        sheet_date = parse_sheet_date(first_cell(row))
        if sheet_date != today:
            continue
        objective = first_cell(objective_values[row_idx]) if row_idx < len(objective_values) else ""
        objective_cells.append(objective.strip())

    for objective in objective_cells:
        if "sem expediente" in objective.casefold():
            raise StopIteration(sheet_name)

    return PersonGoals(name=sheet_name, goals=split_goals(objective_cells))


def header_value(headers: list[dict[str, Any]], name: str) -> str:
    for header in headers:
        if str(header.get("name", "")).casefold() == name.casefold():
            return str(header.get("value", ""))
    return ""


def find_existing_sent_message(gmail: Any, subject: str) -> str | None:
    page_token: str | None = None
    while True:
        kwargs: dict[str, Any] = {
            "userId": "me",
            "q": f'in:sent subject:"{subject}"',
            "maxResults": 500,
        }
        if page_token:
            kwargs["pageToken"] = page_token

        response = google_execute(
            gmail.users().messages().list(**kwargs),
            "Gmail",
            "verificar e-mail duplicado enviado",
        )

        for message in response.get("messages", []):
            message_id = message.get("id")
            if not message_id:
                continue
            details = google_execute(
                gmail.users().messages().get(
                    userId="me",
                    id=message_id,
                    format="metadata",
                    metadataHeaders=["Subject"],
                ),
                "Gmail",
                "ler assunto de e-mail enviado existente",
            )
            headers = details.get("payload", {}).get("headers", [])
            if header_value(headers, "Subject") == subject:
                return str(message_id)

        page_token = response.get("nextPageToken")
        if not page_token:
            return None


def build_email(persons: list[PersonGoals], subject: str) -> EmailMessage:
    text_lines = ["Seguem as metas do dia:", ""]
    html_parts = ["<p>Seguem as metas do dia:</p>"]

    for person in persons:
        text_lines.append(person.name)
        html_parts.append(f"<p><strong>{html.escape(person.name)}</strong></p>")

        if person.goals:
            html_parts.append("<ul>")
            for goal in person.goals:
                text_lines.append(f"• {goal}")
                html_parts.append(f"<li>{html.escape(goal)}</li>")
            html_parts.append("</ul>")
        text_lines.append("")

    message = EmailMessage()
    message["To"] = RECIPIENT
    message["Subject"] = subject
    message.set_content("\n".join(text_lines).rstrip() + "\n")
    message.add_alternative("\n".join(html_parts), subtype="html")
    return message


def send_email(gmail: Any, message: EmailMessage) -> dict[str, Any]:
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii")
    return google_execute(
        gmail.users().messages().send(
            userId="me",
            body={"raw": raw},
        ),
        "Gmail",
        "enviar e-mail",
    )


def run() -> int:
    now, expected_title, subject = current_context()
    today = now.date()
    log(f"Data da execução ({TIMEZONE_NAME}): {today:%d/%m/%Y}")
    log(f"Diário esperado: {expected_title}")

    creds = get_credentials()
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    sheets_api = build("sheets", "v4", credentials=creds, cache_discovery=False)
    gmail = build("gmail", "v1", credentials=creds, cache_discovery=False)

    diary = find_diary(drive, expected_title)
    spreadsheet_id = str(diary["id"])
    tabs = discover_person_sheets(sheets_api, spreadsheet_id)

    persons: list[PersonGoals] = []
    try:
        for tab in tabs:
            persons.append(read_person_goals(sheets_api, spreadsheet_id, tab, today))
    except StopIteration as stop:
        person = str(stop.value or "")
        suffix = f' na aba "{person}"' if person else ""
        log(f"SEM EXPEDIENTE encontrado{suffix}. Nenhum e-mail enviado.")
        return 0

    existing = find_existing_sent_message(gmail, subject)
    if existing:
        log(f'E-mail já enviado: "{subject}". Nenhum envio duplicado realizado.')
        return 0

    message = build_email(persons, subject)
    sent = send_email(gmail, message)
    message_id = sent.get("id", "desconhecido")
    log(f'E-mail enviado: "{subject}".')
    log(f"Message ID: {message_id}")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Envia o e-mail diário da equipe Testes e Qualidade.")
    parser.add_argument(
        "--authorize",
        metavar="CLIENT_SECRET_JSON",
        help="Gera localmente o token OAuth usado pelo GitHub Actions.",
    )
    parser.add_argument(
        "--token-output",
        default="token.json",
        help="Arquivo de saída usado com --authorize (padrão: token.json).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.authorize:
        return authorize(args.authorize, args.token_output)

    try:
        return run()
    except PermissionBlocked as exc:
        log(str(exc))
        return 1
    except WorkflowError as exc:
        log(f"ERRO: {exc}")
        return 1
    except Exception as exc:
        log(f"ERRO INESPERADO: {type(exc).__name__}: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
