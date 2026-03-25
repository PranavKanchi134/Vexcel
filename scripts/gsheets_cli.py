#!/usr/bin/env python3
"""Direct Google Sheets bridge for Codex-driven editing.

Supports:
- OAuth bootstrap for a user account
- service account auth for automation
- spreadsheet aliases via vexcel-sheets.json
- read/write/clear/append workflows with formulas preserved through USER_ENTERED
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import sys
import warnings
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

warnings.filterwarnings("ignore", category=FutureWarning, module=r"google(\.|$)")
warnings.filterwarnings("ignore", message=r".*urllib3 v2 only supports OpenSSL 1\.1\.1\+.*")

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials as UserCredentials
    from google.oauth2.service_account import Credentials as ServiceAccountCredentials
    from googleapiclient.discovery import build
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError as exc:  # pragma: no cover
    print(
        "Missing Google Sheets dependencies. Install them with:\n"
        "  python3 -m pip install -r scripts/requirements-gsheets.txt\n"
        "or create a repo venv and install there.\n"
        f"Import error: {exc}",
        file=sys.stderr,
    )
    raise


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SHEET_URL_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")
DEFAULT_TOKEN_FILE = ".secrets/google-sheets-token.json"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Read and edit live Google Sheets data directly from Codex."
    )
    parser.add_argument(
        "--config",
        default=os.environ.get("VEXCEL_SHEETS_CONFIG", "vexcel-sheets.json"),
        help="Path to alias config JSON. Defaults to vexcel-sheets.json",
    )
    parser.add_argument(
        "--token-file",
        default=os.environ.get("GOOGLE_OAUTH_TOKEN_FILE", DEFAULT_TOKEN_FILE),
        help="Where OAuth user tokens are stored for interactive auth.",
    )
    parser.add_argument(
        "--oauth-client-secret",
        default=os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET_FILE", ""),
        help="OAuth desktop client secret JSON for user auth.",
    )
    parser.add_argument(
        "--service-account",
        default=os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "")),
        help="Service account JSON path for automation.",
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    auth = subparsers.add_parser("auth", help="Run interactive OAuth and cache a user token locally.")
    auth.add_argument(
        "--force",
        action="store_true",
        help="Ignore any existing token file and re-run auth.",
    )

    list_sheets = subparsers.add_parser("list-sheets", help="List sheet tabs in a spreadsheet.")
    add_spreadsheet_arg(list_sheets)

    get_cmd = subparsers.add_parser("get", help="Read a range from Google Sheets.")
    add_spreadsheet_arg(get_cmd)
    get_cmd.add_argument("--range", required=True, help='A1 range, for example "Model!A1:G40".')
    get_cmd.add_argument(
        "--render",
        default="formula",
        choices=["formula", "formatted", "unformatted"],
        help="How values should be rendered when reading.",
    )
    get_cmd.add_argument(
        "--output",
        default="json",
        choices=["json", "tsv", "csv"],
        help="Output format.",
    )

    set_cmd = subparsers.add_parser("set", help="Write values/formulas into a range.")
    add_spreadsheet_arg(set_cmd)
    set_cmd.add_argument("--range", required=True, help='A1 range to update, for example "Model!B2:F20".')
    set_cmd.add_argument("--input-file", help="Optional file to load input from. Defaults to stdin.")
    set_cmd.add_argument(
        "--input-format",
        default="json",
        choices=["json", "tsv", "csv", "text"],
        help="Format of the input payload.",
    )
    set_cmd.add_argument(
        "--input-mode",
        default="USER_ENTERED",
        choices=["USER_ENTERED", "RAW"],
        help="Sheets write mode. USER_ENTERED preserves formulas and date parsing.",
    )

    append_cmd = subparsers.add_parser("append", help="Append rows to a sheet.")
    add_spreadsheet_arg(append_cmd)
    append_cmd.add_argument("--range", required=True, help='Target range, usually a tab name like "Inputs!A:Z".')
    append_cmd.add_argument("--input-file", help="Optional file to load input from. Defaults to stdin.")
    append_cmd.add_argument(
        "--input-format",
        default="json",
        choices=["json", "tsv", "csv"],
        help="Format of the input payload.",
    )
    append_cmd.add_argument(
        "--input-mode",
        default="USER_ENTERED",
        choices=["USER_ENTERED", "RAW"],
        help="Sheets append mode.",
    )

    clear_cmd = subparsers.add_parser("clear", help="Clear a range.")
    add_spreadsheet_arg(clear_cmd)
    clear_cmd.add_argument("--range", required=True, help='A1 range to clear, for example "Model!B2:F20".')

    return parser.parse_args()


def add_spreadsheet_arg(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--spreadsheet",
        required=True,
        help="Spreadsheet alias, spreadsheet ID, or full Google Sheets URL.",
    )


def main() -> int:
    args = parse_args()

    if args.command == "auth":
        creds = get_credentials(args, allow_interactive=True, force=args.force)
        print(
            json.dumps(
                {
                    "ok": True,
                    "authMode": credential_mode(args),
                    "tokenFile": str(Path(args.token_file).resolve()),
                    "hasRefreshToken": bool(getattr(creds, "refresh_token", None)),
                },
                indent=2,
            )
        )
        return 0

    service = sheets_service(args, allow_interactive=False)
    spreadsheet_id = resolve_spreadsheet(args.spreadsheet, Path(args.config))

    if args.command == "list-sheets":
        payload = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields="properties.title,sheets.properties(sheetId,title,index,gridProperties)"
        ).execute()
        print(json.dumps(payload, indent=2))
        return 0

    if args.command == "get":
        render_option = {
            "formula": "FORMULA",
            "formatted": "FORMATTED_VALUE",
            "unformatted": "UNFORMATTED_VALUE",
        }[args.render]
        payload = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=args.range,
            valueRenderOption=render_option,
        ).execute()
        emit_values(payload, args.output)
        return 0

    if args.command == "set":
        values = load_values(args.input_file, args.input_format)
        payload = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=args.range,
            valueInputOption=args.input_mode,
            body={"range": args.range, "majorDimension": "ROWS", "values": values},
        ).execute()
        print(json.dumps(payload, indent=2))
        return 0

    if args.command == "append":
        values = load_values(args.input_file, args.input_format)
        payload = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=args.range,
            valueInputOption=args.input_mode,
            insertDataOption="INSERT_ROWS",
            body={"range": args.range, "majorDimension": "ROWS", "values": values},
        ).execute()
        print(json.dumps(payload, indent=2))
        return 0

    if args.command == "clear":
        payload = service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=args.range,
            body={},
        ).execute()
        print(json.dumps(payload, indent=2))
        return 0

    raise ValueError(f"Unknown command: {args.command}")


def sheets_service(args: argparse.Namespace, allow_interactive: bool) -> Any:
    creds = get_credentials(args, allow_interactive=allow_interactive, force=False)
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def get_credentials(args: argparse.Namespace, allow_interactive: bool, force: bool) -> Any:
    service_account_path = args.service_account.strip()
    if service_account_path:
        return ServiceAccountCredentials.from_service_account_file(service_account_path, scopes=SCOPES)

    client_secret_path = args.oauth_client_secret.strip()
    if not client_secret_path:
        raise SystemExit(
            "No auth method configured.\n"
            "Set GOOGLE_SERVICE_ACCOUNT_JSON for service-account mode, or\n"
            "set GOOGLE_OAUTH_CLIENT_SECRET_FILE and run:\n"
            "  ./scripts/gsheets auth\n"
        )

    token_file = Path(args.token_file)
    creds = None
    if token_file.exists() and not force:
        creds = UserCredentials.from_authorized_user_file(str(token_file), SCOPES)

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        write_token(token_file, json.loads(creds.to_json()))
        return creds

    if not allow_interactive:
        raise SystemExit(
            "OAuth token missing or expired.\n"
            "Run this once from the repo root:\n"
            "  ./scripts/gsheets auth\n"
        )

    flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, SCOPES)
    creds = flow.run_local_server(port=0)
    write_token(token_file, json.loads(creds.to_json()))
    return creds


def write_token(token_file: Path, payload: Dict[str, Any]) -> None:
    token_file.parent.mkdir(parents=True, exist_ok=True)
    token_file.write_text(json.dumps(payload, indent=2))


def credential_mode(args: argparse.Namespace) -> str:
    if args.service_account.strip():
        return "service_account"
    if args.oauth_client_secret.strip():
        return "oauth_user"
    return "unknown"


def resolve_spreadsheet(value: str, config_path: Path) -> str:
    direct = extract_spreadsheet_id(value)
    if direct:
        return direct

    aliases = load_aliases(config_path)
    alias_target = aliases.get(value)
    if alias_target:
        resolved = extract_spreadsheet_id(alias_target)
        if resolved:
            return resolved

    raise SystemExit(
        f"Could not resolve spreadsheet '{value}'. "
        "Pass a spreadsheet ID, a Google Sheets URL, or add it to vexcel-sheets.json."
    )


def load_aliases(config_path: Path) -> Dict[str, str]:
    if not config_path.exists():
        return {}

    payload = json.loads(config_path.read_text())
    if isinstance(payload, dict):
        if isinstance(payload.get("aliases"), dict):
            return {str(k): str(v) for k, v in payload["aliases"].items()}
        return {str(k): str(v) for k, v in payload.items() if isinstance(v, str)}

    raise SystemExit(f"Alias config at {config_path} must be a JSON object.")


def extract_spreadsheet_id(value: str) -> Optional[str]:
    stripped = value.strip()
    match = SHEET_URL_RE.search(stripped)
    if match:
        return match.group(1)

    if re.fullmatch(r"[a-zA-Z0-9-_]{20,}", stripped):
        return stripped

    return None


def load_values(input_file: Optional[str], input_format: str) -> List[List[Any]]:
    raw = read_input(input_file)

    if input_format == "json":
        payload = json.loads(raw)
        if isinstance(payload, dict) and "values" in payload:
            payload = payload["values"]
        if not isinstance(payload, list):
            raise SystemExit("JSON input must be a 2D array or an object with a 'values' key.")
        return normalize_rows(payload)

    if input_format == "text":
        return [[raw.rstrip("\n")]]

    if input_format == "tsv":
        return parse_delimited(raw, delimiter="\t")

    if input_format == "csv":
        return parse_delimited(raw, delimiter=",")

    raise SystemExit(f"Unsupported input format: {input_format}")


def read_input(input_file: Optional[str]) -> str:
    if input_file:
        return Path(input_file).read_text()
    if not sys.stdin.isatty():
        return sys.stdin.read()
    raise SystemExit("No input provided. Use --input-file or pipe data over stdin.")


def parse_delimited(raw: str, delimiter: str) -> List[List[str]]:
    reader = csv.reader(io.StringIO(raw), delimiter=delimiter)
    return [row for row in reader]


def normalize_rows(rows: Iterable[Any]) -> List[List[Any]]:
    normalized: List[List[Any]] = []
    for row in rows:
        if isinstance(row, list):
            normalized.append(row)
        else:
            normalized.append([row])
    return normalized


def emit_values(payload: Dict[str, Any], output_format: str) -> None:
    values = payload.get("values", [])

    if output_format == "json":
        print(json.dumps(payload, indent=2))
        return

    delimiter = "\t" if output_format == "tsv" else ","
    writer = csv.writer(sys.stdout, delimiter=delimiter, lineterminator="\n")
    for row in values:
        writer.writerow(row)


if __name__ == "__main__":
    raise SystemExit(main())
