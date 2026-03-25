# Codex Google Sheets Bridge

Vexcel now includes a repo-local CLI bridge so Codex can read and edit live Google Sheets ranges directly instead of round-tripping through `.xlsx` downloads.

## What It Solves

Instead of:

1. generating a model locally
2. uploading it to Google Sheets
3. downloading it again for each iteration
4. re-uploading it

you can give Codex a sheet alias or spreadsheet URL and let it read or update ranges in place.

## Files

- `scripts/gsheets_cli.py`: main Python CLI
- `scripts/gsheets`: convenience wrapper that prefers the repo `.venv`
- `scripts/requirements-gsheets.txt`: Python dependencies
- `vexcel-sheets.example.json`: alias config example

## One-Time Setup

### Option 1: Service account

Best for automation and team workflows.

1. Create a Google Cloud service account with Sheets API access.
2. Download the JSON key.
3. Share the target spreadsheet with the service account email.
4. Set:
   - `GOOGLE_SERVICE_ACCOUNT_JSON=/absolute/path/to/service-account.json`

### Option 2: OAuth desktop client

Best if you want to work directly as yourself.

1. Create a Google OAuth desktop app client in Google Cloud.
2. Download the client secret JSON.
3. Set:
   - `GOOGLE_OAUTH_CLIENT_SECRET_FILE=/absolute/path/to/oauth-client-secret.json`
4. Run:
   - `./scripts/gsheets auth`

That writes the user token to `.secrets/google-sheets-token.json`.

## Alias Config

Create `vexcel-sheets.json` in the repo root:

```json
{
  "aliases": {
    "finance-model": "https://docs.google.com/spreadsheets/d/abc123/edit",
    "forecast": "spreadsheet_id_here"
  }
}
```

Then Codex can refer to `finance-model` instead of a raw spreadsheet ID.

## Common Commands

List tabs:

```bash
./scripts/gsheets list-sheets --spreadsheet finance-model
```

Read formulas from a live model:

```bash
./scripts/gsheets get \
  --spreadsheet finance-model \
  --range "Model!A1:G40" \
  --render formula
```

Write formulas or values from JSON:

```bash
cat <<'JSON' | ./scripts/gsheets set \
  --spreadsheet finance-model \
  --range "Model!B2:D4" \
  --input-format json
[
  ["=B1*1.1", "=C1*1.1", "=D1*1.1"],
  ["100", "200", "300"],
  ["=SUM(B2:B3)", "=SUM(C2:C3)", "=SUM(D2:D3)"]
]
JSON
```

Append rows:

```bash
cat data.tsv | ./scripts/gsheets append \
  --spreadsheet forecast \
  --range "Raw Data!A:Z" \
  --input-format tsv
```

Clear a range:

```bash
./scripts/gsheets clear --spreadsheet finance-model --range "Scratch!A1:Z100"
```

## Notes

- Writes default to `USER_ENTERED`, so formulas like `=SUM(B2:B10)` are interpreted correctly by Sheets.
- Formatting is not overwritten by the values API.
- The bridge is range-based on purpose: it is simple, scriptable, and easy for Codex to reason about.
