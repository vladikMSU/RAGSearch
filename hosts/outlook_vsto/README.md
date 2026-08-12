# Outlook VSTO host

Windows/.NET Framework host for classic Outlook. It owns the add-in lifecycle, the
WinForms search pane, loopback HTTP calls, startup of the local Python service and
Outlook MAPI connector, and opening a selected result through Outlook Object Model.

The project, assembly, namespace and deployment identity intentionally remain
`RAGSearch`; the directory name describes its architectural role without changing
VSTO registration semantics. The source-specific ingestion implementation lives in
[`connectors/outlook_mapi`](../../connectors/outlook_mapi/README.md), and the search
service lives in [`service`](../../service/README.md).

The host is the only ingestion orchestrator. It starts `OutlookMapiReader.exe`,
reads one strict UTF-8 JSONL record at a time, maps Outlook fields into the neutral
document contract, awaits `POST /v1/documents`, and only then reads the next record.
The reader and service never call or configure each other. Attachment files remain
in a host-owned per-run directory, with an 8 MiB per-document inline budget; no
producer path crosses HTTP. Before a document is sent, the host measures the final
UTF-8 JSON body against the service's 48 MiB request limit. The health handshake
requires service protocol `4`. Outlook subjects are capped at 65,536 characters,
and service responses are streamed through a strict UTF-8 128 MiB client cap before
the search JSON envelope is parsed.

## Development layout and service lifecycle

This is currently a workspace build, not an installed standalone package. The host
reads its registered manifest from
`HKCU\Software\Microsoft\Office\Outlook\Addins\RAGSearch`, expects that manifest
below `hosts\outlook_vsto\bin\<Configuration>`, and walks four parent directories
to the repository root. It then requires:

- `RAGSearch.sln`;
- `service\.venv\Scripts\python.exe`;
- `service\ragsearch_service\__main__.py`.

Ingestion additionally requires
`connectors\outlook_mapi\native\bin\x64\<Configuration>\OutlookMapiReader.exe`;
search and index reset can start/use the service without the native executable.

The reader configuration follows the host build: Debug host uses Debug reader and
Release host uses Release reader. If the exact local directory
`service\models\paraphrase-multilingual-MiniLM-L12-v2` exists, the host starts the
service with the optional sentence-transformers provider; otherwise it uses the
dependency-free hashing provider.

Before search, reset or ingestion, the host probes public `GET /health` without a
token and requires protocol `4`. If the probe cannot reach a healthy process on
`127.0.0.1:8765`, it starts `python -m ragsearch_service` from the service directory
and waits up to 90 seconds. A healthy response with another protocol is reported as
an incompatibility instead; the host does not start a competing process on that
port. A service started by the host is deliberately left running when Outlook or
the add-in closes so it can serve the next session. The host never terminates a
pre-existing service.

## HTTP, cancellation and opening results

All `/v1/*` calls use `X-RAGSearch-Token` read from
`%LOCALAPPDATA%\RAGSearch\service-token`:

- `POST /v1/documents` receives one complete mapped record at a time;
- `POST /v1/search` searches the whole local index (the UI requests at most 25);
- `DELETE /v1/index` clears the disposable document index after confirmation.

The **Stop** command cancels the current ingestion HTTP operation and terminates
only the owned native reader. The reader is assigned to a Windows Job configured for
kill-on-close, with direct process termination as fallback. Its stderr is drained
without being shown or persisted because diagnostics can contain mailbox names,
paths and MAPI identifiers. Per-record attachment files and the host-owned run
directory are deleted best-effort after each successful upsert and in the run's
`finally` cleanup. Deletion/ACL failures are suppressed, so an interrupted or
locked run can leave files under `%LOCALAPPDATA%\RAGSearch\reader-runs` for manual
inspection/removal.

Opening a row is intentionally Outlook-specific. The host accepts only
`locator.connector == "outlook_mapi"`, requires non-empty `store_id` and `entry_id`,
calls `Application.Session.GetItemFromID(entryId, storeId)` on Outlook's UI/STA
thread, verifies the result is a `MailItem`, and calls `Display(false)`. It does not
change the current folder, Outlook Search bar or central message list.

## Build

Debug build from the repository root:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

The script may create or reuse a non-exportable self-signed development certificate
only for Debug. Release deliberately requires an explicit installed production
code-signing certificate with a private key:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 `
  -Configuration Release `
  -CertificateThumbprint YOUR_CERTIFICATE_THUMBPRINT
```

Both commands build the x64 solution, including the native reader. Full
prerequisites and the signing/installation limitations are documented in the
[root README](../../README.md).
