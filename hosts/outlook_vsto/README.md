# Outlook VSTO host

Windows/.NET Framework host for classic Outlook. It owns only the add-in lifecycle,
the WinForms search pane, loopback HTTP calls, launching the Outlook MAPI connector,
and opening a selected result through Outlook Object Model.

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

Build from the repository root with:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1
```
