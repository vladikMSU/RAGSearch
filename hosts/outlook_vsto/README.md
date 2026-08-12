# Outlook VSTO host

Windows/.NET Framework host for classic Outlook. It owns only the add-in lifecycle,
the WinForms search pane, loopback HTTP calls, launching the Outlook MAPI connector,
and opening a selected result through Outlook Object Model.

The project, assembly, namespace and deployment identity intentionally remain
`RAGSearch`; the directory name describes its architectural role without changing
VSTO registration semantics. The source-specific ingestion implementation lives in
[`connectors/outlook_mapi`](../../connectors/outlook_mapi/README.md), and the search
service lives in [`service`](../../service/README.md).

Build from the repository root with:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1
```
