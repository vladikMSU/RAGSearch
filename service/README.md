# RAGSearch local service

Dependency-free Python service for Outlook message ingestion and message-level hybrid search. It stores each message and attachment independently, uses deterministic chunks, combines SQLite FTS5 with embeddings, and always returns Outlook navigation identity (`entry_id`, `store_id`, `folder_entry_id`). It never concatenates a mailbox into one document.

## Run

Python 3.11 or newer is required. The default provider uses no downloads and no third-party packages.
The default Windows data layout requires a non-empty `LOCALAPPDATA`; use
`--data-dir` for an explicit isolated location.

```powershell
Push-Location .\service
.\.venv\Scripts\python.exe -m ragsearch_service
Pop-Location
```

The server binds only to `127.0.0.1:8765`. On first start it creates:

- `%LOCALAPPDATA%\RAGSearch\ragsearch.sqlite3`
- `%LOCALAPPDATA%\RAGSearch\service-token`
- `%LOCALAPPDATA%\RAGSearch\spool\`

The token is generated atomically. On Windows the service removes inherited ACEs from its data/spool roots and grants full control only to the current user, SYSTEM and Administrators; on POSIX it applies user-only modes. Startup fails instead of silently continuing if the Windows ACL cannot be hardened. The VSTO add-in reads the token and sends it in `X-RAGSearch-Token`; the token is never accepted in a URL.

For an isolated development instance:

```powershell
Push-Location .\service
.\.venv\Scripts\python.exe -m ragsearch_service --data-dir .\dev-data --port 8766
Pop-Location
```

## API

`GET /health` is deliberately minimal and does not require a token. Every `/v1/*` endpoint requires `X-RAGSearch-Token`.

### Ingest complete message snapshots

`POST /v1/messages`

```json
{
  "messages": [
    {
      "entry_id": "outlook-entry-id",
      "store_id": "outlook-store-id",
      "folder_entry_id": "outlook-folder-id",
      "folder_path": "\\Mailbox - User\\Inbox",
      "store_name": "Mailbox - User",
      "subject": "Quarterly launch plan",
      "sender_name": "Alex",
      "sender_email": "alex@example.test",
      "to": "user@example.test",
      "cc": "",
      "sent_at": "2026-08-11T09:00:00Z",
      "received_at": "2026-08-11T09:01:00Z",
      "modified_at": "2026-08-11T09:02:00Z",
      "internet_message_id": "<example@example.test>",
      "conversation_id": "conversation-id",
      "body": "The launch was moved to October.",
      "attachments": [
        {
          "name": "notes.txt",
          "size": 4096,
          "content_type": "text/plain",
          "temp_path": "C:\\Users\\user\\AppData\\Local\\RAGSearch\\spool\\notes.txt"
        }
      ]
    }
  ]
}
```

Response:

```json
{"accepted":1,"failed":0,"errors":[]}
```

Identity is the composite `(store_id, entry_id)`. Retrying the same item updates it, replaces only that message's chunks/attachments, and removes stale attachments. A bad item does not roll back other items in the batch.

`to` and `cc` are a single canonical string or `null`; recipient arrays are not
accepted. Each timestamp is either `null` or a non-empty ISO-8601 string with a
`T` separator and explicit UTC offset (`Z` is accepted). Naive and empty values
are rejected. The service converts accepted timestamps to fixed-width UTC such
as `2026-08-11T09:01:00.000000Z` before storing or returning them.

Attachment paths are resolved (including symlinks) and rejected unless they remain inside the configured spool. Text, HTML, JSON/XML/CSV/log/Markdown and DOCX are extracted with the standard library. Unsupported formats remain searchable by message metadata but are recorded as `unsupported`. The default extraction limit is 64 MiB per attachment. Spool files are retained unless `--delete-spool-after-ingest` is explicitly set; with that flag deletion occurs only after commit.

### Search

`POST /v1/search`

```json
{
  "query": "when was the product launch moved",
  "limit": 20,
  "filters": {
    "store_id": "outlook-store-id",
    "folder_path_prefix": "\\Mailbox - User\\Inbox",
    "received_from": "2026-01-01T00:00:00Z",
    "has_attachments": true
  }
}
```

Supported filters: `store_id`, `store_ids`, `folder_entry_id`, `folder_path`, `folder_path_prefix`, `sender_email`, `received_from`, `received_to`, `has_attachments`.
`received_from` and `received_to` use the same strict timestamp grammar and are
normalized to UTC before their inclusive comparisons. Omitting `filters` searches
every indexed OST/PST store.

Response fields are stable for VSTO navigation:

```json
{
  "results": [
    {
      "entry_id": "outlook-entry-id",
      "store_id": "outlook-store-id",
      "folder_entry_id": "outlook-folder-id",
      "subject": "Quarterly launch plan",
      "sender_name": "Alex",
      "sender_email": "alex@example.test",
      "received_at": "2026-08-11T09:01:00.000000Z",
      "folder_path": "\\Mailbox - User\\Inbox",
      "hybrid_score": 0.91,
      "lexical_score": 0.74,
      "lexical_match_kind": "token",
      "vector_similarity": 0.77,
      "vector_distance": 0.23,
      "rank": 1,
      "ranking_basis": "lexical_token",
      "snippet": "The launch was moved to October.",
      "matched_sources": ["body", "attachment:notes.txt"]
    }
  ],
  "mode": "hybrid-semantic",
  "candidate_count": 42,
  "eligible_count": 7,
  "lexical_match_count": 1,
  "lexical_gate": false,
  "cutoff_similarity": 0.67,
  "cutoff_distance": 0.33,
  "max_results": 25,
  "ranking": "lexical_gate_then_vector_distance_asc"
}
```

The service aggregates the best chunk per Outlook message before selecting the
message candidate pools. `vector_similarity` is the best indexed chunk
cosine and `vector_distance = 1 - vector_similarity`. For multi-token queries,
literal matches are followed by semantic candidates ordered by distance, using
an adaptive cutoff (`max(model floor, best similarity - 0.10)`) and a hard cap of
25 messages. The dense-model floor is `0.40`; the hashing provider uses `0.30`
because cosine distributions are model-specific.

Literal retrieval uses both the `unicode61` token index and an FTS5 `trigram`
index. `lexical_match_kind` is `token`, `prefix`, `substring`, or empty. A
prefix/substring hit enables `lexical_gate` and excludes unrelated dense guesses;
for a single-token query any literal hit wins, while a single token with no
literal evidence returns `mode=single-token-no-literal`. This avoids the observed
short-query pathology where the multilingual paraphrase model ranked an unrelated
four-letter body above `киберспорт`. Existing databases must already use schema v3;
older schemas are rejected at startup and must be rebuilt. `DELETE /v1/index`
clears both FTS indexes. Sources remain `message_metadata`, `body`, or
`attachment:<name>`.

### Stats

`GET /v1/stats` returns message, attachment and chunk counts, extraction status
counts, schema version, and the active embedding model, dimension, and immutable
implementation/artifact fingerprint.

### Clear the local index

`DELETE /v1/index` atomically removes all messages, attachments, chunks, FTS
entries, and embedding contract metadata from the local database. It keeps the
database file, schema, service token, and spool files unchanged, and does not run
`VACUUM`. Clearing the embedding metadata is deliberate: the next ingestion fixes
a new model contract for the now-empty index. The endpoint requires
`X-RAGSearch-Token` like every other `/v1/*` route.

Response:

```json
{"deleted_messages":42,"deleted_attachments":49,"deleted_chunks":1877}
```

Calling it again is safe and returns zero counts.

## Embeddings

The default `hashing-v1-256` provider hashes word, bigram and character features into a normalized deterministic vector. It is lightweight and reproducible but is not a neural language model.

An optional locally cached sentence-transformers model can be selected explicitly:

```powershell
Push-Location .\service
.\.venv\Scripts\python.exe -m pip install "sentence-transformers==5.7.0"
.\.venv\Scripts\python.exe -m ragsearch_service `
  --embedding sentence-transformers `
  --model path\to\local-model
Pop-Location
```

`local_files_only=True` is enforced, so starting the service does not download a
model. The first ingestion fixes one provider name, dimension, and immutable
fingerprint for the whole index. The hashing fingerprint covers its feature
algorithm and Unicode database version. The neural fingerprint hashes every file
in the resolved local model snapshot plus the embedding runtime package versions.
Replacing model weights at the same path therefore cannot silently mix vectors.
Starting against a non-empty index with another contract fails immediately;
switch models only after `DELETE /v1/index`, then re-ingest every message.

The conventional ignored workspace location for the multilingual model is:

```text
<repo-root>\service\models\paraphrase-multilingual-MiniLM-L12-v2
```

Run the neural configuration from the repository root:

```powershell
Push-Location .\service
.\.venv\Scripts\python.exe -m ragsearch_service `
  --embedding sentence-transformers `
  --model .\models\paraphrase-multilingual-MiniLM-L12-v2 `
  --delete-spool-after-ingest
Pop-Location
```

The model is loaded with `local_files_only=True`; `service/models/` is intentionally ignored rather than committed. Its model card documents 50-language, 384-dimensional sentence embeddings: [paraphrase-multilingual-MiniLM-L12-v2](https://huggingface.co/sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2).

## Outlook MAPI connector

Source-specific ingestion is intentionally outside the search service, in
[`connectors/outlook_mapi`](../connectors/outlook_mapi/README.md). Its adapter streams
the reader's JSONL contract into this service through the public `/v1/messages` HTTP
boundary.

## Tests

```powershell
Push-Location .\service
.\.venv\Scripts\python.exe -m unittest discover -s tests -t . -v
Pop-Location
```

Tests use explicit temporary data/token/spool directories and do not touch the default `%LOCALAPPDATA%\RAGSearch` state.

The service tests are independent from the connector tests and use only explicit
temporary paths. Connector validation commands are documented in the connector
README.

## Scale note

The stdlib vector path performs a streaming brute-force cosine scan with a bounded top-k heap. It keeps memory bounded and is appropriate for the local prototype, but a multi-year 100 GB archive will eventually need an ANN-backed provider (for example sqlite-vec) behind the same search interface. SQLite FTS5 and all metadata filters are indexed independently.
