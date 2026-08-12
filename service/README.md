# RAGSearch local service

Dependency-free Python service for source-neutral document ingestion and hybrid
search. Producers send complete document snapshots over loopback HTTP. The service
does not know how a document was acquired and does not interpret connector locators.

It stores documents and their parts independently, creates deterministic chunks,
combines SQLite FTS5 with embeddings, and returns the producer's opaque `locator`
unchanged with every search result.

## Run

Python 3.11 or newer is required. The default embedding provider uses no downloads
and no third-party packages.

```powershell
Push-Location .\service
.\.venv\Scripts\python.exe -m ragsearch_service
Pop-Location
```

The server binds only to `127.0.0.1:8765`. On first start it creates:

- `%LOCALAPPDATA%\RAGSearch\ragsearch.sqlite3`
- `%LOCALAPPDATA%\RAGSearch\service-token`

There is no service spool. Binary part content crosses the API inline as base64; no
producer filesystem path is accepted by the service.

The token is generated atomically. On a newly created Windows data root, the
service disables inherited ACEs and grants full control to the current user, SYSTEM
and Administrators; on POSIX it applies user-only modes. The current Windows
hardening step does not remove explicit ACEs already present on a reused/custom
data directory, so verify such a directory is private before using it. Startup
fails if the ACL command itself fails. Clients send the token in
`X-RAGSearch-Token`; it is never accepted in a URL.

For an isolated development instance:

```powershell
Push-Location .\service
$devData = Join-Path $env:LOCALAPPDATA 'RAGSearch-dev'
.\.venv\Scripts\python.exe -m ragsearch_service --data-dir $devData --port 8766
Pop-Location
```

Keep custom data directories outside the checkout: every data root contains a
bearer token as well as the disposable index database.

## API

`GET /health` does not require a token and returns the compatibility contract used
by clients before they reuse a running process:

```json
{"status":"ok","protocol":4}
```

Every `/v1/*` endpoint requires `X-RAGSearch-Token`.

### Upsert one complete document snapshot

`POST /v1/documents`

The JSON body is one document directly, not a batch wrapper:

```json
{
  "source_key": "outlook_mapi:<64-hex-sha256>",
  "kind": "email",
  "title": "Quarterly launch plan",
  "metadata": {
    "sender_name": "Alex",
    "sender_email": "alex@example.test",
    "to": "user@example.test",
    "cc": "",
    "sent_at": "2026-08-11T09:00:00Z",
    "received_at": "2026-08-11T09:01:00Z",
    "modified_at": "2026-08-11T09:02:00Z",
    "folder_path": "Mailbox - User/Inbox",
    "store_name": "Mailbox - User",
    "internet_message_id": "<example@example.test>",
    "conversation_id": "conversation-id"
  },
  "locator": {
    "connector": "outlook_mapi",
    "store_id": "outlook-store-id",
    "entry_id": "outlook-entry-id",
    "folder_entry_id": "outlook-folder-id"
  },
  "parts": [
    {
      "key": "body",
      "kind": "body",
      "name": "body",
      "media_type": "text/plain",
      "size": 32,
      "text": "The launch was moved to October.",
      "truncated": false
    },
    {
      "key": "attachment:0",
      "kind": "attachment",
      "name": "notes.txt",
      "media_type": "text/plain",
      "size": 12,
      "content_base64": "SGVsbG8gd29ybGQh"
    }
  ]
}
```

Required document fields are `source_key`, `kind`, `title`, `metadata`, `locator`
and `parts`; unknown fields are rejected. `source_key` is limited to 65,536
characters and is whitespace-significant (the service does not trim producer
identity), `kind` is limited to 256, `title` to 65,536, and `parts` to 4,096 items.
`metadata` and `locator` must be JSON objects. They are serialized canonically for
storage, but their JSON value is otherwise opaque to the core. The title limit also
keeps a full 25-result response within the host's bounded response envelope.
All request strings and JSON object keys must contain Unicode scalar values;
unpaired UTF-16 surrogate code points are rejected with HTTP 400. Valid astral
characters, including emoji and JSON surrogate-pair escapes, are accepted.
Numbers inside opaque metadata/locator must fit the finite IEEE-754 `Double`
range understood by the .NET host. Large finite values such as `1e100` are
accepted; integer precision beyond `Double` precision is not promised.
Canonical metadata is limited to 1 MiB and canonical locator data to 64 KiB,
measured as UTF-8 JSON. Each object is also limited to depth 16 and 10,000 JSON
nodes. The reserved key `__type` is rejected because the .NET Framework wire
serializer interprets it as runtime type metadata; all other keys remain opaque.

Each part rejects unknown fields and requires a non-empty `key` (at most 32,768
characters) and `kind` (at most 256); part keys must be unique within one document.
Optional fields are `name`
(32,768), `media_type` (4,096), `size`,
`truncated`, `text` and `content_base64`. `size` must be an integer from zero through
2^63-1; it may be omitted for metadata-only or producer-text parts, but is
effectively required with non-empty `content_base64` because it must equal the
decoded byte length. `truncated`, when present, must be boolean. `text` and
`content_base64` are mutually exclusive. Omitting both stores part metadata without
extracted text.

The default decoded binary limit is 8 MiB per part. The HTTP request limit is
48 MiB measured over the fully serialized UTF-8 body; the VSTO host applies the
same byte limit before sending. Text, HTML,
JSON/XML/CSV/log/Markdown and DOCX content is
extracted with the standard library. Unsupported formats are recorded as
`unsupported` but their document metadata remains searchable. Searchable text
extracted from one binary part is capped at 8 MiB characters, and all
binary-derived text in one document shares the same 8 MiB aggregate budget. The
synthesized metadata search projection is capped at 4 MiB characters without
changing stored metadata or part rows. Producer-supplied `text` is not cropped, but
it is deterministically normalized for chunk indexing (line endings, NULs and
horizontal whitespace); the original text itself is not stored, only its hash and
normalized chunks. Producer text is indexed first and returns HTTP 400 if it alone
exceeds a hard per-document limit; derived metadata and binary projections are then
deterministically cropped to the remaining capacity instead of rejecting the
document. The hard limits are 16 Mi characters and 16,384 chunks. Search queries
are limited to 8,192 characters. Together these bounds prevent a valid request
from expanding into unbounded embedding work while keeping the Outlook host
envelope admissible.

Response:

```json
{"source_key":"outlook_mapi:<64-hex-sha256>","status":"upserted"}
```

`source_key` is the sole public identity. Retrying it updates the document and
replaces only that document's parts and chunks, so removed parts do not remain in
the index. Schema/validation errors return HTTP 400; an HTTP body above 48 MiB
returns HTTP 413 `request_too_large`. The removed `/v1/messages` endpoint has no
compatibility alias and returns HTTP 404.

### Search

`POST /v1/search`

```json
{"query":"when was the product launch moved","limit":25}
```

Only `query` and `limit` are accepted. `query` must be non-empty and no longer than
8,192 characters; `limit` must be an integer from 1 through 100. The service still
returns at most its configured hard cap of 25 results. Source-specific filters are
deliberately not part of the neutral core contract.

The response contains generic document fields and the opaque producer data:

The current Outlook host bounds its producer title to 65,536 characters and reads
this response as strict UTF-8 through a 128 MiB streaming/parser cap.

```json
{
  "results": [
    {
      "source_key": "outlook_mapi:<64-hex-sha256>",
      "kind": "email",
      "title": "Quarterly launch plan",
      "metadata": {
        "sender_name": "Alex",
        "sender_email": "alex@example.test",
        "to": "user@example.test",
        "cc": "",
        "sent_at": "2026-08-11T09:00:00Z",
        "received_at": "2026-08-11T09:01:00Z",
        "modified_at": "2026-08-11T09:02:00Z",
        "folder_path": "Mailbox - User/Inbox",
        "store_name": "Mailbox - User",
        "internet_message_id": "<example@example.test>",
        "conversation_id": "conversation-id"
      },
      "locator": {
        "connector": "outlook_mapi",
        "store_id": "outlook-store-id",
        "entry_id": "outlook-entry-id",
        "folder_entry_id": "outlook-folder-id"
      },
      "rank": 1,
      "snippet": "The launch was moved to October.",
      "snippet_part": "body",
      "matched_parts": ["body", "attachment:notes.txt"],
      "hybrid_score": 0.7586,
      "lexical_score": 0.74,
      "lexical_match_kind": "token",
      "vector_similarity": 0.77,
      "vector_distance": 0.23,
      "ranking_basis": "lexical_token"
    }
  ],
  "total": 1,
  "mode": "hybrid-semantic",
  "candidate_count": 1,
  "eligible_count": 1,
  "lexical_match_count": 1,
  "lexical_gate": false,
  "best_vector_similarity": 0.77,
  "best_vector_distance": 0.23,
  "cutoff_similarity": 0.67,
  "cutoff_distance": 0.33,
  "max_results": 25,
  "ranking": "lexical_gate_then_vector_distance_asc"
}
```

Results are aggregated per document. Literal retrieval uses both the FTS5
`unicode61` token index and a trigram index. The ranking policy uses
prefix/substring evidence to enable the lexical gate, requires literal evidence for
single-token queries, and applies an adaptive model-specific cosine cutoff to
multi-token semantic results. `snippet_part` identifies the
exact chunk label used to produce `snippet`; `matched_parts` contains up to eight
distinct matched labels for the document and is therefore a bounded summary, not
an exhaustive list. The service returns at most 25 documents.

### Stats and reset

`GET /v1/stats` returns `documents`, `parts`, `chunks`, `part_extraction`, schema
version, database size, and the active embedding model contract.

`DELETE /v1/index` atomically removes all documents, parts, chunks, FTS entries and
embedding contract metadata. It keeps the database file, schema and token. Response:

```json
{"deleted_documents":42,"deleted_parts":49,"deleted_chunks":1877}
```

Calling it again is safe and returns zero counters. Schema v4 is a clean break;
older databases are rejected at startup and should be deleted and re-indexed.

## Embeddings

The default `hashing-v1-256` provider hashes word, bigram and character features
into a normalized deterministic vector. An optional locally cached
sentence-transformers model can be selected explicitly:

```powershell
Push-Location .\service
.\.venv\Scripts\python.exe -m ragsearch_service `
  --embedding sentence-transformers `
  --model .\models\paraphrase-multilingual-MiniLM-L12-v2
Pop-Location
```

`local_files_only=True` is enforced. The first ingestion fixes one provider name,
dimension and immutable fingerprint for the whole index. Switch models only after
`DELETE /v1/index`, then re-ingest every document.

## Tests

```powershell
Push-Location .\service
.\.venv\Scripts\python.exe -B -m unittest discover -s tests -t . -v
Pop-Location
```

Tests use explicit temporary data and token paths and do not touch default
`%LOCALAPPDATA%\RAGSearch` state.

## Scale note

The standard-library vector path performs a streaming brute-force cosine scan with
a bounded top-k heap. It keeps memory bounded and is appropriate for the local
prototype. A large archive will eventually need an ANN-backed implementation behind
the same search interface; SQLite FTS5 remains independent.
