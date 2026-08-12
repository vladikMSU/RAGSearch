# Outlook Extended MAPI connector

`OutlookMapiReader.exe` — единственный executable этого connector. Это x64
read-only процесс на C++, который читает default profile classic Outlook через
Extended MAPI и потоково выдаёт одно письмо на строку в UTF-8 JSONL.

У connector нет Python adapter, HTTP-клиента, service token или зависимости от
search service. Граница компонентов выглядит так:

```text
classic Outlook profile/providers
              │ Extended MAPI
              ▼
     OutlookMapiReader.exe
              │ JSONL stdout + временные файлы вложений
              ▼
          VSTO host
              │ нейтральный HTTP document contract
              ▼
        Python search service
```

VSTO host владеет запуском и остановкой reader, читает JSONL последовательно с
backpressure, преобразует Outlook-specific record в нейтральный document и
передаёт его сервису. Python service не запускает C++ и не читает его временные
пути напрямую.

## Что делает reader

- вызывает `MAPIInitialize` и `MAPILogonEx` для default Outlook profile;
- перечисляет message stores и IPM-иерархию папок;
- emit делает только для email-like классов `IPM.Note` и `IPM.Note.*`;
- читает subject, sender, recipients, timestamps, identifiers и `PR_BODY`;
- перечисляет metadata вложений;
- по явному запросу сохраняет bounded `ATTACH_BY_VALUE` во временный каталог
  конкретного запуска;
- не использует Outlook Object Model и не содержит send/create/update/delete
  операций над MAPI objects.

Reader не парсит `.pst` или `.ost` как файлы. Данные читает настроенный MAPI store
provider установленного classic Outlook.

## Сборка

Native project включён в корневой `RAGSearch.sln` и использует MSVC v143/C++17.
Из корня репозитория:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

Или только reader:

```powershell
msbuild .\connectors\outlook_mapi\native\OutlookMapiReader.vcxproj `
  /restore /t:Rebuild /p:Configuration=Debug /p:Platform=x64
```

Outputs разделены по configuration:

```text
connectors\outlook_mapi\native\bin\x64\Debug\OutlookMapiReader.exe
connectors\outlook_mapi\native\bin\x64\Release\OutlookMapiReader.exe
```

Extended MAPI headers закреплены в `third_party/MAPIStubLibrary`; готовый EXE в
Git не хранится. Для clean clone нужен Visual Studio workload **Desktop development
with C++**. Битность reader должна совпадать с Outlook; проект намеренно только x64.

## CLI

Полный список параметров не требует Outlook logon:

```powershell
.\connectors\outlook_mapi\native\bin\x64\Debug\OutlookMapiReader.exe --help
```

Небольшой человекочитаемый запуск:

```powershell
.\connectors\outlook_mapi\native\bin\x64\Debug\OutlookMapiReader.exe `
  --max-stores 2 `
  --max-folders 10 `
  --max-messages 5 `
  --body-preview-chars 500
```

JSONL без извлечения файлов:

```powershell
.\connectors\outlook_mapi\native\bin\x64\Debug\OutlookMapiReader.exe `
  --jsonl `
  --max-stores 0 `
  --max-folders 0 `
  --max-messages 0 `
  --body-preview-chars 4000000
```

`0` означает unlimited для stores/folders/messages. Для
`--body-preview-chars` значение `0` отключает body. Диагностика и summary всегда
идут в `stderr`; при `--jsonl` в `stdout` находятся только JSON records.

## JSONL contract

Reader сохраняет Outlook-specific поля внутри connector boundary:

```json
{
  "store_id": "...",
  "store_name": "Mailbox - User",
  "entry_id": "...",
  "folder_entry_id": "...",
  "folder_path": "Mailbox - User/Inbox",
  "subject": "Quarterly plan",
  "body": "Message body",
  "body_available": true,
  "body_truncated": false,
  "attachments_truncated": false,
  "sender_name": "Alex",
  "sender_email": "alex@example.test",
  "to": "user@example.test",
  "cc": "",
  "sent_at": "2026-08-11T09:00:00.000Z",
  "received_at": "2026-08-11T09:01:00.000Z",
  "modified_at": "2026-08-11T09:02:00.000Z",
  "internet_message_id": "<id@example.test>",
  "conversation_id": "...",
  "attachments": [
    {
      "name": "report.pdf",
      "size": 4096,
      "content_type": "application/pdf",
      "temp_path": "C:\\...\\run-id\\rag_..._report.pdf"
    }
  ]
}
```

`sent_at`, `received_at` и `modified_at` — UTC ISO-8601 string либо `null`.
Остальные текстовые поля всегда strings. `temp_path` пуст для metadata-only,
unsupported attachment methods и файлов, не прошедших byte caps.

`body_truncated=true` означает, что `body` достиг переданного
`--body-preview-chars`; consumer обязан сохранить этот признак вместе с part.
`attachments_truncated=true` означает, что attachment table содержала больше
`4095` строк: reader не открывал, не сохранял и не включал в JSONL оставшиеся
вложения. Поле всегда присутствует и обычно равно `false`.

## Каталог вложений и ownership

Каталог создаёт и удаляет вызывающий процесс. Для каждого запуска нужен отдельный
непредсказуемый каталог, который не используется Python service:

```powershell
$runDir = Join-Path $env:TEMP ('RAGSearch-reader-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $runDir | Out-Null
try {
    .\connectors\outlook_mapi\native\bin\x64\Debug\OutlookMapiReader.exe `
      --jsonl `
      --max-stores 2 `
      --max-folders 100 `
      --max-messages 20 `
      --body-preview-chars 200000 `
      --attachment-dir $runDir `
      --max-attachment-bytes 8388608 `
      --max-message-attachment-bytes 8388608 `
      --max-total-attachment-bytes 67108864
}
finally {
    $resolvedRun = (Resolve-Path -LiteralPath $runDir).Path
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
    if ((Split-Path -Parent $resolvedRun) -ne $tempRoot -or
        (Split-Path -Leaf $resolvedRun) -notlike 'RAGSearch-reader-*') {
        throw "Refusing to clean unexpected path: $resolvedRun"
    }
    Remove-Item -LiteralPath $resolvedRun -Recurse -Force
}
```

Правила native writer:

- без `--attachment-dir` файловых записей нет;
- `--attachment-dir` должен быть существующим каталогом и не может быть root;
- путь канонизируется до MAPI logon;
- файл создаётся только непосредственным child этого каталога через `CREATE_NEW`;
- исходное имя превращается в безопасный basename, обрезается и получает
  уникальный prefix;
- извлекается только `ATTACH_BY_VALUE`; embedded message, OLE и by-reference
  остаются metadata-only;
- одна JSONL-запись содержит не более `4095` вложений; reader прекращает обход
  attachment table конкретного письма до открытия, сохранения или вывода лишних
  строк;
- `--max-attachment-bytes` — hard cap одного stream; `0` отключает extraction;
- `--max-message-attachment-bytes` — hard cap суммы сохранённых bytes одного
  письма, сбрасываемый перед следующим письмом; default `67108864`, а `0` отключает
  extraction;
- `--max-total-attachment-bytes` — cap одного процесса; `0` означает unlimited,
  но per-attachment cap продолжает действовать;
- частично записанный или превысивший cap файл удаляется; если удалить его не
  удалось, reader немедленно завершает run с fatal exit code `4`, чтобы не
  продолжать запись поверх уже нарушенного disk budget.

Consumer обязан повторно проверить, что непустой `temp_path` является обычным
файлом внутри его exact run directory, передать содержимое дальше с backpressure,
удалить обработанный файл и в конце очистить принадлежащий ему каталог.

## Проверка

Offline-проверка CLI:

```powershell
powershell.exe -NoProfile -File `
  .\connectors\outlook_mapi\tests\test_reader_smoke.ps1 -OfflineOnly
```

Она проверяет help contract нового per-message cap и exit code `64` для значения
выше hard ceiling, не выполняя MAPI logon.

Live smoke читает default Outlook profile напрямую, валидирует JSONL и containment
вложений; Python/service ему не нужны:

```powershell
powershell.exe -NoProfile -File `
  .\connectors\outlook_mapi\tests\test_reader_smoke.ps1
```

Можно выбрать Release, число писем и store display-name filter:

```powershell
powershell.exe -NoProfile -File `
  .\connectors\outlook_mapi\tests\test_reader_smoke.ps1 `
  -Configuration Release `
  -MaxMessages 3 `
  -StoreContains Archives
```

Smoke требует настроенный default profile и хотя бы одно доступное письмо. Он
создаёт свой временный attachment directory и удаляет только этот каталог.

## Exit codes

- `0` — проход завершён без recoverable errors;
- `1` — records выданы, но были recoverable ошибки отдельных folders/messages;
- `3` — `MAPILogonEx` не смог открыть default profile;
- `4` — fatal runtime failure;
- `64` — неверные CLI arguments.

## Ограничения

- Используется default profile, а не UI-выбор профиля.
- Не перечисляются associated/hidden contents.
- Body читается как `PR_BODY_W`, с fallback на `PR_BODY_A`; HTML/RTF отдельно не
  извлекаются.
- SMTP sender берётся из `PR_SENDER_SMTP_ADDRESS` либо
  `PR_SENT_REPRESENTING_SMTP_ADDRESS`.
- Embedded-message/OLE/by-reference attachments не разворачиваются.
- `EntryID` локален для конкретного MAPI profile/store и не является переносимым
  глобальным идентификатором.
