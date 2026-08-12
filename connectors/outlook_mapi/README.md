# Outlook MAPI connector

Connector объединяет Python-адаптер `adapter.py` и отдельный x64 console reader
`native/OutlookMapiReader.vcxproj`. Reader потоково читает Outlook stores через
**Extended MAPI**, без Outlook Object Model и без прямого чтения заблокированного
`.ost`; adapter валидирует JSONL и передаёт письма локальному search service по HTTP.

Reader:

- вызывает `MAPIInitialize` и `MAPILogonEx` для default Outlook profile;
- перечисляет message stores и всю IPM-иерархию папок, но emit делает только для
  email-like message classes `IPM.Note` и `IPM.Note.*`;
- открывает каждое сообщение read-only и отдаёт identity, store/folder context,
  Subject/Body, sender/recipient display metadata, даты, Internet Message ID,
  conversation identity и metadata вложений;
- умеет опционально читать `ATTACH_BY_VALUE` в явно переданный bounded spool;
- не содержит MAPI `SetProps`, `SaveChanges`, `Create*`, `Delete*`,
  `SubmitMessage`, `MAPI_MODIFY` или иных send/write операций. Единственная запись
  на диск — новые уникальные файлы внутри явно переданного `--spool-dir`;
- использует безопасные конечные лимиты по умолчанию, чтобы первый scan не начал
  обходить многолетний mailbox целиком.

Это не parser файлов PST/OST. Доступ выполняется через установленный MAPI provider
Outlook, поэтому Outlook может быть открыт, а блокировка `.ost` процессом Outlook не
имеет значения.

## Требования

- Classic Outlook for Windows с настроенным default profile. New Outlook Extended
  MAPI не предоставляет.
- Bitness приложения должна совпадать с bitness Outlook/MAPI. Этот проект намеренно
  только `x64`.
- Visual Studio 2022: workload **Desktop development with C++**, MSVC v143 и Windows
  10/11 SDK.
- Минимальный набор Extended MAPI headers из официального Microsoft
  MAPIStubLibrary уже зафиксирован в
  `..\..\third_party\MAPIStubLibrary\include`. Точный upstream commit и MIT-лицензия
  записаны рядом в `README.md` и `LICENSE`.

MSBuild также копирует `THIRD_PARTY_NOTICES.md` и
`MAPIStubLibrary-LICENSE.txt` рядом с EXE. Сохраняйте их в бинарной поставке.

Microsoft отдельно требует совпадения bitness MAPI application и установленного
Outlook: <https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/building-mapi-applications-on-32-bit-and-64-bit-platforms>.

Команда сборки выполняется из корня репозитория; команды прямого запуска — из
`connectors\outlook_mapi`.

## Сборка

Reader и VSTO host входят в одну solution и используют одну конфигурацию:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 `
  -Configuration Debug
```

Артефакты reader разделены по конфигурациям и не перезаписывают друг друга:

```text
connectors\outlook_mapi\native\bin\x64\Debug\OutlookMapiReader.exe
connectors\outlook_mapi\native\bin\x64\Release\OutlookMapiReader.exe
```

`scripts/build.ps1` требует Visual Studio 2022 с MSBuild, Office development и
Desktop development with C++. Альтернативных CMake, include override и direct
compiler путей проект не поддерживает.

## Запуск

Первый короткий positive-control:

```powershell
.\native\bin\x64\Debug\OutlookMapiReader.exe `
  --max-stores 0 `
  --max-folders 25 `
  --max-messages 5 `
  --body-preview-chars 200
```

Для stores/folders/messages `0` означает unlimited. Полный обход потенциально очень
дорогой, поэтому его стоит делать только осознанно:

```powershell
.\native\bin\x64\Debug\OutlookMapiReader.exe `
  --max-folders 0 `
  --max-messages 0 `
  --body-preview-chars 500
```

Можно ограничить proof одним store без изменения профиля, например PST с display
name `Archives`:

```powershell
.\native\bin\x64\Debug\OutlookMapiReader.exe `
  --store-contains Archives `
  --max-folders 25 `
  --max-messages 5
```

Reader логинится без `MAPI_LOGON_UI`; если default profile отсутствует или требует
интерактивного выбора, он завершится с HRESULT вместо показа profile picker.

### JSONL для Python pipeline

`--jsonl` оставляет на `stdout` только UTF-8 JSON Lines — один объект на сообщение.
Progress, recoverable errors и итоговый summary направляются в `stderr`, поэтому
stdout можно безопасно перенаправить в файл или читать построчно без concat всего
mailbox:

```powershell
.\native\bin\x64\Debug\OutlookMapiReader.exe `
  --jsonl `
  --max-folders 100 `
  --max-messages 1000 `
  --body-preview-chars 200000 `
  --spool-dir C:\Users\User\AppData\Local\RAGSearch\spool `
  --max-attachment-bytes 67108864 `
  --max-total-attachment-bytes 0 `
  2> reader.log `
  > messages.jsonl
```

Формат строки (показан развёрнуто только для документации; фактический output —
одна строка на сообщение):

```json
{
  "store_id": "...",
  "store_name": "Mailbox - User",
  "entry_id": "...",
  "folder_entry_id": "...",
  "folder_path": "Mailbox - User/Inbox",
  "subject": "...",
  "body": "...",
  "body_available": true,
  "body_truncated": false,
  "sender_name": "Sender",
  "sender_email": "sender@example.test",
  "to": "Recipient One; Recipient Two",
  "cc": "Copy Recipient",
  "sent_at": "2026-08-11T12:34:56.789Z",
  "received_at": "2026-08-11T12:35:01.123Z",
  "modified_at": null,
  "internet_message_id": "<id@example.test>",
  "conversation_id": "A1B2...",
  "attachments": [
    {
      "name": "report.pdf",
      "size": 12345,
      "content_type": "application/pdf",
      "temp_path": "C:\\...\\spool\\rag_..._report.pdf"
    }
  ]
}
```

`--body-preview-chars` в текущем контракте одновременно является пределом поля `body`;
hard maximum — `4000000` символов (примерно 8 MiB UTF-16 read buffer).
Для production ingestion лучше оставить bounded producer/consumer с backpressure,
а не накапливать весь stdout в памяти вызывающего процесса.

Даты всегда имеют ISO-8601 UTC вид с суффиксом `Z` либо равны JSON `null`.
`conversation_id` содержит только `PR_CONVERSATION_ID`; если provider его не
вернул, поле пустое. Binary value кодируется hex.
`to`/`cc` — provider display strings (`PR_DISPLAY_TO`/`PR_DISPLAY_CC`), не
нормализованный список SMTP-адресов.

Message-class filter выполняется по `PR_MESSAGE_CLASS_W/A` ещё в contents table,
до дорогого `OpenEntry`. Точное `IPM.Note` и prefix `IPM.Note.` сохраняют обычные,
custom-form и S/MIME письма. `REPORT.*` и `IPM.Schedule.Meeting.*` намеренно
исключены: Outlook Object Model материализует их как `ReportItem`/`MeetingItem`, а
предыдущий индексатор принимал только успешный cast `item as Outlook.MailItem`.
Пропуски не расходуют `--max-messages` и отражаются в summary как
`skipped_non_mail`.

### Python adapter

Из корня репозитория adapter можно запустить отдельно от VSTO host:

```powershell
.\service\.venv\Scripts\python.exe .\connectors\outlook_mapi\adapter.py `
  --executable .\connectors\outlook_mapi\native\bin\x64\Debug\OutlookMapiReader.exe `
  --full-scan
```

Adapter читает JSONL потоково, не накапливая mailbox в памяти. Для каждого запуска
он создаёт собственный помеченный каталог ниже `--spool-dir`, принимает только
обычные файлы внутри этого каталога и удаляет только принадлежащий ему run.
Машиночитаемый статус выдаётся строками `RAGSEARCH_PROGRESS {json}`; внешний
`--cancel-file` позволяет запросить корректную остановку.

### Тесты

Из корня репозитория:

```powershell
.\service\.venv\Scripts\python.exe -B -m unittest `
  discover -s connectors\outlook_mapi\tests -t . -v

.\service\.venv\Scripts\python.exe -B `
  .\connectors\outlook_mapi\adapter.py --help
```

Live E2E дополнительно читает реальный Outlook profile и ожидает подготовленный
test store, поэтому запускается отдельно:

```powershell
powershell.exe -NoProfile -File `
  .\connectors\outlook_mapi\tests\test_adapter_e2e.ps1
```

### Безопасность spool вложений

- Без `--spool-dir` producer только перечисляет metadata и вообще не создаёт файлов.
- Извлекается только `ATTACH_BY_VALUE`; embedded messages, OLE и by-reference
  attachments остаются с `temp_path: ""`.
- Имя превращается в безопасный basename, обрезается, получает уникальный prefix;
  файл создаётся через `CREATE_NEW`, поэтому существующий файл не перезаписывается.
- Spool создаётся/канонизируется до MAPI logon, filesystem root отклоняется, а каждый
  target повторно проверяется как непосредственный child канонического spool.
- `--max-attachment-bytes` — hard cap одного фактического stream. `0` отключает
  extraction; default `67108864` (64 MiB).
- `--max-total-attachment-bytes` — cap одного запуска. `0` означает unlimited total,
  но per-attachment cap продолжает действовать. Python consumer обязан читать с
  backpressure и удалять уже обработанные spool-файлы.
- Oversized, unsupported или превысившее total cap вложение всё равно присутствует
  в metadata, но имеет `temp_path: ""`. Частично записанный файл удаляется.

## Что считается успешным экспериментом

1. В консоли появились stores, folder paths, Subject/Body preview, EntryID и StoreID.
2. Outlook Object Model Guard не показал окно согласия.
3. Outlook продолжил нормально работать параллельно с reader.

Extended MAPI — отдельный unmanaged API, а не обход/отключение настроенной политики
Programmatic Access. Итоговое поведение всё равно зависит от доступности MAPI,
настроек профиля, прав пользователя и корпоративных ограничений. Обзор выбора API:
<https://learn.microsoft.com/en-us/office/client-developer/outlook/selecting-an-api-or-technology-for-developing-solutions-for-outlook>.

## Результат проверки на этой машине (2026-08-12)

После clean-break 12 августа 2026 года новый reader собран в раздельные
`native\bin\x64\Debug` и `native\bin\x64\Release`; connector tests прошли 14/14,
VSTO Debug Rebuild — с 0 warnings и 0 errors.

- Исходник без warnings скомпилирован и слинкован как x64 console executable.
- `dumpbin /headers` подтвердил `machine (x64)` и `Windows CUI`; imports содержат
  `MAPI32.dll`.
- После расширения контракта bounded Exchange/OST run прочитал 3 folders и 3
  messages, вернул все обязательные поля и завершился с exit code `0` без ошибок.
- Отдельный bounded `--store-contains Archives` run прочитал 3 PST folders и 12
  messages, перечислил 39 attachments и сохранил 39 by-value streams (1 818 190
  bytes) в отдельный test spool. Все paths находились внутри spool, существовали и
  совпали с заявленным `size`; exit code `0`.
- Никакого Outlook Object Model/COM в процессе нет, и blocking consent dialog во
  время запуска не возник.
- `--jsonl` отдельно распарсен для Exchange и PST: обязательные поля и boolean
  truncation flags присутствуют, diagnostics/summary остаются только в `stderr`.
- Full metadata-only traversal (`stores/folders/messages=0`, без spool) после
  message-class filter завершился за 3.59 s: 2 stores, 41 folders, ровно 42 emitted
  mail messages (`Exchange/OST=16`, `Archives/PST=26`), 31 non-mail objects skipped,
  49 attachment metadata records, 0 recoverable errors. Это совпало с OOM baseline
  42 вместо прежних загрязнённых 73 MAPI message objects.
- Контроль `--max-messages 20` вернул ровно 20 писем, хотя до них/между ними были
  пропущены все 31 non-mail objects: skipped rows лимит emitted mail не уменьшают.

Штатный Visual Studio workload Desktop development with C++ установлен.
Канонический полный `Debug|x64` Rebuild решения и отдельный `Release|x64` Rebuild
reader выполнены через MSVC v143 с 0 warnings и 0 errors. Для воспроизводимой
сборки clean clone нужен заявленный C++ workload.

## Ограничения

- Используется default profile, а не UI-выбор профиля. Если Outlook запущен с
  не-default profile, это может быть другой mailbox.
- Не перечисляются associated/hidden contents.
- Body читается как `PR_BODY_W` stream; если provider публикует только ANSI body,
  используется `PR_BODY_A`. HTML/RTF отдельно не извлекаются.
- SMTP sender берётся только из `PR_SENDER_SMTP_ADDRESS` либо
  `PR_SENT_REPRESENTING_SMTP_ADDRESS`; provider-specific EX address не маскируется
  под SMTP.
- Вложения embedded-message/OLE/by-reference не разворачиваются; для них остаются
  metadata и пустой `temp_path`.
- `EntryID` локален для конкретного MAPI profile/store и не является переносимым
  глобальным идентификатором.
- Ошибки отдельных inaccessible folders/messages считаются recoverable: reader идёт
  дальше и возвращает exit code `1`; fatal init/logon errors имеют другие exit codes.
