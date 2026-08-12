# Native Extended MAPI probe

Отдельный x64 console producer для bounded streaming-чтения Outlook stores через
**Extended MAPI**, без Outlook Object Model и без прямого чтения заблокированного `.ost`.

Probe:

- вызывает `MAPIInitialize` и `MAPILogonEx` для default Outlook profile;
- перечисляет message stores и всю IPM-иерархию папок, но emit делает только для
  email-like message classes `IPM.Note` и `IPM.Note.*`;
- открывает каждое сообщение read-only и отдаёт identity, store/folder context,
  Subject/Body, sender/recipient display metadata, даты, Internet Message ID,
  conversation/search identity и metadata вложений;
- умеет опционально читать `ATTACH_BY_VALUE` в явно переданный bounded spool;
- не содержит MAPI `SetProps`, `SaveChanges`, `Create*`, `Delete*`,
  `SubmitMessage`, `MAPI_MODIFY` или иных send/write операций. Единственная запись
  на диск — новые уникальные файлы внутри явно переданного `--spool-dir`;
- использует безопасные конечные лимиты по умолчанию, чтобы первый probe не начал
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
  `..\third_party\MAPIStubLibrary\include`. Точный upstream commit и MIT-лицензия
  записаны рядом в `README.md` и `LICENSE`.

MSBuild/CMake также копирует `THIRD_PARTY_NOTICES.md` и
`MAPIStubLibrary-LICENSE.txt` рядом с EXE. Сохраняйте их в бинарной поставке.

Microsoft отдельно требует совпадения bitness MAPI application и установленного
Outlook: <https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/building-mapi-applications-on-32-bit-and-64-bit-platforms>.

## Сборка через MSBuild

Из **x64 Native Tools Command Prompt for VS 2022**:

```powershell
msbuild .\NativeMapiProbe.vcxproj /m `
  /p:Configuration=Release `
  /p:Platform=x64
```

Бинарник:

```text
build-direct\NativeMapiProbe.exe
```

## Сборка через CMake

```powershell
cmake -S . -B build -A x64
cmake --build build --config Release
```

Бинарник для Visual Studio generator:

```text
build-direct\NativeMapiProbe.exe
```

### Direct build, использованный на этой машине

На машине нет штатного C++ workload, но есть bundled x64 ScopeCppSDK. Проверенный
эквивалентный Release build выполнен с headers из репозитория:

```powershell
$scopeSdk = 'C:\Program Files\Microsoft Visual Studio\2022\Community\SDK\ScopeCppSDK\vc15'
$mapiHeaders = (Resolve-Path '..\third_party\MAPIStubLibrary\include').Path
$env:INCLUDE = "$mapiHeaders;$scopeSdk\VC\include;$scopeSdk\SDK\include\ucrt;$scopeSdk\SDK\include\um;$scopeSdk\SDK\include\shared"
$env:LIB = "$scopeSdk\VC\lib;$scopeSdk\SDK\lib"
$env:PATH = "$scopeSdk\VC\bin;$scopeSdk\SDK\bin;$env:PATH"

& "$scopeSdk\VC\bin\cl.exe" /nologo /std:c++17 /EHsc /W4 `
  /permissive- /utf-8 /O2 /DNDEBUG /DUNICODE /D_UNICODE `
  /DWIN32_LEAN_AND_MEAN /DNOMINMAX `
  /Fo:build-direct\main.obj /Fe:build-direct\NativeMapiProbe.exe `
  main.cpp /link MAPI32.lib Ole32.lib
```

## Запуск

Первый короткий positive-control:

```powershell
.\build-direct\NativeMapiProbe.exe `
  --max-stores 0 `
  --max-folders 25 `
  --max-messages 5 `
  --body-preview-chars 200
```

Для stores/folders/messages `0` означает unlimited. Полный обход потенциально очень
дорогой, поэтому его стоит делать только осознанно:

```powershell
.\build-direct\NativeMapiProbe.exe `
  --max-folders 0 `
  --max-messages 0 `
  --body-preview-chars 500
```

Можно ограничить proof одним store без изменения профиля, например PST с display
name `Archives`:

```powershell
.\build-direct\NativeMapiProbe.exe `
  --store-contains Archives `
  --max-folders 25 `
  --max-messages 5
```

Probe логинится без `MAPI_LOGON_UI`; если default profile отсутствует или требует
интерактивного выбора, он завершится с HRESULT вместо показа profile picker.

### JSONL для Python pipeline

`--jsonl` оставляет на `stdout` только UTF-8 JSON Lines — один объект на сообщение.
Progress, recoverable errors и итоговый summary направляются в `stderr`, поэтому
stdout можно безопасно читать построчно из Python без concat всего mailbox:

```powershell
.\build-direct\NativeMapiProbe.exe `
  --jsonl `
  --max-folders 100 `
  --max-messages 1000 `
  --body-preview-chars 200000 `
  --spool-dir C:\Users\User\AppData\Local\RAGSearch\spool `
  --max-attachment-bytes 67108864 `
  --max-total-attachment-bytes 0 `
  2> probe.log |
  python .\consume_jsonl.py
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

`--body-preview-chars` в этом prototype одновременно является пределом поля `body`;
hard maximum — `4000000` символов (примерно 8 MiB UTF-16 read buffer).
Для production ingestion лучше оставить bounded producer/consumer с backpressure,
а не накапливать весь stdout в памяти вызывающего процесса.

Даты всегда имеют ISO-8601 UTC вид с суффиксом `Z` либо равны JSON `null`.
`conversation_id` выбирается из `PR_CONVERSATION_ID`, затем
`PR_CONVERSATION_INDEX`, затем `PR_SEARCH_KEY`; binary value кодируется hex.
`to`/`cc` — provider display strings (`PR_DISPLAY_TO`/`PR_DISPLAY_CC`), не
нормализованный список SMTP-адресов.

Message-class filter выполняется по `PR_MESSAGE_CLASS_W/A` ещё в contents table,
до дорогого `OpenEntry`. Точное `IPM.Note` и prefix `IPM.Note.` сохраняют обычные,
custom-form и S/MIME письма. `REPORT.*` и `IPM.Schedule.Meeting.*` намеренно
исключены: Outlook Object Model материализует их как `ReportItem`/`MeetingItem`, а
предыдущий индексатор принимал только успешный cast `item as Outlook.MailItem`.
Пропуски не расходуют `--max-messages` и отражаются в summary как
`skipped_non_mail`.

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
3. Outlook продолжил нормально работать параллельно с probe.

Extended MAPI — отдельный unmanaged API, а не обход/отключение настроенной политики
Programmatic Access. Итоговое поведение всё равно зависит от доступности MAPI,
настроек профиля, прав пользователя и корпоративных ограничений. Обзор выбора API:
<https://learn.microsoft.com/en-us/office/client-developer/outlook/selecting-an-api-or-technology-for-developing-solutions-for-outlook>.

## Результат проверки на этой машине (2026-08-11)

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

Полный штатный build через установленный Visual Studio сейчас недоступен: на машине
не установлен workload Desktop development with C++ (`Microsoft.Cpp.Default.props`
отсутствует), а `cmake` не находится в `PATH`. Для proof использован имеющийся x64
compiler из bundled ScopeCppSDK и зафиксированные в репозитории headers из
официального Microsoft MAPIStubLibrary. Готовый проверенный бинарник лежит в
`build-direct\NativeMapiProbe.exe`.

## Ограничения prototype

- Используется default profile, а не UI-выбор профиля. Если Outlook запущен с
  не-default profile, это может быть другой mailbox.
- Не перечисляются associated/hidden contents.
- Body читается как `PR_BODY_W` stream с ANSI/property fallback. HTML/RTF отдельно не
  извлекаются.
- SMTP sender берётся из `PR_SENDER_SMTP_ADDRESS`/sent-representing с fallback на
  provider email property. Для старых Exchange items fallback может быть legacy EX
  address, а не SMTP.
- Вложения embedded-message/OLE/by-reference не разворачиваются; для них остаются
  metadata и пустой `temp_path`.
- `EntryID` локален для конкретного MAPI profile/store и не является переносимым
  глобальным идентификатором.
- Ошибки отдельных inaccessible folders/messages считаются recoverable: probe идёт
  дальше и возвращает exit code `1`; fatal init/logon errors имеют другие exit codes.
