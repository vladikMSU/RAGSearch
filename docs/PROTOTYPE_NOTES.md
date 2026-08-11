# RAGSearch: журнал прототипа

Дата: 2026-08-11  
Workspace: `<repo-root>`

## Короткий итог

Первоначальная гипотеза «VSTO trusted, поэтому prompt нет» была недостаточной: на машине не было требуемой strict policy. `Application.IsTrusted=True` — классификация Outlook, а не найденный bypass.

Сейчас подготовлены две честно разделённые дорожки:

1. Точный strict interactive Object Model Guard с `AddinTrust=2` и всеми `PromptOOM*=1` был воспроизведён вручную и оставлен внешней политикой; RAGSearch его не переключает.
2. Read-only Extended MAPI reader. Он читает PST и текущий Exchange/OST store через MAPI provider, не использует Outlook Object Model и поэтому не попадает под Object Model Guard. Это поддерживаемая API-граница, а не отключение политики.

Панель перенесена наверх. Семантические результаты теперь отображаются в штатной центральной таблице Outlook через native search, а не в отдельном окне. Реальная UI-проверка дала 16 видимых Outlook items для запроса `куку`.

Strict policy вручную применена и подтверждена реальным Outlook Guard warning. Clean A/B-прогон OOM против Extended MAPI должен учитывать временный grant после `Allow`.

## Машина и данные

- Windows 11 Pro, classic Outlook x64 `16.0.20228.20158`.
- Outlook profile: `user@example.com`.
- Cached Exchange OST: `%LOCALAPPDATA%\Microsoft\Outlook\user@example.com.ost`.
- PST: `%USERPROFILE%\Documents\Outlook Files\archive.pst`, store `Archives`.
- Python 3.12, .NET Framework 4.8, Visual Studio 2022.
- Начальный policy state: явные Object Model Guard policy values отсутствовали.
- На момент эксперимента workspace ещё не был Git-репозиторием.

## Что означает `IsTrusted`

- Внешний PowerShell/Python COM-клиент получает `Outlook.Application.IsTrusted=False`.
- VSTO, загруженный самим Outlook и использующий `ThisAddIn.Application`, обычно получает trusted object chain.
- Это штатное решение Outlook, а не изменение security settings.
- Явная политика `AddinTrust=2` должна заставить Guard проверять и trusted add-ins.
- Если прежний эксперимент в корпоративной среде prompt-ил VSTO, наиболее вероятны `AddinTrust=2` либо создание нового `Outlook.Application`/потеря trusted object chain.

Microsoft описывает это разделение в [Outlook Object Model security behavior](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/security/security-behavior-of-the-outlook-object-model) и [VSTO security considerations](https://learn.microsoft.com/en-us/visualstudio/vsto/specific-security-considerations-for-office-solutions?view=vs-2022).

## Воспроизведение strict interactive Guard

В репозитории оставлен только `scripts\guard_probe.ps1` — внешний untrusted OOM negative control. Код изменения policy и UI-переключатель удалены после завершения эксперимента.

### Точные значения

```text
HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security
  ObjectModelGuard = 1

HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security
  AdminSecurityMode = 3
  AddinTrust = 2
  PromptOOMCustomAction = 1
  PromptOOMSend = 1
  PromptOOMAddressBookAccess = 1
  PromptOOMAddressInformationAccess = 1
  PromptOOMMeetingTaskRequestResponse = 1
  PromptOOMSaveAs = 1
  PromptOOMFormulaAccess = 1
  PromptOOMItemPropertyAccess = 1
  PromptOOMAddressUserPropertyFind = 1
```

Это интерактивный строгий режим. `PromptOOM*=0` означает запрет без возможности Allow и не подходит для воспроизведения требуемого окна. `10 minutes` — выбор в самом Outlook warning dialog; отдельного registry value для него нет. Временное разрешение распространяется на OOM callers в этой Outlook session, поэтому его нельзя считать изоляцией только нашей надстройки. См. [Object Model security warnings](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/security/outlook-object-model-security-warnings).

Эффективные ручные values соответствуют точному списку выше, и Outlook показал реальный warning. Текущее состояние управляется вне RAGSearch; надстройка не читает, не записывает и не восстанавливает эти registry values.

## Negative и positive controls

### Внешний OOM negative control

`scripts\guard_probe.ps1` запускается отдельным PowerShell COM-процессом:

- `Application.IsTrusted=False`;
- на protected access Outlook уже показывал реальный warning об автоматическом доступе к адресной информации;
- после Deny protected body не был прочитан (`body_length=0`).

Это доказывает, что Guard на машине существует, но не заменяет тест exact strict policy с `AddinTrust=2`.

### VSTO baseline

До strict policy VSTO прошёл по двум stores без prompt:

- 42 Outlook items;
- 49 attachment records;
- 1 877 chunks;
- ingestion errors: 0.

Этот результат показывает trusted baseline. Он не доказывает работу при администраторской политике и не является обходом.

### Выполненный A/B live control

1. Все strict values проверены в registry: `ObjectModelGuard=1`, `AdminSecurityMode=3`, `AddinTrust=2`, девять `PromptOOM*=1`.
2. Внешний `guard_probe.ps1` без elevation показал warning; пользователь нажал `Deny`.
3. Сразу после этого x64 Extended MAPI probe прочитал Outlook profile/stores без warning.
4. Полный native scan также завершился без выдачи OOM grant.

Это отделяет реальный Guard negative control от production-маршрута. Ни UAC, ни `Allow`/`Deny` автоматикой не нажимаются: это осознанные security decisions пользователя. Временный Address Book grant относится ко всем OOM callers данной Outlook session, а не только к исходной надстройке, поэтому тест после `Allow` не считается чистым.

### Attribution startup warning

После отдельного clean restart настоящий address-information Guard остался открыт на splash screen. Новая RAGSearch DLL была однозначно загружена из VSTO cache: MVID `3ca8c99d-25dc-44d1-8147-4ddeada45a86`. Ранний `%LOCALAPPDATA%\RAGSearch\startup-trace.log` показал пары `BEGIN/END` для всех этапов с `18:38:26.754` по `18:38:27.145`: filesystem spool, loopback client, WinForms pane, `/health` и безопасный `Application.ActiveExplorer`. В production-startup нет `Session`, `Stores`, `Items`, `MailItem`, `Sender`, `Recipients`, `AddressEntry`, `PropertyAccessor` или `NewMailEx`.

Guard продолжал блокировать main Outlook UI после полного `END ThisAddIn_Startup`; внешние probe/COM scripts в этот момент не работали. В процессе были загружены `SOCIALCONNECTOR.DLL`, OneNote `ONBttnOL.dll`, Skype `UCAddin.dll` и Exchange `UmOutlookAddin.dll`. Точный caller Guard не публикует. Наиболее сильный кандидат — Outlook Social Connector: его штатная функция включает работу с people/GAL, и Microsoft отдельно документирует GAL synchronization и её policy controls: [Manage the Outlook Social Connector](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/configuration/manage-outlook-social-connector). Для точной локализации нужен последовательный A/B с `Deny` и временным отключением по одной чужой add-in; RAGSearch для production-проверки отключать не требуется.

Дополнительный positive control оказался сильнее обычного: пока этот чужой Guard оставался открытым, `test_native_mapi_adapter_e2e.ps1` успешно импортировал 3 PST messages/39 attachments/6 chunks и завершился `PASS`. Значит Extended MAPI worker не ждёт OOM-диалог и не пользуется его grant.

## PST и OST без путаницы

### Debug VSTO/OOM путь

`OutlookIndexer.cs` делает следующее:

1. Ищет только верхнеуровневые `*.pst` в `%USERPROFILE%\Documents\Outlook Files`.
2. Отсутствующие PST подключает через `NameSpace.AddStore`.
3. Перебирает `Application.Session.Stores`.
4. Одинаково обходит folders/items каждого store.
5. Отправляет структурированные DTO Python-сервису.

PST читается не raw-парсером, а подключённым PST store provider. OST также не открывается как файл: Outlook отдаёт текущий Exchange cached store через тот же `Session.Stores`. Для OST нет другого поля/объекта; различается provider за store.

Закрывать Outlook и копировать OST для этого пути не требуется. Но OOM protected properties под strict policy будут под Guard.

Полный legacy scanner больше не используется кнопкой negative control. Предыдущая реализация сначала делала `await GetHealthAsync()`: startup/pane работали на `tid=1`, а continuation и `OutlookIndexer.Start()` реально ушли на thread-pool `tid=11`. Запущенный там `WinForms.Timer` не имел UI message pump, поэтому `Tick` не наступил ни разу. Статус успел показать подготовленные `0/42`, но protected getter ещё не был достигнут — отсутствие Guard ничего не доказывало.

Исправленный `OomGuardProbe.cs` работает синхронно в click-handler без `await` и не обращается к Python-сервису или Timer. Он требует один явно выбранный `MailItem`, не обходит folders/stores и сразу читает документированное protected-свойство `MailItem.SenderEmailAddress`. Значение немедленно отбрасывается и никогда не попадает в UI/log. Фазы и HRESULT записываются в `%LOCALAPPDATA%\RAGSearch\oom-guard-probe.log`; результат различает normal return, документированный для Deny `MAPI_E_NOT_SUPPORTED`, наблюдавшийся на этой сборке Outlook `E_FAIL`, известные `E_ABORT`/`E_ACCESSDENIED` и прочие COM exceptions. На время modal Guard блокируются поиск и остальные действия панели. Нажатия `Allow`/`Deny` не автоматизируются.

Live proof исправления: заранее запущенный window watcher зафиксировал настоящий `#32770` Guard с текстом address-information warning непосредственно после `PROTECTED_GETTER_BEGIN`; после ручного `Deny` Outlook дважды вернул `HRESULT=0x80004005 (E_FAIL)`. Оба прогона завершились, адрес не был прочитан/залогирован, зависания `0/42` больше нет.

### Extended MAPI путь

`native-mapi-probe`:

- x64, чтобы совпасть с Outlook x64;
- `MAPIInitialize` и `MAPILogonEx` без UI;
- stores открываются read-only, без `MDB_WRITE`/`MAPI_MODIFY`;
- обход начинается с `PR_IPM_SUBTREE_ENTRYID`;
- читаются folder hierarchy, реальные store/folder/message IDs, subject/body, sender/recipients, даты, Internet Message ID, Conversation ID и attachment metadata/content;
- write/send APIs отсутствуют;
- `--jsonl` пишет один UTF-8 JSON object на сообщение, diagnostics идут в stderr.

Финальный полный прогон при запущенном Outlook:

- 2 stores, 41 folders, 42 валидных mail rows: OST — 16, PST `Archives` — 26;
- 49 attachment records, 31 non-mail MAPI object корректно пропущен, recoverable errors — 0;
- фильтр message class принимает case-insensitive `IPM.Note` и `IPM.Note.*`, но исключает REPORT/Meeting/Calendar/Contacts, как прежний `item as Outlook.MailItem`;
- blocking OOM popup не появился;
- SHA-256 проверочного EXE: `BDFFD3C5D5718D2771D17FB203A6B89DC6ADE98DB6AC259AA314302E7CC1A00C`.

Почему OST lock не мешает: file handle принадлежит Outlook, но MAPI client разговаривает с profile/store provider. Это не raw-file access. Microsoft рекомендует выбирать API по сценарию и описывает Extended MAPI как native C/C++ API для Outlook/Exchange/PST: [Selecting an API](https://learn.microsoft.com/en-us/office/client-developer/outlook/selecting-an-api-or-technology-for-developing-solutions-for-outlook). `MAPILogonEx` содержит флаг `MAPI_BG_SESSION` для background/indexing scenarios: [MAPILogonEx](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/mapilogonex).

## Python service и native adapter

Service contract:

- `GET /health`;
- `GET /v1/stats`;
- `POST /v1/messages`;
- `POST /v1/search`;
- `DELETE /v1/index` — транзакционная очистка локального индекса без удаления SQLite-файла, token, model metadata или spool.

Все `/v1/*` требуют `X-RAGSearch-Token` из `%LOCALAPPDATA%\RAGSearch\service-token`. Attachment `temp_path` принимается только из spool. Хранилище нормализовано на message/attachment/chunk, FTS5 и embeddings; весь mailbox никогда не конкатенируется в один string.

`service\import_native_mapi.py`:

- запускает probe аргументами с `shell=False`;
- разрешает только loopback HTTP URL и существующий token;
- потоково читает и строго валидирует полный JSONL contract;
- отправляет по одному authenticated `POST /v1/messages` с backpressure;
- создаёт уникальный marker-owned spool run, проверяет canonical attachment paths/размеры и удаляет только собственные временные файлы;
- поддерживает `--full-scan`, cancellation sentinel и progress protocol;
- отклоняет malformed/oversized records и ingestion failures;
- Outlook COM не использует.

E2E positive control: отдельный сервис на `127.0.0.1:8877`, 3 PST messages, 39 attachment records, stats `messages=3`, `chunks=6`, owned spool пуст после импорта, exit PASS. Основной full scan дал `messages=42`, `attachments=49`, `chunks=1877`. Флаг `body_truncated` пока не хранится в БД, но adapter считает и показывает такие записи в progress/final summary.

## UI и нативная фильтрация

Реализовано:

- top-docked task pane, высота 126 px;
- query, `Семантический поиск`, `Сбросить фильтр`;
- production-кнопка `Индексировать PST + OST (MAPI)` запускает full-scan adapter асинхронно без shell/UAC и показывает progress;
- `Очистить базу` показывает warning с `No` по умолчанию, при необходимости запускает локальный service и через authenticated `DELETE /v1/index` удаляет только сообщения/вложения/чанки/FTS из локального индекса;
- `Стоп` сначала создаёт уникальный cancellation sentinel, затем ограниченно ждёт graceful shutdown и только после таймаута завершает принадлежащее этому запуску Windows Job tree;
- отдельная кнопка `Debug OOM: 1 письмо → Guard` немедленно вызывает один protected OOM getter и служит детерминированным negative control; полный legacy scanner к кнопке больше не подключён;
- production-подписка `NewMailEx -> OOM extractor` удалена: при старте RAGSearch protected getters больше не вызываются;
- при недоступном `/health` надстройка может запустить точный workspace `service\.venv\Scripts\python.exe service\run.py`; уже работающие service processes она не завершает;
- отдельный result window/DataGrid удалён;
- `NativeSearchPresenter` вызывает `Explorer.Search(..., olSearchScopeAllFolders)` в привязанном task-pane Explorer и показывает один All Mailboxes rowset из выбранных Outlook stores;
- сервис получает `filters: {}` и ищет по всему локальному индексу, а не только в текущей папке;
- `Сбросить фильтр` вызывает `Explorer.ClearSearch()` и возвращает исходную папку.

Предыдущая текущая-folder реализация через `View.Filter` действительно оставляла Search bar пустым, но архитектурно не могла собрать PST и OST в одном списке. Live probe доказал, что широкий `Explorer.Search` даёт Inbox + Sent + PST `Archives`, а попытка наложить на этот aggregate rowset сохранённый `View.Filter` ничего не меняет: таблица осталась 14/14. Поэтому production-прототип использует финальный AQS как единственный поддержанный cross-store restriction.

Property AQS оказался locale/provider-dependent: canonical `System.Subject:=...` и составные `тема:/откого:/получено:` на целевой ru-RU Windows + en-US Office либо дали 0, либо не распарсились как ожидается. Финальная версия использует короткий OR из обычных кавыченных фраз, полученных из тем результатов. Он приблизителен, но реально работает и заметно понятнее старой технической конструкции.

Финальный реальный прогон:

```text
semantic query: дарова
scope: AllFolders across Outlook stores selected for search
semantic results: structured result list
ranking: literal gate, then vector_distance ascending; adaptive cutoff; UI top-12
native presentation: Explorer.Search(final subject-derived quoted-phrase AQS, olSearchScopeAllFolders)
projection identity: approximate quoted phrases derived from result subjects
Search bar: generated query is visible by Outlook contract
```

Финальный clean UI control выполнен после `Deny` в отдельном startup Guard: production-кнопка панели породила adapter/native process tree, закончила статусом `Native-индексация завершена: 42 писем` и не вызвала нового warning. После upsert основная БД сохранила `messages=42`, `attachments=49`, `chunks=1877`, `PRAGMA integrity_check=ok`; marker-owned native spool directories — 0.

Ограничение: `Explorer.Search` неизбежно использует видимый Instant Search UI, кавыченная subject-фраза не является exact `(StoreID, EntryID)`, а native view не сортирует по внешнему SQLite vector distance. Extended MAPI Search Folder может дать exact `PR_RECORD_KEY`, но только отдельно для каждого OST/PST store; для единого cross-store списка в точном vector-порядке нужен собственный result list. Подробности: [OUTLOOK_SEARCH_PROJECTION.md](OUTLOOK_SEARCH_PROJECTION.md).

## Проверки кода

Последние результаты:

- `python -m unittest discover -s service\tests -t service -v`: 33/33 PASS;
- `scripts\test_native_mapi_adapter_e2e.ps1`: PASS;
- `RAGSearch.sln`, Debug build: 0 warnings, 0 errors;
- native JSONL parsed by Python `json.loads`;
- real PST/OST read-only probes: exit 0.

Standard MSBuild native-проекта сейчас блокируется отсутствующим Visual Studio workload `Desktop development with C++`/`Microsoft.Cpp.Default.props`. Проверочный binary собран доступным x64 compiler toolchain с официальными MAPIStubLibrary headers. Для воспроизводимого production build workload надо установить штатно.

## Целевая production-архитектура

1. Native x64 MAPI worker выполняет initial PST + cached Exchange scan и incremental reconciliation/notifications.
2. Python service извлекает attachments, чанкует внутри сущностей, считает FTS/embeddings и хранит identity/metadata.
3. VSTO отвечает только за верхнюю UI-панель, native Outlook filtering/navigation и жизненный цикл локального service.
4. Graph может дополнять Exchange Online history за пределами OST cache; PST через Graph недоступны.
5. Object Model Guard и корпоративные политики остаются включены.

Такой расклад убирает зависимость массового ingestion от 10-минутного OOM grant. Для новых писем можно использовать Extended MAPI notifications или короткий periodic delta, а VSTO `NewMailEx` оставить лишь необязательным ускорителем.

## До production

- incremental MAPI notifications/checkpoints вместо полного повторного scan;
- durable checkpoints, tombstones, move/delete reconciliation и stable identity (`PR_SEARCH_KEY`/InternetMessageID/content hash);
- exact MAPI Search Folder projection либо явно принять subject-phrase approximation;
- ANN index для больших архивов;
- sandbox/timeouts/AV для PDF, Office, OCR и архивов;
- шифрование БД/секретов, ACL, DLP, retention и audit;
- code signing/installer для VSTO, native worker и service;
- проверка под реальной доменной/cloud policy после локального exact-value experiment;
- отдельный UI/Graph design для new Outlook, где VSTO/COM не поддерживаются.
