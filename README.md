# RAGSearch for classic Outlook

Локальный прототип гибридного поиска по classic Outlook: VSTO-панель сверху, Python-сервис с SQLite/FTS5 и локальными embeddings, а также read-only Extended MAPI reader для чтения PST и текущего Exchange/OST store без Outlook Object Model Guard.

Главный вывод эксперимента: `Application.IsTrusted` не является обходом защиты. Обычная VSTO-надстройка действительно считается Outlook доверенной, но политика `AddinTrust=2` должна принудительно вернуть её под Object Model Guard. Поддерживаемый путь без prompt при неизменной политике — читать store provider через Extended MAPI, а не отключать Guard и не копировать заблокированный `.ost`.

Подробный журнал с точными policy-значениями, доказательствами и ограничениями: [docs/PROTOTYPE_NOTES.md](docs/PROTOTYPE_NOTES.md).

## Что работает

- Верхняя панель `RAG Search` внутри Outlook.
- Гибридный FTS5 + vector поиск по сообщениям и вложениям.
- Результаты проецируются в штатную центральную таблицу Outlook через `Explorer.Search`; отдельного окна/грида нет.
- Кнопка `Сбросить фильтр` возвращает обычный Outlook view.
- Русский AQS определяется по Windows user locale, даже если UI Outlook английский.
- Основная кнопка индексации запускает отдельный read-only Extended MAPI worker; отдельная debug-кнопка синхронно читает один protected OOM getter на одном письме и служит negative control.
- Кнопка `Очистить базу` после подтверждения удаляет только локальный поисковый индекс и позволяет повторить векторизацию с нуля; Outlook, PST и OST не изменяются.
- Надстройка не подписывает OOM extractor на `NewMailEx`, поэтому сама RAGSearch больше не должна вызывать Guard во время старта Outlook.
- Если loopback Python-сервис не отвечает, панель запускает workspace `service\.venv` без shell/UAC, ограниченно ждёт `/health` и затем начинает native import.
- Read-only x64 Extended MAPI probe читает и текущий Exchange/OST store, и `archive.pst` через Outlook profile/store provider; `.ost` как файл не открывается и не копируется.
- Native JSONL adapter отправляет прочитанные сообщения в локальный Python API без Outlook COM.
- Loopback-only API защищён локальным токеном; attachment spool ограничен `%LOCALAPPDATA%\RAGSearch\spool`.

Финальный clean UI-контроль под strict policy: пользователь нажал `Deny` в чужом startup Guard, затем кнопка `Индексировать PST + OST (MAPI)` из панели завершилась статусом `Native-индексация завершена: 42 писем`; нового Guard не появилось.

Проекция в native view строит до 8 компактных AQS-критериев из темы, отправителя и локальной даты получения; если Outlook отвергает уточнённый запрос, надстройка повторяет его только по теме. Семантический результат при этом остаётся структурированным списком сообщений, а не одним склеенным текстом.

Публичный `Explorer.Search` не принимает произвольный набор `(StoreID, EntryID)`, поэтому это всё ещё приближённая проекция: одинаковые тема/отправитель/дата могут дать дополнительные строки, Outlook сам сортирует результаты, а scope зависит от его Search Options. Для математически точного production-набора нужен MAPI Search Folder/служебное индексируемое свойство либо собственный grid.

## PST и OST: что именно читается

Legacy-класс полного VSTO/OOM-индексатора не парсит файлы:

1. `NameSpace.AddStore` подключает отсутствующий PST к Outlook profile.
2. `Session.Stores` перечисляет PST и Exchange cached store.
3. Оба читаются через Outlook Object Model и отправляются DTO в Python.

Поэтому PST читается через подключённый PST provider, а OST — через активный Exchange store provider. Это один и тот же объектный путь; никакого «другого поля OST» нет.

Кнопка `Debug OOM: 1 письмо → Guard` больше не запускает этот полный индексатор. Она требует явно выбранный `MailItem` и сразу читает документированное protected-свойство `MailItem.SenderEmailAddress`. Адрес отбрасывается и не логируется. До исправления кнопка сначала ожидала `/health`, continuation ушёл с Outlook `tid=1` на thread-pool `tid=11`, а запущенный там `WinForms.Timer` не имел message pump; поэтому UI оставался на `0/42` и protected getter вообще не выполнялся.

Native-путь также не трогает raw-файлы. `NativeMapiProbe.exe` выполняет `MAPIInitialize`/`MAPILogonEx`, открывает stores read-only и получает тело/идентификаторы через Extended MAPI. Он уже успешно прочитал оба источника при запущенном Outlook, без Guard popup.

## Outlook Guard — только диагностический контроль

Strict policy установлена вне RAGSearch и подтверждена реальным Outlook Guard warning. Надстройка больше не читает и не изменяет policy/registry; точные воспроизведённые значения сохранены только в журнале прототипа. Для чистого сравнения `OOM prompt` против `Extended MAPI без prompt` нельзя считать отсутствие окна доказательством, пока действует временный grant после `Allow`.

Детерминированный negative control после полного перезапуска Outlook:

1. Выбрать одно обычное письмо в центральном списке и нажать `Debug OOM: 1 письмо → Guard`.
2. Статус дойдёт до `[2/3]` и синхронно вызовет `MailItem.SenderEmailAddress`; решение `Allow`/`Deny` автоматикой не нажимается.
3. После решения статус покажет `GETTER_RETURNED`, `DENY_OR_BLOCKED` либо `COM_EXCEPTION` и точный HRESULT. Значение адреса не выводится; технические фазы пишутся в `%LOCALAPPDATA%\RAGSearch\oom-guard-probe.log`.

`GETTER_RETURNED` сам по себе не различает новый `Allow`, уже действующий временный grant и отсутствие enforcement — OOM не сообщает источник разрешения. Для чистого теста нельзя предварительно выдавать временный grant.

При `AddinTrust=2` warning на старте может вызвать любая загруженная COM-надстройка Outlook. В диагностической сборке RAGSearch ранний `%LOCALAPPDATA%\RAGSearch\startup-trace.log` подтвердил завершение всех startup-этапов без protected mail/address getters; настоящий Guard при этом остался открыт от другого компонента. Загружены, в частности, Outlook Social Connector, OneNote, Skype Meeting и Exchange add-ins. Production-кнопка RAGSearch не зависит от их OOM-решения: native adapter E2E прошёл даже при открытом модальном Guard.

## Запуск Python-сервиса

```powershell
cd "$env:USERPROFILE\Desktop\projects\RAGSearch"
.\service\.venv\Scripts\python.exe .\service\run.py `
  --embedding sentence-transformers `
  --model .\service\models\paraphrase-multilingual-MiniLM-L12-v2 `
  --delete-spool-after-ingest
```

Сервис слушает только `127.0.0.1:8765`. База, token и spool находятся в `%LOCALAPPDATA%\RAGSearch`.

Dependency-free fallback без нейронной модели:

```powershell
python .\service\run.py --delete-spool-after-ingest
```

## Native MAPI ingestion

Локально собранный проверочный binary (build output не хранится в Git):

```powershell
.\native-mapi-probe\build-direct\NativeMapiProbe.exe --max-messages 5
.\native-mapi-probe\build-direct\NativeMapiProbe.exe --store-contains Archives --jsonl
```

Потоковый импорт в основной сервис:

```powershell
.\service\.venv\Scripts\python.exe .\service\import_native_mapi.py `
  --probe .\native-mapi-probe\build-direct\NativeMapiProbe.exe `
  --max-messages 1000
```

JSONL-контракт содержит реальные `store_id`, `entry_id`, `folder_entry_id`, путь/имя store и folder, тему/тело, sender/recipients, даты, Internet Message ID, Conversation ID и attachment metadata. Вложения извлекаются только по явному `--spool-dir`, с canonical path containment и лимитом 64 MiB на файл. Worker выдаёт только `IPM.Note`/`IPM.Note.*`, поэтому Calendar/Contacts/REPORT не загрязняют индекс.

Обычная Visual Studio сборка native-проекта требует workload `Desktop development with C++` и Microsoft MAPIStubLibrary headers. Проверочный EXE собран x64 существующим compiler toolchain; битность совпадает с Outlook x64.

## Сборка VSTO

```powershell
$msbuild = 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe'
& $msbuild .\RAGSearch.sln /t:PrepareForRun /p:Configuration=Debug /m
```

Затем перезапустить classic Outlook. Debug manifest подписан локальным dev-сертификатом и регистрируется в `HKCU\Software\Microsoft\Office\Outlook\Addins\RAGSearch`. Для распространения нужен корпоративно одобренный сертификат/installer.

## Проверки

```powershell
python -m unittest discover -s service\tests -t service -v
powershell.exe -NoProfile -File .\scripts\test_native_mapi_adapter_e2e.ps1
& $msbuild .\RAGSearch.sln /t:Build /p:Configuration=Debug /m
```

Последний прогон: 25/25 Python tests, native-MAPI-to-service E2E PASS даже при открытом Guard, VSTO build — 0 warnings/0 errors. Полный read-only MAPI scan дал ровно 42 письма (16 OST + 26 PST), 49 вложений, 31 корректно пропущенный non-mail объект и 0 ошибок. Основная neural DB содержит 42 сообщения, 49 вложений и 1 877 chunks.

## До production

- добавить инкрементальные MAPI notifications/checkpoints вместо полного повторного scan;
- сделать точную native Search Folder projection, если subject phrase projection недостаточна;
- заменить brute-force cosine на ANN для больших архивов;
- добавить reconciliation/tombstones для перемещений и удалений;
- изолировать тяжёлый extraction, добавить AV/resource quotas;
- согласовать шифрование, DLP, retention и аудит;
- для полной Exchange Online истории рассмотреть Graph; OST содержит только доступное cached window;
- VSTO работает только в classic Outlook.
