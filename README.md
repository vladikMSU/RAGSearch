# RAGSearch

Экспериментальный локальный поиск по почте, совместимый с classic Microsoft Outlook
для Windows x64.
VSTO-надстройка показывает собственную нижнюю WinForms-панель поиска и результатов,
отдельный C++ worker читает настроенный Outlook profile через read-only Extended
MAPI, а Python-сервис индексирует письма и вложения в SQLite/FTS5. Штатные строка
поиска и список писем Outlook при этом не изменяются.

```text
classic Outlook + нижняя VSTO-панель ◄──HTTP──► Python search service ◄──HTTP── Python adapter ◄──JSONL── C++ Extended MAPI reader
          │                                      │                                                       ▲
          │ двойной щелчок: StoreID + EntryID    ▼                                                       │
          ▼                                SQLite + FTS5 + embeddings                         Outlook profile/providers
   исходное письмо
```

Extended MAPI здесь не является скачанным Python-пакетом. Во время сборки
используются закреплённые в репозитории официальные Microsoft headers; линковочные
библиотеки даёт Windows SDK, а MAPI provider во время работы — classic Outlook.
Подробнее: [архитектура и границы компонентов](docs/ARCHITECTURE_AND_MAPI.md).

## Что уже находится в репозитории

- исходники VSTO-надстройки, native x64 worker и Python-сервиса;
- шесть необходимых headers Microsoft MAPIStubLibrary, их MIT-лицензия и точный
  upstream commit в [third_party/MAPIStubLibrary](third_party/MAPIStubLibrary);
- dependency-free Python provider и тесты — для базового режима сторонние
  Python-пакеты и скачивание модели не нужны;
- конфигурации MSBuild и CMake; native EXE собирается в
  `native-mapi-probe\build-direct\NativeMapiProbe.exe`.

В Git намеренно не хранятся `.venv`, build artifacts, база, письма, вложения,
токены, private signing keys и необязательная neural model.

## Требования

Для компиляции из чистого clone:

- Windows x64;
- Visual Studio 2022 с workload `Office/SharePoint development`;
- `.NET Framework 4.8 Targeting Pack`;
- workload `Desktop development with C++`, MSVC v143 x64/x86 и Windows 10/11 SDK.

Для запуска дополнительно нужны Python 3.11+, VSTO Runtime, classic Outlook x64 и
настроенный default Outlook profile. Разрядность native worker и Outlook обязана
совпадать; текущая конфигурация проекта — только x64. CMake 3.20+ необязателен и
нужен лишь при сборке native worker через CMake вместо MSBuild.

## Сборка из чистого clone

Команды ниже выполняются из корня репозитория. Чтобы собранная надстройка затем
могла запустить сервис, создайте ожидаемое ею Python-окружение; базовый режим
использует только стандартную библиотеку Python:

```powershell
py -3 --version # должно быть 3.11+
py -3 -m venv .\service\.venv
```

Соберите оба проекта одной командой. Скрипт находит Visual Studio через `vswhere`,
собирает native worker и создаёт либо повторно использует локальный non-exportable
development certificate в `Cert:\CurrentUser\My`, потому что полный VSTO `Build`
технически требует подписанный manifest:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

Дополнительные MAPI-файлы скачивать или указывать через include path не нужно.
`RAGSearch.sln` пока содержит только VSTO-проект, поэтому native worker собирается
отдельной первой командой.

Сертификат автора и private key в Git не входят. Для повторной сборки можно оставить
созданный скриптом локальный dev certificate. Чтобы использовать свой уже
установленный code-signing certificate с private key:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 `
  -CertificateThumbprint YOUR_CERTIFICATE_THUMBPRINT
```

Автоматический self-signed certificate разрешён скриптом только для `Debug`.
`-VstoConfiguration Release` требует явно переданный production certificate.

Скрипт решает воспроизводимость development-сборки, но не доверие к издателю.
Не коммитьте `.pfx`, private key или локальный thumbprint. Для распространения нужен
нормальный installer и сертификат, которому доверяет целевая организация;
репозиторий их не содержит. Автоматически созданный сертификат также может
потребовать явного доверия в соответствии с локальной ClickOnce/Office policy.
Ручные MSBuild/CMake команды и overrides описаны в
[native-mapi-probe/README.md](native-mapi-probe/README.md).

## Запуск

Python-сервис можно запустить вручную:

```powershell
.\service\.venv\Scripts\python.exe .\service\run.py `
  --delete-spool-after-ingest
```

После сборки и регистрации надстройки перезапустите classic Outlook. Если сервис
не запущен, панель сама использует точный workspace-путь
`service\.venv\Scripts\python.exe`; при отсутствии локальной neural model включается
воспроизводимый hashing provider без дополнительных зависимостей.

Результаты появляются в закреплённой снизу таблице RAG Search в том же порядке,
в котором их вернул backend; максимум — 25 писем. Двойной щелчок по строке открывает
исходный Outlook `MailItem` по точной паре `StoreID + EntryID`. Панель подбирает
светлую или тёмную WinForms-палитру по системной теме, умеет сворачиваться и
отделяться в плавающее окно. Обычные Search bar, текущая папка и список Outlook
остаются без изменений.

Ручной потоковый импорт:

```powershell
.\service\.venv\Scripts\python.exe .\service\import_native_mapi.py `
  --executable .\native-mapi-probe\build-direct\NativeMapiProbe.exe `
  --full-scan
```

Необязательные `sentence-transformers` и multilingual model не включены в Git и
никогда не скачиваются сервисом автоматически. Их локальная настройка описана в
[service/README.md](service/README.md).

## Проверки

```powershell
.\service\.venv\Scripts\python.exe -B -m unittest `
  discover -s service\tests -t service -v

.\native-mapi-probe\build-direct\NativeMapiProbe.exe --help

powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 `
  -Target Rebuild
```

Live E2E требует реального Outlook profile и подготовленного тестового store, поэтому
не является универсальным тестом чистого clone:

```powershell
powershell.exe -NoProfile -File .\scripts\test_native_mapi_adapter_e2e.ps1
```

Последняя локальная проверка: 33/33 Python tests, подписанный VSTO Rebuild, direct
native x64 compile и native-to-service E2E — PASS.

## Честные ограничения

- Это workspace-прототип, а не готовый installer или portable package.
- VSTO работает только в classic Outlook; новый Outlook не поддерживается.
- Worker читает stores через Outlook MAPI providers, а не парсит PST/OST как файлы.
- Индексация пока выполняет полный scan; checkpoints, notifications и tombstones не
  реализованы.
- Если письмо после индексации перемещено или удалено, сохранённая пара
  `StoreID + EntryID` может перестать открываться; нужно обновить индекс.
- Список результатов является собственным WinForms UI надстройки, а не нативным
  списком Outlook; это осознанный выбор ради cross-store выдачи в порядке backend.
- Локальная neural-конфигурация необязательна и не воспроизводится без отдельно
  выбранных package/model artifacts.

Подробности эксперимента и Outlook Guard:
[docs/PROTOTYPE_NOTES.md](docs/PROTOTYPE_NOTES.md). Выбор способа показа результатов
и отвергнутый эксперимент с `Explorer.Search`:
[docs/OUTLOOK_SEARCH_PROJECTION.md](docs/OUTLOOK_SEARCH_PROJECTION.md).

## Лицензии и внешние компоненты

Аудит сделан для сценария закрытого платного продукта. В обязательном default
контуре не обнаружены `Non-Commercial`, GPL/AGPL или иные условия, запрещающие
продажу либо требующие открыть first-party код. Поэтому proprietary-лицензия на
собственный код в принципе совместима с текущим составом — при условии, что автор
владеет правами на contributions и явно исключает сторонние части из своей лицензии.

Это не превращает весь архив в «мою собственность»:

- MAPIStubLibrary headers остаются под MIT; их copyright и полный MIT-текст должны
  сопровождать исходную и бинарную поставку;
- две VSTO Utilities DLL остаются Microsoft redistributables и поставляются без
  изменений на условиях применимой лицензии Visual Studio;
- Windows, .NET Framework, VSTO Runtime и classic Outlook — внешние лицензируемые
  prerequisites, а не часть лицензии RAGSearch;
- optional `sentence-transformers`/model напрямую permissive, но их версии и
  транзитивный набор не закреплены, поэтому neural-поставка пока не прошла полный
  лицензионный аудит.

Для собственного кода корневой `LICENSE` пока не выбран. Полная матрица, границы
проверки и действия перед коммерческой поставкой:
[docs/LICENSE_COMPATIBILITY.md](docs/LICENSE_COMPATIBILITY.md). Файл, который нужно
класть рядом с бинарниками: [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

`RAGSearch` не аффилирован и не одобрен Microsoft. Microsoft и Outlook являются
товарными знаками Microsoft group of companies; здесь они используются только для
описания совместимости.
