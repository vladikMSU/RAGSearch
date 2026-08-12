# RAGSearch

Экспериментальный локальный поиск по почте, совместимый с classic Microsoft Outlook
для Windows x64.
VSTO-надстройка показывает собственную нижнюю WinForms-панель поиска и результатов,
напрямую управляет отдельным C++ reader-ом для read-only Extended MAPI и отправляет
нейтральные документы в Python search service. Штатные строка поиска и список
писем Outlook при этом не изменяются.

```text
                                     ┌── запуск/stop ──► C++ Extended MAPI reader ──► Outlook profile
classic Outlook + нижняя VSTO-панель ┤                    │
          │                          │◄── Outlook JSONL ──┘
          │                          └── neutral HTTP /v1/documents + /v1/search ──► Python service
          │                                                                       │
          │ locator: StoreID + EntryID                                SQLite + FTS5 + embeddings
          ▼
   исходное письмо
```

Между C++ и Python нет прямой связи, общего каталога обмена или adapter-процесса.
VSTO последовательно преобразует source-specific JSONL в `Document`; Outlook IDs
находятся только в opaque `locator`, который сервис хранит и возвращает без
интерпретации.

Extended MAPI здесь не является скачанным Python-пакетом. Во время сборки
используются закреплённые в репозитории официальные Microsoft headers; линковочные
библиотеки даёт Windows SDK, а MAPI provider во время работы — classic Outlook.
Подробнее: [архитектура и границы компонентов](docs/ARCHITECTURE_AND_MAPI.md).

## Что уже находится в репозитории

- исходники VSTO-надстройки, native x64 reader и Python-сервиса;
- шесть необходимых headers Microsoft MAPIStubLibrary, их MIT-лицензия и точный
  upstream commit в [third_party/MAPIStubLibrary](third_party/MAPIStubLibrary);
- dependency-free Python provider и тесты — для базового режима сторонние
  Python-пакеты и скачивание модели не нужны;
- единая MSBuild solution; native EXE собирается в
  `connectors\outlook_mapi\native\bin\x64\<Configuration>\OutlookMapiReader.exe`.

Source-specific native reader находится в `connectors/outlook_mapi`, а
`service/ragsearch_service` содержит только source-neutral локальный search
service. VSTO host находится в `hosts/outlook_vsto`; прежние Python adapter и
экспериментальная Outlook Object Model диагностика из production tree удалены.

В Git намеренно не хранятся `.venv`, build artifacts, база, письма, временные
вложения, токены, private signing keys и необязательная neural model.

## Требования

Для компиляции из чистого clone:

- Windows x64;
- Visual Studio 2022 с workload `Office/SharePoint development`;
- `.NET Framework 4.8 Targeting Pack`;
- workload `Desktop development with C++`, MSVC v143 x64/x86 и Windows 10/11 SDK.

Для запуска дополнительно нужны Python 3.11+, VSTO Runtime, classic Outlook x64 и
настроенный default Outlook profile. Разрядность native reader и Outlook обязана
совпадать; solution предоставляет только платформу x64.

## Сборка из чистого clone

Команды ниже выполняются из корня репозитория. Чтобы собранная надстройка затем
могла запустить сервис, создайте ожидаемое ею Python-окружение; базовый режим
использует только стандартную библиотеку Python:

```powershell
py -3 --version # должно быть 3.11+
py -3 -m venv .\service\.venv
```

Соберите оба проекта одной командой. Скрипт находит Visual Studio через `vswhere`,
собирает native reader и создаёт либо повторно использует локальный non-exportable
development certificate в `Cert:\CurrentUser\My`, потому что полный VSTO `Build`
технически требует подписанный manifest:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

Дополнительные MAPI-файлы скачивать или указывать через include path не нужно.
`RAGSearch.sln` содержит VSTO host и `OutlookMapiReader`; одна конфигурация
`Debug|x64` или `Release|x64` собирает оба компонента.

Сертификат автора и private key в Git не входят. Для повторной сборки можно оставить
созданный скриптом локальный dev certificate. Чтобы использовать свой уже
установленный code-signing certificate с private key:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 `
  -Configuration Release `
  -CertificateThumbprint YOUR_CERTIFICATE_THUMBPRINT
```

Автоматический self-signed certificate разрешён скриптом только для `Debug`.
`-Configuration Release` требует явно переданный production certificate.

Скрипт решает воспроизводимость development-сборки, но не доверие к издателю.
Не коммитьте `.pfx`, private key или локальный thumbprint. Для распространения нужен
нормальный installer и сертификат, которому доверяет целевая организация;
репозиторий их не содержит. Автоматически созданный сертификат также может
потребовать явного доверия в соответствии с локальной ClickOnce/Office policy.
Прямой диагностический запуск reader-а описан в
[connectors/outlook_mapi/README.md](connectors/outlook_mapi/README.md).

## Запуск

Python-сервис можно запустить вручную:

```powershell
Push-Location .\service
.\.venv\Scripts\python.exe -m ragsearch_service
Pop-Location
```

После сборки и регистрации надстройки перезапустите classic Outlook. Если сервис
не запущен, панель сама использует точный workspace-путь
`service\.venv\Scripts\python.exe`; при отсутствии локальной neural model включается
воспроизводимый hashing provider без дополнительных зависимостей.
Health handshake проверяет protocol `4`: если порт 8765 ещё занят старым процессом
RAGSearch service, панель сообщает о несовместимости вместо отправки нового DTO в
старый endpoint.
Непустой индекс привязан к одной embedding-модели и размерности: при смене
конфигурации сервис требует явной очистки индекса и полной переиндексации.

Результаты появляются в закреплённой снизу таблице RAG Search в том же порядке,
в котором их вернул backend; максимум — 25 писем. Двойной щелчок по строке открывает
исходный Outlook `MailItem` по точной паре `StoreID + EntryID` из opaque locator.
Панель подбирает
светлую или тёмную WinForms-палитру по системной теме, умеет сворачиваться и
отделяться в плавающее окно. Обычные Search bar, текущая папка и список Outlook
остаются без изменений.

Reader можно запустить отдельно для диагностики его MAPI/JSONL boundary. Такой
запуск ничего не отправляет в search service:

```powershell
.\connectors\outlook_mapi\native\bin\x64\Debug\OutlookMapiReader.exe `
  --jsonl --max-stores 1 --max-folders 10 --max-messages 5 `
  --body-preview-chars 2000
```

Необязательные `sentence-transformers` и multilingual model не включены в Git и
никогда не скачиваются сервисом автоматически. Их локальная настройка описана в
[service/README.md](service/README.md).

## Проверки

```powershell
.\service\.venv\Scripts\python.exe -B -m unittest `
  discover -s service\tests -t service -v

.\connectors\outlook_mapi\native\bin\x64\Debug\OutlookMapiReader.exe --help

powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\connectors\outlook_mapi\tests\test_reader_smoke.ps1 `
  -Configuration Debug -OfflineOnly

powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 `
  -Target Rebuild
```

Live MAPI-проверка требует реального Outlook profile и поэтому не является
универсальным тестом чистого clone; её команда и bounded параметры приведены в
connector README. Актуальные результаты автоматических тестов и сборки фиксируются
после каждого изменения контракта, а не поддерживаются как вручную подсчитанная
цифра в документации.

## Честные ограничения

- Это workspace-прототип, а не готовый installer или portable package.
- VSTO работает только в classic Outlook; новый Outlook не поддерживается.
- Reader читает stores через Outlook MAPI providers, а не парсит PST/OST как файлы.
- Индексация пока выполняет полный scan; checkpoints, notifications и tombstones не
  реализованы.
- Если письмо после индексации перемещено или удалено, сохранённая пара
  `StoreID + EntryID` может перестать открываться; нужно обновить индекс.
- Список результатов является собственным WinForms UI надстройки, а не нативным
  списком Outlook; это осознанный выбор ради cross-store выдачи в порядке backend.
- Локальная neural-конфигурация необязательна и не воспроизводится без отдельно
  выбранных package/model artifacts.

Граница отображения результатов и причины не использовать `Explorer.Search`:
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
- optional `sentence-transformers==5.7.0`/model напрямую permissive, но
  транзитивные artifacts, hashes и revision модели не закреплены, поэтому
  neural-поставка пока не прошла полный лицензионный аудит.

Для собственного кода корневой `LICENSE` пока не выбран. Полная матрица, границы
проверки и действия перед коммерческой поставкой:
[docs/LICENSE_COMPATIBILITY.md](docs/LICENSE_COMPATIBILITY.md). Файл, который нужно
класть рядом с бинарниками: [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

`RAGSearch` не аффилирован и не одобрен Microsoft. Microsoft и Outlook являются
товарными знаками Microsoft group of companies; здесь они используются только для
описания совместимости.
