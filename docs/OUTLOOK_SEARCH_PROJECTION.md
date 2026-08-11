# Агрегированная выдача в Outlook All Mailboxes

RAGSearch проецирует результаты локального hybrid search в штатный центральный
список classic Outlook через:

```csharp
explorer.Search(finalAqs, Outlook.OlSearchScope.olSearchScopeAllFolders);
```

`olSearchScopeAllFolders` означает почтовые папки того же типа во всех stores,
которые пользователь включил в Outlook `Locations to Search`. На проверенной
машине один native rowset одновременно содержал письма из Inbox, Sent и PST
`Archives`. Deleted Items Outlook по умолчанию в эту область не включает.

## Поток запроса

`NativeSearchPresenter` привязан к тому же `Explorer`, которому принадлежит task
pane; после HTTP он не переходит к случайному `Application.ActiveExplorer()`.

1. Перед запросом presenter запоминает текущие folder/view и версию состояния
   Explorer. Предыдущая выдача остаётся на экране до готовности новой, поэтому
   HTTP-await не мигает исходным Inbox и ошибка сервиса не уничтожает результаты.
2. VSTO отправляет сервису `filters: {}` и `limit: 12`, поэтому поиск идёт по
   всему локальному индексу, а не по текущему `folder_entry_id`, и service limit
   совпадает с максимальным числом Outlook AQS-условий.
3. После await ответ применяется только если пользователь не сменил folder/view
   и не взаимодействовал со списком.
4. Из результатов строится компактный OR: тема каждого результата становится
   обычной кавыченной фразой, максимум 12 условий
   и 2400 символов. Повторяющиеся темы дедуплицируются.
5. Финальный AQS атомарно заменяет предыдущий Instant Search и запускается с
   `olSearchScopeAllFolders`; кнопка сброса вызывает `Explorer.ClearSearch()` и
   возвращает обычную папку.

Пустая service-выдача не вызывает `ClearSearch`, иначе Outlook снова показал бы
все письма исходной папки. Вместо этого запускается заведомо несуществующая
sentinel-фраза и агрегированный список остаётся пустым до сброса.

## Почему в Search bar виден запрос

Microsoft определяет `Explorer.Search` как действие, эквивалентное вводу строки
пользователем в Instant Search UI. У метода нет отдельного results object и нет
callback завершения. Поэтому одновременно получить поддержанный единый
cross-store native rowset и скрыть строку поиска через Outlook Object Model
нельзя.

Проверенный эксперимент с широким All Mailboxes search и последующим
`CurrentView.Filter` это не обошёл: `View.Filter` сохранился, но агрегированная
таблица осталась 14/14 строк. Instant Search provider проигнорировал folder-view
filter. Такой fail-open путь не используется.

Canonical Windows AQS `System.Subject:=...` на ru-RU Windows + en-US classic
Outlook был синтаксически принят, но вернул 0 строк. Локализованные property
keywords также оказались нестабильны. Поэтому текущий финальный запрос состоит
из обычных кавыченных фраз, полученных из subject результатов и соединённых
документированным `OR`. Он короче
прежней конструкции `subject/from/received` и реально работает на целевой
установке.

## Cutoff и literal gate

Сервис агрегирует лучший chunk на уровне сообщения до top-K. Для фразы он
объединяет literal FTS hits с semantic candidates, сортирует vector-часть по
`vector_distance = 1 - cosine_similarity`, применяет model-specific floor и
adaptive window. API hard cap равен 25 сообщениям; VSTO All Mailboxes projection
запрашивает 12, чтобы не получить результаты, которые не поместятся в AQS.

Для одиночного слова действует защита от наблюдавшейся патологии multilingual
embedding model:

- exact token, prefix и substring ищутся через два FTS5 индекса (`unicode61` и
  `trigram`);
- любой literal hit проходит независимо от dense cutoff;
- prefix/substring включает literal gate и скрывает посторонние dense guesses;
- если literal hit отсутствует, single-token semantic guess не показывается.

Поэтому `киберспорт`, `кибер` и `спорт` находят индексированное
`киберспорт`, даже если dense model ставит короткое нерелевантное слово выше.
Schema v2 автоматически rebuild-ит trigram FTS для существующих chunks; повторно
читать PST/OST не требуется.

## Точность identity и порядок строк

Текущая cross-store AQS-проекция приблизительна. Outlook ищет кавыченную фразу
по своему индексу, поэтому может показать дополнительные копии/письма с той же
фразой в теме или теле. Поддержанного Outlook AQS equality по бинарным
`(StoreID, EntryID)` нет; Internet Message-ID также не имеет документированного
Instant Search keyword.

Порядок OR-условий не задаёт порядок строк. Backend возвращает `rank`,
`lexical_match_kind`, `vector_similarity`, `vector_distance` и `hybrid_score`,
но Outlook сортирует aggregate view по собственному relevance/date. Точный
vector order возможен только в собственном result list либо после записи custom
property в письма, чего read-only прототип не делает.

| Механизм | Несколько stores | Пустой Search bar | Exact identity / vector order |
|---|---:|---:|---:|
| `Explorer.Search` / All Mailboxes | да | нет | нет / нет |
| `View.Filter` | нет, одна папка | да | приблизительно / нет |
| `AdvancedSearch.Save` / Search Folder | нет, один store | да | приблизительно / нет |
| Extended MAPI Search Folder | нет, один store | да | да по `PR_RECORD_KEY` / нет |
| собственный result list | да | да | да / да, но не native Outlook list |

## Документация Microsoft

- [Explorer.Search](https://learn.microsoft.com/en-us/office/vba/api/outlook.explorer.search)
- [OlSearchScope](https://learn.microsoft.com/en-us/office/vba/api/outlook.olsearchscope)
- [Use Instant Search across all folders and stores](https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-use-instant-search-to-search-all-folders-and-all-stores-for-a-phrase-in-the-subject)
- [Advanced Query Syntax](https://learn.microsoft.com/en-us/windows/win32/search/-search-3x-advancedquerysyntax)
- [View.Filter](https://learn.microsoft.com/en-us/office/vba/api/outlook.view.filter)
- [Application.AdvancedSearch](https://learn.microsoft.com/en-us/office/vba/api/outlook.application.advancedsearch)
- [PidTagRecordKey](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagrecordkey-canonical-property)
- [Searching a message store](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/searching-a-message-store)
