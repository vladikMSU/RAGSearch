# Собственная cross-store выдача RAGSearch в Outlook

## Принятое решение

RAGSearch **не проецирует** результаты в штатный список classic Outlook. Ответ
локального search service показывается в собственном WinForms `DataGridView`
внутри VSTO Custom Task Pane, закреплённого снизу окна Outlook.

Это важная граница текущей реализации:

- Search bar Outlook не заполняется и не очищается;
- текущая папка, `Explorer.CurrentView` и центральный список писем не меняются;
- один список RAG Search может одновременно содержать письма из OST и PST;
- строки сохраняют порядок, в котором их вернул backend;
- каждая строка несёт opaque locator с точной парой `StoreID + EntryID`, а не
  приблизительный AQS; backend хранит locator, но не интерпретирует его.

## Поток запроса и отображения

1. Пользователь вводит запрос в нижней панели и нажимает **«Найти»** или `Enter`.
2. VSTO отправляет `/v1/search` с текстом запроса и `limit: 25`. Поиск идёт по
   всему локальному индексу, а не только по текущей папке Outlook.
3. Service возвращает уже ранжированный массив `results`.
4. `SearchPaneControl.PopulateResults` последовательно добавляет элементы массива
   в таблицу. Дополнительной клиентской сортировки нет; сортировка колонок
   пользователем отключена.
5. Таблица показывает rank, тему с фрагментом, отправителя, дату и путь
   `store · folder`. Пустой ответ очищает только таблицу RAG Search.

Лимит UI и hard cap API одинаковы — 25 писем. Protocol 4 требует положительный
`rank` у каждой строки; некорректный ответ отклоняется целиком.

Панель использует отдельные светлую и тёмную WinForms-палитры. Начальная палитра
выбирается при создании control по Windows `AppsUseLightTheme`; таблица, selection,
input, buttons и settings menu используют согласованные цвета. Панель можно
свернуть или отделить в плавающее Custom Task Pane, не затрагивая Outlook view.

Основная реализация:

- [`SearchLimit` и построение UI](../hosts/outlook_vsto/SearchPaneControl.cs#L14);
- [`SearchAsync`](../hosts/outlook_vsto/SearchPaneControl.cs#L798);
- [`PopulateResults`](../hosts/outlook_vsto/SearchPaneControl.cs#L898);
- [нижнее размещение Custom Task Pane](../hosts/outlook_vsto/ThisAddIn.cs#L40).

## Exact identity и открытие оригинала

Search response возвращает для каждого документа сохранённый при индексации
`locator`. VSTO проверяет, что это locator connector-а `outlook_mapi`, извлекает
`entry_id` и `store_id`, после чего двойной щелчок по строке или клавиша `Enter`
вызывает на Outlook UI/STA thread:

```csharp
session.GetItemFromID(result.LocatorEntryId, result.LocatorStoreId);
```

Полученный `MailItem` открывается через `Display(false)` в настоящем окне Outlook.
Так открывается именно выбранная backend-строка, даже если одинаковая тема есть в
нескольких папках или stores. Python service при этом остаётся source-neutral:
поля Outlook существуют только внутри locator и интерпретируются host-ом.
Реализация находится в
[`ThisAddIn.OpenSearchResult`](../hosts/outlook_vsto/ThisAddIn.cs#L93).

Точность относится к состоянию store на момент индексации. Если письмо после этого
переместили, удалили либо отключили его PST, Outlook может больше не разрешить
сохранённую пару. Панель показывает ошибку и предлагает обновить индекс.

## Отвергнутый эксперимент: `Explorer.Search`

Предыдущий прототип содержал `NativeSearchPresenter`. Он преобразовывал темы
backend-результатов в OR из кавыченных фраз и вызывал:

```csharp
explorer.Search(finalAqs, Outlook.OlSearchScope.olSearchScopeAllFolders);
```

Эксперимент подтвердил, что `All Mailboxes` способен собрать native rowset из
Inbox, Sent и PST `Archives`, но не удовлетворяет требованиям RAG-выдачи:

- по контракту `Explorer.Search` запрос становится виден в Instant Search bar;
- AQS по тексту темы приблизителен и может вернуть другое письмо или копию;
- бинарные `StoreID + EntryID` нельзя выразить как поддержанное AQS equality;
- порядок OR-условий не управляет строками — Outlook сортирует rowset сам;
- последующий `CurrentView.Filter` не ограничил агрегированную таблицу (в live
  проверке осталось 14/14 строк);
- AQS property syntax зависела от locale/provider: `System.Subject:=...` на
  проверенной ru-RU Windows + en-US Outlook вернул пустую выдачу.

Из-за этих ограничений `NativeSearchPresenter` удалён из production-проекта.
`Explorer.Search`, `Explorer.ClearSearch` и `View.Filter` больше не входят в поток
семантического поиска. Это не запасной режим, а завершённый отвергнутый эксперимент.

| Механизм | Несколько stores | Search bar без изменений | Exact identity | Порядок backend |
|---|---:|---:|---:|---:|
| `Explorer.Search` / All Mailboxes | да | нет | нет | нет |
| `View.Filter` | нет, одна папка | да | нет | нет |
| `AdvancedSearch.Save` / Search Folder | нет, один store | да | приблизительно | нет |
| Extended MAPI Search Folder | нет, один store | да | да по `PR_RECORD_KEY` | нет |
| собственный WinForms result list | да | да | да, `StoreID + EntryID` | да |

## Ranking до UI

Сервис агрегирует лучший chunk на уровне сообщения до top-K. Для фразы он
объединяет literal FTS hits с semantic candidates, сортирует vector-часть по
`vector_distance = 1 - cosine_similarity`, применяет model-specific floor и
adaptive window. UI не пересчитывает ranking и не переставляет результаты.

Для одиночного слова действует защита от наблюдавшейся патологии multilingual
embedding model:

- exact token, prefix и substring ищутся через два FTS5 индекса (`unicode61` и
  `trigram`);
- любой literal hit проходит независимо от dense cutoff;
- prefix/substring включает literal gate и скрывает посторонние dense guesses;
- если literal hit отсутствует, single-token semantic guess не показывается.

Поэтому `киберспорт`, `кибер` и `спорт` находят индексированное `киберспорт`, даже
если dense model ставит короткое нерелевантное слово выше. Service требует
нейтральную document schema v4; более старую Outlook-shaped локальную БД нужно
удалить и заново построить из PST/OST.

## Документация Microsoft для отвергнутого пути

- [Explorer.Search](https://learn.microsoft.com/en-us/office/vba/api/outlook.explorer.search)
- [OlSearchScope](https://learn.microsoft.com/en-us/office/vba/api/outlook.olsearchscope)
- [Use Instant Search across all folders and stores](https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-use-instant-search-to-search-all-folders-and-all-stores-for-a-phrase-in-the-subject)
- [Advanced Query Syntax](https://learn.microsoft.com/en-us/windows/win32/search/-search-3x-advancedquerysyntax)
- [View.Filter](https://learn.microsoft.com/en-us/office/vba/api/outlook.view.filter)
- [Application.AdvancedSearch](https://learn.microsoft.com/en-us/office/vba/api/outlook.application.advancedsearch)
- [PidTagRecordKey](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagrecordkey-canonical-property)
- [Searching a message store](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/searching-a-message-store)
