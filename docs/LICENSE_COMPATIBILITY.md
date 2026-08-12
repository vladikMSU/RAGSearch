# Совместимость лицензий и коммерческая поставка

Дата аудита: 2026-08-12. Это техническая инвентаризация зависимостей, а не
юридическое заключение. Целевой сценарий проверки — закрытая платная поставка
RAGSearch для Windows без передачи прав на first-party исходный код.

## Короткий вывод

В обязательном контуре текущего репозитория не обнаружены лицензии с запретом
коммерческого использования (`Non-Commercial`) или reciprocal/copyleft-условиями
вроде GPL/AGPL, которые потребовали бы открыть first-party код RAGSearch.
Проприетарная коммерческая лицензия на собственный код поэтому в принципе
совместима с найденным составом, если одновременно выполняются три условия:

1. автор действительно владеет правами на весь first-party код и contributions;
2. сторонние компоненты явно исключены из области проприетарной лицензии;
3. их собственные notices и условия поставки сохранены.

Иными словами, весь продукт можно продавать как единое предложение, но нельзя
объявить Microsoft headers/DLL своей собственностью или отменить их лицензии своим
`LICENSE`/EULA.

## Фактический состав

| Компонент | Попадает в Git или build output | Лицензия/режим | Совместимость с закрытой коммерческой поставкой |
|---|---|---|---|
| First-party код RAGSearch | Да | Корневой `LICENSE` пока отсутствует | Владелец прав может выбрать proprietary или open-source условия только для своего кода. Нужно учесть права всех contributors. |
| 6 headers Microsoft MAPIStubLibrary | Да, `third_party/MAPIStubLibrary` | MIT, pinned commit `a9505d7…` | Да. Разрешены использование, изменение, распространение и продажа; copyright и полный MIT-текст нужно сохранить в исходной и бинарной поставке. Сами headers остаются MIT. |
| `Microsoft.Office.Tools.Common.v4.0.Utilities.dll` и `Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll` | Да, копируются в VSTO output без изменений | Microsoft Visual Studio redistributables | Да, только без изменений и с соблюдением применимых условий лицензии Visual Studio. Это не first-party DLL. |
| .NET Framework, VSTO Runtime, Windows, classic Outlook и установленный MAPI provider | Нет, внешние prerequisites | Условия Microsoft для соответствующих продуктов | Лицензия RAGSearch на них не распространяется. Пользователь/организация должны иметь законные установки; их нельзя молча включить в свой installer как собственные файлы. |
| `MAPI32.lib`/`Ole32.lib`, MSVC и Windows SDK | Используются при сборке; import libraries не поставляются как часть RAGSearch | Условия Windows SDK/Visual Studio | Готовый native EXE может быть proprietary. Текущий Release EXE статически включает MSVC runtime code; право использовать toolchain и распространять результат определяется условиями лицензированной редакции Visual Studio/SDK. |
| Python interpreter и standard library | Не включены; пользователь устанавливает Python отдельно | PSF License и notices incorporated software | Текущая поставка не перераспространяет Python. Если позже встроить interpreter, нужно приложить весь Python license stack и notices точной версии. |
| `sentence-transformers` и транзитивные packages | Не включены; только необязательная ручная установка | Прямой проект — Apache-2.0, но dependency range не закреплён | Apache-2.0 сама по себе допускает proprietary/commercial use, но текущий транзитивный набор не зафиксирован и не прошёл полный distribution audit. Не включать в коммерческий installer до lock/SBOM/notices-аудита. |
| `paraphrase-multilingual-MiniLM-L12-v2` | Не включена и не скачивается автоматически | В model card указана Apache-2.0 | В принципе permissive, но перед поставкой нужно закрепить точную revision, сохранить LICENSE/model card и отдельно оценить происхождение данных/модели; license tag не является гарантией отсутствия иных рисков. |
| CMake, PowerShell и Visual Studio | Только build tools | Их собственные условия | Их лицензии обычно не переходят на результат сборки, но сама организация обязана иметь право пользоваться выбранной редакцией инструмента. Для Visual Studio Community есть ограничения по типу и размеру организации. |

## Блокеры и неизвестные перед коммерческим релизом

1. **Лицензия Visual Studio.** Community допускает создание платных приложений
   индивидуальным разработчиком и ограниченное использование небольшой
   организацией. Для иных сценариев, в частности proprietary-разработки в enterprise
   organization, нужна надлежащим образом лицензированная Professional/Enterprise
   или иная разрешённая toolchain. Outlook/VSTO add-in не следует считать
   автоматически попадающим под исключение для open-source/Visual Studio extensions.
2. **EULA конечного продукта.** При распространении VSTO Utilities условия для
   пользователя/дистрибьютора должны защищать Microsoft code как минимум в
   требуемом лицензией Visual Studio объёме, запрещать ложное представление
   Microsoft ownership/endorsement и не накладывать на DLL несовместимую
   `Excluded License`. Это нужно согласовать вместе с first-party EULA.
3. **Release toolchain.** Нынешний локальный native EXE — проверочный dev artifact,
   собранный не штатным v143 `.vcxproj`; он не входит в Git и не считается очищенным
   release binary. Коммерческий EXE нужно заново собрать штатной поддерживаемой
   toolchain под действующей лицензией и повторно записать imports/CRT/SBOM.
4. **Подпись релиза.** Автоматически создаваемый self-signed VSTO certificate и
   неподписанные DLL/EXE годятся только для разработки. Для внешнего релиза нужны
   production code-signing certificate, installer и политика доверия. Это не
   лицензионный copyleft-блокер, но реальный blocker поставки. `scripts/build.ps1`
   поэтому отказывается собирать VSTO `Release` без явно переданного certificate.
5. **Название и trademarks.** Коммерческое имя продукта должно быть `RAGSearch`, а
   Microsoft Outlook следует упоминать только описательно: «совместимо с classic
   Microsoft Outlook». Не использовать Microsoft/Outlook logo и не создавать
   впечатление аффилированности или одобрения Microsoft.
6. **Contributors.** Git-лицензия не исправляет отсутствие прав у автора. Перед
   proprietary-релизом нужно подтвердить авторство/assignment/CLA для всех
   contributions и сторонних фрагментов, если они появятся.

## Что поставлять вместе с default build

Для текущего контура без bundled Python и neural model:

- first-party `LICENSE`/EULA с явным исключением third-party material;
- `THIRD_PARTY_NOTICES.md`;
- `third_party/MAPIStubLibrary/LICENSE` как
  `MAPIStubLibrary-LICENSE.txt`;
- две VSTO Utilities DLL без модификации;
- installer, который проверяет наличие законно установленного .NET Framework,
  VSTO Runtime и classic Outlook, а не копирует Office/MAPI provider из чужой
  установки.

Сборочные проекты автоматически копируют notice и MIT-текст в output. Это помогает
не потерять обязательные файлы, но не заменяет аудит конечного installer/archive.

## Что сделать перед включением neural-контура

1. Выбрать конкретные Python и `sentence-transformers` versions.
2. Создать lock с hashes и SBOM для всех wheels/native libraries.
3. Собрать LICENSE/NOTICE для каждой транзитивной зависимости, включая bundled
   компоненты NumPy/SciPy/PyTorch.
4. Закрепить точный commit/revision модели и сохранить её model card/license.
5. Повторить аудит при каждом обновлении lock или модели.

До этого dependency-free hashing provider — единственный контур, для которого
состав поставки в репозитории полностью определён.

Если позднее поставлять именно локально проверенную `sentence-transformers 5.7.0`,
у неё кроме Apache-2.0 `LICENSE` есть upstream `NOTICE`, который также надо сохранить.
Model card локальной копии указывает Apache-2.0, но отдельного LICENSE-файла рядом
нет; для дистрибуции следует приложить полный Apache-2.0 текст и закрепить revision.

## Зафиксированные evidence текущей dev-сборки

Две VSTO Utilities DLL версии `10.0.30319.1` побайтно совпали с установленными
Visual Studio reference assemblies и имеют Microsoft signature:

- Common SHA-256: `8CB09317C326E9B0F83C337EAE7CCDEAAD3E45E5DA3603E1EBC90C5A06AD1702`;
- Outlook SHA-256: `018F5FE2880C5419CDA8D2AF19CD0AA3C5375EC20378B854FDCE63932EF1D997`.

Это evidence конкретной сборки, а не вечный allow-list: после обновления Visual
Studio hashes могут законно измениться и должны заново попасть в release SBOM.

## Официальные основания

- [MIT-текст vendored MAPIStubLibrary](../third_party/MAPIStubLibrary/LICENSE)
- [Microsoft: Visual Studio 2022 redistribution list](https://learn.microsoft.com/visualstudio/releases/2022/redistribution)
- [Microsoft: условия использования Visual Studio Community](https://visualstudio.microsoft.com/vs/community/)
- [Microsoft: Visual Studio Community 2022 license terms](https://visualstudio.microsoft.com/license-terms/vs2022-ga-community/)
- [Microsoft trademark and brand guidelines](https://www.microsoft.com/legal/intellectualproperty/trademarks)
- [Python Software Foundation License](https://docs.python.org/3/license.html)
- [sentence-transformers LICENSE (Apache-2.0)](https://github.com/huggingface/sentence-transformers/blob/main/LICENSE)
- [model card `paraphrase-multilingual-MiniLM-L12-v2`](https://huggingface.co/sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2)

При выборе окончательной коммерческой EULA/лицензии стоит дать юристу именно эту
матрицу, pinned artifacts и состав готового installer, а не только исходный код.
