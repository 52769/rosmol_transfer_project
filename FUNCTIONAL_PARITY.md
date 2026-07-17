# Контроль функциональной полноты

- Исходный `transfer.py` сохранён без изменений как `legacy_transfer.py`.
- GUI запускает ускоренный адаптер, который импортирует исходное ядро, а не заменяет его.
- Все определения верхнего уровня исходника проверяются `selftest.py` через AST.
- При интерактивных сценариях используется один поток, поэтому консольные выборы не конфликтуют.
- Через поле ответа GUI доступны все вопросы исходного кода: выбор аккаунта, проекта, даты, существующего черновика и действия после ошибки.
- Старые BAT-сценарии представлены отдельными кнопочными/консольными запусками.
- Ускорение не меняет payload, парсер Word, API-клиент, правила валидации или форматы отчётов.


## v2.6 browser runtime

All original transfer functions remain unchanged. The GUI/worker launch layer now sets an explicit Playwright browser directory and the build script provisions portable Chromium under `dist/ProjectTransfer/ms-playwright`.


## v2.7 UTF-8 transport

The transfer core remains byte-identical to the legacy copy. Only the launch/transport layer changed: the worker emits UTF-8, the GUI sends UTF-8 input, and old CP866/CP1251 output is decoded through a readability fallback.
