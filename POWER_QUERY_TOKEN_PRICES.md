# Power Query: Автоматический импорт курсов токенов

## 📋 Описание

Это руководство содержит готовые Power Query M-скрипты для автоматического импорта курсов криптовалют из CoinGecko API в Excel.

## 🎯 Поддерживаемые токены

- **ETH** (Ethereum)
- **USDC** (USD Coin)
- **USDT** (Tether)
- **DAI** (Dai)
- **WBTC** (Wrapped Bitcoin)

## 📦 Требования

- **Excel 2016+** или **Microsoft 365**
- **Подключение к интернету** для обновления данных
- **Power Query** (встроен в Excel 2016+)

## 🚀 Быстрый старт

### Вариант 1: Простая таблица с текущими курсами

Этот запрос возвращает базовую таблицу с текущими курсами токенов в USD.

#### Power Query M-скрипт:

```m
let
    // Определяем список токенов для CoinGecko API
    TokenIds = "ethereum,usd-coin,tether,dai,wrapped-bitcoin",

    // Формируем URL для запроса
    Url = "https://api.coingecko.com/api/v3/simple/price?ids=" & TokenIds & "&vs_currencies=usd&include_last_updated_at=true",

    // Получаем данные от API
    Source = try Json.Document(Web.Contents(Url)) otherwise null,

    // Проверяем успешность запроса
    Result = if Source = null then
        #table(
            {"Токен", "Символ", "Цена USD", "Последнее обновление", "Статус"},
            {
                {"Ethereum", "ETH", null, null, "Ошибка подключения к API"},
                {"USD Coin", "USDC", null, null, "Ошибка подключения к API"},
                {"Tether", "USDT", null, null, "Ошибка подключения к API"},
                {"Dai", "DAI", null, null, "Ошибка подключения к API"},
                {"Wrapped Bitcoin", "WBTC", null, null, "Ошибка подключения к API"}
            }
        )
    else
        // Извлекаем данные для каждого токена
        let
            EthData = if Record.HasFields(Source, "ethereum") then Source[ethereum] else null,
            UsdcData = if Record.HasFields(Source, "usd-coin") then Source[#"usd-coin"] else null,
            UsdtData = if Record.HasFields(Source, "tether") then Source[tether] else null,
            DaiData = if Record.HasFields(Source, "dai") then Source[dai] else null,
            WbtcData = if Record.HasFields(Source, "wrapped-bitcoin") then Source[#"wrapped-bitcoin"] else null,

            // Создаём таблицу с результатами
            TableData = #table(
                {"Токен", "Символ", "Цена USD", "Последнее обновление", "Статус"},
                {
                    {
                        "Ethereum",
                        "ETH",
                        if EthData <> null then EthData[usd] else null,
                        if EthData <> null and Record.HasFields(EthData, "last_updated_at") then
                            #datetime(1970, 1, 1, 0, 0, 0) + #duration(0, 0, 0, EthData[last_updated_at])
                        else null,
                        if EthData <> null then "OK" else "Ошибка"
                    },
                    {
                        "USD Coin",
                        "USDC",
                        if UsdcData <> null then UsdcData[usd] else null,
                        if UsdcData <> null and Record.HasFields(UsdcData, "last_updated_at") then
                            #datetime(1970, 1, 1, 0, 0, 0) + #duration(0, 0, 0, UsdcData[last_updated_at])
                        else null,
                        if UsdcData <> null then "OK" else "Ошибка"
                    },
                    {
                        "Tether",
                        "USDT",
                        if UsdtData <> null then UsdtData[usd] else null,
                        if UsdtData <> null and Record.HasFields(UsdtData, "last_updated_at") then
                            #datetime(1970, 1, 1, 0, 0, 0) + #duration(0, 0, 0, UsdtData[last_updated_at])
                        else null,
                        if UsdtData <> null then "OK" else "Ошибка"
                    },
                    {
                        "Dai",
                        "DAI",
                        if DaiData <> null then DaiData[usd] else null,
                        if DaiData <> null and Record.HasFields(DaiData, "last_updated_at") then
                            #datetime(1970, 1, 1, 0, 0, 0) + #duration(0, 0, 0, DaiData[last_updated_at])
                        else null,
                        if DaiData <> null then "OK" else "Ошибка"
                    },
                    {
                        "Wrapped Bitcoin",
                        "WBTC",
                        if WbtcData <> null then WbtcData[usd] else null,
                        if WbtcData <> null and Record.HasFields(WbtcData, "last_updated_at") then
                            #datetime(1970, 1, 1, 0, 0, 0) + #duration(0, 0, 0, WbtcData[last_updated_at])
                        else null,
                        if WbtcData <> null then "OK" else "Ошибка"
                    }
                }
            )
        in
            TableData
in
    Result
```

### Вариант 2: Расширенная таблица с дополнительными метриками

Этот запрос возвращает подробную информацию о каждом токене, включая изменения цены за 24 часа и 7 дней.

#### Power Query M-скрипт:

```m
let
    // Определяем список токенов для CoinGecko API
    TokenIds = "ethereum,usd-coin,tether,dai,wrapped-bitcoin",

    // Формируем URL для запроса детальных данных
    Url = "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&ids=" & TokenIds & "&order=market_cap_desc&per_page=5&page=1&sparkline=false&price_change_percentage=24h,7d",

    // Получаем данные от API с обработкой ошибок
    Source = try Json.Document(Web.Contents(Url)) otherwise null,

    // Проверяем успешность запроса и преобразуем в таблицу
    Result = if Source = null then
        #table(
            {"Токен", "Символ", "Цена USD", "Изменение 24ч %", "Изменение 7д %", "Макс 24ч", "Мин 24ч", "Рыночная капитализация", "Объём торгов 24ч", "Последнее обновление", "Статус"},
            {
                {"Ethereum", "ETH", null, null, null, null, null, null, null, null, "Ошибка подключения к API"},
                {"USD Coin", "USDC", null, null, null, null, null, null, null, null, "Ошибка подключения к API"},
                {"Tether", "USDT", null, null, null, null, null, null, null, null, "Ошибка подключения к API"},
                {"Dai", "DAI", null, null, null, null, null, null, null, null, "Ошибка подключения к API"},
                {"Wrapped Bitcoin", "WBTC", null, null, null, null, null, null, null, null, "Ошибка подключения к API"}
            }
        )
    else
        let
            // Преобразуем JSON в таблицу
            ConvertedToTable = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
            ExpandedRecords = Table.ExpandRecordColumn(ConvertedToTable, "Column1",
                {"name", "symbol", "current_price", "price_change_percentage_24h_in_currency",
                 "price_change_percentage_7d_in_currency", "high_24h", "low_24h",
                 "market_cap", "total_volume", "last_updated"},
                {"Токен", "Символ", "Цена USD", "Изменение 24ч %", "Изменение 7д %",
                 "Макс 24ч", "Мин 24ч", "Рыночная капитализация", "Объём торгов 24ч", "Последнее обновление"}),

            // Преобразуем символ в верхний регистр
            UppercaseSymbol = Table.TransformColumns(ExpandedRecords, {{"Символ", Text.Upper, type text}}),

            // Добавляем колонку статуса
            AddedStatus = Table.AddColumn(UppercaseSymbol, "Статус", each "OK", type text),

            // Преобразуем типы данных
            ChangedTypes = Table.TransformColumnTypes(AddedStatus, {
                {"Цена USD", type number},
                {"Изменение 24ч %", type number},
                {"Изменение 7д %", type number},
                {"Макс 24ч", type number},
                {"Мин 24ч", type number},
                {"Рыночная капитализация", Int64.Type},
                {"Объём торгов 24ч", Int64.Type},
                {"Последнее обновление", type datetime}
            })
        in
            ChangedTypes
in
    Result
```

### Вариант 3: Компактная таблица для листа "Настройки и константы"

Этот запрос создаёт простую двухколоночную таблицу "Токен - Цена", идеальную для использования в формулах.

#### Power Query M-скрипт:

```m
let
    // Определяем список токенов
    TokenIds = "ethereum,usd-coin,tether,dai,wrapped-bitcoin",

    // Формируем URL
    Url = "https://api.coingecko.com/api/v3/simple/price?ids=" & TokenIds & "&vs_currencies=usd",

    // Получаем данные с обработкой ошибок
    Source = try Json.Document(Web.Contents(Url)) otherwise null,

    // Создаём компактную таблицу
    Result = if Source = null then
        #table(
            {"Токен", "Курс USD"},
            {
                {"ETH", 0},
                {"USDC", 0},
                {"USDT", 0},
                {"DAI", 0},
                {"WBTC", 0}
            }
        )
    else
        #table(
            {"Токен", "Курс USD"},
            {
                {"ETH", if Record.HasFields(Source, "ethereum") then Source[ethereum][usd] else 0},
                {"USDC", if Record.HasFields(Source, "usd-coin") then Source[#"usd-coin"][usd] else 0},
                {"USDT", if Record.HasFields(Source, "tether") then Source[tether][usd] else 0},
                {"DAI", if Record.HasFields(Source, "dai") then Source[dai][usd] else 0},
                {"WBTC", if Record.HasFields(Source, "wrapped-bitcoin") then Source[#"wrapped-bitcoin"][usd] else 0}
            }
        )
in
    Result
```

## 📖 Инструкция по настройке

### Шаг 1: Создание нового запроса Power Query

1. Откройте Excel
2. Перейдите на вкладку **"Данные"** (Data)
3. Нажмите **"Получить данные"** → **"Из других источников"** → **"Пустой запрос"**
   - Или нажмите **"Из Интернета"** → **"Дополнительно"** → **"Пустой запрос"**

### Шаг 2: Открытие расширенного редактора

1. В открывшемся окне Power Query Editor
2. На вкладке **"Главная"** нажмите **"Расширенный редактор"** (Advanced Editor)

### Шаг 3: Вставка M-скрипта

1. Удалите весь существующий код в редакторе
2. Скопируйте и вставьте один из M-скриптов выше (выберите подходящий вариант)
3. Нажмите **"Готово"** (Done)

### Шаг 4: Именование запроса

1. В правой панели **"Настройки запроса"** (Query Settings)
2. В поле **"Имя"** (Name) введите понятное имя, например:
   - `КурсыТокенов` (для варианта 1)
   - `КурсыТокеновДетально` (для варианта 2)
   - `КурсыТокеновКомпакт` (для варианта 3)

### Шаг 5: Загрузка данных в Excel

1. На вкладке **"Главная"** нажмите **"Закрыть и загрузить"** → **"Закрыть и загрузить в..."**
2. Выберите:
   - **"Таблица"** для создания новой таблицы
   - **"Существующий лист"** и укажите ячейку (например, A1 на листе "Настройки и константы")
3. Нажмите **"ОК"**

## 🔄 Настройка автоматического обновления

### Метод 1: Обновление при открытии файла

1. Щёлкните правой кнопкой мыши по таблице с данными
2. Выберите **"Таблица"** → **"Свойства внешних данных"**
3. Установите флажок **"Обновить данные при открытии файла"**
4. Нажмите **"ОК"**

### Метод 2: Периодическое автоматическое обновление

1. Щёлкните правой кнопкой мыши по таблице с данными
2. Выберите **"Таблица"** → **"Свойства внешних данных"**
3. Установите флажок **"Обновлять каждые"**
4. Укажите интервал обновления (например, 15 минут)
5. Нажмите **"ОК"**

> ⚠️ **Важно**: CoinGecko API имеет ограничения на количество запросов (rate limits). Для бесплатного доступа рекомендуется не обновлять данные чаще, чем раз в 1-2 минуты.

### Метод 3: Обновление всех запросов по расписанию

1. Перейдите на вкладку **"Данные"**
2. Нажмите **"Обновить всё"** → **"Свойства подключения"**
3. Настройте параметры обновления для всех подключений

### Метод 4: Ручное обновление

- Щёлкните правой кнопкой по таблице → **"Обновить"**
- Или нажмите **Alt + F5** для обновления текущей таблицы
- Или на вкладке **"Данные"** → **"Обновить всё"** (Ctrl + Alt + F5)

## 🛠️ Расширенные настройки

### Добавление дополнительных токенов

Чтобы добавить другие токены, измените переменную `TokenIds` в скрипте:

```m
TokenIds = "ethereum,usd-coin,tether,dai,wrapped-bitcoin,bitcoin,binancecoin",
```

Найдите ID токена на сайте CoinGecko:
1. Перейдите на страницу токена на coingecko.com
2. ID токена находится в URL: `https://www.coingecko.com/en/coins/{token-id}`

### Изменение валюты котировки

По умолчанию цены указаны в USD. Чтобы изменить валюту, измените параметр `vs_currencies`:

```m
// Для EUR
Url = "https://api.coingecko.com/api/v3/simple/price?ids=" & TokenIds & "&vs_currencies=eur",

// Для нескольких валют одновременно (требует изменения скрипта)
Url = "https://api.coingecko.com/api/v3/simple/price?ids=" & TokenIds & "&vs_currencies=usd,eur,btc",
```

### Настройка таймаутов и повторных попыток

Для более надёжной работы с API можно добавить параметры:

```m
Source = try Json.Document(
    Web.Contents(
        Url,
        [
            Timeout = #duration(0, 0, 0, 30),  // Таймаут 30 секунд
            IsRetry = true                      // Разрешить повторные попытки
        ]
    )
) otherwise null,
```

## ⚠️ Обработка ошибок

Все предоставленные скрипты включают встроенную обработку ошибок:

1. **При недоступности API**: Возвращаются нулевые значения или таблица с сообщением об ошибке
2. **При отсутствии данных для токена**: Используется значение `null` или `0`
3. **При сетевых проблемах**: Используется конструкция `try ... otherwise null`

### Диагностика проблем

Если запрос не работает:

1. **Проверьте подключение к интернету**
2. **Проверьте доступность API**: Откройте в браузере:
   ```
   https://api.coingecko.com/api/v3/ping
   ```
   Должен вернуться JSON: `{"gecko_says":"(V3) To the Moon!"}`

3. **Проверьте настройки безопасности Excel**:
   - Файл → Параметры → Центр управления безопасностью → Параметры центра управления безопасностью
   - Параметры внешнего содержимого → Разрешить содержимое данных

4. **Проверьте лимиты API**: CoinGecko Free API позволяет ~10-50 запросов в минуту

## 📊 Интеграция с листом "Настройки и константы"

### Рекомендуемая структура

На листе "Настройки и константы" создайте таблицу следующего вида:

| Токен | Курс USD | Последнее обновление |
|-------|----------|---------------------|
| ETH   | 3091.94  | 20.01.2026 10:30    |
| USDC  | 0.9997   | 20.01.2026 10:30    |
| USDT  | 0.9989   | 20.01.2026 10:30    |
| DAI   | 0.9999   | 20.01.2026 10:30    |
| WBTC  | 90660    | 20.01.2026 10:30    |

### Использование в формулах

После настройки Power Query таблицы, вы можете использовать функции Excel для получения курсов:

```excel
// Используя ВПР (VLOOKUP)
=ВПР("ETH"; КурсыТокеновКомпакт; 2; ЛОЖЬ)

// Используя ИНДЕКС и ПОИСКПОЗ (INDEX/MATCH)
=ИНДЕКС(КурсыТокеновКомпакт[Курс USD]; ПОИСКПОЗ("ETH"; КурсыТокеновКомпакт[Токен]; 0))

// В русской версии Excel
=ИНДЕКС(КурсыТокеновКомпакт[Курс USD]; ПОИСКПОЗ("ETH"; КурсыТокеновКомпакт[Токен]; 0))
```

### Создание именованных диапазонов

Для удобства создайте именованные диапазоны:

1. Выделите ячейку с ценой ETH
2. В поле "Имя" (слева от строки формул) введите `Курс_ETH`
3. Нажмите Enter

Теперь в формулах можно использовать просто `=Курс_ETH * Количество`

## 🔐 Безопасность и конфиденциальность

- **API ключ не требуется**: CoinGecko Free API не требует регистрации
- **Данные публичные**: Передаются только публичные рыночные данные
- **Нет персональной информации**: Скрипты не отправляют никаких личных данных

## 📈 Альтернативные API

Если CoinGecko API недоступен, можно использовать альтернативы:

### CoinMarketCap API

```m
// Требует бесплатный API ключ с coinmarketcap.com
let
    ApiKey = "YOUR_API_KEY_HERE",
    Url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol=ETH,USDC,USDT,DAI,WBTC&convert=USD",
    Source = Json.Document(Web.Contents(Url, [Headers=[#"X-CMC_PRO_API_KEY"=ApiKey]]))
in
    Source
```

### CryptoCompare API

```m
let
    Url = "https://min-api.cryptocompare.com/data/pricemulti?fsyms=ETH,USDC,USDT,DAI,WBTC&tsyms=USD",
    Source = Json.Document(Web.Contents(Url))
in
    Source
```

## 🎓 Дополнительные ресурсы

- [CoinGecko API Documentation](https://www.coingecko.com/en/api/documentation)
- [Power Query M Language Specification](https://learn.microsoft.com/en-us/powerquery-m/)
- [Excel Power Query Tutorial (Microsoft)](https://support.microsoft.com/en-us/office/about-power-query-in-excel-7104fbee-9e62-4cb9-a02e-5bfb1a6c536a)

## 💡 Советы и лучшие практики

1. **Резервное копирование**: Всегда сохраняйте резервную копию файла перед изменением Power Query
2. **Тестирование**: Проверьте работу запроса на отдельном листе перед интеграцией
3. **Частота обновления**: Не устанавливайте обновление чаще 1 раза в минуту из-за лимитов API
4. **Мониторинг**: Регулярно проверяйте статус обновления данных
5. **Документирование**: Оставьте комментарии в M-скрипте для будущих изменений

## 🐛 Известные проблемы и решения

### Проблема: "DataSource.Error: Web.Contents failed to get contents"

**Решение**:
- Проверьте подключение к интернету
- Убедитесь, что firewall не блокирует запросы к api.coingecko.com
- Попробуйте обновить запрос позже (возможно временные проблемы с API)

### Проблема: "Expression.Error: The name 'Json.Document' wasn't recognized"

**Решение**:
- Убедитесь, что используете Power Query (доступен в Excel 2016+)
- Проверьте, что скопировали весь скрипт полностью

### Проблема: Данные не обновляются автоматически

**Решение**:
- Проверьте настройки обновления в свойствах таблицы
- Убедитесь, что Excel не находится в безопасном режиме
- Проверьте, что макросы и внешние подключения разрешены

## ✅ Чек-лист выполнения задачи

- [x] Изучены возможности Power Query в Excel
- [x] Настроено подключение к API CoinGecko
- [x] Создан запрос для автоматического обновления курсов основных токенов (ETH, USDC, USDT, DAI, WBTC)
- [x] Описано настройка расписания автоматического обновления данных
- [x] Протестирована корректность получаемых данных
- [x] Добавлена обработка ошибок при недоступности API
- [x] Созданы три варианта скриптов для разных сценариев использования
- [x] Добавлена подробная документация по настройке и использованию

---

**Версия документа**: 1.0
**Дата создания**: 20 января 2026
**Автор**: AI Issue Solver

**Готово к использованию! 🚀**
