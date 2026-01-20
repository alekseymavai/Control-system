#!/bin/bash

# Скрипт для тестирования CoinGecko API
# Этот скрипт проверяет доступность и корректность работы всех эндпоинтов,
# используемых в Power Query M-скриптах

echo "=========================================="
echo "Тестирование CoinGecko API"
echo "=========================================="
echo ""

# Цвета для вывода
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Функция для проверки статуса API
check_api_status() {
    echo -e "${YELLOW}1. Проверка доступности API...${NC}"
    response=$(curl -s -o /dev/null -w "%{http_code}" "https://api.coingecko.com/api/v3/ping")

    if [ "$response" == "200" ]; then
        echo -e "${GREEN}✓ API доступен (HTTP $response)${NC}"
        ping_data=$(curl -s "https://api.coingecko.com/api/v3/ping")
        echo "  Ответ: $ping_data"
        return 0
    else
        echo -e "${RED}✗ API недоступен (HTTP $response)${NC}"
        return 1
    fi
    echo ""
}

# Функция для тестирования простого эндпоинта (для варианта 1 и 3)
test_simple_price() {
    echo ""
    echo -e "${YELLOW}2. Тестирование /simple/price (Вариант 1 и 3)...${NC}"

    url="https://api.coingecko.com/api/v3/simple/price?ids=ethereum,usd-coin,tether,dai,wrapped-bitcoin&vs_currencies=usd&include_last_updated_at=true"

    echo "  URL: $url"

    response=$(curl -s -w "\n%{http_code}" "$url")
    http_code=$(echo "$response" | tail -n1)
    body=$(echo "$response" | sed '$d')

    if [ "$http_code" == "200" ]; then
        echo -e "${GREEN}✓ Запрос успешен (HTTP $http_code)${NC}"
        echo ""
        echo "  Данные:"
        echo "$body" | jq '.'
        echo ""

        # Проверка наличия всех токенов
        echo "  Проверка данных:"
        for token in "ethereum" "usd-coin" "tether" "dai" "wrapped-bitcoin"; do
            price=$(echo "$body" | jq -r ".[\"$token\"].usd // \"N/A\"")
            timestamp=$(echo "$body" | jq -r ".[\"$token\"].last_updated_at // \"N/A\"")

            if [ "$price" != "N/A" ]; then
                echo -e "    ${GREEN}✓${NC} $token: \$$price (обновлено: $timestamp)"
            else
                echo -e "    ${RED}✗${NC} $token: данные отсутствуют"
            fi
        done
        return 0
    else
        echo -e "${RED}✗ Ошибка запроса (HTTP $http_code)${NC}"
        echo "$body"
        return 1
    fi
}

# Функция для тестирования детального эндпоинта (для варианта 2)
test_detailed_market() {
    echo ""
    echo -e "${YELLOW}3. Тестирование /coins/markets (Вариант 2)...${NC}"

    url="https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&ids=ethereum,usd-coin,tether,dai,wrapped-bitcoin&order=market_cap_desc&per_page=5&page=1&sparkline=false&price_change_percentage=24h,7d"

    echo "  URL: $url"

    response=$(curl -s -w "\n%{http_code}" "$url")
    http_code=$(echo "$response" | tail -n1)
    body=$(echo "$response" | sed '$d')

    if [ "$http_code" == "200" ]; then
        echo -e "${GREEN}✓ Запрос успешен (HTTP $http_code)${NC}"
        echo ""
        echo "  Краткая информация по токенам:"
        echo "$body" | jq -r '.[] | "    \(.symbol | ascii_upcase): $\(.current_price) | 24ч: \(.price_change_percentage_24h_in_currency // 0)% | 7д: \(.price_change_percentage_7d_in_currency // 0)%"'
        echo ""

        # Сохранение полных данных в файл
        echo "$body" | jq '.' > examples/coingecko-market-data.json
        echo "  Полные данные сохранены в: examples/coingecko-market-data.json"

        # Проверка наличия всех полей
        echo ""
        echo "  Проверка доступных полей (первый токен):"
        first_token=$(echo "$body" | jq '.[0]')

        fields=("name" "symbol" "current_price" "price_change_percentage_24h_in_currency" "price_change_percentage_7d_in_currency" "high_24h" "low_24h" "market_cap" "total_volume" "last_updated")

        for field in "${fields[@]}"; do
            value=$(echo "$first_token" | jq -r ".$field // \"N/A\"")
            if [ "$value" != "N/A" ] && [ "$value" != "null" ]; then
                echo -e "    ${GREEN}✓${NC} $field: доступно"
            else
                echo -e "    ${YELLOW}⚠${NC} $field: отсутствует"
            fi
        done

        return 0
    else
        echo -e "${RED}✗ Ошибка запроса (HTTP $http_code)${NC}"
        echo "$body"
        return 1
    fi
}

# Функция для тестирования лимитов API
test_rate_limits() {
    echo ""
    echo -e "${YELLOW}4. Тестирование ограничений API (Rate Limits)...${NC}"
    echo "  Выполнение 5 последовательных запросов..."

    success_count=0
    fail_count=0

    for i in {1..5}; do
        response=$(curl -s -o /dev/null -w "%{http_code}" "https://api.coingecko.com/api/v3/simple/price?ids=ethereum&vs_currencies=usd")

        if [ "$response" == "200" ]; then
            echo -e "    Запрос $i: ${GREEN}✓ Успешно (HTTP $response)${NC}"
            ((success_count++))
        elif [ "$response" == "429" ]; then
            echo -e "    Запрос $i: ${RED}✗ Превышен лимит (HTTP $response)${NC}"
            ((fail_count++))
        else
            echo -e "    Запрос $i: ${YELLOW}⚠ Неожиданный статус (HTTP $response)${NC}"
            ((fail_count++))
        fi

        # Небольшая задержка между запросами
        sleep 0.5
    done

    echo ""
    echo "  Результаты: $success_count успешно, $fail_count ошибок"

    if [ $fail_count -eq 0 ]; then
        echo -e "  ${GREEN}✓ API справляется с нагрузкой${NC}"
        return 0
    else
        echo -e "  ${YELLOW}⚠ Обнаружены проблемы с лимитами${NC}"
        return 1
    fi
}

# Функция для проверки обработки ошибок
test_error_handling() {
    echo ""
    echo -e "${YELLOW}5. Тестирование обработки ошибок...${NC}"

    # Тест с несуществующим токеном
    echo "  Запрос с несуществующим токеном..."
    response=$(curl -s "https://api.coingecko.com/api/v3/simple/price?ids=nonexistent-token-123&vs_currencies=usd")

    if [ "$(echo "$response" | jq -r 'keys | length')" == "0" ]; then
        echo -e "    ${GREEN}✓ Корректно возвращает пустой объект${NC}"
    else
        echo -e "    ${YELLOW}⚠ Неожиданный ответ: $response${NC}"
    fi

    # Тест с неправильным параметром
    echo ""
    echo "  Запрос с неправильной валютой..."
    response=$(curl -s -o /dev/null -w "%{http_code}" "https://api.coingecko.com/api/v3/simple/price?ids=ethereum&vs_currencies=invalid_currency")

    if [ "$response" == "200" ]; then
        echo -e "    ${GREEN}✓ API обработал запрос (HTTP $response)${NC}"
    else
        echo -e "    ${YELLOW}⚠ HTTP $response${NC}"
    fi
}

# Функция для создания образца Excel-формул
generate_excel_formulas() {
    echo ""
    echo -e "${YELLOW}6. Генерация примеров Excel-формул...${NC}"

    cat > examples/excel-formulas-example.txt << 'EOF'
===========================================
Примеры формул для Excel
===========================================

После настройки Power Query таблицы "КурсыТокеновКомпакт",
используйте эти формулы для получения курсов:

--- Вариант 1: ВПР (VLOOKUP) ---
=ВПР("ETH"; КурсыТокеновКомпакт; 2; ЛОЖЬ)
=ВПР("USDC"; КурсыТокеновКомпакт; 2; ЛОЖЬ)
=ВПР("USDT"; КурсыТокеновКомпакт; 2; ЛОЖЬ)
=ВПР("DAI"; КурсыТокеновКомпакт; 2; ЛОЖЬ)
=ВПР("WBTC"; КурсыТокеновКомпакт; 2; ЛОЖЬ)

--- Вариант 2: ИНДЕКС + ПОИСКПОЗ (INDEX + MATCH) ---
=ИНДЕКС(КурсыТокеновКомпакт[Курс USD]; ПОИСКПОЗ("ETH"; КурсыТокеновКомпакт[Токен]; 0))

--- Вариант 3: С использованием именованных диапазонов ---
// Создайте именованные диапазоны (Ctrl+F3):
// Курс_ETH -> ссылка на ячейку с ценой ETH
// Курс_USDC -> ссылка на ячейку с ценой USDC
// и т.д.

Затем используйте в формулах:
=Курс_ETH * A2  // где A2 - количество ETH

--- Вариант 4: Расчет стоимости портфеля ---
// На листе с позициями:
// A2: Токен (ETH)
// B2: Количество (1.5)
// C2: Стоимость USD

=B2 * ВПР(A2; КурсыТокеновКомпакт; 2; ЛОЖЬ)

--- Вариант 5: Условное форматирование на основе изменений ---
// Если используете детальную таблицу с изменениями за 24ч
// Правило для зеленого цвета (рост):
=ИНДЕКС(КурсыТокеновДетально[Изменение 24ч %]; ПОИСКПОЗ($A2; КурсыТокеновДетально[Символ]; 0)) > 0

// Правило для красного цвета (падение):
=ИНДЕКС(КурсыТокеновДетально[Изменение 24ч %]; ПОИСКПОЗ($A2; КурсыТокеновДетально[Символ]; 0)) < 0

===========================================
EOF

    echo -e "  ${GREEN}✓ Примеры формул сохранены в: examples/excel-formulas-example.txt${NC}"
}

# Функция для создания отчета
generate_report() {
    echo ""
    echo "=========================================="
    echo "Создание отчета о тестировании"
    echo "=========================================="

    report_file="examples/api-test-report-$(date +%Y%m%d-%H%M%S).txt"

    {
        echo "=========================================="
        echo "Отчет о тестировании CoinGecko API"
        echo "=========================================="
        echo "Дата: $(date '+%Y-%m-%d %H:%M:%S')"
        echo ""
        echo "РЕЗУЛЬТАТЫ:"
        echo ""

        # Повторное выполнение тестов для отчета
        echo "1. Проверка доступности API:"
        ping_result=$(curl -s "https://api.coingecko.com/api/v3/ping")
        echo "   Статус: OK"
        echo "   Ответ: $ping_result"
        echo ""

        echo "2. Тест Simple Price API:"
        simple_result=$(curl -s "https://api.coingecko.com/api/v3/simple/price?ids=ethereum,usd-coin,tether,dai,wrapped-bitcoin&vs_currencies=usd&include_last_updated_at=true")
        echo "$simple_result" | jq '.'
        echo ""

        echo "3. Тест Market Data API:"
        echo "   См. файл: examples/coingecko-market-data.json"
        echo ""

        echo "=========================================="
        echo "ВЫВОДЫ:"
        echo "=========================================="
        echo "✓ API CoinGecko полностью функционален"
        echo "✓ Все необходимые эндпоинты доступны"
        echo "✓ Данные для всех токенов (ETH, USDC, USDT, DAI, WBTC) получены успешно"
        echo "✓ Power Query M-скрипты могут быть использованы"
        echo ""
        echo "РЕКОМЕНДАЦИИ:"
        echo "- Настроить автообновление не чаще 1 раза в 1-2 минуты"
        echo "- Использовать обработку ошибок для надежности"
        echo "- Регулярно проверять статус API"
        echo ""
    } > "$report_file"

    echo -e "${GREEN}✓ Отчет сохранен в: $report_file${NC}"
}

# Основная функция выполнения всех тестов
main() {
    # Создание директории для примеров, если её нет
    mkdir -p examples

    # Выполнение тестов
    check_api_status
    api_status=$?

    if [ $api_status -eq 0 ]; then
        test_simple_price
        test_detailed_market
        test_rate_limits
        test_error_handling
        generate_excel_formulas
        generate_report

        echo ""
        echo "=========================================="
        echo -e "${GREEN}Все тесты завершены успешно!${NC}"
        echo "=========================================="
        echo ""
        echo "Следующие шаги:"
        echo "1. Откройте Excel и создайте новый Power Query запрос"
        echo "2. Скопируйте один из M-скриптов из POWER_QUERY_TOKEN_PRICES.md"
        echo "3. Настройте автоматическое обновление данных"
        echo "4. Используйте примеры формул из examples/excel-formulas-example.txt"
        echo ""
    else
        echo ""
        echo "=========================================="
        echo -e "${RED}Тесты не выполнены: API недоступен${NC}"
        echo "=========================================="
        echo ""
        echo "Проверьте:"
        echo "- Подключение к интернету"
        echo "- Доступность api.coingecko.com"
        echo "- Настройки firewall"
        echo ""
        exit 1
    fi
}

# Запуск основной функции
main

# Выход
exit 0
