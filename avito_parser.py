import asyncio
import re
import json
import os
from urllib.parse import urlencode
from playwright.async_api import async_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# ⚙️ ОСНОВНЫЕ НАСТРОЙКИ
# ==========================================
PRICE_MIN = 7000          
PRICE_MAX = 14000         
MAX_PAGES  = 20           
OUTPUT_FILE = "avito_results.xlsx" 
USER_CITY = "ростов-на-дону"

# Конфигурация кэша
CACHE_FILE = "avito_cache.json"

def load_cache() -> dict:
    """Load cached ad details from CACHE_FILE if it exists."""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_cache(cache: dict) -> None:
    """Persist the cache dictionary to CACHE_FILE."""
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)       

SEARCH_QUERIES = [
    "motorola",
    "nothing",
    "google pixel",
    "oneplus",
]

BASE_URL = "https://www.avito.ru/rossiya/telefony/smartfony-ASgBAgICAUSSA8YQ"

def make_url(query: str, page: int = 1) -> str:
    """Генерирует ссылку для поиска на Авито с фильтрами по цене и доставке."""
    params = {"q": query, "pmin": PRICE_MIN, "pmax": PRICE_MAX, "cd": 1}
    if page > 1: params["p"] = page
    return f"{BASE_URL}?{urlencode(params)}"

async def is_captcha_page(page) -> bool:
    """Проверяет, не заблокировал ли нас Авито капчей (по заголовку страницы)."""
    return "ограничен" in (await page.title()).lower()

async def fetch_description(page, url: str) -> dict:
    """
    Заходит в конкретное объявление и вытягивает характеристики (Состояние, Экран, АКБ, Рейтинг, Год регистрации).
    Использует JavaScript для поиска текста по всей странице, обходя скрытые теги Авито.
    """
    result = {
        "description": "Описание не найдено", "condition": "Не указано", 
        "screen": "Не указано", "battery": "—", "rating": 0.0, "reviews": 0, "reg_year": 0
    }
    
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=15_000)
        await asyncio.sleep(3)
        
        # Если вылезла капча - ждем, пока пользователь ее решит руками
        if await is_captcha_page(page):
            print(f"\n⚠ Капча при переходе. Реши в браузере...")
            for _ in range(60):
                await asyncio.sleep(3)
                if not await is_captcha_page(page): break
            else:
                result["description"] = "Капча при загрузке"
                return result

        # Парсим текст описания объявления
        desc_el = await page.query_selector('[data-marker="item-view/item-description"]') or await page.query_selector('[itemprop="description"]')
        if desc_el:
            result["description"] = (await desc_el.inner_text()).strip()

        # Выполняем JS-скрипт внутри страницы для поиска нужных данных в любом месте текста
        js_extractor = r"""
        () => {
            let res = {condition: "Не указано", screen: "Не указано", battery: "—", rating: 0.0, reviews: 0, reg_year: 0};
            let text = document.body.innerText;
            
            // Ищем Состояние и Экран
            let condMatch = text.match(/Состояние:?\s*(Новое|Отличное|Хорошее|Удовлетворительное|Требуется ремонт|На запчасти)/i);
            if (condMatch) res.condition = condMatch[1].charAt(0).toUpperCase() + condMatch[1].slice(1).toLowerCase();
            
            let screenMatch = text.match(/Экран:?\s*(Без дефектов|1[-–]2 мелкие царапины|Много мелких царапин|Глубокие царапины|Трещины|Засветы|Выгорание|Полосы и битые пиксели|Не работает)/i);
            if (screenMatch) res.screen = screenMatch[1].charAt(0).toUpperCase() + screenMatch[1].slice(1).toLowerCase().replace('1-2', '1–2');
            
            // Ищем процент АКБ
            let batMatch = text.match(/(?:акб|батарея|емкость|состояние аккумулятора)[\s:-]*(\d{2,3})\s*%/i);
            if (batMatch) res.battery = batMatch[1] + "%";

            // Ищем рейтинг и отзывы
            let ratingMatch = text.match(/(\d[,.]\d)\s*(?:★\s*)?\n?\s*(\d+)\s+отзыв/i);
            if (ratingMatch) {
                res.rating = parseFloat(ratingMatch[1].replace(',', '.'));
                res.reviews = parseInt(ratingMatch[2]);
            }
            
            // Ищем дату регистрации аккаунта (например: "На Авито с апреля 2023")
            let regMatch = text.match(/На Авито с [а-яА-Я]+ (\d{4})/i);
            if (regMatch) res.reg_year = parseInt(regMatch[1]);

            return res;
        }
        """
        data = await page.evaluate(js_extractor)
        if data: result.update(data)

    except Exception as e:
        result["description"] = f"Ошибка: {e}"
    
    return result


async def parse_page(page, url: str, context) -> list[dict]:
    """Сканирует общую страницу поиска, собирает карточки товаров и делает первичный фильтр."""
    await page.goto(url, wait_until="domcontentloaded", timeout=30_000)
    await asyncio.sleep(3)

    if await is_captcha_page(page): return []

    try:
        await page.wait_for_selector('[data-marker="item"]', timeout=15_000)
    except:
        return [] 

    cards = await page.query_selector_all('[data-marker="item"]')
    items = []

    for card in cards:
        try:
            # Извлекаем базовые данные с карточки (название, цена, город, ссылка)
            title_el = await card.query_selector('[itemprop="name"]') or await card.query_selector('[data-marker="item-title"]')
            title = (await title_el.inner_text()).strip() if title_el else ""

            price_el = await card.query_selector('[itemprop="price"]')
            price = int(await price_el.get_attribute("content") or 0) if price_el else 0

            link_el = await card.query_selector('[data-marker="item-title"]')
            href = await link_el.get_attribute("href") if link_el else ""
            url_item = f"https://www.avito.ru{href}" if href and href.startswith("/") else href

            geo_el = await card.query_selector('[data-marker="item-address"]')
            city = (await geo_el.inner_text()).strip().split(",")[0].strip() if geo_el else ""

            # --- ПЕРВИЧНЫЙ ФИЛЬТР ---
            # Отсекаем мусор и запчасти по названию
            STOP_WORDS = [
                "чехол", "коробка", "аккумулятор", "акб", "дисплей", "экран", "стекло", "камера",
                "плата", "запчасти", "на запчасти", "шлейф", "корпус", "динамик", "зарядка", 
                "ремонт", "наушники", "watch", "buds", "часы", "разбит", "трещина", "выгорание"
            ]
            
            title_lower = title.lower()
            if any(w in title_lower for w in STOP_WORDS): continue

            # Строгий фильтр по нужным моделям телефонов
            is_valid_model = False
            
            # 1. Nothing Phone (Ловит слитное и раздельное написание)
            if "nothing" in title_lower or "phone" in title_lower:
                if " 1 " in title_lower or "cmf" in title_lower: pass 
                # \s* позволяет пропускать пробелы внутри моделей: "2a" == "2 a"
                elif re.search(r'(nothing|phone)\s*\(?(2|2\s*a|2\s*pro|2pro|3|3\s*a|3\s*pro|3pro|4\s*a|4\s*pro|4pro)\b', title_lower):
                    is_valid_model = True
            
            # 2. Google Pixel
            elif "pixel" in title_lower or "пиксель" in title_lower:
                if re.search(r'(pixel|пиксель)\s*(8|9|9\s*pro|9pro|9\s*a|9a)\b', title_lower):
                    is_valid_model = True
            
            # 3. Motorola (Добавил \s* во все модели)
            elif any(brand in title_lower for brand in ["motorola", "moto", "моторола"]):
                allowed_moto = [
                    r"edge\s*30\s*pro", r"edge\s*x30", r"edge\s*40", r"\bx\s*40\b", r"\bs\s*30\b", 
                    r"\bs\s*50\b", r"\bs\s*60\b", r"edge\s*50", r"edge\s*20\s*pro", r"edge\s*60", r"edge\s*70"
                ]
                if any(re.search(m, title_lower) for m in allowed_moto):
                    is_valid_model = True

            # 4. OnePlus (Добавил \s* для Nord, Ace и моделей)
            elif "oneplus" in title_lower or "ванплас" in title_lower:
                # Замени в функции parse_page:
                op_regex = r'(oneplus|ванплас)\s+(11\s*r|12|12\s*r|nord\s*3|nord\s*4|nord\s*5|ace\s*2|ace\s*2\s*pro|ace\s*2\s*v|ace\s*3|ace\s*3\s*v|ace\s*3\s*pro|nord\s*ce\s*3)\b'
                if re.search(op_regex, title_lower):
                    is_valid_model = True
            
            if not is_valid_model: continue

            # Добавляем в список только подходящие по цене варианты
            if price and (PRICE_MIN <= price <= PRICE_MAX):
                items.append({"title": title, "price": price, "city": city, "url": url_item})

        except Exception:
            continue

    return items


def save_to_excel(data: list[dict], filename: str):
    """Сохраняет результаты в Excel с раскраской ячеек и автофильтрами."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Авито — ТОП Смартфоны"

    DARK, ACCENT, LIGHT_ROW, WHITE, GREEN, TEXT_DARK, GOLD, RED = "111827", "6366F1", "F9FAFB", "FFFFFF", "D1FAE5", "1F2937", "F59E0B", "FEE2E2"
    border = Border(left=Side(style="thin", color="E5E7EB"), right=Side(style="thin", color="E5E7EB"), 
                    top=Side(style="thin", color="E5E7EB"), bottom=Side(style="thin", color="E5E7EB"))

    # Шапка
    cols = ["№", "Оценка", "Название", "Цена (₽)", "Город", "Рейтинг", "АКБ", "Комплект", "Состояние", "Экран", "Описание", "Ссылка"]
    ws.merge_cells(f"A1:{chr(64+len(cols))}1")
    header_main = ws["A1"]
    header_main.value = f"🎯 СНАЙПЕРСКАЯ ПОДБОРКА │ {PRICE_MIN:,} – {PRICE_MAX:,} ₽"
    header_main.font = Font(name="Calibri", bold=True, size=14, color=WHITE)
    header_main.fill = PatternFill(start_color=DARK, end_color=DARK, fill_type="solid")
    header_main.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Названия колонок
    col_widths = [5, 10, 35, 12, 18, 22, 10, 15, 15, 22, 60, 25]
    for i, (col_name, width) in enumerate(zip(cols, col_widths), 1):
        cell = ws.cell(row=2, column=i, value=col_name)
        cell.fill = PatternFill(start_color=ACCENT, end_color=ACCENT, fill_type="solid")
        cell.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        ws.column_dimensions[cell.column_letter].width = width
    ws.row_dimensions[2].height = 22

    # Заполнение строк
    row_num = 3
    for idx, item in enumerate(data, 1):
        desc = item.get("description", "—")
        if len(desc) > 300: desc = desc[:297] + "..." 
        
        # --- ФОРМАТИРОВАНИЕ РЕЙТИНГА И ДАТЫ РЕГИСТРАЦИИ ---
        rating_text = f"{item['rating']} ({item['reviews']} отз.)"
        rating_color = None
        
        if item['reviews'] == 0:
            if item['reg_year'] == 0:
                rating_text = "⚠ Нет отзывов"
            elif item['reg_year'] >= 2024:
                rating_text = f"⚠ Опасно (Свежий акк {item['reg_year']})"
                rating_color = RED
            else:
                rating_text = f"Без отзывов (С {item['reg_year']} г.)"
                rating_color = "FEF08A" # Желтый/Оранжевый (старый аккаунт, но без отзывов)

        row_data = [idx, item["score"], item["title"], item["price"], item["city"], rating_text, 
                    item["battery"], item["kit"], item["condition"], item["screen"], desc, item["url"]]

        for col, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col, value=value)
            cell.font = Font(name="Calibri", size=10, color=TEXT_DARK)
            cell.alignment = Alignment(vertical="center", wrap_text=(col in [3, 11]))
            cell.border = border
            if idx % 2 == 0: cell.fill = PatternFill(start_color=LIGHT_ROW, end_color=LIGHT_ROW, fill_type="solid")

        # Раскраска Оценки (зеленый для топов)
        score_cell = ws.cell(row=row_num, column=2)
        score_cell.font = Font(bold=True, size=12, color="FFFFFF" if item["score"] >= 5 else TEXT_DARK)
        score_cell.fill = PatternFill(start_color="10B981" if item["score"] >= 5 else "D1D5DB", fill_type="solid")
        score_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=row_num, column=4).font = Font(bold=True, color="065F46")
        ws.cell(row=row_num, column=4).number_format = '#,##0 ₽'
        ws.cell(row=row_num, column=7).alignment = Alignment(horizontal="center", vertical="center") 
        
        if item["kit"] == "✅ Полный":
            ws.cell(row=row_num, column=8).font = Font(color="059669", bold=True)
            
        # Раскраска рейтинга (красный/желтый)
        if rating_color:
            ws.cell(row=row_num, column=6).fill = PatternFill(start_color=rating_color, end_color=rating_color, fill_type="solid")
            if rating_color == RED: ws.cell(row=row_num, column=6).font = Font(color="B91C1C", bold=True)

        url_cell = ws.cell(row=row_num, column=12)
        url_cell.hyperlink = item["url"]
        url_cell.font = Font(color="4F46E5", underline="single")
        url_cell.value = "Открыть"

        ws.row_dimensions[row_num].height = 45  
        row_num += 1

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{chr(64+len(cols))}{row_num-1}"
    wb.save(filename)
    return len(data)


# ==========================================
# 🚀 ГЛАВНЫЙ АЛГОРИТМ ПАРСЕРА
# ==========================================
async def main():
    all_items = []
    print("🚀 Запуск парсера: СНАЙПЕРСКИЙ РЕЖИМ (С ОЦЕНКАМИ)")

    # Load cache
    cache = load_cache()
    cached_used = 0
    new_added = 0

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False, args=["--disable-blink-features=AutomationControlled"])
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            locale="ru-RU", viewport={"width": 1366, "height": 900})
        await context.add_init_script("Object.defineProperty(navigator, 'webdriver', { get: () => undefined });")
        page = await context.new_page()

        # --- ЭТАП 1: СБОР ССЫЛОК СО СТРАНИЦ ПОИСКА ---
        for query in SEARCH_QUERIES:
            print(f"\n🔍 Поиск: «{query}»")
            empty_attempts = 0
            
            for page_num in range(1, MAX_PAGES + 1):
                url = make_url(query, page_num)
                print(f"  Страница {page_num}... ", end="", flush=True)
                items = await parse_page(page, url, context)
                print(f"Найдено: {len(items)}")
                
                if len(items) == 0:
                    empty_attempts += 1
                    if empty_attempts >= 2:
                        print("  🛑 Пустые страницы, идем дальше.")
                        break
                else:
                    all_items.extend(items)
                    empty_attempts = 0
                
                await asyncio.sleep(3)
        
        # --- ЭТАП 2: Анализ каждого объявления ---
        filtered_items = []
        # Deduplicate items by URL
        seen = set()
        unique_items = []
        for it in all_items:
            if it["url"] not in seen:
                seen.add(it["url"])                
                unique_items.append(it)
        
        if unique_items:
            print("📝 Анализ: Скрытые дефекты, Доставка, Комплект, Рейтинг...")
            for idx, item in enumerate(unique_items, 1):
                url = item["url"]
                if url in cache:
                    # Use cached data
                    cached_item = cache[url]
                    item.update(cached_item)
                    cached_used += 1
                    print(f"  [{idx}/{len(unique_items)}] {item['title'][:20]}... (cached)")
                else:
                    print(f"  [{idx}/{len(unique_items)}] {item['title'][:20]}... ", end="", flush=True)
                    details = await fetch_description(page, url)
                    item.update(details)
                    
                    desc_lower = item["description"].lower()
                    title_lower = item["title"].lower()
                    
                    # 🛑 Фильтр скрытых проблем (скам и блокировки)
                    stop_issues = [
                        "mdm", "мдм", "soft unlock", "софт анлок", "операторский", "демо", "demo", "att", "verizon",
                        "вскрывался", "ремонтировался", "заменен", "менялся", "не работает отпечаток", "отвал",
                        "мерцает", "заблокирован", "icloud", "гугл аккаунт", "google аккаунт", "пятн", "требуется ремонт"
                    ]
                    if any(bad in desc_lower for bad in stop_issues):
                        print("✕ (Отбраковано: Скрытый дефект/Блокировка)")
                        continue
                    # Дополнительная проверка: исключаем объявления с пометкой "Только обмен"
                    if "только обмен" in desc_lower:
                        print("✕ (Отбраковано: Только обмен)")
                        continue

                    # 🛑 Фильтр доставки (Только Ростов)
                    no_delivery = ["без доставки", "доставки нет", "авито доставки нет", "не отправляю", "только личная встреча", "не отправлю", "без пересыла"]
                    if any(word in desc_lower for word in no_delivery):
                        if USER_CITY not in item["city"].lower():
                            print("✕ (Отбраковано: Нет доставки, другой город)")
                            continue

                    # 🛑 Фильтр убитого состояния
                    cond_lower = item["condition"].lower()
                    scr_lower = item["screen"].lower()
                    if cond_lower != "не указано" and not any(good in cond_lower for good in ["отлично", "хороше", "ново"]):
                        print("✕ (Убитое состояние)")
                        continue
                    if scr_lower != "не указано" and not any(good in scr_lower for good in ["без дефект", "1-2", "1–2", "царапин"]):
                        print("✕ (Разбитый экран)")
                        continue

                    # 🛑 Фильтр рейтинга (менее 4.0)
                    if item["rating"] > 0 and item["rating"] < 4.0:
                        print("✕ (Плохой продавец, рейтинг ниже 4.0)")
                        continue

                    # 🎁 Анализ комплекта
                    if re.search(r'полный комплект|коробка|родная зарядка|ориг\\w* блок', desc_lower):
                        item["kit"] = "✅ Полный"
                    else:
                        item["kit"] = "❓ Уточнить"

                    # 🏆 Расчет умной оценки (Скоринг)
                    score = 0
                    if item["condition"] == "Отличное" and item["screen"] == "Без дефектов":
                        score += 3
                    if item["kit"] == "✅ Полный":
                        score += 2
                    
                    # ДЖЕКПОТЫ (+5 баллов)
                    jackpots = [
                        r'(nothing|phone)\\s*\\(?(4\\s*a|4\\s*pro)\\b',
                        r'(pixel|пиксель)\\s*(9\\s*pro|9pro)\\b',
                        r'(edge 60|s60|edge 70)',
                        r'(oneplus|ванплас)\\s+(12|12r|nord 4|nord 5|ace 3 pro)\\b'
                    ]
                    if any(re.search(j, title_lower) for j in jackpots):
                        score += 5
                    
                    item["score"] = score
                    print(f"✅ Оценка: {score} | АКБ: {item['battery']}")
                    filtered_items.append(item)
                    cache[url] = item.copy()
                    new_added += 1
                    await asyncio.sleep(3)
        
        await browser.close()
    
    # Prune stale cache entries (remove URLs no longer present)
    current_urls = {it["url"] for it in all_items}
    stale_keys = [k for k in list(cache.keys()) if k not in current_urls]
    for k in stale_keys:
        del cache[k]
    
    save_cache(cache)
    print(f"🗂️ Кешировано объявлений: {len(cache)} (использовано {cached_used}, новые {new_added})")

    # Финальная сортировка и сохранение
    if filtered_items:
        filtered_items.sort(key=lambda x: (-x["score"], x["price"]))
        total = save_to_excel(filtered_items, OUTPUT_FILE)
        print(f"\n✨ Готово! В {OUTPUT_FILE} сохранено {total} топовых предложений.")
    else:
        print("\n⚠ Ничего не найдено.")




    

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False, args=["--disable-blink-features=AutomationControlled"])
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            locale="ru-RU", viewport={"width": 1366, "height": 900})
        await context.add_init_script("Object.defineProperty(navigator, 'webdriver', { get: () => undefined });")
        page = await context.new_page()

        # --- ЭТАП 1: СБОР ССЫЛОК СО СТРАНИЦ ПОИСКА ---
        for query in SEARCH_QUERIES:
            print(f"\n🔍 Поиск: «{query}»")
            empty_attempts = 0 # Счетчик пустых страниц
            
            for page_num in range(1, MAX_PAGES + 1):
                url = make_url(query, page_num)
                print(f"  Страница {page_num}... ", end="", flush=True)
                items = await parse_page(page, url, context)
                print(f"Найдено: {len(items)}")
                
                if len(items) == 0:
                    empty_attempts += 1
                    # Если подряд 2 пустые страницы — значит точно всё, выходим
                    if empty_attempts >= 2: 
                        print("  🛑 Пустые страницы, идем дальше.")
                        break 
                else:
                    all_items.extend(items)
                    empty_attempts = 0 # Сбрасываем счетчик, если нашли товар
                
                await asyncio.sleep(3) # Увеличим паузу, чтобы Авито меньше нас блокировал
        
        # --- ЭТАП 2: Анализ каждого объявления ---
        filtered_items = []
        if unique_items:
            print("📝 Анализ: Скрытые дефекты, Доставка, Комплект, Рейтинг...")
            for idx, item in enumerate(unique_items, 1):
                print(f"  [{idx}/{len(unique_items)}] {item['title'][:20]}... ", end="", flush=True)
                
                details = await fetch_description(page, item["url"])
                item.update(details)
                
                desc_lower = item["description"].lower()
                title_lower = item["title"].lower()
                
                # 🛑 Фильтр скрытых проблем (скам и блокировки)
                stop_issues = [
                    "mdm", "мдм", "soft unlock", "софт анлок", "операторский", "демо", "demo", "att", "verizon",
                    "вскрывался", "ремонтировался", "заменен", "менялся", "не работает отпечаток", "отвал", 
                    "мерцает", "заблокирован", "icloud", "гугл аккаунт", "google аккаунт", "пятн", "требуется ремонт"
                ]
                if any(bad in desc_lower for bad in stop_issues):
                    print("✕ (Отбраковано: Скрытый дефект/Блокировка)")
                    continue

                # 🛑 Фильтр доставки (Только Ростов)
                no_delivery = ["без доставки", "доставки нет", "авито доставки нет", "не отправляю", "только личная встреча", "не отправлю", "без пересыла"]
                if any(word in desc_lower for word in no_delivery):
                    if USER_CITY not in item["city"].lower():
                        print("✕ (Отбраковано: Нет доставки, другой город)")
                        continue

                # 🛑 Фильтр убитого состояния
                cond_lower = item["condition"].lower()
                scr_lower = item["screen"].lower()
                if cond_lower != "не указано" and not any(good in cond_lower for good in ["отлично", "хороше", "ново"]):
                    print("✕ (Убитое состояние)")
                    continue
                if scr_lower != "не указано" and not any(good in scr_lower for good in ["без дефект", "1-2", "1–2", "царапин"]):
                    print("✕ (Разбитый экран)")
                    continue

                # 🛑 Фильтр рейтинга (менее 4.0 - в мусор)
                if item["rating"] > 0 and item["rating"] < 4.0:
                    print("✕ (Плохой продавец, рейтинг ниже 4.0)")
                    continue

                # 🎁 Анализ комплекта
                if re.search(r'полный комплект|коробка|родная зарядка|ориг\w* блок', desc_lower):
                    item["kit"] = "✅ Полный"
                else:
                    item["kit"] = "❓ Уточнить"

                # 🏆 Расчет умной оценки (Скоринг)
                score = 0
                if item["condition"] == "Отличное" and item["screen"] == "Без дефектов": score += 3
                if item["kit"] == "✅ Полный": score += 2
                
                # ДЖЕКПОТЫ (+5 БАЛЛОВ) за самые топовые модели
                jackpots = [
                    r'(nothing|phone)\s*\(?(4\s*a|4\s*pro)\b',
                    r'(pixel|пиксель)\s*(9\s*pro|9pro)\b',
                    r'(edge 60|s60|edge 70)',
                    r'(oneplus|ванплас)\s+(12|12r|nord 4|nord 5|ace 3 pro)\b'
                ]
                if any(re.search(j, title_lower) for j in jackpots):
                    score += 5
                
                item["score"] = score
                print(f"✅ Оценка: {score} | АКБ: {item['battery']}")
                filtered_items.append(item)
                await asyncio.sleep(3)

        await browser.close()

    # Финальная сортировка и сохранение
    if filtered_items:
        filtered_items.sort(key=lambda x: (-x["score"], x["price"]))
        total = save_to_excel(filtered_items, OUTPUT_FILE)
        print(f"\n✨ Готово! В {OUTPUT_FILE} сохранено {total} топовых предложений.")
    else:
        print("\n⚠ Ничего не найдено.")

if __name__ == "__main__":
    asyncio.run(main())