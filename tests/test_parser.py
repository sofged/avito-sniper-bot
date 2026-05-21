import os
import pytest
from unittest.mock import AsyncMock, MagicMock, patch
import avito_parser

def test_make_url():
    """Проверяет правильность генерации поискового URL."""
    url_p1 = avito_parser.make_url("google pixel", page=1)
    assert "q=google+pixel" in url_p1
    assert "pmin=" in url_p1
    assert "pmax=" in url_p1
    assert "p=" not in url_p1

    url_p2 = avito_parser.make_url("google pixel", page=2)
    assert "p=2" in url_p2


def test_cache_load_save(tmp_path):
    """Тестирует чтение и запись кэша без изменения реального файла."""
    temp_cache_file = tmp_path / "test_cache.json"
    
    with patch("avito_parser.CACHE_FILE", str(temp_cache_file)):
        assert avito_parser.load_cache() == {}

        test_data = {"https://avito.ru/item1": {"title": "Test Phone", "price": 9000}}
        avito_parser.save_cache(test_data)

        assert temp_cache_file.exists()

        loaded = avito_parser.load_cache()
        assert loaded["https://avito.ru/item1"]["title"] == "Test Phone"


@pytest.mark.asyncio
async def test_is_captcha_page():
    """Тестирует детекцию капчи по заголовку страницы."""
    mock_page = AsyncMock()
    
    mock_page.title.return_value = "Доступ временно ограничен"
    assert await avito_parser.is_captcha_page(mock_page) is True

    mock_page.title.return_value = "Купить Nothing Phone (2) на Авито"
    assert await avito_parser.is_captcha_page(mock_page) is False


@pytest.mark.asyncio
async def test_parse_page_filtering():
    """
    Тестирует фильтрацию объявлений на странице поиска.
    Проверяет, что телефоны проходят фильтр, а аксессуары (стоп-слова) отсекаются.
    """
    mock_page = AsyncMock()
    mock_page.title.return_value = "Купить смартфоны на Авито"
    mock_context = MagicMock()

    mock_card_1 = AsyncMock()
    
    title_el_1 = AsyncMock()
    title_el_1.inner_text.return_value = "Смартфон Google Pixel 8 128Gb"
    
    price_el_1 = AsyncMock()
    price_el_1.get_attribute.return_value = "12000" 
    
    link_el_1 = AsyncMock()
    link_el_1.get_attribute.return_value = "/rossiya/telefony/pixel_8"
    
    geo_el_1 = AsyncMock()
    geo_el_1.inner_text.return_value = "Ростов-на-Дону"

    def side_effect_card_1(selector):
        return {
            '[itemprop="name"]': title_el_1,
            '[itemprop="price"]': price_el_1,
            '[data-marker="item-title"]': link_el_1,
            '[data-marker="item-address"]': geo_el_1
        }.get(selector)
    mock_card_1.query_selector.side_effect = side_effect_card_1

    mock_card_2 = AsyncMock()
    
    title_el_2 = AsyncMock()
    title_el_2.inner_text.return_value = "Чехол силиконовый для Pixel 8"
    
    price_el_2 = AsyncMock()
    price_el_2.get_attribute.return_value = "500"
    
    link_el_2 = AsyncMock()
    link_el_2.get_attribute.return_value = "/rossiya/aksessuary/case_pixel"
    
    geo_el_2 = AsyncMock()
    geo_el_2.inner_text.return_value = "Ростов-на-Дону"

    def side_effect_card_2(selector):
        return {
            '[itemprop="name"]': title_el_2,
            '[itemprop="price"]': price_el_2,
            '[data-marker="item-title"]': link_el_2,
            '[data-marker="item-address"]': geo_el_2
        }.get(selector)
    mock_card_2.query_selector.side_effect = side_effect_card_2

    mock_page.query_selector_all.return_value = [mock_card_1, mock_card_2]

    results = await avito_parser.parse_page(mock_page, "https://dummy-url.com", mock_context)

    assert len(results) == 1
    assert results[0]["title"] == "Смартфон Google Pixel 8 128Gb"
    assert results[0]["price"] == 12000
    assert "pixel_8" in results[0]["url"]