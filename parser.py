"""Парсер каталога книг с выгрузкой в Excel.

Сайт: https://books.toscrape.com — учебная площадка, созданная специально для парсинга.

Запуск:
    python parser.py                 # все страницы каталога (50 стр., 1000 книг)
    python parser.py --pages 3       # только первые 3 страницы
    python parser.py --out books.xlsx
"""
import argparse
import time
from datetime import datetime
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

START_URL = "https://books.toscrape.com/catalogue/page-1.html"
HEADERS = {"User-Agent": "Mozilla/5.0 (portfolio demo parser)"}
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
DELAY = 0.5  # пауза между запросами, чтобы не нагружать сайт


def fetch(session: requests.Session, url: str, retries: int = 3) -> str:
    """Скачать страницу с повторными попытками при сбоях сети."""
    for attempt in range(1, retries + 1):
        try:
            resp = session.get(url, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            resp.encoding = "utf-8"  # иначе знак £ превращается в «Â£»
            return resp.text
        except requests.RequestException as e:
            if attempt == retries:
                raise
            print(f"  ошибка ({e}), повтор {attempt}/{retries - 1}...")
            time.sleep(2 * attempt)


def parse_page(html: str, page_url: str) -> tuple[list[dict], str | None]:
    """Вернуть книги со страницы и ссылку на следующую страницу (или None)."""
    soup = BeautifulSoup(html, "html.parser")
    books = []
    for card in soup.select("article.product_pod"):
        link = card.select_one("h3 a")
        price_text = card.select_one("p.price_color").get_text(strip=True)
        rating_cls = card.select_one("p.star-rating")["class"]
        rating_word = next((c for c in rating_cls if c in RATINGS), None)
        stock = card.select_one("p.availability").get_text(strip=True)

        books.append({
            "title": link["title"],
            "price": float(price_text.lstrip("£Â")),
            "rating": RATINGS.get(rating_word, 0),
            "in_stock": "Да" if "In stock" in stock else "Нет",
            "url": urljoin(page_url, link["href"]),
        })

    next_link = soup.select_one("li.next a")
    next_url = urljoin(page_url, next_link["href"]) if next_link else None
    return books, next_url


def scrape(max_pages: int | None) -> list[dict]:
    books, url, page = [], START_URL, 0
    with requests.Session() as session:
        while url and (max_pages is None or page < max_pages):
            page += 1
            print(f"Страница {page}: {url}")
            page_books, url = parse_page(fetch(session, url), url)
            books.extend(page_books)
            time.sleep(DELAY)
    return books


def save_excel(books: list[dict], path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Книги"

    columns = [("№", 6), ("Название", 60), ("Цена, £", 10),
               ("Рейтинг", 10), ("В наличии", 12), ("Ссылка", 50)]
    ws.append([name for name, _ in columns])
    for i, (_, width) in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    header_fill = PatternFill("solid", fgColor="305496")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for n, b in enumerate(books, start=1):
        ws.append([n, b["title"], b["price"], "★" * b["rating"], b["in_stock"], b["url"]])
        row = ws.max_row
        ws.cell(row, 3).number_format = "0.00"
        link_cell = ws.cell(row, 6)
        link_cell.hyperlink = b["url"]
        link_cell.font = Font(color="0563C1", underline="single")

    ws.freeze_panes = "A2"                 # шапка не уезжает при прокрутке
    ws.auto_filter.ref = ws.dimensions     # фильтры/сортировка по столбцам

    # Второй лист — короткая сводка
    summary = wb.create_sheet("Сводка")
    prices = [b["price"] for b in books] or [0]
    rows = [
        ("Дата выгрузки", datetime.now().strftime("%d.%m.%Y %H:%M")),
        ("Всего книг", len(books)),
        ("Средняя цена, £", round(sum(prices) / len(prices), 2)),
        ("Минимальная цена, £", min(prices)),
        ("Максимальная цена, £", max(prices)),
        ("С рейтингом 5★", sum(b["rating"] == 5 for b in books)),
    ]
    for r in rows:
        summary.append(r)
    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 20
    for cell in summary["A"]:
        cell.font = Font(bold=True)

    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description="Парсер books.toscrape.com → Excel")
    ap.add_argument("--pages", type=int, default=None, help="сколько страниц собрать (по умолчанию все)")
    ap.add_argument("--out", default="books.xlsx", help="имя выходного файла")
    args = ap.parse_args()

    books = scrape(args.pages)
    save_excel(books, args.out)
    print(f"Готово: {len(books)} книг сохранено в {args.out}")


if __name__ == "__main__":
    main()
