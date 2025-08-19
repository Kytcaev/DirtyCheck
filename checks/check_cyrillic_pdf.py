import os
import re
from PyPDF2 import PdfReader
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont


def extract_code(filename):
    """Извлекаем код PKS2 (40 символов)"""
    return filename[:40] if filename.startswith("PKS2") else filename


def highlight_word(word, cyrillic_pattern):
    """Подсветка кириллических букв в слове"""
    parts = []
    for ch in word:
        if cyrillic_pattern.match(ch):
            parts.append(TextBlock(InlineFont(color="FF0000", b=True), ch))
        else:
            parts.append(TextBlock(InlineFont(color="000000"), ch))
    return parts


def check(file_path):
    """
    Проверка первых 4 страниц PDF на кириллицу.
    Возвращает список (код, comment_rich, author)
    """
    filename = os.path.basename(file_path)
    code = extract_code(filename)
    results = []

    try:
        reader = PdfReader(file_path)
    except Exception as e:
        return [(code, f"Ошибка чтения PDF: {e}", "автопроверка")]

    num_pages = min(4, len(reader.pages))
    cyrillic_pattern = re.compile(r"[А-Яа-яЁё]")

    for page_num in range(num_pages):
        text = reader.pages[page_num].extract_text() or ""
        matches = cyrillic_pattern.findall(text)

        if not matches:
            continue

        words = text.split()
        cyrillic_words = [w for w in words if cyrillic_pattern.search(w)]

        # Если слишком много кириллицы
        if len(cyrillic_words) > 10:
            comment_text = f"Страница {page_num+1}: на странице преобладают кириллические символы"
            rich = CellRichText(TextBlock(InlineFont(color="000000"), comment_text))
            results.append((code, rich, "автопроверка"))
        else:
            for word in cyrillic_words:
                parts = [TextBlock(InlineFont(color="000000"), f"Страница {page_num+1}: найдена кириллица '")]
                parts.extend(highlight_word(word, cyrillic_pattern))
                parts.append(TextBlock(InlineFont(color="000000"), "'"))
                rich = CellRichText(*parts)
                results.append((code, rich, "автопроверка"))

    return results
