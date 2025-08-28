import os
import re
from PyPDF2 import PdfReader
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont


def clean_code(raw: str) -> str:
    """Очистка кода документа от пробелов/переводов строк по краям."""
    return raw.replace("\n", "").strip()


def extract_code(filename: str) -> str:
    """Фоллбек: извлекаем код PKS2 (первые 40 символов из имени файла)."""
    base = os.path.basename(filename)
    return base[:40] if base.startswith("PKS2") else base


def extract_code_and_revision(reader: PdfReader, filename: str):
    """
    Извлекает код и ревизию из имени файла или с титульных страниц (0–1).
    Если ревизия не найдена — вычисляет по 6-му символу кода.
    Возвращает (code[:40] | None, rev | None).
    """
    base_no_ext = os.path.splitext(os.path.basename(filename))[0]

    # Попытка 1: из имени файла: PKS2...._C01.pdf
    m = re.match(r"(PKS2\.[\w\.&]{10,})_([CBR]\d{2})", base_no_ext)
    if m:
        raw_code, rev = m.groups()
        return clean_code(raw_code)[:40], rev

    # Попытка 2: с титульных (страницы 0–1)
    for page_num in (0, 1):
        try:
            text = reader.pages[page_num].extract_text() or ""
        except Exception:
            continue

        code_match = re.search(r"(PKS2\.[\w\.&]{10,})", text)
        raw_code = clean_code(code_match.group(1)) if code_match else None

        # Ищем ревизию: "Revision C03", "Rev. B12", "REVISED: R07" и т.п.
        rev_match = re.search(
            r"(?i)\b(?:Revision|Rev(?:ision)?\.?|Revised[:]?)[\s:]*([CBR]\d{2})",
            text,
        )

        if raw_code:
            if rev_match:
                rev = rev_match.group(1).upper()
            else:
                # Фоллбек по 6-му символу кода (индекс 5)
                # L -> B, D -> C, иначе R
                try:
                    sixth = raw_code[5].upper()
                except IndexError:
                    sixth = ""
                rev_letter = "B" if sixth == "L" else "C" if sixth == "D" else "R"
                rev = f"{rev_letter}01"
            return raw_code[:40], rev

    return None, None


def highlight_word(word: str, cyrillic_pattern: re.Pattern):
    """Подсветка кириллических букв в слове (RichText для Excel)."""
    parts = []
    for ch in word:
        if cyrillic_pattern.match(ch):
            parts.append(TextBlock(InlineFont(color="FF0000", b=True), ch))
        else:
            parts.append(TextBlock(InlineFont(color="000000"), ch))
    return parts


def check(file_path: str):
    """
    Проверка PDF на кириллицу.
    - Если 40-й символ кода = 'E' → проверяются все страницы,
      иначе → только первые 4 страницы.
    Возвращает список кортежей: (code, CellRichText, 'автопроверка').
    """
    filename = os.path.basename(file_path)
    results = []

    try:
        reader = PdfReader(file_path)
    except Exception as e:
        # Если PDF не читается — возвращаем одно замечание
        return [(extract_code(filename), f"Ошибка чтения PDF: {e}", "автопроверка")]

    # Берём код/рев с титула/имени; если кода нет — фоллбек на имя
    code, rev = extract_code_and_revision(reader, filename)
    if not code:
        code = extract_code(filename)

    # Правило про 40-й символ 'E'
    check_all_pages = len(code) >= 40 and code[39].upper() == "E"
    num_pages_to_check = len(reader.pages) if check_all_pages else min(4, len(reader.pages))

    cyrillic_pattern = re.compile(r"[А-Яа-яЁё]")

    for page_idx in range(num_pages_to_check):
        try:
            text = reader.pages[page_idx].extract_text() or ""
        except Exception:
            continue

        if not cyrillic_pattern.search(text):
            continue

        words = text.split()
        cyrillic_words = [w for w in words if cyrillic_pattern.search(w)]

        # Если слишком много — одно общее замечание на страницу
        if len(cyrillic_words) > 10:
            comment_text = f"Страница {page_idx + 1}: на странице преобладают кириллические символы"
            rich = CellRichText(TextBlock(InlineFont(color="000000"), comment_text))
            results.append((code, rich, "автопроверка"))
            continue

        # Иначе — одно замечание на каждое найденное слово (с подсветкой букв)
        for word in cyrillic_words:
            parts = [
                TextBlock(InlineFont(color="000000"),
                          f"Страница {page_idx + 1}: найдена кириллица '")
            ]
            parts.extend(highlight_word(word, cyrillic_pattern))
            parts.append(TextBlock(InlineFont(color="000000"), "'"))
            rich = CellRichText(*parts)
            results.append((code, rich, "автопроверка"))

    return results
