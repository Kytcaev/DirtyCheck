import os
import re
from PyPDF2 import PdfReader
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont


def _filename_code(filename: str) -> str:
    """Берём код из имени файла целиком (без расширения), даже если длина ≠ 40."""
    name, _ = os.path.splitext(os.path.basename(filename))
    return name[:40] if filename.startswith("PKS2") else filename


def _latin_cyr_pair(letter: str) -> str:
    """Возвращает класс символов для латинской/кириллической пары одной буквы."""
    mapping = {
        "А": "A", "В": "B", "С": "C", "Е": "E", "К": "K", "М": "M",
        "Н": "H", "О": "O", "Р": "P", "Т": "T", "Х": "X", "Л": "L",
        "а": "A", "в": "B", "с": "C", "е": "E", "к": "K", "м": "M",
        "н": "H", "о": "O", "р": "P", "т": "T", "х": "X", "л": "L"
    }
    return mapping.get(letter.upper(), re.escape(letter))


def _clean_code(raw: str) -> str:
    return (raw or "").replace("\n", "").strip()


def _extract_title_code_and_rev(reader: PdfReader):
    """
    Достаём код и ревизию с титульных (стр. 0–1).
    Возвращает (title_code, rev) или (None, None).
    """
    title_code, rev = None, None

    # Более точная регулярка: вытаскиваем код целиком (QC.0006.S и т.д.)
    code_re = re.compile(
        r"(PKS2\.[A-Z]\.[A-Z0-9.&]+?\.\d{3,4}\.[A-Z]{2}\.\d{3,4}\.[ESR])",
        re.I
    )
    rev_words_re = re.compile(r"(?i)\b(?:Revision|Rev(?:ision)?\.?|Revised[:]?)\s*([CBR]\d{2})")
    rev_inline_re = re.compile(r"/\s*([CBR]\d{2})\s*/")

    for p in (0, 1):
        try:
            text = reader.pages[p].extract_text() or ""
        except Exception:
            continue
        if not text:
            continue

        if not title_code:
            m_code = code_re.search(text)
            if m_code:
                title_code = _clean_code(m_code.group(1))

        if not rev:
            m_rev_w = rev_words_re.search(text)
            if m_rev_w:
                rev = m_rev_w.group(1).upper()

        if not rev:
            m_rev_i = rev_inline_re.search(text)
            if m_rev_i:
                rev = m_rev_i.group(1).upper()

        if title_code and rev:
            break

    if title_code and not rev and len(title_code) > 5:
        sixth = title_code[5].upper()
        letter = "B" if sixth == "L" else "C" if sixth == "D" else "R"
        rev = f"{letter}01"

    return title_code, rev


def _validate_code_structure(code: str):
    """
    Проверка структуры PKS2-кода:
    - Длина = 40
    - 9 секций
    - Соответствие stage и stage_code
    Возвращает None, если ошибок нет, иначе строку ошибки.
    """
    if len(code) != 40:
        return f"Код '{code}' имеет неверную длину ({len(code)} символов вместо 40)"

    parts = code.split(".")
    if len(parts) != 9:
        return f"Код '{code}' имеет неверное количество секций ({len(parts)} вместо 9)"

    stage = parts[1]
    stage_code = parts[6]

    # Справочники допустимых значений
    valid_stage_codes = {
        "D": {
            "MDD": {"CK", "JN", "EE", "KC", "SB", "ZG", "SE", "ME", "MB"},
            "MLA": {"CB", "PD", "PC", "YQ", "CZ", "YP", "HA", "QC", "PT", "HW", "EW"},
        },
        "L": {
            "MLA": {"CK", "ZR", "ZB"},
        },
    }

    if stage not in valid_stage_codes:
        return f"Код '{code}' содержит недопустимую стадию '{stage}'"

    allowed_codes = set()
    for group in valid_stage_codes[stage].values():
        allowed_codes.update(group)

    if stage_code not in allowed_codes:
        return f"Код '{code}' содержит недопустимый код '{stage_code}' для стадии '{stage}'"

    return None


def check(file_path: str):
    """
    Проверка:
    - Сравнение кода имени и титула
    - Структура кода (40 символов, 9 секций, stage vs stage_code)
    - Футеры с 3-й страницы
    Возвращает список кортежей: (код, CellRichText, 'автопроверка').
    """
    author = "автопроверка"
    results = []

    filename = os.path.basename(file_path)
    code_from_filename = _filename_code(filename)

    try:
        reader = PdfReader(file_path)
    except Exception as e:
        msg = f"Ошибка чтения PDF: {e}"
        results.append((code_from_filename, CellRichText(TextBlock(InlineFont(color="FF0000", b=True), msg)), author))
        return results

    title_code, rev = _extract_title_code_and_rev(reader)

    # 1) Несовпадение кода имени и титула
    if title_code and code_from_filename and code_from_filename != title_code:
        parts = [
            TextBlock(InlineFont(color="000000"), "Код в имени файла ("),
            TextBlock(InlineFont(color="FF0000", b=True), code_from_filename),
            TextBlock(InlineFont(color="000000"), ") не совпадает с кодом на титуле ("),
            TextBlock(InlineFont(color="FF0000", b=True), title_code),
            TextBlock(InlineFont(color="000000"), ")"),
        ]
        results.append((code_from_filename, CellRichText(*parts), author))

    # 2) Проверка структуры кода
    if title_code:
        err = _validate_code_structure(title_code)
        if err:
            results.append((title_code, CellRichText(TextBlock(InlineFont(color="000000"), err)), author))

    # Без титульного кода/ревизии дальше футер не проверяем
    if not title_code or not rev:
        return results

    total_pages = len(reader.pages)
    if total_pages <= 2:
        return results

    missing = []

    # 3) Проверка футеров с 3-й страницы
    for i in range(2, total_pages):
        try:
            page_text = reader.pages[i].extract_text() or ""
        except Exception:
            missing.append(i + 1)
            continue

        rev_rx = _latin_cyr_pair(rev[0]) + re.escape(rev[1:])
        pattern = re.compile(rf"{re.escape(title_code)}\s*/\s*{rev_rx}\s*/\s*{i + 1}")

        if not pattern.search(page_text):
            missing.append(i + 1)

    if not missing:
        return results

    if len(missing) <= 3:
        for p in missing:
            msg = f"Страница {p}: отсутствует футер (ожидалось '{title_code}/{rev}/{p}')"
            results.append((title_code, CellRichText(TextBlock(InlineFont(color="000000"), msg)), author))
    else:
        example = f"{title_code}/{rev}/N"
        msg = f"Футер отсутствует более чем на 3 страницах ({len(missing)} стр.). Пример: '{example}'"
        results.append((title_code, CellRichText(TextBlock(InlineFont(color="000000"), msg)), author))

    return results
