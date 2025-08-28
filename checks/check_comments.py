import os
import fitz  # PyMuPDF
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont


def _filename_code(filename: str) -> str:
    """Берём код из имени файла целиком (без расширения), даже если длина ≠ 40."""
    name, _ = os.path.splitext(os.path.basename(filename))
    return name[:40] if filename.startswith("PKS2") else filename


def check(file_path: str):
    """
    Проверка наличия встроенных комментариев (аннотаций) в PDF:
    - Стикеры (Text)
    - Выделения/подчёркивания (Highlight, Underline, StrikeOut)
    - Рисунки (Ink)
    Возвращает список кортежей: (код, CellRichText, 'автопроверка').
    """
    results = []
    author = "автопроверка"

    filename = os.path.basename(file_path)
    code = _filename_code(filename)

    try:
        doc = fitz.open(file_path)
    except Exception as e:
        msg = f"Ошибка чтения PDF: {e}"
        rich = CellRichText(TextBlock(InlineFont(color="FF0000", b=True), msg))
        return [(code, rich, author)]

    annotations = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        for annot in page.annots() or []:
            atype = annot.type[1]  # человекочитаемое имя (например "Highlight")
            contents = annot.info.get("content", "")

            # Стикеры
            if atype == "Text":
                annotations.append((page_num + 1, f"Комментарий: {contents}"))

            # Выделения текста
            elif atype in ("Highlight", "Underline", "StrikeOut"):
                annotations.append((page_num + 1, f"Выделение текста ({atype}) {contents}"))

            # Рисунки/карандаш
            elif atype == "Ink":
                annotations.append((page_num + 1, f"Рисунок/пометка (Ink) {contents}"))

            else:
                # другие типы, например линии, стрелки и т. п.
                annotations.append((page_num + 1, f"Аннотация {atype}: {contents}"))

    if not annotations:
        return []

    # ≤ 3 аннотаций — пишем все
    if len(annotations) <= 5:
        for page_num, text in annotations:
            msg = f"Страница {page_num}: {text}"
            rich = CellRichText(TextBlock(InlineFont(color="000000"), msg))
            results.append((code, rich, author))
    else:
        msg = f"Найдено {len(annotations)} встроенных комментариев в документе"
        rich = CellRichText(TextBlock(InlineFont(color="000000"), msg))
        results.append((code, rich, author))

    return results
