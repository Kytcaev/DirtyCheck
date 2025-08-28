import os
import re
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont


def extract_code(filename):
    """Берём код из имени файла целиком (без расширения), даже если длина ≠ 40."""
    name, _ = os.path.splitext(os.path.basename(filename))
    return name [:40] if filename.startswith("PKS2") else filename

def check(file_path):
    """
    Проверка наличия кириллицы в имени файла.
    Возвращает список:
    [
      (код документа, CellRichText, 'автопроверка')
    ]
    """
    filename = os.path.basename(file_path)
    code = extract_code(filename)

    cyrillic_pattern = re.compile(r'[А-Яа-яЁё]')
    matches = list(cyrillic_pattern.finditer(filename))
    if not matches:
        return []

    # Формируем RichText с подсветкой только букв
    normal = TextBlock(InlineFont(color="000000"), "Имя файла содержит кириллицу: ")
    parts = [normal]

    for char in filename:
        if cyrillic_pattern.match(char):
            # подсвечиваем красным жирным
            parts.append(TextBlock(InlineFont(color="FF0000", b=True), char))
        else:
            parts.append(TextBlock(InlineFont(color="000000"), char))

    comment_rich = CellRichText(*parts)

    return [(code, comment_rich, 'автопроверка')]
