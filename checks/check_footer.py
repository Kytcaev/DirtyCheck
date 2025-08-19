import os
from PyPDF2 import PdfReader


def extract_code(filename: str) -> str:
    """Извлекаем код PKS2 (40 символов)"""
    return filename[:40] if filename.startswith("PKS2") else filename


def check(file_path: str):
    """
    Проверка наличия футера на страницах PDF (начиная с 3-й).
    Футер должен содержать: код_документа/ревизия/номер_страницы
    """
    results = []
    code = extract_code(os.path.basename(file_path))

    try:
        reader = PdfReader(file_path)
    except Exception:
        return []

    total_pages = len(reader.pages)

    for page_num in range(2, total_pages):  # начинаем с 3-й страницы (индекс 2)
        try:
            text = reader.pages[page_num].extract_text() or ""
        except Exception:
            continue

        # Проверяем, что есть шаблон вида "PKS2xxx/Rev/стр"
        if code not in text and "/" not in text:
            results.append({
                "Код": code,
                "Комментарий": f"Стр.{page_num+1}: отсутствует футер",
                "Автор": "автопроверка"
            })

    return results
