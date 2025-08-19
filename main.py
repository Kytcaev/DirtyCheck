import os
import importlib
import datetime
from openpyxl import load_workbook
from tkinter import Tk, filedialog

CHECK_MODULES = [
    "checks.check_cyrillic_filename",
    "checks.check_cyrillic_pdf",
]

TEMPLATE_FILE = "Summary of comments.xlsx"
OUTPUT_FILE = f"Summary_of_comments_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"


def find_first_empty_row(ws):
    """Ищем первую пустую строку в колонке B (код документа)"""
    row = 2  # начинаем после заголовков
    while ws.cell(row=row, column=2).value not in (None, ""):
        row += 1
    return row


def main():
    # выбираем папку
    Tk().withdraw()
    folder = filedialog.askdirectory(title="Выберите папку с PDF")
    if not folder:
        print("Папка не выбрана.")
        return

    # загружаем шаблон
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    # загружаем проверки
    checks = [importlib.import_module(m) for m in CHECK_MODULES]

    # собираем pdf
    files = []
    for root, _, filenames in os.walk(folder):
        for f in filenames:
            if f.startswith("PKS2") and f.lower().endswith(".pdf"):
                files.append(os.path.join(root, f))

    print(f"Найдено {len(files)} файлов для проверки.")

    # нумерация
    next_row = find_first_empty_row(ws)
    index = 1

    for file_path in files:
        for check_module in checks:
            results = check_module.check(file_path)
            for res in results:
                # разные модули возвращают разный формат
                if len(res) == 3:  # (код, comment_rich, author)
                    code, comment, author = res
                    page = None
                else:              # (код, page, comment_rich, author)
                    code, page, comment, author = res

                ws.cell(row=next_row, column=1).value = index
                ws.cell(row=next_row, column=2).value = code
                ws.cell(row=next_row, column=3).value = comment
                ws.cell(row=next_row, column=4).value = author
                if page:
                    ws.cell(row=next_row, column=5).value = page

                next_row += 1
                index += 1

    wb.save(folder+"\\"+OUTPUT_FILE)
    print(f"Готово! Результаты сохранены в {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
input("\nГотово. Нажмите Enter, чтобы закрыть окно...")
