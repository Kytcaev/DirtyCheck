import os
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# --- правила ---
MDD_REQUIRED = {"CK", "JN", "EE", "KC", "SB", "ZG", "SE"}  # ME или MB — хотя бы один
MLA_REQUIRED = {"CB", "PD", "PC", "YQ", "CZ", "YP", "HA", "QC",
                "PT", "HW", "EW", "CK", "ZR", "ZB"}
MDD_OPTIONAL = {"ME", "MB"}  # хотя бы один из них

# --- расшифровка кодов ---
DESCRIPTIONS = {
    "MB": "Техническое задание (ToR)",
    "ME": "Технические условия (TU)",
    "SE": "Сборочные чертежи, спецификации, расчёты на прочность",
    "ZG": "Паспорт на оборудование (форма)",
    "SB": "Ведомость замены материалов",
    "KC": "Руководство по эксплуатации",
    "EE": "Инструкция по транспортировке, хранению, консервации",
    "JN": "Перечень запасных частей (ЗИП)",
    "CZ": "Описание процесса изготовления",
    "YQ": "Таблицы контроля качества ТБ1, ТБ2",
    "PC": "Программа и методика испытаний",
    "PD": "Программа контроля качества",
    "CB": "Принципиальная технология производства",
    "HW": "Welding Procedure Qualification Record (WPQR)",
    "EW": "Welding Procedure Specification (WPS)",
    "PT": "Сборники технологических карт контроля",
    "YP": "Процедуры и инструкции",
    "HA": "Перечень форматов отчётных документов",
    "QC": "План качества",
    "CK": "MLA сопроводительный документ",
    "ZR": "Заявления проектировщика",
    "ZB": "Сопроводительная документация"
}


def clean_code(filename: str) -> str:
    """Вернуть PKS2-код без расширений и хвостов типа _C01_ASE"""
    name, _ = os.path.splitext(os.path.basename(filename))
    return name.split("_")[0]  # берём только основной код до "_"


def parse_code(filename: str):
    """Вытащить 7-й сектор и всю структуру"""
    code = clean_code(filename)
    parts = code.split(".")
    if len(parts) < 7:
        return None, None, code, parts
    stage = parts[1].upper()
    sec7 = parts[6].upper()
    return stage, sec7, code, parts


def build_ck_code(parts, stage="D"):
    """Собрать примерный CK-код (если отсутствует)"""
    parts = parts.copy()
    parts[1] = stage  # заменяем стадию
    parts[6] = "CK"
    parts[7] = "0001"
    parts[8] = "E"
    return ".".join(parts[:9])


def check(folder: str, files: list[str]):
    results = []
    author = "автопроверка"

    stage_docs = {"D": set(), "L": set()}
    qc_present = False
    ck_D_code = None
    ck_L_code = None
    any_parts = None

    # Сканируем файлы
    for file in files:
        stage, sec7, code, parts = parse_code(file)
        if not sec7:
            continue
        if stage in stage_docs:
            stage_docs[stage].add(sec7)
        if sec7 == "QC":
            qc_present = True
        if sec7 == "CK" and stage == "D":
            ck_D_code = code
        if sec7 == "CK" and stage == "L":
            ck_L_code = code
        if not any_parts:  # для генерации "примерного кода"
            any_parts = parts

    if not any_parts:
        return results

    # --- Правила CK ---
    if ck_L_code and not ck_D_code:
        ck_D_code = build_ck_code(any_parts, "D")
        results.append((
            ck_L_code,
            CellRichText(TextBlock(InlineFont(b=True, color="FF0000"),
                                   f"Обнаружен L-CK без D-CK — должен быть {ck_D_code}")),
            author
        ))

    if not (ck_D_code or ck_L_code):
        return results  # Нет CK — проверку состава не делаем

    # --- Если есть QC → должен быть L-CK ---
    if qc_present and not ck_L_code:
        ck_L_code = build_ck_code(any_parts, "L")
        results.append((
            ck_L_code,
            CellRichText(TextBlock(InlineFont(),
                                   f"Отсутствует документ {ck_L_code}")),
            author
        ))

    # --- Проверка состава MDD ---
    if ck_D_code:
        for doc in MDD_REQUIRED:
            if doc not in stage_docs["D"]:
                desc = DESCRIPTIONS.get(doc, "")
                results.append((
                    ck_D_code,
                    CellRichText(TextBlock(InlineFont(),
                                           f"Отсутствует документ '{doc}' ({desc})")),
                    author
                ))
        if not (stage_docs["D"] & MDD_OPTIONAL):
            results.append((
                ck_D_code,
                CellRichText(TextBlock(InlineFont(),
                                       "Отсутствует документ 'ME' или 'MB' (должен быть хотя бы один)")),
                author
            ))

    # --- Проверка состава MLA ---
    if ck_L_code:
        for doc in MLA_REQUIRED:
            if doc == "CK" and ck_D_code:
                continue  # CK не дублируем, если есть в MDD
            if doc not in stage_docs["L"]:
                desc = DESCRIPTIONS.get(doc, "")
                results.append((
                    ck_L_code,
                    CellRichText(TextBlock(InlineFont(),
                                           f"Отсутствует документ '{doc}' ({desc})")),
                    author
                ))

    return results
