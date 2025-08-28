import os
import re
import glob
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont


# Пути
#IMS_PATH = r"Z:\7. DLSC\99_Конфигурация\DATA IMS 2.0"
IMS_PATH = r"C:\work"
LOCAL_BACKUP = "2025.08.25_DSR.xlsx"

# Статусы, при которых считаем документ "активным"
ALLOWED_STATUSES = {
    "Approved with remarks",
    "Pending",
    "Approved",
    "Handover of signed copies",
    "Final approval",
    "Electronic signs",
    "Submission for Consent",
    "Meeting Request",
    "DPTRA response",
    "DPTRA remarks",
    "Submission for ES",
    "Submission for electronic signs",
}


def _find_latest_ims_file():
    """Находит последний файл *_DSR.xlsx в IMS_PATH или локальный бэкап."""
    try:
        files = glob.glob(os.path.join(IMS_PATH, "*_DSR.xlsx"))
        if files:
            return max(files, key=os.path.getmtime)
    except Exception:
        pass
    return LOCAL_BACKUP


def _load_ims_codes():
    """Чтение IMS-кодов из последнего DSR.xlsx. Возвращает словарь: code -> (rev_prefix, status)."""
    file = _find_latest_ims_file()
    df = pd.read_excel(file)

    ims = {}
    for _, row in df.iterrows():
        code = str(row.get("IMS_CODE", "")).strip()
        rev = str(row.get("Business revision", ""))[:3]  # только первые 3 символа
        status = str(row.get("Dal Action", "")).strip()
        if code:
            ims[code] = (rev, status)
    return ims


def _clean_code(raw: str) -> str:
    return (raw or "").replace("\n", "").strip()


def _extract_title_code_and_rev(reader: PdfReader):
    """
    Достаём код и ревизию с титульных (стр. 0–1).
    Возвращает (title_code, rev) или (None, None).
    """
    title_code, rev = None, None

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

    # Фоллбек ревизии по 6-му символу кода (если код найден, а рев нет)
    if title_code and not rev and len(title_code) > 5:
        sixth = title_code[5].upper()
        letter = "B" if sixth == "L" else "C" if sixth == "D" else "R"
        rev = f"{letter}01"

    return title_code, rev


def check(file_path: str):
    """
    Проверка на дублирование кода в IMS.
    Возвращает список кортежей: (код, CellRichText, 'автопроверка для АСЭ').
    """
    results = []
    author = "автопроверка для АСЭ"

    filename = os.path.basename(file_path)

    try:
        reader = PdfReader(file_path)
        title_code, rev = _extract_title_code_and_rev(reader)
    except Exception as e:
        msg = f"Ошибка чтения PDF: {e}"
        rich = CellRichText(TextBlock(InlineFont(color="000000"), msg))
        return [(filename, rich, author)]

    if not title_code or not rev:
        return results

    ims_file = _find_latest_ims_file()
    has_network_access = ims_file != LOCAL_BACKUP

    ims = _load_ims_codes()
    rev_prefix = rev[:3]

    if title_code in ims:
        ims_rev, ims_status = ims[title_code]

        if not has_network_access:
            # Локальный режим (без статусов)
            if ims_rev != rev_prefix:
                msg = f"Код документа уже существует в IMS с ревизией {ims_rev}, текущая ревизия {rev_prefix}"
            else:
                msg = f"Код документа уже зарегистрирован в IMS ({ims_rev})"
            rich = CellRichText(TextBlock(InlineFont(color="000000"), msg))
            results.append((title_code, rich, author))

        else:
            # Расширенный режим (с проверкой статуса)
            if rev_prefix > ims_rev:
                if ims_status in ALLOWED_STATUSES:
                    msg = f"Код документа уже зарегистрирован и статус документа \"{ims_status}\""
                    rich = CellRichText(TextBlock(InlineFont(color="000000"), msg))
                    results.append((title_code, rich, author))
            elif rev_prefix == ims_rev:
                # Совпадает — пропускаем
                pass
            else:  # рев ниже
                msg = f"Код документа зарегистрирован с более новой ревизией {ims_rev} (текущая {rev_prefix})"
                rich = CellRichText(TextBlock(InlineFont(color="000000"), msg))
                results.append((title_code, rich, author))

    return results
