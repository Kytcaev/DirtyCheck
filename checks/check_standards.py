# checks/check_standards.py
import os
import re
from typing import List, Dict, Set, Tuple
from datetime import datetime

import fitz  # PyMuPDF
import pandas as pd
from rapidfuzz import process, fuzz

# --- Настройки ---
FUZZY_THRESHOLD = 88  # порог % для "ПОХОЖЕ"
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
MASTER_FILE = os.path.join(BASE_DIR, "Standard.csv")

ADD_NON_APPLICABLE_COMMENT = True
NON_APPLICABLE_TEXT = "не применим в соответствии с PAKSII-PMM-02.1.01.02"

# Лог для Excel
_STANDARDS_LOG: List[Dict] = []

# Регекс-части
PREFIX_PART = r"(?:ISO|IEC|EN|MSZ|GOST|ГОСТ|GOST\s*R?|ГОСТ\s*Р?|OST|BS|DIN|ASME|ASTM|API|TU|СТО|STO|MU|МУ|MI|МИ|MR|МР|TP|TPR|NUREG|IAEA|WANO|HAEA|GKINP|OK|TI|VSN|SO|PCM|PPB(?:-AS)?)"

CANDIDATE_RE = re.compile(
    rf"""
    (?P<prefix_chain>{PREFIX_PART}(?:\s+{PREFIX_PART}){{0,4}})    # цепочка префиксов (ISO, MSZ EN ISO и т.д.)
    [\s,:-]* 
    (?:
        (?P<gkinp>-\d{{2}}-\d{{3}}-\d{{2}})                      # GKINP-02-033-82
        |
        (?P<code>\d{{1,5}}(?:[-/]\d+)*(?:\.\d+)*[A-Z0-9\-]*)      # 13480-2, 3.13, 60721-2-1, 5817, 10250-4 и т.п.
    )
    (?:\s*v(?P<ver>\d+))?                                        # v3
    (?:\s*[:/]\s*(?P<year>\d{{4}}))?                             # :2019 или /2019
    """,
    re.IGNORECASE | re.VERBOSE,
)

# ---------------- вспомогательные ----------------

def _dedupe_prefix_chain(chain: str) -> str:
    """Удаляем дубли в цепочке префиксов: 'MSZ EN MSZ EN ISO' -> 'MSZ EN ISO'; склеиваем 'GOST R' и 'ГОСТ Р'."""
    toks = [t for t in re.split(r"\s+", (chain or "").strip()) if t]
    toks = [t.upper() for t in toks]
    # склейка пар "GOST R" / "ГОСТ Р"
    merged = []
    i = 0
    while i < len(toks):
        if i + 1 < len(toks) and (
            (toks[i] == "GOST" and toks[i+1] == "R") or
            (toks[i] == "ГОСТ" and toks[i+1] == "Р")
        ):
            merged.append(toks[i] + " " + toks[i+1])
            i += 2
        else:
            merged.append(toks[i])
            i += 1
    # удаляем повторения с сохранением порядка
    out, seen = [], set()
    for t in merged:
        if t not in seen:
            out.append(t); seen.add(t)
    return " ".join(out)

def _clean_text_basic(raw: str) -> str:
    """Базовая чистка без склейки страниц."""
    if not raw:
        return ""
    text = raw.replace("\r", "\n")
    # NBSP
    text = text.replace("\xa0", " ")
    # склейка дефис-разрывов: "3834-\n1" -> "3834-1"
    text = re.sub(r"-\s*\n\s*", "-", text)
    # иногда номер разделён переносом без дефиса "ISO 3834\n1:2022" -> "ISO 3834-1:2022"
    text = re.sub(r"(\d)\s*\n\s*(\d{1,2}:?)", r"\1-\2", text)
    # нормализация пробелов
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text

def _fix_split_years(text: str) -> str:
    """
    Чиним кейс '...-2\\n3' -> '...-2:2023' (предполагаем, что вторая часть — последние 2 цифры года).
    Аккуратно: только когда перед этим стоит '-[1-9]' (типичный номер части стандарта).
    Не трогаем размеры типа 'EN 10060-45'.
    """
    def repl(m):
        prefix = m.group("prefix")
        main = m.group("main")
        part = m.group("part")  # одна цифра 1..9
        yy = m.group("yy")      # две цифры
        return f"{prefix}{main}-{part}:20{yy}"

    pattern = re.compile(
        rf"(?P<prefix>\b(?:{PREFIX_PART})\b[^\n]{{0,50}}?)"
        r"(?P<main>\d{3,5})-(?P<part>[1-9])\s*[-\n]\s*(?P<yy>\d{2})(?!\d)"
        , re.IGNORECASE
    )
    return pattern.sub(repl, text)

def _extract_pages(pdf_path: str) -> List[Tuple[int, str]]:
    """Возвращает список (page_number, cleaned_text)."""
    doc = fitz.open(pdf_path)
    out = []
    for i, page in enumerate(doc, start=1):
        try:
            txt = page.get_text("text") or ""
            txt = _clean_text_basic(txt)
            txt = _fix_split_years(txt)
        except Exception:
            txt = ""
        out.append((i, txt))
    doc.close()
    return out

def _last_base(prefix_chain: str) -> str:
    toks = [t.strip().upper() for t in re.split(r"\s+", (prefix_chain or "").strip()) if t.strip()]
    if not toks:
        return ""
    # собрать составные типа "GOST R"
    merged = []
    i = 0
    while i < len(toks):
        if i+1 < len(toks) and ((toks[i] == "GOST" and toks[i+1] == "R") or (toks[i] == "ГОСТ" and toks[i+1] == "Р")):
            merged.append(toks[i] + " " + toks[i+1])
            i += 2
        else:
            merged.append(toks[i])
            i += 1
    return merged[-1]

def _canon(code: str, base: str, year: str | None) -> str:
    code = (code or "").strip().upper()
    code = re.sub(r"[^\w\-\./]", "", code)
    base = (base or "").strip().upper()
    if not base or not code:
        return ""
    key = f"{base} {code}"
    if year:
        key = f"{key}:{year}"
    return key

def _parse_to_canon(prefix_chain: str, code: str, year: str | None) -> Tuple[str, str]:
    """Строим canon_key и красивое отображение с дедупом префиксов."""
    base = _last_base(prefix_chain)
    key = _canon(code, base, year)

    chain_disp = _dedupe_prefix_chain(prefix_chain)
    disp = (f"{chain_disp} {code}".strip()).upper()
    if year and f":{year}" not in disp:
        disp += f":{year}"
    return key, disp

def _detect_encoding_and_read_lines(path: str) -> List[str]:
    encodings = ["utf-8-sig", "utf-8", "cp1251", "latin1"]
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, errors="ignore") as f:
                lines = [ln.strip() for ln in f if ln.strip()]
            if lines:
                return lines
        except Exception:
            continue
    with open(path, "rb") as f:
        lines = [ln.decode("latin1", errors="ignore").strip() for ln in f if ln.strip()]
        return [ln for ln in lines if ln]

def _load_master_list(path: str) -> Dict[str, str]:
    """Возвращает словарь canon_key -> original_display_from_master."""
    lines = []
    # пробуем pandas
    for enc in ("utf-8-sig", "utf-8", "cp1251", "latin1"):
        try:
            df = pd.read_csv(path, encoding=enc, engine="python", header=None, on_bad_lines="skip")
            if df.shape[0] > 0:
                col0 = df.iloc[:, 0].astype(str).str.strip().replace({"nan": ""}).tolist()
                lines = [x for x in col0 if x]
                if lines:
                    break
        except Exception:
            continue
    if not lines:
        lines = _detect_encoding_and_read_lines(path)

    master_map: Dict[str, str] = {}
    for s in lines:
        # пробуем через наш парсер
        m = CANDIDATE_RE.search(s)
        if m:
            prefix_chain = m.group("prefix_chain") or ""
            code = m.group("gkinp") or m.group("code") or ""
            year = m.group("year")
            key, disp = _parse_to_canon(prefix_chain, code, year)
            if key:
                master_map[key] = disp or s.strip()
                continue

        # fallback: грубо
        mnum = re.search(r"(\d{1,5}(?:[-/]\d+)*)", s)
        pp = re.search(PREFIX_PART, s, flags=re.IGNORECASE)
        yy = re.search(r"(\d{4})", s)
        if pp and mnum:
            k = _canon(mnum.group(1), pp.group(0), yy.group(1) if yy else None)
            if k:
                master_map[k] = s.strip()
    return master_map

def _has_cyrillic(s: str) -> bool:
    return bool(re.search(r"[А-Яа-яЁё]", s))

def _split_base_year(key: str) -> Tuple[str, str]:
    if not key:
        return "", ""
    m = re.match(r"^(.+?)(?::(\d{4}))?$", key)
    if not m:
        return key, ""
    return m.group(1), (m.group(2) or "")

def _looks_like_material(fragment: str) -> bool:
    """Отсечём типичные марки материалов: '1.4571', '1.4057' и т.п."""
    return bool(re.search(r"\b1\.\d{3,4}\b", fragment))

# ---------------- основная логика ----------------

def check_pdf(pdf_path: str) -> List[Tuple]:
    print(f"\nПроверка стандартов в {pdf_path}...")

    # master
    try:
        master_map = _load_master_list(MASTER_FILE)
    except Exception as e:
        print(f"  [Ошибка загрузки master: {e}]")
        master_map = {}

    master_keys: Set[str] = set(master_map.keys())
    master_base_map: Dict[str, List[str]] = {}
    for k in master_keys:
        b, _ = _split_base_year(k)
        master_base_map.setdefault(b, []).append(k)

    # двуязычный?
    base_filename = os.path.splitext(os.path.basename(pdf_path))[0]
    ignore_cyrillic_doc = bool(re.search(r"\.S($|_)", base_filename, flags=re.IGNORECASE))

    # постранично
    pages = _extract_pages(pdf_path)

    # seen для защиты от дублей: (page, canon_key, disp_norm)
    seen: Set[Tuple[int, str, str]] = set()

    for page_no, text in pages:
        # собираем кандидаты на странице
        candidates: Dict[str, Set[str]] = {}

        for m in CANDIDATE_RE.finditer(text):
            frag = m.group(0) or ""
            if _looks_like_material(frag):
                continue  # отсечь EN10272-1.4057 и т.п.

            prefix_chain = m.group("prefix_chain") or ""
            code = m.group("gkinp") or m.group("code") or ""
            year = m.group("year")

            key, disp = _parse_to_canon(prefix_chain, code, year)
            if not key:
                continue

            candidates.setdefault(key, set()).add(disp)

        if not candidates:
            continue

        # фильтрация кириллицы (если есть латиница или документ двуязычный)
        for key in list(candidates.keys()):
            disps = candidates[key]
            has_latin = any(not _has_cyrillic(d) for d in disps)
            if has_latin:
                candidates[key] = set(d for d in disps if not _has_cyrillic(d))
            elif ignore_cyrillic_doc:
                candidates[key] = set(d for d in disps if not _has_cyrillic(d))

        # вывод
        for key in sorted(candidates.keys()):
            disps = candidates[key]
            if not disps:
                continue
            disp = sorted(disps, key=len, reverse=True)[0]

            dedup_token = (page_no, key, re.sub(r"\s+", "", disp.upper()))
            if dedup_token in seen:
                continue
            seen.add(dedup_token)

            status = ""
            comment = ""

            if key in master_keys:
                status = "OK"
                print(f"  [стр. {page_no}] OK: {disp}")
            else:
                base, cand_year = _split_base_year(key)

                if base in master_base_map:
                    variants = sorted(master_base_map[base])
                    # есть ли точное совпадение по году?
                    same_year = any(_split_base_year(v)[1] == cand_year and cand_year for v in variants)
                    if same_year:
                        status = "OK"
                        print(f"  [стр. {page_no}] OK: {disp}")
                    else:
                        status = "КОД ИЗВЕСТЕН"
                        comment = f"варианты в master: {', '.join(variants)}"
                        print(f"  [стр. {page_no}] КОД ИЗВЕСТЕН (варианты в master): {disp} → {', '.join(variants)}")
                else:
                    # fuzzy к полному ключу
                    best = process.extractOne(key, master_keys, scorer=fuzz.token_set_ratio) if master_keys else None
                    if best and best[1] >= FUZZY_THRESHOLD:
                        best_key = best[0]
                        score = best[1]
                        bk_base, bk_year = _split_base_year(best_key)
                        master_disp = master_map.get(best_key, best_key)
                        if bk_year and cand_year and bk_year != cand_year:
                            status = "ПОХОЖЕ"
                            comment = f"{master_disp} ({score}%) (год отличается: {cand_year} vs {bk_year})"
                            print(f"  [стр. {page_no}] ПОХОЖЕ: {disp} → {master_disp} ({score}%) (год отличается: {cand_year} vs {bk_year})")
                        else:
                            status = "ПОХОЖЕ"
                            comment = f"{master_disp} ({score}%)"
                            print(f"  [стр. {page_no}] ПОХОЖЕ: {disp} → {master_disp} ({score}%)")
                    else:
                        status = "ВНЕ MASTER"
                        comment = NON_APPLICABLE_TEXT if ADD_NON_APPLICABLE_COMMENT else "Не найдено в master"
                        print(f"  [стр. {page_no}] ВНЕ MASTER: {disp}")

            # логируем всё подряд
            _STANDARDS_LOG.append({
                "Document": os.path.basename(pdf_path),
                "Page": page_no,
                "Canon key": key,
                "Found": disp,
                "Status": status or "—",
                "Comment": comment or ""
            })

    return []  # совместимость с main

def check(file_path: str):
    return check_pdf(file_path)

def finalize(save_dir: str | None = None):
    """Сохраняем подробный отчёт по стандартам в отдельный Excel."""
    if not _STANDARDS_LOG:
        return
    df = pd.DataFrame(_STANDARDS_LOG)
    today = datetime.now().strftime("%Y%m%d")
    out_name = f"Standards_check_{today}.xlsx"
    out_path = os.path.join(save_dir or os.getcwd(), out_name)
    df.to_excel(out_path, index=False)
    print(f"\nПодробный отчёт по стандартам сохранён в {out_path}")
