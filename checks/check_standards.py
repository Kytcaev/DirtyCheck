# checks/check_standards.py
import os
import re
from typing import List, Dict, Set, Tuple

import fitz  # PyMuPDF
import pandas as pd
from rapidfuzz import process, fuzz

# --- Настройки ---
FUZZY_THRESHOLD = 88  # порог % для "ПОХОЖЕ"
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
MASTER_FILE = os.path.join(BASE_DIR, "Standard.csv")

# Регекс-части
PREFIX_PART = r"(?:ISO|IEC|EN|MSZ|GOST|ГОСТ|GOST\s*R?|ГОСТ\s*Р?|OST|BS|DIN|ASME|ASTM|API|TU|СТО|STO|MU|МУ|MI|МИ|MR|МР|TP|TPR|NUREG|IAEA|WANO|HAEA|GKINP|OK|TI|VSN|SO|PCM|PPB(?:-AS)?)"

CANDIDATE_RE = re.compile(
    rf"""
    (?P<prefix_chain>{PREFIX_PART}(?:\s+{PREFIX_PART}){{0,4}})    # цепочка префиксов (ISO, MSZ EN ISO и т.д.)
    [\s,:-]* 
    (?:
        (?P<gkinp>-\d{{2}}-\d{{3}}-\d{{2}})                      # GKINP-02-033-82
        |
        (?P<code>\d{{1,5}}(?:[-/]\d+)*(?:\.\d+)*[A-Z0-9\-]*)      # номера вроде 13480-2, 3.13, 60721-2-1, 5817
    )
    (?:\s*v(?P<ver>\d+))?                                        # v3
    (?:\s*[:/]\s*(?P<year>\d{{4}}))?                             # :2019 или /2019
    """,
    re.IGNORECASE | re.VERBOSE,
)

# ---------------- вспомогательные ----------------

def _clean_text(raw: str) -> str:
    """Склеиваем разрыв по дефису, убираем NBSP, нормализуем пробелы и переводы строк."""
    if not raw:
        return ""
    text = raw.replace("\r", "\n")
    # склейка разрывов по дефису: "3834-\n1" -> "3834-1"
    text = re.sub(r"-\s*\n\s*", "-", text)
    # иногда номер разделён переносом без дефиса "ISO 3834\n1:2022" -> "ISO 3834-1:2022" (хотя осторожно)
    text = re.sub(r"(\d)\s*\n\s*(\d{1,2}:?)", r"\1-\2", text)
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text

def _extract_text(pdf_path: str) -> str:
    doc = fitz.open(pdf_path)
    parts = []
    for page in doc:
        try:
            parts.append(page.get_text("text") or "")
        except Exception:
            parts.append("")
    doc.close()
    return _clean_text("\n".join(parts))

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

def _parse_to_canon(s: str) -> Tuple[str, str]:
    """Попытаться распарсить любую строку в (canon_key, nice_display)."""
    m = CANDIDATE_RE.search(s)
    if not m:
        return "", ""
    prefix_chain = m.group("prefix_chain") or ""
    code = m.group("gkinp") or m.group("code") or ""
    year = m.group("year")
    ver = m.group("ver")
    base = _last_base(prefix_chain)
    disp = (prefix_chain + " " + (m.group(0) or "")).strip()
    if ver:
        disp = disp + f" v{ver}"
    if year:
        # если год уже в disp — ок; иначе добавим
        if f":{year}" not in disp:
            disp = disp + f":{year}"
    key = _canon(code, base, year)
    return key, disp.strip()

def _detect_encoding_and_read_lines(path: str) -> List[str]:
    """Читаем файл построчно, пробуем разные кодировки; возвращаем непустые строки."""
    encodings = ["utf-8-sig", "utf-8", "cp1251", "latin1"]
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, errors="ignore") as f:
                lines = [ln.strip() for ln in f if ln.strip()]
            if lines:
                return lines
        except Exception:
            continue
    # последний фолбек
    with open(path, "rb") as f:
        lines = [ln.decode("latin1", errors="ignore").strip() for ln in f if ln.strip()]
        return [ln for ln in lines if ln]

def _load_master_list(path: str) -> Dict[str, str]:
    """Возвращает словарь canon_key -> original_display_from_master."""
    lines = []
    # попытка через pandas сначала (чтобы корректно разбирались CSV с заголовками и т.д.)
    try:
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
    except Exception:
        lines = []

    if not lines:
        lines = _detect_encoding_and_read_lines(path)

    master_map: Dict[str, str] = {}
    for s in lines:
        key, disp = _parse_to_canon(s)
        if key:
            master_map[key] = disp or s.strip()
        else:
            # простая эвристика: найти цифры и префикс
            m = re.search(r"(\d{1,5}(?:[-/]\d+)*)", s)
            p = re.search(PREFIX_PART, s, flags=re.IGNORECASE)
            y = re.search(r"(\d{4})", s)
            if p and m:
                k = _canon(m.group(1), p.group(0), y.group(1) if y else None)
                if k:
                    master_map[k] = s.strip()
    return master_map

def _has_cyrillic(s: str) -> bool:
    return bool(re.search(r"[А-Яа-яЁё]", s))

def _split_base_year(key: str) -> Tuple[str, str]:
    """Разделяет canon key на base (без :YYYY) и год (или '')."""
    if not key:
        return "", ""
    m = re.match(r"^(.+?)(?::(\d{4}))?$", key)
    if not m:
        return key, ""
    base = m.group(1)
    year = m.group(2) if m.group(2) else ""
    return base, year

# ---------------- основная логика ----------------

def check_pdf(pdf_path: str) -> List[Tuple]:
    print(f"\nПроверка стандартов в {pdf_path}...")

    # загрузка master list
    try:
        master_map = _load_master_list(MASTER_FILE)
    except Exception as e:
        print(f"  [Ошибка загрузки master: {e}]")
        master_map = {}

    master_keys: Set[str] = set(master_map.keys())
    # мапа base -> list of keys (с годами)
    master_base_map: Dict[str, List[str]] = {}
    for k in master_keys:
        base, year = _split_base_year(k)
        master_base_map.setdefault(base, []).append(k)

    # проверка имени файла: если заканчивается на .S (двуязычный), игнорировать кириллицу
    base_filename = os.path.splitext(os.path.basename(pdf_path))[0]
    ignore_cyrillic_doc = bool(re.search(r"\.S($|_)", base_filename, flags=re.IGNORECASE))

    # извлекаем текст
    try:
        text = _extract_text(pdf_path)
    except Exception as e:
        print(f"  [Ошибка чтения PDF: {e}]")
        return []

    # собираем кандидаты (canon_key -> set(display_forms))
    candidates: Dict[str, Set[str]] = {}
    for m in CANDIDATE_RE.finditer(text):
        prefix_chain = m.group("prefix_chain") or ""
        code = m.group("gkinp") or m.group("code") or ""
        year = m.group("year")
        key = _canon(code, _last_base(prefix_chain), year)
        if not key:
            continue
        # nice display
        _, disp = _parse_to_canon(m.group(0) or "")
        if not disp:
            disp = (prefix_chain + " " + (code or "")).strip()
            if year:
                if f":{year}" not in disp:
                    disp = disp + f":{year}"
        candidates.setdefault(key, set()).add(disp)

    if not candidates:
        print("  (ничего похожего на ссылки на стандарты не найдено)")
        return []

    # Убрать кириллицу, если есть латинская версия того же ключа, или если документ двуязычный (.S)
    # Для каждого canon key смотрим disps; если есть disp без кириллицы — фильтруем оставляя только такие;
    # если нет латиницы и doc двуязычный — фильтруем все кириллические disps.
    for key in list(candidates.keys()):
        disps = candidates[key]
        has_latin = any(not _has_cyrillic(d) for d in disps)
        if has_latin:
            candidates[key] = set(d for d in disps if not _has_cyrillic(d))
        elif ignore_cyrillic_doc:
            # если doc двуязычный и нет латинской формы — удаляем кириллические (ничего останется)
            candidates[key] = set(d for d in disps if not _has_cyrillic(d))

    # Вывод результатов: точные попадания, похожие (fuzzy) и "вне master"
    printed: Set[str] = set()
    for key in sorted(candidates.keys()):
        disps = candidates[key]
        if not disps:
            # ничего осталось показать (после фильтрации), пропускаем
            continue
        # берём самую полную форму отображения
        disp = sorted(disps, key=len, reverse=True)[0]

        if key in master_keys:
            print(f"  OK: {disp}")
            continue

        # если base (без года) присутствует в master как варианты — сообщим об этом
        base, cand_year = _split_base_year(key)
        if base in master_base_map:
            variants = master_base_map[base]
            # если candidate содержит год и совпадает по base, но год отличается — сообщим
            found_same_year = False
            for v in variants:
                vb, vy = _split_base_year(v)
                if vy and cand_year and vy == cand_year:
                    found_same_year = True
                    break
            if found_same_year:
                # Попытка найти точное ключ с этим годом — уже handled earlier, но на всякий случай
                print(f"  OK (по базе+году): {disp}")
                continue
            # выводим варианты в master
            variants_disp = ", ".join(sorted(variants))
            print(f"  КОД ИЗВЕСТЕН (варианты в master): {disp} → {variants_disp}")
            continue

        # fuzzy поиск по master_keys
        if master_keys:
            best = process.extractOne(key, master_keys, scorer=fuzz.token_set_ratio)
        else:
            best = None

        if best and best[1] >= FUZZY_THRESHOLD:
            best_key = best[0]
            score = best[1]
            # сравниваем годы, если они есть
            bk_base, bk_year = _split_base_year(best_key)
            cand_base, cand_year = _split_base_year(key)
            master_disp = master_map.get(best_key, best_key)
            if bk_year and cand_year and bk_year != cand_year:
                print(f"  ПОХОЖЕ: {disp} → {master_disp} ({score}%) (год отличается: {cand_year} vs {bk_year})")
            else:
                print(f"  ПОХОЖЕ: {disp} → {master_disp} ({score}%)")
        else:
            # ВНЕ MASTER — реальный новый/неизвестный стандарт
            print(f"  ВНЕ MASTER: {disp}")

    return []  # совместимость с main (ничего не записываем в Excel)

def check(file_path: str):
    return check_pdf(file_path)
