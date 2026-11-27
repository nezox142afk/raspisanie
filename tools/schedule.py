"""Парсер расписания техникума из xlsx (один файл на день)."""
import re
from dataclasses import dataclass
from datetime import date
from typing import Optional

import openpyxl
from openpyxl.utils import get_column_letter

GROUP_SPAN = 6          # каждая группа занимает 6 столбцов (B:G, H:M, N:S, T:Y)
GROUP_RE = re.compile(r"^\d{2,3}-[А-Яа-яA-Za-z]+")
ROOM_RE = re.compile(r"^\d{3,4}$|этаж|зал|спорт", re.I)


@dataclass
class Lesson:
    number: int
    subject: str
    teacher: str
    room: str
    subgroup: Optional[int] = None


def _is_room(v) -> bool:
    if isinstance(v, (int, float)):
        return True
    return bool(ROOM_RE.search(str(v).strip()))


def _txt(v) -> str:
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def file_date(ws) -> Optional[date]:
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", str(ws["A1"].value or ""))
    if not m:
        return None
    d, mo, y = map(int, m.groups())
    return date(y, mo, d)


def _group_headers(ws):
    """[(строка, первый столбец, название группы)] по всему листу."""
    res = []
    for row in ws.iter_rows(min_col=2):
        for c in row:
            if isinstance(c.value, str) and GROUP_RE.match(c.value.strip()):
                res.append((c.row, c.column, c.value.strip()))
    return res


def parse_group(path: str, group: str):
    """Возвращает (дата файла, список пар) или (дата, None), если группы нет в файле."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    fdate = file_date(ws)
    headers = _group_headers(ws)
    target = next((h for h in headers if h[2].lower() == group.lower()), None)
    if not target:
        return fdate, None
    hrow, c0, _ = target
    c1 = c0 + GROUP_SPAN - 1

    # конец блока — ближайшая следующая строка с заголовком группы в тех же столбцах
    end = ws.max_row + 1
    for r, c, _name in sorted(headers):
        if r > hrow and c0 <= c <= c1:
            end = r
            break

    # ширина объединённых ячеек: (строка, столбец) -> ширина
    widths = {}
    for m in ws.merged_cells.ranges:
        widths[(m.min_row, m.min_col)] = m.max_col - m.min_col + 1

    def anchors(row):
        out = []
        for col in range(c0, c1 + 1):
            v = ws.cell(row, col).value
            if v is not None and str(v).strip() != "":
                out.append((col, v))
        return out

    lessons = []
    for r in range(hrow + 1, end):
        num = ws.cell(r, 1).value
        if not isinstance(num, (int, float)):
            continue
        row1, row2 = anchors(r), anchors(r + 1)
        subjects = [(c, v) for c, v in row1 if not _is_room(v)]
        rooms = [(c, v) for c, v in row1 if _is_room(v)]
        for i, (sc, sv) in enumerate(subjects):
            nxt = subjects[i + 1][0] if i + 1 < len(subjects) else c1 + 1
            room = next((_txt(v) for c, v in rooms if sc < c < nxt), "")
            teacher = next((_txt(v) for c, v in row2 if sc <= c < nxt), "")
            sub = None
            if widths.get((r, sc), 1) < 4:
                sub = (sc - c0) // 2 + 1
            lessons.append(Lesson(int(num), _txt(sv), teacher, room, sub))
    lessons.sort(key=lambda l: (l.number, l.subgroup or 0))
    return fdate, lessons


def list_groups(path: str):
    wb = openpyxl.load_workbook(path, data_only=True)
    return sorted({h[2] for h in _group_headers(wb.active)})


def peek_date(path: str) -> Optional[date]:
    """Дата из ячейки A1, не разбирая весь файл."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", str(wb.active["A1"].value or ""))
    finally:
        wb.close()
    if not m:
        return None
    d, mo, y = map(int, m.groups())
    return date(y, mo, d)
