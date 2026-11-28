"""Парсер расписания техникума из xlsx (один файл на день, много групп)."""
import re
from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Optional, Tuple

import openpyxl

GROUP_SPAN = 6                     # каждая группа занимает 6 столбцов (B:G, H:M, N:S, T:Y)
HEADER_COLS = (2, 8, 14, 20)       # столбцы, где стоят названия групп
GROUP_RE = re.compile(r"^\d{2,3}-[А-Яа-яA-Za-z]+")
ROOM_RE = re.compile(r"^\d{3,4}$|этаж|зал|спорт", re.I)
DATE_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4})")


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


def _to_date(text) -> Optional[date]:
    m = DATE_RE.search(str(text or ""))
    if not m:
        return None
    d, mo, y = map(int, m.groups())
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def file_date(ws) -> Optional[date]:
    return _to_date(ws["A1"].value)


def peek_date(path: str) -> Optional[date]:
    """Дата из ячейки A1, не разбирая весь файл."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return _to_date(wb.active["A1"].value)
    finally:
        wb.close()


def _group_headers(ws) -> List[Tuple[int, int, str]]:
    """[(строка, первый столбец, название группы)]."""
    res = []
    for row in ws.iter_rows(min_col=2, max_col=max(HEADER_COLS)):
        for c in row:
            if c.column in HEADER_COLS and isinstance(c.value, str) and GROUP_RE.match(c.value.strip()):
                res.append((c.row, c.column, c.value.strip()))
    return res


def _parse_block(ws, widths, headers, hrow: int, c0: int) -> List[Lesson]:
    c1 = c0 + GROUP_SPAN - 1
    end = ws.max_row + 1
    for r, c, _ in headers:                      # блок кончается на следующем заголовке в тех же столбцах
        if r > hrow and c == c0 and r < end:
            end = r

    def anchors(row):
        out = []
        for col in range(c0, c1 + 1):
            v = ws.cell(row, col).value
            if v is not None and str(v).strip() != "":
                out.append((col, v))
        return out

    lessons: List[Lesson] = []
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
            sub = (sc - c0) // 2 + 1 if widths.get((r, sc), 1) < 4 else None
            lessons.append(Lesson(int(num), _txt(sv), teacher, room, sub))
    lessons.sort(key=lambda l: (l.number, l.subgroup or 0))
    return lessons


def parse_all(path: str) -> Tuple[Optional[date], Dict[str, List[Lesson]]]:
    """Дата файла и пары всех групп: {название группы: [Lesson, ...]}."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    widths = {(m.min_row, m.min_col): m.max_col - m.min_col + 1 for m in ws.merged_cells.ranges}
    headers = _group_headers(ws)
    result: Dict[str, List[Lesson]] = {}
    for hrow, c0, name in headers:
        if name not in result:
            result[name] = _parse_block(ws, widths, headers, hrow, c0)
    return file_date(ws), result


def parse_group(path: str, group: str):
    """(дата файла, пары группы или None, если группы нет в файле)."""
    fdate, allg = parse_all(path)
    for name, lessons in allg.items():
        if name.lower() == group.lower():
            return fdate, lessons
    return fdate, None
