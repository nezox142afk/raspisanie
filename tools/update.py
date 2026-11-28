"""Скачивает расписание с Яндекс Диска, хранит архив по дням и собирает данные для мини-приложения:
docs/index.json (список групп, звонки) и docs/data/<группа>.json (пары группы по дням)."""
import json
import os
import re
import tempfile
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from schedule import parse_all

ROOT = Path(__file__).resolve().parent.parent
DAYS_DIR = ROOT / "data" / "days"
OUT_DIR = ROOT / "docs"
GROUPS_DIR = OUT_DIR / "data"

PUBLIC_KEY = os.getenv("PUBLIC_KEY", "https://disk.yandex.ru/d/bbxOH6odZrMXrQ")
TZ = ZoneInfo(os.getenv("TZ_NAME", "Asia/Novokuznetsk"))
KEEP_DAYS = 21          # сколько последних дней отдаём приложению

DAY_FILES = ["ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА"]

BELLS = {
    "mon": {"extra": "Разговоры о важном  08:30–09:00",
            "pairs": {"1": ["09:10", "10:40"], "2": ["10:55", "12:25"], "3": ["12:55", "14:25"],
                      "4": ["14:40", "16:10"], "5": ["16:20", "17:50"]}},
    "tue_fri": {"extra": "",
                "pairs": {"1": ["08:30", "10:00"], "2": ["10:15", "11:45"], "3": ["12:15", "13:45"],
                          "4": ["14:00", "15:30"], "5": ["15:40", "17:10"], "6": ["17:20", "18:50"]}},
    "sat": {"extra": "",
            "pairs": {"1": ["08:30", "09:30"], "2": ["09:40", "10:40"], "3": ["10:50", "11:50"],
                      "4": ["12:00", "13:00"], "5": ["13:10", "14:10"]}},
}

_TR = dict(zip("абвгдеёжзийклмнопрстуфхцчшщъыьэюя",
               ["a", "b", "v", "g", "d", "e", "e", "zh", "z", "i", "y", "k", "l", "m", "n", "o", "p", "r", "s",
                "t", "u", "f", "h", "c", "ch", "sh", "sch", "", "y", "", "e", "yu", "ya"]))


def slug(name: str) -> str:
    s = "".join(_TR.get(ch, ch) for ch in name.lower())
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-") or "group"


# ---------- скачивание ----------
def http_get(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def fetch_day_file(name: str):
    q = urllib.parse.urlencode({"public_key": PUBLIC_KEY, "path": f"/{name}.xlsx"})
    try:
        href = json.loads(http_get("https://cloud-api.yandex.net/v1/disk/public/resources/download?" + q))["href"]
        return http_get(href)
    except Exception as e:
        print(f"  {name}: нет файла или ошибка ({e})")
        return None


def ingest():
    """Скачать файлы недели и сохранить пары всех групп под датой из таблицы."""
    DAYS_DIR.mkdir(parents=True, exist_ok=True)
    for name in DAY_FILES:
        data = fetch_day_file(name)
        if not data:
            continue
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(data)
            tmp = f.name
        try:
            d, groups = parse_all(tmp)
            if not d or not groups:
                print(f"  {name}: не удалось прочитать (дата={d}, групп={len(groups)}), пропускаю")
                continue
            day = {g: [{"n": l.number, "subject": l.subject, "teacher": l.teacher,
                        "room": l.room, "sub": l.subgroup} for l in lessons]
                   for g, lessons in groups.items()}
            text = json.dumps(day, ensure_ascii=False, indent=0, sort_keys=True)
            path = DAYS_DIR / f"{d.isoformat()}.json"
            if not path.exists() or path.read_text(encoding="utf-8") != text:
                path.write_text(text, encoding="utf-8")
            print(f"  {name}: {d} — групп {len(groups)}")
        except Exception as e:
            print(f"  {name}: ошибка разбора ({e})")
        finally:
            os.unlink(tmp)


# ---------- данные для приложения ----------
def load_days() -> dict:
    days = {}
    for p in sorted(DAYS_DIR.glob("*.json")):
        try:
            v = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        if isinstance(v, dict):
            days[p.stem] = v
    return dict(list(days.items())[-KEEP_DAYS:])


def dump(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"), sort_keys=True)


def build_site():
    days = load_days()
    names = sorted({g for day in days.values() for g in day})
    outputs = {}
    for g in names:
        # группа есть в файле дня — берём её пары; нет в файле — в этот день пар нет
        outputs[GROUPS_DIR / f"{slug(g)}.json"] = dump(
            {"group": g, "days": {d: day.get(g, []) for d, day in days.items()}})

    index = {"groups": [[g, slug(g)] for g in names], "bells": BELLS}
    index_path = OUT_DIR / "index.json"
    changed = False

    for path, text in outputs.items():
        if not path.exists() or path.read_text(encoding="utf-8") != text:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
            changed = True

    old = None
    if index_path.exists():
        try:
            old = json.loads(index_path.read_text(encoding="utf-8"))
        except Exception:
            old = None
    if old is None or {k: v for k, v in old.items() if k != "updated"} != index:
        changed = True
    if changed:
        index["updated"] = datetime.now(TZ).isoformat(timespec="minutes")
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        index_path.write_text(dump(index), encoding="utf-8")
    print(f"Групп: {len(names)}, дней: {len(days)}, изменения: {'да' if changed else 'нет'}")


if __name__ == "__main__":
    print("Скачиваю расписание…")
    ingest()
    build_site()
    print("Готово")
