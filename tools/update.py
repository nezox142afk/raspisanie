"""Скачивает расписание с Яндекс Диска, хранит архив по дням, собирает docs/data.json
и раз в день присылает пары на завтра в Телеграм."""
import json
import os
import tempfile
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from schedule import parse_group, peek_date

ROOT = Path(__file__).resolve().parent.parent
DAYS_DIR = ROOT / "data" / "days"
STATE = ROOT / "data" / "state.json"
SITE_DATA = ROOT / "docs" / "data.json"

GROUP = os.getenv("GROUP", "263-СП")
PUBLIC_KEY = os.getenv("PUBLIC_KEY", "https://disk.yandex.ru/d/bbxOH6odZrMXrQ")
TZ = ZoneInfo(os.getenv("TZ_NAME", "Asia/Novokuznetsk"))
SEND_TIME = os.getenv("SEND_TIME", "20:00")
PAGES_URL = os.getenv("PAGES_URL", "")
KEEP_DAYS = 60

DAY_FILES = ["ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА"]
WD = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]

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


def bells_for(d: date) -> dict:
    return BELLS["mon"] if d.weekday() == 0 else BELLS["sat"] if d.weekday() == 5 else BELLS["tue_fri"]


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
    """Скачать файлы недели, сохранить пары группы по датам из таблиц."""
    DAYS_DIR.mkdir(parents=True, exist_ok=True)
    for name in DAY_FILES:
        data = fetch_day_file(name)
        if not data:
            continue
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(data)
            tmp = f.name
        try:
            d = peek_date(tmp)
            if not d:
                print(f"  {name}: нет даты в A1, пропускаю")
                continue
            _, lessons = parse_group(tmp, GROUP)
            items = [{"n": l.number, "subject": l.subject, "teacher": l.teacher,
                      "room": l.room, "sub": l.subgroup} for l in (lessons or [])]
            (DAYS_DIR / f"{d.isoformat()}.json").write_text(
                json.dumps(items, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
            print(f"  {name}: {d} — пар {len(items)}")
        except Exception as e:
            print(f"  {name}: ошибка разбора ({e})")
        finally:
            os.unlink(tmp)


# ---------- сайт ----------
def load_days() -> dict:
    days = {}
    for p in sorted(DAYS_DIR.glob("*.json")):
        days[p.stem] = json.loads(p.read_text(encoding="utf-8"))
    return dict(list(days.items())[-KEEP_DAYS:])


def build_site():
    new = {"group": GROUP, "bells": BELLS, "days": load_days()}
    old = None
    if SITE_DATA.exists():
        try:
            old = json.loads(SITE_DATA.read_text(encoding="utf-8"))
        except Exception:
            old = None
    if old and {k: v for k, v in old.items() if k != "updated"} == new:
        return  # ничего не изменилось — не плодим лишние коммиты
    new["updated"] = datetime.now(TZ).isoformat(timespec="minutes")
    SITE_DATA.parent.mkdir(parents=True, exist_ok=True)
    SITE_DATA.write_text(json.dumps(new, ensure_ascii=False, indent=1), encoding="utf-8")


# ---------- сообщение в Телеграм ----------
def esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def render_message(d: date, items) -> str:
    head = f"📅 <b>{WD[d.weekday()].capitalize()}, {d:%d.%m.%Y}</b> · {esc(GROUP)}\n"
    if not items:
        return head + "\nПар нет 🎉"
    pairs = bells_for(d)["pairs"]
    out = [head]
    for l in sorted(items, key=lambda x: (x["n"], x["sub"] or 0)):
        a, b = pairs.get(str(l["n"]), ["", ""])
        t = f"{a} — {b}" if a else ""
        room = f"  ·  <code>{esc(l['room'])}</code>" if l["room"] else ""
        sub = f" (подгр. {l['sub']})" if l["sub"] else ""
        out.append(f"<b>[{l['n']}]</b> {t}{room}\n<b>{esc(l['subject'])}</b>{sub}\n<i>{esc(l['teacher'])}</i>\n")
    return "\n".join(out)


def send(text: str) -> bool:
    token, chat = os.getenv("BOT_TOKEN"), os.getenv("CHAT_ID")
    if not token or not chat:
        print("BOT_TOKEN/CHAT_ID не заданы — сообщение не отправляю")
        return False
    body = {"chat_id": chat, "text": text, "parse_mode": "HTML"}
    if PAGES_URL:
        body["reply_markup"] = {"inline_keyboard": [[
            {"text": "📱 Открыть расписание", "web_app": {"url": PAGES_URL}}]]}
    req = urllib.request.Request(f"https://api.telegram.org/bot{token}/sendMessage",
                                 data=json.dumps(body).encode(),
                                 headers={"Content-Type": "application/json"})
    try:
        urllib.request.urlopen(req, timeout=30).read()
        return True
    except Exception as e:
        print("Ошибка отправки:", e)
        return False


def notify():
    now = datetime.now(TZ)
    h, m = map(int, SEND_TIME.split(":"))
    if (now.hour, now.minute) < (h, m):
        return
    state = {}
    if STATE.exists():
        state = json.loads(STATE.read_text(encoding="utf-8"))
    today = now.date().isoformat()
    if state.get("last_sent") == today:
        return
    tomorrow = now.date() + timedelta(days=1)
    if tomorrow.weekday() == 6:                      # завтра воскресенье — молчим
        state["last_sent"] = today
    else:
        items = load_days().get(tomorrow.isoformat())
        if items is not None:
            if send(render_message(tomorrow, items)):
                state["last_sent"] = today
        elif now.hour >= 22:                         # к 22:00 так и не выложили
            if send(f"На {tomorrow:%d.%m} расписание ещё не выложили."):
                state["last_sent"] = today
    STATE.write_text(json.dumps(state), encoding="utf-8")


if __name__ == "__main__":
    print("Скачиваю расписание…")
    ingest()
    build_site()
    notify()
    print("Готово")
