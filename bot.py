import os
import json
import asyncio

import sys   # ← добавь

# 🔥 LOCK вставка ↓
LOCK_FILE = "bot.lock"

if os.path.exists(LOCK_FILE):
    print("❌ BOT ALREADY RUNNING")
    sys.exit()

with open(LOCK_FILE, "w") as f:
    f.write(str(os.getpid()))
# 🔥 LOCK вставка ↑

from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart

from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton

from config import BOT_TOKEN, ALLOWED_USERS

from datetime import datetime
import pytz


def now():
    return datetime.now()
    

print("🔥 BOT STARTED")

bot = Bot(token=BOT_TOKEN)

import requests

requests.get(
    f"https://api.telegram.org/bot{BOT_TOKEN}/deleteWebhook?drop_pending_updates=true"
)
dp = Dispatcher()

from config_paths import path

# =========================
# SUPPLIERS
# =========================
SUPPLIERS = {
    "top_kitchen": {
        "name": "📦 Top Kitchen",
        "script": "parsers/top_kitchen/run.py",
        "file": path("output/top_kitchen/top_kitchen_LIVE.xlsx"),
        "status": path("output/top_kitchen/status.json"),
        "lock": path("output/top_kitchen/lock.txt"),
    },
    "Харьковская 4425-4426 Gold Top": {
        "name": "📦 Харьковская 4425-4426 Gold Top",
        "script": "parsers/4425-4426_Gold_Top/run.py",
        "file": path("output/4425-4426_Gold_Top/Харьковская_4425-4426_Gold_Top_LIVE.xlsx"),
        "status": path("output/4425-4426_Gold_Top/status.json"),
        "lock": path("output/4425-4426_Gold_Top/lock.txt"), 
        #"file": "output/4425-4426_Gold_Top/Харьковская_4425-4426_Gold_Top_LIVE.xlsx",
        #"status": "output/4425-4426_Gold_Top/status.json",
        #"lock": "output/4425-4426_Gold_Top/lock.txt",
    },
    "Харьковская 208": {
        "name": "📦 Харьковская 208",
        "script": "parsers/208/run.py",
        "file": path("output/208/Харьковская_208_LIVE.xlsx"),
        "status": path("output/208/status.json"),
        "lock": path("output/208/lock.txt"),
    },
        "Харьковская 4421-4422 Jmax": {
        "name": "📦 Харьковская 4421-4422 Jmax",
        "script": "parsers/4421-4422_Jmax/run.py",
        "file": path("output/4421-4422_Jmax/Харьковская_4421-4422_Jmax_LIVE.xlsx"),
        "status": path("output/4421-4422_Jmax/status.json"),
        "lock": path("output/4421-4422_Jmax/lock.txt"),
    },
        "Харьковская 4399-4400": {
        "name": "📦 Харьковская 4399-4400",
        "script": "parsers/4399-4400/run.py",
        "file": path("output/4399-4400/Харьковская_4399-4400_LIVE.xlsx"),
        "status": path("output/4399-4400/status.json"),
        "lock": path("output/4399-4400/lock.txt"),
    },
        "Melad": {
        "name": "📦 Melad",
        "script": "parsers/Melad/run.py",
        "file": path("output/Melad/Melad_LIVE.xlsx"),
        "status": path("output/Melad/status.json"),
        "lock": path("output/Melad/lock.txt"),
    },
         
}

RUNNING_PROCESSES = {}
DASHBOARD_MESSAGES = {}
DASHBOARD_OPENED = set()


def file_time(ts):
    return datetime.fromtimestamp(ts, KYIV)
    
# =========================
# ACCESS
# =========================
def is_allowed(user_id):
    return user_id in ALLOWED_USERS

# =========================
# HELPERS
# =========================
def load_json(path):
    if not path or not os.path.exists(path):
        return None

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # защита от пустого JSON
        if not isinstance(data, dict):
            return None

        return data

    except Exception as e:
        print("❌ JSON READ ERROR:", path, e)
        return None

# =========================
# UI HELPERS
# =========================
def bar(p):
    p = int(p)
    return "█" * (p // 10) + "░" * (10 - p // 10) + f" {p}%"

def get_progress(st):
    if not st:
        return 0

    return int(st.get("progress", 0))


async def card_updater(chat_id, msg_id, key):
    while chat_id in DASHBOARD_OPENED:
        s = SUPPLIERS[key]
        st = load_json(s["status"])

        stt = display_status(st, s["file"])[0]
        p = int(st.get("progress", 0)) if st else 0

        text = f"{s['name']}\n\n"
        text += f"📌 {stt} {p}%\n\n"
        text += bar(p)

        try:
            await bot.edit_message_text(
                chat_id=chat_id,
                message_id=msg_id,
                text=text,
                reply_markup=kb_supplier(key, True)
            )
        except:
            pass

        await asyncio.sleep(2)

def status(st):
    if not st:
        return "⚪ НЕТ ДАННЫХ", 0

    # если есть файл — проверим возраст
    file_path = st.get("file_path")

    if file_path and os.path.exists(file_path):
        file_time = os.path.getmtime(file_path)
        age_days = (now() - datetime.fromtimestamp(file_time)).days

        if age_days >= 5:
            return "⚠️ УСТАРЕЛО", st.get("progress", 100)

    if st.get("running"):
        return "🟡 В РАБОТЕ", st.get("progress", 0)

    if st.get("canceled"):
        return "⛔ ОТМЕНЕНО", st.get("progress", 0)

    return "🟢 ГОТОВО", st.get("progress", 100)

def age_days(st):
    if not st or not st.get("time"):
        return 999
    try:
        from datetime import timezone
        
        t = datetime.strptime(st["time"], "%Y-%m-%d %H:%M")
        t = pytz.timezone("Europe/Kyiv").localize(t)
        
        return (now() - t).days
    except:
        return 999

def get_file_time(path):
    if not os.path.exists(path):
        return None

    ts = os.path.getmtime(path)
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")

# =========================
# SAFE EDIT (FIX AIogram ERROR)
# =========================
async def safe_edit(call, text, kb=None):
    try:
        await call.message.edit_text(text, reply_markup=kb)
    except:
        await call.message.answer(text, reply_markup=kb)


def display_status(st, file_path):

    if not st:
        return "⚪ НЕТ ДАННЫХ", 0

    # всегда берём путь из SUPPLIERS (это главный источник)
    exists = file_path and os.path.exists(file_path)

    if st.get("running"):
        return "🟡 В РАБОТЕ", max(1, int(st.get("progress", 1)))

    if st.get("canceled"):
        return "⛔ ОТМЕНЕНО", int(st.get("progress", 0))

    if exists:
        return "🟢 ГОТОВО", 100

    return "⚪ ФАЙЛ ОТСУТСТВУЕТ", 0

# =========================
# DASHBOARD
# =========================
def dashboard_text():
    t = "📊 Дашборд парсеров\n\n"

    for k, s in SUPPLIERS.items():
        print("READ STATUS:", s["status"])
        st = load_json(s["status"])

        stt, p = display_status(st, s["file"])

        warn = ""

        if st and not st.get("running") and os.path.exists(s["file"]):

            kyiv = pytz.timezone("Europe/Kyiv")

            age = (
                now() - datetime.fromtimestamp(os.path.getmtime(s["file"]))
            ).days
           
            if age >= 3:
                warn = "⚠️"

        mini_bar = "█" * (p // 20) + "░" * (5 - p // 20)

        t += f"{s['name']}\n{stt} {mini_bar} {p}% {warn}\n\n"

    return t

def kb_dashboard():
    rows = []

    for k, s in SUPPLIERS.items():
        st = load_json(s["status"])

        stt, p = display_status(st, s["file"])

        mini_bar = "█" * (p // 20) + "░" * (5 - p // 20)

        rows.append([
            InlineKeyboardButton(
                text=f"{s['name']} | {stt} {mini_bar} {p}%",
                callback_data=f"open_{k}"
            )
        ])

    return InlineKeyboardMarkup(inline_keyboard=rows)


def kb_supplier(key, running=False):

    if key not in SUPPLIERS:
        return InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="❌ НЕ НАЙДЕНО", callback_data="back")]
        ])
    st = load_json(SUPPLIERS[key]["status"]) or {}
    file_path = st.get("file_path") if st else SUPPLIERS[key]["file"]
    file_exists = file_path and os.path.exists(file_path)

    if running:
        return InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="⛔ ОТМЕНА", callback_data=f"cancel_{key}")],
            [InlineKeyboardButton(text="🔙 ДАШБОРД", callback_data="back")]
        ])

    rows = []

    if file_exists:
        rows.append([
            InlineKeyboardButton(
                text="⬇️ СКАЧАТЬ EXCEL",
                callback_data=f"download_{key}"
            )
        ])

    rows.append([
        InlineKeyboardButton(
            text="▶ ЗАПУСТИТЬ ЗАНОВО",
            callback_data=f"run_{key}"
        )
    ])

    rows.append([
        InlineKeyboardButton(
            text="🔙 ДАШБОРД",
            callback_data="back"
        )
    ])

    return InlineKeyboardMarkup(inline_keyboard=rows)

# =========================
# START
# =========================
#@dp.message()
@dp.message(CommandStart())
async def start(message: types.Message):
    if not is_allowed(message.from_user.id):
        return

    msg = await message.answer(
        dashboard_text(),
        reply_markup=kb_dashboard()
    )

    DASHBOARD_MESSAGES[message.chat.id] = msg.message_id
    
# =========================
# RUN PARSER
# =========================
async def run_parser(key, user):
    s = SUPPLIERS[key]

    import sys

    proc = await asyncio.create_subprocess_exec(
        sys.executable,
        s["script"],
        user,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE
    )

    RUNNING_PROCESSES[key] = proc

    print(f"🚀 STARTED {key}")
    print(f"📄 SCRIPT: {s['script']}")
    print("WRITE STATUS:", s["status"])

    try:
        while True:
            # читаем статус отмены
            st = load_json(s["status"])

            if st and st.get("canceled"):
                try:
                    import psutil
                    parent = psutil.Process(proc.pid)

                    for child in parent.children(recursive=True):
                        try:
                            child.kill()
                        except:
                            pass

                    try:
                        parent.kill()
                    except:
                        pass

                except Exception as e:
                    print("KILL ERROR:", e)

                RUNNING_PROCESSES.pop(key, None)
                return "canceled"

            # если процесс умер
            if proc.returncode is not None:
                break

            await asyncio.sleep(1)

        stdout, stderr = await proc.communicate()

        if stdout:
            print(stdout.decode(errors="ignore"))

        if stderr:
            print(stderr.decode(errors="ignore"))

        RUNNING_PROCESSES.pop(key, None)
        return proc.returncode

    except Exception as e:
        RUNNING_PROCESSES.pop(key, None)
        print("RUN ERROR:", e)
        return -1
# =========================
# CALLBACKS
# =========================
@dp.callback_query()
async def cb(call: types.CallbackQuery):
    if not is_allowed(call.from_user.id):
        return

    await call.answer()
    data = call.data

    # BACK
    if data == "back":

        msg = await call.message.edit_text(
            dashboard_text(),
            reply_markup=kb_dashboard()
        )

        DASHBOARD_MESSAGES[
            call.message.chat.id
        ] = msg.message_id
        
        DASHBOARD_OPENED.discard(call.message.chat.id)

        return
    
    # OPEN CARD
    if data.startswith("open_"):
        key = data.replace("open_", "")
        s = SUPPLIERS[key]
        st = load_json(s["status"])
        

        DASHBOARD_OPENED.add(call.message.chat.id)

        stt, p = display_status(st, s["file"])

        if stt == "🟢 ГОТОВО":
            p = 100

        running = st.get("running") if st else False

        text = f"{s['name']}\n\n"
        text += f"📌 {stt} {p}%\n\n"

        if st:
            text += f"👤 Запустил: {st.get('user','-')}\n"
            text += f"🕒 Время: {st.get('time','-')}\n\n"

        text += bar(p)

        file_path = st.get("file_path") if st else s["file"]
        file_exists = file_path and os.path.exists(file_path)

        file_old = False
        file_time = None

        if file_exists:
            file_time_raw = os.path.getmtime(file_path)
            file_age_days = (now() - datetime.fromtimestamp(file_time_raw)).days

            file_old = file_age_days >= 3
            file_time = get_file_time(s["file"])

        # 🔥 1) ЕСЛИ БЫЛ ОТМЕНЁН ПАРСИНГ
        if st and st.get("canceled"):
            text += "\n\n⛔ Парсинг был ОТМЕНЁН\n📄 Файл может быть неполным"

        # 🔥 2) ЕСЛИ ФАЙЛ СЛИШКОМ СТАРЫЙ
        elif file_exists and file_old:
            size_mb = round(os.path.getsize(file_path) / 1024 / 1024, 2)

            text += (
                f"\n\n⚠️ ДАННЫЕ УСТАРЕЛИ\n"
                f"📄 Excel есть, но он старый\n"
                f"📦 Размер: {size_mb} МБ\n"
                f"🕒 Обновлён: {file_time}"
            )

        # 🔥 3) НОРМАЛЬНЫЙ ФАЙЛ

        elif file_exists:
            size_mb = round(os.path.getsize(file_path) / 1024 / 1024, 2)
            file_time = get_file_time(file_path)

            text += (
                f"\n\n📄 Excel готов\n"
                f"📦 Размер: {size_mb} МБ\n"
                f"🕒 Обновлён: {file_time}"
            )

        else:
            text += "\n\n📄 Excel отсутствует"

        await call.message.edit_text(
            text,
            reply_markup=kb_supplier(key, running)
        )

    # DOWNLOAD FILE
    if data.startswith("download_"):
        key = data.replace("download_", "")
        st = load_json(SUPPLIERS[key]["status"]) or {}
        path = st.get("file_path") if st else SUPPLIERS[key]["file"]

        if os.path.exists(path):
            await call.message.answer_document(
                types.FSInputFile(path)
            )
        else:
            await call.answer(
                "Файл не найден",
                show_alert=True
            )

        return

    # RUN
    if data.startswith("run_"):
        key = data.replace("run_", "").strip()
        
        if key not in SUPPLIERS:
            await call.message.answer("❌ Парсер не найден")
            return

        s = SUPPLIERS[key]


        await call.message.edit_text(
            "🚀 Запуск...\n" + bar(0),
            reply_markup=kb_supplier(key, True)
        )

        try:
            # запускаем парсер
            code = await run_parser(key, call.from_user.full_name)

            # если его НЕ отменили вручную
            st = load_json(s["status"]) or {}
            
            await call.message.edit_text(
                "✅ ГОТОВО\n" + bar(100),
                reply_markup=kb_supplier(key, False)
            )

        except Exception as e:

            await call.message.edit_text(
                f"❌ ОШИБКА\n{e}",
                reply_markup=kb_supplier(key, False)
            )

        return

    # CANCEL
    if data.startswith("cancel_"):
        key = data.replace("cancel_", "")
        s = SUPPLIERS[key]

        proc = RUNNING_PROCESSES.get(key)

        DASHBOARD_OPENED.discard(call.message.chat.id)

        if proc:
            try:
                import psutil

                parent = psutil.Process(proc.pid)

                # 💣 убиваем ВСЁ дерево (python + chromedriver + chrome)
                children = parent.children(recursive=True)

                for child in children:
                    try:
                        child.kill()
                    except:
                        pass

                try:
                    parent.kill()
                except:
                    pass

            except Exception as e:
                print("KILL ERROR:", e)

            RUNNING_PROCESSES.pop(key, None)

        # 🧹 lock файл
        if "lock" in s and os.path.exists(s["lock"]):
            try:
                os.remove(s["lock"])
            except:
                pass

        await safe_edit(call, "⛔ ОТМЕНЕНО\n" + bar(0), kb_supplier(key, False))
        return

async def dashboard_updater():

    while True:

        for chat_id, msg_id in list(DASHBOARD_MESSAGES.items()):
            # НЕ ОБНОВЛЯЕМ, если пользователь сейчас в карточке
            if chat_id in DASHBOARD_OPENED:
                continue
            
            try:
                await bot.edit_message_text(
                    chat_id=chat_id,
                    message_id=msg_id,
                    text=dashboard_text(),
                    reply_markup=kb_dashboard()
                )

            except Exception:
                pass

        await asyncio.sleep(5)
        

from aiogram.filters import Command

@dp.message(Command("status"))
async def debug_status(message: types.Message):

    path = "/app/output/4425-4426_Gold_Top/status.json"

    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            await message.answer(f"<pre>{f.read()}</pre>", parse_mode="HTML")
    else:
        await message.answer("❌ status.json не найден")

@dp.message(Command("where"))
async def where(message: types.Message):
    await message.answer(os.getcwd())


# =========================
# MAIN
# =========================
async def main():
    print("🚀 BOT RUNNING")

    asyncio.create_task(
        dashboard_updater()
    )

    await dp.start_polling(bot)

import atexit

def remove_lock():
    if os.path.exists("bot.lock"):
        os.remove("bot.lock")

atexit.register(remove_lock)

if __name__ == "__main__":
    asyncio.run(main())
