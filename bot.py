import os
import json
import asyncio
import time
import sys

LOCK_FILE = "bot.lock"

if os.path.exists(LOCK_FILE):
    print("❌ BOT ALREADY RUNNING")
    sys.exit()

with open(LOCK_FILE, "w") as f:
    f.write(str(os.getpid()))


from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart

from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton

from config import BOT_TOKEN, ALLOWED_USERS

from datetime import datetime
import pytz

last_edit = {}

async def safe_edit_message(chat_id, message_id, text, kb=None):
    import time

    key = f"{chat_id}:{message_id}"
    now = time.time()

    if now - last_edit.get(key, 0) < 3:
        return  # не чаще чем раз в 3 сек

    last_edit[key] = now

    try:
        await bot.safe_edit_message(
            chat_id=chat_id,
            message_id=message_id,
            text=text,
            reply_markup=kb
        )
    except:
        pass


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
        "file": path("output/4425-4426_Gold_Top/4425-4426_Gold_Top_LIVE.xlsx"),
        "status": path("output/4425-4426_Gold_Top/status.json"),
        "lock": path("output/4425-4426_Gold_Top/lock.txt"), 
    },
    "Харьковская 208": {
        "name": "📦 Харьковская 208",
        "script": "parsers/208/run.py",
        "file": path("output/208/208_LIVE.xlsx"),
        "status": path("output/208/status.json"),
        "lock": path("output/208/lock.txt"),
    },
        "Харьковская 4421-4422 Jmax": {
        "name": "📦 Харьковская 4421-4422 Jmax",
        "script": "parsers/4421-4422_Jmax/run.py",
        "file": path("output/4421-4422_Jmax/4421-4422_Jmax_LIVE.xlsx"),
        "status": path("output/4421-4422_Jmax/status.json"),
        "lock": path("output/4421-4422_Jmax/lock.txt"),
    },
        "Харьковская 4399-4400": {
        "name": "📦 Харьковская 4399-4400",
        "script": "parsers/4399-4400/run.py",
        "file": path("output/4399-4400/4399-4400_LIVE.xlsx"),
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
        "Melad": {
            "name": "📦 Харьковская 220 Kithen Plus",
            "script": "parsers/220_Kithen_Plus/run.py",
            "file": path("output/220_Kithen_Plus/220_Kithen_Plus_LIVE.xlsx"),
            "status": path("output/220_Kithen_Plus/status.json"),
            "lock": path("output/220_Kithen_Plus/lock.txt"),
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



def ensure_status():
    for s in SUPPLIERS.values():
        os.makedirs(os.path.dirname(s["status"]), exist_ok=True)

        if not os.path.exists(s["status"]):
            with open(s["status"], "w", encoding="utf-8") as f:

                json.dump({
                    "running": False,
                    "progress": 0,
                    "user": "",
                    "time": ""
                }, f, ensure_ascii=False, indent=2)

ensure_status()
# =========================
# HELPERS
# =========================
def load_json(path):
    try:
        # читаем как текст
        with open(path, "r", encoding="utf-8") as f:
            text = f.read().strip()

        # 🔥 ЗАЩИТА ОТ ПУСТОГО ФАЙЛА
        if not text:
            return {
                "running": False,
                "progress": 0,
                "user": "",
                "time": ""
            }
            

        return json.loads(text)

    except Exception as e:
        print("❌ JSON READ ERROR:", path, e)

        return {
            "running": False,
            "progress": 0,
            "user": "",
            "time": ""
        }


ANIM = [
    "▰▱▱▱▱▱▱▱▱▱",
    "▰▰▱▱▱▱▱▱▱▱",
    "▰▰▰▱▱▱▱▱▱▱",
    "▰▰▰▰▱▱▱▱▱▱",
    "▰▰▰▰▰▱▱▱▱▱",
    "▰▰▰▰▰▰▱▱▱▱",
    "▰▰▰▰▰▰▰▱▱▱",
    "▰▰▰▰▰▰▰▰▱▱",
    "▰▰▰▰▰▰▰▰▰▱",
    "▰▰▰▰▰▰▰▰▰▰",
]

def anim_bar(step: int):
    return ANIM[step % len(ANIM)]

# =========================
# UI HELPERS
# =========================

def get_progress(st):
    if not st:
        return 0

    return int(st.get("progress", 0))


async def card_updater(chat_id, msg_id, key):
    while chat_id in DASHBOARD_OPENED:

        s = SUPPLIERS[key]
        st = load_json(s["status"])

        stt, p = display_status(st, s["file"])

        # 🧠 ВАЖНО: создаём text заново
        text = f"{s['name']}\n\n"
        text += f"📌 {stt}\n\n"

        step = int(time.time() * 2)

        text += anim_bar(step)

        try:
            await bot.safe_edit_message(
                chat_id=chat_id,
                message_id=msg_id,
                text=text,
                reply_markup=kb_supplier(key, True)
            )
        except:
            pass

        await asyncio.sleep(1)

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

    if st.get("running"):
        return "🟡 В РАБОТЕ", int(st.get("progress", 0))

    if os.path.exists(file_path):
        return "🟢 ГОТОВО", 100

    return "⚪ ФАЙЛ ОТСУТСТВУЕТ", 0

# =========================
# DASHBOARD
# =========================
def dashboard_text():
    t = "📊 Дашборд парсеров\n\n"

    for k, s in SUPPLIERS.items():
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

        if st.get("running"):
            mini_bar = anim_bar(int(time.time() * 2))
        elif os.path.exists(s["file"]):
            mini_bar = anim_bar(9)
        else:
            mini_bar = ""


        t += f"{s['name']}\n{stt} {mini_bar} {p}% {warn}\n\n"

    return t

def kb_dashboard():
    rows = []

    for k, s in SUPPLIERS.items():
        st = load_json(s["status"])

        stt, p = display_status(st, s["file"])


        if st.get("running"):
            mini_bar = anim_bar(int(time.time() * 2))
        elif os.path.exists(s["file"]):
            mini_bar = anim_bar(9)
        else:
            mini_bar = ""

        rows.append([
            InlineKeyboardButton(
                text=f"{s['name']} | {stt} {mini_bar} {p}%",
                callback_data=f"open_{k}"
            )
        ])

    return InlineKeyboardMarkup(inline_keyboard=rows)


def kb_supplier(key, running=False):

    if key not in SUPPLIERS:
        return InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="❌ НЕ НАЙДЕНО",
                        callback_data="back"
                    )
                ]
            ]
        )

    file_exists = os.path.exists(SUPPLIERS[key]["file"])

    if running:
        return InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="⛔ ОТМЕНА",
                        callback_data=f"cancel_{key}"
                    )
                ],
                [
                    InlineKeyboardButton(
                        text="🔙 ДАШБОРД",
                        callback_data="back"
                    )
                ]
            ]
        )

    rows = []

    if file_exists:
        rows.append([
            InlineKeyboardButton(
                text="⬇️ Скачать Excel",
                callback_data=f"download_{key}"
            )
        ])

    rows.append([
        InlineKeyboardButton(
            text="▶ Запустить заново",
            callback_data=f"run_{key}"
        )
    ])

    rows.append([
        InlineKeyboardButton(
            text="🔙 Дашборд",
            callback_data="back"
        )
    ])

    return InlineKeyboardMarkup(
        inline_keyboard=rows
    )

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
    
        user = st.get("user", "-")





        
        

        from datetime import timedelta

        run_time = st.get("time", "-")
        
        try:
            dt = datetime.strptime(run_time, "%Y-%m-%d %H:%M:%S")
        
            # +3 часа
            dt = dt + timedelta(hours=3)
        
            # формат без секунд
            run_time = dt.strftime("%d.%m.%Y %H:%M")
        
        except:
            pass







        
        #run_time = st.get("time", "-")
    
        text = f"📦 <b>{s['name']}</b>\n\n"
    
        text += f"📌 <b>Статус:</b> {stt}\n"
        text += f"👤 <b>Пользователь:</b> {user}\n"
        text += f"🕒 <b>Запуск:</b> {run_time}\n\n"

        if st.get("running"):
            text += "\n" + anim_bar(int(time.time() * 2))
        elif os.path.exists(s["file"]):
            text += "\n" + anim_bar(9)
            
        if os.path.exists(s["file"]):
    
            size_mb = round(
                os.path.getsize(s["file"]) / 1024 / 1024,
                2
            )
    
            mtime = os.path.getmtime(s["file"])
    
            from datetime import timedelta

            dt = datetime.fromtimestamp(mtime)
            
            # +3 часа
            dt = dt + timedelta(hours=3)
            
            dt_str = dt.strftime("%d.%m.%Y %H:%M")
            
            
            age = (
                datetime.now() -
                datetime.fromtimestamp(mtime)
            ).days
    
            if age >= 3:
                state = f"⚠️ Устарел ({age} дн.)"
            else:
                state = "✅ Актуальный"
    
            text += (
                "\n\n"
                "📄 <b>Excel</b>\n"
                f"├ Размер: {size_mb} МБ\n"
                f"├ Обновлён: {dt_str}\n"
                f"└ Состояние: {state}"
            )
    
        else:
    
            text += (
                "\n\n"
                "📄 <b>Excel</b>\n"
                "└ ❌ Файл отсутствует"
            )
    
        await call.message.edit_text(
            text,
            reply_markup=kb_supplier(
                key,
                st.get("running", False)
            ),
            parse_mode="HTML"
        )
    
        return
    
    # DOWNLOAD FILE

    if data.startswith("download_"):
    
        key = data.replace("download_", "")
    
        path = SUPPLIERS[key]["file"]
    
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
            "🚀 Запуск...\n" + anim_bar(0),
            reply_markup=kb_supplier(key, True)
        )

        try:
            # запускаем парсер
            code = await run_parser(key, call.from_user.full_name)

            # если его НЕ отменили вручную
            st = load_json(s["status"]) or {}
            
            await call.message.edit_text(
                "✅ ГОТОВО\n" + anim_bar(9),
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

        await safe_edit(call, "⛔ ОТМЕНЕНО\n" + anim_bar(0), kb_supplier(key, False))
        return

async def dashboard_updater():

    while True:

        for chat_id, msg_id in list(DASHBOARD_MESSAGES.items()):
            # НЕ ОБНОВЛЯЕМ, если пользователь сейчас в карточке
            if chat_id in DASHBOARD_OPENED:
                continue
            
            try:
                await bot.safe_edit_message(
                    chat_id=chat_id,
                    message_id=msg_id,
                    text=dashboard_text(),
                    reply_markup=kb_dashboard()
                )

            except Exception:
                pass

        await asyncio.sleep(5)
        

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
