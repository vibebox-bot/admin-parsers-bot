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

async def safe_edit_message(chat_id, message_id, text, kb=None, parse_mode=None):
    key = f"{chat_id}:{message_id}"
    last_edit[key] = time.time()

    try:
        await bot.edit_message_text(
            chat_id=chat_id,
            message_id=message_id,
            text=text,
            reply_markup=kb,
            parse_mode=parse_mode
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
        "Харьковская 220 Kithen Plus": {
            "name": "📦 Харьковская 220 Kithen Plus",
            "script": "parsers/220_Kithen_Plus/run.py",
            "file": path("output/220_Kithen_Plus/220_Kithen_Plus_LIVE.xlsx"),
            "status": path("output/220_Kithen_Plus/status.json"),
            "lock": path("output/220_Kithen_Plus/lock.txt"),
        },
        "Dellta": {
            "name": "📦 Dellta",
            "script": "parsers/Dellta/run.py",
            "file": path("output/Dellta/Dellta_LIVE.xlsx"),
            "status": path("output/Dellta/status.json"),
            "lock": path("output/Dellta/lock.txt"),
        },
        "Харьковская 239": {
            "name": "📦 Харьковская 239",
            "script": "parsers/239/run.py",
            "file": path("output/239/239_LIVE.xlsx"),
            "status": path("output/239/status.json"),
            "lock": path("output/239/lock.txt"),
        },
        "Харьковская 228 Rainberg": {
            "name": "📦 Харьковская 228 Rainberg",
            "script": "parsers/228_Rainberg/run.py",
            "file": path("output/228_Rainberg/228_Rainberg_LIVE.xlsx"),
            "status": path("output/228_Rainberg/status.json"),
            "lock": path("output/228_Rainberg/lock.txt"),
        }, 
        "Фабричная 626": {
            "name": "📦 Фабричная 626",
            "script": "parsers/626/run.py",
            "file": path("output/626/626_LIVE.xlsx"),
            "status": path("output/626/status.json"),
            "lock": path("output/626/lock.txt"),
        },
        "Харьковская 205 AND": {
            "name": "📦 Харьковская 205 AND",
            "script": "parsers/205_AND/run.py",
            "file": path("output/205_AND/205_AND_LIVE.xlsx"),
            "status": path("output/205_AND/status.json"),
            "lock": path("output/205_AND/lock.txt"),
        },
        "Харьковская D-Top": {
            "name": "📦 Харьковская D-Top",
            "script": "parsers/D-Top/run.py",
            "file": path("output/D-Top/D-Top_LIVE.xlsx"),
            "status": path("output/D-Top/status.json"),
            "lock": path("output/D-Top/lock.txt"),
        },  
        "Харьковская 219 Магнит": {
            "name": "📦 Харьковская 219 Магнит",
            "script": "parsers/219/run.py",
            "file": path("output/219/219_LIVE.xlsx"),
            "status": path("output/219/status.json"),
            "lock": path("output/219/lock.txt"),
        },
        "Харьковская КМТ": {
            "name": "📦 Харьковская КМТ",
            "script": "parsers/КМТ/run.py",
            "file": path("output/КМТ/КМТ_LIVE.xlsx"),
            "status": path("output/КМТ/status.json"),
            "lock": path("output/КМТ/lock.txt"),
        },    
         
}


RUNNING_PROCESSES = {}
DASHBOARD_MESSAGES = {}
DASHBOARD_OPENED = set()

def is_running(key, st):
    return st.get("running") or key in RUNNING_PROCESSES


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
                    "time": "",
                    "canceled": False,
                    "success": False
                }, f, ensure_ascii=False, indent=2)
                

        else:
            st = load_json(s["status"])

            # если lock-файла нет — значит парсер уже не работает
            if not os.path.exists(s["lock"]):

                st["running"] = False
                st["canceled"] = False

                with open(s["status"], "w", encoding="utf-8") as f:
                    json.dump(st, f, ensure_ascii=False, indent=2)

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

ensure_status()

BLINK = ["🔴", "⚫"]

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

        stt, p = display_status(key, st, s["file"])

        # 🧠 ВАЖНО: создаём text заново
        text = f"{s['name']}\n\n"
        text += f"{stt}\n\n"
        
        try:
            await safe_edit_message(
                chat_id=chat_id,
                message_id=msg_id,
                text=text,
                reply_markup=kb_supplier(key, True)
            )
        except:
            pass

        await asyncio.sleep(0.6)

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


def display_status(key, st, file_path):

    if key in RUNNING_PROCESSES:
        return "🟡 В РАБОТЕ", st.get("progress", 0)

    if st.get("canceled"):
        return "⛔ ОТМЕНЕНО", 0

    if not os.path.exists(file_path):
        return "⚪ НЕТ ФАЙЛА", 0

    size = os.path.getsize(file_path)

    if size < 50:
        return "🔴 ОШИБКА", 0

    return "🟢 ГОТОВО", 100

# =========================
# DASHBOARD
# =========================
def dashboard_text():
    t = "📊 <b>Дашборд парсеров</b>\n\n"

    for k, s in SUPPLIERS.items():

        st = load_json(s["status"])
        stt, p = display_status(k, st, s["file"])

        if "🟢" in stt:
            icon = "🟢"
        elif "🟡" in stt:
            icon = "🟡"
        elif "⛔" in stt:
            icon = "⛔"
        elif "🔴" in stt:
            icon = "🔴"
        else:
            icon = "⚪"

        t += f"{s['name']} {icon}\n"
        t += "\n"

    return t

def kb_dashboard():
    rows = []
    blink_state = int(time.time() * 2) % 2

    for k, s in SUPPLIERS.items():
        st = load_json(s["status"])
        stt, p = display_status(k, st, s["file"])

        if "🟡" in stt:
            icon = "🟡" if blink_state == 0 else "⚪"
        elif "ГОТОВО" in stt:
            icon = "🟢"
        elif "ОТМЕНЕНО" in stt:
            icon = "⛔"
        elif "ОШИБКА" in stt:
            icon = "🔴"
        else:
            icon = "⚪"

        btn = f"{s['name']} {icon}"

        rows.append([
            InlineKeyboardButton(
                text=btn,
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
        reply_markup=kb_dashboard(),
        parse_mode="HTML"
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

            await asyncio.sleep(3)

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
            reply_markup=kb_dashboard(),
            parse_mode="HTML"
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
        import asyncio
    
        stt, p = display_status(key, st, s["file"])
    
        user = st.get("user", "-")

        from datetime import timedelta

        run_time = st.get("time", "-")
        
        try:
            dt = datetime.strptime(run_time, "%Y-%m-%d %H:%M:%S")
        
            # +3 часа
            dt = dt + timedelta(hours=3)
        
            # формат без секунд
            run_time = dt.strftime("%d.%m.%Y %H:%M")
    

        except Exception as e:
            if "message is not modified" not in str(e):
                print("EDIT ERROR:", e)

    
        text = f"📦 <b>{s['name']}</b>\n\n"
        text += f" {stt}\n"
        text += f"👤 <b>Пользователь:</b> {user}\n"
        text += f"🕒 <b>Запуск:</b> {run_time}\n\n"
        
        if st.get("running") or key in RUNNING_PROCESSES:
            text += "\n⏳ Идёт обработка..."
        
        
        elif stt == "🟢 ГОТОВО":

            if os.path.exists(s["file"]):
        
                size_mb = round(
                    os.path.getsize(s["file"]) / 1024 / 1024,
                    2
                )
        
                mtime = os.path.getmtime(s["file"])
        
                dt = datetime.fromtimestamp(mtime)
        
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
            "🚀 Запуск...",
            reply_markup=kb_supplier(key, True)
        )

        st = load_json(s["status"])
            
        st["running"] = True
        st["success"] = False
        st["canceled"] = False
        st["progress"] = 0
            
        with open(s["status"], "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=2)
        
        try:
            # запускаем парсер
            code = await run_parser(key, call.from_user.full_name)
            
            st = load_json(s["status"]) or {}
            
            st["running"] = False
            
            if code == "canceled":
                st["canceled"] = True
                st["success"] = False
            
            elif code == 0:
                st["canceled"] = False
                st["progress"] = 100
                st["success"] = True
            
            else:
                st["canceled"] = False
                st["progress"] = 0
                st["success"] = False
 
            with open(s["status"], "w", encoding="utf-8") as f:
                json.dump(st, f, ensure_ascii=False, indent=2)
            
            if os.path.exists(s["lock"]):
                try:
                    os.remove(s["lock"])
                except:
                    pass
            
            await call.message.edit_text(
                "✅ ГОТОВО",
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

        st = load_json(s["status"])
        
        st["running"] = False
        st["progress"] = 0
        st["canceled"] = True
        
        with open(s["status"], "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=2)

        await safe_edit(call, "⛔ ОТМЕНЕНО", kb_supplier(key, False))
        return

async def dashboard_updater():

    while True:

        for chat_id, msg_id in list(DASHBOARD_MESSAGES.items()):
            # НЕ ОБНОВЛЯЕМ, если пользователь сейчас в карточке
            if chat_id in DASHBOARD_OPENED:
                continue
            
            try:

                await safe_edit_message(
                    chat_id=chat_id,
                    message_id=msg_id,
                    text=dashboard_text(),
                    reply_markup=kb_dashboard(),
                    parse_mode="HTML"
                )
                
            except Exception:
                pass

        await asyncio.sleep(1.5)
        

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
