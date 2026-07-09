# -*- coding: utf-8 -*-
"""
Telegram-бот "Дашборд производственных показателей"
Библиотека: pyTelegramBotAPI (telebot) — самая простая для деплоя без сервера.
"""

import telebot
from telebot import types
import os

# ====== ТОКЕН ======
# Получить у @BotFather -> /newbot -> вставить сюда или задать переменную окружения BOT_TOKEN
TOKEN = os.environ.get("BOT_TOKEN", "ВСТАВЬ_СЮДА_ТОКЕН")

bot = telebot.TeleBot(TOKEN)

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
          "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

# план / три вида факта
IND = {
    "plan": "План по ОП",
    "stat": "Факт по статистике",
    "mark": "Факт марк. замеру",
    "uchet": "Факт принято к учету",
}

# Данные из сводной таблицы (Янв-Апр — реальные, дальше нули до обновления)
CATS = {
    "Горная масса": {
        "plan":  [73, 131, 78, 152, 0, 0, 0, 0, 0, 0, 0, 0],
        "stat":  [64, 88, 99, 115, 0, 0, 0, 0, 0, 0, 0, 0],
        "mark":  [63, 89, 99, 115, 0, 0, 0, 0, 0, 0, 0, 0],
        "uchet": [64, 89, 99, 115, 0, 0, 0, 0, 0, 0, 0, 0],
    },
    "Добыча": {
        "plan":  [0, 0, 0, 9, 0, 0, 0, 0, 0, 0, 0, 0],
        "stat":  [0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0],
        "mark":  [0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0],
        "uchet": [0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 0, 0],
    },
    "Вскрыша и навалы": {
        "plan":  [52, 50, 27, 77, 0, 0, 0, 0, 0, 0, 0, 0],
        "stat":  [24, 55, 37, 65, 0, 0, 0, 0, 0, 0, 0, 0],
        "mark":  [23, 56, 37, 65, 0, 0, 0, 0, 0, 0, 0, 0],
        "uchet": [24, 56, 37, 65, 0, 0, 0, 0, 0, 0, 0, 0],
    },
    "Прочие": {
        "plan":  [21, 81, 51, 66, 0, 0, 0, 0, 0, 0, 0, 0],
        "stat":  [40, 33, 62, 30, 0, 0, 0, 0, 0, 0, 0, 0],
        "mark":  [40, 33, 62, 30, 0, 0, 0, 0, 0, 0, 0, 0],
        "uchet": [40, 33, 62, 30, 0, 0, 0, 0, 0, 0, 0, 0],
    },
    "Переработка (промывка) песков": {
        "plan":  [11, 10, 0, 9, 0, 0, 0, 0, 0, 0, 0, 0],
        "stat":  [5, 10, 0, 10, 0, 0, 0, 0, 0, 0, 0, 0],
        "mark":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        "uchet": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    },
}

# для "Переработка" нет марк.замера и учёта — фолбэк на статистику
NO_FULL = {"Переработка (промывка) песков"}

# ====== состояние пользователя (простое, в памяти) ======
user_state = {}  # chat_id -> {"cat": .., "ind": .., "months": ..}


def main_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("📊 Выбрать показатель")
    return kb


def cats_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    for c in CATS:
        kb.add(c)
    kb.add("⬅️ Назад")
    return kb


def ind_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    kb.add("Факт по статистике")
    kb.add("Факт марк. замеру")
    kb.add("Факт принято к учету")
    kb.add("⬅️ Назад")
    return kb


def period_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    kb.add("Январь–Апрель (есть данные)")
    kb.add("Весь год")
    kb.add("⬅️ Назад")
    return kb


IND_LABEL_TO_KEY = {
    "Факт по статистике": "stat",
    "Факт марк. замеру": "mark",
    "Факт принято к учету": "uchet",
}


@bot.message_handler(commands=["start"])
def start(msg):
    user_state[msg.chat.id] = {}
    bot.send_message(
        msg.chat.id,
        "Дашборд производственных показателей.\nВыбери показатель для отчёта.",
        reply_markup=main_menu(),
    )


@bot.message_handler(func=lambda m: m.text == "📊 Выбрать показатель")
def choose_cat(msg):
    bot.send_message(msg.chat.id, "Выбери направление:", reply_markup=cats_menu())


@bot.message_handler(func=lambda m: m.text in CATS)
def choose_ind(msg):
    user_state.setdefault(msg.chat.id, {})["cat"] = msg.text
    bot.send_message(msg.chat.id, "С чем сравнить план?", reply_markup=ind_menu())


@bot.message_handler(func=lambda m: m.text in IND_LABEL_TO_KEY)
def choose_period(msg):
    st = user_state.setdefault(msg.chat.id, {})
    if "cat" not in st:
        bot.send_message(msg.chat.id, "Сначала выбери направление.", reply_markup=cats_menu())
        return
    st["ind"] = IND_LABEL_TO_KEY[msg.text]
    bot.send_message(msg.chat.id, "За какой период?", reply_markup=period_menu())


@bot.message_handler(func=lambda m: m.text in ("Январь–Апрель (есть данные)", "Весь год"))
def show_report(msg):
    st = user_state.get(msg.chat.id, {})
    if "cat" not in st or "ind" not in st:
        bot.send_message(msg.chat.id, "Начни заново: /start")
        return

    cat = st["cat"]
    ind_key = st["ind"]
    months_range = range(4) if "Апрель" in msg.text else range(12)

    data = CATS[cat]
    ind_key_used = ind_key
    fallback_note = ""
    if cat in NO_FULL and ind_key in ("mark", "uchet"):
        ind_key_used = "stat"
        fallback_note = "\n⚠️ Для этого направления нет марк.замера/учёта — показана статистика."

    plan = data["plan"]
    fact = data[ind_key_used]

    lines = [f"📊 {cat}", f"Сравнение: План по ОП vs {IND[ind_key_used]}", ""]
    total_plan = total_fact = 0
    has_rows = False
    for i in months_range:
        p, f = plan[i], fact[i]
        if p == 0 and f == 0:
            continue
        has_rows = True
        total_plan += p
        total_fact += f
        pct = round(f / p * 100, 1) if p > 0 else None
        pct_str = f"{pct}%" if pct is not None else "—"
        mark = "✅" if pct and pct >= 100 else ("⚠️" if pct else "")
        lines.append(f"{MONTHS[i]}: план {p} / факт {f}  ({pct_str}) {mark}")

    if not has_rows:
        lines.append("Нет данных за выбранный период.")
    else:
        total_pct = round(total_fact / total_plan * 100, 1) if total_plan > 0 else None
        lines.append("")
        lines.append(f"Итого: план {total_plan} / факт {total_fact} ({total_pct}%)" if total_pct is not None else f"Итого: план {total_plan} / факт {total_fact}")

    lines.append(fallback_note) if fallback_note else None

    bot.send_message(msg.chat.id, "\n".join(lines), reply_markup=cats_menu())


@bot.message_handler(func=lambda m: m.text == "⬅️ Назад")
def go_back(msg):
    start(msg)


@bot.message_handler(func=lambda m: True)
def fallback(msg):
    bot.send_message(msg.chat.id, "Не понял. Нажми /start.", reply_markup=main_menu())


if __name__ == "__main__":
    print("Бот запущен...")
    bot.infinity_polling()
