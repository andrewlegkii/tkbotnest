import telebot
from telebot import types
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os
import re

TOKEN = '8045807732:AAEfIFg7FbFVvYUcAbtLwonTjMs1agIIV7g'
bot = telebot.TeleBot(TOKEN)

EXCEL_FILE = 'data.xlsx'
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Документы'
    ws.append(['№', 'Дата/время', 'ТК', 'Номер документа', 'Фото путь'])
    wb.save(EXCEL_FILE)

user_states = {}

tc_list = ["ГТЕ", "МОНОПОЛИЯ", "ОБОЗ", "Л7", "ТТ", "СИЯНИЕ", "ВОЛК", "ОЛИМП"]

def send_tc_selection(chat_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for tc in tc_list:
        markup.add(types.KeyboardButton(tc))
    markup.add(types.KeyboardButton("ℹ Информация"))
    bot.send_message(chat_id, "Выберите ТК из списка или введите вручную:", reply_markup=markup)

@bot.message_handler(commands=['start'])
def start(message):
    user_states[message.chat.id] = {}
    send_tc_selection(message.chat.id)

@bot.message_handler(func=lambda message: message.text == "ℹ Информация")
def info(message):
    info_text = (
        "ℹ️ Инструкция для водителя\n\n"
        "📄 Перед загрузкой накладной на поддоны:\n\n"
        "Передайте 2 экземпляра Торг-12 на поддоны в окно приёмки документов на РЦ вместе с основными документами на груз.\n"
        "⚠️ Нельзя разделять документы на груз и поддоны — накладная на поддоны входит в комплект товарно-сопроводительной документации при поставке на РЦ Тандера.\n\n"
        "Проверьте, чтобы на Торг-12 и ТрН стояли:\n\n"
        "Подпись сотрудника РЦ\n\n"
        "Печать Тандера\n\n"
        "📸 Только после этого прикрепляйте фото документа через бот.\n\n"
        "📧 Контакты: vozvr_podd@magnit.ru\n"
        "В копию ставте адрес email: tatyana.gorlevich@nestle.ru"
    )
    bot.send_message(message.chat.id, info_text)

@bot.message_handler(func=lambda message: message.text == "🔙 Назад")
def go_back(message):
    user_states[message.chat.id] = {}
    send_tc_selection(message.chat.id)

@bot.message_handler(func=lambda message: message.text == "📎 Отправить ещё скан")
def send_another_scan(message):
    user_states[message.chat.id] = {}
    send_tc_selection(message.chat.id)

@bot.message_handler(func=lambda message: True and message.text not in ["ℹ Информация", "🔙 Назад", "📎 Отправить ещё скан"])
def handle_text(message):
    state = user_states.get(message.chat.id, {})
    if 'tc' not in state:
        user_states[message.chat.id] = {'tc': message.text}
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(types.KeyboardButton("🔙 Назад"))
        bot.send_message(message.chat.id, f"Вы выбрали ТК: {message.text}\nТеперь введите номер документа (например, R101...)",
                         reply_markup=markup)
    elif 'doc' not in state:
        doc_number = message.text.strip()
        # Измененная проверка: игнорируем регистр буквы R
        if not re.match(r"^[Rr]101\d+$", doc_number):
            bot.send_message(message.chat.id, "❌ Номер документа должен начинаться с R101 или r101 и содержать только цифры после.")
            return
        user_states[message.chat.id]['doc'] = doc_number
        bot.send_message(
            message.chat.id,
            "Теперь отправьте фото накладной 📷\n‼️ Убедитесь, что на Торг-12 и ТН есть подпись сотрудника РЦ и печать Тандера."
        )

@bot.message_handler(content_types=['photo'])
def handle_photo(message):
    state = user_states.get(message.chat.id)
    if not state or 'doc' not in state or 'tc' not in state:
        bot.send_message(message.chat.id, "Сначала выберите ТК и введите номер документа.")
        return

    doc_number = state['doc']
    tc_name = state['tc']

    # Скачиваем фото
    file_info = bot.get_file(message.photo[-1].file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    folder = "photos"
    os.makedirs(folder, exist_ok=True)
    file_path = os.path.join(folder, f"{doc_number}.jpg")

    with open(file_path, 'wb') as f:
        f.write(downloaded_file)

    # Запись в Excel
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    row_number = ws.max_row
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ws.append([row_number, now, tc_name, doc_number, file_path])
    wb.save(EXCEL_FILE)

    # Ответ пользователю
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add(types.KeyboardButton("📎 Отправить ещё скан"))
    bot.send_message(
        message.chat.id,
        f"✅ Документ {doc_number} загружен и сохранён.\nСпасибо!",
        reply_markup=markup
    )

    user_states.pop(message.chat.id, None)

# Запуск
if __name__ == '__main__':
    print("Bot is running...")
    bot.infinity_polling(timeout=60, long_polling_timeout=60)
