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
admin_id = 360300829  # <- Вставь сюда свой настоящий Telegram ID
DATA_PASSWORD = "2695"  # <- Задай свой пароль для команды /data

tc_list = ["ГТЕ", "МОНОПОЛИЯ", "ОБОЗ", "Л7", "ТТ", "СИЯНИЕ", "ВОЛК", "ОЛИМП"]

def send_tc_selection(chat_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for tc in tc_list:
        markup.add(types.KeyboardButton(tc))
    markup.add(types.KeyboardButton("ℹ Информация"))
    bot.send_message(chat_id, "Выберите ТК кнопкой или введите вручную:", reply_markup=markup)

@bot.message_handler(commands=['start'])
def start(message):
    user_states[message.chat.id] = {}
    send_tc_selection(message.chat.id)

@bot.message_handler(commands=['data'])
def send_data_file(message):
    # Проверяем пароль
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        bot.reply_to(message, "❗ Пожалуйста, укажи пароль после команды, например:\n/data пароль123")
        return
    password = args[1].strip()
    if password != DATA_PASSWORD:
        bot.reply_to(message, "❌ Неверный пароль!")
        return
    try:
        with open(EXCEL_FILE, 'rb') as f:
            bot.send_document(message.chat.id, f)
    except Exception as e:
        bot.reply_to(message, f"Ошибка при отправке файла: {e}")

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
        "В копию ставьте адрес email: tatyana.gorlevich@nestle.ru"
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

@bot.message_handler(func=lambda message: message.text == "🆘 Помощь")
def help_start(message):
    user_states[message.chat.id] = {'help_mode': True}
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add(types.KeyboardButton("🔙 Назад"))
    bot.send_message(message.chat.id, "Напишите сообщение администратору. Чтобы выйти, нажмите '🔙 Назад'.", reply_markup=markup)

@bot.message_handler(func=lambda message: True)
def handle_text(message):
    state = user_states.get(message.chat.id, {})

    # Режим помощи — пересылаем админу
    if state.get('help_mode'):
        if message.text == "🔙 Назад":
            user_states[message.chat.id] = {}
            send_tc_selection(message.chat.id)
            return
        bot.send_message(admin_id, f"Сообщение от @{message.from_user.username or message.from_user.first_name} (ID {message.chat.id}):\n{message.text}")
        bot.send_message(message.chat.id, "✅ Ваше сообщение отправлено администратору. Ожидайте ответ.")
        return

    # Обработка выбора ТК
    if 'tc' not in state:
        user_states[message.chat.id] = {'tc': message.text}
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(types.KeyboardButton("🔙 Назад"))
        markup.add(types.KeyboardButton("🆘 Помощь"))
        bot.send_message(message.chat.id, f"Вы выбрали ТК: {message.text}\nТеперь введите номер документа (например, R101...)", reply_markup=markup)
        return

    # Обработка ввода номера документа
    if 'doc' not in state:
        doc_number = message.text.strip()
        if not re.match(r"^[Rr]101\d+$", doc_number):
            bot.send_message(message.chat.id, "❌ Номер документа должен начинаться с R101 или r101 и содержать только цифры после.")
            return
        user_states[message.chat.id]['doc'] = doc_number
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        markup.add(types.KeyboardButton("🔙 Назад"))
        markup.add(types.KeyboardButton("🆘 Помощь"))
        bot.send_message(
            message.chat.id,
            "Теперь отправьте фото накладной 📷\n‼️ Убедитесь, что на Торг-12 и ТН есть подпись сотрудника РЦ и печать Тандера.",
            reply_markup=markup
        )
        return

    bot.send_message(message.chat.id, "Пожалуйста, отправьте фото накладной или используйте кнопки.")

@bot.message_handler(content_types=['photo'])
def handle_photo(message):
    state = user_states.get(message.chat.id)
    if not state or 'doc' not in state or 'tc' not in state:
        bot.send_message(message.chat.id, "Сначала выберите ТК и введите номер документа.")
        return

    doc_number = state['doc']
    tc_name = state['tc']

    file_info = bot.get_file(message.photo[-1].file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    folder = "photos"
    os.makedirs(folder, exist_ok=True)
    file_path = os.path.join(folder, f"{doc_number}.jpg")

    with open(file_path, 'wb') as f:
        f.write(downloaded_file)

    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    row_number = ws.max_row
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ws.append([row_number, now, tc_name, doc_number, file_path])
    wb.save(EXCEL_FILE)

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add(types.KeyboardButton("📎 Отправить ещё скан"))
    bot.send_message(
        message.chat.id,
        f"✅ Документ {doc_number} загружен и сохранён.\nСпасибо!",
        reply_markup=markup
    )

    user_states.pop(message.chat.id, None)

@bot.message_handler(commands=['reply'])
def admin_reply(message):
    if message.chat.id != admin_id:
        return

    parts = message.text.split(' ', 2)
    if len(parts) < 3:
        bot.send_message(admin_id, "Использование: /reply <user_id> <текст ответа>")
        return
    user_id_str, answer = parts[1], parts[2]

    if not user_id_str.isdigit():
        bot.send_message(admin_id, "ID пользователя должен быть числом.")
        return
    user_id_int = int(user_id_str)

    try:
        bot.send_message(user_id_int, f"💬 Ответ администратора:\n{answer}")
        bot.send_message(admin_id, "Ответ отправлен.")
    except Exception as e:
        bot.send_message(admin_id, f"Ошибка при отправке сообщения: {e}")

if __name__ == '__main__':
    print("Bot is running...")
    bot.infinity_polling(timeout=60, long_polling_timeout=60)
