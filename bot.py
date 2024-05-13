import telebot
import os
from telebot import types
from ExcelEditor import ExcelEditor
from FromSheet import FromSheet
import dotenv
import config

dotenv.load_dotenv()

API_TOKEN = os.getenv('TELEGRAM_BOT_API_KEY')

#bot = telebot.TeleBot('7198822480:AAGcs_xzxNbcJ7vQJLIFPw2zEhnKMx3fhZA')
bot = telebot.TeleBot(API_TOKEN)

allowed = ["1015008397", "1585747030", "1027925215"]

filePath = ""
fileName = "file"
fileExt = ".xlsx"

greetingsStr = ["Добрый день, вы можете начинать работу!",
                "Вас нет в вайт-листе, есди вы учитель, то свяжитесь с нами @cectus1"]
greetingsButtons = ["Изменить расписание", "Получить таблицу", "Внести актуальную таблицу"]
actionSelectorStr = ["Введите данные в формате ДеньНедели КлассБуква НомерУрока. Пример '3 10Т 3'", "Скиньте файл"]
getTableStr = "Файл успешно загружен!"


@bot.message_handler(commands=["start"])
def greetings(message):
    # print(message.chat.id)
    if (str(message.from_user.id) in config.allowed):
        bot.send_message(message.chat.id, config.greetingsStr[0], reply_markup=keyboard(greetingsButtons))
        bot.register_next_step_handler(message, actionSelector)
    else:
        bot.send_message(message.chat.id, config.greetingsStr[1])

@bot.message_handler(func=lambda message: True)
def actionSelector(message):
    if (message.text == greetingsButtons[0]):  # Изменить распсание
        bot.send_message(message.chat.id, actionSelectorStr[0], reply_markup=types.ReplyKeyboardRemove())
        bot.register_next_step_handler(message, scheduleShow)
    elif (message.text == greetingsButtons[1]):  # Получить таблицу
        file = open(config.filePath + config.fileName + config.fileExt, 'rb')
        bot.send_document(message.chat.id, file, reply_markup=keyboard(greetingsButtons))
        file.close()
        bot.register_next_step_handler(message, actionSelector)
    elif (message.text == greetingsButtons[2]):  # Внести актуальную таблицу
        bot.send_message(message.chat.id, config.actionSelectorStr[1], reply_markup=types.ReplyKeyboardRemove())
        bot.register_next_step_handler(message, config.getTable)


def scheduleShow(message):
    dayClassLesson = message.text.split(" ")
    lesson = FromSheet(int(dayClassLesson[0]) - 1, dayClassLesson[1].lower(), int(dayClassLesson[2]) + 1)
    bot.send_message(message.chat.id,
                     "1.Урок: " + lesson.get_everything()[1] + "\n2.Учитель: " + lesson.get_everything()[
                         0] + "\n\nВведите цифру и текст для изменения поля.")
    # print(lesson.get_everything()[5], lesson.get_everything()[6])
    bot.register_next_step_handler(message,
                                   scheduleEdit,
                                   int(dayClassLesson[0]),
                                   dayClassLesson[1],
                                   int(dayClassLesson[2]))


def scheduleEdit(message, day, clas, lesson_number):
    editor = ExcelEditor(config.filePath + config.fileName + config.fileExt)
    editor.make_temp_file()

    if message.text.split(' ')[0] == '1':
        row = str((day - 1) * 8 + 2 + lesson_number)
        editor.edit_cell(config.class_to_excel_cell[clas] + row, message.text.split(' ')[1], 0)
    elif message.text.split(' ')[0] == '2':
        editor.edit_cell(message.text.split(' ')[1], clas, 1)

    bot.send_message(message.chat.id, config.greetingsStr[0], reply_markup=keyboard(greetingsButtons))
    bot.register_next_step_handler(message, actionSelector)



@bot.message_handler(content_types=['document'])
def getTable(message):
    if (str(message.from_user.id) in config.allowed):
        docInfo = bot.get_file(message.document.file_id)
        doc = bot.download_file(file_path=docInfo.file_path)
        with open(config.filePath + config.fileName + config.fileExt, 'wb') as new_file:
            new_file.write(doc)
        bot.send_message(message.chat.id, config.getTableStr, reply_markup=keyboard(greetingsButtons))
        new_file.close()
        bot.register_next_step_handler(message, actionSelector)


def keyboard(button=[]):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for i in range(len(button)):
        markup.add(types.KeyboardButton(button[i]))
    return markup


bot.infinity_polling()
