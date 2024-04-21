import telebot
from telebot import types
from FromSheet import FromSheet

bot = telebot.TeleBot('7198822480:AAGcs_xzxNbcJ7vQJLIFPw2zEhnKMx3fhZA')

allowed = ["1015008397"]

filePath = "C:/Users/Cectus/Desktop/"
fileName = "file"
fileExt = ".xlsx"

greetingsStr = ["Добрый день, вы можете начинать работу!",
                "Вас нет в вайт-листе, есди вы учитель, то свяжитесь с нами @cectus1"]
greetingsButtons = ["Изменить расписание", "Получить таблицу", "Внести актуальную таблицу"]
actionSelectorStr = ["Введите день недели! (1-7)", "Скиньте файл"]
getTableStr = "Файл успешно загружен!"

@bot.message_handler(commands=["start"])
def greetings(message):
    if (str(message.from_user.id) in allowed):
        bot.send_message(message.chat.id, greetingsStr[0], reply_markup=keyboard(greetingsButtons))
        bot.register_next_step_handler(message, actionSelector)
    else:
        bot.send_message(message.chat.id, greetingsStr[1])

def actionSelector(message):
    if (message.text == greetingsButtons[0]):  #Изменить распсание
        bot.send_message(message.chat.id, actionSelectorStr[0], reply_markup=types.ReplyKeyboardRemove())
        bot.register_next_step_handler(message, scheduleShow)
    elif (message.text == greetingsButtons[1]):  #Получить таблицу
        file = open(filePath + fileName + fileExt, 'rb')
        bot.send_document(message.chat.id, file, reply_markup=keyboard(greetingsButtons))
        bot.register_next_step_handler(message, actionSelector)
    elif (message.text == greetingsButtons[2]):  #Внести актуальную таблицу
        bot.send_message(message.chat.id, actionSelectorStr[1], reply_markup=types.ReplyKeyboardRemove())
        bot.register_next_step_handler(message, getTable)

def scheduleShow(message):
    dayClassLesson = message.text.split(" ")
    lesson = FromSheet(int(dayClassLesson[0]), dayClassLesson[1].lower(), int(dayClassLesson[2]))
    bot.send_message(message.chat.id, "1.Урок: "+lesson.get_everything()[1]+"\n2.Учитель: "+lesson.get_everything()[0]+"\n\nВведите цифру и текст для изменения поля.")
    bot.register_next_step_handler(message, getTable, lesson)

def scheduleEdit(message, lesson):
    lesson.


@bot.message_handler(content_types=['document'])
def getTable(message):
    if (str(message.from_user.id) in allowed):
        docInfo = bot.get_file(message.document.file_id)
        doc = bot.download_file(file_path=docInfo.file_path)
        with open(filePath + fileName + fileExt, 'wb') as new_file:
            new_file.write(doc)
        bot.send_message(message.chat.id, getTableStr, reply_markup=keyboard(greetingsButtons))
        new_file.close()
        bot.register_next_step_handler(message, actionSelector)


def keyboard(button=[]):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for i in range(len(button)):
        markup.add(types.KeyboardButton(button[i]))
    return markup


bot.infinity_polling()
