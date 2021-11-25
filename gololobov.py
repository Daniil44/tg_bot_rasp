import telebot
import openpyxl
from dotenv import load_dotenv
from random import randint
from pathlib import Path
import os
from flask import Flask, request
load_dotenv()
TOKEN = os.getenv("TOKEN")
book = openpyxl.open("o.xltm", read_only=True)
sheet = book.active
substring = ['.', '[']
wish = ['Доброго дня!', 'Удачи! ', 'У тебя получится!', 'Не сдавайся!', 'Пятерок!']
bot = telebot.TeleBot('2013032423:AAFEzGyOuHQo0UY8KQEh8AND91A8m78AziQ')

def check(word):
    if word == '-':
        return('-')
    else:    
        if substring[0] in word:
            return('.'.join(word.split('.')[:1]))
        else:
            if substring[1] in word:
                return('.'.join(word.split('[')[:1]))
            else:
                return(word)

count = 13
mo_f_l = check(sheet['C'+str(count)].value)
count1 = count + 2
mo_s_l = check(sheet['C'+str(count1)].value)
count2 = count1 + 2
mo_t_l = check(sheet['C'+str(count2)].value)
#Вторник
count = 29
tu_f_l = check(sheet['C'+str(count)].value)
count1 = count + 2
tu_s_l = check(sheet['C'+str(count1)].value)
count2 = count1 + 2
tu_t_l = check(sheet['C'+str(count2)].value)
#Среда
count = 45
we_f_l = check(sheet['C'+str(count)].value)
count1 = count + 2
we_s_l = check(sheet['C'+str(count1)].value)
count2 = count1 + 2
we_t_l = check(sheet['C'+str(count2)].value)
#Четверг
count = 61
th_f_l = check(sheet['C'+str(count)].value)
count1 = count + 2
th_s_l = check(sheet['C'+str(count1)].value)
count2 = count1 + 2
th_t_l = check(sheet['C'+str(count2)].value)
#Пятница
count = 77
fr_f_l = check(sheet['C'+str(count)].value)
count1 = count + 2
fr_s_l = check(sheet['C'+str(count1)].value)
count2 = count1 + 2
fr_t_l = check(sheet['C'+str(count2)].value)
wish = ['Доброго дня!', 'Удачи! ', 'У тебя получится!', 'Не сдавайся!', 'Пятерок!']

@bot.message_handler(commands=['help'])
def help(message):
    bot.send_message(message.chat.id, 'Команды здесь ТОЛЬКО на английском языке и со слешем. Если хочешь узнать уроки введи первые 2 буквы дня недели. /mo - понедельник /tu - вторник /we - среда /th - четверг /fr - пятница. также бот  может отправлять пожелания. введи /wish')
@bot.message_handler(commands=['wish'])
def wish2(message):
    a = randint(0, 4)
    bot.send_message(message.chat.id, wish[a])
@bot.message_handler(commands=['mo'])
def timetable(message):
    bot.send_message(message.chat.id, mo_f_l)
    bot.send_message(message.chat.id, mo_s_l)
    bot.send_message(message.chat.id, mo_t_l)
@bot.message_handler(commands=['tu'])
def timetable2(message):
    bot.send_message(message.chat.id, tu_f_l)
    bot.send_message(message.chat.id, tu_s_l)
    bot.send_message(message.chat.id, tu_t_l)
@bot.message_handler(commands=['we'])
def timetable3(message):
    bot.send_message(message.chat.id, we_f_l)
    bot.send_message(message.chat.id, we_s_l)
    bot.send_message(message.chat.id, we_t_l)
@bot.message_handler(commands=['th'])
def timetable4(message):
    bot.send_message(message.chat.id, th_f_l)
    bot.send_message(message.chat.id, th_s_l)
    bot.send_message(message.chat.id, th_t_l)
@bot.message_handler(commands=['fr'])
def timetable5(message):
    bot.send_message(message.chat.id, fr_f_l)
    bot.send_message(message.chat.id, fr_s_l)
    bot.send_message(message.chat.id, fr_t_l)
@bot.message_handler(commands=['start'])
def start_command(message):
    bot.send_message(message.chat.id, 'Привет! В этом боте есть расписание уроков пятого класса монтессори! Введи: /help для подробной информации.')
#bot.polling()
@server.route("/")
def webhook():
    bot.remove_webhook()
    bot.set_webhook(url="https://dashboard.heroku.com/apps/montessory-tg-bot") # этот url нужно заменить на url вашего Хероку приложения
    return "?", 200
server.run(host="0.0.0.0", port=os.environ.get('PORT', 80))
