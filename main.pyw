from aiogram import Bot, Dispatcher, executor, types
import config as cfg
import markups as nav
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.contrib.fsm_storage.memory import MemoryStorage
import requests
import os
from dotenv import load_dotenv
import json
from datetime import timedelta, datetime
import openpyxl
from func_uch import *
import asyncio

load_dotenv()

bot = Bot(token=cfg.TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())



class States(StatesGroup):
    step_1 = State()
    step_2 = State()
    step_3 = State()
    step_4 = State()
    step_5 = State()
    step_6 = State()
    step_7 = State()
    step_8 = State()
    step_9 = State()
    step_10 = State()

async def my_func():
    while True:
        if datetime.now().hour == 0 and datetime.now().minute == 10:
            func_b()
            await asyncio.sleep(3000)  
        await asyncio.sleep(5) 

async def on_startup(dp):
    asyncio.create_task(my_func())



@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    if message.from_user.id in [1193989520, 424915104, 6400135693, 336194148]: #  список id пользователей, которым разрешено запускать бота
        await bot.send_message(message.from_user.id, 'Меню', reply_markup=nav.keyb_1)

@dp.callback_query_handler(lambda c: True)
async def call_back_q(callback: types.CallbackQuery):   
    if callback.data == 'btn_1':        
        try:
            await bot.delete_message(callback.message.chat.id, callback.message.message_id)
        except:
            pass  
        z = await bot.send_message(callback.from_user.id, 'Загрузка⏳')
        act = []
        deact = []
        path = "C:/Users/Evgen/YandexDisk-promebelmarket/Маркетплейсы/API Аналитика/Озон/Общая модель.xlsx"
        wb = openpyxl.load_workbook(path)  
        sheet = wb['Участвуют и доступные для участ']  
        for row in sheet.values:
            if str(row[14]) == '-1':
                deact.append({
            "action_id": row[2],
            "product_ids": [row[5]]
        })
                
        for row in sheet.values:
            if str(row[14]) == '1':
                act.append({
            "action_id": row[2],
            "products": [
                {
                    "action_price": row[8],
                    "product_id": row[5],
                    "stock": 10
                }
            ]
        })        
        url = 'https://api-seller.ozon.ru/v1/actions/products/deactivate' #Удалить товары из акции
        headers = {'Client-Id':os.getenv('client_id'), 'Api-Key':os.getenv('client_secret'),}
        for row in deact:
            r = requests.post(url, headers = headers, json=row)
            r.json()

        url = 'https://api-seller.ozon.ru/v1/actions/products/activate' #Добавить товар в акцию
        headers = {'Client-Id':os.getenv('client_id'), 'Api-Key':os.getenv('client_secret'),}
        for row in act:
            r = requests.post(url, headers = headers, json=row)
            r.json()    

        i = 3      
        for row in sheet.values:
            if row[0] == datetime.now().strftime('%d.%m.%Y'):
                sheet.delete_rows(i, amount=10000)
                break 
            i += 1
        wb.save("C:/Users/Evgen/YandexDisk-promebelmarket/Маркетплейсы/API Аналитика/Озон/Общая модель.xlsx")

        path = "C:/Users/Evgen/YandexDisk-promebelmarket/Маркетплейсы/API Аналитика/Озон/Общая модель.xlsx"
        wb = openpyxl.load_workbook(path)  
        sheet = wb['Содержательная бизнес логика'] 
        i = 3        
        for row in sheet.values:
            if row[0] == datetime.now().strftime('%d.%m.%Y'):
                sheet.delete_rows(i, amount=10000)
                break 
            i += 1
        wb.save("C:/Users/Evgen/YandexDisk-promebelmarket/Маркетплейсы/API Аналитика/Озон/Общая модель.xlsx")  
        func_b()       
        try:
            await bot.delete_message(callback.message.chat.id, z.message_id)
        except:
            pass  
        await bot.send_message(callback.from_user.id, 'Процесс закончен', reply_markup=nav.keyb_1)



@dp.message_handler(commands=['1'])
async def start(message: types.Message):
    if message.from_user.id in [1193989520]:
        z = await bot.send_message(message.from_user.id, 'Началось формирование файла⏳') 
        func_b()
        try:
            await bot.delete_message(message.chat.id, z.message_id)
        except:
            pass  
        await bot.send_message(message.from_user.id, 'Формирование файла закончено')
 

if __name__ == "__main__":  
    try:  
        executor.start_polling(dp, skip_updates=True, on_startup=on_startup)
    except:
        file = open('Ошибки.txt', 'a')
        file.write(f"{datetime.now().strftime('%d.%m.%Y %H:%M:%S')} Ошибка, вероятно нет подключения к интернету\n")  