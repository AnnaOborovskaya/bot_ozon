from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton

keyb_1 = InlineKeyboardMarkup(row_width=1)
btn_1 = InlineKeyboardButton(text='Добавить/Удалить акции', callback_data='btn_1')
keyb_1.add(btn_1)
