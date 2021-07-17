from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Updater, CommandHandler, CallbackQueryHandler, CallbackContext, MessageHandler, Filters, \
    ConversationHandler
from Token_BOT_SQL import Token
import logging
from requests import post
from Store_File_and_Send import store_file, save_data_to_another_table

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO
)

logger = logging.getLogger(__name__)

GET, DATA, PHONE_NUMBER, ISSUE, WHICH_ONE, CORRECT_INFO, CUSTOMER_NAME = range(7)


def facts_to_str(user_data) -> str:
    facts = [f'{key} - {value}' for key, value in user_data.items()]
    return "\n".join(facts).join(['\n', '\n'])


def start(update: Update, context: CallbackContext) -> int:
    keyboard = [
        [
            InlineKeyboardButton('Приём.', callback_data='Get_Device'),
            InlineKeyboardButton('Выдача.', callback_data='Return_Device')
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    update.message.reply_text('Что хочешь сделать?', reply_markup=reply_markup)
    return GET


def customer_name(update: Update, context: CallbackContext) -> int:
    context.user_data['Имя клиента'] = update.message.text
    update.message.reply_text('Номер телефона?')
    return PHONE_NUMBER


def phone_number(update: Update, context: CallbackContext) -> int:
    context.user_data['Номер телефона'] = update.message.text
    update.message.reply_text("Что случилось?")
    return ISSUE


def mistake_was_made(update: Update, context: CallbackContext) -> int:
    keyboard = [
            ['Имя клиента', 'Номер телефона', 'Поломка']
    ]
    markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True)
    update.message.reply_text(text='Где ошибка?', reply_markup=markup)
    return WHICH_ONE


def what_happened(update: Update, context: CallbackContext) -> None:
    context.user_data['Поломка'] = update.message.text
    update.message.reply_text(f'Убедись, что написал всё верно. \n{facts_to_str(context.user_data)}'
                              f'Если да, жми /save. \nЕсли ошибся, то /mistake.')
    return ConversationHandler.END


def save_order(update: Update, context: CallbackContext) -> None:
    text = store_file(context.user_data)
    update.message.reply_text(text=text[:11])
    message = f'https://api.telegram.org/bot{Token}/sendDocument?chat_id={update.effective_chat.id}'
    post(message, files={'document': open('Excel_And_Pdf/PDF/' + text[11:], 'rb')})
    context.user_data.clear()


def button(update: Update, context: CallbackContext) -> int:
    query = update.callback_query
    if query['data'] == 'Get_Device':
        query.edit_message_text('Как зовут?')
    if query['data'] == 'Return_Device':
        query.edit_message_text('Номер ремонта:')
    return GET


def category_mistake(update: Update, context: CallbackContext) -> int:
    global category
    category = update.message.text
    update.message.reply_text(f'Понял. Ошибка в {category}. Можешь исправить.')
    return CORRECT_INFO


def search_repair(update: Update, context: CallbackContext) -> str:
    if save_data_to_another_table(update.message.text):
        update.message.reply_text("Спасибо. Ремонт был перемещён в завершённые.")
    return ConversationHandler.END


def correction_info(update: Update, context: CallbackContext) -> int:
    context.user_data[category] = update.message.text
    update.message.reply_text(f'Убедись, что написал всё верно. \n{facts_to_str(context.user_data)}'
                              f'Если да, жми /save. \nЕсли ошибся, то  жми /mistake.')
    return ConversationHandler.END


def main(user_limit: list) -> None:
    updater = Updater(Token)
    add = updater.dispatcher.add_handler
    gather_info_customer = ConversationHandler(
        entry_points=[CommandHandler('start', start, Filters.user(user_limit))],
        states={
            GET: [
                MessageHandler(Filters.regex(r'\d\d.\d\d.\d\d.\d\d'), search_repair),
                MessageHandler(Filters.text, customer_name)
            ],
            DATA: [
                add(CallbackQueryHandler(button))
            ],
            PHONE_NUMBER: [
                MessageHandler(Filters.text, phone_number)
            ],
            ISSUE: [
                MessageHandler(Filters.text, what_happened)
            ],
        },
        fallbacks=[CommandHandler('start', start, Filters.user(user_limit))]
    )
    mistake_handler = ConversationHandler(
        entry_points=[CommandHandler('mistake', mistake_was_made, Filters.user(user_limit))],
        states={
            WHICH_ONE: [
                MessageHandler(Filters.regex('^(Имя клиента|Номер телефона|Поломка)$'), category_mistake)
            ],
            CORRECT_INFO: [
                MessageHandler(Filters.text & ~Filters.command, correction_info)
            ]
        },
        fallbacks=[CommandHandler('mistake', mistake_was_made, Filters.user(user_limit))]
    )
    add(CommandHandler('save', save_order, Filters.user(user_limit)))
    add(gather_info_customer)
    add(mistake_handler)
    updater.start_polling()
    updater.idle()


if __name__ == '__main__':
    user_list = [Your_List_Chat_ID]
    main(user_list)
