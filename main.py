import logging
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters, 
    ConversationHandler, ContextTypes
)
import sqlite3
import pandas as pd
from datetime import datetime
import os

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# Состояния разговора
STORE_NAME, PRODUCT_URL, DESCRIPTION, FULL_NAME, CONFIRMATION, EDIT_CHOICE = range(6)

# Список магазинов
STORES = [
    "🎃1688",
    "🔥Taobao", 
    "💎POISON",
    "❤️Pindoudou"
]

# Инициализация базы данных
def init_db():
    conn = sqlite3.connect('submissions.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            store_name TEXT NOT NULL,
            product_url TEXT NOT NULL,
            description TEXT,
            full_name TEXT NOT NULL,
            photo_filename TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            status TEXT DEFAULT 'Новая заявка',
            user_id INTEGER
        )
    ''')
    conn.commit()
    conn.close()

async def download_photo(photo_file, application_id):
    """Скачивание фото и сохранение на диск с именем по ID заявки"""
    try:
        os.makedirs('photos', exist_ok=True)
        photo = await photo_file.get_file()
        file_name = f"application_{application_id}.jpg"
        file_path = f'photos/{file_name}'
        await photo.download_to_drive(file_path)
        return file_name
    except Exception as e:
        logging.error(f"Ошибка скачивания фото: {e}")
        return None

async def export_to_excel():
    """Экспорт всех заявок в Excel файл"""
    try:
        conn = sqlite3.connect('submissions.db', check_same_thread=False)
        query = """
        SELECT 
            status as "Статус",
            store_name as "Магазин", 
            full_name as "ФИО клиента",
            product_url as "Ссылка на товар", 
            description as "Описание товара",
            photo_filename as "Фото",
            created_at as "Дата заявки",
            id as "ID заявки",
            user_id as "ID пользователя"
        FROM applications 
        ORDER BY created_at DESC
        """
        
        df = pd.read_sql_query(query, conn)
        conn.close()
        
        if df.empty:
            return False
            
        filename = "заявки_на_товары.xlsx"
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Заявки', index=False)
            worksheet = writer.sheets['Заявки']
            worksheet.column_dimensions['A'].width = 15
            worksheet.column_dimensions['B'].width = 15
            worksheet.column_dimensions['C'].width = 20
            worksheet.column_dimensions['D'].width = 30
            worksheet.column_dimensions['E'].width = 40
            worksheet.column_dimensions['F'].width = 20
            worksheet.column_dimensions['G'].width = 20
            worksheet.column_dimensions['H'].width = 10
            worksheet.column_dimensions['I'].width = 15
        
        print(f"✅ Данные экспортированы в {filename}")
        return filename
        
    except Exception as e:
        logging.error(f"Ошибка экспорта в Excel: {e}")
        return False

async def save_to_database(data, user_id):
    """Сохранение заявки в локальную базу данных"""
    try:
        conn = sqlite3.connect('submissions.db', check_same_thread=False)
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO applications (store_name, product_url, description, full_name, user_id)
            VALUES (?, ?, ?, ?, ?)
        ''', (
            data['store_name'],
            data['product_url'],
            data['description'],
            data['full_name'],
            user_id
        ))
        
        application_id = cursor.lastrowid
        
        if data.get('photo_file'):
            photo_filename = await download_photo(data['photo_file'], application_id)
            if photo_filename:
                cursor.execute('''
                    UPDATE applications SET photo_filename = ? WHERE id = ?
                ''', (photo_filename, application_id))
        
        conn.commit()
        conn.close()
        
        await export_to_excel()
        return True
    except Exception as e:
        logging.error(f"Ошибка сохранения в базу данных: {e}")
        return False

# Команда /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    
    # Всегда показываем команды при старте
    commands_text = """
Привет! Ты зашел в бот для сбора заявок "Китайских штучек".

Вот тебе подсказка по пользованию.
📋 Доступные команды:
/start - Начать заявку
/help - Помощь
/cancel - Отмена

А теперь начнём оформление заявки.
Выбери магазин:
    """
    
    keyboard = [[store] for store in STORES]
    reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
    
    await update.message.reply_text(commands_text, reply_markup=reply_markup)
    return STORE_NAME

# Команда /help
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = """
🤖 **Бот для заявок на товары**

📋 **Команды:**
/start - Начать новую заявку
/help - Показать это сообщение  
/cancel - Отменить текущую заявку
/export - Создать Excel файл с заявками

🛒 **Процесс заявки:**
1. Выберите магазин
2. Пришлите ссылку на товар
3. Опишите товар (можно с фото)
4. Укажите ваше ФИО
5. Подтвердите заявку

📊 **Результат:** Заявка сохраняется в базу и Excel файл
    """
    await update.message.reply_text(help_text)

# Команда /new
async def new_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await start(update, context)

# Выбор магазина
async def get_store_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    store_name = update.message.text
    
    if store_name not in STORES:
        keyboard = [[store] for store in STORES]
        reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
        
        await update.message.reply_text(
            "Пожалуйста, выберите магазин из предложенных вариантов:",
            reply_markup=reply_markup
        )
        return STORE_NAME
    
    context.user_data['store_name'] = store_name
    
    await update.message.reply_text(
        f"✅ Выбрано: {store_name}\n\n📎 Теперь пришлите ссылку на товар:",
        reply_markup=ReplyKeyboardRemove()
    )
    return PRODUCT_URL

# Получение ссылки на товар
async def get_product_url(update: Update, context: ContextTypes.DEFAULT_TYPE):
    product_url = update.message.text
    
    if not (product_url.startswith('http://') or product_url.startswith('https://')):
        await update.message.reply_text("❌ Пожалуйста, введите корректную ссылку (http:// или https://):")
        return PRODUCT_URL
    
    context.user_data['product_url'] = product_url
    await update.message.reply_text("✅ Отлично! Уточните параметры (цвет, размер или фото с указанием конкретной модели)")
    return DESCRIPTION

# Получение описания товара
async def get_description(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.photo:
        photo_file = update.message.photo[-1]
        context.user_data['photo_file'] = photo_file
        
        if update.message.caption:
            context.user_data['description'] = update.message.caption
        else:
            context.user_data['description'] = "Фото товара"
    else:
        context.user_data['description'] = update.message.text
        context.user_data['photo_file'] = None
    
    await update.message.reply_text("✅ Принято! Теперь ваше имя и фамилия:")
    return FULL_NAME

# Получение ФИО
async def get_full_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['full_name'] = update.message.text
    await show_summary(update, context)
    return CONFIRMATION

async def show_summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data
    
    summary = f"""📋 **Проверьте данные:**

🏪 Магазин: {data['store_name']}
🔗 Ссылка: {data['product_url']}
📝 Описание: {data['description']}
👤 ФИО: {data['full_name']}"""
    
    if data.get('photo_file'):
        summary += f"\n📷 Фото: будет сохранено"
    
    summary += "\n\n✅ Всё верно?"
    
    if data.get('photo_file'):
        await update.message.reply_photo(
            photo=data['photo_file'].file_id,
            caption=summary,
            reply_markup=ReplyKeyboardMarkup([['✅ Всё верно', '✏️ Изменить что-то']], 
                                           one_time_keyboard=True, resize_keyboard=True)
        )
    else:
        await update.message.reply_text(
            summary,
            reply_markup=ReplyKeyboardMarkup([['✅ Всё верно', '✏️ Изменить что-то']], 
                                           one_time_keyboard=True, resize_keyboard=True)
        )

# Подтверждение данных
async def confirm_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_response = update.message.text
    
    if user_response == '✅ Всё верно':
        user_id = update.message.from_user.id
        data = context.user_data
        
        db_success = await save_to_database(data, user_id)
        
        if db_success:
            await update.message.reply_text(
                "🎉 **Заявка принята!**\n\n"
                "Для новой заявки отправьте /start\n"
                "Все команды: /help",
                reply_markup=ReplyKeyboardRemove()
            )
        else:
            await update.message.reply_text(
                "❌ Ошибка сохранения. Попробуйте /start",
                reply_markup=ReplyKeyboardRemove()
            )
        
        context.user_data.clear()
        return ConversationHandler.END
        
    elif user_response == '✏️ Изменить что-то':
        keyboard = [
            ['🏪 Изменить магазин'],
            ['🔗 Изменить ссылку'],
            ['📝 Изменить описание'],
            ['👤 Изменить ФИО']
        ]
        reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
        
        await update.message.reply_text("Что изменить?", reply_markup=reply_markup)
        return EDIT_CHOICE

# Выбор что изменить
async def edit_choice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    choice = update.message.text
    
    if choice == '🏪 Изменить магазин':
        keyboard = [[store] for store in STORES]
        reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
        await update.message.reply_text("Выберите магазин:", reply_markup=reply_markup)
        return STORE_NAME
        
    elif choice == '🔗 Изменить ссылку':
        await update.message.reply_text("Новая ссылка:")
        return PRODUCT_URL
        
    elif choice == '📝 Изменить описание':
        await update.message.reply_text("Новое описание:")
        return DESCRIPTION
        
    elif choice == '👤 Изменить ФИО':
        await update.message.reply_text("Ваше ФИО:")
        return FULL_NAME

# Отмена
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        '❌ Заявка отменена.\nДля новой заявки: /start',
        reply_markup=ReplyKeyboardRemove()
    )
    context.user_data.clear()
    return ConversationHandler.END

# Команда для экспорта в Excel
async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    filename = await export_to_excel()
    
    if filename:
        await update.message.reply_text(f"✅ Файл создан: {filename}")
    else:
        await update.message.reply_text("❌ Нет данных для экспорта")

def main():
    TOKEN = '7865276355:AAEBBPoWdik7u-xWLD50S6ge7MvRyUKcRS8'
    
    init_db()
    application = Application.builder().token(TOKEN).build()
    
    # Обработчик разговора
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            STORE_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_store_name)],
            PRODUCT_URL: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_product_url)],
            DESCRIPTION: [MessageHandler(filters.TEXT | filters.PHOTO, get_description)],
            FULL_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_full_name)],
            CONFIRMATION: [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm_data)],
            EDIT_CHOICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, edit_choice)],
        },
        fallbacks=[CommandHandler('cancel', cancel)]
    )
    
    # Команды
    application.add_handler(conv_handler)
    application.add_handler(CommandHandler('help', help_command))
    application.add_handler(CommandHandler('new', new_command))
    application.add_handler(CommandHandler('export', export_excel))
    application.add_handler(CommandHandler('cancel', cancel))
    
    print("Бот запущен...")
    application.run_polling()

if __name__ == '__main__':
    main()