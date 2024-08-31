import gspread
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import Workbook
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackContext, ConversationHandler
from telegram.ext import filters  # Menggunakan filters bukan Filters
import io
import logging
from queue import Queue  # Add this import
import os
from dotenv import load_dotenv

# Muat variabel dari file .env
load_dotenv()

TELEGRAM_TOKEN = os.getenv('TELEGRAM_TOKEN')
ALLOWED_USER_ID = int(os.getenv('ALLOWED_USER_ID'))  # Pastikan ini adalah integer

# Set up logging
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

# Google Sheets API credentials
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
CREDS_FILE = 'credentials.json'  # Ganti dengan nama file JSON Anda

# Inisialisasi Google Sheets API
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPES)
client = gspread.authorize(creds)


# Konstanta untuk percakapan
COPY_SPREADSHEET_ID, COPY_SHEET_NAME, COPY_DATA_RANGE, COPY_OUTPUT_FILENAME = range(4, 8)
INFO_SPREADSHEET_ID, INFO_SHEET_NAME = range(8, 10)


async def check_user(update: Update) -> bool:
    user_id = update.effective_user.id
    return user_id == ALLOWED_USER_ID

async def start(update: Update, context: CallbackContext) -> int:
    if not await check_user(update):
        await update.message.reply_text("Akses ditolak. Anda tidak memiliki izin untuk menggunakan bot ini. bot eL! ini BOT KHUSUS OWNER")
        return
    logger.info('Received /start command')
    await update.message.reply_text('Selamat datang di bot eL! ini BOT KHUSUS OWNER! Silahkan ketikan /copy atau /info untuk menggunakan bot ini')
    return ConversationHandler.END

async def copy(update: Update, context: CallbackContext) -> int:
    if not await check_user(update):
        await update.message.reply_text("Akses ditolak. Anda tidak memiliki izin untuk menggunakan bot ini.")
        return
    logger.info('Received /copy command')
    await update.message.reply_text('Silakan kirimkan Spreadsheet ID.')
    return COPY_SPREADSHEET_ID

async def receive_spreadsheet_id(update: Update, context: CallbackContext) -> int:
    logger.info('Received Spreadsheet ID: %s', update.message.text)
    spreadsheet_id = update.message.text.strip()  # Menghapus spasi di awal dan akhir

    # Validasi ID Spreadsheet (contoh sederhana)
    if not spreadsheet_id or len(spreadsheet_id) < 30:  # Pastikan ID cukup panjang
        await update.message.reply_text('ID Spreadsheet tidak valid. Silakan coba lagi.')
        return COPY_SPREADSHEET_ID

    context.user_data['spreadsheet_id'] = spreadsheet_id

    try:
        # Ambil sheet names dari spreadsheet
        sheet = client.open_by_key(spreadsheet_id)
        context.user_data['sheet_names'] = sheet.worksheets()
        context.user_data['sheet_titles'] = [s.title for s in context.user_data['sheet_names']]
        
        await update.message.reply_text(f'Nama sheet yang tersedia: {", ".join(context.user_data["sheet_titles"])}\n'
                                  'Silakan kirimkan nama sheet yang diinginkan.')
        return COPY_SHEET_NAME
    except Exception as e:
        logger.error('Error in receive_spreadsheet_id: %s', e)
        await update.message.reply_text(f'Terjadi kesalahan: {e}')
        return ConversationHandler.END

async def receive_sheet_name(update: Update, context: CallbackContext) -> int:
    logger.info('Received sheet name: %s', update.message.text)
    sheet_name = update.message.text
    if sheet_name not in context.user_data['sheet_titles']:
        await update.message.reply_text('Nama sheet tidak valid. Silakan coba lagi.')
        return COPY_SHEET_NAME

    context.user_data['sheet_name'] = sheet_name
    await update.message.reply_text('Silakan kirimkan rentang data yang ingin disalin (misal: A1:A10).')
    return COPY_DATA_RANGE

async def receive_data_range(update: Update, context: CallbackContext) -> int:
    logger.info('Received data range: %s', update.message.text)
    context.user_data['data_range'] = update.message.text
    await update.message.reply_text('Silakan kirimkan nama file output (.xlsx).')
    return COPY_OUTPUT_FILENAME

async def receive_output_filename(update: Update, context: CallbackContext) -> int:
    logger.info('Received output filename: %s', update.message.text)
    output_filename = update.message.text
    context.user_data['output_filename'] = output_filename

    try:
        # Ambil data dari range
        sheet = client.open_by_key(context.user_data['spreadsheet_id']).worksheet(context.user_data['sheet_name'])
        data = sheet.get(context.user_data['data_range'])
        
        # Pastikan data tidak kosong sebelum membuat file
        if not data:
            await update.message.reply_text('Data tidak ditemukan di rentang yang diberikan.')
            return ConversationHandler.END
        
        # Buat file Excel
        wb = Workbook()
        ws = wb.active
        for row in data:
            ws.append(row)
        file_stream = io.BytesIO()
        wb.save(file_stream)
        file_stream.seek(0)

        # Kirim file ke Telegram
        await update.message.reply_document(document=file_stream, filename=output_filename)

        # Hapus data dari sheet setelah disalin
        start_row = int(context.user_data['data_range'].split(':')[0][1:])  # Mendapatkan nomor baris awal
        end_row = int(context.user_data['data_range'].split(':')[1][1:])  # Mendapatkan nomor baris akhir

        # Ambil semua data dari sheet
        all_data = sheet.get_all_values()

        # Hapus data yang telah disalin dan geser data ke atas
        for row in range(start_row - 1, end_row):  # Menghapus data dari baris yang ditentukan
            all_data[row] = []  # Kosongkan data yang telah disalin

        # Geser data ke atas
        new_data = [row for row in all_data if any(row)]  # Ambil hanya baris yang tidak kosong
        for i in range(len(new_data)):
            all_data[i] = new_data[i]  # Pindahkan data yang tersisa ke atas

        # Kosongkan sisa baris di bawah
        for i in range(len(new_data), len(all_data)):
            all_data[i] = [''] * len(all_data[i])  # Kosongkan baris yang tersisa

        # Update sheet dengan data yang telah digeser
        sheet.update('A1', all_data)  # Update seluruh sheet dengan data baru

        await update.message.reply_text('Data telah disalin dan dihapus dari Google Sheets!')
    except Exception as e:
        logger.error('Error in receive_output_filename: %s', e)
        await update.message.reply_text(f'Terjadi kesalahan: {e}')

    return ConversationHandler.END

async def show_data_info(update: Update, context: CallbackContext) -> int:
    if not await check_user(update):
        await update.message.reply_text("Akses ditolak. Anda tidak memiliki izin untuk menggunakan bot ini.")
        return
    logger.info('Received request for data info')
    spreadsheet_id = context.user_data.get('spreadsheet_id')

    if not spreadsheet_id:
        await update.message.reply_text('Silakan masukkan ID Spreadsheet terlebih dahulu.')
        return INFO_SPREADSHEET_ID

    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        sheet_names = spreadsheet.worksheets()
        
        response = "Daftar Sheet dan Jumlah Data:\n"
        
        for sheet in sheet_names:
            all_data = sheet.get_all_values()
            non_empty_rows = [row for row in all_data if any(row)]  # Ambil hanya baris yang tidak kosong
            data_count = len(non_empty_rows)
            response += f"Sheet: {sheet.title}, Jumlah Data: {data_count}\n"

        await update.message.reply_text(response)

    except gspread.exceptions.SpreadsheetNotFound:
        await update.message.reply_text('Spreadsheet tidak ditemukan. Pastikan ID spreadsheet benar.')
    except gspread.exceptions.WorksheetNotFound as e:
        await update.message.reply_text(f'Sheet tidak ditemukan: {str(e)}. Pastikan nama sheet yang diminta benar.')
    except Exception as e:
        logger.error('Error in show_data_info: %s', e)
        await update.message.reply_text(f'Terjadi kesalahan: {e}')

    return ConversationHandler.END

async def receive_spreadsheet_id_info(update: Update, context: CallbackContext) -> int:
    if not await check_user(update):
        await update.message.reply_text("Akses ditolak. Anda tidak memiliki izin untuk menggunakan bot ini.")
        return
    logger.info('Received Spreadsheet ID for info: %s', update.message.text)
    context.user_data['spreadsheet_id'] = update.message.text
    spreadsheet_id = update.message.text

    try:
        # Ambil sheet names dari spreadsheet
        sheet = client.open_by_key(spreadsheet_id)
        context.user_data['sheet_names'] = sheet.worksheets()
        context.user_data['sheet_titles'] = [s.title for s in context.user_data['sheet_names']]
        
        await update.message.reply_text(f'Nama sheet yang tersedia: {", ".join(context.user_data["sheet_titles"])}\n'
                                  'Silakan kirimkan nama sheet yang diinginkan untuk melihat info.')
        return INFO_SHEET_NAME
    except Exception as e:
        logger.error('Error in receive_spreadsheet_id_info: %s', e)
        await update.message.reply_text(f'Terjadi kesalahan: {e}')
        return ConversationHandler.END

async def receive_sheet_name_info(update: Update, context: CallbackContext) -> int:
    if not await check_user(update):
        await update.message.reply_text("Akses ditolak. Anda tidak memiliki izin untuk menggunakan bot ini.")
        return
    logger.info('Received sheet name for info: %s', update.message.text)
    sheet_name = update.message.text
    if sheet_name not in context.user_data['sheet_titles']:
        await update.message.reply_text('Nama sheet tidak valid. Silakan coba lagi.')
        return INFO_SHEET_NAME

    context.user_data['sheet_name'] = sheet_name
    sheet = [s for s in context.user_data['sheet_names'] if s.title == sheet_name][0]

    # Tampilkan informasi data untuk sheet yang dipilih
    all_data = sheet.get_all_values()
    non_empty_rows = [row for row in all_data if any(row)]  # Ambil hanya baris yang tidak kosong
    data_count = len(non_empty_rows)

    if data_count > 0:
        start_row = all_data.index(non_empty_rows[0]) + 1  # Menambahkan 1 untuk baris yang dimulai dari 1
        end_row = start_row + data_count - 1
        await update.message.reply_text(f'Jumlah data: {data_count}\nRentang data: A{start_row}:A{end_row}')
    else:
        await update.message.reply_text('Tidak ada data yang ditemukan di sheet ini.')

    return ConversationHandler.END

async def cancel(update: Update, context: CallbackContext) -> int:
    logger.info('Operation cancelled')
    await update.message.reply_text('Operasi dibatalkan.')
    return ConversationHandler.END

def main() -> None:
    # Initialize an empty update queue
    update_queue = Queue()  
    # Use ApplicationBuilder instead of Updater
    application = ApplicationBuilder().token(TELEGRAM_TOKEN).build()  

    conversation_handler_copy = ConversationHandler(
        entry_points=[CommandHandler('copy', copy)],
        states={
            COPY_SPREADSHEET_ID: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_spreadsheet_id)],
            COPY_SHEET_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_sheet_name)],
            COPY_DATA_RANGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_data_range)],
            COPY_OUTPUT_FILENAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_output_filename)],
        },
        fallbacks=[CommandHandler('cancel', cancel)]
    )

    conversation_handler_info = ConversationHandler(
        entry_points=[CommandHandler('info', show_data_info)],
        states={
            INFO_SPREADSHEET_ID: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_spreadsheet_id_info)],
            INFO_SHEET_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_sheet_name_info)],
        },
        fallbacks=[CommandHandler('cancel', cancel)]
    )

    application.add_handler(conversation_handler_copy)
    application.add_handler(conversation_handler_info)
    application.add_handler(CommandHandler('start', start))

    application.run_polling()  # Use run_polling instead of start_polling

if __name__ == '__main__':
    main()
