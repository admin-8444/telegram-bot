from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, InputFile
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, ConversationHandler, CallbackContext, CallbackQueryHandler
import pandas as pd
from main import process_docx
import logging
import os

# Logging
logging.basicConfig(level=logging.INFO)

# Bosqichlar
(ISM, LAVOZIM, OBYEKT, RAHBAR, SANASI, NAZORAT, QOIDABUZARLIK, BAND_QIDIRISH) = range(8)

def start(update: Update, context: CallbackContext):
    update.message.reply_text("👋 Assalomu alaykum! Ismingizni kiriting:")
    return ISM

def get_ism(update: Update, context: CallbackContext):
    context.user_data["ism"] = update.message.text
    update.message.reply_text("Lavozimingiz?")
    return LAVOZIM

def get_lavozim(update: Update, context: CallbackContext):
    context.user_data["lavozim"] = update.message.text
    update.message.reply_text("Obyekt nomini kiriting:")
    return OBYEKT

def get_obyekt(update: Update, context: CallbackContext):
    context.user_data["obyekt"] = update.message.text
    update.message.reply_text("Obyekt rahbari kim?")
    return RAHBAR

def get_rahbar(update: Update, context: CallbackContext):
    context.user_data["rahbar"] = update.message.text
    update.message.reply_text("Tekshiruv sanasi (masalan: 2025-05-13):")
    return SANASI

def get_sana(update: Update, context: CallbackContext):
    context.user_data["sana"] = update.message.text
    update.message.reply_text("Nazorat sanasi:")
    return NAZORAT

def get_nazorat(update: Update, context: CallbackContext):
    context.user_data["nazorat"] = update.message.text
    update.message.reply_text("Qoidabuzarlikni kiriting:")
    return QOIDABUZARLIK

def get_qoidabuzarlik(update: Update, context: CallbackContext):
    context.user_data["qoidabuzarlik"] = update.message.text
    update.message.reply_text("🔍 Band uchun kalit so'z kiriting:")
    return BAND_QIDIRISH

def band_qidirish(update: Update, context: CallbackContext):
    query = update.message.text.lower()
    df = pd.read_excel("bandlar.xlsx").dropna()
    topilganlar = df["Hujjat bandi"].dropna().tolist()
    natija = [b for b in topilganlar if query in b.lower()][:5]
    if not natija:
        update.message.reply_text("Hech narsa topilmadi. Yana urinib ko‘ring:")
        return BAND_QIDIRISH

    tugmalar = [[InlineKeyboardButton(band, callback_data=band)] for band in natija]
    reply_markup = InlineKeyboardMarkup(tugmalar)
    update.message.reply_text("Quyidagilardan birini tanlang:", reply_markup=reply_markup)
    return BAND_QIDIRISH

def band_tanlandi(update: Update, context: CallbackContext):
    query = update.callback_query
    query.answer()
    band = query.data
    band_obj = {"matn": band, "muddat": context.user_data["nazorat"]}
    context.user_data.setdefault("bandlar", []).append(band_obj)
    if len(context.user_data["bandlar"]) >= 5:
        query.edit_message_text("Bandlar tanlandi. Hujjatlar tayyorlanmoqda...")
        return generate_docs(query.message.chat_id, context)
    else:
        query.edit_message_text("Band qo‘shildi. Yana band izlang:")
        return BAND_QIDIRISH

def generate_docs(chat_id, context: CallbackContext):
    data = context.user_data
    replacements = {
        "{{ism}}": data.get("ism"),
        "{{lavozim}}": data.get("lavozim"),
        "{{Obyekt_nomi}}": data.get("obyekt"),
        "{{Obyekt_rahbari}}": data.get("rahbar"),
        "{{kun.oy.yil}}": data.get("sana"),
        "{{yil.oy.yil}}": data.get("nazorat"),
        "{{qoidabuzarlik}}": data.get("qoidabuzarlik"),
        "{{Tekshiruvda_qatnashganlar}}": "",
        "{{Termiz}}": "",
    }

    doc1_path = process_docx("hujjat1.docx", replacements, data["bandlar"], include_band_table=True)
    doc2_path = process_docx("hujjat2.docx", replacements, data["bandlar"], include_band_table=False)

    context.bot.send_document(chat_id=chat_id, document=InputFile(doc1_path, filename="Yozma_Korsatma.docx"))
    context.bot.send_document(chat_id=chat_id, document=InputFile(doc2_path, filename="Dalolatnoma.docx"))
    return ConversationHandler.END

def cancel(update: Update, context: CallbackContext):
    update.message.reply_text("❌ Jarayon bekor qilindi.")
    return ConversationHandler.END

def main():
    TOKEN = os.getenv("BOT_TOKEN")
    updater = Updater(TOKEN)
    dp = updater.dispatcher

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            ISM: [MessageHandler(Filters.text & ~Filters.command, get_ism)],
            LAVOZIM: [MessageHandler(Filters.text & ~Filters.command, get_lavozim)],
            OBYEKT: [MessageHandler(Filters.text & ~Filters.command, get_obyekt)],
            RAHBAR: [MessageHandler(Filters.text & ~Filters.command, get_rahbar)],
            SANASI: [MessageHandler(Filters.text & ~Filters.command, get_sana)],
            NAZORAT: [MessageHandler(Filters.text & ~Filters.command, get_nazorat)],
            QOIDABUZARLIK: [MessageHandler(Filters.text & ~Filters.command, get_qoidabuzarlik)],
            BAND_QIDIRISH: [
                MessageHandler(Filters.text & ~Filters.command, band_qidirish),
                CallbackQueryHandler(band_tanlandi)
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    dp.add_handler(conv_handler)
    updater.start_polling()
    updater.idle()

if __name__ == "__main__":
    main()