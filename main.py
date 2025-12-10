import os
import pandas as pd
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    filters,
    ContextTypes
)

from processor import calculate_cvr, calculate_cvi

TOKEN = "8201546747:AAGChpoZ8U9e1qsg0SQKvnuOhFpIAEBMq3M"

user_state = {}

MENU = ReplyKeyboardMarkup(
    [["CVR", "CVI"]],
    resize_keyboard=True
)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "سلام! یکی از تحلیل‌ها رو انتخاب کن:",
        reply_markup=MENU
    )
    user_state[update.effective_chat.id] = None


async def choose_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    chat_id = update.effective_chat.id

    if text not in ["CVR", "CVI"]:
        await update.message.reply_text("فقط یکی از گزینه‌های موجود را انتخاب کن.", reply_markup=MENU)
        return

    user_state[chat_id] = text
    await update.message.reply_text("حالا فایل اکسل را ارسال کن (xlsx).")


async def file_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    mode = user_state.get(chat_id)

    if mode is None:
        await update.message.reply_text("اول با /start شروع کن.")
        return

    file = await update.message.document.get_file()
    filepath = f"/tmp/input_{chat_id}.xlsx"
    await file.download_to_drive(filepath)

    # خواندن فایل
    excel = pd.ExcelFile(filepath)
    first_sheet = excel.sheet_names[0]  # شیت اول
    df = excel.parse(first_sheet)

    # پردازش
    if mode == "CVR":
        result_df = calculate_cvr(df)
        out_name = "CVR"
    elif mode == "CVI":
        result_df = calculate_cvi(df)
        out_name = "CVI"

    # تولید خروجی
    outpath = f"/tmp/output_{chat_id}.xlsx"
    with pd.ExcelWriter(outpath, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name=out_name, index=False)

    await update.message.reply_document(open(outpath, "rb"))

    # پاکسازی
    os.remove(filepath)
    os.remove(outpath)


def main():
    port = int(os.environ.get("PORT", 8080))
    webhook_url = "https://reliablety-bot.onrender.com/webhook"

    application = (
        ApplicationBuilder()
        .token(TOKEN)
        .build()
    )

    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, choose_mode))
    application.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), file_received))

    application.run_webhook(
        listen="0.0.0.0",
        port=port,
        url_path="/webhook",
        webhook_url=webhook_url,
    )


if __name__ == "__main__":
    main()
