import os
import pandas as pd
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

from processor import calculate_cvr, calculate_cvi

TOKEN = "YOUR_BOT_TOKEN"

# برای مدیریت state هر چت
user_state = {}

MENU = ReplyKeyboardMarkup(
    [["CVR", "CVI"], ["هر دو"]],
    resize_keyboard=True
)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "سلام! نوع تحلیل رو انتخاب کن:",
        reply_markup=MENU
    )
    user_state[update.effective_chat.id] = None


async def choose_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    chat_id = update.effective_chat.id

    if text not in ["CVR", "CVI", "هر دو"]:
        await update.message.reply_text("یکی از گزینه‌ها رو انتخاب کن.", reply_markup=MENU)
        return

    user_state[chat_id] = text
    await update.message.reply_text("فایل اکسل رو ارسال کن (xlsx).")


async def file_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    mode = user_state.get(chat_id)

    if mode is None:
        await update.message.reply_text("اول با /start شروع کن.")
        return

    file = await update.message.document.get_file()
    filepath = f"input_{chat_id}.xlsx"
    await file.download_to_drive(filepath)

    # خواندن فایل
    excel = pd.ExcelFile(filepath)

    outputs = {}

    if mode in ("CVR", "هر دو"):
        df_cvr = excel.parse("CVR")
        outputs["CVR"] = calculate_cvr(df_cvr)

    if mode in ("CVI", "هر دو"):
        df_cvi = excel.parse("CVI")
        outputs["CVI"] = calculate_cvi(df_cvi)

    # ساخت فایل خروجی
    outpath = f"output_{chat_id}.xlsx"
    with pd.ExcelWriter(outpath, engine="openpyxl") as writer:
        for name, df in outputs.items():
            df.to_excel(writer, sheet_name=name, index=False)

    await update.message.reply_document(open(outpath, "rb"))

    # پاکسازی
    os.remove(filepath)
    os.remove(outpath)


def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, choose_mode))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), file_received))

    app.run_polling()


if __name__ == "__main__":
    main()
