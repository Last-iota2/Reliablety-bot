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

from processor import calculate_cvr, calculate_cvi, calculate_omega

TOKEN = "8201546747:AAGChpoZ8U9e1qsg0SQKvnuOhFpIAEBMq3M"

user_state = {}

MENU = ReplyKeyboardMarkup(
    [["CVR", "CVI"], ["OMEGA"], ["راهنما"]],
    resize_keyboard=True
)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "سلام! یکی از تحلیل‌ها رو انتخاب کن:",
        reply_markup=MENU
    )
    user_state[update.effective_chat.id] = None


HELP_TEXT = """
📘 راهنمای استفاده از بات

برای هر تحلیل باید یک فایل اکسل با فرمت xlsx ارسال کنید.

فرمت‌ها:

1) **CVR:**
- ستون اول: Item
- باقی ستون‌ها: R1, R2, R3, ...

2) **CVI:**
- ستون اول: Item
- ستون‌ها: Clarity_R1, Clarity_R2...
  Relevance_R1, ...
  Simplicity_R1, ...

3) **OMEGA (اهمیت):**
- ستون اول: Item
- ستون‌های امتیازدهی: R1, R2, R3, ... (نمره 1 تا 5)

اگر نیاز داشتید می‌تونید از فایل نمونه‌ای که ارسال می‌کنم استفاده کنید.
"""

async def help_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP_TEXT)
    
    # ارسال عکس/ویدیو آموزشی (اگر داری)
    # مثال:
    # await update.message.reply_photo(open("guide.jpg", "rb"))
    # await update.message.reply_video(open("guide.mp4", "rb"))

    # ارسال فایل نمونه کلی
    await update.message.reply_document(open("templates/template_cvr.xlsx", "rb"), caption="نمونه CVR")
    await update.message.reply_document(open("templates/template_cvi.xlsx", "rb"), caption="نمونه CVI")
    await update.message.reply_document(open("templates/template_omega.xlsx", "rb"), caption="نمونه OMEGA")

async def choose_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    chat_id = update.effective_chat.id
    
    if text == "راهنما":
        return await help_handler(update, context)

    if text not in ["CVR", "CVI", "OMEGA"]:
        await update.message.reply_text("فقط یکی از گزینه‌های موجود را انتخاب کن.", reply_markup=MENU)
        return
    
    if text == "CVR":
        await update.message.reply_document(open("templates/template_cvr.xlsx", "rb"), caption="این هم فایل نمونه CVR")
    elif text == "CVI":
        await update.message.reply_document(open("templates/template_cvi.xlsx", "rb"), caption="این هم فایل نمونه CVI")
    elif text == "OMEGA":
        await update.message.reply_document(open("templates/template_omega.xlsx", "rb"), caption="این هم فایل نمونه OMEGA")
    
    user_state[chat_id] = text
    await update.message.reply_text("حالا فایل تکمیل شده یا فایل خودت رو ارسال کن.")


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
    elif mode == "OMEGA":
        result_df = calculate_omega(df)
        out_name = "OMEGA"

    # تولید خروجی
    outpath = f"/tmp/{out_name}_{chat_id}.xlsx"
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
