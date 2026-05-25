import os
import tempfile
from pathlib import Path

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

TOKEN = os.environ.get("TELEGRAM_TOKEN")
WEBHOOK_URL = os.environ.get("WEBHOOK_URL")

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"

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
    
    # ارسال فایل نمونه کلی
    for template_name, caption in [
        ("template_cvr.xlsx", "نمونه CVR"),
        ("template_cvi.xlsx", "نمونه CVI"),
        ("template_omega.xlsx", "نمونه OMEGA"),
    ]:
        template_path = TEMPLATES_DIR / template_name
        if template_path.exists():
            with open(template_path, "rb") as doc:
                await update.message.reply_document(doc, caption=caption)

async def choose_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    chat_id = update.effective_chat.id
    
    if text == "راهنما":
        return await help_handler(update, context)

    if text not in ["CVR", "CVI", "OMEGA"]:
        await update.message.reply_text("فقط یکی از گزینه‌های موجود را انتخاب کن.", reply_markup=MENU)
        return
    
    if text == "CVR":
        template_path = TEMPLATES_DIR / "template_cvr.xlsx"
        caption = "این هم فایل نمونه CVR"
    elif text == "CVI":
        template_path = TEMPLATES_DIR / "template_cvi.xlsx"
        caption = "این هم فایل نمونه CVI"
    elif text == "OMEGA":
        template_path = TEMPLATES_DIR / "template_omega.xlsx"
        caption = "این هم فایل نمونه OMEGA"
    else:
        template_path = None
        caption = ""

    if template_path and template_path.exists():
        with open(template_path, "rb") as doc:
            await update.message.reply_document(doc, caption=caption)

    user_state[chat_id] = text
    await update.message.reply_text("حالا فایل تکمیل شده یا فایل خودت رو ارسال کن.")


async def file_received(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    mode = user_state.get(chat_id)

    if mode is None:
        await update.message.reply_text("اول با /start شروع کن.")
        return

    file = await update.message.document.get_file()
    tmp_dir = Path(tempfile.gettempdir())
    filepath = tmp_dir / f"input_{chat_id}.xlsx"
    outpath = tmp_dir / f"{mode}_{chat_id}.xlsx"

    try:
        await file.download_to_drive(str(filepath))

        excel = pd.ExcelFile(str(filepath))
        first_sheet = excel.sheet_names[0]
        df = excel.parse(first_sheet)

        if mode == "CVR":
            result_df = calculate_cvr(df)
            out_name = "CVR"
        elif mode == "CVI":
            result_df = calculate_cvi(df)
            out_name = "CVI"
        elif mode == "OMEGA":
            result_df = calculate_omega(df)
            out_name = "OMEGA"
        else:
            await update.message.reply_text("حالت نامشخص است. لطفاً دوباره /start را بزنید.")
            return

        with pd.ExcelWriter(str(outpath), engine="openpyxl") as writer:
            result_df.to_excel(writer, sheet_name=out_name, index=False)

        with open(outpath, "rb") as doc:
            await update.message.reply_document(doc)
    finally:
        if filepath.exists():
            filepath.unlink()
        if outpath.exists():
            outpath.unlink()


def main():
    if not TOKEN:
        raise RuntimeError("TELEGRAM_TOKEN environment variable is required")
    if not WEBHOOK_URL:
        raise RuntimeError("WEBHOOK_URL environment variable is required")

    port = int(os.environ.get("PORT", 8080))

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
        webhook_url=WEBHOOK_URL,
    )


if __name__ == "__main__":
    main()
