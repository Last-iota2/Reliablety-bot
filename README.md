# Reliability-bot

Reliability Bot is a Telegram bot for calculating questionnaire validation and reliability metrics from Excel spreadsheets.

The bot processes `.xlsx` files and automatically generates analyzed Excel reports for research workflows.

## Supported Metrics

- CVR (Content Validity Ratio)
- CVI (Content Validity Index)
- Omega coefficient / average importance scoring

## Features

- Persian-language Telegram interface
- Automated Excel processing workflow
- Predefined spreadsheet templates
- Webhook-ready deployment architecture
- Environment-based configuration for secure deployment
- Exportable processed Excel reports

## Tech Stack

- Python
- python-telegram-bot
- Pandas
- OpenPyXL

## Project Structure

- `main.py` — Telegram bot entry point
- `templates/` — Excel templates for supported calculations
- `processor.py` — processing modules / calculation and file-processing logic

## Requirements

- Python 3.10+
- `python-telegram-bot[webhooks]`
- `pandas`
- `openpyxl`

Install dependencies:

```bash
pip install -r requirements.txt
```

## Setup

1. Copy the example environment file:

```bash
copy .env.example .env
```

2. Fill in your bot token and public webhook URL in `.env`:

```env
TELEGRAM_TOKEN=your_bot_token
WEBHOOK_URL=https://your-app.example.com/webhook
PORT=8080
```

3. Run the bot:

```bash
python main.py
```

## Templates

The repository includes sample template files in `templates/`:

- `template_cvr.xlsx`
- `template_cvi.xlsx`
- `template_omega.xlsx`

Use these templates to prepare the Excel files before sending them to the bot.

## Deployment Notes

- Do not commit .env files or Telegram bot tokens
- Webhook deployment requires a publicly accessible server
- Recommended platforms: Render, Railway, or Docker-based VPS deployment

## Future Improvements

- Additional reliability metrics
- Statistical report visualization
- Multi-language support
- Admin dashboard and usage analytics
