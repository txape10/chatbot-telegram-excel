from functools import wraps
from telegram import Update
from telegram.ext import ContextTypes
from config import AUTHORIZED_USERS


def solo_autorizados(func):
    @wraps(func)
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE):
        user_id = update.effective_user.id
        if user_id not in AUTHORIZED_USERS:
            await update.message.reply_text("🚫 No tienes acceso a este bot.")
            return
        return await func(update, context)
    return wrapper
