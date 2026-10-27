from telethon.sync import TelegramClient
from telethon.sessions import StringSession

API_ID = int(input("API_ID: "))
API_HASH = input("API_HASH: ")

with TelegramClient(StringSession(), API_ID, API_HASH) as client:
    print("\n=== SESSION_STRING ===")
    print(client.session.save())
    print("=====================\n")
    print("Скопируй строку выше и вставь в переменную TG_SESSION_STRING на Render.")
