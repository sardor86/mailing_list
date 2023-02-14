import asyncio

from pyrogram import Client
from pyrogram.types import InputPhoneContact

import configparser


async def send_message(phone_number: int, photo: str, caption: str, app: Client):
    async with app:
        user_id = (await app.import_contacts([InputPhoneContact(str(phone_number), str(phone_number))])).users[0].id
        await app.send_photo(user_id, photo, caption=caption)


def start_send_message(phone_number: int, photo: str, caption: str) -> None:
    config = configparser.ConfigParser()
    config.read("config.ini")

    api_id = config['pyrogram']['api_id']
    api_hash = config['pyrogram']['api_hash']

    app = Client("my_account", api_id=api_id, api_hash=api_hash)

    app.run(send_message(phone_number, photo, caption, app))
