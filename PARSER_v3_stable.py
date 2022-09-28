from pyrogram import Client, filters, enums
from pyrogram.types import (ReplyKeyboardMarkup, InlineKeyboardMarkup,
                            InlineKeyboardButton)
import datetime
import xlsxwriter
import texts

cd = 'E:/DATA/BACKUP/files/'

api_id = 610888
api_hash = "adc8354e5243aacac6df569567b3bec"
bot_token = "5705921565:AAHVRgCQk4slKjkdR1-c7sFxKAnbn1AHYy"

app = Client(
    "my_bot",
    api_id=api_id, api_hash=api_hash,
    bot_token=bot_token
)


@app.on_message(filters.text & filters.private)
async def echo(client, message):
    global s_msg
    if message.text == '/start':
        s_msg = await app.send_message(message.chat.id, texts.start_text)
    else:
        # --EXCEL--#--------------------------------------------------------------------
        e = datetime.datetime.now()
        today = "%s.%s.%s" % (e.day, e.month, e.year)
        workbook = xlsxwriter.Workbook(f"{cd}{message.text}_{today}_Parsing.xlsx")
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})
        worksheet.write('A1', 'Username', bold)
        worksheet.write('B1', 'Name', bold)
        worksheet.write('C1', 'User_id', bold)
        worksheet.write('D1', 'Phone', bold)
        worksheet.write('E1', 'Status', bold)
        worksheet.write('F1', 'Chat link', bold)

        count = await app.get_chat_members_count(message.text)
        if count >= 5000:
            msg = await app.send_message(message.chat.id,
                                         texts.start_parsing_big % count,
                                         parse_mode=enums.ParseMode.HTML)
        if 1 <= count <= 5000:
            msg = await app.send_message(message.chat.id,
                                         texts.start_parsing_small % count,
                                         parse_mode=enums.ParseMode.HTML)

        member = [x async for x in app.get_chat_members(message.text)]
        for i, w in enumerate(range(len(member)), start=2):
            if member[w].user.username:
                username = f"@{member[w].user.username}"
            else:
                username = ""
            if member[w].user.first_name:
                first_name = member[w].user.first_name
            else:
                first_name = ""

            if member[w].user.last_name:
                last_name = member[w].user.last_name
            else:
                last_name = ""

            if member[w].user.phone_number:
                phone = f"+{member.user.phone_number}"
            else:
                phone = ""

            name = first_name + last_name
            status = f'{member[w].status}'
            status = status.replace('ChatMemberStatus.', '')

            worksheet.set_column(0, 7, 20)
            worksheet.write(f'A{i}', username)
            worksheet.write(f'B{i}', name)
            worksheet.write(f'C{i}', member[w].user.id)
            worksheet.write(f'D{i}', phone)
            worksheet.write(f'E{i}', status)
            worksheet.write(f'F{i}', message.text)
        workbook.close()
        await app.send_document(message.chat.id, f"{cd}{message.text}_{today}_Parsing.xlsx")
        # print(name, member[w].user.id, username, phone, status)

app.run()

# tg: max_reynders