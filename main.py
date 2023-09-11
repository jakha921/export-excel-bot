import os
from asyncio import sleep
from random import random

from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
from aiogram.contrib.middlewares.logging import LoggingMiddleware

from runner import run

# Replace 'YOUR_BOT_TOKEN' with your actual bot token obtained from BotFather
TOKEN = '6016263556:AAEZEx5ph26fKO7VyYfNssNH0Po5wV5lNjA'

# Initialize the bot and dispatcher
bot = Bot(token=TOKEN)
dp = Dispatcher(bot)
dp.middleware.setup(LoggingMiddleware())


# Command handler for start
@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    await message.reply("Hello! Send me any Excel file, and I'll process and return it to you.")


# File handler
@dp.message_handler(content_types=types.ContentTypes.DOCUMENT)
async def handle_file(message: types.Message):
    try:
        # Get the file from the message
        file_id = message.document.file_id
        file_size = message.document.file_size

        # make long proccesing for big files for every 50 kilobytes of file size sleep 1 second and set to random range
        if file_size / 1024 > 50:
            # send message to user
            await message.reply('File is in processing, please wait')

            # get random range if file size more than 700 kilobytes
            random_range = random() * 3 if file_size / 1024 > 700 else random() * 4
            # get sleep time
            sleep_time = file_size / 1024 / 50 * random_range
            # sleep
            print('sleep_time', sleep_time)
            await sleep(sleep_time)

        file_path = await bot.get_file(file_id)

        # Download the file
        file_to_process = await bot.download_file(file_path.file_path)

        # Save the uploaded Excel file as "parsing_file_1.xlsx"
        with open("exeles\parsing_file.xlsx", "wb") as parsing_file:
            parsing_file.write(file_to_process.getvalue())

        # Call the custom function to process the data
        output_data, errors = run()

        # Send the processed file back to the user
        if output_data:
            with open(output_data, 'rb') as output_file:
                await bot.send_document(message.chat.id, output_file, caption="Here's your processed file.")
            print('-' * 50)

            # Delete the temporary files
            if os.path.exists(os.path.join('exeles', 'parsing_file.xlsx')):
                os.remove(os.path.join('exeles', 'parsing_file.xlsx'))
            if os.path.exists(os.path.join('src', 'parsed_file.json')):
                os.remove(os.path.join('src', 'parsed_file.json'))
            os.remove(output_data)

        if errors:
            await message.reply('*Errors:*\n\n' + '\n'.join(errors), parse_mode='Markdown')

    except ValueError as e:
        await message.reply(e)

    except Exception as e:
        print(e)
        await message.reply("Sorry, something went wrong. Please try again.")


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)