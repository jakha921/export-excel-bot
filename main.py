import os
import aiogram
from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
from aiogram.contrib.middlewares.logging import LoggingMiddleware

from runner import run

# Replace 'YOUR_BOT_TOKEN' with your actual bot token obtained from BotFather
TOKEN = '6016263556:AAGhzBWRsVFUSAzRLY8zvKvntcBknLQYuw8'

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
        file_path = await bot.get_file(file_id)
        file_name = file_path.file_path.split('/')[-1]

        # Download the file
        file_to_process = await bot.download_file(file_path.file_path)

        # Save the uploaded Excel file as "parsing_file_1.xlsx"
        with open("exeles\parsing_file.xlsx", "wb") as parsing_file:
            parsing_file.write(file_to_process.getvalue())

        # Call the custom function to process the data
        output_data = run()

        print('output_data', output_data)

        # Send the processed file back to the user
        with open(output_data, 'rb') as output_file:
            await bot.send_document(message.chat.id, output_file, caption="Here's your processed file.")

        # Delete the temporary files
        os.remove("exeles/parsing_file.xlsx")
        os.remove("src\\parsed_file.json")
        os.remove(output_data)

    except ValueError as e:
        await message.reply(e)

    except Exception as e:
        print(e)
        await message.reply("Sorry, something went wrong. Please try again.")


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
