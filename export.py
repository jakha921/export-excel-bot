import os
from datetime import datetime
import json

import openpyxl
import re


def insert_into_template(template_path, data_list, codes):
    workbook = openpyxl.load_workbook(template_path)
    sheet = workbook.sheetnames[0]
    sheet = workbook[sheet]

    # Create a new list to store the duplicated rows
    duplicated_rows = []

    # Create a list to collect error messages
    error_messages = []

    # Starting from row 4, insert the parsed data into the template
    for row_index, data in enumerate(data_list, start=4):
        try:
            # Insert product data
            product_data = data.get('products', [])
            if product_data:
                for index, product in enumerate(product_data, start=0):

                    date_string = str(data['dob']).split()[0]

                    # Accept date format 01.01.2021 or 01,01,2021
                    if re.match(r'\d{2}.\d{2}.\d{4}', date_string) or re.match(r'\d{2},\d{2},\d{4}', date_string):
                        date_string_with_periods = date_string.replace(',', '.')
                        date_obj = datetime.strptime(date_string_with_periods, "%d.%m.%Y")
                        formatted_date = date_obj.strftime("%d.%m.%Y")
                    else:
                        raise ValueError(
                            f"Incorrect data format, should be 01.01.2021 or 01,01,2021 in invoice {data['invoice_number']}")

                    row = [
                        data['invoice_number'],
                        0,
                        data['sender'],
                        data['phone'],
                        data['passport'],
                        formatted_date,
                        data['weight'],
                        data['phone'],
                        1,
                        0,
                        '',
                        '',
                        '796',
                        product['product_quantity'],
                        product['product_price'],
                        '840',
                        product['product_info'],
                        ''
                    ]

                    # Search code in codes list
                    for code in codes:
                        if code['product'] == product['product_name'].lower().strip():
                            row[10] = code['code']
                            row[11] = code['kod']
                            break

                    duplicated_rows.append(row)

        except ValueError as e:
            # Collect error messages
            error_messages.append(f"Error in invoice {data['invoice_number']}: {str(e)}")

    # Insert duplicated rows into the sheet
    for row_index, row_data in enumerate(duplicated_rows,
                                         start=len(data_list) + 3):  # +3 to account for the header rows
        sheet.append(row_data)

    # Create the 'ready' directory if it doesn't exist
    ready_dir = 'ready'
    os.makedirs(ready_dir, exist_ok=True)

    # Save the updated template to a new Excel file
    output_file_path = os.path.join(ready_dir, 'output.xlsx')
    workbook.save(output_file_path)
    print(f"Data has been inserted into '{output_file_path}'.")

    return output_file_path, error_messages


def open_files():
    print('open_files')
    # Replace 'template_path.xlsx' with the path to your template Excel file
    # Replace 'parsed_data' with the list of dictionaries containing the parsed data
    # parsed_data import from backet json file

    # Get the current working directory
    current_dir = os.getcwd()

    # Construct the file paths using os.path.join to handle the correct path separator
    # parsed_file_path = os.path.join(current_dir, 'src', 'parsed_file.json')
    # product_codes_file_path = os.path.join(current_dir, 'src', 'product_codes.json')

    # Construct the file paths using os.path.join
    template_path = os.path.join(current_dir, 'exeles', 'Template.xlsx')

    parsed_file_path = os.path.join('src', 'parsed_file.json')
    product_codes_file_path = os.path.join('src', 'product_codes.json')

    # Load the parsed data from 'parsed_file.json'
    print('parsed_data')
    with open(parsed_file_path, 'r', encoding='utf-8') as json_file:
        parsed_data = json.load(json_file)

    # Load the codes from 'product_codes.json'
    print('codes')
    with open(product_codes_file_path, 'r', encoding='utf-8') as json_file:
        codes = json.load(json_file)

    return insert_into_template(template_path, parsed_data, codes)


if __name__ == '__main__':
    open_files()