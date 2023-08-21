import time
import openai
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Step 1: Set up OpenAI API
openai.api_key = 'Enter Your OPENAI API Key'

# Step 2: Set up Google Sheets API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('ENTER THE JSON API KEY FILE NAME', scope)
client = gspread.authorize(creds)

# Open the Google Sheet
sheet = client.open('ENTER YOUR SHEET FILE NAME').sheet1

# Fetch all rows
rows = sheet.get_all_records()

for idx, row in enumerate(rows):
    existing_title = row['New Title']
    existing_meta_description = row['New Meta Description']

    # If "New Title" or "New Meta Description" already have content, skip this row
    if existing_title or existing_meta_description:
        continue

    h1_text = row['H1-1']

    # Get a new title using OpenAI API based on H1-1
    response_title = openai.ChatCompletion.create(
        model="gpt-3.5-turbo-16k-0613",
        messages=[
            {
                "role": "user",
                "content": f"write an SEO title based on the H1: '{h1_text}' with a length between 40 to 60 characters."
            }
        ],
        temperature=1,
        max_tokens=60,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    new_title = response_title.choices[0].message['content'].strip()

    meta_description = row['Meta Description 1']
    base_text_for_meta = meta_description if meta_description else new_title

    # Get a new meta description using OpenAI API based on either Meta description 1 or the New Title if empty
    response_meta = openai.ChatCompletion.create(
        model="gpt-3.5-turbo-16k-0613",
        messages=[
            {
                "role": "user",
                "content": f"write an SEO meta description based on the H1: '{base_text_for_meta}' with a length between 80 to 120 characters."
            }
        ],
        temperature=1,
        max_tokens=120,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    new_meta_description = response_meta.choices[0].message['content'].strip()

    # Update the new title and meta description in the Google Sheet
    sheet.update_cell(idx + 2, 3, new_title)  # New Title column
    sheet.update_cell(idx + 2, 6, new_meta_description)  # New Meta description column

    # Update the Title 1 length and Meta description 1 Length with new values
    sheet.update_cell(idx + 2, 4, len(new_title))  # Title 1 length
    sheet.update_cell(idx + 2, 7, len(new_meta_description))  # Meta description 1 Length

    time.sleep(5)

print("All titles and meta descriptions updated!")
