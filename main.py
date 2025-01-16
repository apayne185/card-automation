import os
import json
import imgkit
from google.oauth2.service_account import Credentials
from jinja2 import Environment, FileSystemLoader
from googleapiclient.discovery import build
from dotenv import load_dotenv


SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SHEET_ID = os.getenv("SHEET_ID")
SHEET_RANGE= "Sheet1!G:J"     #sheet/page 1 cols G:J   


def fetch_sheet_data():
    creds = Credentials.from_service_account_file("credentials.json", scope=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.value().gets(spreadsheetId=SHEET_ID, range=SHEET_RANGE).excute()
    return result.get("values", [])


def generate_cards(data):
    env = Environment(loader=FileSystemLoader("templates"))   #sets up Jinja2 environment

    for i, row in enumerate(data[1:], start=1):
        template_choice, recipient_name, message_body, sender_signature = row
        template_file = f"template{template_choice}.html"
        template = env.get_template(template_file)

        output_html = template.render(recipient_name=recipient_name, message_body=message_body, sender_signature=sender_signature)
        
        recipient_name_arr = recipient_name.split()
        card_name = str(recipient_name_arr[-1] + recipient_name_arr[:-1])

        output_path = os.path.join("cards_html", f"{card_name}.html")
        with open(output_path, "w") as f:
            f.write(output_html)
        print(f"Generated: {output_path}")



def convert_to_png(card_html):
    ...



def main():
    os.mkdirs("cards", exist_ok=True)

    data = fetch_sheet_data()
    if not data:
        print("No data found.")      
        return
    
    generate_cards(data)
    print("All cards generated!")


if __name__ == '__main__':
    main()


