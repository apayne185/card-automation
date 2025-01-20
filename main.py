import os
import json
import imgkit
from google.oauth2.service_account import Credentials
from jinja2 import Environment, FileSystemLoader
from googleapiclient.discovery import build
from dotenv import load_dotenv


SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SHEET_ID = os.getenv("SHEET_ID")
SHEET_RANGE= "Sheet1D:J"     #sheet/page 1 cols D:J   

'''
    col d:[3] = recipient_name
    col g:[6] = template_choice 
    col h:[7] = message_address    ("dear ____")
    col i:[8] = message_body       ("happy valentines day!")
    col j:[9] = sender_signature   ("love ___")
'''


def fetch_sheet_data():
    creds = Credentials.from_service_account_file("credentials.json", scope=SCOPES)     #authenticates using API w service account json file
    service = build("sheets", "v4", credentials=creds)      #builds sheets API service object
    sheet = service.spreadsheets()
    result = sheet.value().get(spreadsheetId=SHEET_ID, range=SHEET_RANGE).execute()     #fetches range for sheet
    return result.get("values", [])     #list of lists





def generate_cards(data):
    env = Environment(loader=FileSystemLoader("templates"))   #sets up Jinja2 environment

    for i, row in enumerate(data[1:], start=1):     #skips header row 
        try:
            recipient_name = row[0]
            template_choice= row[3]
            message_address = row[4]
            message_body = row[5]
            sender_signature = row[6]
        except:
            print(f"Skipping row {i} due to insufficient data: {row}")
            continue

        template_file = f"template{template_choice}.html"    #chooses template
        template = env.get_template(template_file)

        output_html = template.render(message_address_address=message_address, message_body=message_body, sender_signature=sender_signature)
        output_png = convert_to_png(output_html)

        recipient_name_arr = recipient_name.split()
        card_name = str(recipient_name_arr[-1] + recipient_name_arr[:-1])

        output_path = os.path.join("cards_png", f"{card_name}.png")
        with open(output_path, "w") as f:
            f.write(output_png)
        print(f"Generated: {output_path}")



def convert_to_png(card_html):
    temp_path = "temp_card.png"
    
    options = {
            'format': 'png',  
            'quality': 100,   
            'width': 800,  
            'height': 600   
        }
    
    try:
        imgkit.from_string(card_html, temp_path, options=options)
        with open(temp_path, "rb") as f:
            return f.read()  # returns binary content of png 
    except Exception as e:
        print(f"Error generating PNG: {e}")
        return None



def main():
    os.makedirs("cards", exist_ok=True)

    data = fetch_sheet_data()
    if not data:
        print("No data found.")      
        return
    
    generate_cards(data)
    print("All cards generated!")


if __name__ == '__main__':
    main()


