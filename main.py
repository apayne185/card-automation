import os
import json
import imgkit
from google.oauth2.service_account import Credentials
from jinja2 import Environment, FileSystemLoader
from googleapiclient.discovery import build
from dotenv import load_dotenv

load_dotenv()

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SHEET_ID = os.getenv("SHEET_ID")
# SHEET_RANGE= "Sheet1!D:J"     #sheet/page 1 cols D:J   
SHEET_RANGE= "FormResponses1!C:J"


'''
    col d:[3] = recipient_name
    col g:[6] = template_choice 
    col h:[7] = message_address    ("dear ____")
    col i:[8] = message_body       ("happy valentines day!")
    col j:[9] = sender_signature   ("love ___")
'''


def fetch_sheet_data():
    creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)     #authenticates using API w service account json file
    service = build("sheets", "v4", credentials=creds)      #builds sheets API service object
    sheet = service.spreadsheets()

    try:
        result = sheet.values().get(spreadsheetId=SHEET_ID, range=SHEET_RANGE).execute()
        data = result.get("values", [])

        if not data:
            print("No data found in Google Sheets.")
            return []
        
        print("Fetched Data:", data) 
        return data

    except Exception as e:
        print(f"Error fetching Google Sheets data: {e}")
        return []




def generate_cards(data):
    recipient_count = {}

    env = Environment(loader=FileSystemLoader("templates"))   #sets up Jinja2 environment

    for i, row in enumerate(data[1:], start=1):     #skips header row 
        clean_row = [cell.strip() if isinstance(cell, str) else cell for cell in row]
        print(f"Row {i}: {clean_row}")

        recipient_name = clean_row[0].strip()
        template_choice = clean_row[3].strip()
        message_address = clean_row[4].strip()
        message_body = clean_row[5].strip()
        sender_signature = clean_row[6].strip()

        template_file = f"template{template_choice}.html"

        try:
            template = env.get_template(template_file)
        except Exception as e:
            print(f"Template loading error for {template_file}: {e}")
            continue

        # template_file = f"template{template_choice}.html"    #chooses template
        # template = env.get_template(template_file)


        output_html = template.render(
            message_address=message_address, 
            message_body=message_body, 
            sender_signature=sender_signature)
        
        output_png = convert_to_png(output_html)

        # recipient_name_arr = recipient_name.split()
        card_name = "_".join(recipient_name.split())

        if card_name in recipient_count:
            recipient_count[card_name] += 1
        else:
            recipient_count[card_name] = 1

        numbered_card_name = f"{card_name}_{recipient_count[card_name]:02d}"

        output_path = os.path.join("cards_png", f"_{numbered_card_name}.png")
        # with open(output_path, "w") as f:
        #     f.write(output_png)
            
        with open(output_path, "wb") as f:  
            f.write(output_png)

        print(f"Generated: {output_path}")



def convert_to_png(card_html):
    temp_path = "temp_card.png"
    
    options = {
            'format': 'png',  
            'quality': 100,   
            'width': 1000,  
            'height': 800,   
            'crop-h': '800',
            'crop-w': '1000',
            'enable-local-file-access': ''
        }
    
    try:
        imgkit.from_string(card_html, temp_path, options=options)
        with open(temp_path, "rb") as f:
            output_png = f.read()  # returns binary content of png 
        os.remove(temp_path)  
        return output_png
    except Exception as e:
        print(f"Error generating PNG: {e}")
        return None






def main():
    os.makedirs("cards_png", exist_ok=True)

    data = fetch_sheet_data()
    if not data:
        print("No data found.")      
        return
    
    generate_cards(data)
    print("All cards generated!")


if __name__ == '__main__':
    main()


