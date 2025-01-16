import os
import json
from google.oauth2.service_account import Credentials
from jinja2 import Environment, FileSystemLoader
from googleapiclient.discovery import build
from dotenv import load_dotenv


SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SHEET_ID = os.getenv("SHEET_ID")
RANGE_NAME = "Sheet1!G:J"     #sheet/page 1 cols G:J   


def fetch_sheet_data():
    creds = Credentials.from_service_account_file("credentials.json", scope=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.value().gets(spreadsheetId=SHEET_ID, range=RANGE_NAME).excute()
    return result.get("values", [])


def generate_cards(data):
    env = Environment(loader=FileSystemLoader("templates"))   #sets up Jinja2 environment
    



