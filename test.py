import pandas as pd
import json
import websocket
import requests
from datetime import datetime

WS_URL = "wss://excel-sheet-analyser.onrender.com/"
EXCEL_FILE_URL = "https://drive.google.com/uc?export=download&id=12tpbZ__5DCEh-uguqzFQbLFOvXLLwz6q"  # <-- Change this! to https://drive.google.com/uc?export=download&id=FILE_ID


def send_json_to_ws(json_data):
    try:
        ws = websocket.create_connection(WS_URL)
        ws.send(json.dumps(json_data))
        ws.close()
        print(f"[{datetime.now()}] JSON sent to WebSocket server at {WS_URL}")
    except Exception as e:
        print(f"[{datetime.now()}] WebSocket error: {e}")

def convert_dates_to_iso(df):
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%d-%m-%Y')
    return df

def fetch_excel_file(url):
    response = requests.get(url)
    if response.status_code == 200:
        with open("temp_excel.xlsx", "wb") as f:
            f.write(response.content)
        return "temp_excel.xlsx"
    else:
        raise Exception(f"Failed to download file. Status code: {response.status_code}")

def process_excel_file(file_path):
    xl = pd.ExcelFile(file_path)
    json_data = {}
    for sheet_name in xl.sheet_names[1:]:  # Skip first sheet if needed
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        df.columns = df.columns.str.replace(' ', '_')
        if df.empty:
            df = pd.DataFrame([{col: 0 for col in df.columns}])
        else:
            df = df.fillna(0)
        df = convert_dates_to_iso(df)
        data = json.loads(df.to_json(orient='records'))
        safe_sheet_name = sheet_name.replace(' ', '_')
        json_data[safe_sheet_name] = data
    return json_data

def main():
    try:
        print(f"[{datetime.now()}] Downloading Excel file...")
        file_path = fetch_excel_file(EXCEL_FILE_URL)
        print(f"[{datetime.now()}] Processing Excel file...")
        json_data = process_excel_file(file_path)
        print(f"[{datetime.now()}] Sending JSON to WebSocket server...")
        send_json_to_ws(json_data)
        print(f"[{datetime.now()}] Process completed successfully.")
    except Exception as e:
        print(f"[{datetime.now()}] Error: {e}")

if __name__ == "__main__":
    main()
