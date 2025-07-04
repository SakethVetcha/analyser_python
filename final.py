import pandas as pd
import json
import websocket
from datetime import datetime

WS_URL = "wss://excel-sheet-analyser-1.onrender.com/"
EXCEL_FILE_PATH = r"C:\Users\saket\Downloads\News letter Template.xlsx"  # Local file path
OUTPUT_JSON_FILE = "newsletter_data.json"  # File to save the JSON data

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
            df[col] = df[col].dt.strftime('%Y-%m-%d')
    return df

def process_excel_file(file_path):
    xl = pd.ExcelFile(file_path)
    json_data = {}
    for sheet_name in xl.sheet_names[1:]:  # Skip first sheet
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

def save_json_to_file(json_data, filename):
    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(json_data, f, indent=4)
        print(f"[{datetime.now()}] JSON data saved to {filename}")
    except Exception as e:
        print(f"[{datetime.now()}] Error saving JSON to file: {e}")

def main():
    try:
        print(f"[{datetime.now()}] Processing local Excel file...")
        json_data = process_excel_file(EXCEL_FILE_PATH)
        save_json_to_file(json_data, OUTPUT_JSON_FILE)  # Save data to file
        print(f"[{datetime.now()}] Sending JSON to WebSocket server...")
        send_json_to_ws(json_data)
        print(f"[{datetime.now()}] Process completed successfully.")
    except Exception as e:
        print(f"[{datetime.now()}] Error: {e}")

if __name__ == "__main__":
    main()
