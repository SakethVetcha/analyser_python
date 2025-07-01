import streamlit as st
import pandas as pd
import json
import websocket

WS_URL = "wss://excel-sheet-analyser-1.onrender.com/"

def send_json_to_ws(json_data):
    try:
        ws = websocket.create_connection(WS_URL)
        ws.send(json.dumps(json_data))
        ws.close()
        st.success(f"JSON sent to WebSocket server at {WS_URL}")
    except Exception as e:
        st.error(f"WebSocket error: {e}")

def convert_dates_to_iso(df):
    # Convert any datetime columns to ISO string format
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%Y-%m-%d')
    return df

def get_excel_json():
    st.title("Excel Sheet Data Analyzer")

    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

    if uploaded_file is not None:
        try:
            xl = pd.ExcelFile(uploaded_file)
            json_data = {}

            # Skip the first sheet
            for sheet_name in xl.sheet_names[1:]:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                df.columns = df.columns.str.replace(' ', '_')

                if df.empty:
                    # Create a single row with 0 for each column
                    df = pd.DataFrame([{col: 0 for col in df.columns}])
                else:
                    df = df.fillna(0)

                df = convert_dates_to_iso(df)  # Convert date columns to ISO strings
                data = json.loads(df.to_json(orient='records'))
                safe_sheet_name = sheet_name.replace(' ', '_')
                json_data[safe_sheet_name] = data

            st.subheader("JSON Data of Excel Sheet:")
            st.json(json_data)

            send_json_to_ws(json_data)

        except Exception as e:
            st.error(f"Error reading the file: {str(e)}")

if __name__ == "__main__":
    get_excel_json()
