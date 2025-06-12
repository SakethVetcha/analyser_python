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

def is_date(val):
    try:
        # Accepts both string and numeric dates
        pd.to_datetime(val, errors='raise')
        return True
    except:
        return False

def format_date(val):
    try:
        return pd.to_datetime(val, errors='raise').strftime('%Y-%m-%d')
    except:
        return val

def get_excel_json():
    st.title("Excel Sheet Data Analyzer")

    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

    if uploaded_file is not None:
        try:
            xl = pd.ExcelFile(uploaded_file)
            selected_sheets = xl.sheet_names[1:]  # Skip the first sheet

            dfs = pd.read_excel(
                uploaded_file,
                sheet_name=selected_sheets
            )

            json_data = {}
            for sheet_name in selected_sheets:
                df = dfs[sheet_name]
                df.columns = df.columns.str.replace(' ', '_')
                df = df.fillna(0)

                first_col = df.columns[0]
                date_mask = df[first_col].apply(is_date)

                # Data rows: first column is a date
                data_rows = df[date_mask].copy()
                # Format the first column as ISO date strings
                data_rows[first_col] = data_rows[first_col].apply(format_date)
                data = json.loads(data_rows.to_json(orient='records'))

                # Meta/info rows: first column is not a date
                meta_rows = df[~date_mask]
                meta_info = {}
                for _, row in meta_rows.iterrows():
                    key = str(row[first_col]).strip()
                    if key and key != "0":
                        value = row.iloc[1] if len(row) > 1 else None
                        if pd.isna(value) or value == "":
                            value = None
                        meta_info[key] = value

                # Combine into final structure
                # Sheet name: replace spaces with underscores
                safe_sheet_name = sheet_name.replace(' ', '_')
                sheet_json = {"data": data}
                sheet_json.update(meta_info)
                json_data[safe_sheet_name] = sheet_json

            st.subheader("JSON Data of Excel Sheet:")
            st.json(json_data)

            # Send the JSON to the public Node.js WebSocket server
            send_json_to_ws(json_data)

        except Exception as e:
            st.error(f"Error reading the file: {str(e)}")

if __name__ == "__main__":
    get_excel_json()
