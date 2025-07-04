# Excel Sheet Analyser

This repository contains the production script `test.py`, which downloads an Excel file from Google Drive, processes its contents, and sends the data to a WebSocket server.  
All other scripts in this repository are for development and testing purposes only and are **not** required for deployment.

## Features

- Downloads an Excel file from a direct Google Drive link
- Processes all sheets (except the first, by default)
- Converts date columns to ISO format
- Sends the processed data as JSON to a WebSocket server

## Usage

### 1. Clone the Repository

```bash
git clone https://github.com/SakethVetcha/analyser_python.git
cd analyser_python
```

### 2. Install Dependencies

```bash
pip install pandas websocket-client requests
```

### 3. Set the Excel File Download Link

Open `test.py` and set the `EXCEL_FILE_URL` variable to your **direct Excel download link** from Google Drive.  
**Do not use folder or preview links.**  
The format should be:

```python
EXCEL_FILE_URL = "https://drive.google.com/uc?export=download&id=YOUR_FILE_ID"
```
Replace `YOUR_FILE_ID` with your file’s actual ID.

#### **How to Get the Direct Download Link:**
1. Upload your Excel file to Google Drive.
2. Right-click the file and select **Get link**.
3. Copy the link and extract the file ID from it (the string between `/d/` and `/view`).
4. Construct the direct download link as shown above.

### 4. Run the Script

```bash
python test.py
```

## Notes

- Only `test.py` is intended for deployment. Other scripts in this repository are for testing and development.
- Make sure the Excel file you link is not converted to Google Sheets; it should remain in `.xlsx` format.
- The WebSocket endpoint can be configured in the `WS_URL` variable in `test.py`.

## Example

```python
# Inside test.py
WS_URL = "wss://excel-sheet-analyser.onrender.com/"
EXCEL_FILE_URL = "https://drive.google.com/uc?export=download&id=1A2B3C4D5E6F7G8H9I0J"  # Example file ID
```

## License

[MIT](LICENSE)

**For any issues or questions, please open an issue in this repository.**

