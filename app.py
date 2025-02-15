import os
import pandas as pd
import streamlit as st
from datetime import datetime
from docx import Document
from dotenv import load_dotenv
import requests

# Load API key from .env file
load_dotenv()
API_KEY = os.getenv("NANONETS_API_KEY")
MODEL_ID = os.getenv("NANONETS_MODEL_ID")
print("Using API Key:", API_KEY)  # Debugging API Key Load
print("Using Model ID:", MODEL_ID)

# Function to convert PDF to JSON using NanoNets API
def convert_pdf_to_json(filepath):
    url = f'https://app.nanonets.com/api/v2/OCR/Model/{MODEL_ID}/LabelFile/'
    
    # Open the file and prepare the request data
    payload = {}

    files = [
        ('file', (os.path.basename(filepath), open(filepath, 'rb'), 'application/pdf'))
    ]

    # Prepare the headers with Authorization
    headers = {
        'accept': 'multipart/form-data'
        # 'Authorization': requests.auth.HTTPBasicAuth(API_KEY, '')
    }
        
    # Send POST request to the NanoNets API
    response = requests.request("POST", url, headers=headers, auth=requests.auth.HTTPBasicAuth(API_KEY, ''),data=payload, files=files)

    # Check for successful response and handle errors
    if response.status_code != 200:
        raise ValueError(f"API Error: {response.text}")

    # Parse and return JSON data
    print(response.text)
    return response.json()

# Function to extract tables from JSON and convert to CSV format
def convert_json_to_csv(json_data):
    tables = []
    
    # Extract tables from the JSON response
    for entry in json_data.get("result", []):
        for prediction in entry.get("prediction", []):
            if prediction.get("type") == "table":
                table_data = {}
                
                for cell in prediction.get("cells", []):
                    row, col = cell["row"], cell["col"]
                    text = cell["text"].strip(":").strip()
                    
                    if row not in table_data:
                        table_data[row] = {}
                    table_data[row][col] = text
                
                # Convert to DataFrame
                df = pd.DataFrame.from_dict(table_data, orient="index").sort_index()
                df = df.rename(columns=lambda x: f"Column_{x}")  # Standardize column names
                tables.append(df)

                
    # Merge all DataFrames in the list into one DataFrame
    if tables:
        combined_df = pd.concat(tables, ignore_index=True)
    else:
        raise ValueError("No tables found in the JSON response")
    
    return combined_df

# Function to process the file based on requirements
def process_file(df):

        
    # Remove the first column
    df = df.iloc[:, 1:]

    # Remove rows with blank entries in the 'Column_6' column
    if 'Column_6' in df.columns:
        df = df[df['Column_6'].notna()]

    # Remove duplicate rows
    df = df.drop_duplicates()

    # Remove rows where the third column is blank
    if df.shape[1] > 2:  # Ensure the third column exists
        df = df[df.iloc[:, 2].notna()]

    # Promote the first row to header
    df.columns = df.iloc[0]
    df = df[1:]

    # Apply filters for the 'RESIGNATION' column
    resignation_col = next((col for col in df.columns if 'RESIGNATION' in col.upper()), None)
    if resignation_col:
        current_year = datetime.now().year
        df = df[
            (df[resignation_col].isna()) |  # Blank rows
            (df[resignation_col] == '') |  # Empty strings
            # Extract the year and check if it's more than 5 years ago
            (df[resignation_col].str.extract(r'(\d{4})', expand=False).astype(float) > current_year - 5)
        ]

    #Remove rows where the first column has 'COMPANIES'
    if df.shape[1] > 0:  # Ensure the first column exists
        df = df[~df.iloc[:, 0].astype(str).str.contains("COMPANIES", na=False, case=False)]

    return df

# Function to save the processed file as a Word document
def save_to_word(df, filepath):
    doc = Document()
    doc.add_heading("Processed Table", level=1)
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = "Table Grid"
    for i, column in enumerate(df.columns):
        table.cell(0, i).text = column
    for row in df.itertuples(index=False):
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = str(value) if not pd.isna(value) else ""
    doc.save(filepath)

# Streamlit UI
st.title("SSM Table Extractor")
uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

# Define temp file paths at the start
temp_pdf_path = "temp.pdf"
csv_path = "processed_data.csv"
word_path = "processed_data.docx"

if uploaded_file is not None:
    with open(temp_pdf_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    try:
        json_data = convert_pdf_to_json(temp_pdf_path)
        df = convert_json_to_csv(json_data)
        df = process_file(df)
        df.to_csv(csv_path, index=False)
        save_to_word(df, word_path)
        st.download_button("Download CSV", open(csv_path, "rb"), "processed_data.csv", "text/csv")
        st.download_button("Download Word Document", open(word_path, "rb"), "processed_data.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        st.error(f"Error: {e}")
    
    # Cleanup: Remove temp files after the app runs
    try:
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if os.path.exists(csv_path):
            os.remove(csv_path)
        if os.path.exists(word_path):
            os.remove(word_path)
    except Exception as e:
        st.warning(f"Could not delete temp files: {e}")