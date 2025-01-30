import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
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

# Function to load the file (PDF)
def load_file():
    filepath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if filepath:
        try:
            json_data = convert_pdf_to_json(filepath)
            df = convert_json_to_csv(json_data)
            if not df.empty:
                process_file(df)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

# Function to process the file based on requirements
def process_file(df):

    try:
        
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

        save_file(df)
        save_to_word(df)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process file: {e}")

# Function to save the processed file
def save_file(df):
    filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    if filepath:
        try:
            df.to_csv(filepath, index=False)
            messagebox.showinfo("Success", "File processed and saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")

# Function to save the processed file as a Word document
def save_to_word(df):
    filepath = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
    if filepath:
        try:
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
            messagebox.showinfo("Success", "File processed and saved as Word document successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file as Word document: {e}")

# Main application GUI
def main():
    root = tk.Tk()
    root.title("SSM Table Extractor")
    tk.Label(root, text="SSM Table Extractor", font=("Arial", 16)).pack(pady=10)
    load_button = tk.Button(root, text="Upload your shit here", command=load_file, font=("Arial", 12))
    load_button.pack(pady=20)
    root.mainloop()

if __name__ == "__main__":
    main()