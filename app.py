import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from docx import Document

# Function to load the CSV file
def load_file():
    filepath = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if filepath:
        try:
            df = pd.read_csv(filepath)
            process_file(df)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

# Function to process the file based on requirements
def process_file(df):
    try:
        # Remove the first 4 columns
        df = df.iloc[:, 4:]

        # Remove rows with blank entries in the 'col6' column
        if 'col6' in df.columns:
            df = df[df['col6'].notna()]

        # Remove duplicate rows
        df = df.drop_duplicates()

        # Remove rows where the second column is blank
        if df.shape[1] > 1:  # Ensure the second column exists
            df = df[df.iloc[:, 1].notna()]

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

# Function to save the processed file as Word document
def save_to_word(df):
    filepath = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
    if filepath:
        try:
            # Create a Word document
            doc = Document()
            doc.add_heading("Processed Table", level=1)

            # Add the table to the Word document
            table = doc.add_table(rows=1, cols=len(df.columns))
            table.style = "Table Grid"

            # Add headers
            for i, column in enumerate(df.columns):
                table.cell(0, i).text = column

            # Add data rows
            for row in df.itertuples(index=False):
                cells = table.add_row().cells
                for i, value in enumerate(row):
                    cells[i].text = str(value) if not pd.isna(value) else ""

            # Save the Word document
            doc.save(filepath)
            messagebox.showinfo("Success", "File processed and saved as Word document successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file as Word document: {e}")

# Main application GUI
def main():
    root = tk.Tk()
    root.title("CSV Table Editor")

    tk.Label(root, text="CSV Table Editor", font=("Arial", 16)).pack(pady=10)

    load_button = tk.Button(root, text="Load CSV File", command=load_file, font=("Arial", 12))
    load_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
