# SSM Table Extractor

A Python-based application to process PDF files and export the results to both CSV and Word formats. This application utilises the Nanonets free API. Please sign up for an account at Nanonets to obtain your own API key and model ID.

## Features
- Extracts tables from PDF files and processes the data.
- Removes specific columns and filters rows based on user-defined criteria.
- Processes resignation dates to filter data accurately.
- Exports the processed data to CSV and Word formats.
- User-friendly GUI for easy interaction.

## Requirements
- Python 3.7+
- tkinter (comes pre-installed with Python)
- pandas
- python-docx
- python-dotenv (for handling environment variables)

## Installation

### Option 1: Clone the Repository

#### Step 1: Navigate to Your Desired Folder
Open your **terminal** (macOS/Linux) or **Command Prompt/PowerShell** (Windows), and navigate to the folder where you want to clone the repository:
- **Windows**:
  ```bash
  cd %HOMEPATH%\Desktop
  ```
- **macOS/Linux**:
  ```bash
  cd ~/Desktop
  ```

#### Step 2: Clone the Repository
Run the following command to clone the repository:
```bash
git clone https://github.com/azlinamanap/table-extractor.git
cd table-extractor
```

#### Step 3: Create a Virtual Environment
In the same terminal/command prompt, execute:
```bash
python -m venv venv
```

#### Step 4: Activate the Virtual Environment
- **Windows**:
  - If you're using Command Prompt:
    ```bash
    venv\Scripts\activate
    ```
  - If you're using PowerShell:
    ```bash
    .\venv\Scripts\Activate.ps1
    ```
- **macOS/Linux**:
  ```bash
  source venv/bin/activate
  ```

When activated, your terminal will display `(venv)` at the beginning of the command line.

#### Step 5: Install Dependencies
Run the following in the terminal/command prompt to install all required packages:
```bash
pip install -r requirements.txt
```

#### Step 6: Verify Installation
Ensure all dependencies are correctly installed by running:
```bash
pip list
```

### Option 2: Download the ZIP File

#### Step 1: Download the ZIP File
1. Navigate to the repository page on GitHub.
2. Click the green **Code** button.
3. Select **Download ZIP**.
4. Extract the ZIP file to your desired location.

#### Step 2: Navigate to the Project Directory
Open your **terminal** (macOS/Linux) or **Command Prompt/PowerShell** (Windows) and navigate to the extracted folder:
- **Windows**:
  ```bash
  cd %HOMEPATH%\Downloads\table-extractor-main
  ```
- **macOS/Linux**:
  ```bash
  cd ~/Downloads/table-extractor-main
  ```

#### Step 3: Create a Virtual Environment
Follow the same instructions as in Option 1, Step 3.

#### Step 4: Activate the Virtual Environment
Follow the same instructions as in Option 1, Step 4.

#### Step 5: Install Dependencies
Follow the same instructions as in Option 1, Step 5.

### Adding Environment Variables

#### Step 1: Create an .env File

To set up the environment variables for the API_KEY and MODEL_ID, you need to create a .env file in the root of the project directory.

1. Create a new file named .env in the project folder.

2. Add the following lines to the .env file:

`NANONETS_API_KEY=your_api_key_here
NANONETS_MODEL_ID=your_model_id_here`

Replace your_api_key_here and your_model_id_here with the actual API key and model ID provided by the Nanonets API.

## Usage

### Running the Application
1. **Open Terminal or Command Prompt**:
   - Navigate to the project directory if you're not already there:
     ```bash
     cd /path/to/table-extractor
     ```

2. **Run the Application**:
   ```bash
   python app.py
   ```

### Interacting with the Application
1. Follow the GUI instructions:
   - Click "Upload your shit here" to ~~upload your shit here~~ select a file for processing.
   - The application will process the file based on predefined rules.
   - Save the processed file as both CSV and Word documents.

### macOS-Specific Note
For macOS users, ensure Python3 is installed. If it's not, install it via Homebrew:
```bash
brew install python
```

You can confirm the installation by running:
```bash
python3 --version
```

### Windows-Specific Note
For Windows users, ensure Python is added to your PATH during installation. You can verify this by running:
```bash
python --version
```
If the command isn't recognized, re-install Python and select the option to "Add Python to PATH" during the setup process.

### Using Visual Studio Code (Optional)
1. Open the project folder in **VS Code**:
   - Open VS Code and click "File > Open Folder".
   - Select the folder where the repository is cloned or extracted.

2. Configure the Virtual Environment:
   - Open the Command Palette (`Ctrl+Shift+P` or `Cmd+Shift+P` on macOS).
   - Type "Python: Select Interpreter" and choose the virtual environment from the list.

3. Run the Application:
   - Open the `app.py` file.
   - Press `F5` or click on "Run > Start Debugging" to run the script.

## License
This project is licensed under the MIT License.
