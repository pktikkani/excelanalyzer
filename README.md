# 🔍 Excel File Analyzer

Analyze Excel files with multiple sheets using **pandas** and **Streamlit** for the UI.

Upload `.xlsx` / `.xls` files and ask questions about your data interactively.

---

## 📦 Installation

1. **Create and activate a virtual environment** (recommended):

   ```bash
   python -m venv venv
   source venv/bin/activate        # macOS / Linux
   venv\Scripts\activate           # Windows
   ```

2. **Install dependencies**:

   ```bash
   pip install -r requirements.txt
   ```



---

## 🚀 Usage

### Option 1 — Streamlit Web App (recommended)

```bash
streamlit run main.py
```

This opens a browser-based UI where you can:

- Upload one or more Excel files via the sidebar 
- View per-file and per-sheet summaries
- Preview data in interactive tables
- Ask natural-language questions about your data
- Use quick-action buttons for common queries

### Option 2 — Command-Line Interface

```bash
python cli_analyzer.py                          # start empty, then load files interactively
python cli_analyzer.py data1.xlsx data2.xlsx    # pre-load files at startup
```

Once inside the REPL you can run any of the commands listed below.

---

## 🛠️ CLI Commands

| Command     | Description                              |
| ----------- | ---------------------------------------- |
| `load`      | Load an Excel file by path               |
| `files`     | List all currently loaded files          |
| `summary`   | Show high-level summary of a file        |
| `info`      | Show detailed info for a specific sheet  |
| `head`      | Preview the first N rows of a sheet      |
| `tail`      | Preview the last N rows of a sheet       |
| `describe`  | Show descriptive statistics for a sheet  |
| `nulls`     | Show null / missing value counts         |
| `unique`    | List unique values in a column           |
| `counts`    | Show value counts for a column           |
| `filter`    | Filter rows by a condition               |
| `compare`   | Compare all loaded files side-by-side    |
| `help`      | Print the help message                   |
| `quit`      | Exit the analyzer                        |

---

## 💡 Example Questions (Streamlit App)

- *"How many rows are in each sheet?"*
- *"What columns are available?"*
- *"Give me a summary of the data"*
- *"What's the average of the sales column?"*
- *"What is the maximum value?"*
- *"Compare the two files"*

---

## 📁 Project Structure

```
excelanalyzer/
├── main.py            # Streamlit web app
├── cli_analyzer.py    # Command-line interface
├── requirements.txt   # Python dependencies
├── README.md          # This file
└── venv/              # Virtual environment (not committed)
```

---

## 🔧 Requirements

- Python 3.9+
- pandas
- numpy
- streamlit
- openpyxl (for reading `.xlsx` files)

---

## 📄 License

This project is provided as-is for personal and educational use.