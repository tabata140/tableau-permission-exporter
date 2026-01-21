# Tableau Permission Exporter

GUI tool for bulk exporting Tableau permissions with user-level analysis.

## Features

- Export permissions for projects, workbooks, data sources, views, and flows
- Administrator template expansion (shows actual capabilities)
- Optional group membership export
- User-friendly interface - no coding required

## Requirements

- Python 3.8+
- Tableau Cloud account with Personal Access Token

## 🚀 Quick Start

### For Non-Technical Users (Easy!)

**Download and run - no installation needed:**

1. Click here: [Download TableauPermissionExporter.exe](https://github.com/tabata140/tableau-permission-exporter/releases/latest)
2. Find the downloaded file (usually in your Downloads folder)
3. Double-click `TableauPermissionExporter.exe`
4. If Windows shows a warning:
   - Click "More info"
   - Click "Run anyway"
5. The tool will open - ready to use!

**That's it!** No Python, no coding, no technical setup required.

---

### For Developers (Run from Source)

<details>
<summary>Click to expand instructions</summary>

**What you need:**
- [Python 3.8+](https://www.python.org/downloads/) installed
- [Git](https://git-scm.com/downloads/) installed

**Steps:**

1. Open Command Prompt (Windows) or Terminal (Mac/Linux)

2. Clone the repository:
```bash
   git clone https://github.com/tabata140/tableau-permission-exporter.git
   cd tableau-permission-exporter
```

3. Install dependencies:
```bash
   pip install -r requirements.txt
```

4. Run:
```bash
   python gui_app.py
```

**Alternative: Download without Git**

If you don't have Git:
1. Click the green "Code" button at the top of this page
2. Select "Download ZIP"
3. Extract the ZIP file
4. Open Command Prompt in that folder
5. Run: `pip install -r requirements.txt`
6. Run: `python gui_app.py`

</details>

## Usage

1. Login with your Tableau Cloud credentials
2. Select content types to export
3. Choose projects from the tree view
4. Click "Export Permissions"

## Built With

Python, Tkinter, Tableau REST API

## License

MIT License
