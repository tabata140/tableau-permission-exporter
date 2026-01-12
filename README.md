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

## Installation
```bash
pip install -r requirements.txt
python gui_app.py
```

## Usage

1. Login with your Tableau Cloud credentials
2. Select content types to export
3. Choose projects from the tree view
4. Click "Export Permissions"

## Built With

Python, Tkinter, Tableau REST API

## License

MIT License
