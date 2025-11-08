# 🧰 AutoCAD Automation Tool

A lightweight desktop application built with **Python + Tkinter + PyAutoCAD** to automate common tasks in AutoCAD and Civil 3D, such as reading, editing, and replacing Multileader contents, managing terminal data, and interacting with Excel sheets.

This tool includes:
- A simple Tkinter GUI with buttons and progress logs  
- Excel integration for reading and exporting structured data  
- AutoCAD COM automation (via `pyautocad` or `win32com.client`)  
- Self-contained `.exe` build (no external dependencies)  
- Embedded `.ico` icon (no need to ship extra files)

---

## ⚙️ Features
- ✅ Read and parse Multileader text in AutoCAD drawings  
- ✅ Filter by layer (e.g., `_units`)  
- ✅ Replace contents using regex rules (IN/OUT/TERMINAL OUTPUT)  
- ✅ Zoom and highlight entities by handle  
- ✅ Integrate Excel sheets for terminal data management  
- ✅ Export structured data to `.txt` or `.xlsx`  
- ✅ User-friendly Tkinter UI with progress bar and status console  

---

## 🧩 Requirements (for development)
- Python 3.9+  
- AutoCAD or Civil 3D installed on the same machine  
- Required packages:
  ```bash
  pip install pyautocad pywin32 pandas openpyxl tkinter2