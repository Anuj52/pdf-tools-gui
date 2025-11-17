# PDF Tools — Password Remover, PDF Merger, Word → PDF Converter  
*A single-file, GUI-only toolkit for working with PDF and Word files.*

<p align="center">

<a href="https://github.com/Anuj52/pdf-tools-gui/stargazers">
  <img src="https://img.shields.io/github/stars/Anuj52/pdf-tools-gui?style=for-the-badge" />
</a>

<a href="https://github.com/Anuj52/pdf-tools-gui/issues">
  <img src="https://img.shields.io/github/issues/Anuj52/pdf-tools-gui?style=for-the-badge" />
</a>

<a href="https://github.com/Anuj52/pdf-tools-gui/network/members">
  <img src="https://img.shields.io/github/forks/Anuj52/pdf-tools-gui?style=for-the-badge" />
</a>

<img src="https://img.shields.io/badge/Python-3.8+-blue?style=for-the-badge" />

<img src="https://img.shields.io/badge/Platform-Windows-green?style=for-the-badge" />

</p>


This project provides a **modern, tabbed Tkinter GUI** that includes:

- **PDF Password Tool**  
  - Remove passwords / decrypt PDFs  
  - Re-encrypt with a new password  
  - Batch processing with multi-threading  
  - CSV per-file password mapping  
  - Backup overwritten files  
  - Drag & Drop support  

- **PDF Merge Tool**  
  - Add files or entire folders  
  - Reorder using Move Up / Move Down  
  - Merge into a single PDF  
  - Drag & Drop support  

- **Word → PDF Converter**  
  - Convert `.docx` → `.pdf`  
  - Preserve folder structure (optional)  
  - Multi-file and multi-folder support  
  - Real-time results table (File | Status | Message | Output Path)  
  - “Open Output Folder” button  
  - Requires **Microsoft Word** (Windows only)  
  - Uses `docx2pdf` internally  

This project is a **GUI-only** tool — **no CLI**.

---

## 🚀 Features

### ✔ Password Tool
- Decrypt PDFs using:
  - Common password
  - Per-file password map (CSV)
- Re-encrypt with a new password  
- Skip already-unlocked files  
- Backup originals before overwrite  
- Threaded batch processor  
- Progress bar + status updates

### ✔ Merge PDFs
- Add files or full folders  
- Drag & drop supported  
- Reorder using Move Up / Move Down  
- Save as a single merged PDF  

### ✔ Word → PDF
- Add `.docx` files or entire folders  
- Optional: preserve subfolder structure  
- Detailed results table  
- Shows success/failure for each file  
- One-click: **Open Output Folder**  

---

## 🛠 Installation

### 1. Install Python dependencies

```bash
pip install PyPDF2 pandas docx2pdf tkinterdnd2




Optional:

* `pandas` → for CSV mapping
* `tkinterdnd2` → enables drag-and-drop
* `docx2pdf` → required for Word → PDF

### 2. Run the application

```bash
python pdf_tools_tabbed_word_improved.py
```

---

## 🧩 Requirements

### Mandatory:

* Python 3.8+
* Windows 10 or newer (for Word→PDF)

### For Word → PDF conversion:

* **Microsoft Word must be installed**
* `docx2pdf` uses Word automation (COM)

---

## 📁 Folder Structure

```
your_project/
│
├── pdf_tools_tabbed_word_improved.py   # main GUI app
├── README.md
└── pdf_tools.spec (optional, PyInstaller)
```

---

## 🧪 Word → PDF Behavior

### Preserve Subfolders

If enabled:

```
Input:
    C:\Docs\Reports\2024\jan\file1.docx
Output folder:
    C:\out

Result:
    C:\out\Reports\2024\jan\file1.pdf
```

If disabled:

```
C:\out\file1.pdf
```

---

## 📦 Building a Windows EXE (Optional)

A complete **PyInstaller spec file** is included in the Python script under:

```
PYINSTALLER_SPEC
```

Basic usage:

1. Save it as `pdf_tools.spec`
2. Run:

```bash
pyinstaller pdf_tools.spec
```

This creates:

```
dist/pdf_tools/pdf_tools.exe
```

### Notes:

* Word → PDF requires MS Word even in EXE version
* Drag & Drop (tkinterdnd2) works in EXE

---

## 📜 License

MIT License. Free for personal and commercial use.

---

## ⭐ Credits

* `PyPDF2` — PDF read/write
* `docx2pdf` — Word automation
* `tkinterdnd2` — drag and drop
* `pandas` — CSV reading

