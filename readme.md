# FluxPDF

**The All-in-One PDF Utility: Remove Passwords, Merge Files, and Convert Word Documents.**

<p align="center">
  <a href="https://github.com/Anuj52/pdf-tools-gui/releases/latest">
    <img src="https://img.shields.io/github/v/release/Anuj52/pdf-tools-gui?style=for-the-badge&color=brightgreen&label=Download%20Latest%20Installer" alt="Download Latest Release" />
  </a>
</p>

<p align="center">
  <a href="https://github.com/Anuj52/pdf-tools-gui/stargazers">
    <img src="https://img.shields.io/github/stars/Anuj52/pdf-tools-gui?style=for-the-badge" alt="Stars"/>
  </a>
  <a href="https://github.com/Anuj52/pdf-tools-gui/issues">
    <img src="https://img.shields.io/github/issues/Anuj52/pdf-tools-gui?style=for-the-badge" alt="Issues"/>
  </a>
  <a href="https://github.com/Anuj52/pdf-tools-gui/network/members">
    <img src="https://img.shields.io/github/forks/Anuj52/pdf-tools-gui?style=for-the-badge" alt="Forks"/>
  </a>
  <img src="https://img.shields.io/badge/Platform-Windows-blue?style=for-the-badge" alt="Platform Windows" />
  <img src="https://img.shields.io/badge/License-MIT-orange?style=for-the-badge" alt="License" />
</p>

---

FluxPDF is a modern, tabbed application designed for productivity. It allows you to manipulate PDF files and convert documents without needing complex command-line tools.

## 📥 Download & Install

**No Python required!** Simply download the Windows Installer.

1.  Go to the [Latest Release Page](https://github.com/Anuj52/pdf-tools-gui/releases/latest).
2.  Download `FluxPDF_Setup.exe`.
3.  Run the installer.
4.  Launch **FluxPDF** from your Desktop or Start Menu.

---

## 🚀 Features

### 🔓 PDF Password Tool
* **Bulk Decrypt:** Remove passwords from multiple PDFs at once.
* **Smart Handling:** Uses a common password or a specific CSV map for different files.
* **Re-Encrypt:** Optionally secure output files with a new password.
* **Safe:** Automatically backs up original files before overwriting.

### 📑 PDF Merge Tool
* **Drag & Drop:** Easily add files or entire folders.
* **Reorder:** Use "Move Up" and "Move Down" buttons to arrange pages.
* **Fast Merge:** Combines distinct PDFs into a single document instantly.

### 📝 Word → PDF Converter
* **Batch Convert:** Turn `.docx` files into `.pdf` using native Microsoft Word automation.
* **Folder Preservation:** Option to keep your subfolder structure in the output directory.
* **Live Tracking:** View a detailed results table (Status, Message, Output Path).
* **One-Click Access:** Button to immediately open the output folder.

---

## 🧪 Word → PDF Behavior

The application includes logic to handle directory structures intelligently.

**Preserve Subfolders Option:**

| Option State | Input Path | Output Path | Result |
| :--- | :--- | :--- | :--- |
| **Enabled** | `C:\Docs\Reports\2024\jan\file1.docx` | `C:\out\Reports\2024\jan\file1.pdf` | Mirrors structure |
| **Disabled** | `C:\Docs\Reports\2024\jan\file1.docx` | `C:\out\file1.pdf` | Flattens all files |

---

## 💻 For Developers: Running from Source

If you prefer to run the Python script directly or contribute to the code, follow these steps.

### 1. Prerequisites
* **Python 3.10+**
* **Microsoft Word** (Required for `.docx` conversion on Windows)

### 2. Installation

```bash
# Clone the repository
git clone [https://github.com/Anuj52/pdf-tools-gui.git](https://github.com/Anuj52/pdf-tools-gui.git)
cd pdf-tools-gui

# Install dependencies
pip install PyPDF2 pandas docx2pdf tkinterdnd2 comtypes

```bash

### 3. Run the App

```bash
python pdf_tools_tabbed_word_improved.py
```bash
## 🛠 Building the Installer (CI/CD)

This project uses GitHub Actions to automatically build and release the Windows installer.

* **Push to Main:** Triggers a build check.
* **Push a Tag (e.g., `v1.0.0`):**
    * Compiles the Python code using `PyInstaller` (OneDir mode).
    * Builds a Setup EXE using `Inno Setup`.
    * Publishes a new GitHub Release with the installer attached.

### Build Locally
Ensure you have **Inno Setup 6** installed.

```bash
pyinstaller FluxPDF.spec --clean --noconfirm
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup.iss

```bash

## ⭐ Credits & License

* **PyPDF2:** PDF manipulation.
* **docx2pdf:** Microsoft Word automation.
* **tkinterdnd2:** Drag-and-drop support for Tkinter.
* **Inno Setup:** Windows installer generation.

Distributed under the **MIT License**.