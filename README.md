# 📄 PPT to PDF Converter (Python + COMTypes)

[![Python](https://img.shields.io/badge/Python-3.x-blue)](https://www.python.org/)
[![Windows](https://img.shields.io/badge/OS-Windows-lightgrey)](https://www.microsoft.com/windows)
[![PowerPoint](https://img.shields.io/badge/Requires-Microsoft%20PowerPoint-orange)](https://www.microsoft.com/microsoft-365/powerpoint)
[![License](https://img.shields.io/badge/License-MIT-green)](https://opensource.org/licenses/MIT)

A professional **PowerPoint to PDF converter** using Python and Microsoft PowerPoint COM automation. Converts PPT and PPTX files to PDFs efficiently, including batch folder processing.

---

## 🚀 Features

| Feature            | Description                               |
| ------------------ | ----------------------------------------- |
| Convert PPT & PPTX | Automatically convert all files in folder |
| Batch processing   | Scan folders and subfolders (optional)    |
| Output directory   | PDFs saved in target folder               |
| Lightweight & Fast | Minimal code, fast conversion             |
| COM API            | Uses official Microsoft PowerPoint API    |

---

## 🛠️ Requirements

* Windows OS (COM automation)
* Microsoft PowerPoint installed
* Python 3.8+
* `comtypes` library

```bash
pip install comtypes
```

---

## 📦 Installation & Setup

### 1. Clone or Download Project

```bash
git clone https://github.com/kamrulislam27/PPT-to-PDF-Converter-Python-COMTypes.git
cd PPT-to-PDF-Converter-Python-COMTypes
```

### 2. Install Dependencies

```bash
pip install comtypes
```

### 3. Ensure PowerPoint is Installed

---

## ▶️ How to Use

1. Set source folder (PPT/PPTX) and target folder (PDF output) in script:

```python
source = r"D:\sourceFolder"
target = r"D:\sourceFolder\targetFolder"
ppt_to_pdf(source, target)
```

2. Run the script:

```bash
python pptToPdf.py
```

3. Optional: Use command line args for file/folder:

```bash
python convert.py "C:/path/to/folder" --recursive
```

---

---

## 📁 Folder Structure

```
ppt-to-pdf-converter/
│── pptToPdf.py
│── README.md
└── pptFolder/   # optional input folder
└── pdfFolder/  # optional output folder
```

---

## ⚠️ Notes & Limitations

* PowerPoint **must** be installed
* Only works on **Windows**
* PDF output quality depends on PowerPoint export settings

---

## 📄 License

This project is licensed under the **MIT License**. Modify and redistribute freely.

---

## 📬 Contact

📧 Email: kamrul@ahut.edu.cn  
💻 GitHub: [https://github.com/kamrulislam27](https://github.com/kamrulislam27)

---


