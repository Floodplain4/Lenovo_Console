# 🖥️ Lenovo Case Tracker
### Fast, lightweight Lenovo case and parts tracking tool

[![Release](https://img.shields.io/github/v/release/Floodplain4/Lenovo_Console)](https://github.com/Floodplain4/Lenovo_Console/releases)
[![Downloads](https://img.shields.io/github/downloads/Floodplain4/Lenovo_Console/total)](https://github.com/Floodplain4/Lenovo_Console/releases)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Built With](https://img.shields.io/badge/built%20with-PySide6-green)
![License](https://img.shields.io/badge/license-MIT-blue)

---

## 🎬 Demo

![Lenovo Console Demo](Lenovo%20Console%20Demo.gif)

---

## 🚀 Download

👉 **[Download Latest Release](https://github.com/Floodplain4/Lenovo_Console/releases/tag/V2.28)**

> ⚠️ Windows may display a SmartScreen warning on first run.  
> Click **"More Info" → "Run Anyway"**.


## ✨ Features

### 📋 Case Management
- Add entries using:
  - Work Order number
  - Serial Number
  - Status
  - Notes  
- Structured tracking stored in CSV format

---

### 🔧 Parts Tracking
- Quick-select buttons for common parts:
  - Top lid  
  - Hinges  
  - Bezel  
  - LCD  
  - Keyboard  
  - Motherboard  
- Optional **“Other”** field for uncommon parts

---

### ⚡ Fast Workflow Tools
- **Paste from Lenovo ticket**
  - Automatically parses clipboard text  
  - Auto-fills Work Order & Serial Number  
  - Detects parts from case descriptions  

---

### 📊 Dashboard Overview
- Real-time stats for:
  - Total entries  
  - Ordered  
  - Pending  
  - Replaced  
  - Returned  
  - Complete  
- Visual status indicators

---

### 🔄 Entry Management
- Update status with dropdown + button  
- Edit entries via double-click  
- Delete single or multiple entries  
- Bulk actions with confirmation  

---

### 🔍 Filtering & Search
- Search across all fields  
- Filter by:
  - Status  
  - Part type  
- Dynamic log updates  

---

### 📁 Data Handling
- Automatic logging to `lcd_log.csv`  
- Import existing logs  
- Export for backup  
- Automatic backup before import  

---

### 🧠 Quality-of-Life Features
- Duplicate entry detection  
- Timestamp updates on changes  
- Right-click context menu:
  - Copy serial number  
  - Copy work order  
  - Copy full case summary  
- LCD script helper window for quick ticket notes  

---

## 🛠 Installation

### Option 1: Download EXE
1. Download from the release page  
2. Run the `.exe` file  
3. If prompted by Windows:
   - Click **More Info**
   - Click **Run Anyway**

---

### Option 2: Run from Source

```bash
pip install PySide6 pyperclip
python Lenovo_Case_Tracker2.28.py
