<div align="center">

# 💎 Gemstone Measurement Consolidator

### *AI-Powered Quality Control for Precision Manufacturing*

[![Version](https://img.shields.io/badge/version-2.0.0-blue?style=for-the-badge)](https://github.com/yourusername/gemstone-consolidator)
[![Python](https://img.shields.io/badge/python-3.7+-brightgreen?style=for-the-badge&logo=python)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green?style=for-the-badge)](LICENSE)
[![Status](https://img.shields.io/badge/status-production-success?style=for-the-badge)](https://github.com/yourusername/gemstone-consolidator)

**Transform hours of manual work into seconds with intelligent automation**

[Quick Start](#-quick-start) • [Features](#-features) • [Install](#-installation) • [Usage](#-usage)

---

</div>

## 🎯 What It Does

**Orava Gemstone Master Reporter** automatically consolidates multiple Excel measurement files, validates values against tolerances, and generates professional color-coded reports for quality control.

```
📊 Upload Gem Measurement Report Excel Files → ⚙️ Set Tolerances → ✅ Auto-Validate → 📈 Export Master Report Excel File
```

### ✨ Key Features

- ⚡ **Instant Processing** - Handle 1000+ measurements in seconds
- 🎨 **Visual Validation** - Red/green color-coded pass/fail indicators  
- 📊 **Professional Reports** - Formatted Excel with tolerance tables
- 🔄 **Smart Parsing** - Auto-detects headers and measurement types
- 💾 **Session Memory** - Retains tolerance settings until export
- 🏠 **Easy Navigation** - Intuitive 3-screen workflow

---

## 🚀 Quick Start

```bash
# Clone repository
git https://github.com/Nirmana-KAS/gemstone-measurement-consolidator
cd gemstone-measurement-consolidator

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run application
python main.py
```

**That's it! 🎉**

---

## 💻 Installation

### Requirements
- Python 3.7+
- Windows 10/11, macOS 10.14+, or Linux
- 4GB RAM (8GB recommended)

### Dependencies
```txt
PyQt5>=5.15.0       # Modern GUI framework
openpyxl>=3.0.0     # Excel file handling
python-dateutil>=2.8.0
```

---

## 📖 Usage

### Simple 4-Step Workflow

1. **Launch** - Run `python main.py` and click "Get Started"
2. **Upload** - Add multiple Excel files with measurement data
3. **Configure** - Set nominal values and ±tolerances for each type
4. **Export** - Generate professional master report with validation

### Input Excel Format
```
| ID   | Date Time          | Type           | Unit | Value |
|------|--------------------|----------------|------|-------|
| C462 | 2025-12-02 10:30  | Diameter       | mm   | 1.98  |
| C463 | 2025-12-02 10:31  | Concentricity  | µ    | 0.03  |
```

### Output Features
- ✅ Tolerance reference table (light green headers)
- ✅ Color-coded cells (red = fail, black = pass)
- ✅ Final status column (green/red backgrounds)
- ✅ Metadata (inspector name, timestamp)
- ✅ Auto-sorted by file ID

---

## 🏗️ Project Structure

```
gemstone-measurement-consolidator/
├── main.py                      # Application entry
├── app/
│   ├── gui/
│   │   ├── mainwindow.py        # Main GUI controller
│   │   └── tolerancedialog.py   # Tolerance input dialog
│   └── core/
│       ├── parser.py            # Excel parsing
│       ├── validator.py         # Tolerance validation
│       └── excelwriter.py       # Report generation
├── requirements.txt
└── README.md
```

---

## 🔧 Configuration

### Change Default Tolerances
**File:** `app/gui/tolerancedialog.py`
```python
plus.setText("0.05")   # Change default ± tolerance
minus.setText("0.05")
```

### Customize Colors
**File:** `app/core/excelwriter.py`
```python
# Pass status (green)
passfill = PatternFill(start_color="92D050", ...)

# Fail status (red)  
failfill = PatternFill(start_color="FF0000", ...)
```

---

## 🐛 Troubleshooting

| Problem | Solution |
|---------|----------|
| Import errors | Check function names match exactly in `parser.py` |
| All values show FAIL | Ensure tolerance dict keys match column headers with units |
| Empty cells marked FAIL | Update `validator.py` to return `True` for `None` values |
| Slow performance | Split large files (<10MB each) or upgrade to Python 3.9+ |

---

## 🤝 Contributing

```bash
# Fork repo, create branch
git checkout -b feature/YourFeature

# Make changes, commit
git commit -m "✨ Add feature"

# Push and create PR
git push origin feature/YourFeature
```

**Guidelines:** Follow PEP 8, add docstrings, write tests, update README

---

## 📄 License

MIT License - Copyright (c) 2025 Your Name

Permission granted for use, modification, and distribution. See [LICENSE](LICENSE) for details.

---

## 👤 Contact

**Your Name**
- GitHub: [@yourusername](https://github.com/Nirmana-KAS/gemstone-measurement-consolidator)
- Email: shehannirmana.orava@example.com

**Support:** [Open an Issue](https://github.com/Nirmana-KAS/gemstone-measurement-consolidator/issues)

---

<div align="center">

### ⭐ Star this repo if it helped you!

**Made with ❤️ for precision manufacturing**

[🔝 Back to Top](#-gemstone-measurement-consolidator)

</div>

