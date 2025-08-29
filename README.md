# 📋 CoreSummaryStat: Excel Summary Generator App

![Python](https://img.shields.io/badge/Python-3.10-blue?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?logo=pandas&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-217346?logo=microsoftexcel&logoColor=white)
![GitHub](https://img.shields.io/badge/GitHub-CoreSummaryStat-black?logo=github)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-1.0.0-orange)

---

## 📝 Overview
**CoreSummaryStat** is a Python-based tool that generates **comprehensive summary statistics** from datasets and exports them into Excel.  
The **Excel Summary Generator App** provides an interactive interface to load data, compute statistics, and save results, even for users with no programming experience.

---

## 🔄 Pipeline / Workflow

### 📥 Load Data
Import CSV or Excel datasets using `pandas`.

### 📊 Compute Summary Statistics
Metrics include:
- Mean, Median, Mode  
- Variance, Standard Deviation  
- Minimum, Maximum, Quartiles  
- Custom calculations

### 💾 Export to Excel
Summary statistics are saved automatically in `.xlsx` format using `openpyxl`.



---

⚙️ Installation and Standalone Setup (Detailed)

Follow these steps to set up CoreSummaryStat + Excel Summary Generator App on your machine and optionally create a standalone executable so users can run it without Python.

1️⃣ Install Python

Download Python 3.10 or later from python.org
.

During installation, check “Add Python to PATH”.

Verify installation:
```
python --version
```

It should show something like:

Python 3.10.x

2️⃣ Install Git (optional, but recommended)

Download Git from git-scm.com
.

Install with default options.

Verify installation:
```
git --version
```
3️⃣ Clone the repository

Open your terminal or Command Prompt, navigate to your desired folder, and run:
```
git clone https://github.com/username/CoreSummaryStat.git
cd CoreSummaryStat
```

If you don’t want to use Git, you can download the ZIP from GitHub and extract it.

4️⃣ Create a Virtual Environment (Recommended)

Isolate your project dependencies:
```
python -m venv venv

```
Activate the environment:

Windows:
```
venv\Scripts\activate

```
Mac/Linux:
```
source venv/bin/activate
```

Your terminal should show (venv).

5️⃣ Install Dependencies
```
pip install -r requirements.txt
```

Installs all necessary packages: pandas, openpyxl, PySimpleGUI (or any GUI library used).

Verify by importing in Python:
```
import pandas as pd
import openpyxl
```

No errors mean the setup is successful.

6️⃣ Run the App
```
python app.py
```

Follow the prompts to select a dataset.

The app generates summary statistics and saves them as an Excel file in the output/ folder.

7️⃣ Create a Standalone Executable (Optional)

If you want users to run the app without installing Python, create a Windows executable:

Install PyInstaller:
```
pip install pyinstaller
```
Place the icon image in the same folder (.ico format)
Build the executable:
```
pyinstaller --noconsole --onefile --icon=icon.ico CoreSummarystat.py
```
Place the icon in the same folder (.ico format)
Locate the .exe in the dist/ folder.

Share the .exe with any required data files; users can double-click to run the app.

💡 Tips:

Place sample_data.csv in the same folder as the .exe or update the path in the app.

You can also create a folder for output/ so Excel files are generated in the correct location.
📂 Project Structure
```
CoreSummaryStat/
│
├── core_summary.py        # Core functions for summary statistics
├── app.py                 # Excel Summary Generator interface
├── requirements.txt       # Python dependencies
├── README.md              # Project documentation
├── .gitignore             # Ignore unnecessary files
└── data/                  # Sample datasets
```
🤝 Contributing

Fork the repository

Create a new branch:
```
git checkout -b feature-name
```

Commit your changes:
```
git commit -m "Add feature"
```

Push to the branch:
```
git push origin feature-name

```
Open a Pull Request

📜 License

This project is licensed under the MIT License.

🙌 Acknowledgements

Developed by Vivek Menon Sreekumar, 
Ph.D. Agricultural Statistics, Department of Agricutural Statistics, 
Faculty of Agriculture, Bidhan Chandra Krishi Viswavidyalaya, Nadia, West Bengal.
