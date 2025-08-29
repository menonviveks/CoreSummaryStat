# 🚀 CoreSummaryStat: Excel Summary Generator App

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

### 📈 Optional Visualizations
Generate histograms, boxplots, or charts for exploratory analysis.

---

## ⚙️ Installation

Clone the repository:

```bash
git clone https://github.com/username/CoreSummaryStat.git
cd CoreSummaryStat
```

🛠 Usage
Using as a Python module
```
from core_summary import summarize
import pandas as pd

df = pd.read_csv('data/sample_data.csv')
summary_df = summarize(df)
summary_df.to_excel('output/summary.xlsx', index=False)
```
Using the Excel Summary Generator App
```
python app.py

```
Follow the prompts to select a dataset.

The app generates summary statistics and saves them as an Excel file.

🖥 Making the App External / Executable

You can convert the Python app into a standalone Windows executable:

Install pyinstaller:
```
pip install pyinstaller

```
Build the executable:
```
pyinstaller --onefile app.py

```
Find the executable in the dist folder.

Share the .exe along with any required data files; users can run it by double-clicking.

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
