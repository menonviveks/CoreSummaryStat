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
1. **📥 Load Data**  
   - Import CSV or Excel datasets using `pandas`.

2. **📊 Compute Summary Statistics**  
   - Metrics include mean, median, mode, variance, standard deviation, min, max, quartiles, and custom calculations.

3. **💾 Export to Excel**  
   - Summary statistics saved automatically in `.xlsx` format using `openpyxl`.

4. **📈 Optional Visualizations**  
   - Generate histograms, boxplots, or charts for exploratory analysis.

---

## ⚙️ Installation

1. Clone the repository:

```bash
git clone https://github.com/username/CoreSummaryStat.git
cd CoreSummaryStat
Install dependencies:

bash
Copy code
pip install -r requirements.txt
🛠 Usage
Using as a Python module:
python
Copy code
from core_summary import summarize
import pandas as pd

df = pd.read_csv('data/sample_data.csv')
summary_df = summarize(df)
summary_df.to_excel('output/summary.xlsx', index=False)
Using the Excel Summary Generator App:
bash
Copy code
python app.py
Follow the prompts to select a dataset.

The app generates summary statistics and saves them as an Excel file.

🖥 Making the App External / Executable
You can convert the Python app into a standalone Windows executable:

Install pyinstaller:

bash
Copy code
pip install pyinstaller
Build the executable:

bash
Copy code
pyinstaller --onefile app.py
Find the executable in the dist folder.

Share the .exe along with any required data files; users can run it by double-clicking.

📂 Project Structure
bash
Copy code
CoreSummaryStat/
│
├── core_summary.py        # Core functions for summary statistics
├── app.py                 # Excel Summary Generator interface
├── requirements.txt       # Python dependencies
├── README.md              # Project documentation
├── .gitignore             # Ignore unnecessary files
└── data/                  # Sample datasets
🤝 Contributing
Fork the repository

Create a new branch (git checkout -b feature-name)

Commit your changes (git commit -m 'Add feature')

Push to the branch (git push origin feature-name)

Open a Pull Request

📜 License
This project is licensed under the MIT License.
