### 📋 Description

This script automates the processing of CSV reports generated from Nessus scans and creates HTML files with analyses for different risk levels (Medium, High, and Critical). The main functionality is to extract the newest CSV file from the `InputData` folder and generate reports in the `Reports` folder.

### 🚀 How It Works

1. 📁 **Place your CSV report** in the `InputData` folder. The script will only manage the newest CSV file in the folder.
2. ▶️ **Run `Run.ps1`.**
3. 📊 After the script executes, three HTML files will be created in the `Reports` folder, containing reports for Medium, High, and Critical risks.
