# 📋 MSU-IIT ROTC Database Automation Tools

This project contains Python scripts designed to automate the verification, cleaning, and synchronization of the **AER Master List** and **Sectioning Lists** for the MSU-IIT ROTC Unit.

## 🛠️ Prerequisites

Before running the scripts, ensure you have Python installed and the required libraries:

```bash
pip install pandas openpyxl
📂 Folder StructureEnsure your project folder is organized exactly like this for the scripts to detect files correctly:Plaintext/ROTC_Project
    │
    ├── MSU-IIT AER 2S25-26 DATA.xlsx    <-- The Master List
    ├── match_and_highlight.py           <-- Script 1 (The Matcher)
    ├── reset_and_clean.py               <-- Script 2 (The Cleaner)
    │
    └── /Sectioning Lists                <-- FOLDER containing all Section files
            ├── A COY.xlsx
            ├── B COY.xlsx
            ├── C COY.xlsx
            └── ... (other section files)
🚀 Tool 1: The Matcher & HighlighterFilename: match_and_highlight.pyThis script verifies student enrollment by comparing every Section List against the Master AER List.What it does:Scans the AER Master List to memorize all officially enrolled students.Checks every Section File in the Sectioning Lists folder.If a student in a section list is NOT found in the Master List, their row is highlighted RED in the section file.Updates the AER Master List:If a student in the Master List IS found in any section file, their name is highlighted BLUE in the Master List.How to Run:Open your terminal in the project folder and run:Bashpython match_and_highlight.py
Output: Modified Section files (Red highlights applied) and a new master file named MSU-IIT AER_Highlighted.xlsx.🧹 Tool 2: The Cleaner & ResetterFilename: reset_and_clean.pyThis script resets the "Sectioning Lists" to a clean, default state. Use this if you need to start over or remove all highlights and temporary data.What it does:Clears Content: Deletes all text/data in Columns J through O (System/Status columns).Removes Colors: Resets all background fills (removes Red highlights).Resets Formatting: Sets all Row Heights and Column Widths back to Excel defaults.How to Run:Open your terminal in the project folder and run:Bashpython reset_and_clean.py
Output: All Excel files in the Sectioning Lists folder are cleaned and reset to default formatting.⚙️ ConfigurationIf your file names or folder locations change, you must edit the CONFIGURATION section at the top of the python scripts:Python# Open the .py file and look for this section:
# --- CONFIGURATION ---
AER_FILE_PATH = r"C:\Your\Path\To\MSU-IIT AER 2S25-26 DATA.xlsx"
SECTION_FOLDER_PATH = r"C:\Your\Path\To\Sectioning Lists"
📝 TroubleshootingError MessageSolutionModuleNotFoundError: No module named 'openpyxl'Run pip install pandas openpyxl in your terminal.Permission DeniedClose all Excel files before running the scripts. Python cannot edit open files.FileNotFoundErrorCheck that the paths in the script match the actual location of your files.