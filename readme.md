<div align="center">

# ⚔️ MSU-IIT ROTC Database Automation

### Automated Enrollment Verification & Data Sanitization System
**MSU-IIT ROTC Unit (2nd Semester, A.Y. 2025-2026)**

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?style=for-the-badge&logo=python&logoColor=white)
![Status](https://img.shields.io/badge/Status-Active-success?style=for-the-badge)
![License](https://img.shields.io/badge/License-Open%20Source-orange?style=for-the-badge)

</div>

---

## 📖 Overview

This suite of automation tools is designed to streamline the administrative workflow for the **MSU-IIT ROTC Unit**. It eliminates manual cross-checking between the **AER Master List** and individual **Sectioning Lists** (Coy/Platoon files).

**Key Capabilities:**
* ✅ **Instant Verification:** Cross-references thousands of students in seconds.
* ✅ **Visual Auditing:** Automatically highlights missing or unverified students in **RED**.
* ✅ **Master List Sync:** Updates the main database with **BLUE** indicators for confirmed enrollees.
* ✅ **One-Click Cleanup:** Instantly resets and sanitizes all data files for fresh runs.

---

## 📂 Project Structure

For the automation to work, your directory must be organized **exactly** as shown below:

```text
/ROTC_Project
│
├── 📜 MSU-IIT AER 2S25-26 DATA.xlsx      # 🔒 The Source of Truth (Master List)
├── 🐍 match_and_highlight.py             # ⚙️ Script 1: The Verifier
├── 🐍 reset_and_clean.py                 # 🧹 Script 2: The Cleaner
│
└── 📁 Sectioning Lists                   # 📂 Folder containing all Platoon/Coy files
    ├── A COY.xlsx
    ├── B COY.xlsx
    ├── ...
    └── Z COY.xlsx
🛠️ Installation & SetupEnsure you have Python installed. Then, install the required dependencies:Bashpip install pandas openpyxl
🚀 Usage Guide1️⃣ The Matcher & HighlighterRun this to verify enrollment.This script compares every student in the Sectioning Lists folder against the Master AER List.Logic:If a student in a Section List is NOT in the Master List → Row turns <span style="background-color: #ffcccc; color: black; padding: 2px 5px; border-radius: 3px;">🔴 RED</span>.If a student in the Master List IS found in a Section → Name turns <span style="background-color: #cceeff; color: black; padding: 2px 5px; border-radius: 3px;">🔵 BLUE</span>.Command:Bashpython match_and_highlight.py
Output: Updates all section files and generates MSU-IIT AER_Highlighted.xlsx.2️⃣ The Cleaner & ResetterRun this to reset files to default.Use this tool when you need to start over. It wipes all processing data and formatting.Actions:🗑️ Sanitize: Clears data in System Columns (J through O).🎨 Reset Colors: Removes all Red/Blue background fills.ea Format: Resets Row Height (15) and Column Width (8.43).Command:Bashpython reset_and_clean.py
Output: All files in Sectioning Lists are scrubbed clean.⚙️ ConfigurationIf file names or folder paths change, update the CONFIGURATION block at the top of the Python scripts:Python# --- CONFIGURATION ---
AER_FILE_PATH = r"C:\Your\Path\To\MSU-IIT AER 2S25-26 DATA.xlsx"
SECTION_FOLDER_PATH = r"C:\Your\Path\To\Sectioning Lists"
❓ TroubleshootingIssueCauseSolutionModuleNotFoundErrorMissing libraries.Run pip install pandas openpyxlPermissionDeniedExcel file is open.Close all Excel files and try again.FileNotFoundErrorWrong paths.Check the SECTION_FOLDER_PATH in the script.<div align="center">Built for the MSU-IIT ROTC Corps of CadetsServe the people. Secure the land.</div>
