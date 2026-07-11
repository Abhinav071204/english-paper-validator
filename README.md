# ELA Document Validator & Auto-Corrector ✍️

An automated quality assurance (QA) and formatting compliance tool for English Language Arts (ELA) assessment papers (`.docx`). It scans documents for formatting rules, flags unresolved placeholders, standardizes wording guidelines, checks passage readability grade levels, and provides an auto-corrected document download instantly.


## 🚀 Key Features

* **Formatting Audit**: Checks font sizes, types (e.g., Arial 10.0pt/11.0pt compliance), and heading hierarchies.
* **Wording & Guideline Rules**: Enforces standardization rules (e.g. converting *"refer to the audio"* to *"listen to the audio"*).
* **Placeholder Finder**: Flags leftover bracket markers like `[Insert ID]` or `##` indicators.
* **Readability Analyzer**: Calculates Flesch-Kincaid Grade Levels and word counts for passages.
* **In-Context Highlights**: Displays interactive layout preview with hover tooltips detailing QA violations.
* **1-Click Corrector**: Corrects formatting errors and wording mismatches and downloads a clean Word file directly.


## 🛠️ Tech Stack

* **Backend**: Python (using `python-docx` for document parsing and formatting correction)
* **Frontend Option A**: HTML5, Vanilla CSS3, and JavaScript (with a light HTTP server)
* **Frontend Option B**: Streamlit (Python-only rapid UI layer)


## 📂 Project Structure

* `validator.py` - Core validation engine and document auto-correction logic.
* `streamlit_app.py` - Streamlit application dashboard (best for cloud deployment).
* `server.py` - Lightweight local HTTP server serving the static files.
* `index.html` - Main HTML5 dashboard UI.
* `app.js` - Client-side state manager and preview renderer.
* `style.css` - Custom styling tokens and dark/light layouts.
* `preview.html` - Standalone interactive document preview template with Light/Dark theme.
* `requirements.txt` - Python project package dependencies.


## 💻 How to Run Locally

First, clone the repository and install dependencies:
```bash
pip install -r requirements.txt
