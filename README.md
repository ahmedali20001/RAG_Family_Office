# Family Office SEC Extraction & RAG Pipeline

This repository contains an end-to-end pipeline designed to extract structured regulatory profiles for Family Office investment advisers from the SEC IAPD (Investment Adviser Public Disclosure) database, and an AI-powered analytical RAG (Retrieval-Augmented Generation) dashboard to query the dataset.

## 🏗️ Architecture & Process

This project is divided into two main components:

### 1. Data Extraction (`extract_family_offices.py`)
A comprehensive Python extraction script that:
*   Reads an initial list of adviser CRD numbers from `crd.csv`.
*   Interrogates the SEC IAPD REST API (`https://api.adviserinfo.sec.gov/`) to download the full, raw profile for each individual.
*   Saves the unmodified, deeply-nested JSON records directly into the `raw_json/` folder.
*   Validates and cleans the data (parsing exams, employment histories, calculating experience, and flattening disciplinary disclosures).
*   Compiles the structured records into a clean tabular dataset (`family_offices_v2.xlsx`).

### 2. AI Analytical Agent (`app.py`)
A Streamlit web application that functions as an intelligent Data Router and RAG Pipeline:
*   **Two-Step LLM Workflow:** When a user asks a question, the AI first analyzes the column headers of the Excel dataset to determine which specific columns contain the relevant information.
*   **Precise Context Injection:** It uses Pandas to slice the exact columns required (drastically saving tokens) and passes the CSV text directly into the LLM prompt.
*   **Global Aggregation:** Because the AI reads the tabular data across all 50 profiles simultaneously, it achieves 100% accuracy on global math and comparison queries (e.g., *"Who has the most years of experience?"* or *"Who paid the highest total penalties?"*).

---

## 🚀 How to Run Locally

### Prerequisites
*   Python 3.9+
*   An active Google Gemini or OpenAI API Key.

### Setup Instructions
1. **Clone the repository:**
   ```bash
   git clone https://github.com/ahmedali20001/RAG_Family_Office.git
   cd RAG_Family_Office
   ```

2. **Install the dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. *(Optional)* **Run the Extraction Script:**
   If you wish to re-fetch the data directly from the SEC:
   ```bash
   python extract_family_offices.py
   ```
   *This will update the `raw_json/` folder and rebuild `family_offices_v2.xlsx`.*

4. **Launch the AI Dashboard:**
   ```bash
   streamlit run app.py
   ```
   *The application will open in your default web browser at `http://localhost:8501`. Enter your API key in the sidebar to begin querying.*

---

## 🌐 How to Deploy to Streamlit Community Cloud
You can instantly host this dashboard on the internet for free using Streamlit Cloud:
1. Go to [share.streamlit.io](https://share.streamlit.io/).
2. Log in with your GitHub account.
3. Click **Create app**.
4. Select the repository `YOUR_USERNAME/YOUR_REPO_NAME` and set the Main file path to `app.py`.
5. Click **Deploy!**

Within 60 seconds, your application will be live with a public URL!
