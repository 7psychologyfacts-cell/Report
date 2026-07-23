# 📊 Hospital Credit Collection & Claim Recovery MIS Report Engine

An enterprise-grade Python application engineered for high-volume **hospital insurance claim recovery, collection tracking, and automated MIS reporting**.

Designed to process complex financial ledgers and TPA/insurer claim files (such as `.xls` / `.xlsx` claim dumps), this system normalizes messy financial records, computes key performance metrics (Outstanding Amount, Collection Recovery %, Pending Claims), and automates executive MIS generation.

---

## ⚙️ Operational Impact

* **Automated Claims Ledger Ingestion:** Replaces manual spreadsheet consolidation of unit-wise hospital claim reports and TPA collection dumps.
* **Instant Financial Reconciliation:** Automatically matches insurance claim IDs, settlement notes, and collection figures across multiple units without manual ledger cross-checking.
* **Reduction in Processing Time:** Reduces the time required to generate weekly/monthly Credit Collection MIS reports from days to a few seconds.
* **Standardized Status & Deduction Tracking:** Categorizes outstanding claims (e.g., Pending with TPA, Approved, Shortfall/Deduction, Disputed) to streamline RCM follow-up workflows.

---

## 💼 Business Impact

* **Maximised Cash Flow Recovery:** Expedites identification of delayed or partial insurance settlements, enabling timely appeals for disputed deductions.
* **Scalable RCM Operations:** Allows multi-hospital healthcare groups to process large volumes of insurance claims and credit ledgers across units without increasing administrative headcount.
* **Executive Visibility & Decision Support:** Provides hospital leadership with standardized, real-time collection metrics and unit-wise performance analytics.
* **Audit Compliance & Accuracy:** Eliminates human transcription errors, ensuring precise financial reporting for internal and external audits.

---

## 👨‍💻 Developer Information

* **Core Engine:** Built on Python 3 with Flask for lightweight WSGI routing and backend processing.
* **Data Processing & Analytics:** Leverages `pandas`, `xlrd`, and `openpyxl` to parse legacy `.xls` and modern `.xlsx` spreadsheet dumps with optimized memory consumption.
* **Modular Code Structure:** Clean separation of file parsing, data transformation logic, and API endpoints for easy extensibility.
* **Serverless & Cloud Deployment Ready:** Designed for quick deployment on cloud platforms (e.g., Vercel / WSGI servers) with minimal setup requirements.

---

## 🌟 Key Capabilities & Features

### 📥 1. Intelligent File Parsing & Normalization
* Supports legacy excel formats (`.xls`) and modern spreadsheets (`.xlsx`).
* Auto-detects and maps dynamic column headers across different hospital unit report exports.
* Cleans numeric strings, currency symbols, and date formats automatically.

### 📈 2. Automated MIS & Collection Analytics
* Calculates net outstanding balances, recovered amounts, and deduction percentages.
* Generates unit-level and payer-level summary breakdowns (CGHS, ECHS, Private TPAs, Corporate Insurers).

### 🖥️ 3. Interactive Web Dashboard & Export
* User-friendly web interface to upload claims dumps and instantly view key analytics.
* Exports structured, clean Excel workbooks ready for executive presentation.

---

## 🏗️ Architecture & Technical Stack

```
   ┌──────────────────┐     Upload (.xls/.xlsx)   ┌───────────────────────┐
   │ Hospital Billing │──────────────────────────►│  Flask Web Engine     │
   │ / RCM Team UI    │                           │  (app.py)             │
   └──────────────────┘                           └──────────┬────────────┘
                                                             │
                                                             ▼
                                                  ┌───────────────────────┐
                                                  │ Data Transformation   │
                                                  │ (Pandas / OpenPyXL)   │
                                                  └──────────┬────────────┘
                                                             │
                                                             ▼
   ┌──────────────────┐     Structured MIS        ┌───────────────────────┐
   │ Analysis-Ready   │◄──────────────────────────┤ Export Engine & UI    │
   │ Excel / Dashboard│                           │ Dashboard             │
   └──────────────────┘                           └───────────────────────┘
```

* **Language:** Python 3.10+
* **Backend Framework:** Flask
* **Data Libraries:** `pandas`, `xlrd`, `openpyxl`, `numpy`

---

## 📁 Repository Structure

```
.
├── app.py                                                   # Main Flask application logic & processing routes
├── collection-report-claim-latest (9)- Sample Data.xls      # Sample input dataset for claims collection
├── requirements.txt                                         # Python package dependencies
└── README.md                                                # Project documentation
```

---

## 🚀 Setup & Installation

### 1. Prerequisites
* Python 3.9 or higher
* `pip` package manager

### 2. Local Setup

```bash
# Clone the repository
git clone https://github.com/your-username/Report-main.git
cd Report-main

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### 3. Running Locally

```bash
python app.py
```
Open `http://localhost:5000` in your web browser.

---

## 🛡️ License & Contributing

Open-source under the **MIT License**. Contributions and improvements are welcome!
