# ANP EnhancedBadgeApp — Intelligent Port Access Control & Analytics

[![Python](https://img.shields.io/badge/Python-3.9%20%7C%203.10%20%7C%203.11%20%7C%203.12-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Data-Pandas-150458?logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![GUI](https://img.shields.io/badge/GUI-Tkinter%20%2F%20ttk-orange)](https://docs.python.org/3/library/tkinter.html)
[![Institution](https://img.shields.io/badge/Partner-Agence%20Nationale%20des%20Ports%20(ANP)-00529b)](https://www.anp.org.ma/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**EnhancedBadgeApp** is an operational desktop data analytics application developed in partnership with the **Agence Nationale des Ports (ANP - National Ports Agency of Morocco)**. The platform streamlines port security management by ingesting access badge registries from structured Excel files, providing fuzzy search, automated activity metric dashboards, and a natural language query assistant for access control personnel.

---

## ️ System Architecture

```
 ┌────────────────────────────────────────────────────────────────────────┐
 │                      User Interface Layer (Tkinter)                    │
 │    • Dynamic ttk Data Table (Treeview) with Custom Color Palette       │
 │    • Multi-criteria Filter Bar & Interactive Fuzzy Query Input         │
 │    • Integrated Matplotlib Embedded Figure Canvas                     │
 └───────────────────────────────────┬────────────────────────────────────┘
                                     │
                                     ▼
 ┌────────────────────────────────────────────────────────────────────────┐
 │                      Business & Processing Logic                       │
 │    • DateTime Normalization & Timezone Alignment (pytz / UTC+1)        │
 │    • Dynamic Status Computations (Active, Expired, Revoked)            │
 │    • Statistical Aggregations (Departmental, Port-Terminal Breakdown)  │
 └─────────────────┬──────────────────────────────────┬───────────────────┘
                   │                                  │
                   ▼                                  ▼
 ┌───────────────────────────────────┐  ┌─────────────────────────────────┐
 │       Pandas Data Engine Layer    │  │   Natural Language Assistant    │
 │   • openpyxl Ingestion & Schema   │  │   • difflib Heuristic Matcher   │
 │   • In-Memory DataFrame Indexing  │  │   • Keyword Parsing & Stats     │
 └───────────────────────────────────┘  └─────────────────────────────────┘
```

---

## Key Capabilities

1. **Automated Excel Registry Ingestion:**
   - Cleans and standardizes column formats across large port badge datasets.
   - Robust datetime parsing for activation, expiration, and token modification timestamps.

2. **Fuzzy Search & Sub-Second Querying:**
   - Employs similarity algorithms (`difflib.get_close_matches`) to tolerate typos in employee names, CINs, or badge serial numbers.

3. **Embedded Statistical Visualizations:**
   - Real-time Matplotlib charts rendered directly inside the Tkinter UI without external browser windows.
   - Breakdown of badge statuses, issuance trends over time, and terminal allocation.

4. **Natural Language Query Assistant:**
   - Access officers can query dataset metrics using conversational queries:
     - *"How many active badges are currently registered?"*
     - *"Which department holds the highest number of badges?"*
     - *"What is the distribution of badges issued this month?"*

---

## ️ Installation & Setup

### Prerequisites
- Python 3.9+
- Tkinter (included with standard Python installers on Windows and macOS; on Linux install via `sudo apt-get install python3-tk`)

### Steps
```bash
# 1. Clone the repository
git clone https://github.com/yasmenmad/anp-projects.git
cd anp-projects

# 2. (Optional) Initialize a virtual environment
python -m venv venv
# On Windows:
venv\Scripts\activate
# On Linux/macOS:
source venv/bin/activate

# 3. Install required packages
pip install -r requirements.txt

# 4. Launch the application
python main.py
```

---

## Project Structure

```
anp-projects/
├── anpp.py             # Primary application source
├── main.py             # Application entrypoint
├── requirements.txt    # Pinned Python package dependencies
├── .gitignore          # Repository ignore rules
└── README.md           # Engineering documentation
```

---

## ‍ Author & Institutional Attribution

**Yasmine Daoudi**  
*Computer Science & Software Engineering Student*  
Université Ibn Tofail — Kénitra, Morocco  
Partnership / Practical Training: **Agence Nationale des Ports (ANP)**  
- LinkedIn: [linkedin.com/in/yasmine-daoudi-473260340](https://www.linkedin.com/in/yasmine-daoudi-473260340)
- Email: [yasmineedaoudii@gmail.com](mailto:yasmineedaoudii@gmail.com)
