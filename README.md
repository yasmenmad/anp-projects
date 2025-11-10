
ANP PROJECT
---

# 🪪 EnhancedBadgeApp

**EnhancedBadgeApp** is a Python desktop application built with **Tkinter** and **pandas** that allows users to manage and analyze badge data directly from an **Excel file**.
It combines a simple **graphical interface**, **data processing**, and an integrated **AI assistant** to provide smart insights and statistics.

---

##  Project Overview

The app helps users:

* Load, view, and manage badge data from an Excel sheet.
* Search and filter records interactively.
* Visualize statistics and charts using **matplotlib**.
* Interact with a built-in **AI assistant** that answers natural language questions about the data (e.g., “Which employee has the highest number of badges?”).

It’s a complete project that bridges **data management**, **automation**, and **AI-powered insights** — all within a friendly graphical interface.

---

##  Architecture Overview

```
[User Interface: Tkinter]
          ↓
[Application Logic: Python Classes & Functions]
          ↓
[Data Layer: pandas DataFrame]
          ↓
[Excel File (.xlsx)]
          ↘
 [AI Module – Virtual Assistant + Statistics Engine]
```

---

##  Main Features

 **Excel Data Integration**
Easily load badge data stored in Excel files. The app uses pandas for fast data handling and filtering.

 **Interactive Search**
Find specific employees or badge records instantly using flexible search with fuzzy matching.

 **Statistics Dashboard**
Visualize data through charts and metrics powered by matplotlib.

**AI Assistant Module**
Ask natural-language questions about your dataset — the assistant analyzes and summarizes relevant statistics.

 **User-Friendly Interface**
Modern Tkinter UI using `ttk` widgets and `scrolledtext`, designed for clarity and ease of use.

---

##  Technologies Used

| Category        | Libraries / Tools                                     |
| --------------- | ----------------------------------------------------- |
| GUI             | `tkinter`, `ttk`, `scrolledtext`, `messagebox`        |
| Data Processing | `pandas`, `datetime`, `pytz`                          |
| AI Logic        | `difflib` (for fuzzy matching, natural query parsing) |
| Visualization   | `matplotlib`                                          |
| File I/O        | `openpyxl` (Excel data)                               |

---

## 🪜 How It Works

1. **Load Data** → Select an Excel file (`.xlsx`) to import badge data into the interface.
2. **Explore & Search** → Filter, view, and update records interactively.
3. **Ask Questions** → Use the AI module to request statistics or summaries.
4. **Visualize** → Generate plots showing distributions and key performance indicators.

---

##  Example Questions for the AI Module

* “Show me the total number of badges.”
* “Who has the most badges?”
* “Give me the average number of badges per user.”
* “How many badges were created this month?”

---

##  Project Structure

```
EnhancedBadgeApp/
│
├── main.py                # Main Tkinter application
├── assistant_module.py     # AI logic and natural language handling
├── data/                   # Excel files (input/output)
├── screenshots/             # UI and feature previews
├── requirements.txt         # Python dependencies
└── README.md                # Documentation
```

---

##  Installation

1. **Clone this repository:**

   ```bash
   git clone https://github.com/yourusername/EnhancedBadgeApp.git
   cd EnhancedBadgeApp
   ```

2. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

3. **Run the app:**

   ```bash
   python main.py
   ```

---

##  Example Dependencies

```txt
pandas
matplotlib
openpyxl
pytz
```

Tkinter is built into Python by default (no installation needed).

---



##  Future Improvements

*  Connect to a real database instead of Excel.
*  Add export options (PDF, CSV).
*  Implement voice interaction for the AI assistant.
*  Dark/light theme toggle for Tkinter UI.

---

##  Author

**Yasmine Daoudi**
*Engineering Student in Computer Science & Data Science*
Université Ibn Tofail — Kenitra, Morocco

---



---

