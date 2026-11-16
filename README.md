# ANP EnhancedBadgeApp - Port Logistics Intelligence & Access Control

[![Python](https://img.shields.io/badge/Python-3.9%20%7C%203.10%20%7C%203.11%20%7C%203.12-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Data-Pandas-150458?logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![GUI](https://img.shields.io/badge/GUI-Tkinter%20%2F%20ttk-orange)](https://docs.python.org/3/library/tkinter.html)
[![Institution](https://img.shields.io/badge/Deployment-National_Ports_Agency_(ANP)-005596)](https://www.anp.org.ma/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

An industrial data processing and logistics auditing suite engineered for the **National Ports Agency (Agence Nationale des Ports - ANP, Morocco)**. The platform streamlines port access credential verification, statistical audit generation, and batch record validation.

---

## Data Pipeline Architecture

```mermaid
flowchart TD
    RawData[Raw Maritime Port Excel Sheets] --> Ingest[High-Throughput Pandas Ingestion Engine]
    Ingest --> Clean[Data Normalization & Null Handling]
    Clean --> Fuzzy[Levenshtein Distance Fuzzy Matcher]
    Clean --> Agg[Statistical Aggregation & KPI Engine]
    Fuzzy --> GUI[Interactive Tkinter / ttk Desktop Console]
    Agg --> Export[Export Clean Datasets & Analytical Reports]
```

---

## Technical Capabilities

### 1. High-Throughput Batch Processing
- Ingests and harmonizes heterogeneous spreadsheet structures containing thousands of badge entries.
- Robust schema validation with automatic type inference and exception handling.

### 2. Fault-Tolerant String Matching
- Employs fuzzy string distance metrics based on normalized Levenshtein edit distance:
  
  $$\text{similarity}(s_1, s_2) = 1 - \frac{\text{lev}(s_1, s_2)}{\max(|s_1|, |s_2|)}$$

- Enables rapid identification of badge holders despite typographical errors or phonetic transcription variants.

### 3. Desktop Analytical Console
- Modular Tkinter graphical interface optimized for low-latency search and real-time record inspection in restricted port security environments.

---

## Execution Guide

```bash
# Clone the repository
git clone https://github.com/yasmenmad/anp-projects.git
cd anp-projects

# Install required runtime dependencies
pip install -r requirements.txt

# Launch application console
python main.py
```

---

## Author & Organizational Context

- **Developer:** Yasmine Daoudi
- **Client Organization:** Agence Nationale des Ports (ANP), Kingdom of Morocco
