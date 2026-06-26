# job-market-scraper-indonesia-

# 🇮🇩 Job Market Scraper Indonesia

Python-based web scraper that collects **Data Analyst**, **Business Intelligence**, **Reporting Analyst**, and **Data Scientist** job vacancies from multiple Indonesian job portals.

The scraper automatically extracts:

- Job Position
- Company
- Location
- Industry
- Salary
- Experience
- Skills Required
- Job Requirements
- Posting Date
- Source Website

and exports everything into a formatted Excel report.

---

# Features

✔ Search multiple job websites automatically

✔ Scrape job information

✔ Extract required skills using keyword detection

✔ Generate skill frequency analysis

✔ Generate source summary

✔ Export to formatted Excel

✔ Automatic logging

---

# Supported Sources

- LinkedIn Jobs
- JobStreet
- Glints
- Kalibrr
- Indeed
- DuckDuckGo Search

---

# Output

Excel Workbook

- Jobs
- Skill Frequency
- Source Summary
- Scraping Log

---

# Technologies

- Python
- Pandas
- Requests
- BeautifulSoup4
- DuckDuckGo Search
- OpenPyXL

---

# Installation

Clone repository

```bash
git clone https://github.com/yourusername/job-market-scraper-indonesia.git
```

Install packages

```bash
pip install -r requirements.txt
```

Run

```bash
python job_market_scraper.py
```

---

# Sample Output

| Position | Company | Location | SQL | Python |
|----------|----------|----------|------|---------|
| Data Analyst | ABC Company | Jakarta | Yes | Yes |

---

# Folder Structure

```
job-market-scraper-indonesia
│
├── job_market_scraper.py
├── requirements.txt
├── README.md
├── output/
└── docs/
```

---

# Author

Goldina Raina Putri

LinkedIn

https://www.linkedin.com/in/goldina-raina/

---

# License

MIT License
