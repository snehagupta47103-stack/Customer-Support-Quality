<div align="center">

# -- ! Customer Care Service ! --
### *Customer Support Quality Analysis using Excel, SQL & Python*
**Data Analysis Practical Exam — Set E | Student ID: `YOUR-STUDENT-ID`**

[![Excel](https://img.shields.io/badge/Excel-XLOOKUP%20%7C%20COUNTIFS%20%7C%20Pivot-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/excel)
[![SQL](https://img.shields.io/badge/SQL-JOIN%20%7C%20GROUP%20BY%20%7C%20HAVING-4479A1?style=for-the-badge&logo=postgresql&logoColor=white)](https://www.w3schools.com/sql/)
[![Python](https://img.shields.io/badge/Python-pandas%20%7C%20matplotlib-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Analysis](https://img.shields.io/badge/Analysis-SLA%20Breach%20Rate-FF6F00?style=for-the-badge&logo=databricks&logoColor=white)](https://github.com/)

<br/>

> *"Great service is not about speed alone — it is about resolving issues before the clock runs out."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🗂️ Dataset & Data Dictionary](#️-dataset--data-dictionary)
- [🧹 Data Cleaning & Metric Definitions](#-data-cleaning--metric-definitions)
- [📗 Part A — Excel Analysis](#-part-a--excel-analysis)
- [🗄️ Part B — SQL Analysis](#️-part-b--sql-analysis)
- [🐍 Part C — Python Analysis](#-part-c--python-analysis)
- [🔗 Cross-Tool Reconciliation](#-cross-tool-reconciliation)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [💡 Recommendation & Limitation](#-recommendation--limitation)
- [🎥 Video Explanation](#-video-explanation)
- [🏆 Advantages](#-advantages)
- [📚 References](#-references)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

**Customer Care Service** is a customer support quality analysis project that studies how quickly support tickets are resolved, how satisfied customers are, and how service quality varies across **support teams, departments, channels, and months**. The same dataset is analysed independently in **Excel**, **SQL**, and **Python**, and the results are cross-checked against each other to make sure they agree.

This project is designed to:
- Clean raw data by detecting and removing an exact duplicate record
- Join a fact table (`tickets`) with a lookup table (`teams`) using `team_id`
- Create a derived field (`breach_flag`) to measure SLA performance
- Compare teams, departments, channels, and months using Excel formulas, SQL queries, and Python code
- Reconcile one aggregate value across all three tools

> **Business Question:** *Which support team should improve resolution performance, and how does service quality vary by channel?*

---

## 🎯 Problem Statement

> **Objective:** Identify which support team needs to improve its resolution performance, and understand how service quality differs across Email, Chat, and Phone channels.

A customer support centre handles tickets through three channels across four teams. A ticket **meets the SLA** if it is resolved within **24 hours**. The analysis answers two business questions:

1. **Which team or department is breaching the SLA most often?**
2. **How does resolution time and customer satisfaction vary by channel and by month?**

| 📂 Dataset | 📄 Type | 🔍 Description |
|------------|---------|----------------|
| `tickets.csv` | Fact table | 13 rows (12 unique + 1 duplicate) of support tickets |
| `teams.csv` | Lookup table | 4 rows mapping each team to its department |

The goal is to demonstrate **practical data analysis skills** using Excel, SQL, and Python on one consistent dataset.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🧹 **Duplicate Removal** | Detects and removes 1 exact duplicate row — 13 rows become 12 |
| 🔗 **Lookup Join** | Links tickets to teams using `team_id` (XLOOKUP / JOIN / `merge`) |
| 🚩 **SLA Breach Flag** | `breach_flag = 1` when `resolution_hours > 24`, otherwise `0` |
| 📊 **Excel PivotTable** | Average resolution hours by department × month (Jan → Feb → Mar) |
| 🗄️ **SQL Queries** | Three labelled queries plus a data-integrity check |
| 🐍 **Python Pipeline** | Load → clean → merge → assert → derive → summarise → plot → export |
| ✅ **Validation Checks** | Row-count assertions and zero-unmatched-key checks in every tool |
| 🔁 **Reproducible** | Relative paths and run-in-order scripts that work on any computer |

---

## 🏗️ Project Structure

```
📦 data-analysis-set-e-YOUR-STUDENT-ID/
│
├── 📁 data/
│   └── 📁 raw/
│       ├── 📄 tickets.csv            ← Raw fact file (13 rows incl. 1 duplicate)
│       └── 📄 teams.csv              ← Lookup file (4 rows)
│
├── 📁 excel/
│   └── 📗 analysis.xlsx              ← Raw, Lookup, Clean, Summary sheets
│
├── 📁 sql/
│   ├── 📄 setup.sql                  ← CREATE TABLE + INSERT (12 + 4 rows)
│   └── 📄 queries.sql                ← S2a, S2b, S2c + integrity check
│
├── 📁 python/
│   └── 🐍 analysis.py                ← Cleaning, merge, derivation, chart, exports
│
├── 📁 outputs/
│   ├── 📄 clean_data.csv             ← Merged 12-row clean dataset
│   ├── 📄 python_summary.csv         ← Department breach-rate summary
│   ├── 🖼️ python_chart.png           ← Monthly average resolution hours chart
│   └── 📁 sql/                       ← Saved SQL query results
│       ├── 📄 s2a_avg_resolution_by_department.csv
│       ├── 📄 s2b_teams_breaching_sla.csv
│       └── 📄 s2c_top_two_channels_by_breach.csv
│
├── 📄 requirements.txt               ← Python packages (pandas, matplotlib)
├── 📄 .gitignore                     ← Excludes environments, caches, credentials
└── 📄 README.md                      ← Project documentation
```

---

## 🔄 Project Workflow

```
Raw Data (tickets.csv + teams.csv)
              │
              ▼
┌─────────────────────────────────┐
│  Verify Raw Files               │  ← 13 ticket rows, 4 team rows
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│  Remove Exact Duplicate         │  ← Ticket 12 appears twice → 12 rows remain
└────────────────┬────────────────┘
                 │
      ┌──────────┼──────────┐
      ▼          ▼          ▼
┌──────────┐ ┌──────────┐ ┌───────────┐
│  EXCEL   │ │   SQL    │ │  PYTHON   │
│ XLOOKUP  │ │ JOIN     │ │ merge()   │
│ IF flag  │ │ GROUP BY │ │ groupby   │
│ COUNTIFS │ │ HAVING   │ │ matplotlib│
│ Pivot    │ │ LIMIT    │ │ to_csv    │
└────┬─────┘ └────┬─────┘ └────┬──────┘
     │            │            │
     └────────────┼────────────┘
                  ▼
┌─────────────────────────────────┐
│  Cross-Tool Reconciliation      │  ← Same value in Excel, SQL & Python
└────────────────┬────────────────┘
                 │
                 ▼
     Findings + Recommendation ✅
```

---

## 🗂️ Dataset & Data Dictionary

### 📄 `tickets.csv` — Fact Table

| Column | Type | Meaning |
|--------|------|---------|
| `ticket_id` | Integer | Unique ticket number |
| `month` | Text (ordered) | Month of the ticket — **Jan → Feb → Mar** |
| `team_id` | Text | Support team key (T1–T4), links to `teams.csv` |
| `channel` | Text | Contact channel: Email, Chat, or Phone |
| `resolution_hours` | Numeric | Hours taken to resolve the ticket |
| `satisfaction` | Numeric | Customer rating on a **1–5** scale |

### 📄 `teams.csv` — Lookup Table

| Column | Type | Meaning |
|--------|------|---------|
| `team_id` | Text | Unique team key (T1–T4) |
| `team` | Text | Team name |
| `department` | Text | Department: Service or Technical |

### 👥 Team Mapping

| team_id | Team | Department |
|---------|------|------------|
| T1 | AccountCare | Service |
| T2 | BillingHelp | Service |
| T3 | AppSupport | Technical |
| T4 | DeviceHelp | Technical |

> One `teams` row maps to many `tickets` rows (**one-to-many** on `team_id`).

---

## 🧹 Data Cleaning & Metric Definitions

### 🧼 Cleaning Steps

| Step | Action | Result |
|------|--------|--------|
| 1️⃣ | Load raw `tickets.csv` and `teams.csv` unchanged | 13 + 4 rows |
| 2️⃣ | Identify the exact duplicate (ticket `12, Mar, T4, Phone, 24, 5`) | 1 duplicate found |
| 3️⃣ | Remove the duplicate | **13 → 12 rows** |
| 4️⃣ | Set data types (integer, numeric, text) | Types confirmed |
| 5️⃣ | Join tickets to teams on `team_id` | 0 unmatched keys |
| 6️⃣ | Add `breach_flag` | 5 breached, 7 within SLA |

### 📐 Metric Definitions

| Metric | Rule |
|--------|------|
| 🚩 **breach_flag** | `1` if `resolution_hours > 24`, else `0` (exactly 24 hours **meets** the SLA) |
| 📉 **SLA Breach Rate** | `Tickets resolved in more than 24 hours ÷ All tickets × 100` |
| ⏱️ **Avg Resolution Hours** | Mean of `resolution_hours` for the selected group |
| 😊 **Avg Satisfaction** | Mean of `satisfaction` (1–5 scale) |

> ⚠️ Rates are always calculated from **underlying counts**, never by averaging subgroup percentages. Numeric results are shown to **two decimal places**, and all tied entities are reported for highest/lowest.

---

## 📗 Part A — Excel Analysis

### 📝 1. Workbook Guide — `excel/analysis.xlsx`

| Sheet | Purpose |
|-------|---------|
| 📄 **Raw** | Original 13-row `tickets.csv`, pasted unchanged |
| 🔍 **Lookup** | The 4-row `teams.csv` |
| 🧹 **Clean** | 12 unique rows, `department` column, `breach_flag` column, before (13) / after (12) counts |
| 📊 **Summary** | COUNTIFS breach table, PivotTable, and column chart |

---

### 🔍 2. Department Lookup (XLOOKUP)

```excel
=XLOOKUP(C2, Lookup!A:A, Lookup!C:C)
```

*Pulls `department` from the Lookup sheet using `team_id`. `INDEX/MATCH` works equally well.*

---

### 🚩 3. Breach Flag

```excel
=IF(E2>24, 1, 0)
```

*Applied to all 12 rows of the Clean sheet.*

---

### 🔢 4. Breached Tickets by Channel (COUNTIFS)

```excel
=COUNTIFS(Clean!D:D, "Chat", Clean!G:G, 1)
```

| Channel | Breached Tickets |
|---------|------------------|
| 📧 Email | 0 |
| 💬 Chat | 3 |
| 📞 Phone | 2 |

---

### 📊 5. PivotTable & Chart

- **Rows:** `department`  |  **Columns:** `month` (ordered Jan → Feb → Mar)  |  **Values:** Average of `resolution_hours`
- A **column chart** with a title, axis labels, and legend is placed on the Summary sheet.

| Department | Jan | Feb | Mar |
|------------|-----|-----|-----|
| Service | 20.00 | 19.00 | 19.00 |
| Technical | 28.00 | 29.00 | 28.00 |

> All formulas and the PivotTable stay **live and editable** — no screenshots-only results.

---

## 🗄️ Part B — SQL Analysis

### 📝 6. Setup — `sql/setup.sql`

- Defines `teams` and `tickets` with primary keys and a **foreign key** `tickets.team_id → teams.team_id`
- Loads exactly **4 team rows** and **12 ticket rows** (the duplicate is excluded)
- Starts with a comment stating the **SQL dialect and version**

```sql
-- Dialect: [Your SQL Engine & Version, e.g. SQLite 3.x / MySQL 8.x]

CREATE TABLE teams (
    team_id    VARCHAR(5)  PRIMARY KEY,
    team       VARCHAR(50) NOT NULL,
    department VARCHAR(50) NOT NULL
);

CREATE TABLE tickets (
    ticket_id        INT          PRIMARY KEY,
    month            VARCHAR(3)   NOT NULL,
    team_id          VARCHAR(5)   NOT NULL,
    channel          VARCHAR(10)  NOT NULL,
    resolution_hours DECIMAL(6,2) NOT NULL,
    satisfaction     INT          NOT NULL,
    FOREIGN KEY (team_id) REFERENCES teams(team_id)
);
```

---

### 📊 7. S2a — Average Resolution Time by Department

```sql
SELECT t.department,
       ROUND(AVG(k.resolution_hours), 2) AS avg_resolution_hours
FROM tickets k
JOIN teams t ON k.team_id = t.team_id
GROUP BY t.department
ORDER BY avg_resolution_hours DESC;
```

**Result:**

| department | avg_resolution_hours |
|------------|----------------------|
| Technical | 28.33 |
| Service | 19.33 |

---

### 🚨 8. S2b — Teams Breaching SLA (Average > 24 Hours)

```sql
SELECT t.team,
       ROUND(AVG(k.resolution_hours), 2) AS avg_resolution_hours
FROM tickets k
JOIN teams t ON k.team_id = t.team_id
GROUP BY t.team
HAVING AVG(k.resolution_hours) > 24
ORDER BY avg_resolution_hours DESC;
```

**Result:**

| team | avg_resolution_hours |
|------|----------------------|
| AppSupport | 28.67 |
| DeviceHelp | 28.00 |
| BillingHelp | 26.67 |

---

### 🏅 9. S2c — Top Two Channels by Breach Count

```sql
SELECT channel,
       COUNT(*) AS breach_count
FROM tickets
WHERE resolution_hours > 24
GROUP BY channel
ORDER BY breach_count DESC, channel ASC
LIMIT 2;
```

**Result:**

| channel | breach_count |
|---------|--------------|
| Chat | 3 |
| Phone | 2 |

---

### 🔎 10. Data Integrity Check (LEFT JOIN)

```sql
SELECT t.team_id, t.team, COUNT(k.ticket_id) AS ticket_count
FROM teams t
LEFT JOIN tickets k ON t.team_id = k.team_id
GROUP BY t.team_id, t.team;
```

*Confirms every `team_id` in the fact table matches a lookup row — **zero unmatched keys**.*

---

### ▶️ 11. SQL Execution Order

| Step | File | Action |
|------|------|--------|
| 1️⃣ | `sql/setup.sql` | Create tables and load 4 + 12 rows |
| 2️⃣ | `sql/queries.sql` | Run S2a → S2b → S2c → integrity check |
| 3️⃣ | `outputs/sql/` | Save each result as a labelled CSV |

> **SQL Dialect & Version:** `[Fill in — e.g. SQLite 3.45 / MySQL 8.0 / PostgreSQL 16]`

---

## 🐍 Part C — Python Analysis

### 📝 12. Setup & Run

```bash
pip install -r requirements.txt
python python/analysis.py
```

*Run from the **repository root** — all file paths are relative, so no edits are needed on another machine.*

---

### 📥 13. Load, Clean & Merge

```python
import pandas as pd

tickets = pd.read_csv("data/raw/tickets.csv")
teams   = pd.read_csv("data/raw/teams.csv")

tickets = tickets.drop_duplicates()                       # 13 → 12 rows
df = tickets.merge(teams, on="team_id", how="left")

assert len(df) == 12
assert df["department"].isna().sum() == 0                 # no unmatched keys
```

---

### 🚩 14. Derived Field & Department Summary

```python
df["breach_flag"] = (df["resolution_hours"] > 24).astype(int)

dept = df.groupby("department").agg(
    total_tickets=("ticket_id", "count"),
    breached=("breach_flag", "sum")
).reset_index()
dept["sla_breach_rate_pct"] = (dept["breached"] / dept["total_tickets"] * 100).round(2)
```

**Department Summary:**

| department | total_tickets | breached | sla_breach_rate_pct |
|------------|---------------|----------|---------------------|
| Service | 6 | 2 | 33.33 |
| Technical | 6 | 3 | 50.00 |

**Team with Highest Breach Rate:**

| team | breached (numerator) | total tickets (denominator) | breach rate |
|------|----------------------|-----------------------------|-------------|
| BillingHelp (T2) | 2 | 3 | 66.67% |
| AppSupport (T3) | 2 | 3 | 66.67% |

> 🤝 Two teams are **tied** for the highest breach rate, so both are reported.

---

### 📈 15. Chart & Exports

```python
import matplotlib.pyplot as plt

order = ["Jan", "Feb", "Mar"]
monthly = df.groupby("month")["resolution_hours"].mean().reindex(order)

plt.bar(monthly.index, monthly.values)
plt.title("Monthly Average Resolution Hours")
plt.xlabel("Month")
plt.ylabel("Average Resolution Hours")
plt.savefig("outputs/python_chart.png", dpi=150, bbox_inches="tight")

df.to_csv("outputs/clean_data.csv", index=False)
dept.to_csv("outputs/python_summary.csv", index=False)
```

| Month | Avg Resolution Hours |
|-------|----------------------|
| Jan | 24.00 |
| Feb | 24.00 |
| Mar | 23.50 |

---

## 🔗 Cross-Tool Reconciliation

**Aggregate checked:** Average resolution hours for the **Technical** department.

| Tool | Method | Result |
|------|--------|--------|
| 📗 **Excel** | PivotTable / `AVERAGEIFS` on the Clean sheet | 28.33 |
| 🗄️ **SQL** | S2a query | 28.33 |
| 🐍 **Python** | Department `groupby` summary | 28.33 |

**Second check — overall SLA breach rate:** 5 breached ÷ 12 tickets = **41.67%** in Excel (`COUNTIFS`), SQL (`COUNT … WHERE resolution_hours > 24`), and Python (`breach_flag.sum() / len(df)`).

> ✅ All three tools agree. Rounding note: `[Add any rounding differences you notice, or write "No rounding differences — values match to 2 decimals"]`.

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📗 **Microsoft Excel** | `[365 / 2019+]` | XLOOKUP, IF, COUNTIFS, PivotTable, charts |
| 🗄️ **SQL** | `[Engine & version]` | Table design, JOIN, GROUP BY, HAVING, LIMIT |
| 🐍 **Python** | `[3.8+]` | Core scripting language |
| 🐼 **pandas** | `[version]` | Data loading, cleaning, merging, aggregation |
| 📊 **matplotlib** | `[version]` | Charting |
| 🔀 **Git / GitHub** | Latest | Version control and submission |

---

## 📈 Results & Insights

After running the analysis, the following findings are produced:

- 🧹 **Clean Dataset** — 13 raw rows reduced to 12 unique records
- 🚩 **Overall SLA Breach Rate** — **41.67%** (5 of 12 tickets took more than 24 hours)
- 🏢 **Department Gap** — Technical has a **50.00%** breach rate (3 of 6) and a **28.33-hour** average, versus **33.33%** (2 of 6) and **19.33 hours** for Service
- 👥 **Weakest Teams** — BillingHelp and AppSupport are tied at **66.67%** breach rate (2 of 3 tickets each); AppSupport has the longest average at **28.67 hours**
- ✅ **Best Team** — AccountCare has **0 breaches** and the fastest average of **12.00 hours**
- 📡 **Channel Quality** — Chat has the most breaches (**3**) and a 27.00-hour average; Email has **0 breaches**, an 18.00-hour average, and the highest average satisfaction (**4.00**)
- 📅 **Monthly Trend** — Average resolution time is steady: **24.00** (Jan), **24.00** (Feb), **23.50** (Mar)
- 😊 **Satisfaction vs SLA** — Breached tickets average **2.60** satisfaction compared with **4.29** for tickets within SLA

---

## 💡 Recommendation & Limitation

### 🎯 Recommendation

Prioritise **resolution-time improvement for the Technical department**, starting with **AppSupport** and **BillingHelp** (both 66.67% breach rate), and review **Chat** workflows, since Chat has the most SLA breaches. Email and AccountCare processes can serve as an internal benchmark, as they show no breaches.

### ⚠️ Limitation

The dataset is **synthetic and very small** (12 tickets, 3 per team), so one ticket changes a team's breach rate by 33 percentage points. Findings show direction, not statistical proof.


---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Beginner Friendly** | Applies Excel, SQL, and Python to one simple, easy-to-follow dataset |
| 🔁 **Reproducible** | Relative paths and ordered run steps work on any computer |
| ✅ **Validated** | Row-count assertions and zero-unmatched-key checks in each tool |
| 🔗 **Cross-Verified** | Key results match across Excel, SQL, and Python |
| 🧮 **Count-Based Rates** | Breach rates come from underlying counts, not averaged percentages |
| 🧩 **Modular** | Each tool lives in its own folder and can be run independently |
| 📖 **Well Documented** | Data dictionary, metric definitions, and run steps included |
| 🧪 **Extensible** | Easy to add more months, teams, or SLA thresholds |

---

## 📚 References


> **Authorship Declaration:** *All work in this repository is my own except where cited.*

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Sneha Gupta

> *"Data tells the story — analysis helps us listen."*

**🎓 Role:** BBA Student | Data Analysis Learner\
**🆔 Student ID:** `12827` |
**📍 Location:** India\
**🛠️ Skills:** Excel · SQL · Python · Data Cleaning · Data Analysis

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 🏫 **Red & White Skill Education** — For the practical exam and dataset
- 📚 [Python Official Docs](https://docs.python.org/3/) — Official Python language reference
- 🐼 [pandas Documentation](https://pandas.pydata.org/docs/) — Data analysis library reference
- 📊 [matplotlib Documentation](https://matplotlib.org/stable/index.html) — Plotting reference
- 🗄️ [W3Schools SQL](https://www.w3schools.com/sql/) — Beginner SQL reference
- 📗 [Microsoft Excel Support](https://support.microsoft.com/excel) — Excel formula and PivotTable help
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---

*Made with ❤️ and ☕ — Last updated: 3 October, 2026*

</div>
