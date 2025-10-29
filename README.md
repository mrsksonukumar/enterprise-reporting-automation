# Enterprise Reporting Automation System

[![Production](https://img.shields.io/badge/Status-Production-success)](https://github.com/mrsksonukumar/enterprise-reporting-automation)
[![Python](https://img.shields.io/badge/Python-3.8+-blue)](https://www.python.org/)
[![Monthly](https://img.shields.io/badge/Execution-Monthly-orange)](https://github.com/mrsksonukumar/enterprise-reporting-automation)

> Production-grade ETL and reporting automation system reducing monthly reporting time by **75%** for **105+ suppliers** at Tesco Property Insights

## 🚀 Project Overview

**Challenge:** Manual generation of supplier KPI reports was taking 3-4 hours monthly, processing data for 105+ suppliers with multiple KPI categories, requiring extensive manual Excel work and email distribution.

**Solution:** Built end-to-end automation system that extracts data from Oracle database, generates customized Excel reports with business logic, and creates draft emails—all in under 1 hour.

**Impact:**
- ⏱️ **75% time reduction** (3-4 hours → 1 hour)
- 📊 **105+ suppliers** served monthly
- 🔄 **6 KPI categories** automated
- 💾 **30-40 hours** saved annually
- ✅ **Zero errors** in production since deployment

---

## 💡 Technical Architecture

### System Components

┌─────────────────┐
│ Oracle DB │ ← SQL Queries (6 KPI Categories)
│ (Production) │
└────────┬────────┘


│ ▼ ┌
────────────────┐ │ Data Pipeline │ ←
pandas DataFrames │ - Validation │ -
ategory Filtering │ - Transform │
Target Matching └
─
─
───┬────────┘
│ ▼ ┌─────────────────┐ │ Excel E
gine │ ← openpyxl (10-15s/report) │ -
Template │ - Fast Bulk Write │
- Formatting │


Multi-sheet └──────
─┬────────┘ │ ▼ ┌─────
───────────┐ │ Email Service │ ← Ou
look Automation │ - Draft Create │


---

## 🛠️ Key Features

### 1. **Intelligent Data Extraction**
- Parameterized SQL queries for 6 KPI categories
- Dynamic period detection from control table
- Category-based filtering and validation
- Automatic target data retrieval

### 2. **High-Performance Excel Generation**
- Fast bulk-write method (10-15 seconds per report)
- Template-based approach maintaining formatting
- Multi-sheet handling (Time KPIs, Compliance, Inspections, etc.)
- Automatic cell clearing and data refresh

### 3. **Fault-Tolerant Architecture**
- JSON-based checkpoint system
- Resume capability after interruptions
- Comprehensive logging at all levels
- Progress tracking with ETA calculations

### 4. **Email Automation**
- Automated Outlook draft creation
- Configurable To/CC distribution lists
- Attachment handling for 105+ files
- Placeholder email support for testing

### 5. **Production-Ready Design**
- Configuration-driven (JSON config file)
- Modular architecture (core, generators, utils)
- Data validation with schema checking
- Error recovery and graceful degradation

---

## 📊 Performance Metrics

| Metric | Value |
|--------|-------|
| **Total Suppliers** | 105+ |
| **Processing Time** | ~60 minutes |
| **Time per Supplier** | ~30-40 seconds |
| **Excel Generation** | 10-15 seconds/report |
| **Success Rate** | 99%+ |
| **Manual Time Saved** | 30-40 hours/year |
| **Execution Frequency** | Monthly |

---

## 🏗️ Technical Stack

**Core Technologies:**
- **Python 3.8+** - Main automation logic
- **pandas** - Data manipulation and transformation
- **openpyxl** - Fast Excel file generation
- **cx_Oracle** - Database connectivity

**Database:**
- **Oracle SQL** - Production data warehouse
- Complex joins across supplier, KPI, and target tables
- Parameterized queries for security

**Libraries & Tools:**
- `json` - Configuration and checkpointing
- `logging` - Comprehensive tracking
- `pathlib` - Cross-platform file handling
- `win32com` - Outlook automation (Windows)

---

## 📁 Project Structure

enterprise-reporting-automation/
├── core/
│ ├── logger.py # Centralized logging
│ ├── database.py # Oracle DB connection
│ ├── validator.py # Data validation
│ └── init.py
├── generators/
│ ├── excel_generator.py # Fast Excel creation
│ ├── email_generator.py # Outlook automation
│ └── init.py
├── sql/
│ └── queries.py # SQL query definitions
├── utils/
│ ├── file_utils.py # File operations
│ ├── data_utils.py # Data helpers
│ └── init.py
├── config/
│ ├── config.json # System configuration
│ └── schema.json # Data validation rules
├── main.py # Production entry point
└── README.md

---

## 🔧 Key Algorithms & Optimizations

### Fast Excel Generation

Original: 60-90s per report (cell-by-cell)
Optimized: 10-15s per report (bulk write)
for r_idx, row_data in enumerate(df.values, 2):
for c_idx, value in enumerate(row_data, 1):
sheet.cell(row=r_idx, column=c_idx, value=value)

**Result:** 6x performance improvement

### Progress Checkpointing
{
"completed": [101, 102, 103, ...],
"updated_at": "2024-10-29T10:30:00",
"total_completed": 45
}

**Benefit:** Resume from any point, no duplicate work

### Multi-Method Error Handling
- Database connection retry logic
- Excel generation fallback strategies
- Email creation error recovery
- Graceful degradation throughout

---

## 📈 Business Impact

### Quantifiable Benefits
- **Time Savings:** 2-3 hours per month = 30-40 hours annually
- **Consistency:** 100% standardized reporting format
- **Scalability:** Can handle 200+ suppliers without code changes
- **Reliability:** 99%+ success rate in production
- **Audit Trail:** Complete logging for compliance

### Operational Improvements
- ✅ Freed analyst time for high-value work
- ✅ Eliminated manual copy-paste errors
- ✅ Enabled consistent period-over-period analysis
- ✅ Improved supplier communication timeliness
- ✅ Reduced email distribution effort

---

## 🎯 Problem-Solving Approach

### Challenges Overcome

**1. Performance Bottleneck**
- **Problem:** Initial cell-by-cell writing took 60-90s per report
- **Solution:** Implemented bulk-write optimization
- **Result:** 6x faster (10-15s per report)

**2. State Management**
- **Problem:** Crashes required full restart (3-4 hours wasted)
- **Solution:** JSON checkpoint system with resume capability
- **Result:** Zero time wasted on re-processing

**3. Data Quality**
- **Problem:** Inconsistent category naming across sources
- **Solution:** Normalization layer with validation
- **Result:** 100% data accuracy

**4. Email Distribution**
- **Problem:** Manual email creation for 105+ suppliers
- **Solution:** Automated Outlook draft creation
- **Result:** One-click review and send

---

## 🏆 Recognition

- ✅ **Running in production** since 2024
- ✅ **Monthly execution** supporting business operations
- ✅ **Zero post-deployment issues**
- ✅ **Adopted as standard practice** for similar automation needs

---

## 💼 Skills Demonstrated

**Data Engineering:**
- ETL pipeline design and implementation
- Database query optimization
- Data validation and quality control
- Schema design and normalization

**Software Engineering:**
- Modular architecture design
- Error handling and fault tolerance
- Logging and observability
- Configuration management

**Business Analysis:**
- Requirements gathering from stakeholders
- Process optimization and workflow design
- KPI definition and measurement
- Stakeholder communication

**DevOps/Production:**
- Production system deployment
- Checkpoint/resume pattern implementation
- Performance monitoring and optimization
- Documentation for maintenance

---

## 📝 Note on Code

Due to proprietary nature and company IP considerations, the full implementation code is not publicly available. This README demonstrates the technical approach, architecture, and business impact of the solution.

**Architecture patterns and approaches shown here can be applied to:**
- Automated reporting systems
- ETL pipelines with email distribution
- High-volume data processing tasks
- Enterprise automation projects

For technical discussions or similar project inquiries, please connect via LinkedIn.

---

## 🔗 Related Projects

- [Enterprise Certificate Scraper](https://github.com/mrsksonukumar/enterprise-automation-case-study) - Award-winning web scraping automation (70K+ documents)

---

**Project Context:** Tesco Property Insights Team  
**Role:** Data Analyst  
**Period:** 2024 - Present  
**Status:** Production (Monthly Execution)  
**Author:** Sonu Kumar

[![LinkedIn](https://img.shields.io/badge/-Connect_on_LinkedIn-0077B5?style=flat&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/mrsksonukumar)
[![GitHub](https://img.shields.io/badge/-GitHub_Profile-181717?style=flat&logo=github&logoColor=white)](https://github.com/mrsksonukumar)
[![Website](https://img.shields.io/badge/-Portfolio-4CAF50?style=flat&logo=google-chrome&logoColor=white)](https://mrsksonukumar.github.io)
