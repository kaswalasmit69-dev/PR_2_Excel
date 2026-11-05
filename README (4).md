<div align="center">

# 📊 PR. 1 — FUNDAMENTAL BOOSTER

### ✦ Excel Data Analysis · Formula Automation · Dynamic Reporting ✦

**An Academic Practical Project in Microsoft Excel**

<br>

![Project](https://img.shields.io/badge/PROJECT-FUNDAMENTAL%20BOOSTER-1d4ed8?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Status](https://img.shields.io/badge/STATUS-VIDEO%20LINK%20PENDING-0f766e?style=for-the-badge)
![Course](https://img.shields.io/badge/FOCUS-DATA%20ANALYSIS-7c3aed?style=for-the-badge)

<br>

**👨‍🎓 Sumit Kaswala** &nbsp; • &nbsp; **🎓 B.Tech — Final Year**  
**📚 Data Analysis Course** &nbsp; • &nbsp; **👨‍🏫 Faculty Guide: Prof. Girish Gondaliya**

<br>

### [▶️ WATCH THE PROJECT VIDEO PRESENTATION](PASTE_YOUR_NEW_VIDEO_LINK_HERE)

*Click above to open the full video demonstration in Google Drive.*

---

**📘 Documentation** · **🧮 Formula Implementation** · **📈 Practical Analysis** · **🎥 Video Walkthrough**

</div>

> [!IMPORTANT]
> **Evaluator quick access:** [Open the project demonstration video](PASTE_YOUR_NEW_VIDEO_LINK_HERE). **Video link pending:** Replace `PASTE_YOUR_NEW_VIDEO_LINK_HERE` with your own video URL and enable viewer access. The Excel workbook should be submitted alongside this README.

## 🧭 Table of Contents

- [✨ Project Overview](#-project-overview)
- [🎯 Project Objectives](#-project-objectives)
- [🗂️ Workbook Architecture](#️-workbook-architecture)
- [🔍 Worksheet-by-Worksheet Breakdown](#-worksheet-by-worksheet-breakdown)
- [🧠 Excel Formula Toolkit](#-excel-formula-toolkit)
- [🧪 Practical Demonstrations](#-practical-demonstrations)
- [🎬 Video Presentation](#-video-presentation)
- [🚀 How to Review](#-how-to-review)
- [✅ Submission Checklist](#-submission-checklist)
- [🏆 Learning Outcomes](#-learning-outcomes)
- [👨‍🎓 Academic Details](#-academic-details)

---

## ✨ Project Overview

**PR. 1 — Fundamental Booster** is a structured Excel project demonstrating how raw data can be transformed into meaningful insights using **calculations, logical conditions, lookups, text processing, date functions, and dynamic references**.

The project combines realistic academic and business-style datasets with interactive formula demonstrations. Its focus is not just to display results, but to show **how those results are calculated and how they respond when inputs change**.

### 📌 Project Snapshot

| 📍 Metric | 📊 Project Scope |
|:--|:--|
| 🗂️ Total worksheets | **7** |
| 👨‍🎓 Student records | **20** |
| 🛍️ Sales transactions | **10** |
| 👥 Employee records | **10** |
| 📈 Salespersons analyzed | **10** |
| 🧩 Main capabilities | Conditional logic, multi-criteria analysis, lookups, dynamic references |
| 🎥 Presentation | [Add your video link here](PASTE_YOUR_NEW_VIDEO_LINK_HERE) |

### 🔄 Analysis Workflow

```text
   DATA ENTRY          EXCEL FORMULAS           INSIGHTS & OUTPUTS
       │                     │                         │
       ▼                     ▼                         ▼
 Student / Sales  ──►  IF · SUMIFS · XLOOKUP  ──►  Grades / Totals
 Employee / Text  ──►  INDEX · MATCH · TEXT   ──►  Search / Cleaning
 Monthly Trends   ──►  INDIRECT · OFFSET     ──►  Dynamic Analysis
```

---

## 🎯 Project Objectives

- 🧮 **Automate calculations** instead of calculating results manually.
- 🧠 **Apply logical rules** using `IF`, nested `IF`, `AND`, and `OR`.
- 📊 **Analyze data with multiple criteria** using `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS`.
- 🔎 **Retrieve accurate records** using `VLOOKUP`, `MATCH`, and `INDEX` + `MATCH`.
- 🧹 **Clean and standardize text** with Excel text functions.
- 📅 **Calculate age and experience** using date functions.
- 🔁 **Build responsive calculations** using relative/absolute references, `INDIRECT`, and `OFFSET`.

---

## 🗂️ Workbook Architecture

| # | 📑 Worksheet | 🎯 Purpose | ⚙️ Core Functions |
|:--:|:--|:--|:--|
| 01 | **Project Dashboard** | Introduction and workbook overview | Structured navigation and project scope |
| 02 | **Student Data** | Student marks, grades, eligibility and analysis | `SUM` · `AVERAGE` · `IF` · `COUNTIFS` · `AVERAGEIFS` · `VLOOKUP` · `INDEX` + `AGGREGATE` |
| 03 | **Sales Data** | Sales, discounts and product search | `IF` · `OR` · `SUMIFS` · `VLOOKUP` · `MATCH` · rounding functions |
| 04 | **Employee Data** | Employee age, experience and search | `TODAY` · `DATEDIF` · `INDEX` · `MATCH` |
| 05 | **Salesperson Performance** | Monthly sales and performance comparison | `SUM` · `AVERAGE` · nested `IF` · `INDEX` + `MATCH` |
| 06 | **Text Functions** | Text extraction and standardization | `LEFT` · `RIGHT` · `FIND` · `LEN` · `UPPER` · `LOWER` |
| 07 | **Dynamic References** | Fixed references and rolling sales analysis | `$` references · `INDIRECT` · `OFFSET` |

---

## 🔍 Worksheet-by-Worksheet Breakdown

### 01 · 🏠 Project Dashboard

The starting point of the workbook. Introduces the project, the practical tasks and the Excel concepts covered.

### 02 · 🎓 Student Data — Academic Performance Analysis

**Dataset:** 20 students, with Math, Science and English marks.

**Implemented features:** Total and average marks, nested-IF grading, AND-based eligibility checks, conditional counts, conditional averages, student-ID lookup and Excel 2016-compatible top-performer extraction.

| 📌 Analysis | 📈 Result |
|:--|--:|
| Math score above 50 | **16 students** |
| Average score above 60 | **Calculated dynamically in workbook** |
| Math **and** Science above 60 | **Calculated dynamically in workbook** |
| Average of qualifying scores above 60 | **Calculated dynamically in workbook** |
| Above 60 in all three subjects | **Calculated dynamically in workbook** |

**Key idea:** Criteria-based formulas generate summaries automatically when source marks are updated.

### 03 · 🛒 Sales Data — Transaction & Discount Analysis

**Dataset:** 10 sales records with product codes, regions, salespersons, quantities and unit prices.

**Business rule:** Sales of **₹40,000 or more → 10% discount**; otherwise **5% discount**.

**Implemented features:** Gross sales, discount, final amount, eligibility, region-and-product `SUMIFS`, product-price `VLOOKUP`, product-position `MATCH` (Excel 2016 compatible), and numeric rounding.

| 🔎 Lookup | ⌨️ Input | ✅ Expected Result |
|:--|:--|:--|
| `VLOOKUP` | `Q211` | **₹58,000** unit price |
| `MATCH` | `Q214` | **4** (position within product lookup list) |

**Actual formulas demonstrated:**

```excel
=VLOOKUP(N14,$N$8:$O$12,2,FALSE)
=MATCH(N17,$N$8:$N$12,0)
```

> **Why this matters:** The product code is matched against column **C**, and the exact-match lookup returns the corresponding value or position. `$` locks the lookup range when copying a formula.

### 04 · 👥 Employee Data — Workforce Information

**Dataset:** 10 employees, with department, designation, salary, date of birth and date of joining.

**Implemented features:** Age, years of experience, total service days, salary retrieval by employee ID and dynamic name/department/designation lookup.

**Key idea:** `MATCH` identifies a record's position; `INDEX` returns the required field from that position.

### 05 · 📈 Salesperson Performance — Monthly Analysis

**Dataset:** 10 salespersons across January, February, March and April.

**Implemented features:** Total sales, monthly average, performance labels, salesperson lookup and month-specific sales retrieval.

| 👤 Salesperson | 🗓️ Month | 💰 Sales |
|:--|:--|--:|
| Pihu | March | **(calculated in workbook)** |
| Pihu | April | **(calculated in workbook)** |

**Key idea:** Changing the selected month updates the matching result without rewriting the lookup formula.

### 06 · 🔤 Text Functions — Data Preparation

**Implemented features:** Extracting first/last names, converting text to uppercase/lowercase and counting characters.

| ⚙️ Function | 📝 Application |
|:--|:--|
| `LEFT` + `FIND` | Extract first name |
| `RIGHT` + `LEN` + `FIND` | Extract last name |
| `UPPER` | Convert text to uppercase |
| `LOWER` | Convert text to lowercase |
| `LEN` | Count characters, including spaces |

### 07 · 🔁 Dynamic References — Responsive Calculations

**Dataset:** January–December monthly sales.

**Implemented features:** Relative references, absolute references for a fixed **7% bonus rate**, `INDIRECT` cell lookup and `OFFSET`-based recent-month calculations.

**Live scenario:** Change the recent-month input from **3** to **6** to update the total and average for the selected rolling period.

---

## 🧠 Excel Formula Toolkit

| 🧩 Category | ⚙️ Functions | 💡 What They Do |
|:--|:--|:--|
| Arithmetic | `SUM`, `AVERAGE` | Calculate totals and means |
| Logical | `IF`, nested `IF`, `AND`, `OR` | Evaluate conditions and make decisions |
| Multi-criteria | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` | Analyze matching records |
| Lookup | `VLOOKUP`, `MATCH` | Find values and positions |
| Dynamic lookup | `INDEX` + `MATCH` | Return data from matched positions |
| Text | `LEFT`, `RIGHT`, `FIND`, `LEN`, `UPPER`, `LOWER` | Clean and transform strings |
| Date | `TODAY`, `DATEDIF` | Calculate age and tenure |
| Dynamic arrays | `INDEX` + `AGGREGATE` | Return qualifying rows in Excel 2016 |
| References | Relative, absolute, `INDIRECT`, `OFFSET` | Create reusable and responsive formulas |
| Math | `ROUND`, `CEILING`, `FLOOR` | Round numeric results |

### 💻 Selected Formula Patterns

```excel
=SUM(C4:E4)
=COUNTIFS(C4:C23,">50")
=SUMIFS(sum_range,region_range,selected_region,product_range,selected_product)
=INDEX(salary_range,MATCH(employee_id,id_range,0))
=INDEX(return_range,MATCH(lookup_value,lookup_range,0))
=IFERROR(INDEX(data_range,AGGREGATE(15,6,row_positions/(condition),k)),"")
=INDIRECT(cell_containing_address)
=SUM(OFFSET(start_cell,rows_offset,0,number_of_rows,1))
```

*These are learning examples, not a claim that every sample cell reference is used verbatim in the submitted workbook. The workbook itself contains the actual formulas.*

---

## 🧪 Practical Demonstrations

| 🎯 Demo | 🖱️ Action | 👀 Expected Observation |
|:--|:--|:--|
| Student analysis | Select conditional-summary cells | Counts and averages reflect criteria |
| Product lookup | Search `Q211` | Price **₹58,000** |
| Product position | Search `Q214` | First matching position **4** |
| Employee lookup | Change employee ID | Corresponding employee details update |
| Salesperson analysis | Change the selected salesperson or month | **updated result** |
| Dynamic period | Change recent months from 3 to 6 | Total and average recalculate |

> [!TIP]
> While reviewing, select a calculated cell and inspect the **Formula Bar** to understand the underlying logic.

---

## 🎬 Video Presentation

<div align="center">

### ▶️ [CLICK HERE TO WATCH THE COMPLETE VIDEO](PASTE_YOUR_NEW_VIDEO_LINK_HERE)

**📹 Format:** MP4 &nbsp; | &nbsp; **🗣️ Explanation:** Hinglish &nbsp; | &nbsp; **📚 Content:** Workbook walkthrough

[![Watch Video](https://img.shields.io/badge/▶%20WATCH%20PROJECT%20DEMO-GOOGLE%20DRIVE-e11d48?style=for-the-badge&logo=googledrive&logoColor=white)](PASTE_YOUR_NEW_VIDEO_LINK_HERE)

</div>

**Presentation sequence:** Introduction → Student Data → Sales Data → Employee Data → Salesperson Performance → Text Functions → Dynamic References → Conclusion.

> [!NOTE]
> **Video URL placeholder:** Replace `PASTE_YOUR_NEW_VIDEO_LINK_HERE` everywhere with the final uploaded video link. If using Google Drive, set **General access → Anyone with the link → Viewer**.

---

## 🚀 How to Review

1. 📂 Open the submitted **Excel workbook**.
2. 🏠 Start with **Project Dashboard** to understand the workbook structure.
3. 📑 Open each worksheet in the order shown above.
4. 🔍 Select formula cells and inspect the **Formula Bar**.
5. 🔁 Change a lookup input or selected month to observe the dynamic result.
6. ▶️ Open the **[video presentation](PASTE_YOUR_NEW_VIDEO_LINK_HERE)** for the guided walkthrough.

**💻 Compatibility:** This project uses Excel 2016-compatible `INDEX` + `MATCH` and `INDEX` + `AGGREGATE` in place of newer `XLOOKUP`, `XMATCH`, and `FILTER` functions.

---

## ✅ Submission Checklist

- [x] Seven-sheet workbook structure documented
- [x] Core formulas and their purposes explained
- [x] Practical lookup examples included
- [ ] Replace the placeholder with your own video presentation link
- [ ] Configure your video sharing as **Anyone with the link → Viewer**
- [ ] Confirm the video opens in an incognito/private window
- [ ] Recheck the final Excel workbook before submission
- [ ] Submit the **Excel workbook + README.md** together

---

## 🏆 Learning Outcomes

Through this project, I developed practical skills in **spreadsheet organization, formula-based automation, conditional analysis, exact-match searching, text cleaning, date calculations and dynamic reporting**.

The project also strengthened my ability to **explain technical work clearly**, verify formula outputs and present structured data analysis in an academic setting.

---

## 👨‍🎓 Academic Details

| 🏷️ Field | 📋 Information |
|:--|:--|
| **Student Name** | **Sumit Kaswala** |
| **Academic Program** | Bachelor of Technology (B.Tech) — Final Year |
| **Course Focus** | Data Analysis |
| **Project Title** | PR. 1 — Fundamental Booster (New Records Edition) |
| **Faculty Guide** | **Prof. Girish Gondaliya** |
| **Tools** | Microsoft Excel · Markdown · Video hosting (to be added) |
| **Deliverables** | Excel Workbook · README Documentation · Video Demonstration |

---

<div align="center">

### 🌟 Thank You for Reviewing My Project!

**Designed, documented and presented by Sumit Kaswala**  
*PR. 1 — Fundamental Booster | Excel Data Analysis*

[⬆️ Back to Top](#-pr-1--fundamental-booster)

</div>
