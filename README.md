# 📊 Excel VBA Expense Dashboard

An interactive and automated **Expense Management Dashboard built in Microsoft Excel using VBA**.

The project transforms raw expense data into an interactive dashboard containing KPI cards, PivotTables, charts and slicers. The VBA automation prepares the data, builds the analytical layer and generates the dashboard with minimal manual intervention.

<img width="1562" height="638" alt="expenses-dashboard" src="https://github.com/user-attachments/assets/87752873-a3bd-4594-8de9-2c0349e5bbee" />


---

## 🎯 Project Objective

The objective of this project was to automate the process of converting expense transaction data into a professional management dashboard.

Instead of manually creating PivotTables, charts, KPI calculations and dashboard elements, the VBA solution automates the complete dashboard-building workflow through a single public macro:

`BuildExpensesDashboard`

The VBA architecture separates data preparation, PivotTable creation, chart helper generation, dashboard construction, chart formatting and slicer creation.

---

## 📌 Dashboard Overview

The dashboard provides a high-level view of expense performance and allows users to analyse spending across different dimensions.

### Key KPIs

* **Total Expenses**
* **Budget Allocated**
* **Budget Remaining**
* **Total Transactions**
* **Approval Rate**
* **Average Expense**

The KPI values are linked to PivotTable-based calculations and are designed to respond to dashboard filters.

---

## 📈 Dashboard Visualizations

The dashboard includes five main visualizations:

### 1. Expense by Department

Shows total expenses across departments such as:

* Engineering
* Finance
* HR
* Marketing
* Operations
* Sales

### 2. Expense by Category

A doughnut chart showing the distribution of expenses across categories.

### 3. Monthly Trend

A line chart showing how expenses change throughout the year.

### 4. Budget vs Actual

Compares actual spending against allocated budgets for each department.

### 5. Payment Method

Shows expense distribution across different payment methods.

## The VBA project creates these five charts automatically from dedicated helper ranges generated from PivotTable data.

## 🎛️ Interactive Filters

The dashboard includes interactive slicers for:

* **Month**
* **Department**

## The slicers are connected to the dashboard's PivotTables, allowing users to filter the analysis dynamically.

## ⚙️ Automation Features

The VBA automation performs the following tasks:

1. Validates required data columns
2. Adds the Month field when required
3. Converts the source data into an Excel Table
4. Applies data formatting
5. Creates the PivotTable layer
6. Builds seven PivotTables
7. Generates chart-helper ranges
8. Creates the dashboard sheet
9. Generates KPI cards
10. Creates dashboard charts
11. Adds Month and Department slicers
12. Applies dashboard formatting
13. Refreshes the workbook data

The source data is expected to contain fields including Expense ID, Date, Department, Category, Amount (INR), Payment Method, Approved, Budget Allocated and Budget Remaining.

---

## 🧩 PivotTable Architecture

The analytical layer contains seven PivotTables:

| PivotTable            | Purpose                     |
| --------------------- | --------------------------- |
| Expense by Department | Department-level spending   |
| Expense by Category   | Category-level spending     |
| Monthly Trend         | Monthly expense analysis    |
| Approved Status       | Approval analysis           |
| KPI                   | Dashboard KPI calculations  |
| Payment Method        | Payment-method analysis     |
| Budget vs Actual      | Budget performance analysis |

The VBA build process creates these PivotTables sequentially from a common PivotCache.

---

## 🛠️ Tools & Technologies

* **Microsoft Excel**
* **VBA (Visual Basic for Applications)**
* **PivotTables**
* **PivotCharts**
* **Excel Slicers**
* **Excel Tables**
* **Excel Formulas**
* **Dashboard Design**
* **Data Analysis**
* **Process Automation**

---

## 📂 Project Structure

```text
excel-vba-expense-dashboard/
│
├── README.md
├── Expenses_Dashboard.xlsm
│
├── VBA/
│   └── BuildExpensesDashboard.bas
│
├── expenses-dashboard.png
│   
│
└── Sample_Data/
    └── expenses_sample.xlsx
```

---

## 🚀 How to Use

### Step 1 — Open the workbook

Open:

`Expenses_Dashboard.xlsm`

in **Microsoft Excel for Windows** with macros enabled.

### Step 2 — Prepare the Data sheet

Create or provide a sheet named:

`Data`

The required headers are:

```text
Expense ID
Date
Department
Category
Amount (INR)
Payment Method
Approved
Budget Allocated
Budget Remaining
```

The VBA process can automatically add the `Month` column when it is missing.

### Step 3 — Run the macro

Run:

```text
BuildExpensesDashboard
```

The macro prepares the data, builds the PivotTables and generates the dashboard.

### Step 4 — Interact with the dashboard

Use the:

* Month slicer
* Department slicer

to filter the dashboard and analyse the expense data.

---

## 💡 Business Questions Answered

This dashboard can help answer questions such as:

* Which department has the highest expenses?
* How much of the allocated budget has been spent?
* How much budget remains?
* Which expense categories contribute most to spending?
* How are expenses trending month by month?
* Which payment method accounts for the highest spending?
* What percentage of transactions have been approved?
* What is the average expense per transaction?
* Which departments have the largest gap between budget and actual spending?

---

## 🔄 VBA Build Workflow

```text
Raw Expense Data
       ↓
Data Validation
       ↓
Excel Table
       ↓
PivotCache
       ↓
7 PivotTables
       ↓
Chart Helper Data
       ↓
KPI Calculations
       ↓
Dashboard
       ↓
Charts + Slicers
       ↓
Interactive Expense Analysis
```

The VBA project uses `BuildExpensesDashboard` as the main entry point and separates the workflow into dedicated procedures for data preparation, PivotTables, dashboard creation, charts and slicers.

---

## 📊 Skills Demonstrated

This project demonstrates practical experience in:

**Excel**

* Advanced Excel
* PivotTables
* PivotCharts
* Slicers
* Excel Tables
* Dashboard development
* KPI reporting

**VBA**

* Macro automation
* Modular VBA development
* Worksheet and range manipulation
* Chart automation
* PivotTable automation
* Slicer automation
* Error handling
* Dynamic formatting

**Data Analytics**

* KPI development
* Budget vs Actual analysis
* Trend analysis
* Department analysis
* Category analysis
* Transaction analysis

**Process Improvement**

* Reducing repetitive manual dashboard-building tasks
* Standardizing reporting workflows
* Automating data preparation and visualization

---

## 🤖 AI-Assisted Development

AI tools were used as a development assistant during the VBA implementation process.

The solution was reviewed and tested with focus on understanding the underlying Excel/VBA logic, dashboard architecture and automation workflow.

The purpose of using AI was to accelerate development while maintaining understanding and control over the final solution.

---

## 📷 Dashboard Preview

<img width="1562" height="638" alt="expenses-dashboard" src="https://github.com/user-attachments/assets/849b6f17-3fe6-44c0-9773-97d94173e5f0" />


---

## 👤 Author

**Ashim Barman**

Interested in **Data Analytics | Business Analysis | Operations Analytics | Excel Automation | Power BI | SQL**

---

## ⭐ Project Highlights

> **Excel VBA + Data Analysis + Dashboard Automation**

A practical example of using Excel VBA to transform transactional data into an interactive management reporting solution.
