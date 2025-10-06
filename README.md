# 💼 Excel Salary Dashboard

![1_Salary_Dashboard_Chart2.png](Resource/1_Salary_Dashboard_Final_Dashboard.gif)

## 📘 Introduction

I built this **Data Jobs Salary Dashboard** in Excel to help job seekers explore salary trends across various data-related roles.  
The goal was to provide clear, data-driven insights into how factors like job title, country, and schedule type influence compensation — helping professionals understand their market value.

---

### 📂 Dashboard File

You can access my completed Excel dashboard here:  
👉 [**1_Salary_Dashboard.xlsx**](1_Salary_Dashboard.xlsx)

---

## 🧠 Excel Skills Applied

Throughout this project, I leveraged several Excel features and functions to build an interactive and insightful dashboard:

- **📊 Charts**
- **🧮 Formulas & Functions**
- **✅ Data Validation**

---

## 📊 Dataset Overview

The dataset I worked with contains **real-world data science job information from 2023**, covering key details such as:

- 👨‍💼 **Job Titles**  
- 💰 **Salaries**  
- 📍 **Countries**  
- 🛠️ **Skills**

This dataset served as the foundation for creating visual and analytical insights in Excel.

---

## 🏗️ Dashboard Build

### 📉 Charts

#### **1️⃣ Data Science Job Salaries — Bar Chart**

<img src="/Resource/1_Salary_Dashboard_Chart1.png" width="850" height="550" alt="Salary Dashboard Chart1">

- 🛠️ **Excel Features Used:** Bar charts with formatted salary values and an optimized layout for clarity.  
- 🎨 **Design Choice:** I used a **horizontal bar chart** to make it easier to compare median salaries visually.  
- 📊 **Data Handling:** Salaries were sorted in descending order to highlight top-paying roles first.  
- 💡 **Key Insight:** Senior-level positions and engineering roles consistently offer higher pay than analyst roles.

---

#### **2️⃣ Country Median Salaries — Map Chart**

![1_Salary_Dashboard_Chart2.png](Resource/1_Salary_Dashboard_Country_Map.gif)

- 🗺️ **Excel Features Used:** Leveraged **Excel’s map chart** to visualize salary differences across countries.  
- 🎨 **Color-Coding:** Each country’s median salary was represented through a gradient color scale.  
- 👁️ **Visualization Benefit:** Instantly communicates salary disparities between regions.  
- 💡 **Key Insight:** Certain countries stand out with notably higher median salaries, helping job seekers target specific markets.

---

### 🧮 Formulas & Functions

#### 💰 **Median Salary by Job Title**

```excel
=MEDIAN(
IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
)
)
`        
