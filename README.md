# 📊 Excel Salary Dashboard

![Salary Dashboard](images/1_Salary_Dashboard_Final_Dashboard.gif)

## Introduction

This data jobs salary dashboard was created to help job seekers investigate salaries for their desired roles and ensure they are being adequately compensated.

This project was built by following **[Luke Barousse's Excel Course](https://www.lukebarousse.com/excel)**, which provides a hands-on foundation in analyzing real-world data with Excel. All credit for the course structure, teaching, and the underlying dataset goes to Luke — this repo documents my own build of the final dashboard as I worked through the course.

The dataset contains real-world data science job postings from 2023, including:

- 👨‍💼 **Job titles**
- 💰 **Salaries**
- 📍 **Locations**
- 🛠️ **Skills**

### Dashboard File
The final dashboard is in [1_Salary_Dashboard.xlsx](1_Salary_Dashboard.xlsx).

### Excel Skills Used

- **📉 Charts**
- **🧮 Formulas and Functions**
- **❎ Data Validation**

## Dashboard Build

### 📉 Charts

#### 📊 Data Science Job Salaries — Bar Chart

<img src="images/1_Salary_Dashboard_Chart1.png" width="850" height="550" alt="Salary Dashboard Chart1">

- 🛠️ **Excel Features:** Bar chart with formatted salary values, laid out for clarity.
- 🎨 **Design Choice:** Horizontal bar chart for easy visual comparison of median salaries.
- 📉 **Data Organization:** Job titles sorted by descending salary for readability.
- 💡 **Insights Gained:** Quick identification of salary trends — Senior roles and Engineers out-earn Analyst roles.

#### 🗺️ Country Median Salaries — Map Chart

![Country Map Chart](images/1_Salary_Dashboard_Country_Map.gif)

- 🛠️ **Excel Features:** Excel's map chart feature, used to plot median salaries globally.
- 🎨 **Design Choice:** Color-coded map to visually differentiate salary levels across regions.
- 📊 **Data Representation:** Median salary plotted for each country with available data.
- 👁️ **Visual Enhancement:** Improved readability and immediate understanding of geographic salary trends.
- 💡 **Insights Gained:** Quick grasp of global salary disparities and which regions pay highest/lowest.

### 🧮 Formulas and Functions

#### 💰 Median Salary by Job Title

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
```

- 🔍 **Multi-Criteria Filtering:** Checks job title, country, and schedule type, and excludes blank salaries.
- 📊 **Array Formula:** Uses `MEDIAN()` with a nested `IF()` to evaluate an array.
- 🎯 **Tailored Insights:** Returns salary information specific to a given job title, region, and schedule type.
- **🔢 Formula Purpose:** Populates the background table that drives the dashboard's job title dropdown.

🍽️ Background Table

![Background Table 1](images/1_Salary_Dashboard_Screenshot1.png)

📉 Dashboard Implementation

<img src="images/1_Salary_Dashboard_Job_Title.png" width="400" height="500" alt="Salary Dashboard Title">

#### ⏰ Count of Job Schedule Type

```excel
=FILTER(J2#,(NOT(ISNUMBER(SEARCH("and",J2#))+ISNUMBER(SEARCH(",",J2#))))*(J2#<>0))
```

- 🔍 **Unique List Generation:** Uses `FILTER()` to exclude entries containing "and" or commas, and to omit zero values.
- **🔢 Formula Purpose:** Populates the background table that gives a clean list of unique job schedule types.

🍽️ Background Table

![Background Table 2](images/1_Salary_Dashboard_Screenshot2.png)

📉 Dashboard Implementation

<img src="images/1_Salary_Dashboard_Type.png" width="350" height="500" alt="Salary Dashboard Type">

### ❎ Data Validation

#### 🔍 Filtered List

Implementing the filtered lists as data validation rules under the `Job Title`, `Country`, and `Type` fields ensures:

- 🎯 User input is restricted to predefined, valid options
- 🚫 Incorrect or inconsistent entries are prevented
- 👥 Overall usability of the dashboard is improved

<img src="images/1_Salary_Dashboard_Data_Validation.gif" width="425" height="400" alt="Salary Dashboard Data Validation">

## Conclusion

I built this dashboard to explore salary trends across data-related job titles and to make informed decisions about my own career path — comparing how location and job type influence pay. Big thanks to **Luke Barousse** for the excellent course this project is based on.

## Credit

- Course: [Luke Barousse's Excel Course](https://www.lukebarousse.com/excel)
- YouTube: [Luke Barousse](https://www.youtube.com/@LukeBarousse)
- Dataset: provided as part of the course
