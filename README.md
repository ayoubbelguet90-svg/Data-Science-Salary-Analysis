# Data-Science-Salary-Analysis
An Advanced Excel Project analyz
# Data Science Salary & Job Market Dashboard

An interactive, data-driven Excel dashboard designed to analyze global data science salaries, employment trends, and job distribution across various roles, countries, and employment types. Powered by modern dynamic array formulas, this project bridges raw market data with clean, actionable executive insights.

---

## 🖥️ Dashboard Overview

The user interface is engineered around three main operational zones:

### 1. Global Interactive Slicers (Controls)
Located at the top of the interface to dynamically filter the entire data model in real-time:
*   **Job Title:** Dropdown selector to filter by specific roles (e.g., *Data Analyst*, *Machine Learning Engineer*).
*   **Country:** Dropdown selector to filter metrics by geographic location (e.g., *United States*).
*   **Work Type:** Dropdown selector to isolate employment structures (e.g., *Full-time*, *Part-time*, *Contractor*).

### 2. Analytical Visualizations
*   **Salary by Job Title (Horizontal Bar Chart):** Compares median compensation scales across high-demand positions like Senior Data Scientist, Data Engineer, and Cloud Engineer.
*   **Geographic Map Visualization:** A live-shaded map plotting global data density and workforce distribution.
*   **Salary by Work Type (Horizontal Bar Chart):** Details compensation variances based on contract categories (Full-time, Part-time, Contractor, Temp work, Internship).

### 3. Executive KPI Cards
Summary metric boxes at the base providing immediate top-line answers:
*   **Median Salary:** Displays the exact central salary point for the selected criteria (e.g., **\$90,000**).
*   **Top Job Platform:** Highlights the dominant hiring platform sourcing these positions (e.g., **Indeed**).
*   **Job Count:** Indicates the total sample pool size or live openings matching the current filters (e.g., **6,480**).

---

## ⚙️ Technical Architecture & Backend Logic

This dashboard bypasses legacy VBA macros by leveraging Excel's modern **Dynamic Array Engine** and boolean logic arrays. The entire backend scales automatically as new rows are appended to the core `jobs` or `Data` tables.

### 1. Unique Platform Extraction
Automatically extracts a clean, deduplicated list of job sourcing platforms:
```excel
=UNIQUE(jobs[job_via])
```
*   **Application:** Populates rows dynamically with clean platform data (e.g., LinkedIn, Indeed) without static copy-pasting.

### 2. Multi-Criteria Job Volumetric Counting
Computes total job availability per platform while maintaining complete interactivity with active dashboard filters:
```excel
=COUNTIFS(jobs[job_via], A1, jobs[job_title_short], title, jobs[job_country], country, jobs[job_schedule_type], type)
```
*   **Application:** Acts as the data engine behind the **Job Count** KPI card and volume charts, processing multiple criteria simultaneously.

### 3. Advanced Dynamic Median Salary Engine
Calculates the true mathematical median salary from a dataset of over 32,000 rows while isolating statistical noise and enforcing user filters:
```excel
=MEDIAN(IF((Data!$K$2:$K$32673=B5)*(Data!$M$2:$M$32673<>0)*(Data[job_title_short]=title)*(Data[job_schedule_type]=type), Data!$M$2:$M$32673))
```
*   **Application:** Drives the **Median Salary** KPI card.
*   **Logic Breakdown:**
    *   `Data!$K$2:$K$32673=B5`: Isolates rows belonging to the specific platform in cell `B5`.
    *   `Data!$M$2:$M$32673<>0`: Excludes missing or `$0` entries to prevent dragging down the median calculation.
    *   `*(Data[job_title_short]=title)*(Data[job_schedule_type]=type)`: Employs boolean array multiplication (`AND` logic) to filter only the rows that match the active dashboard named ranges.

### 4. Smart Validation & Error Treatment
Isolates out-of-scope data fields dynamically to clean up backend calculation layers:
```excel
=IF($A3<>title, $B3, NA())
```
*   **Application:** Compares row metrics with active user selections. Returning an `#N/A` error is intentionally utilized here to force Excel charts to automatically ignore non-relevant data points instead of plotting them as zero lines.

### 5. Dynamic Array Lookups
Queries and maps attributes across fluid, expanding spill range vectors:
```excel
=XLOOKUP(A4, A1#, A1#, "", 0)
```
*   **Application:** Scans for exact matches (`0`) inside the dynamic array spill reference (`A1#`). If the reference cannot find a match, it cleanly outputs an empty string (`""`) to maintain a clean UI presentation.

---

## 🚀 How to Run the Project
1. Clone or download this repository.
2. Open the workbook using **Microsoft Excel 365** or **Excel 2021+** (required to support dynamic array features and spill range operator `#`).
3. Navigate to the main dashboard tab and utilize the slicers at the top to slice and dice the data market metrics instantly.
[Live Excel Sheet on OneDrive](https://1drv.ms/x/c/e32b32bc7ac0127f/IQDna172mIHpSJ3k2niaB7nbARv1AlKeCXsR_HStKMSGEQk)
<img width="1350" height="560" alt="Animation" src="https://github.com/user-attachments/assets/99b874d0-db79-435f-b3ff-b76edd145dfa" />
