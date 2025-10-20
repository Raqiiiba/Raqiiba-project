# Excel Data Projects 📊

This repository contains exploratory data analysis case studies using Excel. The files represent real-world datasets and showcases practical Excel skills such as sorting, filtering and basic data cleaning.

## 📁 Included Files

### 1. Ask A Manager Salary Survey

- Dataset: Salary responses from professionals in various roles.
- insights:
    - Salary ranges by job title, location, or industry.
   - Visualizations using Excel charts and pivot tables.

### 2. Actors Case Study

- Dataset: it includes film, actor, or industry-related data.
- use: This summarizes actors activities, performance stats and trends.

## 3. Steps to cleaning the data
  
  **Step 1: Preparing the Files**

I first saved a backup copy of the original dataset in a folder called raw, then created another folder named clean for the cleaned version.
I opened the dataset in Excel to explore the structure, check the number of rows and columns, and note where there were missing values or strange symbols. 

**Step 2: Renaming Columns**

Next, I renamed the columns to shorter and clearer names to make them easier to work with.
I kept all column names in lowercase and removed spaces for consistency.

**Step 3: Cleaning Data Types**

I noticed that the salary column had symbols like $, commas, and “k”.
So, I used Find and Replace to remove them:
+Replaced $ and , with nothing.
Replaced k with 000 (for example, 50k → 50000).
Then, I changed the cell format to Number so that all salary values were numeric.
If any salary was still text, I used the formula =VALUE(A2) to convert it.
For date columns, I used =DATEVALUE() to fix formatting issues.


**Step 4: Handling Missing Data**

To deal with missing information, I filtered each column using Data → Filter → Blanks.
I removed rows that were missing important data like salary.
For less important columns, I filled blank spaces with “Unknown”.
I also removed unrealistic salary values (for example, below 1,000 or above 5,000,000).

**Step 5: Standardizing Text Data**

I worked on text columns to make them consistent.
Country: Used =PROPER(A2) to fix letter casing.
        Changed “U.S.”, “USA”, and “United States of America” to one standard name  “United States.”
Job Title: Used =LOWER(A2) to make everything lowercase.For example:
“developer”, “programmer”, “coder” → “software engineer
“teacher”, “tutor” → “teacher”
Created a small mapping table to group similar roles.
Gender: Replaced “F”, “female”, and “woman” with “Female”.
        Replaced “M”, “male”, and “man” with “Male”

**Step 6: Cleaning Experience and Education Columns**

For experience ranges like “3–5 years”, I converted them into average numbers such as 4.
For education, I grouped different responses into simpler categories:
“Bachelor’s”
“Master’s”
“PhD”
This made the data easier to analyze.

 Step 7: Removing Duplicates**

I selected the full dataset, went to Data → Remove Duplicates, and checked columns like salary, job title, and country to ensure only unique rows remained.

**Step 8: Creating New Columns**

I added two new columns to help with analysis later:
Column	Formula	Purpose
Salary Level	=IF(B2<50000,"Low",IF(B2<100000,"Medium","High"))	Grouped salaries into levels
Experience Level	=IF(C2<3,"Junior",IF(C2<10,"Mid","Senior"))	Categorized experience levels

 **Step 9: Checking and Validating the Data**
Before saving, I checked everything carefully:
No blank salary cells
Correct text and number formats
Consistent capitalization and naming
I also created quick pivot tables to view:
Average salary by country
Average salary by job title
Salary by experience level

**Step 10: Saving and Documenting**

After cleaning, I saved the file as ask_a_manager_salary_cleaned.xlsx in the clean folder.
I also added a sheet called “Cleaning Notes” where I recorded the main changes I made.

**Step 11: Protecting Personal Information**

I deleted all personal or identifying data before saving the final file to make sure it only included general job, salary, and demographic information.

**Step 12: Final Outcome**

After cleaning, the dataset was:
Consistent and well-organized
Free of duplicates and missing data


**Skills I Used**

Data cleaning and organization
Handling missing and inconsistent data
Text standardization and formatting
Using formulas for data transformation
Documenting and validating data in Excel

## 🛠 Tools Used

- Microsoft Excel

## 📌 Note

These datasets are for educational and demonstrative purposes only.
