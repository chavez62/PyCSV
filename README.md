# Access Database to CSV Converter

This is a Python GUI application that allows users to run SQL queries on an Access database and export the results to a CSV file. The application is built using Tkinter and ttkbootstrap for the GUI, pyodbc for database connectivity, and pandas for data manipulation.

## Features

+ Select an Access database file (.accdb or .mdb).
+ Input a custom SQL query.
+ Export the query results to a CSV file.
+ Modern and user-friendly interface using ttkbootstrap.

## Installation
Prerequisites
+ Python 3.x
+ The following Python libraries:
+ tkinter (usually included with Python)
+ ttkbootstrap
+ pyodbc
+ pandas

## Step-by-Step Guide
1. Clone the repository:
```
git clone https://github.com/your-username/access-database-to-csv.git
cd access-database-to-csv
```
2. Install the required libraries:
```
pip install ttkbootstrap pyodbc pandas
```
3. Run the application:
```
python access_to_csv.py
```
## Usage
1. Select the Access Database File:
   + Click the "Browse..." button to select your Access database file (.accdb or .mdb).
2. Input the SQL Query:
   + Enter your SQL query in the provided text area.
3. Generate CSV:
   + Click the "Generate CSV" button to execute the query and save the results to a CSV file. You will be prompted to choose the save location.
