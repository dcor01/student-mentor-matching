# student-mentor-matching

# Student-Mentor Matching Algorithm
A Python script to automate the process of matching student mentees with mentors based on a prioritized set of rules, including age, gender preference, and campus location.
## Project Overview
This script provides an automated solution to the complex task of creating meaningful student-mentor pairings. It moves beyond simple preference matching by implementing a hierarchical logic that prioritizes age, enforces mandatory gender preferences, and considers logistical factors like campus location. The entire process is data-driven, reading from and writing to simple Excel files, and provides clear output for both successful and unsuccessful matches.
## Features
Age-Based Priority: Matches the oldest students with the oldest available mentors first.
Mandatory Gender Preference: Strictly enforces student-stated gender preferences ("Male" or "Female") while correctly interpreting flexible options ("Either way is fine!").
Campus Proximity Matching: Uses campus location as a primary tie-breaker to pair mentors and students who are physically closer.
Mentor Capacity Control: Ensures no mentor is assigned more than a configurable maximum number of mentees.
Automated Data Preprocessing:
Converts mentor titles (Mr./Ms.) into a standard gender column.
Maps detailed faculty names into a simple campus identifier ('1' or '2').
Safely converts age data from strings to numbers.
Clear Output: Generates a clean Excel report of all successful matches and prints a list of any unmatched students to the console for manual review.
## How to Use
### 1. Prerequisites
Make sure you have Python installed. You will also need the pandas and openpyxl libraries. You can install them using pip:
```
pip install pandas openpyxl
```
### 2. Input File Setup
The script requires an Excel file named mentors.xlsx to be in the same directory. This file must contain two sheets named Students and Mentors.
#### Sheet 1: Students
This sheet must contain the following columns (headers can be uppercase or lowercase):
age: The student's age (e.g., "21" or "21 years").
program: The student's program of study (e.g., "Masters").
faculty: The student's full faculty name (e.g., "Law Faculty").
gender_preference: The student's preference for their mentor's gender. Must be one of:
Male
Female
Either

#### Sheet 2: Mentors
This sheet must contain the following columns:
age: The mentor's age.
program: The mentor's program or field.
mr./ms.: The mentor's title (e.g., "Mr.", "Ms.", "Mrs.").
faculty: The mentor's faculty.

### 3. Running the Script
Place your mentors.xlsx file in the same folder as the Python script.
Open a terminal or command prompt in that folder.
Run the script using the following command:
```
python your_script_name.py
```
### 4. Configuration
You can easily change the maximum number of mentees a mentor can have by editing the MENTOR_CAPACITY variable at the top of the main execution block in the script.
```
--- Main Execution ---
if name == "main":
MENTOR_CAPACITY = 10 # Change this value as needed
```
## Output
After running, the script will produce two outputs:
mentor_matches_final.xlsx: A new Excel file will be created in the same directory containing a list of all successful pairings and their relevant details.
Console Output: If any students could not be matched due to rule conflicts or lack of available mentors, their details will be printed directly to the terminal for easy follow-up.
