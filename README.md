# Lab Report Automation Using Google Apps Script
This project automates the generation of lab reports from Form submissions into a Google Sheet. Leveraging Google Apps Script, it processes form data, organizes tests by department, and formats test results into a structured report in Google Docs. The script dynamically adjusts report layout and prevents page breaks within individual test tables, creating a professional and organized output.

[*Note: This is a simple automation project which was created for a temporary solution purpose only until a complete Lab Report Management System is created. It just gets the job done somehow <img src="https://raw.githubusercontent.com/Tarikul-Islam-Anik/Animated-Fluent-Emojis/master/Emojis/Smilies/Nerd%20Face.png" alt="Nerd Face" width="15" height="15" />*]

## Project Overview
This script project is designed to streamline the process of generating lab reports. Once a lab technician submits test data via a Form, the responses populate a Google Sheet. The Apps Script then:

1. Detects the Change in google sheet and triggers the script to run after checking if the detected change is acutally a new data row appended or not. [*Every Form submission appends a new row and it prevents running the script from manual data manipulation and only runs the full script after verifying if new row was appended or not*]
2. Creates folders with patient name inside the specified folder with REPORT_FOLDER_ID
3. Make a copy of doc with REPORT_TEMPLATE_ID and replace placeholders (*datas enclosed by << and >> inside the templates*) in Header with the data passed to sheet from form (*formData*) and similarly begin table insertion.
4. Categorizes the tests by department and insert tables so that tests that comes under same department are inserted.
5. After each table insertion, it checks if there is available space to adjust another test table without breaking the table into two pages. If there is enough space available for test of the same department to be inserted, the table is inserted. Otherwise, a page break is inserted.
6. Formats the report with department-specific tables.
7. Ensures page breaks only occur between departments, not within tables, maintaining a clean layout for easy readability.
8. Replace the placeholders with formData as mapped inside tesObjects.gs
9. Insert signature. [*signature are formated as a table without border*]
10. Creates a doc and pdf file with header. Similarly, the doc and pdf file without header is as well created and saved in 2 separate folders.
11. Sends the files and link to the folder as an mail in the gmail specified.

**Folders**
1. *templates*: The templates folder contains the docs with the table templates created and used for the report generation
2. *testing*: The testing folder only contains some test and scratch function and logic I came up with trying out some new features and improvements for the Lab Report Automation. This folder is just for my reference and don't have any useful scripts that are needed for Lab Report Automation.

## Known Issues
- Limited Page Break Control: Google Docs API has limitations in managing exact positioning, which may cause minor layout inconsistencies.
- Estimated Table Height: Table height is estimated based on average row size, which may vary slightly with different font styles or sizes.