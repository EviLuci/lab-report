# Lab Report Automation Using Google Apps Script
This project automates the generation of lab reports from Form submissions into a Google Sheet. Leveraging Google Apps Script, it processes form data, organizes tests by department, and formats test results into a structured report in Google Docs. The script dynamically adjusts report layout and prevents page breaks within individual test tables, creating a professional and organized output.

## Project Overview
This script project is designed to streamline the process of generating lab reports. Once a lab technician submits test data via a Form, the responses populate a Google Sheet. The Apps Script then:

- Categorizes the tests by department.
- Formats the report with department-specific tables.
- Ensures page breaks only occur between departments, not within tables, maintaining a clean layout for easy readability.
- Creates a doc and pdf file with header and without header into separate folder

*templates*: The templates folder contains the docs with the table templates created and used for the report generation
*testing*: The testing folder only contains some test and scratch function and logic I came up with trying out some new features and improvements for the Lab Report Automation. This folder is just for my reference and don't have any useful scripts that are needed for Lab Report Automation.

## Known Issues
- Limited Page Break Control: Google Docs API has limitations in managing exact positioning, which may cause minor layout inconsistencies.
- Estimated Table Height: Table height is estimated based on average row size, which may vary slightly with different font styles or sizes.