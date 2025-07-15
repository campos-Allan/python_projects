# Index of Python projects
I wrote a bunch of code to automate boring tasks like data entry, updating spreadsheets, scraping data from PDF reports and Excel reports with weird formatting that wouldn't fit well in a pandas DataFrame. 

## [Report Generator](https://github.com/campos-Allan/report_generator) 
This Python-based solution automates the entire Excel reporting workflow. It opens multiple Excel workbooks containing Power Query queries, intelligently monitors Excel’s CPU usage to detect when data refreshes finish, exports specified worksheets as timestamped PDF files, and cleanly closes all files. Additionally, it generates a monthly styled graphic from updated data, exports it as a PNG image, and triggers a Power Automate flow that posts everything directly to a Teams channel.

Designed for corporate environments that regularly generate reports from Excel dashboards with external data sources, this automation reduces manual report generation time from approximately one hour to just 15 minutes.

The solution requires no VBA or macros, works seamlessly on Windows with Excel installed, and provides reliable, error-free report generation and communication. It’s ideal for logistics, operations, and other data-intensive workflows that benefit from automated refresh, export, and visual reporting integration.

## [Master Scraper](https://github.com/campos-Allan/master_scraper)
This is a Python-based automation tool designed to simplify and accelerate the daily process of consolidating logistics data. By reading and processing multiple PDF and Excel files, it extracts key information about product storage and transportation across various locations and formats the results into a unified Excel file. The tool handles inconsistent file formats, and automatically organizes, and archives processed files. This ensures accurate, up-to-date reporting with minimal manual effort, making it ideal for teams managing complex logistics operations.

## [Sheet Generator](https://github.com/campos-Allan/sheet_generator) 
This project automates the weekly preparation of Excel reports by copying the most recent file, renaming it for the upcoming week, and updating key cells with new dates and labels. It also refreshes all data connections and saves the updated file, eliminating the need for manual intervention in recurring reporting tasks.

Ideal for business dashboards and regular reporting cycles, this script ensures consistency, saves time, and reduces errors. Simply configure your source and destination folders, and let the automation handle the rest.

