# Index of Python projects
I wrote a bunch of code to automate boring tasks like data entry, updating spreadsheets, scraping data from PDF reports and Excel reports with weird formatting that wouldn't fit well in a pandas DataFrame. 

## [Report Generator](https://github.com/campos-Allan/report_generator) 
This Python-based solution automates the entire Excel reporting workflow. It opens multiple Excel workbooks containing Power Query queries, intelligently monitors Excel’s CPU usage to detect when data refreshes finish, exports specified worksheets as timestamped PDF files, and cleanly closes all files. Additionally, it generates a monthly styled graphic from updated data, exports it as a PNG image, and triggers a Power Automate flow that posts everything directly to a Teams channel.

Designed for corporate environments that regularly generate reports from Excel dashboards with external data sources, this automation reduces manual report generation time from approximately one hour to just 15 minutes.

The solution requires no VBA or macros, works seamlessly on Windows with Excel installed, and provides reliable, error-free report generation and communication. It’s ideal for logistics, operations, and other data-intensive workflows that benefit from automated refresh, export, and visual reporting integration.

## [Sheet Generator](https://github.com/campos-Allan/sheet_generator) 
This project automates the weekly preparation of Excel reports by copying the most recent file, renaming it for the upcoming week, and updating key cells with new dates and labels. It also refreshes all data connections and saves the updated file, eliminating the need for manual intervention in recurring reporting tasks.

Ideal for business dashboards and regular reporting cycles, this script ensures consistency, saves time, and reduces errors. Simply configure your source and destination folders, and let the automation handle the rest.

## [Easy Info](https://github.com/campos-Allan/easy_info) 
A rudimentary system to extract about 300 data points from scrambled images, which needed to be manually transferred to 10 spreadsheets with different formatations. I automated the process to input the data once and format it correctly in each spreadsheet. Over time, I updated the system as most of the data was made available in a spreadsheets instead of structureless images, speeding up the process from a 3 day data entry process, to 3 minutes. I haven't update the code here with this last change.

**Tools used:** Openpyxl.

## [Master Scraper](https://github.com/campos-Allan/master_scraper)
I developed this program to automate a repetitive task of retrieving values daily from 5 different PDF files and 5 Excel spreadsheets, and then formatting everything in a specific way to be inserted into a large shared Excel spreadsheet. 

**Tools used:** Pandas, Openpyxl, Tabula, Tkinter, PyAutoGUI, PyPDF.

## [Scraping Model](https://github.com/campos-Allan/scraping_model)
Based on all the previous code, I created this general clean model to scrape data from PDF and Excel files.

**Tools used:** Pandas, Openpyxl, Tabula.
