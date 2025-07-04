# Index of Python projects
I wrote a bunch of code to automate boring tasks like data entry, updating spreadsheets, scraping data from PDF reports and Excel reports with weird formatting that wouldn't fit well in a pandas DataFrame. 

## [Report Generator](https://github.com/campos-Allan/report_generator) 
Is a Python-based solution that eliminates the tedious manual process of managing Excel workbooks with external data connections. The tool automatically opens multiple Excel files containing Power Query queries, intelligently monitors CPU usage to detect when data refreshes are complete, exports specified worksheets to timestamped PDF files, and closes everything cleanly without any user intervention. This automation reduces manual reporting tasks from approximately one hour to just 15 minutes, making it ideal for corporate environments that regularly generate reports from Excel dashboards connected to external data sources. The solution requires no VBA code or macros, works seamlessly with Windows and Excel installations, and provides reliable, error-free report generation for logistics, operations, and other data-intensive workflows.

## [Excel Generator](https://github.com/campos-Allan/excel_generator) 
Basic script that picks a spreadsheet model based on the month in question, changes a few cells to update the queries and then saves this spreadsheet with specific name formatting. 

**Tools used:** Pywin32.

## [Easy Info](https://github.com/campos-Allan/easy_info) 
A rudimentary system to extract about 300 data points from scrambled images, which needed to be manually transferred to 10 spreadsheets with different formatations. I automated the process to input the data once and format it correctly in each spreadsheet. Over time, I updated the system as most of the data was made available in a spreadsheets instead of structureless images, speeding up the process from a 3 day data entry process, to 3 minutes. I haven't update the code here with this last change.

**Tools used:** Openpyxl.

## [Excel Updater](https://github.com/campos-Allan/excel_updater) 
This code opens a spreadsheet, updates the queries in it, retrieves scrambled data from a 'realized' sheet, and cross-references it with 'planned' data from other spreadsheets in a sheet that shows the final result of this data cross-checking.

**Tools used:** Pywin32.

## [Master Scraper](https://github.com/campos-Allan/master_scraper)
I developed this program to automate a repetitive task of retrieving values daily from 5 different PDF files and 5 Excel spreadsheets, and then formatting everything in a specific way to be inserted into a large shared Excel spreadsheet. 

**Tools used:** Pandas, Openpyxl, Tabula, Tkinter, PyAutoGUI, PyPDF.

## [Scraping Model](https://github.com/campos-Allan/scraping_model)
Based on all the previous code, I created this general clean model to scrape data from PDF and Excel files.

**Tools used:** Pandas, Openpyxl, Tabula.
