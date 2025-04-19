# Index of Python projects
I wrote a bunch of code to automate boring tasks like data entry, updating spreadsheets, scraping data from PDF reports and Excel reports with weird formatting that wouldn't fit well in a pandas DataFrame. **I used this in some jobs, and because of that, I created mock-up data to post here as an example. I also changed variable names and other aspects of the code to preserve the real data, so the code could be a bit messy and hard to understand.**

## [Excel Generator](https://github.com/campos-Allan/excel_generator) 
Basic script that picks a spreadsheet model based on the month in question, changes a few cells to update the queries and then saves this spreadsheet with specific name formatting. 

**Tools used:** Pywin32.

## [PDF Generator](https://github.com/campos-Allan/pdf_generator) 
The code opens a few spreadsheets that would update its queries upon opened, and then generated PDF files from each spreadsheet. Then the code would acess another spreadsheet, update its queries, copy the data to a third spreadsheet, paste it there in a specific place and take a print screen of the result. 

**Tools used:** Pywin32, PyAutoGUI.

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
