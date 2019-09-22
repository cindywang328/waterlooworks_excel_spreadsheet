# waterlooworks_excel_spreadsheet
Pulls job descriptions from WaterlooWorks into an Excel spreadsheet using Selenium Web Driver and Excel VBA
"Inspected element" a lot but it was worth! Now I can search for keywords in job postings. 

Opens WaterlooWorks, logs in (fill in the "UWATERLOOUSERNAME"), clicks "Hire Waterloo Co-op", then goes through all the pages and opens each posting, saving info such as:
- job ID
- title
- how many openings
- location
- level (junior, intermediate, senior)
- current number of applications
- job summary
- job responsibilities
