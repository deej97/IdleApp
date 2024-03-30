# IdleApp
IdleApp
Motivation: Python is a very handy language that can be used to automate repetitive tasks like report generation for business reviews. In the support industry, tickets are logged and customers are aligned with engineers to work with them until the issue is resolved. Customer satisfaction is very important for any successful business. Besides providing quicker resolution, keeping customers updated regarding the progress of the case is essential.

To ensure that customers are informed and kept updated, identifying if the ticket is idle or not is essential. This tool will read the input file and identify if the ticket is idle or not using the last communication date. if the previous communication was sent before 5 days, it will mark the ticket as idle.

This idleapp.exe is built from a python script that can read data from .xlsx files and outputs the report in a .xlsx file with charts and tables.

How to run the .exe and generate output: Inside the dist folder, double click on idleapp.exe. It will prompt for the input file name. I have added an input file in the same directory as inputdata.xlsx. Enter the input file name with inputdata.xlsx And this will publish the output in result.xlsx

Technical information: Python packages such as NumPy, pandas have been used for reading and processing input data from excel. Basic data preprocessing was done to calculate the idle days as the date was in a different format. Used XlsxWriter https://pypi.org/project/XlsxWriter/ to create reports with charts and tables in excel. Used pyinstaller module https://pypi.org/project/pyinstaller/ to convert .py to .exe, so that the tool can run on devices without python installed.
