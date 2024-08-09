# Personal-progress-Optimzer
A end-to-end Computer program analysis Project which analyse and process data through Machine Learning and AI.


Activity Monitoring and Analysis Tool
Overview
This tool monitors and logs computer activity, categorizes applications, and provides insights into usage patterns. It helps users understand their productivity, track time spent on various activities, and offers recommendations based on their usage.

Prerequisites
Python Installation: Ensure Python 3.x is installed on your system.

Required Libraries: Install the necessary Python libraries using pip. You can use the following command:

bash
Copy code
pip install psutil win32gui pandas openpyxl scikit-learn matplotlib gspread oauth2client docx
Google Sheets Credentials: If using Google Sheets for master database storage, you'll need a JSON key file for Google Sheets API access. Follow Google's documentation to create and download the credentials.

Excel File: Ensure the Excel file for the master database is available at the specified path.

Setup
Download or Clone the Project:

Download or clone the repository containing the project code.
Configure Paths and Credentials:

Update the EXCEL_PATH variable in the code with the path to your Excel master sheet.
If using Google Sheets, ensure the path to your credentials JSON file is correct.
Usage Instructions
Start the Tool:

Run the main Python script using:
bash
Copy code
python your_script_name.py
GUI Overview:

Start Monitoring: Click the "Start Monitoring" button to begin logging activity. This will create an analysis folder and start recording data.
Stop Monitoring: Click the "Stop Monitoring" button to stop logging activity and generate the analysis report.
Monitoring Activity:

The tool will log active window titles and CPU usage at regular intervals.
Categorizing Applications:

If the tool encounters a new application, it will prompt you to categorize it. Choose a category from the dropdown list and click "Save."
Analyzing Data:

After stopping monitoring, the tool will analyze the logged data, detect anomalies, and generate a report.
The analysis report includes insights on computer usage, productivity, and recommendations.
Interpreting the Report
Computer Goodness:

Provides an overview of system performance, including CPU usage and temperature.
Indicates if the computer's performance is "Good" or "Needs Attention."
Human Goodness:

Assesses the balance between productive and entertainment activities.
Provides feedback on how well the time is managed between different categories.
Detailed Analysis:

Includes a breakdown of time spent on various applications.
Lists any detected anomalies and offers advice based on usage patterns.
Features a pie chart visualizing time distribution across categories.
Troubleshooting
Error Messages:

Check for specific error messages in the console or GUI. Common issues may include file path errors or missing libraries.
Excel File Issues:

Ensure the Excel file is not open or locked by another application when the script tries to save data.
Google Sheets Access:

Verify that your credentials are correctly set up and that you have access to the specified Google Sheet.
Contact and Support
For further assistance or to report issues, please contact the project maintainers at ganesh.khosur@gmail.com
