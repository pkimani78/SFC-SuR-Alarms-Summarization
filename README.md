# SFC-SuR-Alarms-Summarization
This script creates a summary report for all alarms observed in a given period. 
The provided .xlsx file MUST have SuR,4G,3G and 2G alarms worksheets for this script to work. 
It creates a new "_updated[date].xlsx" file on the directory the script is run.
A new worksheet "SuR Analysis" is created in the above document where further analysis can be conducted.
Inputs: minimum unavailability threshold, alarms dates. 
Sites with unavailability mins above the minimum threshold will have their alarms summarized for the given dates

How to Run the script
---------------------
1. Does not need python installed on local machine
    run the executable script on cmd followed by the path_to_xlsx_file.xlsx
    i.e sfc_alarms_script.exe test.xlsx
2. Clone the project and execute with python in a virtual env in the same way as explained above.
   More complex as you may need to install python + all dependencies needed on Pipfile 
    i.e sfc_alarms_script.py test.xlsx
