# SFC-SuR-Alarms-Summarization
This script creates a summary report for all alarms observed in a given period. 
The provided .xlsx file MUST have SuR,4G,3G and 2G alarms worksheets for this script to work. 
It creates a new "_updated[date].xlsx" file on the directory the script is run.
A new worksheet "SuR Analysis" is created in the above document where further analysis can be conducted.
Inputs: minimum unavailability threshold, alarms dates. 
Sites with unavailability mins above the minimum threshold will have their alarms summarized for the given dates
