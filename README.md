# NessusCSV2XL

Python script automates IP based XL reporting for Nessus scan results. 
Input is the csv file exported from Nessus. 
The script performs the following operations and creates an Excel file containing vulnerabilities for each IP in separate sheets.

-	Drops Duplicates
-	Removes vulnerabilities with no risk (Info)
-	Removes some additional columns 
-	Sorted based on severity. 
-	Does formatting. 
-	Creates an additional sheet with all identified open ports.

## Usage
Copy the exported CSV file from Nessus to the same directory as the script.  
Usage: python3 nessusipreport.py [csvfilename.csv]

## Installing required libraries
pip3 install pandas  
pip3 install openpyxl
