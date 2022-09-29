# What for?

The script is used for aggregating database extracts from mySAGW.

# Prerequisites

To run the script successfully, the following requirements must be met:

- Python3
- pip installed
- Libraries: openpyxl, pandas
- Access to mySAGW as admin
- XLSX file with the correct column names

Tested under MacOS 12.6 (Intel)

# Structure of the Excel files

The Excel files must have at least the following columns:

SECTION
MEMBER INSTITUTION
FORM_ID
REFERENCE NO.

# What happens during execution?

1. the user selects the XLSX file with the data
2. the script creates a main data frame
3. script creates separate data frames for the individual sections
4. user selects storage location
5. script creates own XLSX files from the individual data frames
