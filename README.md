# xlsxDigest
Python LIB to digest ServiceAID adhoc report by the date of creation, in order to generate a simple xlsx heat map of incidents

It creates a xlsx file listing incidents in a simple table view, using the format of hour of the incident vs date of the incident (as shown in the sample file)

In order to digest it, make sure you have the exported csv from ServiceAID converted to xslx and also the "ticket creation date" in the 3rd column of the xlsx.

ps.: you can also filter by the "ticket close date", just make sure you place it as the 3rd column.

The requirements to use this lib is openpyxl version 2.5.3 or higher.

The way to use Digest lib is pretty simple:

# Import the LIB
from xlsxDigest import digest

# Initiate the file you wish to use (make sure that the file is in the same path of the lib, or just reference the full path)
file = digest("adhoc_report.xlsx")

# Populates the DataMap with the content from the file you just initiated
file.populateData()

# Export the dataMap to a xlsx file (a file with the name of sample.xlsx will be created in the same path of the lib)
file.generateReport()

# Feel free to contact me in case of any doubts or issues.
