# googleSheetsFlows
#### Conditionally copy rows from one Google Sheet to another based on cell value. 

This is meant to be run regularly to maintain an eventual, one-directional sync between a source sheet and its child(ren) sheet(s). It is built on the Google Sheets Python API. It assumes you are using a credentials.json file for authentication.

### BACKGROUND
This was built for a volunteer project to filter a list of all volunteers into initiative-specific sheets based on the input from the volunteer forms. Since people may register as volunteers at any time, it was helpful to have an automated script for checking for new volunteers that were interested in specific initiatives. This script enables initiative-leaders to receive only the relevant info about only the volunteers who were interested in their initiatives.

