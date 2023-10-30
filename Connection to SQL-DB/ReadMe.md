# VBA Script for Connecting to a SQL Database

This VBA script is designed to connect to a SQL database and retrieve data. It uses the ADODB library for database connectivity and is intended for use with Microsoft Excel.

## Features

- Connect to a SQL database.
- Execute an SQL query.
- Retrieve data from the database and populate it into an Excel worksheet.

## Prerequisites

Before using this script, you should have the following:

- Microsoft Excel installed on your computer.
- Appropriate permissions to access the SQL database.
-

## Setup

1. **Open the VBA Editor in Excel:**
   - Press `ALT + F11` to open the VBA editor.

2. **Insert a Module:**
   - In the VBA editor, go to `Insert > Module`. This will create a new module where you can paste the VBA code.

3. **Paste the VBA Code:**
   - Copy and paste the VBA code from the [MyConnectionToDB](MyConnectionToDB.vba) file into the module (just change the variables
   specified in the code.)

4. **Specify Database Connection Details:**
   - In the VBA code, locate the section where it says "Specify your database connection details." Replace the placeholders with your actual database information:

   ```vba
   ' Specify your database connection details
   dbServer = "YourServerName"     ' Replace with your database server
   dbCatalog = "YourDatabaseName"  ' Replace with your database name
   dbUser = "YourUsername"        ' Replace with your database username
   dbPassword = "YourPassword"    ' Replace with your database password

## NOTE 

To enhance security, consider storing sensitive information like usernames and passwords in a more secure manner, such as using Excel's built-in "Protected Sheets and Workbooks" feature, or even better, using environment variables to store these credentials outside of the code. This way, you can access the credentials without hardcoding them in the VBA code.



