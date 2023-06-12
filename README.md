# Python-script-for-SQL
This is a Python Script that connects to a remote linux server to execute SQL Queries and then create a Word or PDF report based on the template provided currently.

This is what the script performs in steps:

1. Connects to linux Host
2. Reads an Sql file that contains Sql queries
3. prints the queries
4. Executes the Sql queries in the Database
5. checks the current Date (to be later used for report)
6. Create a Word Document
7. reads and replace the table contents of each table with the query results
8. Saves the file 
9. Closes the connection

