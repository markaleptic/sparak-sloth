# Sparak Accounting Sloth
*Automate book keeping within Sparak (D+H AIS) Transaction Input Area*

### Description
Sparak Accounting Sloth automates manual book keeping entries into the Sparak Accounting System (D+H AIS)
by receiving entries from an Excel file and booking each entry through the Transaction Input Area.
There is additional support for deleting entries by inputting a desired count.

### How to Use
*Sparak*   
Open Sparak to to the *Transaction Input* window. This is the window after you have specified your batch number 
and selected *Process*. This window shows the green plus button, red x button, edit entry button, and the list of transactions.


*Enter Transactions*   
1. Open Sparak Accounting Sloth and go to *Enter Transactions* menu.  
2. Click *Load Transactions* button to select desired excel file and load in entries.  
3. Confirm Debit / Credit Totals and transaction count are accurate.  
4. Click *Enter Transactions*  



*Delete Transactions*  
1. Open Sparak Accounting Sloth and go to *Enter Transactions* menu.  
2. Click *Delete Transactions* button.  
3. A pop up window will appear. Type in desired number of entries to delete and   



### Excel Requirements & VBA
For the Excel file to be read into Sparak Accounting Sloth correctly, you must format and manage Excel's *Used Range* diligently. 

*VBA*  
Without using VBA to create, format, and add entries this application does not make much sense for you. **Contact me and I will work with you to create VBA code that will take care of everything described below.**  

*Sheet Name*  
In a rush to get this application completed, there are some limitations in the input file. The sheet that has the entries must be 
called *Payment Sheet* or no entries will be loaded in. Allowing for the user to select a sheet during entry loading will occur in 
future updates. 
 
 
*File Formatting*  
Each row in the Excel sheet represents an entry. Each column represents values based upon this table:

| Column A      | Column B      | Column C      | Column D      | Column E      | Column F      |
|:-------------:|:-------------:|:-------------:|:-------------:|:-------------:|:-------------:|
| Distribution Code | Account Number | Tran Code | Entry Amount | Effective Date | Description |

You can either manually type in the transaction information (this doesn't make much sense if you can automate with VBA) or paste in some data and use VBA to fill in the remaining portions automatically from predetermined settings. **Remember, I want to help you do this so reach out to me.**

Keep in mind if you're convinced manual entries are better:
* Inncorrect or nonexistent dist or tran codes, account numbers, or dates will be denied by Sparak. If this occurs, the method of entering payments will likely skip this entry and move on to the next.
* Distribution and Tran codes never go beyonds the tens place and are greater than zero.
* Entry Amounts are greater than zero.
* Effective Dates cannot be in the future.
* Empty descriptions mean just leaving the cell empty.

There is limited handling of data constraints within this application. Meaning, this program relies on the user to see these mistakes once the file is 
loaded in and incorrect types to be restricted by Sparak during the entry period.   

Future updates will provide additional popup messages to inform the user that there could be input errors.
 
 
*Managing the Used Range*  
Sparak Accounting Sloth determines the number of entries to make based upon multiple factors including the used range in a sheet. If used range is expanded beyond Column F and the user specified number of rows then the incorrect values will be read into the application.    

[Microsoft's](https://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.worksheet.usedrange.aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1)
definition of Used Range:
 > A used range includes any cell that has ever been used. For example, if cell A1 contains a value, and then you delete the value, then cell A1 is considered used. In this case, the UsedRange property will return a range that includes cell A1.

The best way to work around Used Range issues:   
1. Create a new sheet called *Payment Sheet* (mind the capitals and spaces).  
2. Copy your entries (A1:F[row number where your entries end]).  
3. Paste the entries into the new sheet starting in the first row in columns A to F.

*Supported File Types*
* .xlsx 
* .xlsm  

### System Requirements
* 64-bit Windows 10
* 2007 Excel or newer
 
### How to Install
* CVB Employees reach out to me for an installer.
* I use cx_Freeze to create an executable. Download source files and run ```python setup.py build``` in the terminal.
