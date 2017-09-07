**VBA Module Code**

Primary code that creates entries that can be entered into Sparak via Sparak Sloth. GL Accounts Numbers are redacted. 
This code was run every day for about a year and worked beautifully. After copying payments from our check register 
into Payment Sheet tab, the code determines GL entries for regular, ORE, and deficiency accounts, splits entries
based on preferred accounting guidelines per CFO, formats entries, looks for common errors, error handling, and
accepts user input for wire totals.

--------------------------------------------------------------------------------------------------------------------
_main_
```
Sub main()
' ***************************************************************
' *                                                             *
' *  Project:       PART 1b: C.R.E.A.M. Project                 *
' *  Created by:    Mark Allred                                 *
' *  Purpose:       This code is part of the cream project      *
' *                 that automates B49 daily book keeping.      *
' *                 This part of the project lets a user        *
' *                 paste in the daily payments into the        *
' *                 Payment Sheet then reorganizes the sheet,   *
' *                 checks for ORE & Other Income, checks for   *
' *                 errors, and then allows for user            *
' *                 customization. After code runs the sheet    *
' *                 will be ready to be read into the Python    *
' *                 pyautogui application that will input the   *
' *                 values into Sparak.                         *
' *                                                             *
' ***************************************************************

    Application.ScreenUpdating = False

' Delete rows and columns
    On Error GoTo DeleteError
        deleteColumn (4)
        deleteRow (1)
        deleteColumn (4)
        deleteRow (1)
    On Error GoTo 0
    
    
' Get count of rows for cycling
    Dim rowCount As Integer
    rowCounter rowCount, "Payment Sheet"

' Format and Swap Data
    On Error GoTo FormatError
        formatSwap
    On Error GoTo 0
    
' Create Entries
    On Error GoTo CreateEntryError
        create_GL_Entries rowCount
    On Error GoTo 0
    
' Change ORE borrowers to post to GL for ORE
    On Error GoTo ORECheckError
        checkGL "ORE Borrowers"
    On Error GoTo 0
    
' Change deficiency borrowers to post into Other Income
    On Error GoTo ORECheckError
        checkGL "ORE"
    On Error GoTo 0
    
' Change deficiency borrowers to post into Other Income
    On Error GoTo OtherCheckError
        checkGL "Other Income"
    On Error GoTo 0
    
' Reduce redundant entries for ORE and Other Income Entries
    On Error GoTo RedundancyError
        remove_redundancy
    On Error GoTo 0
    
' Create Wire Entries
    On Error GoTo CreateWireError
        create_wire_entries
    On Error GoTo 0

' Format for readability
    On Error GoTo FormatterError
        formatter
    On Error GoTo 0

' Look for errors
    On Error GoTo IronicError
        errorChecker
    On Error GoTo 0

    Application.ScreenUpdating = True

' Create import sheet based upon user input
    response = MsgBox("Would you like to create an import sheet for Slothful Accounting", vbYesNo)
  
    If response = vbYes Then
        exportSheet
    End If



  
    Exit Sub
' **************************
' * Error Handling Section *
' **************************
DeleteError:
    MsgBox ("Delete function encountered error: " & Err.Number)
    Exit Sub
FormatError:
    MsgBox ("Format Swap encountered error: " & Err.Number)
    Exit Sub
CreateEntryError:
    MsgBox ("Create Entry function encountered error: " & Err.Number)
    Exit Sub
ORECheckError:
    MsgBox ("Error encountered with ORE data. Error: " & Err.Number & "Proceeding with Subroutine.")
    Resume Next
OtherCheckError:
    MsgBox ("Error encountered with ORE data. Error: " & Err.Number & "Proceeding with Subroutine.")
    Resume Next
RedundancyError:
    MsgBox ("Error encountered removing redundancies. Error: " & Err.Number & "Proceeding with Subroutine.")
    Resume Next
CreateWireError:
    MsgBox ("Error encountered while creating Entries. Error: " & Err.Number & "Proceeding with Subroutine.")
    Resume Next
FormatterError:
    MsgBox ("Error with formatting. Executing program suicide.exe")
    Resume Next
IronicError:
    MsgBox ("An error occured while looking for errors. How Ironic!")
    Resume Next

End Sub
```
_deleteColumn_
```
Sub deleteColumn(ByVal colNum As Integer)
    ' Function deletes column number passed to function
    Sheets("Payment Sheet").Columns(colNum).EntireColumn.delete
End Sub
```
_deleteRow_
```
Sub deleteRow(ByVal rowNum As Integer)
    ' Function deletes row number passed to function
    Sheets("Payment Sheet").Rows(rowNum).EntireRow.delete
End Sub
```
_insertRow_
```
Sub insertRow(ByVal rowNum As Integer)
    ' Function inserts row at the row number passed to the function
    Sheets("Payment Sheet").Rows(rowNum).EntireRow.Insert
End Sub
```
_rowCounter_
```
Sub rowCounter(ByRef counter As Integer, ByVal sheetName As String)
' Counts rows in Payment Sheet
    With Sheets(sheetName)
        counter = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
End Sub
```
_formatSwap_
```
Sub formatSwap()
    ' Document should be formatted to read left to right like Transaction Express
    ' Distribution Code --> Account --> Tran Code --> Amount --> Effective Date --> Description
        
    Dim WS As Worksheet
    Set WS = Sheets("Payment Sheet")
        
    ' Insert
    WS.Columns(1).EntireColumn.Insert
    WS.Columns(3).EntireColumn.Insert
    WS.Columns(3).EntireColumn.Insert
    
    ' Swap
    WS.Columns(7).Cut Destination:=Range("D:D")
    WS.Columns(5).EntireColumn.Insert
    WS.Columns(7).Cut Destination:=Range("E:E")
    deleteColumn colNum:=7
    deleteColumn colNum:=7
    
    ' Clear borrowers of description values in F column
    WS.Columns(6).ClearContents
    
End Sub
```
_create GL Entries_
```
Sub create_GL_Entries(ByVal counter As Integer)
' Loop through sheet based on the counter to insert a row directly above
' each borrower row. Insert the distribution code and Tran Code for
' GL & Borrower rows.
    Dim WS As Worksheet
    Dim GL_Dist As String
    Dim GL_Account As String
    Dim GL_Tran As String
    Dim GL_Amount As Double
    Dim GL_Date As String
    Dim GL_Description As Double
    Dim borrower_Dist As String
    Dim borrower_Tran As String
    Dim nonaccrual_Tran As String
    Dim principal_Tran As String
    Dim closed_Tran As String
    Dim distCode_Col As Integer
    Dim accountNum_Col As Integer
    Dim tranCode_Col As Integer
    Dim amount_Col As Integer
    Dim date_Col As Integer
    Dim descript_Col As Integer
    Dim dif_Col As Integer
    Dim pmt_Col As Integer
    Dim status_Col As Integer
    Dim borrowerPrinAmt As Double
    Dim payoff_Flag As String
    Dim modFee_Flag As String
    Dim mod_Dist As String
    Dim GL_Other As String
    Dim other_Tran As String
    Set WS = Sheets("Payment Sheet")
    
    ' Constant GL entry values
    GL_Dist = "19"
    GL_Account = "1234567890"
    GL_Tran = "44"
    GL_Date = Date
    borrower_Dist = "22"
    borrower_Tran = "7"
    nonaccrual_Tran = "14"
    principal_Tran = "15"
    closed_Tran = "2"
    payoff_Flag = "59"
    modFee_Flag = "9"
    
    mod_Dist = "18"
    GL_Other = "6135715308"
    other_Tran = "3"
    
    ' Column reference numbers
    distCode_Col = 1
    accountNum_Col = 2
    tranCode_Col = 3
    amount_Col = 4
    date_Col = 5
    descript_Col = 6
    pmt_Col = 7
    dif_Col = 8
    status_Col = 9
    
    ' Double counter to account for additional entries
    counter = (counter * 2)

    Dim i As Long
    i = 1
    Do
        ' These values are tied to the current Borrower Row Value
        GL_Amount = WS.Cells(i, amount_Col).Value
        GL_Description = WS.Cells(i, accountNum_Col).Value
        
        ' Insert row to create GL Entry
        insertRow rowNum:=i
            'Previous Insert Code  = WS.Rows(i).EntireRow.Insert
        
        ' Assign GL values from constants above & variables within loop
        WS.Cells(i, distCode_Col).Value = GL_Dist
        WS.Cells(i, accountNum_Col).Value = GL_Account
        WS.Cells(i, tranCode_Col).Value = GL_Tran
        WS.Cells(i, amount_Col).Value = GL_Amount
        WS.Cells(i, date_Col).Value = GL_Date
        WS.Cells(i, descript_Col).Value = GL_Description
        
        ' Set Distribution Code & Tran Code for Borrower
        ' Variables for secondary row for borrower principal allocation
        Dim borrowerRow As Integer
        Dim secondaryRow As Integer
        Dim borrowerPayment As Double
        
        borrowerRow = i + 1
        secondaryRow = i + 2
        
        WS.Cells(borrowerRow, distCode_Col).Value = borrower_Dist
        ' Check for Non-Accrual or Closed Status
        If WS.Cells(borrowerRow, status_Col).Value = nonaccrual_Tran Then
            WS.Cells(borrowerRow, tranCode_Col).Value = nonaccrual_Tran
        ElseIf WS.Cells(borrowerRow, status_Col).Value = 0 Then
            WS.Cells(borrowerRow, tranCode_Col).Value = borrower_Tran
        ElseIf WS.Cells(borrowerRow, status_Col).Value = closed_Tran Then
            WS.Cells(borrowerRow, tranCode_Col).Value = closed_Tran
        End If
        
        
        ' Look for Payoff / Mod Fee Flags
        If WS.Cells(borrowerRow, status_Col).Value = payoff_Flag Then
            WS.Cells(borrowerRow, tranCode_Col).Value = payoff_Flag
        ElseIf WS.Cells(borrowerRow, status_Col).Value = modFee_Flag Then
            WS.Cells(borrowerRow, descript_Col).Value = WS.Cells(borrowerRow, accountNum_Col).Value
            WS.Cells(borrowerRow, distCode_Col).Value = mod_Dist
            WS.Cells(borrowerRow, accountNum_Col).Value = GL_Other
            WS.Cells(borrowerRow, tranCode_Col).Value = other_Tran
            WS.Cells(borrowerRow, amount_Col).Value = GL_Amount
            WS.Cells(borrowerRow, descript_Col).Value = GL_Description
        End If
        
        
        ' Determine if Borrower Payment amount needs to be split into two rows
        If (WS.Cells(borrowerRow, dif_Col).Value > 0) And (WS.Cells(borrowerRow, tranCode_Col) <> payoff_Flag) Then
            ' Determine row values, insert row, and adjust original payment allocation
            insertRow rowNum:=secondaryRow
            borrowerPrinAmt = WS.Cells(borrowerRow, dif_Col).Value
            borrowerPayment = WS.Cells(borrowerRow, pmt_Col).Value

            ' Insert Values for secondary row
            WS.Cells(secondaryRow, distCode_Col).Value = borrower_Dist
            WS.Cells(secondaryRow, accountNum_Col).Value = WS.Cells(secondaryRow - 1, accountNum_Col).Value
            WS.Cells(secondaryRow, tranCode_Col).Value = principal_Tran
            WS.Cells(secondaryRow, amount_Col).Value = borrowerPrinAmt
            WS.Cells(secondaryRow, date_Col).Value = WS.Cells(secondaryRow - 1, date_Col).Value

            ' Insert Payment Value for borrower row
            WS.Cells(borrowerRow, amount_Col).Value = borrowerPayment

            ' Move i in order to account for the secondary row creation
            i = borrowerRow
            ' Increase counter to account for the additional row
            counter = counter + 1
        End If
        
    i = i + 2
    Loop While i <= counter
End Sub
```
_WorksheetExists_
```
Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function
```
_checkGL_
```
Sub checkGL(ByVal checkSheet As String)

' Confirm worksheet passed exists
Dim wkSheet As Boolean
wkSheet = WorksheetExists(checkSheet)

If wkSheet = True Then
    Dim i As Long                   ' Looping variable
    Dim otherBorrowers As Variant   ' Array of borrowers in ORE or other income
    Dim rowCount As Integer         ' Determine size of range
    Dim lowerBound As Variant       ' Bottom of array
    Dim upperBound As Variant       ' Top of array
    Dim rng As String               ' Array as a string
    
' Create Array of ORE Borrowers
    ' Count for lowerBound
    With Sheets(checkSheet)
        rowCount = .Cells(.Rows.Count, "B").End(xlUp).Row
    End With
        
    lowerBound = Cells(rowCount, 2).Address
    upperBound = "$B$1"
    ' combine lower & upper to create a string for range object
    rng = upperBound & ":" & lowerBound
    ' Fill array
    otherBorrowers = Sheets(checkSheet).Range(rng)

    counter = 0         ' Reset counter to pass to rowCounter
    rowCounter counter:=rowCount, sheetName:="Payment Sheet"
    
    For i = LBound(otherBorrowers, 1) To UBound(otherBorrowers, 1)
        otherBorrowers(i, 1) = CStr(otherBorrowers(i, 1))
    Next

' Determine if any payments in the payment sheet are ORE loans
    Dim WS As Worksheet
    Dim isOtherBorrower As Variant
    Dim loanNumDub As Double
    Dim loanNumString As String
    Dim distCode_Col As Integer
    Dim accountNum_Col As Integer
    Dim tranCode_Col As Integer
    Dim descript_Col As Integer
    Dim status_Col As Integer
    Dim stringOne As String
    Dim stringTwo As String
    Dim GL_Dist As String
    Dim GL_Account As String
    Dim GL_Tran As String
    Dim GL_Description As String
    
    Set WS = Sheets("Payment Sheet")
    
    distCode_Col = 1
    accountNum_Col = 2
    tranCode_Col = 3
    descript_Col = 6
    status_Col = 9
    
    ' Set GL variable values
    GL_Dist = "18"
    GL_Tran = "3"
    If checkSheet = "ORE Borrowers" Then
        GL_Account = "7410852096"
    ElseIf checkSheet = "Other Income" Then
        GL_Account = "6135715308"
    ElseIf checkSheet = "ORE" Then
        GL_Account = "0369025801"
    End If
    
    
    ' Initialize i for loop
    i = 1
    ' Single-Column array's must be transposed so the Filter
    ' function doesn't throw a run time error
    otherBorrowers = Application.Transpose(otherBorrowers)
    
    Do
        ' Determine description for ORE GL entry
        stringOne = Trim(WS.Cells(i, 6).Value)
        stringTwo = Trim(CStr(WS.Cells(i, 2).Value))
        GL_Description = stringOne & " " & stringTwo
        
        ' Analyze loan number and compare it to ORE number array
        loanNumDub = WS.Cells(i, 2).Value
        loanNumString = CStr(loanNumDub)
        isOtherBorrower = Filter(otherBorrowers, loanNumString, True)
        
        ' UBound value returns 0 or higher int value if the array is found
        If UBound(isOtherBorrower) >= 0 And loanNumString <> "" Then  ' If account number is ORE
            WS.Cells(i, distCode_Col).Value = GL_Dist
            WS.Cells(i, accountNum_Col).Value = GL_Account
            WS.Cells(i, tranCode_Col).Value = GL_Tran
            WS.Cells(i, descript_Col).Value = GL_Description
        ' UVound returns -1 if the value passed by the array function is
        ' not the correct string value
        ElseIf UBound(isOtherBorrower) < 0 Then
        Else
            MsgBox ("ruh roh... looks like there was an error")
        End If

    i = i + 1
    Loop While i <= rowCount
Else
    MsgBox ("There was an error in calling functions to check for ORE loans and Other Income loans")
End If

End Sub
```
_remove redundancy_
```
Sub remove_redundancy()
' Determine size of sheet
    Dim rowCount As Integer
    rowCounter counter:=rowCount, sheetName:="Payment Sheet"

    Dim primaryRow_amount As Variant
    Dim mainRow_amount As Variant
    Dim secondaryRow_amount As Variant

    Dim WS As Worksheet
    Set WS = Sheets("Payment Sheet")

    Dim i As Long
    Dim mainRow As Long
    Dim secondaryRow As Long
    i = 1


' Loop through entries and determine which rows to adjust based on whether
' they are credit or debit entries, whether they are divided, they're GL
' entries, and after moving the entry amounts accordingly
    Do
    'mainRow = i + 1
    'secondaryRow = i + 2
        If WS.Cells(i, 1).Value = 19 Then   ' Check that the primary row is a GL debit entry
            If WS.Cells((i + 1), 1).Value = 18 And WS.Cells((i + 2), 1).Value = 18 Then     ' Confirm that there are two credit entires
                If WS.Cells((i + 1), 2).Value = WS.Cells((i + 2), 2).Value Then     ' Confirm the two credit entries go to the same account
                    primaryRow_amount = WS.Cells(i, 4).Value                ' Determine the value in the Debit Entry
        
                    mainRow_amount = WS.Cells((i + 1), 4).Value             ' Get the value from the first credit entry
                    secondaryRow_amount = WS.Cells((i + 2), 4).Value        ' Get the value from the second credit entry
                    mainRow_amount = mainRow_amount + secondaryRow_amount   ' Combine the value into the credit entry value variables
        
                    If primaryRow_amount = mainRow_amount Then              ' Verif the Debit Entry value equals the Credit Entry value
                        WS.Cells((i + 1), 4).Value = mainRow_amount         ' Set the value in the top credit entry
                        WS.Rows(i + 2).EntireRow.delete                     ' Delete row because it no longer should be entered
                        rowCount = rowCount - 1                             ' Remove 1 from rowCount because there is one less row
                    End If
                End If
            End If
        End If
    i = i + 1
    Loop While i <= rowCount
End Sub
```
_formatter (Print Pretty)_
```
Sub formatter()
    Dim WS As Worksheet
    Set WS = Sheets("Payment Sheet")
    
' Clear Formatting for whole sheet
    WS.Range("$A:$I").ClearFormats
    
' Set amount column as accounting format
    WS.Range("D:D").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
' Set date column as a date format
    WS.Range("E:E").NumberFormat = "mm/dd/yyyy"

' Set borders for readability
    Dim i As Long
    Dim rowCount As Integer
    Dim rng As String
    i = 1
    rowCounter rowCount, "Payment Sheet"
    
    
    Do
    ' Put a top border on the entries that have 19 distribution code
        If WS.Cells(i, 1).Value = 19 Then
            rng = "A" & i & ":F" & i
            WS.Range(rng).Borders(xlEdgeTop).LineStyle = xlContinuous
            WS.Range(rng).Borders(xlEdgeTop).Weight = xlThick
            WS.Range(rng).Interior.ColorIndex = 15
        End If
        
        If WS.Cells(i, 1).Value <> 19 Then
            rng = "A" & i & ":F" & i
            WS.Range(rng).Interior.ColorIndex = 35
        End If
        
    i = i + 1
    Loop While i <= rowCount

    ' Put border on the right side to segement the manual formulas from the actual entries
    WS.Range("F1:F" & rowCount).Borders(xlEdgeRight).LineStyle = xlContinuous
    WS.Range("F1:F" & rowCount).Borders(xlEdgeRight).Weight = xlThick
    ' Put border on the bottom of the entries
    WS.Range("A" & rowCount & ":F" & rowCount).Borders(xlEdgeBottom).LineStyle = xlContinuous
    WS.Range("A" & rowCount & ":F" & rowCount).Borders(xlEdgeBottom).Weight = xlThick
End Sub
```
_create wire entries_
```
Sub create_wire_entries()
    Dim WS As Worksheet
    Set WS = Sheets("Payment Sheet")
    
    ' Get user input for the number of wires
    numberOfWiresVar = Application.InputBox(prompt:="How many wires will you be sending today?", Type:=1)
    
    ' Check if the number of wires isn't a typo (typed wire amount instead of # of wires)
    If numberOfWiresVar >= 6 Then
        Dim Ans As Integer
        Dim text As String
        text = "Are you sure you want to send " & numberOfWiresVar & " wires?"
        Ans = MsgBox(text, vbYesNo)
        
        If Ans = vbNo Then
            MsgBox ("No wire entries will be created. Please start over to create a payment sheet")
            numberOfWiresVar = 0    ' This will not pass the the next if statement,
        End If
    End If
    
    
    ' Confirm the number of wires is greater than or equal to 1
    If numberOfWiresVar >= 1 Then
        Dim i As Long               ' Looper Value
        Dim wire_amount As Double   ' Wire Amount Value
        ' Citi 1834927351 values for entries
        Const citi_dist As String = "19"
        Const citi_account As String = "1834927351"
        Const citi_tran As String = "44"
        ' Cash 9157583402 values for entries
        Const cash_dist As String = "18"
        Const cash_account As String = "9157583402"
        Const cash_tran As String = "3"
    
        Do While i < numberOfWiresVar
            wireAmount wire_amount
            If wire_amount >= 0 And Not wire_amount = False Then
                Dim rowCount As Integer
                rowCounter rowCount, "Payment Sheet"
                ' Set all citi GL transaction values
                WS.Cells(rowCount + 1, 1).Value = citi_dist     ' Distribution Value in Column A
                WS.Cells(rowCount + 1, 2).Value = citi_account  ' Account Value in Column B
                WS.Cells(rowCount + 1, 3).Value = citi_tran     ' Tran Code in Column C
                WS.Cells(rowCount + 1, 4).Value = wire_amount   ' Wire Amount in Column D
                WS.Cells(rowCount + 1, 5).Value = Date          ' Current date for wire in Column E
                ' No description set
                
                ' Set all cash collection GL transaction values
                WS.Cells(rowCount + 2, 1).Value = cash_dist     ' Distribution Value in Column A
                WS.Cells(rowCount + 2, 2).Value = cash_account  ' Account Value in Column B
                WS.Cells(rowCount + 2, 3).Value = cash_tran     ' Tran Code in Column C
                WS.Cells(rowCount + 2, 4).Value = wire_amount   ' Wire Amount in Column D
                WS.Cells(rowCount + 2, 5).Value = Date          ' Current date for wire in Column E
                ' No description set
            End If
        i = i + 1
        Loop
    End If
       
    
End Sub
```
_wireAmount_
```
Sub wireAmount(ByRef sendAmount As Double)
    sendAmount = Application.InputBox(prompt:="Input Wire Amount", Type:=1)
End Sub
```
_errorChecker_
```
Sub errorChecker()
    Dim WS As Worksheet
    Set WS = Sheets("Payment Sheet")

    Dim rowCount As Integer
    rowCounter rowCount, "Payment Sheet"

    ' 19 & loan number
    ' 22 & 1234567890
    ' Difference column is positive
    
    Dim i As Long
    i = 1
    
    Do
        ' Check if the entry is a GL entry and confirm the appropriate distribution code, account number, and tran code
        If WS.Cells(i, 1).Value = 19 And ((WS.Cells(i, 2).Value <> 1234567890# And WS.Cells(i, 2).Value <> 9157583402# _
                                            And WS.Cells(i, 2).Value <> 1834927351#) _
                                            Or WS.Cells(i, 3).Value <> 44) Then
                                            ' Loan number must 478, 122, 120 in order to not return false
            WS.Range("A" & i & ":F" & i).Interior.ColorIndex = 3
        End If
        
        ' Check if the entry is a credit GL entry and confirm appropriate distribution code, account number, and tran code
        If WS.Cells(i, 1).Value = 18 And ((WS.Cells(i, 2).Value <> 9157583402# And WS.Cells(i, 2).Value <> 7410852096# _
                                            And WS.Cells(i, 2).Value <> 6135715308# And WS.Cells(i, 2).Value <> 0369025801#) _
                                            Or WS.Cells(i, 3).Value <> 3) Then
                                            ' Loan number must 122, 6952, 682 in order to not return false
            WS.Range("A" & i & ":F" & i).Interior.ColorIndex = 3
        End If
        
        
        ' Check if the entry is a loan payment, if it is the loan number <> 1234567890, and the tran code must be appropriate
        If WS.Cells(i, 1).Value = 22 And _
            ((WS.Cells(i, 2).Value = 1234567890#) Or (WS.Cells(i, 3).Value <> 14 And WS.Cells(i, 3).Value <> 7 And _
                WS.Cells(i, 3).Value <> 2 And WS.Cells(i, 3).Value <> 15 And WS.Cells(i, 3).Value <> 59)) Then
            WS.Range("A" & i & ":F" & i).Interior.ColorIndex = 3
        End If
        
        ' Determine if the payment less the amount paid column is positive.
        ' Create entry function should effectively guarantee this will never happen
        ' but not a problem to check.
        If WS.Cells(i, 8).Value > 0 Then
            WS.Cells(i, 8).Interior.ColorIndex = 3
        End If
    
    i = i + 1
    Loop While i <= rowCount


End Sub
```
_exportSheet_
```
Sub exportSheet()
    Dim WS As Worksheet
    Set WS = Sheets("Payment Sheet")
    
    ' Change name of current sheet
    WS.Name = "Payment Info Sheet"
    Set infoSheet = Sheets("Payment Info Sheet")
    
    ' Add sheet
    Sheets.Add.Name = "Payment Sheet"
    Set pmtSheet = Sheets("Payment Sheet")
    
    Dim rowCount As Integer
    rowCount = 0
    
    rowCounter counter:=rowCount, sheetName:="Payment Info Sheet"
    
    infoSheet.Range("A1:F" & rowCount).Copy _
        Destination:=pmtSheet.Range("A1:F" & rowCount)
          
End Sub
```
