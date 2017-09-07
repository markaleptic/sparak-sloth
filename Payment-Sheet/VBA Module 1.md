# Insert Sub Routine to move exported Sparak data into Import Sheete
```
Sub insert_import()
    
' ***************************************************************
' *                                                             *
' *  Project:       PART 1a: C.R.E.A.M. Project                 *
' *  Created by:    Mark Allred                                 *
' *  Purpose:       This code is part of the cream project      *
' *                 that automates B49 daily book keeping.      *
' *                 This part of the project takes Sparak's     *
' *                 exported data and pastes it into the        *
' *                 "Import Sheet" of the Payment Sheeet        *
' *                 workbook.                                   *
' *                                                             *
' ***************************************************************
    
    Set pmt_wkb = ThisWorkbook
    
    ' Set directory to folder where I leave Sparak Exports
    ChDrive "O"
    new_dir = "O:\Accounting\Accounting\Daily Balancing and Entries\Payments\Payment Imports\"
    ChDir (new_dir)
    
    ' Find file via File Explorer
    Dim MyFile As Variant
    MyFile = Application.GetOpenFilename()
    
    ' Open file as a xlDelimited file with the delimiter set to semicolon
    ' If cancel is selected then Exit Sub
    If MyFile = False Then
        Exit Sub
    Else
        Workbooks.OpenText Filename:=MyFile, _
                        DataType:=xlDelimited, _
                        Semicolon:=True
    End If
    
    ' Copy import data from the selected sheet into the worksheet Import Sheet in the Payment Sheet workbook
    ActiveWorkbook.Worksheets(1).Range("A1:H756").Copy Destination:=pmt_wkb.Worksheets("Import Sheet").Range("A2:H757")
    
    ' Close the import data workbook
    ActiveWorkbook.Close
        
End Sub
```
