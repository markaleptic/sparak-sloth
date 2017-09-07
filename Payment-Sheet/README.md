# Payment Sheet

To see the VBA code modules that run this payment sheet, click here: [Module 1](https://github.com/markaleptic/sparak-sloth/blob/master/Payment-Sheet/VBA%20Module%201.md) and [Module 2](https://github.com/markaleptic/sparak-sloth/blob/master/Payment-Sheet/VBA%20Module%202.md).

Below is a screenshot of the Payment Sheet main view. The gridded area is where you paste entries from the check register.

![Image of Payment Sheet](https://github.com/markaleptic/sparak-sloth/blob/master/Payment-Sheet/Payment_Sheet.PNG "Payment Sheet screenshot")

_Normal Pmt Formula_
```
=IFERROR(IF(VLOOKUP(A3,'Import Sheet'!$A$2:$E$757,4,FALSE)=0,"",
VLOOKUP(A3,'Import Sheet'!$A$2:$E$757,4,FALSE)),"ERROR")
```
_Difference Formula_
```
=IFERROR(F4-G4,0)
```
_Status Formula_
```
=IFERROR(IF(VLOOKUP(A4,'Import Sheet'!$A$2:$I$757,9,FALSE)=TRUE,"2",IF((VLOOKUP(A4,'Import
Sheet'!$A$2:$E$757,3,FALSE))=11,"11",IF(VLOOKUP(A4,'Import Sheet'!$A$2:$E$757,3,FALSE)=18,"14",
VLOOKUP(A4,'Import Sheet'!$A$2:$E$757,3,FALSE)))),"ERROR")
```

--------------------------------------------------------------------------------------------------------------------
# Import Sheet
Below is a screenshot of the Import Sheet. This area is not handled by the individual, but the import module from [Module 1](https://github.com/markaleptic/sparak-sloth/blob/master/Payment-Sheet/VBA%20Module%201.md).
Private data is redacted, while other data is included to show the type of data provided by Sparak.

![Image of Import Sheet](https://github.com/markaleptic/sparak-sloth/blob/master/Payment-Sheet/Import_Sheet.PNG "Import Sheet screenshot")

_NA - Int Only Formula_

This code is used by the Payment Sheet status formula to determine if a loan is both Non-Accrual and Interest Only.
```
=IFERROR(IF((AND(F2=0,G2>0,C2<>11)),TRUE, FALSE),FALSE)
```
