'***************************************
'************** MODULE 2 *************
'***************************************
'
' This module contains 2 subroutines (do something on the sheet), used to manually launch a full verification of the whole sheet
'
' - 2 subroutines: VerifyLibrary, VerifySCLibrary
'
'***************************************


Sub VerifyLibrary()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("library")

    ' Find the last row in column B
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    ' Find the last column in row 2
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.count).End(xlToLeft).Column
    
    ' Save the results in the range starting from B2
    Dim resultRange As range
    
    ' Subtract 1 row to exclude first row to exclude B2
    ' Subtract 3 columns to exclude column A and annotator+date
    Set resultRange = ws.range("B2").Resize(lastRow - 1, lastCol - 3)

    Application.Run "Feuil1.Worksheet_Change", resultRange
    
End Sub



Sub VerifySCLibrary()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("sc-library")

    ' Find the last row in column B
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    ' Find the last column in row 2
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.count).End(xlToLeft).Column
    
    ' Save the results in the range starting from B2
    Dim resultRange As range
    
    ' Subtract 1 row to exclude first row to exclude B2
    ' Subtract 3 columns to exclude column A and annotator+date
    Set resultRange = ws.range("B2").Resize(lastRow - 1, lastCol - 3)

    Application.Run "Sheet4.Worksheet_Change", resultRange
    
End Sub



