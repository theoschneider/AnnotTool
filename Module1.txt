'***************************************
'************** MODULE 1 *************
'***************************************
'
' This module contains functions (return something) and subroutines (do something on the sheet)
'
' - 6 functions: FindMatchingValues, SortArray, IsStringInArray, Count_Libraries, ProtocolStatus, SCProtocolStatus
' - 3 formatting subroutines: Warning, Fatal, ClearFormatting
' - 3 subroutines: ExperimentStatus, AnnotationStatus, BiologicalStatus
'
'***************************************


Public Sub Warning(cell As range)

    ' Fill the cell in yellow
    cell.Interior.Color = RGB(255, 255, 155)
    
    ' Fill the first col in yellow
    Cells(cell.row, 1).Interior.Color = RGB(255, 255, 155)
    
    ' Find header and put it in first column
    Dim header As String
    header = Cells(1, cell.Column).Value
    
    If InStr(1, Cells(cell.row, 1).Value, header, vbTextCompare) = 0 Then
        Cells(cell.row, 1).Value = Cells(cell.row, 1).Value + header + ", "
    End If

End Sub



Public Sub Fatal(cell As range)

    ' Fill the cell in red
    cell.Interior.Color = RGB(255, 0, 0)
    
    ' Change text to red
    cell.Font.Color = RGB(255, 255, 255)
    
    ' Fill the first col in red
    Cells(cell.row, 1).Interior.Color = RGB(255, 255, 155)
    
    ' Find header and put it in first column
    Dim header As String
    header = Cells(1, cell.Column).Value
    
    If InStr(1, Cells(cell.row, 1).Value, header, vbTextCompare) = 0 Then
        Cells(cell.row, 1).Value = Cells(cell.row, 1).Value + header + ", "
    End If

End Sub



Public Sub ClearFormatting(cell As range)

    ' Remove color from cell
    cell.Interior.Color = xlNone
    
    ' Find header and remove it from first column
    Dim header As String
    header = Cells(1, cell.Column).Value
    Cells(cell.row, 1).Value = Replace(Cells(cell.row, 1).Value, header + ", ", "")
    
    ' If cell in first col is empty, remove its color
    If Cells(cell.row, 1).Value = "" Then
        Cells(cell.row, 1).Interior.Color = xlNone
    End If

End Sub



Public Sub ExperimentStatus(cell As range)

    Dim lastRow As Long
    Dim arr As Variant
    
    ' Find the last row of possible terms
    lastRow = Worksheets("settings").Cells(Rows.count, 1).End(xlUp).row
    
    arr = Application.Transpose(Worksheets("settings").range("A2:A" & lastRow).Value)
    
    If IsStringInArray(cell.Value, arr) Then
    
        cell.Validation.Delete
    
    Else
        ' Fill the cell with dropdown
        With cell.Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=Join(arr, ",")
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = False
        End With

    End If

End Sub



Public Sub AnnotationStatus(cell As range)
    
    Dim lastRow As Long
    Dim arr As Variant
    
    ' Find the last row of possible terms
    lastRow = Worksheets("settings").Cells(Rows.count, 2).End(xlUp).row
    
    arr = Application.Transpose(Worksheets("settings").range("B2:B" & lastRow).Value)
    
    If IsStringInArray(cell.Value, arr) Then
        
        ClearFormatting cell
        cell.Validation.Delete
        
    Else
    
        Warning cell
        ' Fill the cell with dropdown
        With cell.Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=Join(arr, ",")
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = False
        End With
    
    End If

End Sub



Public Sub BiologicalStatus(cell As range)

    Dim lastRow As Long
    Dim arr As Variant
    
    ' Find the last row of possible terms
    lastRow = Worksheets("settings").Cells(Rows.count, 3).End(xlUp).row
    
    arr = Application.Transpose(Worksheets("settings").range("C2:C" & lastRow).Value)
    
    If IsStringInArray(cell.Value, arr) Then
    
        ClearFormatting cell
        cell.Validation.Delete
        
    Else
    
        Warning cell
        ' Fill the cell with dropdown
        With cell.Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=Join(arr, ",")
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = False
        End With
    
    End If

End Sub



Function ProtocolStatus(protocol As String, protocolType As String, RNASelection As String, dbsheet As Worksheet) As Variant

    Dim dbLastRow As Long
    Dim rowCount As Long
    Dim searchResults As Variant
    Dim checkedProtocols() As Variant
    Dim protocol_values() As Variant
    Dim type_values() As Variant
    Dim RNASel_values() As Variant
    Dim i As Long
    
    ' Get the last row number in the "protocols-db" sheet
    dbLastRow = dbsheet.Cells(dbsheet.Rows.count, "A").End(xlUp).row
    
    ' Fill the data arrays with the 3 columns of interest
    protocol_values = Application.Transpose(dbsheet.range("A2:A" & dbLastRow))
    type_values = Application.Transpose(dbsheet.range("D2:D" & dbLastRow))
    RNASel_values = Application.Transpose(dbsheet.range("C2:C" & dbLastRow))
    
    ' Loop through each row in dbsheet and check the conditions
    rowCount = 0
    
    ReDim searchResults(1 To 3, 0 To 0)
    
    For i = 1 To UBound(protocol_values)
        If InStr(1, protocol_values(i), protocol, vbTextCompare) > 0 _
            And (Not IsStringInArray(protocol, checkedProtocols)) _
            And InStr(1, type_values(i), protocolType, vbTextCompare) > 0 _
            And InStr(1, RNASel_values(i), RNASelection, vbTextCompare) > 0 Then
            ' Add the matching values from columns A and B to the array
            rowCount = rowCount + 1
            ReDim Preserve checkedProtocols(1 To rowCount)
            checkedProtocols(rowCount) = protocol_values(i)
            ReDim Preserve searchResults(1 To 3, 1 To rowCount)
            searchResults(1, rowCount) = protocol_values(i)
            searchResults(2, rowCount) = type_values(i)
            searchResults(3, rowCount) = RNASel_values(i)
        End If
    Next i
    
    ProtocolStatus = searchResults
    
End Function



Function SCProtocolStatus(protocol As String, protocolType As String, dbsheet As Worksheet) As Variant

    Dim dbLastRow As Long
    Dim rowCount As Long
    Dim searchResults As Variant
    Dim checkedProtocols() As Variant
    Dim protocol_values() As Variant
    Dim type_values() As Variant
    Dim i As Long
    
    ' Get the last row number in the "protocols-db" sheet
    dbLastRow = dbsheet.Cells(dbsheet.Rows.count, "A").End(xlUp).row
    
    ' Fill the data arrays with the 2 columns of interest
    protocol_values = Application.Transpose(dbsheet.range("A2:A" & dbLastRow))
    type_values = Application.Transpose(dbsheet.range("B2:B" & dbLastRow))
    
    ' Loop through each row in dbsheet and check the conditions
    rowCount = 0
    
    ReDim searchResults(1 To 2, 0 To 0)
    
    For i = 1 To UBound(protocol_values)
        If InStr(1, protocol_values(i), protocol, vbTextCompare) > 0 _
            And (Not IsStringInArray(protocol, checkedProtocols)) _
            And InStr(1, type_values(i), protocolType, vbTextCompare) > 0 Then
            ' Add the matching values from columns A and B to the array
            rowCount = rowCount + 1
            ReDim Preserve checkedProtocols(1 To rowCount)
            checkedProtocols(rowCount) = protocol_values(i)
            ReDim Preserve searchResults(1 To 2, 1 To rowCount)
            searchResults(1, rowCount) = protocol_values(i)
            searchResults(2, rowCount) = type_values(i)
        End If
    Next i
    
    SCProtocolStatus = searchResults
    
End Function



Function SortArray(arr As Variant, ByVal sortRowIndex As Long) As Variant
    Dim numCols As Long
    Dim numRows As Long
    Dim i As Long, j As Long, k As Long
    Dim temp As String
    
    ' Get the number of columns/rows in the array
    numCols = UBound(arr, 2)
    numRows = UBound(arr, 1)
    
    ' Perform the sorting using Bubble Sort
    For i = 1 To numCols - 1
        For j = 1 To numCols - 1
            If Len(arr(sortRowIndex, j)) > Len(arr(sortRowIndex, j + 1)) Then
                ' Swap columns directly (for all rows)
                For k = 1 To numRows
                    temp = arr(k, j)
                    arr(k, j) = arr(k, j + 1)
                    arr(k, j + 1) = temp
                Next k
            End If
        Next j
    Next i
    
    If UBound(arr, 2) > 15 Then
        ReDim Preserve arr(1 To 2, 1 To 15)
    
    End If
    
    SortArray = arr

End Function



Function FindMatchingValues(Id As String, Term As String, Species As String, RefSheet As Worksheet) As Variant
    Dim lastRowRefSheet As Long
    Dim rowCount As Long
    Dim matchingValuesArray As Variant
    Dim ID_col() As Variant
    Dim Term_col() As Variant
    Dim Species_col() As Variant
    Dim i As Long
    
    ' Get the last row number in the "organ-db" sheet
    lastRowRefSheet = RefSheet.Cells(RefSheet.Rows.count, "A").End(xlUp).row
    
    ' Fill the data arrays with the 3 columns of interest
    ID_col = Application.Transpose(RefSheet.range("A2:A" & lastRowRefSheet))
    Term_col = Application.Transpose(RefSheet.range("B2:B" & lastRowRefSheet))
    Species_col = Application.Transpose(RefSheet.range("C2:C" & lastRowRefSheet))
    
    ' Loop through each row in RefSheet and check the conditions
    rowCount = 0
    
    ReDim matchingValuesArray(1 To 2, 0 To 0)
    
    For i = 1 To UBound(ID_col)
        If InStr(1, ID_col(i), Id, vbTextCompare) > 0 _
            And InStr(1, Term_col(i), Term, vbTextCompare) > 0 _
            And InStr(1, Species_col(i), Species, vbTextCompare) > 0 Then
            ' Add the matching values from columns A and B to the array
            rowCount = rowCount + 1
            ReDim Preserve matchingValuesArray(1 To 2, 1 To rowCount)
            matchingValuesArray(1, rowCount) = ID_col(i)
            matchingValuesArray(2, rowCount) = Term_col(i)
        End If
    Next i
    
    FindMatchingValues = matchingValuesArray

End Function



Function IsStringInArray(searchString As String, searchArray As Variant) As Boolean
    
    ' Check if array is empty
    If Len(Join(searchArray, "")) = 0 Then
        IsStringInArray = False
        Exit Function
    End If
    
    ' If not empty, proceed to search
    Dim ii As Long
    
    For ii = LBound(searchArray) To UBound(searchArray)
        If searchArray(ii) = searchString Then
            IsStringInArray = True
            Exit Function
        End If
    Next ii
    
    ' If not found, element is not in array
    IsStringInArray = False
    
End Function


Function Count_Libraries(expID As String, expSheet As Worksheet, libSheet As Worksheet) As Long
    
    ' Declare variables
    Dim counter As Long
    Dim exp_expID_col As String
    Dim lib_expID_col As String
    Dim lib_libID_col As String
    Dim lib_lastrow As Long
    Dim exp_lastrow As Long
    
    ' Get values of interest (columns and last rows)
    exp_expID_col = Split(expSheet.Cells(1, Application.Match("#experimentId", expSheet.Rows(1), 0)).address, "$")(1)
    lib_expID_col = Split(libSheet.Cells(1, Application.Match("experimentId", libSheet.Rows(1), 0)).address, "$")(1)
    lib_libID_col = Split(libSheet.Cells(1, Application.Match("#libraryId", libSheet.Rows(1), 0)).address, "$")(1)
    lib_lastrow = libSheet.Cells(libSheet.Rows.count, lib_libID_col).End(xlUp).row
    exp_lastrow = expSheet.Cells(expSheet.Rows.count, exp_expID_col).End(xlUp).row
    
    ' Declare 2 data arrays
    Dim libIDs() As Variant
    Dim expIDs() As Variant
    
    ' Save the data range in 2 arrays (library ID and experiment ID of library sheet)
    libIDs = Application.Transpose(libSheet.range(lib_libID_col & "2:" & lib_libID_col & lib_lastrow).Value)
    expIDs = Application.Transpose(libSheet.range(lib_expID_col & "2:" & lib_expID_col & lib_lastrow).Value)
    
    ' Initialize the counter
    counter = 0

    ' Loop through the 2D array and check the conditions
    For rowCounter = LBound(libIDs) To UBound(libIDs)
        If (Not libIDs(rowCounter) Like "[#]*") And (expIDs(rowCounter) = expID) Then
            ' Both conditions are met, so increment the counter
            counter = counter + 1
        End If
    Next rowCounter
    
    Count_Libraries = counter
    
End Function



