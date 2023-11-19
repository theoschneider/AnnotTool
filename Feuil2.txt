Private Sub Worksheet_Change(ByVal Modified As range)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' --- GET ALL SHEETS ---
    Dim exp_Sheet As Worksheet
    Set exp_Sheet = Sheets("experiment")
    Dim lib_Sheet As Worksheet
    Set lib_Sheet = Sheets("library")
    
    ' --- GET ALL COLUMNS OF INTEREST ---
    Dim expID_col As String
    expID_col = Split(exp_Sheet.Cells(1, Application.Match("#experimentId", Rows(1), 0)).address, "$")(1)
    Dim number_col As String
    number_col = Split(exp_Sheet.Cells(1, Application.Match("numberOfAnnotatedLibraries", Rows(1), 0)).address, "$")(1)
    Dim expStatus_col As String
    expStatus_col = Split(exp_Sheet.Cells(1, Application.Match("experimentStatus", Rows(1), 0)).address, "$")(1)
    
    
    ' --- CHECK EVERY MODIFIED CELL (MAIN LOOP) ---
    For Each Target In Modified
    
        ' --- GET COL AND ROW OF MODIFIED CELL ---
        Dim col As String
        col = Split(Cells(1, Target.Column).address, "$")(1)
        Dim row As Long
        row = Target.row
        
        ' --- DECLARE OTHER VALUES OF INTEREST ---
        Dim numberOfLibs As Long
        Dim exp_ID As String
        exp_ID = range(expID_col & row).Value
        
        
        If (col = expID_col) And (row > 1) Then
             ' --- COUNT LIBRARIES PART ---
             
            ' --- Run the libraries counter
            numberOfLibs = Count_Libraries(exp_ID, exp_Sheet, lib_Sheet)
            
            ' --- Fill the cell with the number
            exp_Sheet.range(number_col & row).Value = CStr(numberOfLibs) & " libraries"
    
        End If
        
        
        If (col = expID_col Or col = expStatus_col) And (row > 1) Then
            ' --- EXPERIMENT STATUS ---
            
            ExperimentStatus Worksheets("experiment").range(expStatus_col & row)
        
        End If
    
    
    Next Target
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
