Private Sub Worksheet_Change(ByVal Modified As range)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' --- GET ALL SHEETS ---
    Dim librarySheet As Worksheet
    Set librarySheet = ThisWorkbook.Worksheets("library")
    Dim experimentSheet As Worksheet
    Set experimentSheet = ThisWorkbook.Worksheets("experiment")
    Dim dbsheet As Worksheet
    
    ' --- GET ALL COLUMNS OF THE SHEET ---
    Dim libID_col As String
    libID_col = Split(Cells(1, Application.Match("#libraryId", Rows(1), 0)).address, "$")(1)
    Dim expID_col As String
    expID_col = Split(Cells(1, Application.Match("experimentId", Rows(1), 0)).address, "$")(1)
    Dim platform_col As String
    platform_col = Split(Cells(1, Application.Match("platform", Rows(1), 0)).address, "$")(1)
    Dim SRSId_col As String
    SRSId_col = Split(Cells(1, Application.Match("SRSId", Rows(1), 0)).address, "$")(1)
    Dim anatId_col As String
    anatId_col = Split(Cells(1, Application.Match("anatId", Rows(1), 0)).address, "$")(1)
    Dim anatName_col As String
    anatName_col = Split(Cells(1, Application.Match("anatName", Rows(1), 0)).address, "$")(1)
    Dim stageId_col As String
    stageId_col = Split(Cells(1, Application.Match("stageId", Rows(1), 0)).address, "$")(1)
    Dim stageName_col As String
    stageName_col = Split(Cells(1, Application.Match("stageName", Rows(1), 0)).address, "$")(1)
    Dim anatAnnStatus_col As String
    anatAnnStatus_col = Split(Cells(1, Application.Match("anatAnnotationStatus", Rows(1), 0)).address, "$")(1)
    Dim anatBioStatus_col As String
    anatBioStatus_col = Split(Cells(1, Application.Match("anatBiologicalStatus", Rows(1), 0)).address, "$")(1)
    Dim stageAnnStatus_col As String
    stageAnnStatus_col = Split(Cells(1, Application.Match("stageAnnotationStatus", Rows(1), 0)).address, "$")(1)
    Dim sex_col As String
    sex_col = Split(Cells(1, Application.Match("sex", Rows(1), 0)).address, "$")(1)
    Dim strain_col As String
    strain_col = Split(Cells(1, Application.Match("strain", Rows(1), 0)).address, "$")(1)
    Dim genotype_col As String
    strain_col = Split(Cells(1, Application.Match("genotype", Rows(1), 0)).address, "$")(1)
    Dim Species_col As String
    Species_col = Split(Cells(1, Application.Match("speciesId", Rows(1), 0)).address, "$")(1)
    Dim proto_col As String
    proto_col = Split(Cells(1, Application.Match("protocol", Rows(1), 0)).address, "$")(1)
    Dim proto_type_col As String
    proto_type_col = Split(Cells(1, Application.Match("protocolType", Rows(1), 0)).address, "$")(1)
    Dim RNASelection_col As String
    RNASelection_col = Split(Cells(1, Application.Match("RNASelection", Rows(1), 0)).address, "$")(1)
    Dim sampleTitle_col As String
    sampleTitle_col = Split(Cells(1, Application.Match("sampleTitle", Rows(1), 0)).address, "$")(1)
    Dim annotatorCol As String
    annotatorCol = Split(Cells(1, Application.Match("annotatorId", Rows(1), 0)).address, "$")(1)
    Dim lastModifiedCol As String
    lastModifiedCol = Split(Cells(1, Application.Match("lastModificationDate", Rows(1), 0)).address, "$")(1)
    
    Dim exp_number_col As String
    exp_number_col = Split(experimentSheet.Cells(1, Application.Match("numberOfAnnotatedLibraries", experimentSheet.Rows(1), 0)).address, "$")(1)
    Dim exp_expID_col As String
    exp_expID_col = Split(Cells(1, Application.Match("#experimentId", experimentSheet.Rows(1), 0)).address, "$")(1)
    
    
    ' --- GET LASTROW ---
    Dim lib_lastrow As Long
    lib_lastrow = Cells(Rows.count, libID_col).End(xlUp).row
    
    ' --- INITIATE LIST OF COL THAT CAN'T BE EMPTY ---
    ' (all the status columns are already getting checked)
    Dim mandatory As Variant
    mandatory = Array(expID_col, platform_col, SRSId_col, anatId_col, anatName_col, stageId_col, stageName_col, _
    sex_col, strain_col, Species_col, proto_col, proto_type_col, RNASelection_col, sampleTitle_col)
    
    
    ' --- CHECK EVERY MODIFIED CELL (MAIN LOOP) ---
    For Each Target In Modified

        ' --- GET COL AND ROW OF MODIFIED CELL ---
        Dim col As String
        col = Split(Cells(1, Target.Column).address, "$")(1)
        Dim row As Long
        row = Target.row
        
        ' --- DECLARE OTHER VALUES OF INTEREST ---
        Dim userName As String
        Dim anatName As String
        Dim anatId As String
        Dim cellTypeId As String
        Dim cellTypeName As String
        Dim stageId As String
        Dim stageName As String
        Dim Species As String
        Dim nResults As Long
        Dim i As Long
        Dim mergedValuesArray() As Variant
        Dim splitted() As String
        Dim numberOfLibs As Long
        Dim exp_ID As String
        Dim count As Long
        Dim strains_data As Variant
        Dim species_data As Variant
        Dim RNASel_data As Variant
    
        
        If (range(libID_col & row).Value <> "") And (IsStringInArray(col, mandatory)) And (range(col & row).Value = "") Then
            Warning range(col & row)
        End If


        If (col <> annotatorCol) And (col <> lastModifiedCol) And (row > 1) Then
        ' --- ANNOTATOR PART (RUN EVERYTIME) ---
        
            ' Get the username of the person who made the change and update column, as well as date
            userName = Application.userName
            Cells(row, annotatorCol).Value = userName
            Cells(row, lastModifiedCol).Value = Date
                
        End If
        
        
        If (col = libID_col Or col = expID_col) And (row > 1) Then
             ' --- COUNT LIBRARIES PART ---
             
            exp_ID = range(expID_col & row).Value
            
            ' --- Run the libraries counter
            numberOfLibs = Count_Libraries(exp_ID, experimentSheet, librarySheet)
        
            ' --- Find the correct experiment row and fill it with the number
            Dim foundCell As range
            
            ' --- Find the search string in the column
            Set foundCell = experimentSheet.range(exp_expID_col & ":" & exp_expID_col).Find(What:=exp_ID, LookIn:=xlValues, LookAt:=xlWhole)
            
            ' Check if the search string was found
            If Not foundCell Is Nothing Then
                experimentSheet.range(exp_number_col & foundCell.row).Value = CStr(numberOfLibs) & " libraries"
                
            End If
    
        End If
        
        
        If (col = libID_col Or col = anatAnnStatus_col) And (row > 1) Then
            ' --- ANAT ANNOTATION STATUS PART ---
            AnnotationStatus Worksheets("library").range(anatAnnStatus_col & row)
        
        End If
        
        
        If (col = libID_col Or col = stageAnnStatus_col) And (row > 1) Then
            ' --- STAGE ANNOTATION STATUS PART ---
            AnnotationStatus Worksheets("library").range(stageAnnStatus_col & row)
        
        End If
        
        
        If (col = libID_col Or col = anatBioStatus_col) And (row > 1) Then
            ' --- ANAT BIOLOGICAL STATUS PART ---
            BiologicalStatus Worksheets("library").range(anatBioStatus_col & row)
        
        End If
        
        
        If (col = anatId_col Or col = anatName_col Or col = Species_col) And (row > 1) Then
        ' --- ORGAN PART ---
        
            ' Set references to the relevant sheet(s)
            Set dbsheet = Nothing
            Set dbsheet = ThisWorkbook.Worksheets("organ-db")
    
            ' Get the values for anatId, anatName and species
            anatName = CStr(librarySheet.range(anatName_col & row).Value)
            anatId = CStr(librarySheet.range(anatId_col & row).Value)
            Species = CStr(librarySheet.range(Species_col & row).Value)
            
            If (anatName = "") And (anatId = "") Then
                ' Remove the drop-down in both columns, to make sure
                librarySheet.range(anatId_col & row, anatName_col & row).Validation.Delete
                ' Remove formatting
                ClearFormatting range(anatId_col & row)
                ClearFormatting range(anatName_col & row)

            ElseIf librarySheet.range(anatId_col & row).Value Like "[A-Za-z]*[:]#* [A-Za-z]*" Then
            ' If something has been selected previously, fill it (ID Term)
                splitted = Split(librarySheet.range(anatId_col & row).Value, " ", 2)
                Cells(row, anatId_col).Value = splitted(0)
                Cells(row, anatName_col).Value = splitted(1)
                ' Remove the drop-down in both columns, to make sure
                librarySheet.range(anatId_col & row, anatName_col & row).Validation.Delete
                ' Remove formatting
                ClearFormatting range(anatId_col & row)
                ClearFormatting range(anatName_col & row)

            ElseIf librarySheet.range(anatName_col & row).Value Like "[A-Za-z]*[:]#* [A-Za-z]*" Then
                splitted = Split(librarySheet.range(anatName_col & row).Value, " ", 2)
                Cells(row, anatId_col).Value = splitted(0)
                Cells(row, anatName_col).Value = splitted(1)
                ' Remove the drop-down in both columns, to make sure
                librarySheet.range(anatId_col & row, anatName_col & row).Validation.Delete
                ' Remove formatting
                ClearFormatting range(anatId_col & row)
                ClearFormatting range(anatName_col & row)

            Else
                ' Run the search
                matchingValuesArray = FindMatchingValues(anatId, anatName, Species, dbsheet)
                nResults = UBound(matchingValuesArray, 2)
                
                ' If 0 options, put a warning
                If nResults = 0 Then
                    Warning range(anatId_col & row)
                    Warning range(anatName_col & row)
                ' If only 1 option, fill it directly
                ElseIf nResults = 1 Then
                    librarySheet.range(anatId_col & row).Value = matchingValuesArray(1, 1)
                    librarySheet.range(anatName_col & row).Value = matchingValuesArray(2, 1)
                    ' Remove formatting
                    ClearFormatting range(anatId_col & row)
                    ClearFormatting range(anatName_col & row)
                Else
                    ' Sort the array by length of string
                    matchingValuesArray = SortArray(matchingValuesArray, 2)
                
                    ' Transform the 2D array into a 1D array with each element separated by a space
                    ReDim mergedValuesArray(1 To UBound(matchingValuesArray, 2))
                    For i = 1 To UBound(mergedValuesArray)
                        mergedValuesArray(i) = matchingValuesArray(1, i) & " " & matchingValuesArray(2, i)
                    Next i
                    
                    ' Set the drop-down list validation, in both columns
                    With librarySheet.range(anatId_col & row, anatName_col & row).Validation
                        .Delete
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(mergedValuesArray, ",")
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .ShowInput = True
                        .ShowError = False
                    End With
                
                End If
            
            End If
            
        End If
        
        
        If (col = stageId_col Or col = stageName_col Or col = Species_col) And (row > 1) Then
        ' --- STAGE PART ---
        
            ' Set references to the relevant sheet(s)
            Set dbsheet = Nothing
            Set dbsheet = ThisWorkbook.Worksheets("stage-db")
    
            ' Get the values for cellTypeId, cellTypeName and species
            stageName = CStr(librarySheet.range(stageName_col & row).Value)
            stageId = CStr(librarySheet.range(stageId_col & row).Value)
            Species = CStr(librarySheet.range(Species_col & row).Value)
            
            If (stageId = "") And (stageName = "") Then
                ' Remove the drop-down in both columns, to make sure
                librarySheet.range(stageId_col & row, stageName_col & row).Validation.Delete
                ' Remove formatting
                ClearFormatting range(stageId_col & row)
                ClearFormatting range(stageName_col & row)

            ElseIf librarySheet.range(stageId_col & row).Value Like "[A-Za-z]*[:]#* [A-Za-z]*" Then
            ' If something has been selected previously, fill it (ID Term)
                splitted = Split(librarySheet.range(stageId_col & row).Value, " ", 2)
                Cells(row, stageId_col).Value = splitted(0)
                Cells(row, stageName_col).Value = splitted(1)
                ' Remove the drop-down in both columns, to make sure
                librarySheet.range(stageId_col & row, stageName_col & row).Validation.Delete
                ' Remove formatting
                ClearFormatting range(stageId_col & row)
                ClearFormatting range(stageName_col & row)

            ElseIf librarySheet.range(stageName_col & row).Value Like "[A-Za-z]*[:]#* [A-Za-z]*" Then
                splitted = Split(librarySheet.range(stageName_col & row).Value, " ", 2)
                Cells(row, stageId_col).Value = splitted(0)
                Cells(row, stageName_col).Value = splitted(1)
                ' Remove the drop-down in both columns, to make sure
                librarySheet.range(stageId_col & row, stageName_col & row).Validation.Delete
                ' Remove formatting
                ClearFormatting range(stageId_col & row)
                ClearFormatting range(stageName_col & row)

            Else
                ' Run the search
                matchingValuesArray = FindMatchingValues(stageId, stageName, Species, dbsheet)
                nResults = UBound(matchingValuesArray, 2)
                
                ' If 0 options, color cell
                If nResults = 0 Then
                    Warning range(stageId_col & row)
                    Warning range(stageName_col & row)
                ' If only 1 option, fill it directly
                ElseIf nResults = 1 Then
                    librarySheet.range(stageId_col & row).Value = matchingValuesArray(1, 1)
                    librarySheet.range(stageName_col & row).Value = matchingValuesArray(2, 1)
                    ' Remove formatting
                    ClearFormatting range(stageId_col & row)
                    ClearFormatting range(stageName_col & row)
                Else
                    ' Sort the array by length of string
                    matchingValuesArray = SortArray(matchingValuesArray, 2)
                    ' Transform the 2D array into a 1D array with each element separated by a space
                    ReDim mergedValuesArray(1 To UBound(matchingValuesArray, 2))
                    For i = 1 To UBound(mergedValuesArray)
                        mergedValuesArray(i) = matchingValuesArray(1, i) & " " & matchingValuesArray(2, i)
                    Next i
                    ' Set the drop-down list validation, in both columns
                    With librarySheet.range(stageId_col & row, stageName_col & row).Validation
                        .Delete
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(mergedValuesArray, ",")
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .ShowInput = True
                        .ShowError = False
                    End With
                End If
                
            End If
        
        End If
        
        
        If (col = sex_col Or col = Species_col) And (row > 1) Then
        ' --- SEX PART ---
        
            ' Find species of the line
            Species = CStr(librarySheet.range(Species_col & row).Value)
            
            sex_data = Application.Transpose(range(sex_col & "2:" & sex_col & lib_lastrow).Value)
            species_data = Application.Transpose(range(Species_col & "2:" & Species_col & lib_lastrow).Value)
            
            count = 0
            
            Dim compatible_sex() As Variant
            
            For i = LBound(sex_data) To UBound(sex_data)
                If (CStr(species_data(i)) = Species) And (Not IsStringInArray(CStr(sex_data(i)), compatible_sex)) And (Not sex_data(i) = "") Then
                    count = count + 1
                    ReDim Preserve compatible_sex(1 To count)
                    compatible_sex(UBound(compatible_sex)) = sex_data(i)
                End If
            Next i
            
            If (count > 0) Then
                With librarySheet.range(sex_col & row).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(compatible_sex, ",")
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = False
                End With
            End If
        
        End If
        
        
        If (col = strain_col Or col = Species_col) And (row > 1) Then
        ' --- STRAIN PART ---
        
            ' Find species of the line
            Species = CStr(librarySheet.range(Species_col & row).Value)
            
            strains_data = Application.Transpose(range(strain_col & "2:" & strain_col & lib_lastrow).Value)
            species_data = Application.Transpose(range(Species_col & "2:" & Species_col & lib_lastrow).Value)
            
            count = 0
            
            Dim compatible_strains() As Variant
            
            For i = LBound(strains_data) To UBound(strains_data)
                If (CStr(species_data(i)) = Species) And (Not IsStringInArray(CStr(strains_data(i)), compatible_strains)) And (Not strains_data(i) = "") Then
                    count = count + 1
                    ReDim Preserve compatible_strains(1 To count)
                    compatible_strains(UBound(compatible_strains)) = strains_data(i)
                End If
            Next i
            
            If (count > 0) Then
                With librarySheet.range(strain_col & row).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(compatible_strains, ",")
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = False
                End With
            End If
        
        End If
        
        
        If (col = proto_col Or col = proto_type_col Or col = RNASelection_col) And (row > 1) Then
        ' --- PROTOCOL PART ---
            
            Dim protocol As String
            protocol = range(proto_col & row).Value
            Dim protocolType As String
            protocolType = range(proto_type_col & row).Value
            Dim RNASelection As String
            RNASelection = range(RNASelection_col & row).Value
            Dim possible_protocols As Variant
            
            Set dbsheet = Nothing
            Set dbsheet = ThisWorkbook.Worksheets("protocols-db")
            
            ' Perform search
            searchResults = ProtocolStatus(protocol, protocolType, RNASelection, dbsheet)
            
            nResults = UBound(searchResults, 2)
            
            If range(proto_col & row).Value = "" Then
                range(proto_col & row).Validation.Delete
                ClearFormatting range(proto_col & row)
                ClearFormatting range(proto_type_col & row)
                ClearFormatting range(RNASelection_col & row)
            
            ElseIf nResults = 0 Then
                range(proto_col & row).Validation.Delete
                Warning range(proto_col & row)
                Warning range(proto_type_col & row)
                Warning range(RNASelection_col & row)
            
            ElseIf nResults = 1 Then
                range(proto_col & row).Validation.Delete
                range(proto_col & row).Value = searchResults(1, 1)
                ClearFormatting range(proto_col & row)
                ClearFormatting range(proto_type_col & row)
                ClearFormatting range(RNASelection_col & row)
                
                If (range(proto_type_col & row).Value = "") And (range(RNASelection_col & row).Value = "") Then
                    range(proto_type_col & row).Value = searchResults(2, 1)
                    range(RNASelection_col & row).Value = searchResults(3, 1)
                End If
            
            Else
                ClearFormatting range(proto_col & row)
                ClearFormatting range(proto_type_col & row)
                ClearFormatting range(RNASelection_col & row)
                ' Extract just the protocol names
                ReDim possible_protocols(1 To nResults)
                For i = 1 To nResults
                    possible_protocols(i) = searchResults(1, i)
                Next i
                
                ' Put them in a dropdown
                With range(proto_col & row).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(possible_protocols, ",")
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = False
                End With
                
            End If

        End If
        
        
    
    Next Target
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub



