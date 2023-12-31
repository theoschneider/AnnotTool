'***************************************
'************** MODULE 3 *************
'***************************************
'
' This module contains 1 subroutine to export all the other scripts in text format, before comitting to git.
' /!\ Make sure that exportPath exists /!\
'
' - 1 subroutine: ExportAllModules
'
'***************************************


Sub ExportAllModules()

    Dim moduleComponent As Object
    Dim exportPath As String
    Dim moduleName As String
    Dim fileNumber As Integer
    Dim componentType As Integer

    ' Set the export path (change this to your desired directory)
    exportPath = "/Users/theo/Desktop/Uni/Bgee/AnnotTool"


   ' Loop through all modules in the workbook
    For Each moduleComponent In ThisWorkbook.VBProject.VBComponents
    
        componentType = moduleComponent.Type

        ' Check if the component is a code module (Type 1) or a sheet module (Type 100)
        If componentType = 1 Or componentType = 100 Then
        
            ' Check if the code module is not empty
            If moduleComponent.CodeModule.CountOfLines > 0 Then
                
                On Error Resume Next
                moduleName = moduleComponent.Properties("Name").Value
                On Error GoTo 0
                
                ' Open a text file for writing
                fileNumber = FreeFile
                Open exportPath & "/" & moduleName & ".bas" For Output As fileNumber
    
                ' Write the code to the text file
                Print #fileNumber, moduleComponent.CodeModule.Lines(1, moduleComponent.CodeModule.CountOfLines)
    
                ' Close the text file
                Close fileNumber
            End If
            
        End If
    Next moduleComponent
    
End Sub





