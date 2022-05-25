Option Explicit

' Run GitSave() to export code and modules.
'
' Source:
' https://github.com/Vitosh/VBA_personal/blob/master/VBE/GitSave.vb
' The code below is slightly modified to include a list of
' modules you want to ignore.

    Dim ignoreList As Variant
    
    Dim parentFolder As String
    
    Const dirNameCode As String = "\Code"
    Const dirNameModules As String = "\Modules"
    
Sub GitSave()
    
    ignoreList = Array("Module1_to_ignore", "Module2_to_ignore")
    
    Call DeleteAndMake
    Call ExportModules
    Call PrintAllCode
    Call PrintModulesCode
    Call PrintAllContainers
    
End Sub

Sub DeleteAndMake()
    
    Dim childA As String
    Dim childB As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    parentFolder = ThisWorkbook.path
    childA = parentFolder & dirNameCode
    childB = parentFolder & dirNameModules
        
    On Error Resume Next
    fso.DeleteFolder childA
    fso.DeleteFolder childB
    On Error GoTo 0

    MkDir childA
    MkDir childB
    
End Sub

Sub PrintAllCode()
' Print all modules' code in one .vb file.
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        If Not IsStringInList(item.Name, ignoreList) Then
            
            lineToPrint = vbNewLine & "' MODULE: " & item.CodeModule.Name & vbNewLine
            If item.CodeModule.CountOfLines > 0 Then
                lineToPrint = lineToPrint & item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
            Else
                lineToPrint = lineToPrint & "' empty" & vbNewLine
            End If
'            Debug.Print lineToPrint
            textToPrint = textToPrint & vbCrLf & lineToPrint
            
        End If
    Next item
    
    Dim pathToExport As String: pathToExport = parentFolder & dirNameCode
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    SaveTextToFile textToPrint, pathToExport & "\all_code.vb"
    
End Sub

Sub PrintModulesCode()
' Print all modules' code in separate .vb files.

    Dim item  As Variant
    Dim lineToPrint As String
    Dim pathToExport As String
    
    pathToExport = parentFolder & dirNameCode
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        If Not IsStringInList(item.Name, ignoreList) Then
            
            If item.CodeModule.CountOfLines > 0 Then
                lineToPrint = item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
            
            Else
                lineToPrint = "' empty"
            End If
            
            
            
            If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
            Call SaveTextToFile(lineToPrint, pathToExport & "\" & item.CodeModule.Name & "_code.vb")
        
        End If
    Next item

End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        If Not IsStringInList(lineToPrint, ignoreList) Then
'            Debug.Print lineToPrint
            textToPrint = textToPrint & vbCrLf & lineToPrint
        End If
    Next item
    
    Dim pathToExport As String: pathToExport = parentFolder & dirNameCode
    SaveTextToFile textToPrint, pathToExport & "\all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = parentFolder & dirNameModules
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
        
        If Not IsStringInList(filePath, ignoreList) Then
        
            Select Case component.Type
                Case vbext_ct_ClassModule
                    filePath = filePath & ".cls"
                Case vbext_ct_MSForm
                    filePath = filePath & ".frm"
                Case vbext_ct_StdModule
                    filePath = filePath & ".bas"
                Case vbext_ct_Document
                    tryExport = False
            End Select
        
            If tryExport Then
'                Debug.Print unitsCount & " exporting " & filePath
                component.Export pathToExport & "\" & filePath
            End If
            
        End If
        
    Next

'    Debug.Print "Exported at " & pathToExport
    
End Sub

Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub

Function IsStringInList(ByVal whatString As String, whatList As Variant) As Boolean
' True if string is found in the list.
' Pass the list as Array.

    IsStringInList = Not (IsError(Application.Match(whatString, whatList, 0)))

End Function
