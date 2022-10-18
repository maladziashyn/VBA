Attribute VB_Name = "JFX_create"
Option Explicit

Const myFraction As Double = 0.005   ' 0.0067 = 0.67%
Const parZRow As Integer = 22
Const parFRow As Integer = 23

Dim parLRow As Integer
Dim ws As Worksheet
Dim c As Range

Dim defaultInstrument As String
Dim defaultPeriod As String
Dim algoTag As String
Dim auxIns As String

Dim strategyName As String, insAbbrev As String
Dim edHeadRow As Integer
Dim edSkipFRow As Integer
Dim edSkipLRow As Integer
Dim edVarsFRow As Integer
Dim edVarsLRow As Integer
Dim params() As Variant

Private Sub Create_JFX_file_Main()
    
    Dim i As Integer, j As Integer
    Dim replacedHeading As String
    Dim Rng As Range, cell As Range
    
    Application.ScreenUpdating = False
    Set ws = ActiveSheet
    Set c = ws.Cells
    strategyName = Strategy_name(c(1, 2))
    insAbbrev = Instrument_abbreviation(c(2, 2))
    algoTag = strategyName & insAbbrev
    parLRow = c(22, 1).End(xlDown).Row
' move Parameters to arr
    ReDim params(1 To parLRow - parFRow + 1, 1 To 2)
    For i = LBound(params, 1) To UBound(params, 1)
        For j = LBound(params, 2) To UBound(params, 2)
            params(i, j) = c(parZRow + i, j)
        Next j
'        Debug.Print params(i, 1) & " - " & params(i, 2)
    Next i
    edHeadRow = c(parLRow, 1).End(xlDown).Row
    edSkipFRow = edHeadRow + 2
    edSkipLRow = c(edSkipFRow, 1).End(xlDown).Row
    edVarsFRow = edSkipLRow + 2
    edVarsLRow = c(edVarsFRow, 1).End(xlDown).Row
' replace heading
    replacedHeading = Editor_new_heading(c(edHeadRow, 1), algoTag)
    c(edHeadRow, 2) = replacedHeading
' copy "Skip" part
    Set Rng = ws.Range(c(edSkipFRow, 1), c(edSkipLRow, 1))
    Rng.Copy c(edSkipFRow, 2)
' loop through "Variables" part
    Set Rng = ws.Range(c(edVarsFRow, 1), c(edVarsLRow, 1))
    For Each cell In Rng
        cell.Offset(0, 1) = Replaced_var(cell)
    Next cell
    ws.Range(c(edHeadRow, 2), c(edVarsLRow, 2)).Select
    Application.ScreenUpdating = True

End Sub

Private Function Replaced_var(ByVal origCell As String) As String
    
    Dim varName As String
    Dim varValue As String
    Dim modPostfix As String
    Dim j As Integer, k As Integer, posInArr As Integer
    
    If Mid(origCell, 5, 1) = "@" Then
        Replaced_var = origCell
    Else
        j = InStr(1, origCell, "=", vbTextCompare)  ' find "="
        k = InStrRev(origCell, " ", j - 2, vbTextCompare)   ' find space before varName
        varName = Mid(origCell, k + 1, j - k - 2)
        posInArr = Index_in_array(params, varName)
        If posInArr = 0 Then
            varValue = "***ATTENTION***NOT*FOUND_IN_PARAMS***"
        Else
            varValue = params(posInArr, 2)
        End If
        ' insert var Value from Excel GetStats
        If varName = "defaultInstrument" Or Mid(varName, 1, 9) = "_aux_ins_" Then
            varValue = Replace(varValue, "/", "", 1, 1, vbTextCompare)
            modPostfix = " Instrument." & varValue & ";"
        ElseIf varName = "defaultPeriod" Then
            modPostfix = " Period." & JConverted_Period(varValue) & ";"
        ElseIf varName = "_tag" Then
            modPostfix = " """ & algoTag & """;"
        ElseIf varName = "_algo_comment" Then
            modPostfix = " """ & algoTag & """;"
        ElseIf varName = "_fraction" Then
            varValue = Replace(CStr(myFraction), ",", ".", 1, 1, vbTextCompare)
            modPostfix = " " & varValue & ";"
        ElseIf Mid(origCell, 12, 7) = "boolean" Then
            modPostfix = " " & LCase(CStr(varValue)) & ";"
        Else
            ' insert NUMERIC value of the varName
            varValue = Replace(varValue, ",", ".", 1, 1, vbTextCompare)
            modPostfix = " " & varValue & ";"
        End If
        Replaced_var = Left(origCell, j) & modPostfix
    End If

End Function

Private Function JConverted_Period(ByVal p As String) As String
    
    Dim jcp As String
    
    Select Case p
        Case Is = "4 Hours"
            jcp = "FOUR_HOURS"
        Case Is = "Daily"
            jcp = "DAILY"
        Case Else
            jcp = "*****ATTENTION*****"
    End Select
    JConverted_Period = jcp

End Function

Private Function Index_in_array(ByVal objArr As Variant, _
        ByVal objStr As String) As Integer
    
    Dim pos As Integer
    Dim i As Integer
    Dim foundStr As Boolean
    
    foundStr = False
    For i = LBound(objArr, 1) To UBound(objArr, 1)
        If objArr(i, 1) = objStr Then
            pos = i
            foundStr = True
            Exit For
        End If
    Next i
    If foundStr Then
        Index_in_array = pos
    Else
        Index_in_array = 0
    End If

End Function

Private Function Editor_new_heading(strOrig, insertTag) As String
    
    Dim j As Integer

    j = InStr(14, strOrig, " ", vbTextCompare)
    Editor_new_heading = Left(strOrig, 13) & insertTag & Right(strOrig, Len(strOrig) - j + 1)

End Function

Private Function Strategy_name(ByVal s As String) As String
    
    Dim pfx As String
    
    pfx = Right(s, 4)
    If pfx = "_mxu" Or pfx = "_mux" Or pfx = "_cxu" Or pfx = "_cux" Then
        Strategy_name = Left(s, Len(s) - 4)
    ElseIf InStr(1, s, "_mxu_", vbTextCompare) > 0 Then
        Strategy_name = Replace(s, "_mxu_", "_", 1, 1, vbTextCompare)
    ElseIf InStr(1, s, "_mux_", vbTextCompare) > 0 Then
        Strategy_name = Replace(s, "_mux_", "_", 1, 1, vbTextCompare)
    ElseIf InStr(1, s, "_cxu_", vbTextCompare) > 0 Then
        Strategy_name = Replace(s, "_cxu_", "_", 1, 1, vbTextCompare)
    ElseIf InStr(1, s, "_cux_", vbTextCompare) > 0 Then
        Strategy_name = Replace(s, "_cux_", "_", 1, 1, vbTextCompare)
    Else
        Strategy_name = s
    End If

End Function

Private Function Instrument_abbreviation(ByVal s As String) As String
    
    Dim numer As String
    Dim denom As String
    Dim char1 As String, char2 As String
    
    numer = LCase(Left(s, 3))
    denom = LCase(Right(s, 3))
    Select Case numer
        Case Is = "chf"
            char1 = "f"
            char2 = Left(denom, 1)
        Case Is = "xau"
            char1 = "g"
            char2 = "l"
        Case Is = "xag"
            char1 = "s"
            char2 = "i"
        Case Else
            char1 = Left(numer, 1)
            If denom = "chf" Then
                char2 = "f"
            Else
                char2 = Left(denom, 1)
            End If
    End Select
    Instrument_abbreviation = "_" & char1 & char2

End Function

Sub Settings_To_Launch_Log()

    Dim i As Integer, first_row As Integer, last_row As Integer
    Dim this_col As Integer
    Dim c As Range, Rng As Range, cell As Range
    Dim s As String, stg As String, algo_tag As String
    Dim k As Integer
    
    Application.ScreenUpdating = False
    Set c = ActiveSheet.Cells
    Set Rng = Selection
    first_row = Rng.Rows(1).Row
    this_col = Rng.Columns(1).Column
    last_row = first_row + Rng.Rows.count - 1
    For Each cell In Rng
        s = cell.Value
        If InStr(1, s, "    public ", vbTextCompare) > 0 Then
            If InStr(1, s, " _tag = ", vbTextCompare) > 0 Then
                algo_tag = s
                algo_tag = Replace(algo_tag, "    public String _tag = """, "", 1)
                algo_tag = Left(algo_tag, Len(algo_tag) - 2)
            ElseIf InStr(1, s, " _algo_comment = ", vbTextCompare) > 0 Then
                algo_tag = s
                algo_tag = Replace(algo_tag, "    public String _algo_comment = """, "", 1)
                algo_tag = Left(algo_tag, Len(algo_tag) - 2)
            End If
            s = Replace(s, "    public ", "", 1)
            k = InStr(1, s, " ", vbTextCompare)
            s = Right(s, Len(s) - k)
            s = Replace(s, "Instrument.", "", 1)

            stg = stg & s & " "
        End If
    Next cell
    stg = Left(stg, Len(stg) - 2) & "."
    Rng.Clear
    c(first_row, this_col) = stg
    c(first_row, 2) = algo_tag
    c(first_row + 1, this_col).Select
    Application.ScreenUpdating = True

End Sub


