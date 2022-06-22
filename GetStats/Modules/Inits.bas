Attribute VB_Name = "Inits"
Option Explicit

Const addInFName As String = "GetStats_BackTest_v1.13.xlsm"
Const settingsSheetName As String = "hSettings"
Const backSheetName As String = "Back-test"

Const maxHtmls As Integer = 2000
Const reportType As String = "GS_Pro_Single_Core"
Const depoIniOK As Double = 10000

Const stratFdRow As Integer = 2 ' strategy folder row
Const stratFdCol As Integer = 1 ' strategy folder column
Const stratNmRow As Integer = 7 ' strategy name row
Const stratNmCol As Integer = 1 ' strategy name column

Const instrFRow As Integer = 2
Const instrLRow As Integer = 47
Const instrCol As Integer = 2
Const instrGrpFRow As Integer = 2
Const instrGrpLRow As Integer = 47
Const instrGrpFCol As Integer = 4
Const instrGrpLCol As Integer = 5

Const dateFromRow As Integer = 10
Const dateFromCol As Integer = 2
Const dateToRow As Integer = 11
Const dateToCol As Integer = 2
Const htmlCountRow As Integer = 12
Const htmlCountCol As Integer = 2

Const readyRepFRow As Integer = 10
Const readyRepFCol As Integer = 3
Const readyRepLCol As Integer = 10
Const readyDateCol As Integer = 4
Const readyCountCol As Integer = 5
Const readyDepoIniCol As Integer = 6

'======
Const readyRobotNameCol As Integer = 7
'======

Const readyTimeFromCol As Integer = 8
Const readyTimeToCol As Integer = 9
Const readyLinkCol As Integer = 10

Sub Init_Bt_Settings_Sheets(ByRef wbAddIn As Workbook, _
        ByRef setWs As Worksheet, ByRef btWs As Worksheet, _
        ByRef btC As Range, ByRef activeInstrumentsList As Variant, _
        ByRef instrumentLotGroup As Variant, ByRef stratFdPath As String, _
        ByRef stratNm As String, ByRef dateFrom As Date, _
        ByRef dateTo As Date, ByRef htmlCount As Integer, _
        ByRef dateFromStr As String, ByRef dateToStr As String, _
        ByRef btNextFreeRow As Integer, ByRef maxHtmlCount As Integer, _
        ByRef repType As String, ByRef macroVer As String, _
        ByRef depoIniCheck As Double, ByRef rdRepNameCol As Integer, _
        ByRef rdRepDateCol As Integer, ByRef rdRepCountCol As Integer, _
        ByRef rdRepDepoIniCol As Integer, ByRef rdRepRobotNameCol As Integer, _
        ByRef rdRepTimeFromCol As Integer, ByRef rdRepTimeToCol As Integer, _
        ByRef rdRepLinkCol As Integer)
    
    Dim setC As Range
    Dim instrumentsList As Range
    Dim lastCh As String
    
    Set wbAddIn = Workbooks(addInFName)
    Set btWs = wbAddIn.Sheets(backSheetName)
    Set btC = btWs.Cells
    Set setWs = wbAddIn.Sheets(settingsSheetName)
    Set setC = setWs.Cells
    Set instrumentsList = setWs.Range(setC(instrFRow, instrCol), setC(instrLRow, instrCol))
    activeInstrumentsList = ListActiveInstruments(instrumentsList)
    instrumentLotGroup = GetInstrumentLotGroups(setC, _
            instrGrpFRow, instrGrpLRow, instrGrpFCol, instrGrpLCol)
    stratFdPath = btC(stratFdRow, stratFdCol)
    ' remove "\" at path end
    lastCh = Right(stratFdPath, 1)
    If lastCh = "\" Then
        stratFdPath = Left(stratFdPath, Len(stratFdPath) - 1)
        btC(stratFdRow, stratFdCol) = stratFdPath
    End If
    stratNm = btC(stratNmRow, stratNmCol)
    dateFrom = btC(dateFromRow, dateFromCol)
    dateTo = btC(dateToRow, dateToCol)
    htmlCount = btC(htmlCountRow, htmlCountCol)
    dateFromStr = ConvertDateToString(dateFrom)
    dateToStr = ConvertDateToString(dateTo)
    btNextFreeRow = btC(btWs.Rows.count, readyRepFCol).End(xlUp).Row + 1
    maxHtmlCount = maxHtmls
    repType = reportType
    macroVer = addInFName
    depoIniCheck = depoIniOK
    rdRepNameCol = readyRepFCol
    rdRepDateCol = readyDateCol
    rdRepCountCol = readyCountCol
    rdRepDepoIniCol = readyDepoIniCol
    rdRepRobotNameCol = readyRobotNameCol
    rdRepTimeFromCol = readyTimeFromCol
    rdRepTimeToCol = readyTimeToCol
    rdRepLinkCol = readyLinkCol
    
End Sub

Function GetInstrumentLotGroups(ByRef Rng As Range, _
            ByRef firstRow As Integer, _
            ByRef lastRow As Integer, _
            ByRef firstCol As Integer, _
            ByRef lastCol As Integer) As Variant
    
    Dim a() As Variant
    Dim i As Integer, j As Integer
    Dim ubndRows As Integer
    ubndRows = lastRow - firstRow + 1
    ReDim a(1 To ubndRows, 1 To 2)
    For i = firstRow To lastRow
        j = i - 1
        a(j, 1) = Rng(i, firstCol)
        a(j, 2) = Rng(i, lastCol)
    Next i
    GetInstrumentLotGroups = a
    
End Function

Function ConvertDateToString(ByVal someDate As Date) As String
    
    Dim sY As String, sM As String, sD As String
    
    sY = Right(Year(someDate), 2)
    sM = Format(Month(someDate), "00")
    sD = Format(Day(someDate), "00")
    ConvertDateToString = sY & sM & sD
    
End Function

Function ListActiveInstruments(ByVal instrumentsList As Range) As Variant
    
    Dim arr() As Variant
    Dim cell As Range
    Dim rngSum As Integer, i As Integer
' Args: Instruments True/False list
' Returns: Variant array of active instruments
' if 0 active instruments, redims arr(0 To 0)

    rngSum = 0
    
    For Each cell In instrumentsList
        If cell Then
            rngSum = rngSum + 1
        End If
    Next cell
    
    If rngSum > 0 Then
        ReDim arr(1 To rngSum)
        i = 1
        For Each cell In instrumentsList
            If cell Then
                arr(i) = cell.Offset(0, -1)
                i = i + 1
            End If
        Next cell
    Else
        ReDim arr(0 To 0)
    End If
    
    ListActiveInstruments = arr
    
End Function

Sub Init_Pick_Strategy_Folder(ByRef stratFdRng As Range, _
            ByRef stratNmRng As Range)
    
    Set stratFdRng = Workbooks(addInFName).Sheets(backSheetName).Cells(stratFdRow, stratFdCol)
    Set stratNmRng = Workbooks(addInFName).Sheets(backSheetName).Cells(stratNmRow, stratNmCol)

End Sub

Sub Init_DeSelect_Instruments(ByRef setWs As Worksheet, _
            ByRef btWs As Worksheet, _
            ByRef setC As Range, _
            ByRef btC As Range, _
            ByRef selectAll As Range, _
            ByRef instrumentsList As Range)
    
    Set setWs = Workbooks(addInFName).Sheets(settingsSheetName)
    Set setC = setWs.Cells
    Set btWs = Workbooks(addInFName).Sheets(backSheetName)
    Set btC = btWs.Cells
    Set selectAll = setC(1, 2)
    Set instrumentsList = setWs.Range(setC(2, 2), setC(47, 2))
    
End Sub

Sub Init_Clear_Ready_Reports(ByRef btWs As Worksheet, _
            ByRef btC As Range, _
            ByRef upperRow As Integer, _
            ByRef leftCol As Integer, _
            ByRef rightCol As Integer)
    
    Set btWs = Workbooks(addInFName).Sheets(backSheetName)
    Set btC = btWs.Cells
    upperRow = readyRepFRow
    leftCol = readyRepFCol
    rightCol = readyRepLCol
    
End Sub

Sub Separator_Auto_Switcher(ByRef currentDecimal As String, _
            ByRef undoSep As Boolean, _
            ByRef undoUseSyst As Boolean)
    
    undoSep = False
    undoUseSyst = False
    
    If Application.UseSystemSeparators Then     ' SYS - ON
        Application.UseSystemSeparators = False
        If Not Application.International(xlDecimalSeparator) = "." Then
            currentDecimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            undoSep = True                     ' undo condition 1
            undoUseSyst = True                 ' undo condition 2
        End If
    Else                                        ' SYS - OFF
        If Not Application.DecimalSeparator = "." Then
            currentDecimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            undoSep = True                     ' undo condition 1
            undoUseSyst = False                ' undo condition 2
        End If
    End If
    
End Sub

Sub Separator_Auto_Switcher_Undo(ByRef currentDecimal As String, _
            ByRef undoSep As Boolean, _
            ByRef undoUseSyst As Boolean)
    
    If undoSep Then
        Application.DecimalSeparator = currentDecimal
        If undoUseSyst Then
            Application.UseSystemSeparators = True
        End If
    End If
    
End Sub

Sub InitPositionTags(ByRef positionTags As Dictionary)
    
    positionTags.Add "_tag", Nothing
    positionTags.Add "_algo_comment", Nothing
    
End Sub
