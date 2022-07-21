Attribute VB_Name = "Tools"
Option Explicit
Sub Click_Locate_Target()
    Application.ScreenUpdating = False
    Call Click_LocateTarget_AddSource(False)
    Application.ScreenUpdating = True
End Sub
Sub Click_Add_Source()
    Application.ScreenUpdating = False
    Call Click_LocateTarget_AddSource(True)
    Application.ScreenUpdating = True
End Sub
Sub Click_Clear_Sources()
    Dim ws As Worksheet
    Dim clrRng As Range
    
    Application.ScreenUpdating = False
    Call Click_Clear_Sources_Inits(ws, clrRng)
    clrRng.Clear
    Application.ScreenUpdating = True
End Sub
Sub Click_LocateTarget_AddSource(ByRef addSource As Boolean)
' Sub pastes directory path to 'A6'
' callable by other subs above, ignore screen updating
    Dim fd As FileDialog
    Dim dirPath As String
    Dim ws As Worksheet
    Dim fillC As Range
    Dim fillRow As Integer
    Dim fillCol As Integer
    Dim dialTitle As String
    Dim okBtnName As String
    
    Call Click_Locate_Target_Inits(ws, fillC, fillRow, fillCol, _
            dialTitle, okBtnName, addSource)
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = dialTitle
        .AllowMultiSelect = True
        .ButtonName = okBtnName
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    dirPath = CStr(fd.SelectedItems(1))
    fillC(fillRow, fillCol) = dirPath
    ws.columns(fillCol).AutoFit
End Sub
Sub Click_Copy_Dates_From_Selection()
    Dim mainWs As Worksheet
    Dim mainC As Range
    Dim date0Row As Integer
    Dim date9Row As Integer
    Dim datesCol As Integer
    Dim dirsCol As Integer
    Dim srcDirsZeroRow As Integer
    Dim date0 As Date
    Dim date9 As Date
    Dim selectedFilePath As String
    Dim checkCellValue As String
    
    Call Click_Copy_Dates_From_Selection_Inits(mainWs, mainC, _
            date0Row, date9Row, datesCol, dirsCol, srcDirsZeroRow)
' sanity check
' not empty, row after "Source directories", right column
    If Selection.Row > srcDirsZeroRow _
        And Selection.Column = dirsCol _
        And mainC(Selection.Row, Selection.Column) <> "" Then
' highlight
        mainC(date0Row, datesCol).Interior.Color = RGB(255, 0, 0)
        mainC(date9Row, datesCol).Interior.Color = RGB(255, 0, 0)
    
        Application.ScreenUpdating = False
        selectedFilePath = FirstFileFromSelectedDir(mainC, Selection, dirsCol)
        Call ExtractTestDates(selectedFilePath, date0, date9)
        mainC(date0Row, datesCol) = date0
        mainC(date9Row, datesCol) = date9
' remove highlight
        mainC(date0Row, datesCol).Interior.Pattern = xlNone
        mainC(date9Row, datesCol).Interior.Pattern = xlNone
    Else
        MsgBox "Select one of Source Directories.", vbCritical, "Error"
    End If
    Application.ScreenUpdating = True
End Sub
Sub DeSelect_KPIs()
    Dim stgWs As Worksheet
    Dim stgC As Range
    Dim cell As Range
    Dim checkAll As Range
    Dim kpisList As Range
    Application.ScreenUpdating = False
    Call DeSelect_KPIs_Inits(stgWs, stgC, checkAll, kpisList)
    If checkAll Then
        For Each cell In kpisList
            cell = True
        Next cell
    Else
        For Each cell In kpisList
            cell = False
        Next cell
    End If
    Application.ScreenUpdating = True
End Sub
Sub ExtractTestDates(ByRef filePath As String, _
            ByRef date0 As Date, _
            ByRef date9 As Date)
    Dim wbDates As Workbook
    Set wbDates = Workbooks.Open(filePath)
    date0 = wbDates.Sheets(3).Cells(8, 2)
    date9 = wbDates.Sheets(3).Cells(9, 2)
    wbDates.Close savechanges:=False
End Sub

Function FirstFileFromSelectedDir(ByVal mainCells As Range, _
            ByVal userSelection As Range, _
            ByVal srcCol As Integer) As String
    Dim srcRow As Integer
    Dim dirPath As String
    Dim filePath As String
    srcRow = userSelection.Areas.item(1).Row
    dirPath = mainCells(srcRow, srcCol).Value
    filePath = DirectoryFilesList(dirPath, True, False)(1)
    If Right(filePath, 4) = "xlsx" Then
        FirstFileFromSelectedDir = dirPath & "\" & filePath
    Else
        FirstFileFromSelectedDir = ""
    End If
End Function
Sub Print_1D_Array(ByVal print_arr As Variant, _
            ByVal col_offset As Integer, _
            ByVal print_cells As Range)
' Procedure prints any 1-dimensional array in a new Workbook, sheet 1.
' Arguments:
'       1) 1-D array
    
'    Dim wb_print As Workbook
    Dim r As Integer
    Dim print_row As Integer, print_col As Integer
    Dim add_rows As Integer

    If LBound(print_arr) = 0 Then
        add_rows = 1
    Else
        add_rows = 0
    End If
'    Set wb_print = Workbooks.Add
'    Set c_print = wb_print.Sheets(1).cells
    print_col = 1 + col_offset
    For r = LBound(print_arr) To UBound(print_arr)
        print_row = r + add_rows
        print_cells(print_row, print_col) = print_arr(r)
    Next r
End Sub
Sub Print_2D_Array(ByVal print_arr As Variant, _
            ByVal is_inverted As Boolean, _
            ByVal row_offset As Integer, _
            ByVal col_offset As Integer, _
            ByVal print_cells As Range)
' Procedure prints any 2-dimensional array in a new Workbook, sheet 1.
' Arguments:
'       1) 2-D array
'       2) rows-colums (is_inverted = False) or columns-rows (is_inverted = True)
    
'    Dim wb_print As Workbook
'    Dim c_print As Range
    Dim r As Integer, C As Integer
    Dim print_row As Integer, print_col As Integer
    Dim row_dim As Integer, col_dim As Integer
    Dim add_rows As Integer, add_cols As Integer

    If is_inverted Then
        row_dim = 2
        col_dim = 1
    Else
        row_dim = 1
        col_dim = 2
    End If
    If LBound(print_arr, row_dim) = 0 Then
        add_rows = 1
    Else
        add_rows = 0
    End If
    If LBound(print_arr, col_dim) = 0 Then
        add_cols = 1
    Else
        add_cols = 0
    End If
'    Set wb_print = Workbooks.Add
'    Set c_print = wb_print.Sheets(1).cells
    For r = LBound(print_arr, row_dim) To UBound(print_arr, row_dim)
        print_row = r + add_rows + row_offset
        For C = LBound(print_arr, col_dim) To UBound(print_arr, col_dim)
            print_col = C + add_cols + col_offset
            If is_inverted Then
                print_cells(print_row, print_col) = print_arr(C, r)
            Else
                print_cells(print_row, print_col) = print_arr(r, C)
            End If
        Next C
    Next r
End Sub
Sub GenerateIsOsCodes()
    Dim i As Integer
    Dim rg As Range
    Dim exitError As Boolean
    Dim addInName As String
    exitError = False
    Call GenerateIsOsCodes_Inits(rg, exitError, addInName)
    If exitError = True Then
        MsgBox "Please, provide In-Sample & Out-of-Sample windows.", vbCritical, addInName
        Exit Sub
    End If
    For i = 1 To rg.rows.Count
        rg(i, 3) = "i" & Int(rg(i, 1) / 52) & "o" & Int(rg(i, 2) / 52)
    Next i
End Sub
Sub WFAChartClassic(ByRef sc As Range, _
                ByVal ulr As Integer, _
                ByVal ulc As Integer, _
                ByVal logScale As Boolean, _
                ByVal finRes As Double)
' Build chart for IS/OS/Forward date slots.
'
' ** CONSTANTS **
    Const ch_hght_cells As Integer = 20
    Const ch_wdth_cells As Integer = 5
    Const my_rnd = 0.1
' ** VARIABLES **
' Double
    Dim maxVal As Double
    Dim MinVal As Double
' Integer
    Dim chFontSize As Integer
    Dim chObj_idx As Integer
    Dim last_date_row As Integer
' Range
    Dim rng_to_cover As Range
    Dim rngX As Range
    Dim rngY As Range
' String
    Dim ChTitle As String
    
    chObj_idx = ActiveSheet.ChartObjects.Count + 1
    ChTitle = sc(1, ulc) & ", FinRes = " & finRes
    If Left(sc(1, ulc), 2) = "IS" And logScale Then
        ChTitle = ChTitle & ", log scale"         ' log scale
    End If
    last_date_row = sc(2, ulc + 4).End(xlDown).Row
    chFontSize = 12
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + ch_hght_cells, ulc + ch_wdth_cells))
    Set rngX = Range(sc(2, ulc + 4), sc(last_date_row, ulc + 4))
    Set rngY = Range(sc(2, ulc + 5), sc(last_date_row, ulc + 5))
    MinVal = WorksheetFunction.Min(rngY)
    maxVal = WorksheetFunction.Max(rngY)
    MinVal = my_rnd * Int(MinVal / my_rnd)
    maxVal = my_rnd * Int(maxVal / my_rnd) + my_rnd
    rngY.Select
    ActiveSheet.Shapes.AddChart.Select
    With ActiveSheet.ChartObjects(chObj_idx)
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = rng_to_cover.Width
        .Height = rng_to_cover.Height
'        .Placement = xlFreeFloating ' do not resize chart if cells resized
    End With
    With ActiveChart
        .SetSourceData Source:=Application.Union(rngX, rngY)
        .ChartType = xlLine
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = maxVal
        If Left(sc(1, ulc), 2) = "IS" And logScale Then
            .Axes(xlValue).ScaleType = xlScaleLogarithmic   ' log scale
        End If
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
    sc(1, ulc + 1) = chObj_idx
    sc(1, ulc).Select
End Sub
Sub ChartRangesXandY(ByVal rngX As Range, _
            ByVal rngY As Range, _
            ByRef sc As Range, _
            ByVal ulr As Integer, _
            ByVal ulc As Integer, _
            ByVal logScale As Boolean, _
            ByVal selRow As Integer)
' Build chart for IS/OS/Forward date slots.
'
' ** CONSTANTS **
    Const ch_hght_cells As Integer = 19
    Const ch_wdth_cells As Integer = 9
    Const my_rnd = 0.1
' ** VARIABLES **
' Double
    Dim maxVal As Double
    Dim MinVal As Double
' Integer
    Dim chFontSize As Integer
    Dim chObj_idx As Integer
' Range
    Dim rng_to_cover As Range
'    Dim rngX As Range
'    Dim rngY As Range
' String
    
    chObj_idx = ActiveSheet.ChartObjects.Count + 1
    chFontSize = 12
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + ch_hght_cells, ulc + ch_wdth_cells))
    MinVal = WorksheetFunction.Min(rngY)
    maxVal = WorksheetFunction.Max(rngY)
    MinVal = my_rnd * Int(MinVal / my_rnd)
    maxVal = my_rnd * Int(maxVal / my_rnd) + my_rnd
    rngY.Select
    ActiveSheet.Shapes.AddChart.Select
    With ActiveSheet.ChartObjects(chObj_idx)
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = rng_to_cover.Width
        .Height = rng_to_cover.Height
'        .Placement = xlFreeFloating ' do not resize chart if cells resized
    End With
    With ActiveChart
        .SetSourceData Source:=Application.Union(rngX, rngY)
        .ChartType = xlLine
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = maxVal
        If logScale Then
            .Axes(xlValue).ScaleType = xlScaleLogarithmic   ' log scale
        End If
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
        .chartTitle.Delete
'        .ChartTitle.Text = ChTitle
'        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
    sc(selRow, ulc).Select
End Sub
Sub StatementChartRangesXandY(ByVal rngX As Range, _
            ByVal rngY As Range, _
            ByRef sc As Range, _
            ByVal ulr As Integer, _
            ByVal ulc As Integer, _
            ByVal logScale As Boolean, _
            ByVal selRow As Integer, _
            ByVal chartTitle As String)
' Build chart for IS/OS/Forward date slots.
'
' ** CONSTANTS **
    Const ch_hght_cells As Integer = 19
    Const ch_wdth_cells As Integer = 9
    Const my_rnd = 0.02
' ** VARIABLES **
' Double
    Dim maxVal As Double
    Dim MinVal As Double
' Integer
    Dim chFontSize As Integer
    Dim chObj_idx As Integer
' Range
    Dim rng_to_cover As Range
    
    chObj_idx = ActiveSheet.ChartObjects.Count + 1
    chFontSize = 12
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + ch_hght_cells, ulc + ch_wdth_cells))
    MinVal = WorksheetFunction.Min(rngY)
    maxVal = WorksheetFunction.Max(rngY)
    MinVal = my_rnd * Int(MinVal / my_rnd)
    maxVal = my_rnd * Int(maxVal / my_rnd) + my_rnd
    rngY.Select
    ActiveSheet.Shapes.AddChart.Select
    With ActiveSheet.ChartObjects(chObj_idx)
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = rng_to_cover.Width
        .Height = rng_to_cover.Height
'        .Placement = xlFreeFloating ' do not resize chart if cells resized
    End With
    With ActiveChart
        .SetSourceData Source:=Application.Union(rngX, rngY)
        .ChartType = xlLine
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = maxVal
        If logScale Then
            .Axes(xlValue).ScaleType = xlScaleLogarithmic   ' log scale
        End If
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
'        .chartTitle.Delete
        .chartTitle.Text = chartTitle
'        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
    sc(selRow, ulc - 2).Select
End Sub

Sub WfaWinnersRemoveDuplicates()
' WFA single report, any date slot: IS-n, OS-n, Forward Compiled.
' Take source column, copy below, remove duplicats, sort ascending.
' Integer
    Dim firstCol As Integer
    Dim srcCol As Integer
    Dim thisCol As Integer
' Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim lastRowAlt As Long
' Range
    Dim rg As Range
    
    If IsEmpty(ActiveCell) Then  ' Sanity check
        Exit Sub
    End If
    Application.ScreenUpdating = False
    thisCol = ActiveCell.Column
    If Not IsEmpty(ActiveCell.Offset(0, -1)) Then
        firstCol = Cells(2, thisCol).End(xlToLeft).Column
    Else
        firstCol = thisCol
    End If
    lastRow = Cells(ActiveSheet.rows.Count, firstCol).End(xlUp).Row
    If lastRow < 3 Then  ' Sanity check
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Set rg = ActiveCell.CurrentRegion
    If rg.columns.Count = 1 Then
        ' Remove existing copy #1
        rg.Clear
'        Cells(1, thisCol - 2).Select
    Else
        ' Remove existing copy #2
        srcCol = firstCol + 2
        lastRowAlt = Cells(ActiveSheet.rows.Count, srcCol).End(xlUp).Row
        If lastRow <> lastRowAlt Then
            Set rg = Cells(lastRowAlt, thisCol).CurrentRegion
            rg.Clear
            Application.ScreenUpdating = True
            Exit Sub
        End If
        ' Copy, remove duplicates, sort ascending
        Set rg = Range(Cells(3, srcCol), Cells(lastRow, srcCol))
        rg.Copy Cells(lastRow + 3, srcCol)
        firstRow = lastRow + 3
        lastRow = Cells(ActiveSheet.rows.Count, srcCol).End(xlUp).Row
        Set rg = Range(Cells(firstRow, srcCol), Cells(lastRow, srcCol))
        rg.RemoveDuplicates columns:=1, Header:=xlNo
        Set rg = Cells(firstRow, srcCol).CurrentRegion
        rg.Sort Key1:=rg.Cells(1), Order1:=xlAscending, Header:=xlNo
        rg.Cells(0, 1).Activate
    End If
    Application.ScreenUpdating = True
End Sub
Sub OpenWfaSource()
' Open workbook and activate sheet with WFA source.
' FileDialog
    Dim fd As FileDialog
' Integer
    Dim uscorePos As Integer
' Range
    Dim rg As Range
' String
    Dim dirPath As String
    Dim filePath As String
    Dim sheetName As String
' Workbook
    Dim wb As Workbook
    
    If IsEmpty(ActiveCell) Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Set rg = ActiveCell.CurrentRegion
    dirPath = rg.Cells(1, 1)
    If Not LooksLikeDirectory(dirPath) Then
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        With fd
            .Title = "Pick source folder"
            .AllowMultiSelect = False
            .ButtonName = "Okey Dokey"
        End With
        If fd.Show = 0 Then
            Exit Sub
            Application.ScreenUpdating = True
        End If
        dirPath = CStr(fd.SelectedItems(1))
        rg.Cells(0, 1) = dirPath
    End If
    uscorePos = InStrRev(ActiveCell, "_", -1, vbTextCompare)
    filePath = dirPath & "\" & Left(ActiveCell.Value, uscorePos - 1) & ".xlsx"
    sheetName = Right(ActiveCell, Len(ActiveCell) - uscorePos)
    Set wb = Workbooks.Open(filePath)
    wb.Sheets(sheetName).Activate
    Application.ScreenUpdating = True
End Sub
Function LooksLikeDirectory(ByRef dirPath As String) As Boolean
' Find out if a suggested path is a directory and it exists.
' Boolean
    Dim result As Boolean
    result = False
    If InStr(1, dirPath, "\", vbTextCompare) > 0 Then
        dirPath = StringRemoveBackslash(dirPath)
        result = True
    End If
    LooksLikeDirectory = result
End Function
Sub ManuallyApplyDateFilter()
    Dim i As Integer
    Dim kpis As Dictionary
    Dim calDays As Variant
    Dim calDaysLong As Variant
    Dim tradeList As Variant
    Dim date0 As Variant
    Dim date9 As Variant
    Dim sRow As Integer
    Dim sCol As Integer
    Dim rg As Range
    Dim printRow As Integer
    Dim printCol As Integer
    Dim daysAndEquity As Variant
    Dim someShape As Shape
    Dim rangeX As Range
    Dim rangeY As Range
    Const rowShift As Integer = 5
    Const shiftPrintRow As Integer = 2
    Application.ScreenUpdating = False
    sRow = Selection.Row
    sCol = Selection.Column
    date0 = Cells(sRow, sCol)
    date9 = Cells(sRow, sCol + 1)
    tradeList = GetTradeListFromSheet(ActiveSheet, date0, date9, ActiveWorkbook.Name)
    If Not IsEmpty(Cells(sRow + rowShift + 1, sCol)) Then
        For Each someShape In ActiveSheet.Shapes
            someShape.Delete
        Next someShape
        Set rg = Cells(sRow + rowShift + 1, sCol).CurrentRegion
        rg.Clear
        Set rg = Cells(sRow + shiftPrintRow + 1, sCol).CurrentRegion
        rg.Clear
    End If
    Call Print_2D_Array(tradeList, True, sRow + rowShift, sCol - 1, Cells)
    ' calculate KPIs
    calDays = GenerateCalendarDays(date0, date9)
    calDaysLong = GenerateLongDays(UBound(calDays))
    Set kpis = CalcKPIs(tradeList, date0, date9, calDays, calDaysLong)
    Call PrintDictionary(kpis, True, sRow + shiftPrintRow, sCol - 1, Cells)
    daysAndEquity = GetDailyEquityFromTradeSet(tradeList, date0, date9)
    Call Print_2D_Array(daysAndEquity, False, sRow + rowShift, sCol + 3, Cells)
    Set rangeX = Cells(sRow + rowShift + 1, sCol).CurrentRegion
    Set rangeX = rangeX.Offset(0, 4).Resize(rangeX.rows.Count, rangeX.columns.Count - 5)
    Set rangeY = Cells(sRow + rowShift + 1, sCol).CurrentRegion
    Set rangeY = rangeY.Offset(0, 5).Resize(rangeY.rows.Count, rangeY.columns.Count - 5)
    Call ChartRangesXandY(rangeX, rangeY, Cells, sRow + rowShift + 2, sCol, False, sRow)
    Application.ScreenUpdating = True
End Sub
Sub PrintDictionary(ByVal pDict As Dictionary, _
            ByVal arrHorizontal As Boolean, _
            ByVal rowOffset As Integer, _
            ByVal colOffset As Integer, _
            ByVal pCells As Range)
    Dim prArr As Variant
    prArr = DictionaryToArray(pDict, arrHorizontal)
    Call Print_2D_Array(prArr, False, rowOffset, colOffset, pCells)
End Sub
Sub CreateNewWorkbookSheetsCountNames(ByRef targetWb As Workbook, _
            ByVal newSheetsCount As Integer)
    Dim i As Integer
    Call NewWorkbookSheetsCount(targetWb, newSheetsCount)
    targetWb.Sheets(1).Name = "Summary"
    For i = 2 To targetWb.Sheets.Count
        targetWb.Sheets(i).Name = CStr(i - 1)
    Next i
End Sub
Sub NewWorkbookSheetsCount(ByRef theWb As Workbook, _
            ByVal newShCount As Integer)
    Dim oldSheetsCount As Integer
    oldSheetsCount = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = newShCount
    Set theWb = Workbooks.Add
    Application.SheetsInNewWorkbook = oldSheetsCount
End Sub
Sub StatementChartRangesXYZ(ByVal rngX As Range, _
            ByVal rngY As Range, _
            ByVal rngZ As Range, _
            ByRef sc As Range, _
            ByVal ulr As Integer, _
            ByVal ulc As Integer)
' Build chart for IS/OS/Forward date slots.
'
' ** CONSTANTS **
    Const ch_hght_cells As Integer = 24
    Const ch_wdth_cells As Integer = 13
    Const my_rnd = 0.1
    Const my_rnd2 = 0.05
' ** VARIABLES **
' Double
    Dim minValY As Double
    Dim maxValY As Double
    Dim minValZ As Double
    Dim maxValZ As Double
' Integer
    Const chFontSize As Integer = 12
    Dim chObj_idx As Integer
' Range
    Dim rng_to_cover As Range
' String
    
    chObj_idx = ActiveSheet.ChartObjects.Count + 1
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + ch_hght_cells, ulc + ch_wdth_cells))
    
    minValY = WorksheetFunction.Min(rngY)
    maxValY = WorksheetFunction.Max(rngY)
    minValY = my_rnd * Int(minValY / my_rnd2)
    maxValY = my_rnd * Int(maxValY / my_rnd2) + my_rnd2
    
    minValZ = WorksheetFunction.Min(rngZ)
    maxValZ = WorksheetFunction.Max(rngZ)
    minValZ = my_rnd * Int(minValZ / my_rnd)
    maxValZ = my_rnd * Int(maxValZ / my_rnd) + my_rnd
    rngZ.Select
    ActiveSheet.Shapes.AddChart.Select
    With ActiveSheet.ChartObjects(chObj_idx)
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = rng_to_cover.Width
        .Height = rng_to_cover.Height
'        .Placement = xlFreeFloating ' do not resize chart if cells resized
    End With
    With ActiveChart
        .SetSourceData Source:=Application.Union(rngX, rngY, rngZ)
        .ChartType = xlLine
'        .Legend.Delete
        .Axes(xlValue).MinimumScale = minValZ
        .Axes(xlValue).MaximumScale = maxValZ
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
'        .ChartTitle.Delete
        .chartTitle.Text = "Cumulative Returns"
        .chartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
        .SeriesCollection(1).AxisGroup = 2
        .SeriesCollection(1).ChartType = xlColumnClustered
        .Axes(xlValue).TickLabels.NumberFormat = "0.0%"
        .Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0.0%"
        .SeriesCollection(1).Name = "DayReturn"
        .SeriesCollection(2).Name = "ReturnCurve"
        .Axes(xlValue, xlSecondary).MinimumScale = minValY
        .Axes(xlValue, xlSecondary).MaximumScale = maxValY
    End With
    sc(1, 1).Select
End Sub
Sub SortSheetsAlphabetically()
' Sort sheets in alphabetical order
    Dim i As Integer
    Dim areSorted As Boolean
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim sortPoints As Integer
    Dim iterCount As Long
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    areSorted = False
    iterCount = 0
    Do While areSorted = False
        Application.StatusBar = iterCount
        sortPoints = 0
        For i = 1 To Sheets.Count - 1
            Set ws1 = Sheets(i)
            Set ws2 = Sheets(i + 1)
            If LCase(ws2.Name) > LCase(ws1.Name) Then
                sortPoints = sortPoints + 1
            End If
        Next i
        If sortPoints = Sheets.Count - 1 Then
            areSorted = True
            Exit Do
        Else
            For i = 1 To Sheets.Count - 1
                Set ws1 = Sheets(i)
                Set ws2 = Sheets(i + 1)
                If LCase(ws1.Name) > LCase(ws2.Name) Then
                    Sheets(ws2.Index).Move before:=Sheets(ws1.Index)
                End If
            Next i
        End If
        iterCount = iterCount + 1
    Loop
    Sheets(1).Activate
    
    With Application
        .StatusBar = False
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    MsgBox "Done. Iterations: " & iterCount
End Sub
