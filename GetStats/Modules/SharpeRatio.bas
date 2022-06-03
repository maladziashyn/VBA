Attribute VB_Name = "SharpeRatio"
Option Explicit
    
Dim new_col As Integer
    
Sub Scatter_Sharpe()
    Const plotWidth As Integer = 10
    Const plotHeight As Integer = 20
    
    Dim colsList() As Variant
    Dim wsResult As Worksheet
    Dim c As Range, rngX As Range, rngY As Range
    Dim colNumber As Integer
    Dim colName As String
    Dim plotTitle As String
    Dim lastRow As Integer
    Dim lastCol As Integer
    Dim i As Integer
    Dim ulr As Integer, ulc As Integer
    Dim chObj As Integer
    Dim newMinMax() As Variant
    Dim newMin As Double, newMax As Double, newStep As Double
    
    Application.ScreenUpdating = False
    Set wsResult = Worksheets(2)
    Set c = wsResult.Cells
    lastRow = c(wsResult.Rows.count, 1).End(xlUp).Row
    If lastRow = 1 Then
        MsgBox "Error. No data."
        Exit Sub
    End If
    lastCol = c(1, 1).End(xlToRight).Column
    ulc = lastCol + 2
    ulr = 2
    chObj = 1
    colsList = SelectedColumnsIDs(Selection)
    c(1, 1).Activate
    Call RemoveScatters
    For i = LBound(colsList) To UBound(colsList)
        colNumber = colsList(i)
        colName = c(1, colNumber)
        Set rngX = Range(Cells(2, colNumber), Cells(lastRow, colNumber))
        Set rngY = Range(Cells(2, lastCol), Cells(lastRow, lastCol))
        plotTitle = colName & " vs. Sharpe"
        
        newMinMax = PlotXMinMax(rngX)
        newMin = newMinMax(1)
        newMax = newMinMax(2)
        newStep = newMinMax(3)
        Call Scatterplot_Sharpe(wsResult, ulr, ulc, plotWidth, plotHeight, _
                                rngX, rngY, plotTitle, chObj, _
                                newMin, newMax, newStep)
        ulr = ulr + plotHeight
        chObj = chObj + 1
    Next i
    c(1, 1).Activate
    Application.ScreenUpdating = True
End Sub
Sub RemoveScatters()
    Dim img As Shape
    For Each img In ActiveSheet.Shapes
        img.Delete
    Next
End Sub
Sub MergeScatters()
    Dim fd As FileDialog

    Dim sel_count As Integer, i As Integer, lr As Integer
    Dim pos As Integer
    Dim tstr As String
    Dim wbA As Workbook, wbB As Workbook
    Dim s As Worksheet
    Dim Rng As Range
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "GetStats: Выбрать отчеты"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Отчеты GetStats", "*.xlsx"
        .ButtonName = "Вперед"
    End With
    If fd.Show = 0 Then
        MsgBox "Файлы не выбраны!"
        Exit Sub
    End If
    sel_count = fd.SelectedItems.count
    
    Set wbA = Workbooks.Add
    Application.ScreenUpdating = False
    If wbA.Sheets.count > 1 Then
        Do Until wbA.Sheets.count = 1
            Application.DisplayAlerts = False
            wbA.Sheets(2).Delete
            Application.DisplayAlerts = True
        Loop
    End If
    For i = 1 To sel_count
        Application.StatusBar = "Добавляю лист " & i & " (" & sel_count & ")."
        Set wbB = Workbooks.Open(fd.SelectedItems(i))
        tstr = wbB.Name
        pos = InStr(1, tstr, "-", 1)
        tstr = Right(Left(tstr, pos + 6), 6)
        If wbB.Sheets(2).Name = "результаты" Then
            wbB.Sheets("результаты").Copy after:=wbA.Sheets(wbA.Sheets.count)
            Set s = wbA.Sheets(wbA.Sheets.count)
            s.Name = i & "_" & tstr
            ' remove hyperlinks
            lr = s.Cells(1, 1).End(xlDown).Row
            Set Rng = s.Range(s.Cells(2, 1), s.Cells(lr, 1))
            Rng.Hyperlinks.Delete
            ' Add hyperlink to original book into cell "A1"
            s.Hyperlinks.Add anchor:=s.Cells(1, 1), Address:=wbB.path & "\" & wbB.Name
        End If
        wbB.Close savechanges:=False
    Next i
    Application.DisplayAlerts = False
    wbA.Sheets(1).Delete
    wbA.Sheets(1).Activate
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Готово. Сохраните файл """ & wbA.Name & """ по вашему усмотрению.", , "GetStats Pro"
End Sub
Sub SharpeScattersToOneBook()
    Dim wbFrom As Workbook, wbTo As Workbook
    Dim shCopy As Worksheet
    Dim newName As String
    
    Application.ScreenUpdating = False
    Set wbFrom = ActiveWorkbook
    Set shCopy = ActiveSheet
    newName = wbFrom.Name
    Set wbTo = Workbooks("mixer.xlsx")
' copy
    shCopy.Copy after:=wbTo.Sheets(wbTo.Sheets.count)
    wbTo.ActiveSheet.Name = newName
'    wb_from.Activate
'    wb_from.Close savechanges:=False
    Application.ScreenUpdating = True
End Sub
Function PlotXMinMax(ByVal Rng As Range) As Variant
    Dim result(1 To 3) As Variant
    Dim rngMin As Double, rngMax As Double, rngStep As Double
    Dim listVals As Object
    Dim cell As Range
    
    rngMin = WorksheetFunction.Min(Rng)
    rngMax = WorksheetFunction.Max(Rng)
    If rngMin <> rngMax Then
        Set listVals = CreateObject("Scripting.Dictionary")
        For Each cell In Rng
            If Not listVals.Exists(cell.Value) Then
                listVals.Add cell.Value, Nothing    ' add key, value to dictionary
            End If
        Next cell
        rngStep = (rngMax - rngMin) / (listVals.count - 1)
        result(1) = rngMin - rngStep
        result(2) = rngMax + rngStep
        result(3) = rngStep
    Else
        result(1) = rngMin
        result(2) = rngMax
        result(3) = 0.1
    End If
    PlotXMinMax = result
End Function
Sub Scatterplot_Sharpe(chsht As Worksheet, _
                        ulr As Integer, _
                        ulc As Integer, _
                        chW As Integer, _
                        chH As Integer, _
                        rngX As Range, _
                        rngY As Range, _
                        ChTitle As String, _
                        chObj As Integer, _
                        xMin As Double, _
                        xMax As Double, _
                        xStep As Double)
    Const chFontSize As Integer = 14    ' chart title font size
    chW = chW * 48      ' chart width, pixels
    chH = chH * 15      ' chart height, pixels
    chsht.Activate                      ' activate sheet
    rngY.Select
    chsht.Shapes.AddChart.Select        ' build chart
    With ActiveChart
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = rngX
        .SeriesCollection(1).Values = rngY
        .ChartType = xlXYScatter
        .Legend.Delete
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow     ' axis position
        ' X-axis values
        If xMin <> xMax Then
            .Axes(xlCategory).MinimumScale = xMin
            .Axes(xlCategory).MaximumScale = xMax
            .Axes(xlCategory).MajorUnit = xStep
        End If
'        ' Y-axis values
'        .Axes(xlValue).MinimumScale = yMin
'        .Axes(xlValue).MaximumScale = yMax
    End With
    With chsht.ChartObjects(chObj)    ' adjust chart placement
        .Left = chsht.Cells(ulr, ulc).Left
        .Top = chsht.Cells(ulr, ulc).Top
        .Width = chW
        .Height = chH
        If xMin <> xMax Then
            .Chart.SeriesCollection(1).Trendlines.Add Type:=xlPolynomial, Order:=2
        End If
'        .Placement = xlFreeFloating     ' do not resize chart if cells resized
    End With
    Cells(ulr, ulc).Activate
End Sub
Sub tmp_ClearAll()
    Cells.Clear
    ActiveWindow.FreezePanes = False
End Sub
Function SelectedColumnsIDs(ByVal userSelection As Range) As Variant
    Dim colsList() As Variant
    Dim i As Integer, j As Integer, colsCount As Integer
    Dim aFirstCol As Integer, aLastCol As Integer, aColCount As Integer
    
    ReDim colsList(0 To 0)
    colsCount = 0
    For i = 1 To userSelection.Areas.count
        aFirstCol = userSelection.Areas.item(i).Column
        aColCount = userSelection.Areas.item(i).Columns.count
        aLastCol = aFirstCol + aColCount - 1
        For j = aFirstCol To aLastCol
            colsCount = colsCount + 1
            If UBound(colsList) = 0 Then
                ReDim colsList(1 To 1)
            Else
                ReDim Preserve colsList(1 To colsCount)
            End If
            colsList(colsCount) = j
        Next j
    Next i
    SelectedColumnsIDs = colsList
End Function
Sub Calc_Sharpe_Ratio()
    Dim annual_std As Double
    Dim last_row As Integer
    Dim current_balance As Double, net_return As Double, cagr As Double
    Dim days_count As Long, i As Long
    Dim Rng As Range
    'Application.ScreenUpdating = False
    If Cells(21, 1) <> "" Then
        Set Rng = Range(Cells(21, 1), Cells(21, 2))
        Rng.Clear
    Else
        last_row = Cells(Rows.count, 13).End(xlUp).Row
        If last_row < 3 Then
            annual_std = 0
        Else
            Set Rng = Range(Cells(2, 13), Cells(last_row, 13))
            annual_std = WorksheetFunction.StDev(Rng) * Sqr(250)
        End If
        Cells(21, 1) = "Sharpe Ratio"

' EXPERIMENT START
' if Joined report, without CAGR > calculate CAGR
        If Cells(4, 2) = "" Then
            ' Calculate Equity Curve (with % returns) to find NetReturn
            current_balance = 1
            For i = 2 To last_row
                current_balance = current_balance * (1 + Cells(i, 13))
            Next i
            
            days_count = Cells(9, 2) - Cells(8, 2) + 1
            net_return = current_balance - 1
            cagr = (1 + net_return) ^ (365 / days_count) - 1
            With Cells(4, 2)
                .Value = cagr
                .NumberFormat = "0.00%"
            End With
        End If
' EXPERIMENT END
        
        With Cells(21, 2)
            If annual_std = 0 Then
                .Value = 0
            Else
                .Value = Cells(4, 2).Value / annual_std
            End If
            .NumberFormat = "0.00"
            .Font.Bold = True
        End With
    End If
    'Application.ScreenUpdating = True
End Sub
Sub Sharpe_to_all()
    Dim i As Integer
    Dim ws As Worksheet
    Dim c As Range
    Application.ScreenUpdating = False
    Set ws = Sheets(2)
    Set c = ws.Cells
    new_col = c(1, 1).End(xlToRight).Column + 1
    c(1, new_col) = "sharpe_ratio"
    For i = 3 To Sheets.count
        Sheets(i).Activate
        Call Calc_Sharpe_Ratio
        With c(i - 1, new_col)
            .Value = Cells(21, 2)
            .NumberFormat = "0.00"
        End With
    Next i
    Sheets(2).Activate
    Rows(1).AutoFilter
    Rows(1).AutoFilter
    Application.ScreenUpdating = True
End Sub
Sub SharpePivot()
    Dim fd As FileDialog
    Dim wb As Workbook
    Dim tWb As Workbook
    Dim ws As Worksheet
    Dim tWs As Worksheet
    Dim tC As Range
    Dim i As Integer
    Dim rg As Range
    Dim insertRow As Long
    
    Application.ScreenUpdating = False
    
    Set tWb = Workbooks.Add
    Set tWs = tWb.Sheets(1)
    Set tC = tWs.Cells
    insertRow = 1
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "GetStats: Pick files"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "GetStats joint reports", "*.xlsx"
        .ButtonName = "Okey dokey"
    End With
    If fd.Show = 0 Then
        MsgBox "No files selected."
        Exit Sub
    End If
    
    For i = 1 To fd.SelectedItems.count
        Set wb = Workbooks.Open(fd.SelectedItems(i))
        wb.Sheets(2).Activate
        Call Params_To_Summary
        Call Sharpe_to_all
        
        Set rg = ActiveCell.CurrentRegion
        rg.Copy tC(insertRow, 1)
        
        insertRow = tC(tC(tWs.Rows.count, 1).End(xlUp).Row, 1).Row + 2
    
        wb.Close savechanges:=False
    Next i
    
    Application.ScreenUpdating = True
End Sub
