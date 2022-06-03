Option Explicit

Sub Chart_Classic(ulr As Integer, ulc As Integer, RngX As Range, RngY As Range, ChTitle As String)
' standard macro for any chart
' for all times and peoples
    
    Dim chW As Integer, chH As Integer          ' chart width, chart height
    Dim chFontSize As Integer                   ' chart title font size

    Application.ScreenUpdating = False
    chW = 480   ' standrad cell width = 48 pix  '480
    chH = 300   ' standard cell height = 15 pix '300
    chFontSize = 14
' build chart
    ActiveSheet.Shapes.AddChart.Select
' adjust chart placement
    With ActiveSheet.ChartObjects(1)
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = chW
        .Height = chH
' do not resize chart if cells resized
        .Placement = xlFreeFloating
    End With
    With ActiveChart
        .SetSourceData Source:=Application.Union(RngX, RngY)
' chart type - line
        .ChartType = xlLine
' delete legend
        .Legend.Delete
' chart title position
        .SetElement (msoElementChartTitleAboveChart)
' chart title
        .ChartTitle.Text = ChTitle
' title font size
        .ChartTitle.Characters.Font.Size = chFontSize
' axis position
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
'
    Cells(ulr, ulc).Activate
    Application.ScreenUpdating = True

End Sub

Sub Chart_Classic_wMinMax(ulr As Integer, ulc As Integer, RngX As Range, RngY As Range, ChTitle As String, MinVal As Long, MaxVal As Long)
' standard macro for any chart
' for all times and peoples
    
    Dim chW As Integer, chH As Integer          ' chart width, chart height
    Dim chFontSize As Integer                   ' chart title font size

    Application.ScreenUpdating = False
    chW = 480   ' standrad cell width = 48 pix
    chH = 300   ' standard cell height = 15 pix
    chFontSize = 14
' build chart
    ActiveSheet.Shapes.AddChart.Select
' adjust chart placement
    With ActiveSheet.ChartObjects(1)
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = chW
        .Height = chH
' do not resize chart if cells resized
        .Placement = xlFreeFloating
    End With
    With ActiveChart
        .SetSourceData Source:=Application.Union(RngX, RngY)
' chart type - line
        .ChartType = xlLine
' delete legend
        .Legend.Delete
' minimum and maximum Y axis values
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = MaxVal
' chart title position
        .SetElement (msoElementChartTitleAboveChart)
' chart title
        .ChartTitle.Text = ChTitle
' title font size
        .ChartTitle.Characters.Font.Size = chFontSize
' axis position
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
'
    Cells(ulr, ulc).Activate
    Application.ScreenUpdating = True
    
End Sub

Sub Charts_OHLC_wMinMax(chsht As Worksheet, _
                        ulr As Integer, _
                        ulc As Integer, _
                        chW As Integer, _
                        chH As Integer, _
                        RngX As Range, _
                        RngY As Range, _
                        ChTitle As String, _
                        MinVal As Long, _
                        MaxVal As Long, _
                        chobj_n As Integer)
    
    Const chFontSize As Integer = 14    ' chart title font size
    chW = chW * 48      ' chart width, pixels
    chH = chH * 15      ' chart height, pixels
    chsht.Activate                      ' activate sheet
    RngY.Select
    chsht.Shapes.AddChart.Select        ' build chart
    With ActiveChart
        .SetSourceData Source:=RngY
        .ChartType = xlStockOHLC                        ' chart type - OHLC
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = MaxVal
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow     ' axis position
        .SeriesCollection(1).XValues = RngX             ' lower axis data
    End With
    With chsht.ChartObjects(chobj_n)    ' adjust chart placement
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = chW
        .Height = chH
        .Placement = xlFreeFloating     ' do not resize chart if cells resized
    End With
    Cells(ulr, ulc).Activate

End Sub

Sub Charts_Line_Y_wMinMax(chsht As Worksheet, _
                        ulr As Integer, _
                        ulc As Integer, _
                        chW As Integer, _
                        chH As Integer, _
                        RngY As Range, _
                        ChTitle As String, _
                        MinVal As Long, _
                        MaxVal As Long, _
                        chobj_n As Integer)
    
    Const chFontSize As Integer = 14    ' chart title font size
    chW = chW * 48      ' chart width, pixels
    chH = chH * 15      ' chart height, pixels
    chsht.Activate                      ' activate sheet
    RngY.Select
    chsht.Shapes.AddChart.Select        ' build chart
    
    With ActiveChart
        .SetSourceData Source:=RngY
        .ChartType = xlLine                        ' chart type - OHLC
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = MaxVal
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow     ' axis position
'        .SeriesCollection(1).XValues = RngX             ' lower axis data
    End With
    
    With chsht.ChartObjects(chobj_n)    ' adjust chart placement
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = chW
        .Height = chH
        .Placement = xlFreeFloating     ' do not resize chart if cells resized
    End With
    Cells(ulr, ulc).Activate

End Sub

Sub Charts_Line_XY_wMinMax(chsht As Worksheet, _
                        ulr As Integer, _
                        ulc As Integer, _
                        chW As Integer, _
                        chH As Integer, _
                        RngX As Range, _
                        RngY As Range, _
                        ChTitle As String, _
                        MinVal As Long, _
                        MaxVal As Long, _
                        chobj_n As Integer)

'Charts_Line_XY_wMinMax(chsht, ulr, ulc, chW, chH, RngX, RngY, ChTitle, MinVal, MaxVal, chobj_n)
    Const chFontSize As Integer = 14    ' chart title font size
    chW = chW * 48      ' chart width, pixels
    chH = chH * 15      ' chart height, pixels
    chsht.Activate                      ' activate sheet
    RngY.Select
    chsht.Shapes.AddChart.Select        ' build chart
    
    With ActiveChart
        .SetSourceData Source:=RngY
        .ChartType = xlLine                        ' chart type - OHLC
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = MaxVal
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow     ' axis position
        .SeriesCollection(1).XValues = RngX             ' lower axis data
    End With
    With chsht.ChartObjects(chobj_n)    ' adjust chart placement
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = chW
        .Height = chH
        .Placement = xlFreeFloating     ' do not resize chart if cells resized
    End With
    Cells(ulr, ulc).Activate

End Sub

Sub Charts_Hist_Y(chsht As Worksheet, _
                            ulr As Integer, _
                            ulc As Integer, _
                            chW As Integer, _
                            chH As Integer, _
                            RngY As Range, _
                            ChTitle As String, _
                            chobj_n As Integer)
    
    Const chFontSize As Integer = 14    ' chart title font size
    chW = chW * 48      ' chart width, pixels
    chH = chH * 15      ' chart height, pixels
    chsht.Activate                      ' activate sheet
    RngY.Select
    chsht.Shapes.AddChart.Select        ' build chart
    
    With ActiveChart
        .SetSourceData Source:=RngY
        .ChartType = xlColumnClustered                  ' chart type - histogram
        .Legend.Delete
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow     ' axis position
    End With
    With chsht.ChartObjects(chobj_n)    ' adjust chart placement
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = chW
        .Height = chH
        .Placement = xlFreeFloating     ' do not resize chart if cells resized
    End With
    Cells(ulr, ulc).Activate
End Sub


Sub Charts_Remove_From_All_Sheets()
    Dim x As Integer
    Dim img As Shape
    Application.ScreenUpdating = False
    For x = 1 To Sheets.Count
        Application.StatusBar = "x = " & x & " (" & Sheets.Count & ")"
        For Each img In Sheets(x).Shapes
            img.Delete
        Next
    Next x
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "done"
End Sub
Sub Charts_Add_To_All_Sheets()
    Dim x As Integer
    Application.ScreenUpdating = False
    For x = 2 To Sheets.Count
        Application.StatusBar = "x = " & x & " (" & Sheets.Count & ")"
        Sheets(x).Activate
        Call CreateCharts
    Next x
    Sheets(1).Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "done"
End Sub


Sub Conditional_Formatting_Add_DataBar(ByRef RngDataBar As Range, ByRef cRed As Integer, ByRef cGreen As Integer, ByRef cBlue As Integer)

    RngDataBar.FormatConditions.AddDatabar
    RngDataBar.FormatConditions(RngDataBar.FormatConditions.Count).ShowValue = True
    With RngDataBar.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueLowestValue
        .MaxPoint.Modify newtype:=xlConditionValueHighestValue
'        .BarColor.Color = RGB(99, 142, 198)
        .BarColor.Color = RGB(cRed, cGreen, cBlue)
    End With
End Sub

Sub Histogram_Classic(ulr As Integer, ulc As Integer, _
                RngX As Range, RngY As Range, _
                ChTitle As String, Hsht As Worksheet, _
                chobj As Integer)
    Dim chW As Integer, chH As Integer          ' chart width, chart height
    Dim chFontSize As Integer                   ' chart title font size

    ' 10 x 10 - width x height
    chW = 480   ' standrad cell width = 48 pix  '480
    chH = 300   ' standard cell height = 15 pix '300
    chFontSize = 14
' build chart
    Hsht.Activate
    RngY.Select
    Hsht.Shapes.AddChart.Select
    With ActiveChart
        .SetSourceData Source:=RngY     ' source
        .ChartType = xlColumnClustered  ' chart type - histogram
        .Legend.Delete                  ' delete legend
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .ChartTitle.Text = ChTitle      ' chart title
        .ChartTitle.Characters.Font.Size = chFontSize   ' title font size
        .SeriesCollection(1).XValues = RngX             ' axis signatures
    End With
' adjust chart placement
    With Hsht.ChartObjects(chobj)
        .Left = Hsht.Cells(ulr, ulc).Left
        .Top = Hsht.Cells(ulr, ulc).Top
        .Width = chW
        .Height = chH
' do not resize chart if cells resized
        .Placement = xlFreeFloating
    End With
    Hsht.Cells(ulr, ulc).Activate
End Sub
Sub Scatterplot_My(chsht As Worksheet, _
                        ulr As Integer, _
                        ulc As Integer, _
                        chW As Integer, _
                        chH As Integer, _
                        RngX As Range, _
                        RngY As Range, _
                        ChTitle As String, _
                        chobj As Integer)
    Const chFontSize As Integer = 14    ' chart title font size
    chW = chW * 48      ' chart width, pixels
    chH = chH * 15      ' chart height, pixels
    chsht.Activate                      ' activate sheet
'    RngX.Select
    chsht.Shapes.AddChart.Select        ' build chart
    With ActiveChart
'        .SetSourceData Source:=RngY
        .SeriesCollection.NewSeries
'        .SeriesCollection(1).Name = "sadfasdf"
        .SeriesCollection(1).XValues = RngX
        .SeriesCollection(1).Values = RngY
        .ChartType = xlXYScatter                        ' chart type - OHLC
        .Legend.Delete
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
'        .Axes(xlCategory).TickLabelPosition = xlLow     ' axis position
    End With
    With chsht.ChartObjects(chobj)    ' adjust chart placement
        .Left = chsht.Cells(ulr, ulc).Left
        .Top = chsht.Cells(ulr, ulc).Top
        .Width = chW
        .Height = chH
'        .Placement = xlFreeFloating     ' do not resize chart if cells resized
    End With
    Cells(ulr, ulc).Activate
End Sub


Sub Query_Table_Add()
    Dim i As Integer
    Dim fd As FileDialog
    Dim fpath As String, fname As String, fCount As Integer
    Dim lr As Long
    Dim wb As Workbook, ws As Worksheet, wc As Range
    Dim shc As Integer
    Set wb = ActiveWorkbook
'    lr = Cells(rows.Count, 1).End(xlUp).Row + 3
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Pick a file"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel", "*.csv"
        .ButtonName = "GO!"
    End With
    If fd.Show = -1 Then
' get files count
        fCount = fd.SelectedItems.Count
    Else
' exit if no files picked
        MsgBox "No files picked.", , "Cancel"
'        Failed = True
        Exit Sub
    End If
    Application.ScreenUpdating = False
    For i = 1 To fCount
    ' get name of the folder containing html reports
        fpath = fd.SelectedItems(i)
        fname = Dir(fpath, vbDirectory)
        fname = Left(fname, Len(fname) - 4)
        shc = wb.Sheets.Count
        If shc < i Then
            wb.Sheets.Add after:=wb.Sheets(shc)
        End If
        Set ws = wb.Sheets(i)
        Set wc = ws.Cells
        With ws.QueryTables.Add(Connection:= _
            "TEXT;" & fpath, Destination:=wc(1, 1))
            .Name = fname
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 1252 ' 1251
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = True
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With
        ws.Name = Left(fname, 6)
    Next i
    Application.ScreenUpdating = True
End Sub




