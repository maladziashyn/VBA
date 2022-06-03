

' MODULE: ThisWorkbook
' empty


' MODULE: Sheet1
' empty


' MODULE: TB_Subs
Option Explicit

Sub ActiveSheetDelete()
' Delete active worksheet without displaying alert.
' No undo.

    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True

End Sub

Sub RefStyleToggle()
' Toggle between R1C1 and "A1" reference style.

    If Application.ReferenceStyle = 1 Then
        Application.ReferenceStyle = -4150
    Else
        Application.ReferenceStyle = 1
    End If

End Sub

Sub GetCellType()
' By John Walkenbach
    
    Dim c As Range
    
    Set c = ActiveCell
    
    Select Case True
        Case IsEmpty(c): MsgBox "Blank"
        Case Application.IsText(c): MsgBox "Text"
        Case Application.IsLogical(c): MsgBox "Logical"
        Case Application.IsErr(c): MsgBox "Error"
        Case IsDate(c): MsgBox "Date"
        Case InStr(1, c.Text, ":") <> 0: MsgBox "Time"
        Case IsNumeric(c): MsgBox "Value"
    End Select

End Sub

Sub ImagesDelete()
' Remove all images from active worksheet

    Dim img As Shape
    
    For Each img In ActiveSheet.Shapes
        img.Delete
    Next

End Sub

Sub RunStuffBefore()
' Run this before main macro to speed up execution.
    
    Application.ScreenUpdating = False

End Sub

Sub RunStuffAfter()
' Run this after main macro to restore things as they were.

    Application.StatusBar = False
    Application.ScreenUpdating = True

End Sub

Sub ExportChartToImage()
' Export chart to image
    
    Dim ws As Excel.Worksheet
    Dim SaveToDirectory As String, myFileName As String
    Dim objChrt As ChartObject
    Dim myChart As Chart

    SaveToDirectory = Environ("UserProfile") & "\Desktop\ImgExp\"
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        For Each objChrt In ws.ChartObjects
            objChrt.Activate
            Set myChart = objChrt.Chart
            myFileName = SaveToDirectory & ws.Name & "_" & objChrt.Index & ".gif"
            myChart.Export fileName:=myFileName, Filtername:="GIF"
        Next
    Next

End Sub

Sub ExportAllChartsToImages()
' Export chart from one sheet to image
    
    Dim ws As Worksheet
    Dim SaveToDirectory As String, myFileName As String
    Dim objChrt As ChartObject
    Dim myChart As Chart

    SaveToDirectory = Environ("UserProfile") & "\Desktop\"
    Set ws = ActiveSheet
    
    For Each objChrt In ws.ChartObjects
        objChrt.Activate
        Set myChart = objChrt.Chart
        myFileName = SaveToDirectory & ws.Name & "_" & objChrt.Index & ".gif"
        myChart.Export fileName:=myFileName, Filtername:="GIF"
    Next objChrt
    
    ws.Cells(1, 1).Activate

End Sub

Sub ShowFileDialogFilePicker()
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Your Title"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Your Filter Text", "*.xlsx"
        .ButtonName = "Your Button Name"
    End With
    
    ' If user cancels
    If fd.Show = 0 Then
        MsgBox "No files selected!", , "Message box title"
        Exit Sub
    End If
    
    itemsSelected = fd.SelectedItems
    itemsCount = fd.SelectedItems.Count

End Sub

Sub CheckIfDirectoryExists()
    
    If Dir("C:\Windows", vbDirectory) <> vbNullString Then ' or ""
        MsgBox "Exists"
    Else
        MsgBox "Doesn't exist"
    End If
End Sub

Sub FetchNames()

    Dim myPath As String, myfile As String
    Dim r As Long
    
    myPath = "C:\SomeFolder1\SomeFolder2\"
    myfile = Dir(myPath & "*.html")
    
    r = 2
    Do While myfile <> ""
        Cells(r, 1).Value = myfile
        r = r + 1
        myfile = Dir
    Loop

End Sub

Sub ColorScatterPoints()
' Taken from somewhere on the internet

    Dim cht As Chart
    Dim srs As Series
    Dim pt As Point
    Dim p As Long
    Dim Vals$, lTrim#, rTrim#
    Dim valRange As Range, cL As Range
    Dim myColor As Long

    Set cht = ActiveSheet.ChartObjects(1).Chart
    Set srs = cht.SeriesCollection(1)

    ' Get the series Y-Values range address
    lTrim = InStrRev(srs.Formula, ",", InStrRev(srs.Formula, ",") - 1, vbBinaryCompare) + 1
    rTrim = InStrRev(srs.Formula, ",")
    Vals = Mid(srs.Formula, lTrim, rTrim - lTrim)
    Set valRange = Range(Vals)

    For p = 1 To srs.Points.Count
        Set pt = srs.Points(p)
        Set cL = valRange(p).Offset(0, 1) '## assume color is in the next column.

        With pt.Format.Fill
            .Visible = msoTrue
            '.Solid  'I commented this out, but you can un-comment and it should still work
            '## Assign Long color value based on the cell value
            '## Add additional cases as needed.
            Select Case LCase(cL)
                Case "red"
                    myColor = RGB(255, 0, 0)
                Case "orange"
                    myColor = RGB(255, 192, 0)
                Case "green"
                    myColor = RGB(0, 255, 0)
            End Select

            .ForeColor.RGB = myColor

        End With
    Next

End Sub

Sub GetSelectedRows()
' Show first and last selected row in message box.

   Dim iRowFirst As Long
   Dim iRowLast As Long

   iRowFirst = Selection.Row
   iRowLast = Selection.Row + Selection.Rows.Count - 1

   MsgBox iRowFirst & " to " & iRowLast

End Sub

Sub SelectionConditionalFormattingDelete()
' Delete format conditions from selection
    
    Selection.FormatConditions.Delete
    
End Sub

Sub SelectionConditionalFormattingDataBar()
' Apply conditional formatting to the selected region.
' Paint green/red databars.

    Dim Thresh As Double
    Dim aRed As Integer, aGreen As Integer, aBlue As Integer
    Dim bRed As Integer, bGreen As Integer, bBlue As Integer
    Dim dbMinVal As Double, dbMaxVal As Double
    Dim RngDataBar As Range

    Call RunStuffBefore

    Set RngDataBar = Selection
    RngDataBar.FormatConditions.Delete
    ' Settings
    Thresh = 0
    ' Above threshhold = green, below = red
    aRed = 0
    aGreen = 255
    aBlue = 0
    bRed = 255
    bGreen = 0
    bBlue = 0
    
    ' Find Min and Max values
    dbMinVal = WorksheetFunction.Min(RngDataBar)
    dbMaxVal = WorksheetFunction.Max(RngDataBar)
    
    Call AddDataBarPaintSep( _
        RngDataBar, _
        Thresh, dbMinVal, dbMaxVal, _
        aRed, aGreen, aBlue, _
        bRed, bGreen, bBlue)

    Call RunStuffAfter

End Sub


Sub AddDataBarPaintSep( _
        ByRef RngDataBar As Range, _
        ByRef Thresh As Double, ByRef dbMinVal As Double, ByRef dbMaxVal As Double, _
        ByRef aRed As Integer, ByRef aGreen As Integer, _
        ByRef aBlue As Integer, ByRef bRed As Integer, _
        ByRef bGreen As Integer, ByRef bBlue As Integer)
' This is part of Sub "SelectionConditionalFormattingDataBar"

    Dim dbCell As Range
    
    For Each dbCell In RngDataBar
        dbCell.FormatConditions.AddDatabar
        dbCell.FormatConditions(dbCell.FormatConditions.Count).ShowValue = True
        With dbCell.FormatConditions(1)
'            .MinPoint.Modify newtype:=xlConditionValueLowestValue
'            .MaxPoint.Modify newtype:=xlConditionValueHighestValue
            .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=dbMinVal
            .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=dbMaxVal

            ' Paint cell
            If Thresh = 0 Then
                If dbCell.Value > Thresh Then
                    .BarColor.Color = RGB(aRed, aGreen, aBlue)
                ElseIf dbCell.Value = Thresh Then
                    .BarColor.Color = RGB(255, 255, 255)
                Else
                    .BarColor.Color = RGB(bRed, bGreen, bBlue)
                End If
            Else
                If dbCell.Value >= Thresh Then
                    .BarColor.Color = RGB(aRed, aGreen, aBlue)
                Else
                    .BarColor.Color = RGB(bRed, bGreen, bBlue)
                End If
            End If
        End With
    Next dbCell
    
End Sub

Sub SelectedColumnsIDs()
' Print info on selected columns, in Immediate window.
' Works with multiple selections, noncontiguous.

    Debug.Print Selection.Row
    Debug.Print Selection.Rows.Count
    Debug.Print Selection.Column
    Debug.Print Selection.Columns.Count
    Debug.Print Selection.Columns.item(1).Rows.Count & vbNewLine
    Debug.Print "areas " & Selection.Areas.Count & vbNewLine
    Debug.Print "cols in area " & Selection.Areas.item(1).Columns.Count
    
End Sub


Sub Print2DimArray(ByVal print_arr As Variant, ByVal is_inverted As Boolean, _
                   ByVal row_offset As Integer, ByVal col_offset As Integer, _
                   ByVal print_cells As Range)
' Procedure prints any 2-dimensional array in a new Workbook, sheet 1.
' Arguments:
' - print_arr: 2-dimensional array as Variant
' - is_inverted: True if 1st dimension is columns, 2nd dimension is rows; False otherwise
' - row_offset
' - col_offset
' - print_cells: cells of the worksheet where to print the array
    
    Dim wb_print As Workbook
    Dim c_print As Range
    Dim r As Integer, c As Integer
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
    
    Set wb_print = Workbooks.Add
    Set c_print = wb_print.Sheets(1).Cells
    
    For r = LBound(print_arr, row_dim) To UBound(print_arr, row_dim)
        print_row = r + add_rows + row_offset
        For c = LBound(print_arr, col_dim) To UBound(print_arr, col_dim)
            print_col = c + add_cols + col_offset
            If is_inverted Then
                print_cells(print_row, print_col) = print_arr(c, r)
            Else
                print_cells(print_row, print_col) = print_arr(r, c)
            End If
        Next c
    Next r
    
End Sub

Sub Print1DimArray(ByVal print_arr As Variant, ByVal col_offset As Integer, ByVal print_cells As Range)
' Procedure prints any 1-dimensional array in a new Workbook, sheet 1.
' Parameters:
' - print_arr: 1-dimensional array as Variant
' - row_offset
' - col_offset
' - print_cells: cells of the worksheet where to print the array
    
    Dim wb_print As Workbook
    Dim r As Integer
    Dim print_row As Integer, print_col As Integer
    Dim add_rows As Integer

    If LBound(print_arr) = 0 Then
        add_rows = 1
    Else
        add_rows = 0
    End If
    
    Set wb_print = Workbooks.Add
    Set c_print = wb_print.Sheets(1).Cells
    print_col = 1 + col_offset
    
    For r = LBound(print_arr) To UBound(print_arr)
        print_row = r + add_rows
        print_cells(print_row, print_col) = print_arr(r)
    Next r

End Sub

' MODULE: TB_SnippetsDontRun
Sub Snippets()

' Insert a column before colNumber
    Columns(colNumber).Insert

' Apply autofilter to rowNumber
    Rows(rowNumber).AutoFilter

' Set column width
    Columns(colNumber).ColumnWidth = 10

' Remove cell color
    With ActiveSheet.Columns("A:A").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

' Enable error handling, go to the following statement
    On Error Resume Next

' Disable error handling
    On Error GoTo 0

' Copy worksheet to another workbook
    shToCopy.Copy after:=wbToInsert.Sheets(wbToInsert.Worksheets.Count)

' Save as XLSB
    ActiveWorkbook.SaveAs fileName:=yourFileName, FileFormat:=xlExcel12

' Add hyperlink to cell
    ActiveSheet.Hyperlinks.Add anchor:=yourCell, Address:="", SubAddress:="Sheet1!A1"
    ' , TextToDisplay:="MyHyperlink"

' Make a directory
    MkDir yourPath

' Find row or column
    f = yourSheet.Cells.Find(what:="FindString", _
        after:=yourSheet.Cells(1, 1), _
        LookIn:=xlValues, _
        lookat:=xlWhole, _
        searchorder:=xlByRows, _
        searchdirection:=xlNext, _
        MatchCase:=False, _
        searchformat:=False).Column

' A simpler find
    f = Cells.Find(what:="what").Column

' Text property
    Set c = Worksheets(1).Range("B14")
    c.Value = 1198.3
    c.NumberFormat = "$#,##0_);($#,##0)"
    MsgBox c.Value
    MsgBox c.Text

' Count cells in range
    MsgBox Range("A1:C3").Count

' Get range address
    MsgBox Range(Cells(1, 1), Cells(5, 5)).Address

End Sub

' MODULE: Unsorted
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





' MODULE: TB_Functions
Option Explicit

Function ListFiles(ByVal sPath As String) As Variant
' Return list of files in a directory
' as an array.
    
    Dim vaArray As Variant
    Dim i As Integer
    Dim oFile As Object
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFiles As Object
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    Set oFiles = oFolder.Files
    If oFiles.Count = 0 Then Exit Function
    ReDim vaArray(1 To oFiles.Count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = oFile.Name
        i = i + 1
    Next
    ListFiles = vaArray

End Function

Function IsStringInList(ByVal whatString As String, whatList As Variant) As Boolean
' True if string is found in the list.
' Pass the list as Array.

    IsStringInList = Not (IsError(Application.Match(whatString, whatList, 0)))

End Function


' MODULE: VersionControl
Option Explicit

' Run GitSave() to export code and modules.
'
' Source: https://github.com/Vitosh/VBA_personal/blob/master/VBE/GitSave.vb
' Source is slightly modified to include a list of modules to ignore.

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
    
    parentFolder = ThisWorkbook.Path
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
    Dim pathToExport As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        If Not IsStringInList(item.Name, ignoreList) Then
            lineToPrint = vbNewLine & "' MODULE: " & item.CodeModule.Name & vbNewLine
            If item.CodeModule.CountOfLines > 0 Then
                lineToPrint = lineToPrint & item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
            Else
                lineToPrint = lineToPrint & "' empty" & vbNewLine
            End If
            textToPrint = textToPrint & vbCrLf & lineToPrint
        End If
    Next item
    
    pathToExport = parentFolder & dirNameCode
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
    
    Call SaveTextToFile(textToPrint, pathToExport & "\all_code.vb")
    
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
            
            If Dir(pathToExport) <> "" Then
                Kill pathToExport & "*.*"
            End If
            
            Call SaveTextToFile(lineToPrint, pathToExport & "\" & item.CodeModule.Name & "_code.vb")
        End If
    Next item

End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    Dim pathToExport As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        If Not IsStringInList(lineToPrint, ignoreList) Then
            textToPrint = textToPrint & vbCrLf & lineToPrint
        End If
    Next item
    
    pathToExport = parentFolder & dirNameCode
    
    Call SaveTextToFile(textToPrint, pathToExport & "\all_modules.vb")
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String
    Dim wkb As Workbook
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean
    
    pathToExport = parentFolder & dirNameModules
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
    
    Set wkb = Excel.Workbooks(ThisWorkbook.Name)

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
                component.Export pathToExport & "\" & filePath
            End If
        End If
    Next
    
End Sub

Sub SaveTextToFile(ByRef dataToPrint As String, ByRef pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim newFile  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then
        MkDir ThisWorkbook.Path & newFile
    End If
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close

End Sub

Function IsStringInList(ByVal whatString As String, whatList As Variant) As Boolean
' True if string is found in the list.
' Pass the list as Array.

    IsStringInList = Not (IsError(Application.Match(whatString, whatList, 0)))

End Function

