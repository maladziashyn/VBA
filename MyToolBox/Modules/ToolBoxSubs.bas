Attribute VB_Name = "ToolBoxSubs"
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
