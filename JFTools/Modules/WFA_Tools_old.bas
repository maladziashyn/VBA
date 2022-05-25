Attribute VB_Name = "WFA_Tools_old"
Option Explicit
    Const logScale As Boolean = False
    Const addin_name As String = "JFTools_0.01.xlsm"
    Const settings_shname As String = "WFA Main"
    Dim first_row As Integer, last_row As Integer
    Dim first_col As Integer, last_col As Integer
Sub Stats_And_Chart()
    Dim tset() As Variant
    Dim dset() As Variant
    Dim wc As Range
    Dim this_col As Integer
    Dim ch_obj_id As Integer
    
    ' sanity check
    If Cells(2, 1) <> "Parameters" Then
        Exit Sub
    ElseIf ActiveCell.Column < 11 Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    If IsEmpty(ActiveCell) Then
        Exit Sub
    End If
    Set wc = ActiveSheet.Cells
    this_col = ActiveCell.Column
    first_row = 3
    If Not IsEmpty(ActiveCell.Offset(0, -1)) Then
        first_col = wc(2, this_col).End(xlToLeft).Column
    Else
        first_col = this_col
    End If
    If IsEmpty(wc(first_row, first_col)) Then
        Exit Sub
    End If
    last_row = Cells(first_row - 1, first_col).End(xlDown).Row
    last_col = first_col + 4
    
    ch_obj_id = Cells(1, first_col + 1)
    If ch_obj_id > 0 Then
        Call Clean_Days_And_Chart(ch_obj_id)
        Exit Sub
    End If
    ' move to RAM
    tset = Load_Slot_to_RAM(wc)
    ' add Calendar x2 columns
    dset = Get_Calendar_Days_Equity(tset)
    ' print out
    Call tWFA_Print_2D_Array(dset, True, 1, last_col, wc)
    ' build chart
    Call WFA_Chart_Classic(wc, 3, first_col)
    Application.ScreenUpdating = True
End Sub
Sub Clear_Folders()
' clears column "A" on "settings" sheet
    Dim Rng As Range
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Set ws = ActiveSheet
    Set Rng = ws.Range(Cells(2, 1), Cells(ws.rows.Count, 1))
    Rng.Clear
    Application.ScreenUpdating = True
End Sub
Sub Insert_Default_Folders()
' "T" = 20th column
    Dim Rng As Range, c As Range
    Dim ws As Worksheet
    Dim last_row As Integer
    
    Application.ScreenUpdating = False
    Set ws = ActiveSheet
    Set c = ws.Cells
    last_row = c(ws.rows.Count, 20).End(xlUp).Row
    If last_row > 1 Then
        Set Rng = ws.Range(c(2, 20), c(last_row, 20))
        Rng.Copy c(2, 1)
    Else
        c(2, 1) = "default folders not found"
    End If
    Application.ScreenUpdating = True
End Sub
Sub Clean_Days_And_Chart(ByVal ch_obj_id As Integer)
    Dim Rng As Range
    Dim days_last_row As Integer
    
    ActiveSheet.ChartObjects(ch_obj_id).Delete
    Cells(1, first_col + 1).Clear
    Call Decrease_Ch_Index(ch_obj_id)
    days_last_row = Cells(2, last_col + 1).End(xlDown).Row
    Set Rng = Range(Cells(first_row - 1, last_col + 1), Cells(days_last_row, last_col + 2))
    Rng.Clear
End Sub
Sub Decrease_Ch_Index(ByVal ch_obj_id As Integer)
    Dim i As Integer
    Dim the_last_col As Integer
    
    the_last_col = Cells(1, columns.Count).End(xlToLeft).Column
    For i = 12 To the_last_col + 1 Step 10
        If Cells(1, i).Value > ch_obj_id Then
            Cells(1, i).Value = Cells(1, i).Value - 1
        End If
    Next i
End Sub
Sub WFA_Chart_Classic(sc As Range, _
                ulr As Integer, _
                ulc As Integer)
    Const ch_wdth_cells As Integer = 9
    Const ch_hght_cells As Integer = 20
    Const my_rnd = 0.1
    Dim last_date_row As Integer
    Dim rngX As Range, rngY As Range
    Dim ChTitle As String
    Dim MinVal As Double, maxVal As Double
    Dim chFontSize As Integer                   ' chart title font size
    Dim rng_to_cover As Range
    Dim chObj_idx As Integer
    
    chObj_idx = ActiveSheet.ChartObjects.Count + 1
    ChTitle = sc(1, first_col)
    If Left(sc(1, first_col), 2) = "IS" And logScale Then
        ChTitle = ChTitle & ", log scale"         ' log scale
    End If
    last_date_row = sc(2, last_col + 1).End(xlDown).Row
    chFontSize = 12
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + ch_hght_cells, ulc + ch_wdth_cells))
    Set rngX = Range(sc(2, last_col + 1), sc(last_date_row, last_col + 1))
    Set rngY = Range(sc(2, last_col + 2), sc(last_date_row, last_col + 2))
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
        If Left(sc(1, first_col), 2) = "IS" And logScale Then
            .Axes(xlValue).ScaleType = xlScaleLogarithmic   ' log scale
        End If
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
    sc(1, first_col + 1) = chObj_idx
    sc(1, first_col).Select
End Sub
Function Get_Calendar_Days_Equity(ByVal tset As Variant) As Variant
    Dim i As Integer, j As Integer
    Dim arr() As Variant
    Dim date_0 As Date
    Dim date_1 As Date
    Dim calendar_days As Integer
    
    date_0 = Int(tset(1, 1))
    date_1 = Int(tset(2, UBound(tset, 2)))
    calendar_days = date_1 - date_0 + 2
    ReDim arr(1 To 2, 1 To calendar_days)
        ' 1. calendar days
        ' 2. equity curve
    arr(1, 1) = date_0 - 1
    arr(2, 1) = 1
    j = 1
    For i = 2 To UBound(arr, 2)
        arr(1, i) = arr(1, i - 1) + 1   ' populate with dates
        arr(2, i) = arr(2, i - 1)
        If arr(1, i) = Int(tset(2, j)) Then
            Do While arr(1, i) = Int(tset(2, j)) ' And j <= UBound(trades_arr, 2)
                arr(2, i) = arr(2, i) * (1 + tset(3, j))
                If j < UBound(tset, 2) Then
                    j = j + 1
                ElseIf j = UBound(tset, 2) Then
                    Exit Do
                End If
            Loop
        End If
    Next i
    Get_Calendar_Days_Equity = arr
End Function
Function Load_Slot_to_RAM(ByVal wc As Range) As Variant
' Function loads excel report from WFA-sheet to RAM
' Returns (1 To 3, 1 To trades_count) array - INVERTED
    Dim arr() As Variant
    Dim i As Integer, j As Integer
    
    ReDim arr(1 To 3, 1 To last_row - first_row + 1)
    For i = LBound(arr, 2) To UBound(arr, 2)
        j = i + 2
        arr(1, i) = wc(j, first_col)        ' open date
        arr(2, i) = wc(j, first_col + 1)    ' close date
        arr(3, i) = wc(j, first_col + 3)    ' return
    Next i
    Load_Slot_to_RAM = arr
End Function
Sub Copy_Dates_Close_Book()
    Dim wb As Workbook
    Dim macro_book As Workbook
    Dim date_1_copy As Date
    Dim date_2_copy As Date
    Dim wb_path As String
    
    Application.ScreenUpdating = False
    Set macro_book = Workbooks(addin_name)
    Set wb = ActiveWorkbook
    If wb.Sheets(3).Cells(8, 1) <> "Начало теста" Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    date_1_copy = wb.Sheets(3).Cells(8, 2)
    date_2_copy = wb.Sheets(3).Cells(9, 2)
    wb_path = wb.Path
    wb.Close savechanges:=False
    macro_book.Sheets(settings_shname).Cells(2, 4) = date_1_copy
    macro_book.Sheets(settings_shname).Cells(3, 4) = date_2_copy
    macro_book.Sheets(settings_shname).Cells(9, 4) = wb_path
    Application.ScreenUpdating = True
End Sub
Private Sub tWFA_Print_2D_Array(ByVal print_arr As Variant, ByVal is_inverted As Boolean, _
                       ByVal row_offset As Integer, ByVal col_offset As Integer, _
                       ByVal print_cells As Range)
' Procedure prints any 2-dimensional array in a new Workbook, sheet 1.
' Arguments:
'       1) 2-D array
'       2) rows-colums (is_inverted = False) or columns-rows (is_inverted = True)
    Dim r As Long
    Dim c As Integer
    Dim print_row As Long
    Dim print_col As Integer
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

