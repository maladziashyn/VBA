Attribute VB_Name = "ISvsOS"
Option Explicit

Dim AlgoStmtSheet As Long

' RANGE
Dim cMix As Range
Dim cSt As Range
Dim cTg As Range

' WORKBOOK
Dim wbMix As Workbook   ' mixer
Dim wbSt As Workbook    ' statement "Positions Close"
Dim wbTg As Workbook    ' target

' WORKSHEET
Dim wsMix As Worksheet  ' mixer
Dim wsSt As Worksheet   ' statement
Dim wsTg As Worksheet   ' target

' RegEx
Dim rePostFix As Object
Dim Matches As Object, Match As Object


Sub Run_ISvsOS()
    
    Const TgFName As String = "is_vs_os"
    
    Dim i As Long
    Dim SheetsInMix As Long
    Dim TgFPath As String
    
    Call OnStart
    
    On Error GoTo eh
    
    ' Target
    Set wbTg = Workbooks.Add()
    
    ' Statement
    Set wbSt = Workbooks.Open(wsIsOs.Range("PathStatement"))
    AlgoStmtSheet = wbSt.Sheets.Count + 1
    Set wsSt = wbSt.Sheets("Positions Close")
    Set cSt = wsSt.Cells
    
    ' Mixer
    Set wbMix = Workbooks.Open(wsIsOs.Range("PathMixer"))
    
    Set rePostFix = New RegExp
    With rePostFix
        .Pattern = "_(mux|mxu|cux|cxu){1}$"
        .Global = False
        .MultiLine = False
    End With
    
    SheetsInMix = wbMix.Sheets.Count - 1
    For i = 1 To SheetsInMix
        Application.StatusBar = "Sheet " & i & " of " & SheetsInMix & "."
        Call OneFromMixer(i)
        wbTg.Sheets(1).Activate
    Next i
    
' Close workbooks
    TgFPath = PathIncrementIndex(wsIsOs.Range("PathTargetDir") & "\" & TgFName, True)
    wbTg.SaveAs fileName:=TgFPath
    wbMix.Close SaveChanges:=False
    wbSt.Close SaveChanges:=False

eh:
    Call OnExit

End Sub
Sub OneFromMixer(ByVal ShId As Long)
    
    Dim StratTag As String
    Dim rgCopy As Range
    Dim rgDescription As Range
    Dim cAlgoStmt As Range
    Dim wsAlgoStmt As Worksheet
    Dim img As Shape
    Dim rangeX As Range, rangeY As Range
    
    Set wsMix = wbMix.Sheets(ShId)
    Set cMix = wsMix.Cells
    StratTag = GetStrategyTag(cMix(1, 2), cMix(2, 2))
    Set rgDescription = cSt.Find(What:=StratTag, after:=cSt(1, 1))
    
    If wbTg.Sheets.Count < ShId Then
        wbTg.Sheets.Add after:=wbTg.Sheets(wbTg.Sheets.Count)
    End If
    Set wsTg = wbTg.Sheets(ShId)
    
    If rgDescription Is Nothing Then
        wsTg.Name = StratTag & "_NOT_FOUND"
        Exit Sub
    Else
        wsTg.Name = StratTag
    End If
    Set cTg = wsTg.Cells
    
    ' Copy from mixer
    wsMix.Activate
    Set rgCopy = wsMix.columns("A:M")
    rgCopy.Copy cTg(1, 1)
    wsTg.Activate
    Call Stats_Chart_from_Joined_Windows_ISOS
    
    
    ' Copy from statement
    wbSt.Activate
    wsSt.Activate
    rgDescription.Activate
    Call DescriptionFilterChart
    For Each img In ActiveSheet.Shapes
        img.Delete
    Next img
    
    Set wsAlgoStmt = wbSt.Sheets(AlgoStmtSheet)
    Set cAlgoStmt = wsAlgoStmt.Cells
    Set rgCopy = wsAlgoStmt.columns("A:AF")
    rgCopy.Copy cTg(1, 28)

    wbTg.Activate
    wsTg.Activate
    
    Set rangeX = cTg(4, 48).CurrentRegion
    Set rangeX = rangeX.Resize(rangeX.rows.Count, rangeX.columns.Count - 1)
    Set rangeY = cTg(4, 48).CurrentRegion
    Set rangeY = rangeY.Offset(0, 1).Resize(rangeY.rows.Count, rangeY.columns.Count - 1)
    Call StatementChartRangesXandY(rangeX, rangeY, cTg, 4, 50, False, _
        1, "Equity curve, filter=" & "description")
    cTg(1, 1).Activate
    
    Application.DisplayAlerts = False
    wsAlgoStmt.Delete
    Application.DisplayAlerts = True
    Set wsAlgoStmt = Nothing
    
End Sub

Sub Stats_Chart_from_Joined_Windows_ISOS()
    
    Dim ws As Worksheet
    Dim Rng As Range
    Dim ubnd As Long
    Dim tradesSet() As Variant
    Dim daysSet() As Variant
    
    Set ws = ActiveSheet
    Set Rng = ws.Cells
    ubnd = Rng(ws.rows.Count, 3).End(xlUp).Row - 1

    ' move to RAM
    tradesSet = Load_Slot_to_RAM2(Rng, ubnd)
    ' add Calendar x2 columns
    daysSet = Get_Calendar_Days_Equity2(tradesSet, Rng)
    ' print out
    Call Print_2D_Array2(daysSet, True, 0, 14, Rng)
    ' build chart
    Call WFA_Chart_Classic2(Rng, 1, 17)

End Sub

Sub WFA_Chart_Classic2(sc As Range, _
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
    ChTitle = "Equity curve, " & sc(1, 2)
'    If Left(sc(1, first_col), 2) = "IS" And logScale Then
'        ChTitle = ChTitle & ", log scale"         ' log scale
'    End If
    last_date_row = sc(2, 15).End(xlDown).Row
    chFontSize = 12
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + ch_hght_cells, ulc + ch_wdth_cells))
    Set rngX = Range(sc(1, 15), sc(last_date_row, 15))
    Set rngY = Range(sc(1, 16), sc(last_date_row, 16))
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
'        If Left(sc(1, first_col), 2) = "IS" And logScale Then
'            .Axes(xlValue).ScaleType = xlScaleLogarithmic   ' log scale
'        End If
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
'    sc(1, first_col + 1) = chObj_idx
    sc(1, 15).Select

End Sub


Private Sub Print_2D_Array2(ByVal print_arr As Variant, _
        ByVal is_inverted As Boolean, _
        ByVal row_offset As Integer, _
        ByVal col_offset As Integer, _
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



Function Get_Calendar_Days_Equity2(ByVal tset As Variant, _
                                   ByVal wc As Range) As Variant
' INVERTED: columns, rows
    
    Dim i As Integer, j As Integer
    Dim arr() As Variant
    Dim date_0 As Date
    Dim date_1 As Date
    Dim calendar_days As Integer
    
    date_0 = Int(wc(8, 2))
    date_1 = Int(wc(9, 2))
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
    Get_Calendar_Days_Equity2 = arr

End Function


Function Load_Slot_to_RAM2(ByVal wc As Range, _
                           ByVal upBnd As Long) As Variant
' Function loads excel report from WFA-sheet to RAM
' Returns (1 To 3, 1 To trades_count) array - INVERTED
    
    Dim arr() As Variant
    Dim i As Long, j As Long
    
    ReDim arr(1 To 3, 1 To upBnd)
    For i = LBound(arr, 2) To UBound(arr, 2)
        j = i + 1
        arr(1, i) = wc(j, 9)     ' open date
        arr(2, i) = wc(j, 10)    ' close date
        arr(3, i) = wc(j, 13)    ' return
    Next i
    Load_Slot_to_RAM2 = arr

End Function

Function GetStrategyTag(ByVal StratName As String, _
        ByVal CurrPair As String) As String
    
    Dim Postfix As String
    Dim Result As String
    
    Result = IIf(rePostFix.Test(StratName), _
        rePostFix.Replace(StratName, "_"), _
        StratName & "_")
    Postfix = GetCurrPairAsPostfix(CurrPair)
    
    GetStrategyTag = Result & Postfix
    
End Function

Function GetCurrPairAsPostfix(ByVal CurrPair As String) As String
    
    Dim c1 As String, c2 As String
    
    If Len(CurrPair) = 7 _
            And InStr(1, CurrPair, "/", vbTextCompare) > 0 Then
        c1 = LCase(Left(CurrPair, 3))
        c1 = IIf(CurrIsCHF(c1), "f", Left(c1, 1))
        c2 = LCase(Right(CurrPair, 3))
        c2 = IIf(CurrIsCHF(c2), "f", Left(c2, 1))
        GetCurrPairAsPostfix = c1 & c2
    Else
        Err.Raise 801, , "Unknown currency pair"
    End If
    
End Function

Private Function CurrIsCHF(ByVal Curr As String) As Boolean
    If Curr = "chf" Then
        CurrIsCHF = True
    End If
End Function

Sub ClickLocatePathMixer()
    Call ClickLocateSomething(True, "Locate 'MIXER' file", _
        wsIsOs.Range("PathMixer"))
End Sub
Sub ClickLocatePathStatement()
    Call ClickLocateSomething(True, "Locate 'STATEMENT' file", _
        wsIsOs.Range("PathStatement"))
End Sub
Sub ClickLocatePathTarget()
    Call ClickLocateSomething(False, "Locate target directory", _
        wsIsOs.Range("PathTargetDir"))
End Sub
Sub ClickLocateSomething(ByVal blFilePicker As Boolean, _
        ByVal strTitle As String, _
        ByRef rgTarget As Range)
' Show file dialog, let user pick a file directory
    Dim fd As FileDialog
    
    If blFilePicker Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    With fd
        .Title = strTitle
        If blFilePicker Then
            .AllowMultiSelect = False
        End If
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    rgTarget.Value = fd.SelectedItems(1)
    wsIsOs.columns(rgTarget.Column).AutoFit
    
End Sub

