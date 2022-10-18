Attribute VB_Name = "Mixer"
Option Explicit
Option Base 1

Const close_date_col As Integer = 8
Const depo_ini As Integer = 10000

Private Sub GSPR_show_sheet_index()
    
    Const msg As String = "Sheet number "
    
    On Error Resume Next
    MsgBox msg & ActiveSheet.Index & "."

End Sub

Private Sub GSPR_Go_to_sheet_index()
    
    Dim sh_idx As Integer

    On Error Resume Next
    sh_idx = InputBox("Enter sheet number:")
    Sheets(sh_idx).Activate

End Sub

Private Sub GSPR_robo_mixer()
    
    Const mix_sheet_name As String = "mix"
    
    Dim i As Integer, j As Integer
    Dim ws As Worksheet, mix_ws As Worksheet
    Dim wc As Range, mix_c As Range, Rng As Range
    Dim first_row As Integer, last_row As Integer
    Dim next_empty_row As Integer
    Dim backtest_days As Integer
    Dim algos_mixed As Integer
    Dim sh_ini As Integer, sh_fin As Integer
    
' +++++++++++++++++++++++++++++++++++++
    Const rf_lower As Double = 0      ' MINIMUM RECOVERY FACTOR
    Const rf_upper As Double = 990    ' MAXIMUM RECOVERY FACTOR
    Const max_tpm As Double = 99      ' MAXIMUM TPM
    Const min_tpm As Double = 0       ' MINIMUM TPM
    Const min_SR As Double = 0
    Const max_SR As Double = 99
' +++++++++++++++++++++++++++++++++++++
    
    sh_ini = InputBox("Enter first sheet index to join trade lists:")
    sh_fin = InputBox("Enter last sheet index to join trade lists:")
    
'    sh_ini = 1
'    sh_fin = 27

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

' create / assign "mix" sheet
    If Sheets(Sheets.count).Name = mix_sheet_name Then
        Set mix_ws = Sheets(Sheets.count)
        mix_ws.Cells.Clear
    Else
        Set mix_ws = Sheets.Add(after:=Sheets(Sheets.count))
        mix_ws.Name = mix_sheet_name
    End If
    Set mix_c = mix_ws.Cells
' copy trades to "mix" sheet
    j = 1
    For i = sh_ini To sh_fin ' Sheets.Count - 1
        Set ws = Sheets(i)
        Set wc = ws.Cells
' **********************************
        If wc(6, 2) >= rf_lower And wc(6, 2) <= rf_upper _
                And wc(3, 2) <= max_tpm And wc(3, 2) >= min_tpm Then
' **********************************
        
' **********************************
'        If wc(21, 2) >= min_SR Then
' **********************************
        
            last_row = wc(1, 3).End(xlDown).Row
            If j = 1 Then
                first_row = 1
                next_empty_row = 3
            Else
                first_row = 2
                next_empty_row = mix_c(mix_ws.Rows.count, 1).End(xlUp).Row + 1
            End If
            Set Rng = ws.Range(wc(first_row, 3), wc(last_row, 13))
            Rng.Copy mix_c(next_empty_row, 1)
            j = j + 1
        End If
    Next i
    Application.StatusBar = "calc"
' robots included
    algos_mixed = j - 1
    mix_c(1, 2) = rf_lower
    mix_c(2, 2) = rf_upper
    mix_c(1, 3) = "robots"
    mix_c(2, 3) = algos_mixed
' add autofilter
    mix_ws.Activate
    mix_ws.Rows("3:3").AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 3
        .FreezePanes = True
    End With
' sort close date ascending
    last_row = mix_c(3, 1).End(xlDown).Row
    Set Rng = mix_ws.Range(mix_c(3, 1), mix_c(last_row, 11))
    Rng.Sort key1:=mix_c(3, close_date_col), order1:=xlAscending, Header:=xlYes
' calculate winning/losing trades
    mix_c(1, 5) = "plus"
    mix_c(2, 5) = "minus"
    Set Rng = mix_ws.Range(mix_c(4, 5), mix_c(last_row, 5))
    mix_c(1, 6) = WorksheetFunction.CountIf(Rng, ">0")
    mix_c(2, 6) = WorksheetFunction.CountIf(Rng, "<=0")
    With mix_c(1, 7)
        .Value = mix_c(1, 6) / (mix_c(1, 6) + mix_c(2, 6))
        .NumberFormat = "0%"
        .Interior.Color = RGB(0, 255, 0)
    End With
    With mix_c(2, 7)
        .Value = mix_c(2, 6) / (mix_c(1, 6) + mix_c(2, 6))
        .NumberFormat = "0%"
        .Interior.Color = RGB(255, 0, 0)
    End With
'
    backtest_days = mix_c(last_row, 8) - mix_c(4, 7)
    mix_c(1, 8) = "days"
    mix_c(2, 8) = backtest_days
    With mix_c(1, 9)
        .Value = 0.01
        .NumberFormat = "0.00%"
    End With
'    mix_c(2, 10) = "mult"
'    mix_c(2, 11) = 1
    mix_c(1, 11) = "start cap."
    mix_c(1, 12) = depo_ini
    mix_c(3, 12) = depo_ini
    mix_c(3, 13) = depo_ini
' calculate trade-to-trade equity curve, hwm
    mix_c(3, 14) = "dd"
    For i = 4 To last_row
        If i Mod 500 = 0 Then
            Application.StatusBar = "Adding formula " & i & " (" & last_row & ")."
        End If
        mix_c(i, 12).FormulaR1C1 = "=R[-1]C*(1+RC[-1]*R1C9*100)"
        mix_c(i, 13).FormulaR1C1 = "=MAX(RC[-1],R[-1]C)"
        With mix_c(i, 14)
            .FormulaR1C1 = "=(RC[-1]-RC[-2])/RC[-1]"
            .NumberFormat = "0.00%"
        End With
    Next i
    mix_c(2, 11) = "end cap."
    mix_c(2, 12).Formula = "=R" & last_row & "C"
' print out statistics
    mix_c(1, 14) = "MDD"
    mix_c(2, 14).FormulaR1C1 = "=MAX(R4C:R" & last_row & "C)"
    mix_c(1, 15) = "Net %"
    With mix_c(2, 15)
        .FormulaR1C1 = "=(R2C12-R3C12)/R3C12"
        .NumberFormat = "0.00%"
    End With
    mix_c(1, 16) = "Ann %"
    With mix_c(2, 16)
        .FormulaR1C1 = "=(1+R2C15)^(365/R2C8)-1"
        .NumberFormat = "0.00%"
    End With
    mix_c(1, 17) = "Recov"
    With mix_c(2, 17)
        .FormulaR1C1 = "=R2C16/R2C14"
        .NumberFormat = "0.00"
    End With
    mix_c(1, 18) = "Trades"
    mix_c(2, 18).FormulaR1C1 = "=COUNT(R4C8:R" & last_row & "C8)"
    mix_c(1, 19) = "Per Mn"
    With mix_c(2, 19)
        .FormulaR1C1 = "=R2C18/(R2C8/30.4)"
        .NumberFormat = "0"
    End With
    mix_c(1, 20) = "R2"
' win count
    mix_c(1, 22) = "wCount"
    With mix_c(2, 22)
        .FormulaR1C1 = "=COUNTIF(R4C11:R" & last_row & "C11,"">""&0)"
        .NumberFormat = "0"
    End With
' los count
    mix_c(1, 23) = "losCount"
    With mix_c(2, 23)
        .FormulaR1C1 = "=COUNTIF(R4C11:R" & last_row & "C11,""<=""&0)"
        .NumberFormat = "0"
    End With
' avg winner
    mix_c(1, 24) = "avWin"
    With mix_c(2, 24)
        .FormulaR1C1 = "=AVERAGEIF(R4C11:R" & last_row & "C11,"">""&0)"
        .NumberFormat = "0.00%"
    End With
' avg loser
    mix_c(1, 25) = "avLos"
    With mix_c(2, 25)
        .FormulaR1C1 = "=ABS(AVERAGEIF(R4C11:R" & last_row & "C11,""<=""&0))"
        .NumberFormat = "0.00%"
    End With
' average winner / loser
    mix_c(1, 21) = "avW/avL"
    With mix_c(2, 21)
        .FormulaR1C1 = "=R2C24/R2C25"
        .NumberFormat = "0.0000"
    End With
' sharpe ratio
    mix_c(1, 26) = "SharpeRatio"
    With mix_c(2, 26)
        .FormulaR1C1 = "=R2C16/(STDEV(R2C11:R" & last_row & "C11)*SQRT(250))"
        .NumberFormat = "0.000"
    End With
    mix_ws.Name = mix_sheet_name & "_" & algos_mixed & "_" & Sheets.count
    
    With mix_c(3, 26)
        .Value = min_SR
        .NumberFormat = "0.00"
    End With
    With Application
        .StatusBar = False
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    MsgBox "Done"

End Sub

Private Sub GSPR_trades_to_days()
    
    Dim date_from As Date
    Dim date_to As Date
    Dim last_row As Integer
    Const first_row As Integer = 4
    Dim wc As Range
    Const dt_pr_fr As Integer = 4
    Const dt_pr_fc As Integer = 16
    Const ch_build_1st_col = 18
    Dim pc_coeff As Double
    Dim dates() As Long
    Dim day_pc() As Double
    Dim day_fr() As Double
    Dim i As Integer, r As Integer, c As Integer, j As Integer
    ' chart
    Dim rng_x As Range, rng_y As Range
    Dim ch_title As String
    Dim min_val As Long
    Dim max_val As Currency
    
    Application.ScreenUpdating = False
    Set wc = ActiveSheet.Cells
    last_row = wc(first_row, 1).End(xlDown).Row
    date_from = Int(wc(first_row, close_date_col - 1)) - 1
    date_to = Int(wc(last_row, close_date_col))
    pc_coeff = wc(1, 9).Value * 100
    
    ReDim dates(1 To date_to - date_from + 1)
    ReDim day_pc(1 To UBound(dates))
    ReDim day_fr(1 To UBound(dates))
' populate dates, daily % changes
    dates(1) = date_from
    day_pc(1) = 1
    For i = 2 To UBound(dates)
        dates(i) = dates(i - 1) + 1
        day_pc(i) = 1
    Next i
' populate daily percent changes and financial result
    day_fr(1) = depo_ini
    j = first_row
    For i = 2 To UBound(day_pc)
        If dates(i) = CDate(Int(wc(j, close_date_col))) Then
            Do While CDate(Int(wc(j, close_date_col))) = dates(i) And j <= last_row
                day_pc(i) = day_pc(i) * (1 + wc(j, 11) * pc_coeff)
                j = j + 1
            Loop
            day_fr(i) = day_fr(i - 1) * day_pc(i)
        Else
            day_fr(i) = day_fr(i - 1) * day_pc(i)
        End If
    Next i
' print dates array
    For r = dt_pr_fr To UBound(dates) + dt_pr_fr - 1
        i = r - dt_pr_fr + 1
        With wc(r, dt_pr_fc)
            .Value = dates(i)
            .NumberFormat = "dd.mm.yy"
        End With
'        wc(r, dt_pr_fc + 2) = day_pc(i)
        With wc(r, dt_pr_fc + 1)
            .Value = day_fr(i)
            .NumberFormat = "0"
        End With
    Next r
' R-square
    With wc(2, 20)
        .Value = WorksheetFunction.RSq(dates, day_fr)
        .NumberFormat = "0.00"
    End With
' build chart
    Set rng_x = Range(wc(first_row, dt_pr_fc), wc(first_row + UBound(dates) - 1, dt_pr_fc))
    Set rng_y = Range(wc(first_row, dt_pr_fc + 1), wc(first_row + UBound(dates) - 1, dt_pr_fc))
    ch_title = "All algos, annualized ret. (" & wc(2, 3).Value & "), year=" & Round(wc(2, 16).Value * 100, 0) & "%, depo fin.=" & Round(day_fr(UBound(day_fr)), 0) & " usd. Log scale."
'    ch_title = "All algos, annualized ret.=" & Round(wc(2, 16).Value * 100, 0) & "%, depo fin.=" & Round(day_fr(UBound(day_fr)), 0) & " usd. Log scale."
    min_val = WorksheetFunction.Min(day_fr) - 1
    max_val = WorksheetFunction.Max(day_fr)
    Call Merged_Chart_Classic_wMinMax(dt_pr_fr, ch_build_1st_col, _
                    rng_x, rng_y, ch_title, _
                    min_val, max_val)
    wc(first_row, dt_pr_fc).Select
    Application.ScreenUpdating = True

End Sub

Private Sub Merged_Remove_Charts()
    
    Dim img As Shape
    
    For Each img In ActiveSheet.Shapes
        img.Delete
    Next

End Sub

Private Sub Merged_Chart_Classic_wMinMax(ulr As Integer, _
        ulc As Integer, _
        rngX As Range, _
        rngY As Range, _
        ChTitle As String, _
        MinVal As Long, _
        maxVal As Currency)
    
    Dim chW As Integer, chH As Integer          ' chart width, chart height
    Dim chFontSize As Integer                   ' chart title font size
    
'    With Application
'        .ScreenUpdating = False
'        .Calculation = xlCalculationManual
'        .EnableEvents = False
'    End With
    
    Call Merged_Remove_Charts
    chW = 624   ' standrad cell width = 48 pix
    chH = 330   ' standard cell height = 15 pix
    chFontSize = 12
' build chart
    rngY.Select
    ActiveSheet.Shapes.AddChart.Select
' adjust chart placement
    With ActiveSheet.ChartObjects(1)
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = chW
        .Height = chH
'        .Placement = xlFreeFloating ' do not resize chart if cells resized
    End With
    With ActiveChart
        .SetSourceData Source:=Application.Union(rngX, rngY)
        .ChartType = xlLine
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = maxVal
        .Axes(xlValue).ScaleType = xlLogarithmic    ' LOGARITHMIC
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
'    With Application
'        .EnableEvents = True
'        .Calculation = xlCalculationAutomatic
'        .ScreenUpdating = True
'    End With

End Sub

Private Sub GSPR_Mixer_Copy_Sheet_To_Book()
    
    Dim wb_from As Workbook, wb_to As Workbook
    Dim sh_copy As Worksheet
    Dim new_name As String, ins_abbrev As String
    
    Application.ScreenUpdating = False
    Set wb_from = ActiveWorkbook
    Set sh_copy = ActiveSheet
    ins_abbrev = gspr_get_instrument_abbrev(sh_copy.Cells(2, 2))
    new_name = sh_copy.Cells(1, 2).Value & "_" & ins_abbrev & "_" & sh_copy.Name
    Set wb_to = Workbooks("mixer.xlsx")
' copy
    sh_copy.Copy after:=wb_to.Sheets(wb_to.Sheets.count)
    wb_to.ActiveSheet.Name = new_name
'    wb_from.Activate
'    wb_from.Close savechanges:=False
    Application.ScreenUpdating = True

End Sub

Private Function gspr_get_instrument_abbrev(ByVal fed_ins As String) As String
    
    Dim nomin As String, denom As String
    
    nomin = Left(fed_ins, 3)
    denom = Right(fed_ins, 3)
    If nomin = "CHF" Then
        nomin = "f"
    Else
        nomin = LCase(Left(nomin, 1))
    End If
    If denom = "CHF" Then
        denom = "f"
    Else
        denom = LCase(Left(denom, 1))
    End If
    gspr_get_instrument_abbrev = nomin & denom

End Function

Sub SharpesToSeparateSheet()
    
    Dim i As Integer
    Dim ws As Worksheet
    Dim c As Range
    
    Set ws = ActiveSheet
    Set c = ws.Cells
    c(1, 1) = "sheet"
    c(1, 2) = "SR"
    
    For i = 1 To 27
        c(i + 1, 1) = i
        c(i + 1, 2) = Sheets(i).Cells(21, 2)
    Next i

End Sub


