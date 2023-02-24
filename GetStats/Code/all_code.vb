

' MODULE: ThisWorkbook
Option Explicit

Private Sub Workbook_Open()
    
    Call GSPR_Create_CommandBar

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    Call GSPR_Remove_CommandBar

End Sub

' MODULE: Mixer
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



' MODULE: Rep_Extra
Option Explicit
Option Base 1
    Const rep_type As String = "GS_Pro_Single_Core"
    Dim ch_rep_type As Boolean
' macro version
    Const macro_name As String = "GetStats Pro v1.21"
    Const report_type As String = "GS_Pro_Single_Extra"
'
    Const logs_ubd As Integer = 13
' display logs
    Const di_r As Integer = 40
    Const di_c As Integer = 5
' display charts
    Dim rep_adr As String
' arrays
    Dim SV() As Variant                 ' stats values
    Dim SN() As String                  ' stats names
    Dim sv_sn_ubound As Integer
'
    Dim Tlog() As Variant               ' trade log
    Dim Tlog_head(1 To 11) As String
    Dim Aux_T() As Variant             ' auxiliary trade log
'
    Dim Dlog() As Variant               ' daily log & instrument
    Dim Dlog_head(1 To logs_ubd) As String
    Dim Aux_D() As Variant
    Dim Aux_D_Sk_inv() As Integer
'
    Dim Wlog() As Variant               ' weekly log
    Dim Wlog_head(1 To logs_ubd) As String
    Dim Aux_W() As Variant
'
    Dim Mlog() As Variant               ' monthly log
    Dim Mlog_head(1 To logs_ubd) As String
    Dim Aux_M() As Variant
'
    Dim Par() As Variant                ' parameters
'
' formats
    Dim sect_fm_A() As Variant          ' sections formats - rows numbers - column 1
    Dim sect_fm_C() As Variant          ' sections formats - rows numbers - column 3
    Dim fm_date_A() As Variant
    Dim fm_date_C() As Variant
    Dim fm_0_A() As Variant
    Dim fm_0_C() As Variant
    Dim fm_0p00_A() As Variant
    Dim fm_0p00_C() As Variant
    Dim fm_0p00pc_A() As Variant
    Dim fm_0p00pc_C() As Variant
'
    Dim r_s_report As Integer, r_name As Integer, r_version As Integer, r_type As Integer, r_date_gen As Integer, r_time_gen As Integer, r_file As Integer, r_s_basic As Integer, r_strat As Integer, r_ac As Integer, r_ins As Integer, r_init_depo As Integer
    Dim r_fin_depo As Integer, r_s_return As Integer, r_net_pc As Integer, r_net_ac As Integer, r_mon_won As Integer, r_mon_lost As Integer, r_ann_ret As Integer, r_mn_ret As Integer, r_s_pips As Integer, r_net_pp As Integer
    Dim r_won_pp As Integer, r_lost_pp As Integer, r_per_yr_pp As Integer, r_per_mn_pp As Integer, r_per_w_pp As Integer, r_s_rsq As Integer, r_rsq_tr_cve As Integer, r_rsq_eq_cve As Integer, r_s_pf As Integer, r_pf_ac As Integer
    Dim r_pf_pp As Integer, r_s_rf As Integer, r_rf_ac As Integer, r_rf_pp As Integer, r_s_avgs_pp As Integer, r_avg_td_pp As Integer, r_avg_win_pp As Integer, r_avg_los_pp As Integer, r_avg_win_los_pp As Integer, r_s_avgs_ac As Integer
    Dim r_avg_td_ac As Integer, r_avg_win_ac As Integer, r_avg_los_ac As Integer, r_avg_win_los_ac As Integer, r_s_intvl As Integer, r_mn_win As Integer, r_mn_los As Integer, r_mn_no_tds As Integer, r_mn_win_los As Integer, r_w_win As Integer
    Dim r_w_los As Integer, r_w_no_tds As Integer, r_w_win_los As Integer, r_d_win As Integer, r_d_los As Integer, r_d_no_tds As Integer, r_d_win_los As Integer, r_s_act_intvl As Integer, r_mn_act As Integer, r_mn_act_all As Integer
    Dim r_w_act As Integer, r_w_act_all As Integer, r_d_act As Integer, r_d_act_all As Integer
    
    Dim r_s_std As Integer, r_std_tds_pp As Integer
    Dim r_std_tds_ac As Integer
    
    Dim r_s_time As Integer, r_dt_begin As Integer, r_dt_end As Integer, r_yrs As Integer, r_mns As Integer, r_wks As Integer, r_cds As Integer
    Dim r_s_cmsn As Integer, r_cmsn_amnt_ac As Integer, r_cmsn_avg_per_d As Integer, r_s_mdd_ac As Integer, r_mdd_ec_ac As Integer, r_mfe_ec_ac As Integer, r_abs_hi_ac As Integer, r_abs_lo_ac As Integer, r_s_mdd_pp As Integer, r_mdd_ec_pp As Integer
    Dim r_mfe_ec_pp As Integer, r_abs_hi_pp As Integer, r_abs_lo_pp As Integer, r_s_trades As Integer, r_tds_closed As Integer, r_tds_per_yr As Integer, r_tds_per_mn As Integer, r_tds_per_w As Integer, r_tds_max_per_d As Integer, r_tds_win_count As Integer
    Dim r_tds_los_count As Integer, r_tds_win_pc As Integer, r_tds_lg As Integer, r_tds_sh As Integer, r_tds_lg_sh As Integer, r_tds_lg_win_pc As Integer, r_tds_sh_win_pc As Integer, r_s_dur As Integer, r_avg_dur As Integer, r_avg_win_dur As Integer
    Dim r_avg_los_dur As Integer, r_avg_dur_win_los As Integer, r_s_stks As Integer, r_stk_win_tds As Integer, r_stk_los_tds As Integer, r_stk_win_mns As Integer, r_stk_los_mns As Integer, r_stk_win_wks As Integer, r_stk_los_wks As Integer, r_stk_win_ds As Integer
    Dim r_stk_los_ds As Integer, r_runs_tds As Integer, r_zscore_tds As Integer, r_runs_wks As Integer, r_zscore_wks As Integer, r_s_over As Integer, r_over_amnt_pp As Integer, r_ds_over As Integer, r_dwo_per_mn As Integer, r_s_expo As Integer
    Dim r_tm_in_tds As Integer, r_tm_in_win_tds As Integer, r_tm_in_los_tds As Integer, r_tm_win_los As Integer, r_s_orders As Integer, r_ord_sent As Integer, r_ord_tds As Integer
    Dim split_row As Integer
'
    Dim open_fail As Boolean
    Dim open_from_cell As Boolean
    Dim no_trades_found As Boolean
'
' MAIN BOOK - book - sheet - new sheet - new sheet cells
    Dim mb As Workbook      ' macro's workbook
    Dim ms As Worksheet     ' 1st sheet
    Dim ns As Worksheet     ' new worksheet for REPORT
    Dim nc As Range         ' new worksheet CElls
' REPORT book - sheet - cells
    Dim rb As Workbook
    Dim rs As Worksheet
    Dim rc As Range
' INTEGER
    Dim ins_td_r As Integer
    Dim tl_r As Integer
'
    Const chFontSize As Integer = 10    ' chart title font size
' separator variables
    Dim current_decimal As String
    Dim undo_sep As Boolean, undo_usesyst As Boolean
Private Sub GSPR_Single_Extra()
    Dim MbAnswer As Variant
    
    On Error Resume Next
    
    If MsgBox("Works only for reports made with Platform version 1. CONTINUE?", _
        vbYesNo) <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Creating report ""Extra""."
    Call GSPR_Separator_Auto_Switcher_Single_Extra
    Call GSPR_Open_Report_f
    If open_fail = True Then
        Call GSPR_Closing_Extra
        Exit Sub
    End If
    Call GSPR_Prep_SV
    Call GSPR_Get_Basic_Data
    If no_trades_found = True Then
        Call GSPR_Closing_Extra
        Exit Sub
    End If
    Call GSPR_Fill_Trade_Log
    Call GSPR_Fill_Parameters
    Call GSPR_Fill_Intervals_Logs
    Call GSPR_Fill_Aux_Logs
    Call GSPR_Get_All_Stats_part_1
    Call GSPR_Get_All_Stats_part_2
    Call GSPR_Prep_SN
    Call GSPR_Prep_all_fm
    Call GSPR_Report_Status
    Call GSPR_Show_SN_SV
    Call GSPR_Format_SV_SN
    Call GSPR_Fill_Logs_Heads
    Call GSPR_Show_All_Logs
    Call GSPR_Build_Charts
    Call GSPR_Separator_Undo_Single_Extra
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
Private Sub GSPR_Separator_Auto_Switcher_Single_Extra()
    undo_sep = False
    undo_usesyst = False
    If Application.UseSystemSeparators Then     ' SYS - ON
        If Not Application.International(xlDecimalSeparator) = "." Then
            Application.UseSystemSeparators = False
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            undo_sep = True                     ' undo condition 1
            undo_usesyst = True                 ' undo condition 2
        End If
    Else                                        ' SYS - OFF
        If Not Application.DecimalSeparator = "." Then
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            undo_sep = True                     ' undo condition 1
            undo_usesyst = False                ' undo condition 2
        End If
    End If
End Sub
Private Sub GSPR_Separator_Undo_Single_Extra()
    If undo_sep Then
        Application.DecimalSeparator = current_decimal
        If undo_usesyst Then
            Application.UseSystemSeparators = True
        End If
    End If
End Sub
Private Sub GSPR_Closing_Extra()
    Call GSPR_Separator_Undo_Single_Extra
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
Private Sub GSPR_Open_Report_f()
    Dim y As Integer
    Dim adr_row As Integer
    Dim hs As String
    Const ctrl_str As String = "file:///"
'
    Set mb = ActiveWorkbook
    Set ms = ActiveSheet
    If Not ms.Cells.Find(what:=rep_type) Is Nothing Then
        adr_row = ActiveSheet.Cells.Find(what:=rep_type).Row
        hs = Cells(adr_row - 1, 2).Hyperlinks(1).Address
        If InStr(1, hs, ":") = 0 Then   ' if file on Desktop
            hs = Environ("UserProfile") & "\Desktop\" & hs
        End If
        y = InStr(1, hs, ":\") - 2
        hs = Right(hs, Len(hs) - y)
        rep_adr = ctrl_str & hs
    Else
        rep_adr = ActiveCell.Value
    End If
    open_fail = False
' error check
' 1. empty cell
    If rep_adr = "" Then
        open_fail = True
        MsgBox "Wrong address format. Please copy from browser.", 48, "GetStats"
        Exit Sub
' 2. wrong address
    ElseIf Left(rep_adr, 8) <> ctrl_str Then
        open_fail = True
        MsgBox "Wrong address format. Please copy from browser.", 48, "GetStats"
        Exit Sub
    ElseIf Right(rep_adr, 5) <> ".html" Then
        open_fail = True
        MsgBox "Wrong address format. Please copy from browser.", 48, "GetStats"
        Exit Sub
    End If
' encoding %20
    If InStr(1, rep_adr, "%20", 1) Then
        rep_adr = Replace(rep_adr, "%20", " ", 1, 1, 1)
    End If
' remove "file:///"
    rep_adr = Right(rep_adr, Len(rep_adr) - 8)
' 3. check if file exists
    If Dir(rep_adr) = "" Then
        MsgBox "Wrong address format. Please copy from browser.", 48, "GetStats"
        open_fail = True
        Exit Sub
    End If
' OPEN FILE
    Set rb = Workbooks.Open(rep_adr)
    Set rs = rb.Sheets(1)
    Set rc = rs.Cells                  ' report sheet cells
End Sub
Private Sub GSPR_Prep_SV()
' ROWS
' Report
    r_s_report = 1
    r_name = r_s_report + 1
'    r_version = r_name + 1
    r_type = r_name + 1
    r_date_gen = r_type + 1
    r_time_gen = r_date_gen + 1
    r_file = r_time_gen + 1
' Basic
    r_s_basic = r_file + 1
    r_strat = r_s_basic + 1
    r_ac = r_strat + 1
    r_ins = r_ac + 1
    r_init_depo = r_ins + 1
    r_fin_depo = r_init_depo + 1
' Return
    r_s_return = r_fin_depo + 1
    r_net_pc = r_s_return + 1
    r_net_ac = r_net_pc + 1
    r_mon_won = r_net_ac + 1
    r_mon_lost = r_mon_won + 1
    r_ann_ret = r_mon_lost + 1
    r_mn_ret = r_ann_ret + 1
' Pips
    r_s_pips = r_mn_ret + 1
    r_net_pp = r_s_pips + 1
    r_won_pp = r_net_pp + 1
    r_lost_pp = r_won_pp + 1
    r_per_yr_pp = r_lost_pp + 1
    r_per_mn_pp = r_per_yr_pp + 1
    r_per_w_pp = r_per_mn_pp + 1
' R-sq
    r_s_rsq = r_per_w_pp + 1
    r_rsq_tr_cve = r_s_rsq + 1
    r_rsq_eq_cve = r_rsq_tr_cve + 1
' Profit factor
    r_s_pf = r_rsq_eq_cve + 1
    r_pf_ac = r_s_pf + 1
    r_pf_pp = r_pf_ac + 1
' Recovery factor
    r_s_rf = r_pf_pp + 1
    r_rf_ac = r_s_rf + 1
    r_rf_pp = r_rf_ac + 1
' Averages in pips
    r_s_avgs_pp = r_rf_pp + 1
    r_avg_td_pp = r_s_avgs_pp + 1
    r_avg_win_pp = r_avg_td_pp + 1
    r_avg_los_pp = r_avg_win_pp + 1
    r_avg_win_los_pp = r_avg_los_pp + 1
' Averages in a/c
    r_s_avgs_ac = r_avg_win_los_pp + 1
    r_avg_td_ac = r_s_avgs_ac + 1
    r_avg_win_ac = r_avg_td_ac + 1
    r_avg_los_ac = r_avg_win_ac + 1
    r_avg_win_los_ac = r_avg_los_ac + 1
' Invervals w/l
    r_s_intvl = r_avg_win_los_ac + 1
    r_mn_win = r_s_intvl + 1
    r_mn_los = r_mn_win + 1
    r_mn_no_tds = r_mn_los + 1
    r_mn_win_los = r_mn_no_tds + 1
    r_w_win = r_mn_win_los + 1
    r_w_los = r_w_win + 1
    r_w_no_tds = r_w_los + 1
    r_w_win_los = r_w_no_tds + 1
    r_d_win = r_w_win_los + 1
    r_d_los = r_d_win + 1
    r_d_no_tds = r_d_los + 1
    r_d_win_los = r_d_no_tds + 1
' Active intervals
    r_s_act_intvl = r_d_win_los + 1
    r_mn_act = r_s_act_intvl + 1
    r_mn_act_all = r_mn_act + 1
    r_w_act = r_mn_act_all + 1
    r_w_act_all = r_w_act + 1
    r_d_act = r_w_act_all + 1
    r_d_act_all = r_d_act + 1
' Standard deviations
    r_s_std = r_d_act_all + 1
    r_std_tds_pp = r_s_std + 1
    r_std_tds_ac = r_std_tds_pp + 1
' ====== NEW COLUMN ====== STATISTICS
    split_row = r_std_tds_ac
' ====== NEW COLUMN ====== STATISTICS
' Time
    r_s_time = r_std_tds_ac + 1
    r_dt_begin = r_s_time + 1
    r_dt_end = r_dt_begin + 1
    r_yrs = r_dt_end + 1
    r_mns = r_yrs + 1
    r_wks = r_mns + 1
    r_cds = r_wks + 1
' Commissions
    r_s_cmsn = r_cds + 1
    r_cmsn_amnt_ac = r_s_cmsn + 1
    r_cmsn_avg_per_d = r_cmsn_amnt_ac + 1
' MDD, MFE in a/c
    r_s_mdd_ac = r_cmsn_avg_per_d + 1
    r_mdd_ec_ac = r_s_mdd_ac + 1
    r_mfe_ec_ac = r_mdd_ec_ac + 1
    r_abs_hi_ac = r_mfe_ec_ac + 1
    r_abs_lo_ac = r_abs_hi_ac + 1
' MDD, MFE in pips
    r_s_mdd_pp = r_abs_lo_ac + 1
    r_mdd_ec_pp = r_s_mdd_pp + 1
    r_mfe_ec_pp = r_mdd_ec_pp + 1
    r_abs_hi_pp = r_mfe_ec_pp + 1
    r_abs_lo_pp = r_abs_hi_pp + 1
' Trades
    r_s_trades = r_abs_lo_pp + 1
    r_tds_closed = r_s_trades + 1
    r_tds_per_yr = r_tds_closed + 1
    r_tds_per_mn = r_tds_per_yr + 1
    r_tds_per_w = r_tds_per_mn + 1
    r_tds_max_per_d = r_tds_per_w + 1
    r_tds_win_count = r_tds_max_per_d + 1
    r_tds_los_count = r_tds_win_count + 1
    r_tds_win_pc = r_tds_los_count + 1
    r_tds_lg = r_tds_win_pc + 1
    r_tds_sh = r_tds_lg + 1
    r_tds_lg_sh = r_tds_sh + 1
    r_tds_lg_win_pc = r_tds_lg_sh + 1
    r_tds_sh_win_pc = r_tds_lg_win_pc + 1
' Trade duration
    r_s_dur = r_tds_sh_win_pc + 1
    r_avg_dur = r_s_dur + 1
    r_avg_win_dur = r_avg_dur + 1
    r_avg_los_dur = r_avg_win_dur + 1
    r_avg_dur_win_los = r_avg_los_dur + 1
' Streaks
    r_s_stks = r_avg_dur_win_los + 1
    r_stk_win_tds = r_s_stks + 1
    r_stk_los_tds = r_stk_win_tds + 1
    r_stk_win_mns = r_stk_los_tds + 1
    r_stk_los_mns = r_stk_win_mns + 1
    r_stk_win_wks = r_stk_los_mns + 1
    r_stk_los_wks = r_stk_win_wks + 1
    r_stk_win_ds = r_stk_los_wks + 1
    r_stk_los_ds = r_stk_win_ds + 1
    r_runs_tds = r_stk_los_ds + 1
    r_zscore_tds = r_runs_tds + 1
    r_runs_wks = r_zscore_tds + 1
    r_zscore_wks = r_runs_wks + 1
' Overnights
    r_s_over = r_zscore_wks + 1
    r_over_amnt_pp = r_s_over + 1
    r_ds_over = r_over_amnt_pp + 1
    r_dwo_per_mn = r_ds_over + 1
' Exposition
    r_s_expo = r_dwo_per_mn + 1
    r_tm_in_tds = r_s_expo + 1
    r_tm_in_win_tds = r_tm_in_tds + 1
    r_tm_in_los_tds = r_tm_in_win_tds + 1
    r_tm_win_los = r_tm_in_los_tds + 1
' Orders
    r_s_orders = r_tm_win_los + 1
    r_ord_sent = r_s_orders + 1
    r_ord_tds = r_ord_sent + 1
'sv_sn_ubound
    sv_sn_ubound = r_ord_tds
    ReDim SV(sv_sn_ubound)
    ReDim SN(sv_sn_ubound)
End Sub
Private Sub GSPR_Get_Basic_Data()
    Dim bI As Integer, cI As Integer, dI As Integer
    Dim used_inss As Integer
    Dim s As String, ch As String
    ' instrument with trades, row
    no_trades_found = False
    Set rc = rs.Cells                  ' report sheet cells
' file
    SV(r_file) = rep_adr
' get strategy name
    s = rc(3, 1).Value
    bI = InStr(1, s, " strategy report for", 1)
    cI = InStr(bI, s, " instrument(s) from", 1)
' strat_name
    SV(r_strat) = Left(s, bI - 1)
' calculate number of used instruments
    used_inss = 0
    For dI = bI To cI
        ch = Mid(s, dI, 1)
        If ch = "," Then
            used_inss = used_inss + 1
        End If
    Next dI
    used_inss = used_inss + 1
' find relevant instrument, with trades
    ins_td_r = 10
    For bI = 1 To used_inss
        ins_td_r = rc.Find(what:="Closed positions", after:=rc(ins_td_r, 1), LookIn:=xlValues, LookAt _
            :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
            False, searchformat:=False).Row
        If rc(ins_td_r, 2) <> 0 Then
' Exit, because no instruments with trades > 0
            Exit For
        End If
    Next bI
    If rc(ins_td_r, 2) = 0 Then
        no_trades_found = True
        MsgBox "ERROR. No instruments with closed trades found."
        Exit Sub
    End If
' td_closed
    SV(r_tds_closed) = rc(ins_td_r, 2)
' instrument
    s = rc(ins_td_r - 9, 1)
    bI = InStr(1, s, " ", 1)
    s = Right(s, Len(s) - bI)
    SV(r_ins) = s
' orders_sent
    SV(r_ord_sent) = rc(ins_td_r + 1, 2)
' ord_tds
    SV(r_ord_tds) = SV(r_ord_sent) / SV(r_tds_closed)
' cmsn
    SV(r_cmsn_amnt_ac) = rc(ins_td_r + 5, 2)
'' date_begin
'    SV(r_dt_begin) = rc(ins_td_r - 7, 2)
'' date_end
'    SV(r_dt_end) = rc(ins_td_r - 4, 2)
' Test begin
    SV(r_dt_begin) = CDate(rc(ins_td_r - 7, 2))
' Test end
'    SV(s_date_end, 2) = Int(rc(ins_td_r - 4, 2))
    SV(r_dt_end) = CDate(rc(ins_td_r - 4, 2)) - 1
' ac
    SV(r_ac) = rc(4, 2)
' depo_initial
    SV(r_init_depo) = CDbl(rc(5, 2))
'' depo_finish 'a
'    SV(r_fin_depo) = CDbl(rsc(6, 2))
' cds
    SV(r_cds) = WorksheetFunction.RoundUp(SV(r_dt_end) - SV(r_dt_begin), 0)
' yrs
    SV(r_yrs) = SV(r_cds) / 365
' mns
    SV(r_mns) = SV(r_yrs) * 12
' wks
    SV(r_wks) = SV(r_cds) / 7
' cmsn_per_d
    SV(r_cmsn_avg_per_d) = SV(r_cmsn_amnt_ac) / SV(r_cds)
' tds_per_yr
    SV(r_tds_per_yr) = SV(r_tds_closed) / SV(r_yrs)
' tds_per_mn
    SV(r_tds_per_mn) = SV(r_tds_closed) / SV(r_mns)
' tds_per_w
    SV(r_tds_per_w) = SV(r_tds_closed) / SV(r_wks)
End Sub
Private Sub GSPR_Fill_Trade_Log()
    Dim cI As Integer, rI As Integer
    Dim av_c As Double
' get trade log first row - header
    tl_r = rc.Find(what:="Closed orders:", after:=rc(ins_td_r, 1), LookIn:=xlValues, LookAt _
        :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
        False, searchformat:=False).Row + 2
    ReDim Tlog(1 To SV(r_tds_closed), 1 To 11)
    ' 1. amount
    ' 2. direction
    ' 3. open price
    ' 4. close price
    ' 5. profit/loss
    ' 6. profit/loss in pips
    ' 7. open date
    ' 8. close date
    ' 9. comment
    ' 10. equity curve
    ' 11. pips sum
' Tlog() - trades only
    For rI = LBound(Tlog, 1) To UBound(Tlog, 1)
        For cI = LBound(Tlog, 2) To UBound(Tlog, 2) - 2
            Tlog(rI, cI) = rc(tl_r + rI, cI + 1)
        Next cI
    Next rI
' vars to update absolute high and low
    SV(r_abs_hi_ac) = SV(r_init_depo)
    SV(r_abs_lo_ac) = SV(r_init_depo)
    SV(r_abs_hi_pp) = 0
    SV(r_abs_lo_pp) = 0
' Tlog equity curve
    av_c = SV(r_cmsn_amnt_ac) / SV(r_tds_closed)
    Tlog(1, 10) = Tlog(1, 5) + SV(r_init_depo) - av_c
    For rI = 2 To UBound(Tlog)
        Tlog(rI, 10) = Tlog(rI - 1, 10) + Tlog(rI, 5) - av_c
' Update absolute high
        ' a/c
        If Tlog(rI, 10) > SV(r_abs_hi_ac) Then
            SV(r_abs_hi_ac) = Tlog(rI, 10)
        End If
' Update absolute low
        ' a/c
        If Tlog(rI, 10) < SV(r_abs_lo_ac) Then
            SV(r_abs_lo_ac) = Tlog(rI, 10)
        End If
    Next rI
' Tlog pips curve
    Tlog(1, 11) = Tlog(1, 6)
    For rI = 2 To UBound(Tlog)
        Tlog(rI, 11) = Tlog(rI - 1, 11) + Tlog(rI, 6)
' Update absolute high
        ' pips
        If Tlog(rI, 11) > SV(r_abs_hi_pp) Then
            SV(r_abs_hi_pp) = Tlog(rI, 11)
        End If
' Update absolute low
        ' pips
        If Tlog(rI, 11) < SV(r_abs_lo_pp) Then
            SV(r_abs_lo_pp) = Tlog(rI, 11)
        End If
    Next rI
'
' Tlog_head() - header
    For cI = LBound(Tlog_head) To UBound(Tlog_head) - 2
        Tlog_head(cI) = rc(tl_r, cI + 1)
    Next cI
    Tlog_head(10) = "Equity curve"
    Tlog_head(11) = "Pips sum"
' trade amount to CDbl
    For rI = LBound(Tlog, 1) To UBound(Tlog, 1)
        Tlog(rI, 1) = CDbl(Tlog(rI, 1))
    Next rI
End Sub
Private Sub GSPR_Fill_Parameters()
    Dim p_fr As Integer, p_lr As Integer, p_tot As Integer
    Dim rI As Integer, cI As Integer
' get parameters first & last row
    p_fr = 12
    p_lr = rc(12, 1).End(xlDown).Row
    p_tot = p_lr - p_fr + 1
    ReDim Par(1 To p_tot, 1 To 2)
' fill array
    For rI = 1 To p_tot
        For cI = 1 To 2
            Par(rI, cI) = rc(p_fr - 1 + rI, cI)
        Next cI
    Next rI
End Sub
Private Sub GSPR_Fill_Intervals_Logs()
    Dim cI As Integer, rI As Integer, rol_d As Integer
    Dim i_retc As Double, i_hic As Double, i_loc As Double, i_clc As Double
    Dim i_retp As Double, i_hip As Double, i_lop As Double, i_clp As Double
    Dim i_com As Double, i_sw As Double
    Dim tdt As Long         ' temporary date
    Dim oc_fr As Integer, oc_lr As Long, rL As Long
    Dim aI As Integer
    Dim ro_d As Long
    Dim s As String
'
    ReDim Dlog(0 To WorksheetFunction.RoundUp(SV(r_cds), 0), 1 To logs_ubd)
    ReDim Wlog(0 To WorksheetFunction.RoundUp(SV(r_wks), 0), 1 To logs_ubd)
    ReDim Mlog(0 To WorksheetFunction.RoundUp(SV(r_mns), 0), 1 To logs_ubd)
' ------------------------------------------
' DAILY LOG
' fill zero row
    Dlog(0, 1) = SV(r_dt_begin) - 1     ' previous day end
    Dlog(0, 2) = 0                      ' account currency
    Dlog(0, 3) = SV(r_init_depo)
    Dlog(0, 4) = SV(r_init_depo)
    Dlog(0, 5) = SV(r_init_depo)
    Dlog(0, 6) = SV(r_init_depo)
'    Dlog(0, 7) = 0                      ' swaps skip
'    Dlog(0, 8) = 0                      ' commissions skip
    Dlog(0, 9) = 0                      ' in pips
    Dlog(0, 10) = 0
    Dlog(0, 11) = 0
    Dlog(0, 12) = 0
    Dlog(0, 13) = 0
' fill dates
    For rI = 1 To UBound(Dlog, 1)
        Dlog(rI, 1) = Dlog(rI - 1, 1) + 1
    Next rI
' fill Overnights & Commissions
    oc_fr = rc.Find(what:="Event log:", after:=rc(ins_td_r + SV(r_tds_closed), 1), LookIn:=xlValues, LookAt _
        :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
        False, searchformat:=False).Row + 2
    oc_lr = rc(oc_fr, 1).End(xlDown).Row
'
    ro_d = 1
    For rI = 1 To UBound(Dlog, 1)
' compare dates
        If Dlog(rI, 1) = Int(CDate(rc(oc_fr + ro_d, 1))) Then   ' *! cdate
            Do While Dlog(rI, 1) = Int(CDate(rc(oc_fr + ro_d, 1)))  ' *! cdate
                Select Case rc(oc_fr + ro_d, 2)
                    Case Is = "Commissions"
                        s = rc(oc_fr + ro_d, 3)
                        s = Right(s, Len(s) - 13)
                        s = Left(s, Len(s) - 1)
                        s = Replace(s, ".", ",", 1, 1, 1)
                        If CDbl(s) <> 0 Then
                            Dlog(rI, 8) = CDbl(s)
                        End If
                    Case Is = "Overnights"
                        s = rc(oc_fr + ro_d, 3)
                        s = Right(s, Len(s) - 22)
                        aI = InStr(1, s, "]", 1)
                        s = Left(s, aI - 1)
                        s = Replace(s, ".", ",", 1, 1, 1)
                        Dlog(rI, 7) = Dlog(rI, 7) + CDbl(s)
                    End Select
                ' go to next trade in Trade log
                If ro_d + oc_fr < oc_lr Then
                    ro_d = ro_d + 1
                Else
' or exit do, when after last row processed
                    Exit Do
                End If
            Loop
        End If
    Next rI
' fill OHLC
    rol_d = 1 '8 col in Tlog(i_ret, 8)
    For rI = 1 To UBound(Dlog, 1)
        i_retc = 0                          ' assign zero return, a/c - commissions
        i_retp = 0                          ' assign zero return, pips
' account currency
        Dlog(rI, 3) = Dlog(rI - 1, 6)       ' open price, a/c = prev close
        i_hic = Dlog(rI, 3)                 ' assign start hi, a/c = open
        i_loc = Dlog(rI, 3) + i_retc        ' assign start lo, a/c = open
        i_clc = Dlog(rI, 3) + i_retc        ' assign close, a/c = open
' pips
        Dlog(rI, 10) = Dlog(rI - 1, 13)     ' open price, pips = prev close
        i_hip = Dlog(rI, 10)                ' assign start hi, pips = open
        i_lop = Dlog(rI, 10)                ' assign start lo, pips = open
        i_clp = Dlog(rI, 10)                ' assign close, pips = open
'
        If Dlog(rI, 1) = Int(CDate(Tlog(rol_d, 8))) Then
            Do While Dlog(rI, 1) = Int(CDate(Tlog(rol_d, 8)))
                i_retc = i_retc + Tlog(rol_d, 5)    ' update return, a/c
                i_retp = i_retp + Tlog(rol_d, 6)    ' update return, pips
' High & Low, a/c
                i_clc = Dlog(rI, 3) + i_retc
                Select Case i_clc
                    Case Is > i_hic         ' update high
                        i_hic = i_clc
                    Case Is < i_loc         ' update low
                        i_loc = i_clc
                End Select
' High & Low, pips
                i_clp = Dlog(rI, 10) + i_retp
                Select Case i_clp
                    Case Is > i_hip         ' update high
                        i_hip = i_clp
                    Case Is < i_lop         ' update low
                        i_lop = i_clp
                End Select
' go to next trade in Trade log
                If rol_d < SV(r_tds_closed) Then
                    rol_d = rol_d + 1
                Else
' or exit do, when after last trade processed
                    Exit Do
                End If
            Loop
        End If
' ACCOUNT CURRENCY
        Dlog(rI, 2) = i_retc                ' return = 0
        Dlog(rI, 4) = i_hic                 ' hi = open
        Dlog(rI, 5) = i_loc                 ' lo = open
        Dlog(rI, 6) = i_clc - Dlog(rI, 8)   ' close = open - cmsn
' update day's low after deducting commissions
        If Dlog(rI, 6) < Dlog(rI, 5) Then
            Dlog(rI, 5) = Dlog(rI, 6)
        End If
' PIPS
        Dlog(rI, 9) = i_retp                ' return = 0
        Dlog(rI, 11) = i_hip                ' hi = open
        Dlog(rI, 12) = i_lop                ' lo = open
        Dlog(rI, 13) = i_clp                ' close = open
    Next rI
' Close HTML report
    rb.Close savechanges:=False
' ------------------------------------------
' WEEKLY LOG
' fill zero row
    tdt = Int(SV(r_dt_begin))
    Do Until Weekday(tdt) = 7       ' find previous Saturday
        tdt = tdt - 1
    Loop
    Wlog(0, 1) = CDate(tdt)         ' previous Saturday
    Wlog(0, 2) = 0                  ' account currency
    Wlog(0, 3) = SV(r_init_depo)
    Wlog(0, 4) = SV(r_init_depo)
    Wlog(0, 5) = SV(r_init_depo)
    Wlog(0, 6) = SV(r_init_depo)
    Wlog(0, 9) = 0                  ' pips
    Wlog(0, 10) = 0
    Wlog(0, 11) = 0
    Wlog(0, 12) = 0
    Wlog(0, 13) = 0
' fill dates
    For rI = 1 To UBound(Wlog, 1)
        Wlog(rI, 1) = Wlog(rI - 1, 1) + 7
    Next rI
' fill Weekly ohlc and commissions - in a/c
    rol_d = 1 '8 col in Tlog(i_ret, 8)
    For rI = 1 To UBound(Wlog, 1)
        i_retc = 0                          ' assign zero return, a/c
        i_retp = 0                          ' assign zero return, pips
        i_sw = 0                            ' assign zero to swaps
        i_com = 0                           ' assign zero to commissions
' account currency
        Wlog(rI, 3) = Wlog(rI - 1, 6)       ' open price, a/c = prev close
        i_hic = Wlog(rI, 3)                 ' assign start hi, a/c = open
        i_loc = Wlog(rI, 3)                 ' assign start lo, a/c = open
        i_clc = Wlog(rI, 3)                 ' assign close, a/c = open
' pips
        Wlog(rI, 10) = Wlog(rI - 1, 13)     ' open price, pips = prev close
        i_hip = Wlog(rI, 10)                ' assign start hi, pips = open
        i_lop = Wlog(rI, 10)                ' assign start lo, pips = open
        i_clp = Wlog(rI, 10)                ' assign close, pips = open
'
        If Dlog(rol_d, 1) <= Wlog(rI, 1) And Dlog(rol_d, 1) > Wlog(rI - 1, 1) Then
            Do While Dlog(rol_d, 1) <= Wlog(rI, 1) And _
                    Dlog(rol_d, 1) > Wlog(rI - 1, 1)
                i_retc = i_retc + Dlog(rol_d, 2) - Dlog(rol_d, 8) ' update return, a/c
                i_retp = i_retp + Dlog(rol_d, 9)    ' update return, pips
                i_sw = i_sw + Dlog(rol_d, 7)
                i_com = i_com + Dlog(rol_d, 8)
' High & Low, a/c
                i_clc = Wlog(rI, 3) + i_retc
                Select Case i_clc
                    Case Is > i_hic         ' update high
                        i_hic = i_clc
                    Case Is < i_loc         ' update low
                        i_loc = i_clc
                End Select
' High & Low, pips
                i_clp = Wlog(rI, 10) + i_retp
                Select Case i_clp
                    Case Is > i_hip         ' update high
                        i_hip = i_clp
                    Case Is < i_lop         ' update low
                        i_lop = i_clp
                End Select
' go to next day in day log
'                rol_d = rol_d + 1
                If rol_d < UBound(Dlog, 1) Then
                    rol_d = rol_d + 1
                Else
' or exit do, when after last day processed
                    Exit Do
                End If
            Loop
        End If
' ACCOUNT CURRENCY
        Wlog(rI, 2) = i_retc                ' return = 0
        Wlog(rI, 4) = i_hic                 ' hi = open
        Wlog(rI, 5) = i_loc                 ' lo = open
        Wlog(rI, 6) = i_clc                 ' close = open - cmsn
        If i_sw <> 0 Then
            Wlog(rI, 7) = i_sw
        End If
        If i_com <> 0 Then
            Wlog(rI, 8) = i_com
        End If
' PIPS
        Wlog(rI, 9) = i_retp                ' return = 0
        Wlog(rI, 11) = i_hip                ' hi = open
        Wlog(rI, 12) = i_lop                ' lo = open
        Wlog(rI, 13) = i_clp                ' close = open
    Next rI
' ------------------------------------------
' MONTHLY LOG
' fill zero row
    tdt = Int(SV(r_dt_begin))
    tdt = tdt - Day(tdt)            ' find last day of previous month
    Mlog(0, 1) = CDate(tdt)         ' previous month end
    Mlog(0, 2) = 0                  ' account currency
    Mlog(0, 3) = SV(r_init_depo)
    Mlog(0, 4) = SV(r_init_depo)
    Mlog(0, 5) = SV(r_init_depo)
    Mlog(0, 6) = SV(r_init_depo)
    Mlog(0, 9) = 0                  ' pips
    Mlog(0, 10) = 0
    Mlog(0, 11) = 0
    Mlog(0, 12) = 0
    Mlog(0, 13) = 0
' fill dates
    For rI = 1 To UBound(Mlog, 1)
        Mlog(rI, 1) = DateAdd("m", 1, Mlog(rI - 1, 1) + 1) - 1
    Next rI
' fill Monthly ohlc and commissions - in a/c
    rol_d = 1 '8 col in Tlog(i_ret, 8)
    For rI = 1 To UBound(Mlog, 1)
        i_retc = 0                          ' assign zero return, a/c
        i_retp = 0                          ' assign zero return, pips
        i_sw = 0                            ' assign zero to swaps
        i_com = 0                           ' assign zero to commissions
' account currency
        Mlog(rI, 3) = Mlog(rI - 1, 6)       ' open price, a/c = prev close
        i_hic = Mlog(rI, 3)                 ' assign start hi, a/c = open
        i_loc = Mlog(rI, 3)                 ' assign start lo, a/c = open
        i_clc = Mlog(rI, 3)                 ' assign close, a/c = open
' pips
        Mlog(rI, 10) = Mlog(rI - 1, 13)     ' open price, pips = prev close
        i_hip = Mlog(rI, 10)                ' assign start hi, pips = open
        i_lop = Mlog(rI, 10)                ' assign start lo, pips = open
        i_clp = Mlog(rI, 10)                ' assign close, pips = open
'
        If Dlog(rol_d, 1) <= Mlog(rI, 1) And Dlog(rol_d, 1) > Mlog(rI - 1, 1) Then
            Do While Dlog(rol_d, 1) <= Mlog(rI, 1) And _
                    Dlog(rol_d, 1) > Mlog(rI - 1, 1)
                i_retc = i_retc + Dlog(rol_d, 2) - Dlog(rol_d, 8) ' update return, a/c
                i_retp = i_retp + Dlog(rol_d, 9)    ' update return, pips
                i_sw = i_sw + Dlog(rol_d, 7)
                i_com = i_com + Dlog(rol_d, 8)
' High & Low, a/c
                i_clc = Mlog(rI, 3) + i_retc
                Select Case i_clc
                    Case Is > i_hic         ' update high
                        i_hic = i_clc
                    Case Is < i_loc         ' update low
                        i_loc = i_clc
                End Select
' High & Low, pips
                i_clp = Mlog(rI, 10) + i_retp
                Select Case i_clp
                    Case Is > i_hip         ' update high
                        i_hip = i_clp
                    Case Is < i_lop         ' update low
                        i_lop = i_clp
                End Select
' go to next day in day log
'                rol_d = rol_d + 1
                If rol_d < UBound(Dlog, 1) Then
                    rol_d = rol_d + 1
                Else
' or exit do, when after last day processed
                    Exit Do
                End If
            Loop
        End If
' ACCOUNT CURRENCY
        Mlog(rI, 2) = i_retc                ' return = 0
        Mlog(rI, 4) = i_hic                 ' hi = open
        Mlog(rI, 5) = i_loc                 ' lo = open
        Mlog(rI, 6) = i_clc                 ' close = open - cmsn
        If i_sw <> 0 Then
            Mlog(rI, 7) = i_sw
        End If
        If i_com <> 0 Then
            Mlog(rI, 8) = i_com
        End If
' PIPS
        Mlog(rI, 9) = i_retp                ' return = 0
        Mlog(rI, 11) = i_hip                ' hi = open
        Mlog(rI, 12) = i_lop                ' lo = open
        Mlog(rI, 13) = i_clp                ' close = open
    Next rI
End Sub
Private Sub GSPR_Fill_Aux_Logs()
    Dim rI As Integer
    Dim sum_index As Long           ' for R-sq
    Dim mean_index As Double        ' for R-sq
    Dim sum_curve As Double
    Dim mean_curve As Double
    Dim std_num_ac As Double        ' for STD a/c, SUM of: (x - x_mean)^2
    Dim std_num_pp As Double        ' for STD pp, SUM of: (x - x_mean)^2
    Dim r_num As Double, r_den1_sum As Double
        Dim r_den2_sum As Double ' for R-sq
    Dim tr_days As Integer
' Auxiliary Trades
    ReDim Aux_T(0 To SV(r_tds_closed), 0 To 16)    ' 0 column = index
' 0 index
' 1 duration
' ------------- account currency
' 2 hwm
' 3 lwm
' 4 dd
' 5 fe
' 6. std_x_x_mean
'' ------------------pips
' 7. hwm
' 8. lwm
' 9. dd
' 10. fe
' 11. x_x_mean
' 12. r2_y_y_mean

' fill "0" row values
    Aux_T(0, 0) = 0                 ' index
' account currency
    Aux_T(0, 2) = SV(r_init_depo)   ' hwm
    Aux_T(0, 3) = SV(r_init_depo)   ' lwm
' pips
    Aux_T(0, 7) = 0                 ' hwm
    Aux_T(0, 8) = 0                 ' lwm
' streaks
    Aux_T(0, 14) = 2                ' winner_loser
    Aux_T(0, 15) = 0                ' win_streak
    Aux_T(0, 16) = 0                ' los_streak
' fill indexes
    sum_index = 0                               ' for R-sq
    For rI = 1 To UBound(Aux_T, 1)
        Aux_T(rI, 0) = Aux_T(rI - 1, 0) + 1
        sum_index = sum_index + Aux_T(rI, 0)    ' for R-sq
    Next rI
    mean_index = sum_index / SV(r_tds_closed)   ' for R-sq
' fill hwm & lwm, streaks, exposition
    ' assignments
    SV(r_stk_win_tds) = 0               ' streaks: max win trades
    SV(r_stk_los_tds) = 0               ' streaks: max los trades
    SV(r_tm_in_tds) = 0         ' exposition: time in trades
    SV(r_tm_in_win_tds) = 0     ' exposition: time in winners
    SV(r_tm_in_los_tds) = 0     ' exposition: time in losers
    SV(r_mdd_ec_ac) = 0                 ' MDD of EC, a/c
    SV(r_mfe_ec_ac) = 0                 ' MFE of EC, a/c
    SV(r_mdd_ec_pp) = 0         ' MDD of EC, pips
    SV(r_mfe_ec_pp) = 0         ' MFE of EC, pips
    For rI = 1 To UBound(Aux_T, 1)
' duration
        Aux_T(rI, 1) = CDate(Tlog(rI, 8)) - CDate(Tlog(rI, 7))  ' *! cdate x2
' increment exposition
        SV(r_tm_in_tds) = SV(r_tm_in_tds) + Aux_T(rI, 1)
' hwm
        Aux_T(rI, 2) = WorksheetFunction.Max(Aux_T(rI - 1, 2), Tlog(rI, 10))    ' a/c
        Aux_T(rI, 7) = WorksheetFunction.Max(Aux_T(rI - 1, 7), Tlog(rI, 11))    ' pips
' lwm
        Aux_T(rI, 3) = WorksheetFunction.Min(Aux_T(rI - 1, 3), Tlog(rI, 10))    ' a/c
        Aux_T(rI, 8) = WorksheetFunction.Min(Aux_T(rI - 1, 8), Tlog(rI, 11))    ' pip
' dd
        Aux_T(rI, 4) = (Aux_T(rI, 2) - Tlog(rI, 10)) / Tlog(rI, 10) ' a/c
        Aux_T(rI, 9) = Aux_T(rI, 7) - Tlog(rI, 11)                 ' pips
' update Max DD
        ' a/c
        If Aux_T(rI, 4) > SV(r_mdd_ec_ac) Then
            SV(r_mdd_ec_ac) = Aux_T(rI, 4)
        End If
        ' pips
        If Aux_T(rI, 9) > SV(r_mdd_ec_pp) Then
            SV(r_mdd_ec_pp) = Aux_T(rI, 9)
        End If
' fe ----------------------------------------------------------------------------
        Aux_T(rI, 5) = (Tlog(rI, 10) - Aux_T(rI, 3)) / Tlog(rI, 10) ' a/c
        Aux_T(rI, 10) = Tlog(rI, 11) - Aux_T(rI, 8)                 ' pips ==========
' update Max FE
        ' a/c
        If Aux_T(rI, 5) > SV(r_mfe_ec_ac) Then
            SV(r_mfe_ec_ac) = Aux_T(rI, 5)
        End If
        ' pips
        If Aux_T(rI, 11) > SV(r_mfe_ec_pp) Then '1================================
            SV(r_mfe_ec_pp) = Aux_T(rI, 10)
        End If
' streaks
        If Tlog(rI, 6) > 0 Then
' mark as winner
            Aux_T(rI, 14) = 1
' increment winning streak
            Aux_T(rI, 15) = Aux_T(rI - 1, 15) + 1
' recalculate max winning streak
            If Aux_T(rI, 15) > SV(r_stk_win_tds) Then
                SV(r_stk_win_tds) = Aux_T(rI, 15)
            End If
' reset losing streak to zero
            Aux_T(rI, 16) = 0
' increment exposition
            SV(r_tm_in_win_tds) = SV(r_tm_in_win_tds) + Aux_T(rI, 1)
        Else
' mark as loser
            Aux_T(rI, 14) = 0
' increment losing streak
            Aux_T(rI, 16) = Aux_T(rI - 1, 16) + 1
' recalculate max losing streak
            If Aux_T(rI, 16) > SV(r_stk_los_tds) Then
                SV(r_stk_los_tds) = Aux_T(rI, 16)
            End If
' reset winning streak to zero
            Aux_T(rI, 15) = 0
' increment exposition
            SV(r_tm_in_los_tds) = SV(r_tm_in_los_tds) + Aux_T(rI, 1)
        End If
        ' runs_count
        If Aux_T(rI, 14) <> Aux_T(rI - 1, 14) Then
            SV(r_runs_tds) = SV(r_runs_tds) + 1
        End If
    Next rI

' STD & R-squared ==============
' Net, a/c
    SV(r_net_ac) = Dlog(SV(r_cds), 6) - SV(r_init_depo)
' average trade in a/c
    SV(r_avg_td_ac) = SV(r_net_ac) / SV(r_tds_closed)
' Net, pips
    SV(r_net_pp) = Round(Dlog(SV(r_cds), 13), 2)
' average trade in pips
    SV(r_avg_td_pp) = SV(r_net_pp) / SV(r_tds_closed)
' mean in pip curve
    sum_curve = 0
    For rI = 1 To UBound(Tlog, 1)
        sum_curve = sum_curve + Tlog(rI, 11)
    Next rI
    mean_curve = sum_curve / SV(r_tds_closed)
' Assignments
    ' a/c
    std_num_ac = 0      ' STD calc: SUM of (x - x_mean)^2, rolling
    ' pips
    std_num_pp = 0      ' STD calc: SUM of (x - x_mean)^2, rolling
    r_num = 0            ' R2 calc: SUM of (x - x_mean) * (y - y_mean), rolling
    r_den1_sum = 0       ' R2 calc: SUM of (x - x_mean)^2, rolling
    r_den2_sum = 0       ' R2 calc: SUM of (y - y_mean)^2, rolling
    For rI = 1 To SV(r_tds_closed)
    ' STANDARD DEVIATION - a/c
        ' x - x_mean
        Aux_T(rI, 6) = Tlog(rI, 5) - SV(r_avg_td_ac)
        ' sum of (x - x_mean)
        std_num_ac = std_num_ac + Aux_T(rI, 6) ^ 2
    ' STANDARD DEVIATION - pips
        ' x - x_mean
        Aux_T(rI, 11) = Tlog(rI, 6) - SV(r_avg_td_pp)
        ' sum of (x - x_mean)
        std_num_pp = std_num_pp + Aux_T(rI, 11) ^ 2
    ' R-SQUARED - pips
        ' x - x_mean -> trade indexes
        Aux_T(rI, 12) = Aux_T(rI, 0) - mean_index
        ' y - y_mean -> equity curve, pips
        Aux_T(rI, 13) = Tlog(rI, 11) - mean_curve
        ' r numerator
        r_num = r_num + Aux_T(rI, 12) * Aux_T(rI, 13)
        ' r denominator
        r_den1_sum = r_den1_sum + Aux_T(rI, 12) ^ 2
        r_den2_sum = r_den2_sum + Aux_T(rI, 13) ^ 2
    Next rI
' STD, a/c
    SV(r_std_tds_ac) = Sqr(std_num_ac / (SV(r_tds_closed) - 1))
' STD, pips
    SV(r_std_tds_pp) = Sqr(std_num_pp / (SV(r_tds_closed) - 1))
' R2, pips
    SV(r_rsq_tr_cve) = (r_num / Sqr(r_den1_sum * r_den2_sum)) ^ 2

' DAILY LOG, AUX
    
    ReDim Aux_D(0 To SV(r_cds), 0 To 2)
' 0) index; 1) x - x_mean; 2) y - y_mean

' fill "0" row values
    Aux_D(0, 0) = 0                 ' index
'    Aux_D(0, 3) = 3 ' streaks       ' winner_loser
'    Aux_D(0, 4) = 0                 ' win_streak
'    Aux_D(0, 5) = 0                 ' los_streak
' fill indexes
    sum_index = 0                               ' for R-sq
    For rI = 1 To UBound(Aux_D, 1)
        Aux_D(rI, 0) = Aux_D(rI - 1, 0) + 1
        sum_index = sum_index + Aux_D(rI, 0)    ' for R-sq
    Next rI
    mean_index = sum_index / SV(r_cds)   ' for R-sq
' mean in pip curve
    sum_curve = 0
    For rI = 1 To UBound(Dlog, 1)
        sum_curve = sum_curve + Dlog(rI, 6)
    Next rI
    mean_curve = sum_curve / SV(r_cds)
' assignments: R2, numerator, denominator
    r_num = 0            ' R2 calc: SUM of (x - x_mean) * (y - y_mean), rolling
    r_den1_sum = 0       ' R2 calc: SUM of (x - x_mean)^2, rolling
    r_den2_sum = 0       ' R2 calc: SUM of (y - y_mean)^2, rolling
' assignments: streaks
    tr_days = 0
    SV(r_stk_win_ds) = 0               ' streaks: max win trades
    SV(r_stk_los_ds) = 0               ' streaks: max los trades
    
' streaks
    ReDim Aux_D_Sk_inv(1 To 2, 0 To 1)
' 1) winning streak; 2) losing streak
    Aux_D_Sk_inv(1, 0) = 0
    Aux_D_Sk_inv(2, 0) = 0
' Fill Aux_D & Aux_D_Sk_inv arrays
    For rI = 1 To UBound(Aux_D, 1)
' R-SQUARE - daily close
        ' x - x_mean -> trade indexes
        Aux_D(rI, 1) = Aux_D(rI, 0) - mean_index
        ' y - y_mean -> equity curve, daily close
        Aux_D(rI, 2) = Dlog(rI, 6) - mean_curve
        ' r numerator
        r_num = r_num + Aux_D(rI, 1) * Aux_D(rI, 2)
        ' r denominator
        r_den1_sum = r_den1_sum + Aux_D(rI, 1) ^ 2
        r_den2_sum = r_den2_sum + Aux_D(rI, 2) ^ 2
' count trading days - tr_days
        If Weekday(Dlog(rI, 1)) < 7 Then
            tr_days = tr_days + 1
            ReDim Preserve Aux_D_Sk_inv(1 To 2, 0 To tr_days)
' fill Aux_D_Sk_inv()
            Select Case Dlog(rI, 2)
                Case Is > 0
' increment winning streak
                    Aux_D_Sk_inv(1, tr_days) = Aux_D_Sk_inv(1, tr_days - 1) + 1
' recalculate max winning streak
                    If Aux_D_Sk_inv(1, tr_days) > SV(r_stk_win_ds) Then
                        SV(r_stk_win_ds) = Aux_D_Sk_inv(1, tr_days)
                    End If
' reset losing streak to 0
                    Aux_D_Sk_inv(2, tr_days) = 0
                Case Is = 0
' reset winning streak to 0
                    Aux_D_Sk_inv(1, tr_days) = 0
' reset losing streak to 0
                    Aux_D_Sk_inv(2, tr_days) = 0
                Case Is < 0
' increment losing streak
                    Aux_D_Sk_inv(2, tr_days) = Aux_D_Sk_inv(2, tr_days - 1) + 1
' recalculate max losing streak
                    If Aux_D_Sk_inv(2, tr_days) > SV(r_stk_los_ds) Then
                        SV(r_stk_los_ds) = Aux_D_Sk_inv(2, tr_days)
                    End If
' reset winning streak to 0
                    Aux_D_Sk_inv(1, tr_days) = 0
            End Select
        End If
    Next rI
' R2, daily close
    SV(r_rsq_eq_cve) = (r_num / Sqr(r_den1_sum * r_den2_sum)) ^ 2
    
' WEEKLY LOG, Aux

    ReDim Aux_W(0 To UBound(Wlog, 1), 1 To 3)
' 1) winner/loser; 2) winning streak; 3) losing streak
' fill Aux_W
    ' fill zeros
    Aux_W(0, 1) = 2
    Aux_W(0, 2) = 0
    Aux_W(0, 3) = 0
    ' assignments
    SV(r_stk_win_wks) = 0
    SV(r_stk_los_wks) = 0
    SV(r_runs_wks) = 0
    For rI = 1 To UBound(Aux_W, 1)
        Select Case Wlog(rI, 2)
            Case Is > 0
' mark as winner (for "runs")
                Aux_W(rI, 1) = 1
' increment winning streak
                Aux_W(rI, 2) = Aux_W(rI - 1, 2) + 1
' recalculate max winning streak
                If Aux_W(rI, 2) > SV(r_stk_win_wks) Then
                    SV(r_stk_win_wks) = Aux_W(rI, 2)
                End If
' reset losing streak to 0
                Aux_W(rI, 3) = 0
            Case Is = 0
' mark as loser (for "runs")
                Aux_W(rI, 1) = 0
' reset winning streak to 0
                Aux_W(rI, 2) = 0
' reset losing streak to 0
                Aux_W(rI, 3) = 0
            Case Is < 0
' mark as loser (for "runs")
                Aux_W(rI, 1) = 0
' increment losing streak
                Aux_W(rI, 3) = Aux_W(rI - 1, 3) + 1
' recalculate max losing streak
                If Aux_W(rI, 3) > SV(r_stk_los_wks) Then
                    SV(r_stk_los_wks) = Aux_W(rI, 3)
                End If
' reset winning streak to 0
                Aux_W(rI, 2) = 0
        End Select
' runs_count
        If Aux_W(rI, 1) <> Aux_W(rI - 1, 1) Then
            SV(r_runs_wks) = SV(r_runs_wks) + 1
        End If
    Next rI

' MONTHLY LOG, Aux
    ReDim Aux_M(0 To UBound(Mlog, 1), 1 To 2)
' 1) winning streak; 2) losing streak
' fill Aux_M
    ' fill zeros
    Aux_M(0, 1) = 0
    Aux_M(0, 2) = 0
    ' assignments
    SV(r_stk_win_mns) = 0
    SV(r_stk_los_mns) = 0
    For rI = 1 To UBound(Aux_M, 1)
        Select Case Mlog(rI, 2)
            Case Is > 0
' increment winning streak
                Aux_M(rI, 1) = Aux_M(rI - 1, 1) + 1
' recalculate max winning streak
                If Aux_M(rI, 1) > SV(r_stk_win_mns) Then
                    SV(r_stk_win_mns) = Aux_M(rI, 1)
                End If
' reset losing streak to 0
                Aux_W(rI, 2) = 0
            Case Is = 0
' reset winning streak to 0
                Aux_M(rI, 1) = 0
' reset losing streak to 0
                Aux_M(rI, 2) = 0
            Case Is < 0
' increment losing streak
                Aux_M(rI, 2) = Aux_M(rI - 1, 2) + 1
' recalculate max losing streak
                If Aux_M(rI, 2) > SV(r_stk_los_mns) Then
                    SV(r_stk_los_mns) = Aux_M(rI, 2)
                End If
' reset winning streak to 0
                Aux_W(rI, 1) = 0
        End Select
    Next rI
End Sub
Private Sub GSPR_Get_All_Stats_part_1()
    Dim rol_t As Integer
    Dim xI As Integer, yI As Integer, zI As Integer
' depo_finish
    SV(r_fin_depo) = Round(Dlog(SV(r_cds), 6), 2)
' net_pc
    SV(r_net_pc) = SV(r_net_ac) / SV(r_init_depo)
' ann_ret
    SV(r_ann_ret) = (1 + SV(r_net_pc)) ^ (365 / SV(r_cds)) - 1
' mn_ret
    SV(r_mn_ret) = (1 + SV(r_net_pc)) ^ (30.417 / SV(r_cds)) - 1
' money won & lost; pips won & lost; avg winning/losing trades in pips and a/c
    SV(r_mon_won) = 0   ' money won
    SV(r_mon_lost) = 0  ' money lost
    SV(r_won_pp) = 0    ' pips won
    SV(r_lost_pp) = 0   ' pips lost
    SV(r_tds_win_count) = 0 ' winning trades count
    SV(r_tds_los_count) = 0 ' losing trades count
    SV(r_tds_lg) = 0    ' longs count
    SV(r_tds_sh) = 0    ' shorts count
    yI = 0              ' longs winning
    zI = 0              ' shorts winning
    For xI = LBound(Tlog, 1) To UBound(Tlog, 1)
        If Tlog(xI, 5) > 0 Then     ' won
            SV(r_mon_won) = SV(r_mon_won) + Tlog(xI, 5)
            SV(r_won_pp) = SV(r_won_pp) + Tlog(xI, 6)
            SV(r_tds_win_count) = SV(r_tds_win_count) + 1
        Else                        ' lost
            SV(r_mon_lost) = SV(r_mon_lost) + Tlog(xI, 5)
            SV(r_lost_pp) = SV(r_lost_pp) + Tlog(xI, 6)
            SV(r_tds_los_count) = SV(r_tds_los_count) + 1
        End If
    ' longs, shorts
        If Tlog(xI, 2) = "BUY" Then
            SV(r_tds_lg) = SV(r_tds_lg) + 1
            If Tlog(xI, 6) > 0 Then     ' count winning longs
                yI = yI + 1
            End If
        Else
            SV(r_tds_sh) = SV(r_tds_sh) + 1
            If Tlog(xI, 6) > 0 Then     ' count winning shorts
                zI = zI + 1
            End If
        End If
    Next xI
' longs/shorts, winning longs & shorts
    ' longs/shorts
    If SV(r_tds_sh) = 0 Then
        SV(r_tds_lg_sh) = 0
    Else
        SV(r_tds_lg_sh) = SV(r_tds_lg) / SV(r_tds_sh)
    End If
    ' longs winning
    If SV(r_tds_lg) = 0 Then
        SV(r_tds_lg_win_pc) = 0
    Else
        SV(r_tds_lg_win_pc) = yI / SV(r_tds_lg)
    End If
    ' shorts winning
    If SV(r_tds_sh) = 0 Then
        SV(r_tds_sh_win_pc) = 0
    Else
        SV(r_tds_sh_win_pc) = zI / SV(r_tds_sh)
    End If
'' pips net
'    SV(r_net_pp) = Dlog(SV(r_cds), 13)
' pips per year
    SV(r_per_yr_pp) = SV(r_net_pp) / SV(r_yrs)
' pips per month
    SV(r_per_mn_pp) = SV(r_net_pp) / SV(r_mns)
' pips per week
    SV(r_per_w_pp) = SV(r_net_pp) / SV(r_wks)
' profit factor, a/c
    SV(r_pf_ac) = Abs(SV(r_mon_won) / SV(r_mon_lost))
' profit factor, pips
    SV(r_pf_pp) = Abs(SV(r_won_pp) / SV(r_lost_pp))
'' average trade in pips
'    SV(r_avg_td_pp) = SV(r_net_pp) / SV(r_tds_closed)
' average trade in a/c
    SV(r_avg_td_ac) = SV(r_net_ac) / SV(r_tds_closed)
' avg winner, pips
    If SV(r_tds_win_count) = 0 Then
        SV(r_avg_win_pp) = 0
    Else
        SV(r_avg_win_pp) = SV(r_won_pp) / SV(r_tds_win_count)
    End If
' avg loser, pips
    If SV(r_tds_los_count) = 0 Then
        SV(r_avg_los_pp) = 0
    Else
        SV(r_avg_los_pp) = SV(r_lost_pp) / SV(r_tds_los_count)
    End If
' avg win/los, pips
    SV(r_avg_win_los_pp) = SV(r_avg_win_pp) / SV(r_avg_los_pp)
' avg winner, a/c
    If SV(r_tds_win_count) = 0 Then
        SV(r_avg_win_ac) = 0
    Else
        SV(r_avg_win_ac) = SV(r_mon_won) / SV(r_tds_win_count)
    End If
' avg loser, a/c
    If SV(r_tds_los_count) = 0 Then
        SV(r_avg_los_ac) = 0
    Else
        SV(r_avg_los_ac) = SV(r_mon_lost) / SV(r_tds_los_count)
    End If
' avg win/los, a/c
    SV(r_avg_win_los_ac) = SV(r_avg_win_ac) / SV(r_avg_los_ac)
' winning trades, %
    SV(r_tds_win_pc) = SV(r_tds_win_count) / SV(r_tds_closed)
' max trades per day
    SV(r_tds_max_per_d) = 0
    yI = 1
    For xI = 2 To UBound(Tlog, 1)
        If Int(CDate(Tlog(xI, 7))) = Int(CDate(Tlog(xI - 1, 7))) Then   ' *! cdate x 2
            yI = yI + 1
        Else
            yI = 1
        End If
        If yI > SV(r_tds_max_per_d) Then
            SV(r_tds_max_per_d) = yI
        End If
    Next xI
' intervals winning/ losing; active
    ' months
    SV(r_mn_win) = 0
    SV(r_mn_los) = 0
    SV(r_mn_no_tds) = 0
    For xI = 1 To UBound(Mlog, 1)
        Select Case Mlog(xI, 2)
            Case Is > 0: SV(r_mn_win) = SV(r_mn_win) + 1
            Case Is < 0: SV(r_mn_los) = SV(r_mn_los) + 1
            Case Is = 0: SV(r_mn_no_tds) = SV(r_mn_no_tds) + 1
        End Select
    Next xI
    If SV(r_mn_los) = 0 Then
        SV(r_mn_win_los) = 0
    Else
        SV(r_mn_win_los) = SV(r_mn_win) / SV(r_mn_los)
    End If
    ' weeks
    SV(r_w_win) = 0
    SV(r_w_los) = 0
    SV(r_w_no_tds) = 0
    For xI = 1 To UBound(Wlog, 1)
        Select Case Wlog(xI, 2)
            Case Is > 0: SV(r_w_win) = SV(r_w_win) + 1
            Case Is < 0: SV(r_w_los) = SV(r_w_los) + 1
            Case Is = 0: SV(r_w_no_tds) = SV(r_w_no_tds) + 1
        End Select
    Next xI
    If SV(r_w_los) = 0 Then
        SV(r_w_win_los) = 0
    Else
        SV(r_w_win_los) = SV(r_w_win) / SV(r_w_los)
    End If
    ' days
    SV(r_d_win) = 0
    SV(r_d_los) = 0
    SV(r_d_no_tds) = 0
    For xI = 1 To UBound(Dlog, 1)
        Select Case Dlog(xI, 2)
            Case Is > 0: SV(r_d_win) = SV(r_d_win) + 1
            Case Is < 0: SV(r_d_los) = SV(r_d_los) + 1
            Case Is = 0: SV(r_d_no_tds) = SV(r_d_no_tds) + 1
        End Select
    Next xI
    If SV(r_d_los) = 0 Then
        SV(r_d_win_los) = 0
    Else
        SV(r_d_win_los) = SV(r_d_win) / SV(r_d_los)
    End If
' overnights
    SV(r_over_amnt_pp) = 0
    SV(r_ds_over) = 0
    For xI = 1 To UBound(Dlog, 1)
        If Dlog(xI, 7) <> 0 Then
            SV(r_ds_over) = SV(r_ds_over) + 1
            SV(r_over_amnt_pp) = SV(r_over_amnt_pp) + Dlog(xI, 7)
        End If
    Next xI
    SV(r_dwo_per_mn) = SV(r_ds_over) / SV(r_mns)
'
' active intervals
' months
    SV(r_mn_act) = 0
    rol_t = 1
    For xI = 1 To UBound(Mlog, 1)
        yI = 0
        If Tlog(rol_t, 7) > Mlog(xI - 1, 1) And Tlog(rol_t, 7) <= Int(Mlog(xI, 1)) Then
            Do While Tlog(rol_t, 7) > Mlog(xI - 1, 1) And Tlog(rol_t, 7) <= Int(Mlog(xI, 1))
                yI = yI + 1
                If rol_t < UBound(Tlog, 1) Then
                    rol_t = rol_t + 1
                Else
                    Exit Do
                End If
            Loop
        End If
        If yI > 0 Then
            SV(r_mn_act) = SV(r_mn_act) + 1
        End If
    Next xI
    SV(r_mn_act_all) = SV(r_mn_act) / SV(r_mns)
' weeks
    SV(r_w_act) = 0
    rol_t = 1
    For xI = 1 To UBound(Wlog, 1)
        yI = 0
        If Tlog(rol_t, 7) > Wlog(xI - 1, 1) And Tlog(rol_t, 7) <= Wlog(xI, 1) Then
            Do While Tlog(rol_t, 7) > Wlog(xI - 1, 1) And Tlog(rol_t, 7) <= Wlog(xI, 1)
                yI = yI + 1
                If rol_t < UBound(Tlog, 1) Then
                    rol_t = rol_t + 1
                Else
                    Exit Do
                End If
            Loop
        End If
        If yI > 0 Then
            SV(r_w_act) = SV(r_w_act) + 1
        End If
    Next xI
    SV(r_w_act_all) = SV(r_w_act) / SV(r_wks)
' days
    SV(r_d_act) = 0
    rol_t = 1
    For xI = 1 To UBound(Dlog, 1)
        yI = 0
        If Int(CDate(Tlog(rol_t, 7))) = Dlog(xI, 1) Then
            Do While Int(CDate(Tlog(rol_t, 7))) = Dlog(xI, 1)
                yI = yI + 1
                If rol_t < UBound(Tlog, 1) Then
                    rol_t = rol_t + 1
                Else
                    Exit Do
                End If
            Loop
        End If
        If yI > 0 Then
            SV(r_d_act) = SV(r_d_act) + 1
        End If
    Next xI
    SV(r_d_act_all) = SV(r_d_act) / SV(r_cds)
End Sub
Private Sub GSPR_Get_All_Stats_part_2()
    Dim cI As Integer, rol_t As Integer, xI As Integer
    Dim x_z_score As Double, p2_zs As Double, p3_zs As Double   ' Z score variables
' Exposition: Time in winners, days / Time in losers, days
    SV(r_tm_win_los) = SV(r_tm_in_win_tds) / SV(r_tm_in_los_tds)
' Trade duration: Avg, days
    SV(r_avg_dur) = SV(r_tm_in_tds) / SV(r_tds_closed)
' Trade duration: Avg winning, days
    SV(r_avg_win_dur) = SV(r_tm_in_win_tds) / SV(r_tds_win_count)
' Trade duration: Avg losing, days
    SV(r_avg_los_dur) = SV(r_tm_in_los_tds) / SV(r_tds_los_count)
' Trade duration: Avg win/los
    SV(r_avg_dur_win_los) = SV(r_avg_win_dur) / SV(r_avg_los_dur)
' Recovery factor: In a/c
    SV(r_rf_ac) = SV(r_ann_ret) / SV(r_mdd_ec_ac)
    If SV(r_rf_ac) < 0 Then
        SV(r_rf_ac) = 0
    End If
' Recovery factor: In pips
    SV(r_rf_pp) = SV(r_net_pp) / SV(r_mdd_ec_pp)
    If SV(r_rf_pp) < 0 Then
        SV(r_rf_pp) = 0
    End If
' Trades: Z score
    x_z_score = 2 * SV(r_tds_win_count) * SV(r_tds_los_count)
    p2_zs = SV(r_tds_closed) * (SV(r_runs_tds) - 0.5) - x_z_score
    p3_zs = Sqr(x_z_score * (x_z_score - SV(r_tds_closed)) / (SV(r_tds_closed) - 1))
    SV(r_zscore_tds) = p2_zs / p3_zs
' Weeks: Z score
    x_z_score = 2 * SV(r_w_win) * (SV(r_w_los) + SV(r_w_no_tds))
    p2_zs = SV(r_wks) * (SV(r_runs_wks) - 0.5) - x_z_score
    p3_zs = Sqr(x_z_score * (x_z_score - SV(r_wks)) / (SV(r_wks) - 1))
    SV(r_zscore_wks) = p2_zs / p3_zs
End Sub
Private Sub GSPR_Prep_SN()
' STATISTICS NAMES
    SN(r_s_report) = "REPORT"
    SN(r_name) = "Macro"
    SN(r_type) = "Report type"
    SN(r_date_gen) = "Date obtained"
    SN(r_time_gen) = "Time obtained"
    SN(r_file) = "Dukascopy report, link"
    SN(r_s_basic) = "MAIN DATA"
    SN(r_strat) = "Strategy name"
    SN(r_ac) = "Account currency (a/c)"
    SN(r_ins) = "Instrument"
    SN(r_init_depo) = "Deposit start"
    SN(r_fin_depo) = "Deposit finish"
    SN(r_s_return) = "RETURNS"
    SN(r_net_pc) = "Net, %"
    SN(r_net_ac) = "Net, a/c"
    SN(r_mon_won) = "Winning trades sum"
    SN(r_mon_lost) = "Losing trades sum"
    SN(r_ann_ret) = "Annualized return, %"
    SN(r_mn_ret) = "Monthly return, %"
    SN(r_s_pips) = "PIPS"
    SN(r_net_pp) = "Sum"
    SN(r_won_pp) = "Winning trades"
    SN(r_lost_pp) = "Losing trades"
    SN(r_per_yr_pp) = "Avg yearly"
    SN(r_per_mn_pp) = "Avg monthly"
    SN(r_per_w_pp) = "Avg weekly"
    SN(r_s_rsq) = "R-SQUARED"
    SN(r_rsq_tr_cve) = "Pips sum curve"
    SN(r_rsq_eq_cve) = "Equity curve"
    SN(r_s_pf) = "PROFIT FACTOR"
    SN(r_pf_ac) = "A/c"
    SN(r_pf_pp) = "Pips"
    SN(r_s_rf) = "RECOVERY FACTOR"
    SN(r_rf_ac) = "A/c"
    SN(r_rf_pp) = "Pips"
    SN(r_s_avgs_pp) = "AVG PIPS"
    SN(r_avg_td_pp) = "Trade"
    SN(r_avg_win_pp) = "Winner"
    SN(r_avg_los_pp) = "Loser"
    SN(r_avg_win_los_pp) = "Winner/Loser"
    SN(r_s_avgs_ac) = "AVG IN A/C"
    SN(r_avg_td_ac) = "Trade"
    SN(r_avg_win_ac) = "Winner"
    SN(r_avg_los_ac) = "Loser"
    SN(r_avg_win_los_ac) = "Winner/Loser"
    SN(r_s_intvl) = "TIME INTERVALS"
    SN(r_mn_win) = "Months winning"
    SN(r_mn_los) = "Months losing"
    SN(r_mn_no_tds) = "Months no trades"
    SN(r_mn_win_los) = "Months win/los"
    SN(r_w_win) = "Weeks winning"
    SN(r_w_los) = "Weeks losing"
    SN(r_w_no_tds) = "Weeks no trades"
    SN(r_w_win_los) = "Weeks win/los"
    SN(r_d_win) = "Days winninng"
    SN(r_d_los) = "Days losing"
    SN(r_d_no_tds) = "Days no trades"
    SN(r_d_win_los) = "Days win/los"
    SN(r_s_act_intvl) = "ACTIVE INTERVALS"
    SN(r_mn_act) = "Months active"
    SN(r_mn_act_all) = "Months act/all"
    SN(r_w_act) = "Weeks active"
    SN(r_w_act_all) = "Weeks act/all"
    SN(r_d_act) = "Days active"
    SN(r_d_act_all) = "Days act/all"
    SN(r_s_std) = "STD DEV"
    SN(r_std_tds_pp) = "Trades (pips)"
    SN(r_std_tds_ac) = "Trades (a/c)"
'
    SN(r_s_time) = "HISTORICAL WINDOWS"
    SN(r_dt_begin) = "Test begin"
    SN(r_dt_end) = "Test end"
    SN(r_yrs) = "Years"
    SN(r_mns) = "Months"
    SN(r_wks) = "Weeks"
    SN(r_cds) = "Calendar days"
    SN(r_s_cmsn) = "COMMISSIONS"
    SN(r_cmsn_amnt_ac) = "Sum in a/c"
    SN(r_cmsn_avg_per_d) = "Avg per day in a/c"
    SN(r_s_mdd_ac) = "MDD, MFE IN A/C"
    SN(r_mdd_ec_ac) = "MDD on equity curve"
    SN(r_mfe_ec_ac) = "MFE on equity curve"
    SN(r_abs_hi_ac) = "Abs max"
    SN(r_abs_lo_ac) = "Abs min"
    SN(r_s_mdd_pp) = "MDD, MFE IN PIPS"
    SN(r_mdd_ec_pp) = "MDD on pips curve"
    SN(r_mfe_ec_pp) = "MFE on pips curve"
    SN(r_abs_hi_pp) = "Abs max"
    SN(r_abs_lo_pp) = "Abs min"
    SN(r_s_trades) = "TRADES"
    SN(r_tds_closed) = "Closed"
    SN(r_tds_per_yr) = "Per year"
    SN(r_tds_per_mn) = "Per month"
    SN(r_tds_per_w) = "Per week"
    SN(r_tds_max_per_d) = "Max per day"
    SN(r_tds_win_count) = "Winners"
    SN(r_tds_los_count) = "Losers"
    SN(r_tds_win_pc) = "Winners, %"
    SN(r_tds_lg) = "Long"
    SN(r_tds_sh) = "Short"
    SN(r_tds_lg_sh) = "Long/Short"
    SN(r_tds_lg_win_pc) = "Long, winners, %"
    SN(r_tds_sh_win_pc) = "Short, winners, %"
    SN(r_s_dur) = "TRADE DURATION"
    SN(r_avg_dur) = "Avg, days"
    SN(r_avg_win_dur) = "Avg winner, days"
    SN(r_avg_los_dur) = "Avg loser, days"
    SN(r_avg_dur_win_los) = "Avg win/los"
    SN(r_s_stks) = "STREAKS"
    SN(r_stk_win_tds) = "Max winning, trades"
    SN(r_stk_los_tds) = "Max losing, trades"
    SN(r_stk_win_mns) = "Max winning, months"
    SN(r_stk_los_mns) = "Max losing, months"
    SN(r_stk_win_wks) = "Max winning, weeks"
    SN(r_stk_los_wks) = "Max losing, weeks"
    SN(r_stk_win_ds) = "Max winning, days"
    SN(r_stk_los_ds) = "Max losing, days"
    SN(r_runs_tds) = "Streaks, trades"
    SN(r_zscore_tds) = "Z-score, trades"
    SN(r_runs_wks) = "Streaks, weeks"
    SN(r_zscore_wks) = "Z-score, weeks"
    SN(r_s_over) = "OVERNIGHTS"
    SN(r_over_amnt_pp) = "Sum, pips"
    SN(r_ds_over) = "Days with overnights"
    SN(r_dwo_per_mn) = "Days with o/n per mn"
    SN(r_s_expo) = "EXPOSURE"
    SN(r_tm_in_tds) = "Days in positions"
    SN(r_tm_in_win_tds) = "Days in winners"
    SN(r_tm_in_los_tds) = "Days in losers"
    SN(r_tm_win_los) = "Days in win/los"
    SN(r_s_orders) = "ORDERS"
    SN(r_ord_sent) = "Orders sent"
    SN(r_ord_tds) = "Orders/Positions ratio"
End Sub
Private Sub GSPR_Prep_all_fm()
' Column A
    ReDim sect_fm_A(12)
    sect_fm_A(1) = r_s_report
    sect_fm_A(2) = r_s_basic
    sect_fm_A(3) = r_s_return
    sect_fm_A(4) = r_s_pips
    sect_fm_A(5) = r_s_rsq
    sect_fm_A(6) = r_s_pf
    sect_fm_A(7) = r_s_rf
    sect_fm_A(8) = r_s_avgs_pp
    sect_fm_A(9) = r_s_avgs_ac
    sect_fm_A(10) = r_s_intvl
    sect_fm_A(11) = r_s_act_intvl
    sect_fm_A(12) = r_s_std
' Column C
    ReDim sect_fm_C(10)
    sect_fm_C(1) = r_s_time - split_row
    sect_fm_C(2) = r_s_cmsn - split_row
    sect_fm_C(3) = r_s_mdd_ac - split_row
    sect_fm_C(4) = r_s_mdd_pp - split_row
    sect_fm_C(5) = r_s_trades - split_row
    sect_fm_C(6) = r_s_dur - split_row
    sect_fm_C(7) = r_s_stks - split_row
    sect_fm_C(8) = r_s_over - split_row
    sect_fm_C(9) = r_s_expo - split_row
    sect_fm_C(10) = r_s_orders - split_row
' "yyyy-mm-dd"
    ReDim fm_date_A(1)
    fm_date_A(1) = r_date_gen
    ReDim fm_date_C(2)
    fm_date_C(1) = r_dt_begin - split_row
    fm_date_C(2) = r_dt_end - split_row
' "0"
    ReDim fm_0_A(12)
    fm_0_A(1) = r_mn_win
    fm_0_A(2) = r_mn_los
    fm_0_A(3) = r_mn_no_tds
    fm_0_A(4) = r_w_win
    fm_0_A(5) = r_w_los
    fm_0_A(6) = r_w_no_tds
    fm_0_A(7) = r_d_win
    fm_0_A(8) = r_d_los
    fm_0_A(9) = r_d_no_tds
    fm_0_A(10) = r_mn_act
    fm_0_A(11) = r_w_act
    fm_0_A(12) = r_d_act
    ReDim fm_0_C(19)
    fm_0_C(1) = r_cds - split_row
    fm_0_C(2) = r_tds_closed - split_row
    fm_0_C(3) = r_tds_max_per_d - split_row
    fm_0_C(4) = r_tds_win_count - split_row
    fm_0_C(5) = r_tds_los_count - split_row
    fm_0_C(6) = r_tds_lg - split_row
    fm_0_C(7) = r_tds_sh - split_row
    fm_0_C(8) = r_stk_win_tds - split_row
    fm_0_C(9) = r_stk_los_tds - split_row
    fm_0_C(10) = r_stk_win_mns - split_row
    fm_0_C(11) = r_stk_los_mns - split_row
    fm_0_C(12) = r_stk_win_wks - split_row
    fm_0_C(13) = r_stk_los_wks - split_row
    fm_0_C(14) = r_stk_win_ds - split_row
    fm_0_C(15) = r_stk_los_ds - split_row
    fm_0_C(16) = r_runs_tds - split_row
    fm_0_C(17) = r_runs_wks - split_row
    fm_0_C(18) = r_ds_over - split_row
    fm_0_C(19) = r_ord_sent - split_row
' "0.00"
    ReDim fm_0p00_A(33)
    fm_0p00_A(1) = r_init_depo
    fm_0p00_A(2) = r_fin_depo
    fm_0p00_A(3) = r_net_ac
    fm_0p00_A(4) = r_mon_won
    fm_0p00_A(5) = r_mon_lost
    fm_0p00_A(6) = r_net_pp
    fm_0p00_A(7) = r_won_pp
    fm_0p00_A(8) = r_lost_pp
    fm_0p00_A(9) = r_per_yr_pp
    fm_0p00_A(10) = r_per_mn_pp
    fm_0p00_A(11) = r_per_w_pp
    fm_0p00_A(12) = r_rsq_tr_cve
    fm_0p00_A(13) = r_rsq_eq_cve
    fm_0p00_A(14) = r_pf_ac
    fm_0p00_A(15) = r_pf_pp
    fm_0p00_A(16) = r_rf_ac
    fm_0p00_A(17) = r_rf_pp
    fm_0p00_A(18) = r_avg_td_pp
    fm_0p00_A(19) = r_avg_win_pp
    fm_0p00_A(20) = r_avg_los_pp
    fm_0p00_A(21) = r_avg_win_los_pp
    fm_0p00_A(22) = r_avg_td_ac
    fm_0p00_A(23) = r_avg_win_ac
    fm_0p00_A(24) = r_avg_los_ac
    fm_0p00_A(25) = r_avg_win_los_ac
    fm_0p00_A(26) = r_mn_win_los
    fm_0p00_A(27) = r_w_win_los
    fm_0p00_A(28) = r_d_win_los
    fm_0p00_A(29) = r_mn_act_all
    fm_0p00_A(30) = r_w_act_all
    fm_0p00_A(31) = r_d_act_all
    fm_0p00_A(32) = r_std_tds_pp
    fm_0p00_A(33) = r_std_tds_ac
    ReDim fm_0p00_C(28)
    fm_0p00_C(1) = r_yrs - split_row
    fm_0p00_C(2) = r_mns - split_row
    fm_0p00_C(3) = r_wks - split_row
    fm_0p00_C(4) = r_cmsn_amnt_ac - split_row
    fm_0p00_C(5) = r_cmsn_avg_per_d - split_row
    fm_0p00_C(6) = r_abs_hi_ac - split_row
    fm_0p00_C(7) = r_abs_lo_ac - split_row
    fm_0p00_C(8) = r_mdd_ec_pp - split_row
    fm_0p00_C(9) = r_mfe_ec_pp - split_row
    fm_0p00_C(10) = r_abs_hi_pp - split_row
    fm_0p00_C(11) = r_abs_lo_pp - split_row
    fm_0p00_C(12) = r_tds_per_yr - split_row
    fm_0p00_C(13) = r_tds_per_mn - split_row
    fm_0p00_C(14) = r_tds_per_w - split_row
    fm_0p00_C(15) = r_tds_lg_sh - split_row
    fm_0p00_C(16) = r_avg_dur - split_row
    fm_0p00_C(17) = r_avg_win_dur - split_row
    fm_0p00_C(18) = r_avg_los_dur - split_row
    fm_0p00_C(19) = r_avg_dur_win_los - split_row
    fm_0p00_C(20) = r_zscore_tds - split_row
    fm_0p00_C(21) = r_zscore_wks - split_row
    fm_0p00_C(22) = r_over_amnt_pp - split_row
    fm_0p00_C(23) = r_dwo_per_mn - split_row
    fm_0p00_C(24) = r_tm_in_tds - split_row
    fm_0p00_C(25) = r_tm_in_win_tds - split_row
    fm_0p00_C(26) = r_tm_in_los_tds - split_row
    fm_0p00_C(27) = r_tm_win_los - split_row
    fm_0p00_C(28) = r_ord_tds - split_row
' "0.00%"
    ReDim fm_0p00pc_A(3)
    fm_0p00pc_A(1) = r_net_pc
    fm_0p00pc_A(2) = r_ann_ret
    fm_0p00pc_A(3) = r_mn_ret
    ReDim fm_0p00pc_C(5)
    fm_0p00pc_C(1) = r_mdd_ec_ac - split_row
    fm_0p00pc_C(2) = r_mfe_ec_ac - split_row
    fm_0p00pc_C(3) = r_tds_win_pc - split_row
    fm_0p00pc_C(4) = r_tds_lg_win_pc - split_row
    fm_0p00pc_C(5) = r_tds_sh_win_pc - split_row
End Sub
Private Sub GSPR_Report_Status()
    SV(r_name) = macro_name
    SV(r_type) = report_type
    SV(r_date_gen) = Date
    SV(r_time_gen) = Time
End Sub
Private Sub GSPR_Show_SN_SV()
    Dim xI As Integer
' add new sheet
    Set ns = mb.Sheets.Add(after:=ms)
    Set nc = ns.Cells
' show statistics
    For xI = LBound(SN) To split_row
        nc(xI, 1) = SN(xI)
        nc(xI, 2) = SV(xI)
    Next xI
    For xI = split_row + 1 To UBound(SN)
        nc(xI - split_row, 3) = SN(xI)
        nc(xI - split_row, 4) = SV(xI)
    Next xI
' autofit columns
    ns.Columns(1).AutoFit
    ns.Columns(3).AutoFit
End Sub
Private Sub GSPR_Format_SV_SN()
    Dim xI As Integer
' section color
    For xI = LBound(sect_fm_A) To UBound(sect_fm_A)
        With nc(sect_fm_A(xI), 1)
            .Interior.Color = RGB(219, 229, 241)
            .HorizontalAlignment = xlCenter
        End With
    Next xI
    For xI = LBound(sect_fm_C) To UBound(sect_fm_C)
        With nc(sect_fm_C(xI), 3)
            .Interior.Color = RGB(219, 229, 241)
            .HorizontalAlignment = xlCenter
        End With
    Next xI
' "yyyy-mm-dd"
    For xI = LBound(fm_date_A) To UBound(fm_date_A)
        nc(fm_date_A(xI), 2).NumberFormat = "yyyy-mm-dd"
    Next xI
    For xI = LBound(fm_date_C) To UBound(fm_date_C)
        nc(fm_date_C(xI), 4).NumberFormat = "yyyy-mm-dd"
    Next xI
' "hh:mm:ss"
    nc(r_time_gen, 2).NumberFormat = "hh:mm:ss"
' hyperlink
    ns.Hyperlinks.Add anchor:=nc(r_file, 2), Address:=rep_adr
' "0"
    For xI = LBound(fm_0_A) To UBound(fm_0_A)
        nc(fm_0_A(xI), 2).NumberFormat = "0"
    Next xI
    For xI = LBound(fm_0_C) To UBound(fm_0_C)
        nc(fm_0_C(xI), 4).NumberFormat = "0"
    Next xI
' "0.00"
    For xI = LBound(fm_0p00_A) To UBound(fm_0p00_A)
        nc(fm_0p00_A(xI), 2).NumberFormat = "0.00"
    Next xI
    For xI = LBound(fm_0p00_C) To UBound(fm_0p00_C)
        nc(fm_0p00_C(xI), 4).NumberFormat = "0.00"
    Next xI
' "0.00%"
    For xI = LBound(fm_0p00pc_A) To UBound(fm_0p00pc_A)
        nc(fm_0p00pc_A(xI), 2).NumberFormat = "0.00%"
    Next xI
    For xI = LBound(fm_0p00pc_C) To UBound(fm_0p00pc_C)
        nc(fm_0p00pc_C(xI), 4).NumberFormat = "0.00%"
    Next xI
End Sub
Private Sub GSPR_Fill_Logs_Heads()
' daily log head
    Dlog_head(1) = "EOD"
    Dlog_head(2) = "Return"
    Dlog_head(3) = "Open"
    Dlog_head(4) = "High"
    Dlog_head(5) = "Low"
    Dlog_head(6) = "Close"
    Dlog_head(7) = "Swaps, pips"
    Dlog_head(8) = "Cmsn"
    Dlog_head(9) = "Return"
    Dlog_head(10) = "Open"
    Dlog_head(11) = "High"
    Dlog_head(12) = "Low"
    Dlog_head(13) = "Clsoe"
' weekly log head
    Wlog_head(1) = "EOW"
    Wlog_head(2) = "Return"
    Wlog_head(3) = "Open"
    Wlog_head(4) = "High"
    Wlog_head(5) = "Low"
    Wlog_head(6) = "Close"
    Wlog_head(7) = "Swaps, pips"
    Wlog_head(8) = "Cmsn"
    Wlog_head(9) = "Return"
    Wlog_head(10) = "Open"
    Wlog_head(11) = "High"
    Wlog_head(12) = "Low"
    Wlog_head(13) = "Close"
' monthly log head
    Mlog_head(1) = "EOM"
    Mlog_head(2) = "Return"
    Mlog_head(3) = "Open"
    Mlog_head(4) = "High"
    Mlog_head(5) = "Low"
    Mlog_head(6) = "Close"
    Mlog_head(7) = "Swaps, pips"
    Mlog_head(8) = "Cmsn"
    Mlog_head(9) = "Return"
    Mlog_head(10) = "Open"
    Mlog_head(11) = "High"
    Mlog_head(12) = "Low"
    Mlog_head(13) = "Close"
End Sub
Private Sub GSPR_Show_All_Logs()
    Dim rI As Integer, cI As Integer, tI_co As Integer
' PARAMETERS
    nc(di_r, di_c) = "PARAMETERS"
' show parameters
    For rI = LBound(Par, 1) To UBound(Par, 1)
        For cI = LBound(Par, 2) To UBound(Par, 2)
            nc(di_r + rI, di_c - 1 + cI) = Par(rI, cI)
        Next cI
    Next rI
' sort alphabetically - parameter name
    cI = nc(di_r, di_c).End(xlDown).Row    ' last row of parameters
    With ns.Sort
        .SortFields.Clear
        .SortFields.Add Key:=nc(di_r + 1, di_c), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange ns.Range(nc(di_r + 1, di_c), nc(cI, di_c + 1))
        .Header = xlNo
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
' TRADE LOG
    tI_co = di_c + UBound(Par, 2)
    nc(di_r, tI_co) = "Trade log"
' show trade log head
    For cI = LBound(Tlog_head) To UBound(Tlog_head)
        nc(di_r + 1, tI_co - 1 + cI) = Tlog_head(cI)
    Next cI
' show trade log
    For rI = LBound(Tlog, 1) To UBound(Tlog, 1)
        For cI = LBound(Tlog, 2) To UBound(Tlog, 2)
            nc(di_r + 1 + rI, tI_co - 1 + cI) = Tlog(rI, cI)
        Next cI
    Next rI
' DAILY LOG
    tI_co = tI_co + UBound(Tlog_head)
    nc(di_r, tI_co) = "Trade log, calendar days"
    nc(di_r, tI_co + 1) = "In account currency"
    nc(di_r, tI_co + 8) = "In pips"
' show head
    For cI = LBound(Dlog_head) To UBound(Dlog_head)
        nc(di_r + 1, tI_co - 1 + cI) = Dlog_head(cI)
    Next cI
' show log
    For rI = LBound(Dlog, 1) To UBound(Dlog, 1)
        For cI = LBound(Dlog, 2) To UBound(Dlog, 2)
            nc(di_r + 2 + rI, tI_co - 1 + cI) = Dlog(rI, cI)
        Next cI
    Next rI
' WEEKLY LOG
    tI_co = tI_co + UBound(Wlog_head)
    nc(di_r, tI_co) = "Trade log, weeks"
    nc(di_r, tI_co + 1) = "In account currency"
    nc(di_r, tI_co + 8) = "In pips"
' show head
    For cI = LBound(Wlog_head) To UBound(Wlog_head)
        nc(di_r + 1, tI_co - 1 + cI) = Wlog_head(cI)
    Next cI
' show log
    For rI = LBound(Wlog, 1) To UBound(Wlog, 1)
        For cI = LBound(Wlog, 2) To UBound(Wlog, 2)
            nc(di_r + 2 + rI, tI_co - 1 + cI) = Wlog(rI, cI)
        Next cI
    Next rI
' MONTHLY LOG
    tI_co = tI_co + UBound(Mlog_head)
    nc(di_r, tI_co) = "Trade log in monghts"
    nc(di_r, tI_co + 1) = "In account currency"
    nc(di_r, tI_co + 8) = "In pips"
' show head
    For cI = LBound(Mlog_head) To UBound(Mlog_head)
        nc(di_r + 1, tI_co - 1 + cI) = Mlog_head(cI)
    Next cI
' show log
    For rI = LBound(Mlog, 1) To UBound(Mlog, 1)
        For cI = LBound(Mlog, 2) To UBound(Mlog, 2)
            nc(di_r + 2 + rI, tI_co - 1 + cI) = Mlog(rI, cI)
        Next cI
    Next rI
End Sub
Private Sub GSPR_Build_Charts()
    Dim chsht As Worksheet
    Dim ulr As Integer, ulc As Integer      ' upper left row, upper left column
    Dim chW As Integer, chH As Integer      ' chart width, height - in cells
    Dim rngX As Range, rngY As Range        ' source and axis
    Dim ChTitle As String                   ' chart title
    Dim MinVal As Long, maxVal As Long      ' Min & Max values
    Dim chobj_n As Integer                  ' number of chart object
    Dim t_co As Integer
    Dim t_rof As Integer, t_rol As Integer
    Const ch_round As Integer = 100
    
    ns.Activate
    Set chsht = ActiveSheet
    chobj_n = 0
' =====================================
' Pips curve - line
    ulr = 1
    ulc = di_c + UBound(Par, 2)
    chW = UBound(Tlog, 2)
    chH = Int(di_r * 0.6)
    t_co = di_c + UBound(Par, 2) + 10    ' Tlog "pips curve" column
    t_rof = di_r + 2
    t_rol = di_r + 1 + UBound(Tlog, 1)
    Set rngY = Range(nc(t_rof, t_co), nc(t_rol, t_co))
    ChTitle = "PIPS SUM. Strategy: " & SV(r_strat) & ". Instrument: " & SV(r_ins) & ". Dates: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & ". Result = " & SV(r_net_pp) & " pips."
' pips abs high low
    If SV(r_abs_lo_pp) < 0 Then
        MinVal = 100 * Int(SV(r_abs_lo_pp) / 100)
    Else
        MinVal = 100 * Int(SV(r_abs_lo_pp) / 100) - 100
    End If
    maxVal = 100 * Int(SV(r_abs_hi_pp) / 100) + 100
    chobj_n = chobj_n + 1
    Call GSPR_Charts_Line_Y_wMinMax(chsht, ulr, ulc, chW, chH, rngY, ChTitle, MinVal, maxVal, chobj_n)
' =====================================
' Pips histogram
    ulr = Int(di_r * 0.6) + 1
    ulc = di_c + UBound(Par, 2)
    chW = UBound(Tlog, 2)
    chH = Int(di_r * 0.4)
    t_co = di_c + UBound(Par, 2) + 5    ' Tlog "profit/loss in pips" column
    t_rof = di_r + 2
    t_rol = di_r + 1 + UBound(Tlog, 1)
    Set rngY = Range(nc(t_rof, t_co), nc(t_rol, t_co))
    ChTitle = "TRADES IN PIPS. Strategy: " & SV(r_strat) & ". Instrument: " & SV(r_ins) & ". Dates: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & "."
    chobj_n = chobj_n + 1
    Call GSPR_Charts_Hist_Y(chsht, ulr, ulc, chW, chH, rngY, ChTitle, chobj_n)
' =====================================
' Daily log - line
    ulr = 1
    ulc = ulc + UBound(Tlog, 2)
    chW = UBound(Dlog, 2)
    chH = Int(di_r * 0.6)
    t_co = ulc + 5      ' daily close column
    t_rof = di_r + 2
    t_rol = di_r + 2 + UBound(Dlog, 1)
    Set rngX = Range(nc(t_rof, t_co - 5), nc(t_rol, t_co - 5))
    Set rngY = Range(nc(t_rof, t_co), nc(t_rol, t_co))
    ChTitle = "DAILY EQUITY CURVE. Strategy: " & SV(r_strat) & ". Instrument: " & SV(r_ins) & ". Dates: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & ". Result = " & SV(r_fin_depo) & " " & SV(r_ac) & "."
' pips abs high low
    MinVal = ch_round * Int(SV(r_abs_lo_ac) / ch_round) ' minus 1000 - no
    maxVal = ch_round * Int(SV(r_abs_hi_ac) / ch_round) + ch_round
    chobj_n = chobj_n + 1
    Call GSPR_Charts_Line_XY_wMinMax(chsht, ulr, ulc, chW, chH, rngX, rngY, ChTitle, MinVal, maxVal, chobj_n)
' =====================================
' Daily log histogram
    ulr = Int(di_r * 0.6) + 1
    chW = UBound(Dlog, 2)
    chH = Int(di_r * 0.4)
    t_co = ulc + 1
    Set rngY = Range(nc(t_rof, t_co), nc(t_rol, t_co))
    ChTitle = "DAY RESULT IN " & SV(r_ac) & ". Strategy: " & SV(r_strat) & ". Instrument: " & SV(r_ins) & ". Dates: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & "."
    chobj_n = chobj_n + 1
    Call GSPR_Charts_Hist_Y(chsht, ulr, ulc, chW, chH, rngY, ChTitle, chobj_n)
' =====================================
' Weekly log - OHLC
    ulr = 1
    ulc = ulc + UBound(Dlog, 2)
    chW = UBound(Wlog, 2)
    chH = Int(di_r * 0.6)
    t_co = ulc + 2      ' daily close column
    t_rof = di_r + 2
    t_rol = di_r + 2 + UBound(Wlog, 1)
    Set rngX = Range(nc(t_rof, t_co - 2), nc(t_rol, t_co - 2))
    Set rngY = Range(nc(t_rof, t_co), nc(t_rol, t_co + 3))
    ChTitle = "WEEKLY EQUITY CURVE, OHLC. Strategy: " & SV(r_strat) & ". Instrument: " & SV(r_ins) & ". Sates: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & ". Result = " & SV(r_fin_depo) & " " & SV(r_ac) & "."
' pips abs high low
    MinVal = ch_round * Int(SV(r_abs_lo_ac) / ch_round)
    maxVal = ch_round * Int(SV(r_abs_hi_ac) / ch_round) + ch_round
    chobj_n = chobj_n + 1
    Call GSPR_Charts_OHLC_wMinMax(chsht, ulr, ulc, chW, chH, rngX, rngY, ChTitle, MinVal, maxVal, chobj_n)
' =====================================
' Weekly log histogram
    ulr = Int(di_r * 0.6) + 1
    chW = UBound(Wlog, 2)
    chH = Int(di_r * 0.4)
    t_co = ulc + 1
    Set rngY = Range(nc(t_rof, t_co), nc(t_rol, t_co))
    ChTitle = "WEEK RESULT INT " & SV(r_ac) & ". Strategy: " & SV(r_strat) & ". Instrument: " & SV(r_ins) & ". Dates: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & "."
    chobj_n = chobj_n + 1
    Call GSPR_Charts_Hist_Y(chsht, ulr, ulc, chW, chH, rngY, ChTitle, chobj_n)
' =====================================
' Monthly log - OHLC
    ulr = 1
    ulc = ulc + UBound(Wlog, 2)
    chW = UBound(Mlog, 2)
    chH = Int(di_r * 0.6)
    t_co = ulc + 2
    t_rof = di_r + 2
    t_rol = di_r + 2 + UBound(Mlog, 1)
    Set rngX = Range(nc(t_rof, t_co - 2), nc(t_rol, t_co - 2))
    Set rngY = Range(nc(t_rof, t_co), nc(t_rol, t_co + 3))
    ChTitle = "MONTHLY EQUITY CURVE, OHLC. Strategy: " & SV(r_strat) & ". Instrument: " & SV(r_ins) & ". Dates: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & ". Result = " & SV(r_fin_depo) & " " & SV(r_ac) & "."
' pips abs high low
    MinVal = ch_round * Int(SV(r_abs_lo_ac) / ch_round)
    maxVal = ch_round * Int(SV(r_abs_hi_ac) / ch_round) + ch_round
    chobj_n = chobj_n + 1
    Call GSPR_Charts_OHLC_wMinMax(chsht, ulr, ulc, chW, chH, rngX, rngY, ChTitle, MinVal, maxVal, chobj_n)
' =====================================
' Monthly log histogram
    ulr = Int(di_r * 0.6) + 1
    chW = UBound(Mlog, 2)
    chH = Int(di_r * 0.4)
    t_co = ulc + 1
    Set rngY = Range(nc(t_rof, t_co), nc(t_rol, t_co))
    ChTitle = "MONTH RESULT IN " & SV(r_ac) & ". Strategy: " & SV(r_strat) & ". Instrument: " & SV(r_ins) & ". Dates: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & "."
    chobj_n = chobj_n + 1
    Call GSPR_Charts_Hist_Y(chsht, ulr, ulc, chW, chH, rngY, ChTitle, chobj_n)
    nc(1, 1).Activate
End Sub
Private Sub GSPR_Charts_OHLC_wMinMax(chsht As Worksheet, _
                        ulr As Integer, _
                        ulc As Integer, _
                        chW As Integer, _
                        chH As Integer, _
                        rngX As Range, _
                        rngY As Range, _
                        ChTitle As String, _
                        MinVal As Long, _
                        maxVal As Long, _
                        chobj_n As Integer)
    chW = chW * 48      ' chart width, pixels
    chH = chH * 15      ' chart height, pixels
    chsht.Activate                      ' activate sheet
    rngX.Select ' previously RngY.Select
    chsht.Shapes.AddChart.Select        ' build chart
    With ActiveChart
        .SetSourceData Source:=rngY
        .ChartType = xlStockOHLC                        ' chart type - OHLC
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = maxVal
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow     ' axis position
        .SeriesCollection(1).XValues = rngX             ' lower axis data
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
Private Sub GSPR_Charts_Line_Y_wMinMax(chsht As Worksheet, _
                        ulr As Integer, _
                        ulc As Integer, _
                        chW As Integer, _
                        chH As Integer, _
                        rngY As Range, _
                        ChTitle As String, _
                        MinVal As Long, _
                        maxVal As Long, _
                        chobj_n As Integer)
    chW = chW * 48      ' chart width, pixels
    chH = chH * 15      ' chart height, pixels
    chsht.Activate                      ' activate sheet
    rngY.Select
    chsht.Shapes.AddChart.Select        ' build chart
    With ActiveChart
        .SetSourceData Source:=rngY
        .ChartType = xlLine                        ' chart type - OHLC
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = maxVal
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
Private Sub GSPR_Charts_Line_XY_wMinMax(chsht As Worksheet, _
                        ulr As Integer, _
                        ulc As Integer, _
                        chW As Integer, _
                        chH As Integer, _
                        rngX As Range, _
                        rngY As Range, _
                        ChTitle As String, _
                        MinVal As Long, _
                        maxVal As Long, _
                        chobj_n As Integer)
    chW = chW * 48      ' chart width, pixels
    chH = chH * 15      ' chart height, pixels
    chsht.Activate                      ' activate sheet
    rngY.Select
    chsht.Shapes.AddChart.Select        ' build chart
    With ActiveChart
        .SetSourceData Source:=rngY
        .ChartType = xlLine                        ' chart type - OHLC
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = maxVal
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow     ' axis position
        .SeriesCollection(1).XValues = rngX             ' lower axis data
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
Private Sub GSPR_Charts_Hist_Y(chsht As Worksheet, _
                            ulr As Integer, _
                            ulc As Integer, _
                            chW As Integer, _
                            chH As Integer, _
                            rngY As Range, _
                            ChTitle As String, _
                            chobj_n As Integer)
    chW = chW * 48      ' chart width, pixels
    chH = chH * 15      ' chart height, pixels
    chsht.Activate                      ' activate sheet
    rngY.Select
    chsht.Shapes.AddChart.Select        ' build chart
    With ActiveChart
        .SetSourceData Source:=rngY
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



' MODULE: Rep_Multiple
Option Explicit
Option Base 1
    Const addin_file_name As String = "GetStats_BackTest_v1.21.xlsm"
    Const rep_type As String = "GS_Pro_Single_Core"
    Const macro_ver As String = "GetStats Pro v1.20"
    Const max_htmls As Integer = 999
    Const depo_ini_ok As Double = 10000
    
    Dim fd As FileDialog
    Dim open_fail As Boolean
' ARRAYS
    Dim ov() As Variant
    Dim SV() As Variant
    Dim Par() As Variant, par_sum() As Variant, par_sum_head() As String
    Dim sM() As Variant
    Dim t1() As Variant
    Dim t2() As Variant
    Dim fm_date(1 To 2) As Integer, fm_0p00(1 To 10) As Integer, fm_0p00pc(1 To 3) As Integer, fm_clr(1 To 5) As Integer     ' count before changing
' OBJECTS
    Dim mb As Workbook
    Dim addin_book As Workbook
'
    Dim addin_c As Range
    Dim last_row_reports As Integer
'---SHIFTS
'------ sn, sv
    Dim s_strat As Integer, s_ins As Integer
    Dim s_tpm As Integer, s_ar As Integer, s_mdd As Integer, s_rf As Integer, s_rsq As Integer
    Dim s_date_begin As Integer, s_date_end As Integer, s_mns As Integer
    Dim s_trades As Integer, s_win_pc As Integer, s_pips As Integer, s_avg_w2l As Integer, s_avg_pip As Integer
    Dim s_depo_ini As Integer, s_depo_fin As Integer, s_cmsn As Integer, s_link As Integer, s_rep_type As Integer
'------ ov_bas
    Dim s_ov_strat As Integer, s_ov_ins As Integer, s_ov_htmls As Integer
    Dim s_ov_mns As Integer, s_ov_from As Integer, s_ov_to As Integer
    Dim s_ov_params As Integer, s_ov_params_vbl As Integer, s_ov_created As Integer, s_ov_macro_ver As Integer
' separator variables
    Dim current_decimal As String
    Dim undo_sep As Boolean, undo_usesyst As Boolean
' BOOLEAN
    Dim all_zeros As Boolean
' STRING
    Dim Folder_To_Save As String
    Dim fNm As String
Private Sub GSPRM_Merge_Sharpe()
'
' RIBBON > BUTTON "Merge_SR"
'
    Dim wbA As Workbook, wbB As Workbook
    Dim wbksSelected As Integer
    Dim i As Integer
    Dim lr As Integer
    Dim pos As Integer
    Dim tstr As String

    Dim s As Worksheet
    Dim rg As Range
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select reports to merge by Sharpe Ratio"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Reports by GetStats", "*.xlsx"
'        .ButtonName = "OK"
    End With
    If fd.Show = 0 Then
        MsgBox "No files picked!"
        Exit Sub
    End If
    wbksSelected = fd.SelectedItems.count
    ' Create new workbook with 1 sheet
    Call Create_WB_N_Sheets(wbA, 1)
    
    Application.ScreenUpdating = False
    For i = 1 To wbksSelected
        Application.StatusBar = "Adding sheet " & i & " (" & wbksSelected & ")."
        Set wbB = Workbooks.Open(fd.SelectedItems(i))
        ' Add Parameters to summary sheet
        If wbB.Sheets("results").Cells(1, 1) = vbEmpty Then
            Call Params_To_Summary_Sharpe(wbB)
        End If
        ' Calculate Sharpe ratios
        Call SharpeBeforeMerge(wbB)
        tstr = wbB.Name
        pos = InStr(1, tstr, "-", 1)
        tstr = Right(Left(tstr, pos + 6), 6)
        If wbB.Sheets(2).Name = "results" Then
            wbB.Sheets("results").Copy after:=wbA.Sheets(wbA.Sheets.count)
            Set s = wbA.Sheets(wbA.Sheets.count)
            s.Name = i & "_" & tstr
            lr = s.Cells(1, 1).End(xlDown).Row
            Set rg = s.Range(s.Cells(2, 1), s.Cells(lr, 1))
            rg.Hyperlinks.Delete
            s.Rows(1).EntireRow.Insert
            s.Cells(1, 1) = "Open file: " & wbB.Name
            s.Hyperlinks.Add anchor:=s.Cells(1, 1), Address:=wbB.path & "\" & wbB.Name
        End If
        wbB.Close savechanges:=False
    Next i
    Application.DisplayAlerts = False
    wbA.Sheets(1).Delete
    wbA.Sheets(1).Activate
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Call GSPR_Summary_of_summaries_Sharpe
    Application.ScreenUpdating = True
    MsgBox "Done. Please, save """ & wbA.Name & """ as needed.", , "GetStats Pro"
End Sub
Sub Params_To_Summary_Sharpe(ByRef wb As Workbook)
    Const parFRow As Integer = 23
    Dim parLRow As Integer
    Dim i As Integer, j  As Integer, k As Integer, m As Integer
    Dim wsRes As Worksheet, ws As Worksheet
    Dim cRes As Range, c As Range
    Dim clz As Range
    Dim repNum As Integer
    
' copy param names
    Set clz = wb.Sheets(3).Cells
    Set wsRes = wb.Sheets(2)
    Set cRes = wsRes.Cells
    
    parLRow = clz(parFRow, 1).End(xlDown).Row
    j = 2
    cRes(1, 1) = "#_link"
    For i = parFRow To parLRow
        cRes(1, j) = clz(i, 1)
        j = j + 1
    Next i
' copy parameters
    For i = 3 To Sheets.count
        repNum = i - 2
        j = i - 1
        m = 2
        Set ws = Sheets(i)
        Set c = ws.Cells
        cRes(i - 1, 1) = repNum
        For k = parFRow To parLRow
            cRes(j, m) = c(k, 2)
            m = m + 1
        Next k
        ' Add hyperlink to report sheet
        wsRes.Hyperlinks.Add anchor:=cRes(j, 1), Address:="", SubAddress:="'" & repNum & "'!R22C2"
        ' print "back to summary" link
        With c(22, 2)
            .Value = "results"
            .HorizontalAlignment = xlRight
        End With
        ws.Hyperlinks.Add anchor:=c(22, 2), Address:="", SubAddress:="'results'!A" & j
    Next i
    wsRes.Activate
    cRes(2, 2).Activate
    wsRes.Rows("1:1").AutoFilter
    ActiveWindow.FreezePanes = True
End Sub
Sub SharpeBeforeMerge(ByRef wb As Workbook)
    Dim i As Integer
    Dim ws As Worksheet
    Dim c As Range, cSh As Range
    Dim new_col As Integer
    
    Set ws = wb.Sheets(2)
    Set c = ws.Cells
    new_col = c(1, 1).End(xlToRight).Column + 1
    c(1, new_col) = "sharpe_ratio"
    For i = 3 To wb.Sheets.count
        wb.Sheets(i).Activate
        Set cSh = wb.Sheets(i).Cells
        Call Calc_Sharpe_Ratio_Sheet
        
        With c(i - 1, new_col)
            .Value = cSh(21, 2)
            .NumberFormat = "0.00"
        End With
    Next i
    wb.Sheets(2).Activate
    Rows(1).AutoFilter
    Rows(1).AutoFilter
End Sub
Sub Calc_Sharpe_Ratio_Sheet()
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
            On Error Resume Next
            cagr = (1 + net_return) ^ (365 / days_count) - 1
            If Err.Number = 5 Then
                cagr = 0
            End If
            On Error GoTo 0

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

Sub Create_WB_N_Sheets(ByRef newWB As Workbook, _
                       ByVal newSheetsCount As Integer)
' Create new Workbook with new sheets count.
' Parameters:
' newWB             new workbook
' newSheetsCount    new number of sheets in new wb
    Dim origSheetsCount As Integer
    
    origSheetsCount = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = newSheetsCount
    Set newWB = Workbooks.Add
    Application.SheetsInNewWorkbook = origSheetsCount
End Sub

Private Sub GSPRM_Multiple_Main()
'
' RIBBON > BUTTON "Группа"
'
'    On Error Resume Next
    Application.ScreenUpdating = False
' separator ON
    Call GSPR_Separator_Auto_Switcher_Multiple
' Prepare arrays and formats
    Call GSPRM_Prepare_sv_ov_fm
' Pick many reports
    open_fail = False
    Call GSPRM_Open_Reports
    If open_fail Then
        Call GSPR_Separator_Undo_Multiple
        Exit Sub
    End If
' Process one report and print
    Call GSPRM_Process_Each_Print
' check window
    Call GSPR_Check_Window
' save
    Call GSPRM_Save_To_Desktop
'    mb.Sheets(1).Activate
' separator OFF
    Call GSPR_Separator_Undo_Multiple
    Application.ScreenUpdating = True
End Sub
Private Sub GSPR_Separator_Auto_Switcher_Multiple()
    undo_sep = False
    undo_usesyst = False
    If Application.UseSystemSeparators Then     ' SYS - ON
        If Not Application.International(xlDecimalSeparator) = "." Then
            Application.UseSystemSeparators = False
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            undo_sep = True                     ' undo condition 1
            undo_usesyst = True                 ' undo condition 2
        End If
    Else                                        ' SYS - OFF
        If Not Application.DecimalSeparator = "." Then
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            undo_sep = True                     ' undo condition 1
            undo_usesyst = False                ' undo condition 2
        End If
    End If
End Sub
Private Sub GSPR_Separator_Undo_Multiple()
    If undo_sep Then
        Application.DecimalSeparator = current_decimal
        If undo_usesyst Then
            Application.UseSystemSeparators = True
        End If
    End If
End Sub
Private Sub GSPRM_Prepare_sv_ov_fm()
    s_strat = 1
    s_ins = s_strat + 1
    s_tpm = s_ins + 1
    s_ar = s_tpm + 1
    s_mdd = s_ar + 1
    s_rf = s_mdd + 1
    s_rsq = s_rf + 1
    s_date_begin = s_rsq + 1
    s_date_end = s_date_begin + 1
    s_mns = s_date_end + 1
    s_trades = s_mns + 1
    s_win_pc = s_trades + 1
    s_pips = s_win_pc + 1
    s_avg_w2l = s_pips + 1
    s_avg_pip = s_avg_w2l + 1
    s_depo_ini = s_avg_pip + 1
    s_depo_fin = s_depo_ini + 1
    s_cmsn = s_depo_fin + 1
    s_link = s_cmsn + 1
    s_rep_type = s_link + 1
    ReDim SV(1 To s_rep_type, 1 To 2)
' SN
    SV(s_strat, 1) = "Стратегия"
    SV(s_ins, 1) = "Инструмент"
    SV(s_tpm, 1) = "Сделок в месяц"
    SV(s_ar, 1) = "Годовой прирост, %"
    SV(s_mdd, 1) = "Максимальная просадка, %"
    SV(s_rf, 1) = "Коэффициент восстановления"
    SV(s_rsq, 1) = "R-квадрат"
    SV(s_date_begin, 1) = "Начало теста"
    SV(s_date_end, 1) = "Конец теста"
    SV(s_mns, 1) = "Месяцев"
    SV(s_trades, 1) = "Сделок"
    SV(s_win_pc, 1) = "Прибыльных сделок, %"
    SV(s_pips, 1) = "Пунктов"
    SV(s_avg_w2l, 1) = "Сред.приб/убыт, пп"
    SV(s_avg_pip, 1) = "Средняя сделка, пп"
    SV(s_depo_ini, 1) = "Начальный капитал"
    SV(s_depo_fin, 1) = "Конечный капитал"
    SV(s_cmsn, 1) = "Комиссии"
    SV(s_link, 1) = "Размер отчета (МБ), ссылка"
    SV(s_rep_type, 1) = "Тип отчета"
    SV(s_rep_type, 2) = rep_type
' overview
    s_ov_strat = 1
    s_ov_ins = s_ov_strat + 1
    s_ov_htmls = s_ov_ins + 1
    s_ov_mns = s_ov_htmls + 1
    s_ov_from = s_ov_mns + 1
    s_ov_to = s_ov_from + 1
    s_ov_params = s_ov_to + 1
'    s_ov_params_vbl = s_ov_params + 1
    s_ov_created = s_ov_params + 1
    s_ov_macro_ver = s_ov_created + 1
    ReDim ov(1 To s_ov_macro_ver, 1 To 2)
    ov(s_ov_strat, 1) = "Стратегия"
    ov(s_ov_ins, 1) = "Инструмент"
    ov(s_ov_htmls, 1) = "Обработано отчетов"
    ov(s_ov_mns, 1) = "Истор. окно, месяцев"
    ov(s_ov_from, 1) = "Начало теста"
    ov(s_ov_to, 1) = "Конец теста"
    ov(s_ov_params, 1) = "Параметров робота"
'    ov(s_ov_params_vbl, 1) = "Parameters variable"
    ov(s_ov_created, 1) = "Отчет создан"
    ov(s_ov_macro_ver, 1) = "Версия"
' formats
    ' "yyyy-mm-dd"
    fm_date(1) = s_date_begin
    fm_date(2) = s_date_end
    ' "0.00"
    fm_0p00(1) = s_tpm
    fm_0p00(2) = s_rf
    fm_0p00(3) = s_rsq
    fm_0p00(4) = s_mns
    fm_0p00(5) = s_avg_w2l
    fm_0p00(6) = s_avg_pip
    fm_0p00(7) = s_depo_ini
    fm_0p00(8) = s_depo_fin
    fm_0p00(9) = s_cmsn
    fm_0p00(10) = s_link
    ' "0.00%"
    fm_0p00pc(1) = s_ar
    fm_0p00pc(2) = s_mdd
    fm_0p00pc(3) = s_win_pc
    ' color, bold, center
    fm_clr(1) = s_tpm
    fm_clr(2) = s_ar
    fm_clr(3) = s_mdd
    fm_clr(4) = s_rf
    fm_clr(5) = s_rsq
End Sub
Private Sub GSPRM_Open_Reports()
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "GetStats: Выбрать HTML отчеты (максимум " & max_htmls & ")"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Отчеты оптимизатора JForex", "*.html"
        .ButtonName = "Вперед"
    End With
    If fd.Show = 0 Then
        open_fail = True
        MsgBox "No files picked!"
        Exit Sub
    End If
    ov(s_ov_htmls, 2) = fd.SelectedItems.count
    If ov(s_ov_htmls, 2) > max_htmls Then
        MsgBox "GetStats не может обработать более " & max_htmls & " отчетов. Отмена."
        open_fail = True
        Exit Sub
    End If
    ov(s_ov_htmls, 2) = fd.SelectedItems.count
    Folder_To_Save = GSPRM_Folder_To_Save(fd.SelectedItems(1))
    ReDim sM(0 To ov(s_ov_htmls, 2), 0 To 7)
End Sub
Private Function GSPRM_Folder_To_Save(ByVal file_path As String) As String
    Dim q As Integer, i As Integer
    
    q = 0
    i = Len(file_path) + 1
    Do Until q = 2
        i = i - 1
        If Mid(file_path, i, 1) = "\" Then
            q = q + 1
        End If
    Loop
    GSPRM_Folder_To_Save = Left(file_path, i)
End Function
Private Sub GSPRM_Process_Each_Print()
' On Error Resume Next
    Dim i As Integer
    Dim rb As Workbook
    Dim os As Worksheet, hs As Worksheet   ' report sheet, overview sheet, html-processed sheet
    Dim ss As Worksheet                    ' summary sheet
    Dim time_started As Double
    Dim time_now As Double
    Const timer_step As Integer = 5
    Dim time_remaining As Double
    Dim rem_min As Integer, rem_sec As Integer, rem_sec_s As String
    Dim time_rem As String
    Dim counter_timer As Integer
    Dim sta As String
    
' create book
    Set mb = Workbooks.Add
    If mb.Sheets.count > 2 Then
        For i = 1 To mb.Sheets.count - 2
            Application.DisplayAlerts = False
                Sheets(mb.Sheets.count).Delete
            Application.DisplayAlerts = True
        Next i
    ElseIf mb.Sheets.count < 2 Then
        mb.Sheets.Add after:=mb.Sheets(mb.Sheets.count)
    End If
    Set os = mb.Sheets(1)
    os.Name = "сводка"
    Set ss = mb.Sheets(2)
    ss.Name = "результаты"
' open and process each html-report
    time_started = Now
    counter_timer = 0
    For i = 1 To ov(s_ov_htmls, 2)
        counter_timer = counter_timer + 1
        If counter_timer = timer_step Then
            counter_timer = 0
            time_now = Now
            time_remaining = Round(((ov(s_ov_htmls, 2) - i) / i) * (time_now - time_started) * 1440, 2)
            rem_min = Int(time_remaining)
            rem_sec = Int((time_remaining - Int(time_remaining)) * 60)
            rem_sec_s = rem_sec
            If rem_sec < 10 Then
                rem_sec_s = "0" & rem_sec
            End If
            time_rem = rem_min & ":" & rem_sec_s
        End If
        If i < timer_step Then
            sta = "Обрабатываю отчет " & i & " (" & ov(s_ov_htmls, 2) & ")."
        Else
            sta = "Обрабатываю отчет " & i & " (" & ov(s_ov_htmls, 2) & "). Осталось времени " & time_rem
        End If
        Application.StatusBar = sta
        Set rb = Workbooks.Open(fd.SelectedItems(i))
        Set hs = mb.Sheets.Add(after:=mb.Sheets(mb.Sheets.count))
        Select Case i
            Case Is < 10
                hs.Name = "00" & i
            Case 10 To 99
                hs.Name = "0" & i
            Case Else
                hs.Name = i
        End Select
        ' extract statistics
        all_zeros = False
        Call GSPRM_Proc_Extract_stats(rb, i)
        ' print statistics on sheet
        Call GSPRM_Proc_Print_stats(hs, i)
    Next i
    Application.StatusBar = False
' Overview: extract stats and print
    Call GSPRM_Overview_Summary_Extract_Print(os, ss)
End Sub
Private Sub GSPRM_Proc_Extract_stats(ByRef rb As Workbook, ByRef i As Integer)
    Dim used_inss As Integer, ins_td_r As Integer
    Dim j As Integer, k As Integer, l As Integer
    Dim p_fr As Integer, p_lr As Integer
    Dim s As String, ch As String
    Dim rc As Range
    
    Set rc = rb.Sheets(1).Cells
    s = rc(3, 1).Value
' Test end
    SV(s_date_end, 2) = CDate(Left(Right(s, 19), 10))
' get strategy name
    j = InStr(1, s, " strategy report for", 1)
    SV(s_strat, 2) = Left(s, j - 1)
' get trades count
    k = InStr(j, s, " instrument(s) from", 1)
    ' calculate number of used instruments
    used_inss = 0
    For l = j To k
        ch = Mid(s, l, 1)
        If ch = "," Then
            used_inss = used_inss + 1
        End If
    Next l
    used_inss = used_inss + 1
    ' find relevant instrument, with trades
    ins_td_r = 10
    For j = 1 To used_inss
        ins_td_r = rc.Find(what:="Closed positions", after:=rc(ins_td_r, 1), LookIn:=xlValues, LookAt _
            :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
            False, searchformat:=False).Row
        If rc(ins_td_r, 2) <> 0 Then
            Exit For        ' found instrument with trades
        End If
    Next j
' get parameters - Par()
    ' get parameters first & last row
    p_fr = 12
    p_lr = rc(12, 1).End(xlDown).Row
    ov(s_ov_params, 2) = p_lr - p_fr + 1
    ReDim Par(1 To ov(s_ov_params, 2), 1 To 2)
    ' fill Par
    For j = LBound(Par, 1) To UBound(Par, 1)
        For k = LBound(Par, 2) To UBound(Par, 2)
            Par(j, k) = rc(p_fr - 1 + j, k)
        Next k
    Next j
'     sort parameters alphabetically
    Call GSPRM_Par_Bubblesort
    If i = 1 Then
        ReDim par_sum(1 To ov(s_ov_htmls, 2), 1 To UBound(Par, 1))
        ReDim par_sum_head(1 To UBound(Par, 1))
        For j = LBound(par_sum_head) To UBound(par_sum_head)
            par_sum_head(j) = Par(j, 1)
        Next j
    End If
    For j = LBound(par_sum, 2) To UBound(par_sum, 2)
        par_sum(i, j) = Par(j, 2)
    Next j
' File size and link
    SV(s_link, 2) = Round(FileLen(fd.SelectedItems(i)) / 1024 ^ 2, 2)
    
' trades closed
    SV(s_trades, 2) = rc(ins_td_r, 2)
' get instrument
    s = rc(ins_td_r - 9, 1)
    j = InStr(1, s, " ", 1)
    s = Right(s, Len(s) - j)
    SV(s_ins, 2) = s

' Test begin
    SV(s_date_begin, 2) = CDate(Int(rc(ins_td_r - 7, 2))) ' *! new cdate
' Test end
    SV(s_date_end, 2) = CDate(Int(rc(ins_td_r - 4, 2).Value))    ' *! removed "-1"
' Months
    SV(s_mns, 2) = (SV(s_date_end, 2) - SV(s_date_begin, 2)) * 12 / 365
' TPM
'    If rc(ins_td_r, 2) = 0 Then
'        SV(s_tpm, 2) = 0
'    Else
        SV(s_tpm, 2) = Round(SV(s_trades, 2) / SV(s_mns, 2), 2)
'    End If
' Initial deposit
    SV(s_depo_ini, 2) = rc(5, 2)
'' Finish deposit
'    sv(s_depo_fin, 2) = CDbl(rc(6, 2))
' Commissions
    SV(s_cmsn, 2) = rc(8, 2)
    
    If rc(ins_td_r, 2) = 0 Then
        all_zeros = True
        sM(i, 0) = fd.SelectedItems(i)
        sM(i, 1) = i
        sM(i, 2) = 0
        sM(i, 3) = 0
        sM(i, 4) = 0
        sM(i, 5) = 0
        sM(i, 6) = 0
        sM(i, 7) = 0
        rb.Close savechanges:=False
        Exit Sub
    End If
    
    Call GSPRM_Fill_Tradelogs(rc, ins_td_r)
' fill summary stats
    sM(i, 0) = fd.SelectedItems(i)
    sM(i, 1) = i
    sM(i, 2) = SV(s_tpm, 2)
    sM(i, 3) = SV(s_ar, 2)
    sM(i, 4) = SV(s_mdd, 2)
    sM(i, 5) = SV(s_rf, 2)
    sM(i, 6) = SV(s_rsq, 2)
    sM(i, 7) = SV(s_avg_pip, 2)
    rb.Close savechanges:=False
End Sub
Private Sub GSPRM_Par_Bubblesort()
    Dim sj As Integer, sk As Integer
    Dim tmp_c1 As Variant, tmp_c2 As Variant
    
    For sj = 1 To UBound(Par, 1) - 1
        For sk = sj + 1 To UBound(Par, 1)
            If Par(sj, 1) > Par(sk, 1) Then
                tmp_c1 = Par(sk, 1)
                tmp_c2 = Par(sk, 2)
                Par(sk, 1) = Par(sj, 1)
                Par(sk, 2) = Par(sj, 2)
                Par(sj, 1) = tmp_c1
                Par(sj, 2) = tmp_c2
            End If
        Next sk
    Next sj
End Sub
Private Sub GSPRM_Fill_Tradelogs(ByRef rc As Range, ByRef ins_td_r As Integer)
    Dim r As Integer, c As Integer, k As Integer
    Dim oc_fr As Integer
    Dim oc_lr As Long, ro_d As Long
    Dim tl_r As Integer, cds As Integer
    Dim win_sum As Double, los_sum As Double
    Dim win_ct As Integer, los_ct As Integer
    Dim rsqX() As Double, rsqY() As Double
    Dim s As String
    Dim ArrCommis As Variant
    
' get trade log first row - header
    tl_r = rc.Find(what:="Closed orders:", after:=rc(ins_td_r, 1), LookIn:=xlValues, LookAt _
        :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
        False, searchformat:=False).Row + 2
' FILL T1
    ReDim t1(0 To SV(s_trades, 2), 1 To 13)
    t1(0, 10) = SV(s_depo_ini, 2)
    t1(0, 11) = "pc_chg"
    t1(0, 12) = "hwm"
    t1(0, 13) = "dd"
    SV(s_pips, 2) = 0
    win_sum = 0
    los_sum = 0
    win_ct = 0
    los_ct = 0
    For r = LBound(t1, 1) To UBound(t1, 1)
        For c = LBound(t1, 2) To UBound(t1, 2) - 4      ' to 9th column included
            If r > 0 Then
                Select Case c
                    Case Is = 1
                        t1(r, c) = Replace(rc(tl_r + r, c + 1), ",", ".", 1, 1, 1)
                    Case Is = 6
                        t1(r, c) = rc(tl_r + r, c + 1)
                        If t1(r, c) > 0 Then
                            win_ct = win_ct + 1
                            win_sum = win_sum + t1(r, c)
                        Else
                            los_ct = los_ct + 1
                            los_sum = los_sum + t1(r, c)
                        End If
                        SV(s_pips, 2) = SV(s_pips, 2) + t1(r, c)
                    Case Else
                        t1(r, c) = rc(tl_r + r, c + 1)
                End Select
            Else
                t1(r, c) = rc(tl_r + r, c + 1)
            End If
        Next c
    Next r
' winning trades, %
    SV(s_win_pc, 2) = win_ct / SV(s_trades, 2)
' average winner / average loser
    If win_ct = 0 Or los_ct = 0 Then
        SV(s_avg_w2l, 2) = 0
    Else
        SV(s_avg_w2l, 2) = Abs((win_sum / win_ct) / (los_sum / los_ct))
    End If
' average trade, pips
    SV(s_avg_pip, 2) = SV(s_pips, 2) / SV(s_trades, 2)
' FILL T2
    cds = SV(s_date_end, 2) - SV(s_date_begin, 2) + 2
    ReDim t2(0 To cds, 1 To 5)
' 1) date, 2) EOD fin.res., 3) trades closed this day, 4) sum cmsn, 5) avg cmsn
    t2(0, 1) = "date"
    t2(0, 2) = "EOD_fin_res"
    t2(0, 3) = "days_cmsn"
    t2(0, 4) = "tds_closed"
    For r = 1 To UBound(t2, 1)
        t2(r, 4) = 0
    Next r
    t2(0, 5) = "tds_avg_cmsn"
' fill dates
    t2(1, 1) = SV(s_date_begin, 2) - 1
    For r = 2 To UBound(t2, 1)
        t2(r, 1) = t2(r - 1, 1) + 1
    Next r
' fill returns
    t2(1, 2) = SV(s_depo_ini, 2)
' fill sum_cmsn
    oc_fr = rc.Find(what:="Event log:", after:=rc(ins_td_r + SV(s_trades, 2), 1), LookIn:=xlValues, LookAt _
        :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
        False, searchformat:=False).Row + 2 ' header row
    oc_lr = rc(oc_fr, 1).End(xlDown).Row
'
    ro_d = 1
    For r = 2 To UBound(t2, 1)
' compare dates
        If t2(r, 1) = Int(CDate(rc(oc_fr + ro_d, 1))) Then
            Do While t2(r, 1) = Int(CDate(rc(oc_fr + ro_d, 1)))
                If Left(rc(oc_fr + ro_d, 2), 6) = "Commis" Then
                    s = rc(oc_fr + ro_d, 3)
                    ArrCommis = Split(s, " ")
                    s = ArrCommis(UBound(ArrCommis))
                    s = Replace(s, ".", ",")
                    s = Replace(s, "[", "")
                    s = Replace(s, "]", "")
                    If CDbl(s) <> 0 Then
                        t2(r, 3) = CDbl(s)              ' enter cmsn
                    End If
                End If
                ' go to next trade in Trade log
                If ro_d + oc_fr < oc_lr Then
                    ro_d = ro_d + 1
                Else
                    Exit Do ' or exit do, when after last row processed
                End If
            Loop
        End If
    Next r
' calculate trades closed per each day
    c = 1   ' count trades for that day
    k = 1
    For r = 1 To UBound(t1, 1)
        If t1(r, 8) = "ERROR" Then  ' clean ERROR data
            t1(r, 4) = t1(r, 3)
            t1(r, 5) = 0
            t1(r, 6) = 0
            t1(r, 8) = t1(r, 7)
            t1(r, 9) = "ERROR"
        End If
        ' find this day in t2
        Do Until Int(CDate(t2(k, 1))) = Int(CDate(t1(r, 8)))
            k = k + 1
        Loop
        t2(k, 4) = t2(k, 4) + 1
    Next r
' calculate average commission for a trade
    c = 1
    For r = 1 To UBound(t2, 1)
        If t2(r, 4) > 0 Then
            t2(r, 5) = 0
            ' sum cmsns
            For k = c To r
                t2(r, 5) = t2(r, 5) + t2(k, 3)
            Next k
            t2(r, 5) = t2(r, 5) / t2(r, 4)
            c = r + 1
        End If
    Next r
' fill t1 - equity curve
    k = 1
    For r = 1 To UBound(t1, 1)
        Do Until t2(k, 1) = Int(CDate(t1(r, 8)))
            k = k + 1
        Loop
        t1(r, 10) = t1(r - 1, 10) + t1(r, 5) - t2(k, 5)
        ' percent change
        t1(r, 11) = (t1(r, 10) - t1(r - 1, 10)) / t1(r - 1, 10)
    Next r
' hwm & mdd - high watermark & maximum drawdown
    t1(1, 12) = WorksheetFunction.Max(t1(0, 10), t1(1, 10))
    t1(1, 13) = (t1(1, 12) - t1(1, 10)) / t1(1, 12)
    SV(s_mdd, 2) = 0
    For r = 2 To UBound(t1, 1)
        t1(r, 12) = WorksheetFunction.Max(t1(r - 1, 12), t1(r, 10))
        t1(r, 13) = (t1(r, 12) - t1(r, 10)) / t1(r, 12)
        ' update mdd
        If t1(r, 13) > SV(s_mdd, 2) Then
            SV(s_mdd, 2) = t1(r, 13)
        End If
    Next r
' finish deposit
    SV(s_depo_fin, 2) = t1(UBound(t1, 1), 10)
    If SV(s_depo_fin, 2) < 0 Then
        SV(s_depo_fin, 2) = 0
    End If
' annual return = (1 + sv(i, s_depo_fin))
    SV(s_ar, 2) = (1 + (SV(s_depo_fin, 2) - SV(s_depo_ini, 2)) / SV(s_depo_ini, 2)) ^ (12 / SV(s_mns, 2)) - 1
' rf - recovery factor: annualized return / mdd
    If SV(s_ar, 2) > 0 Then
        If SV(s_mdd, 2) > 0 Then
            SV(s_rf, 2) = SV(s_ar, 2) / SV(s_mdd, 2)
        Else
            SV(s_rf, 2) = 999
        End If
    Else
        SV(s_rf, 2) = 0
    End If
' fill t2 EOD fin_res
    k = 1
    For r = 1 To UBound(t1, 1)
        If Int(CDate(t1(r, 8))) > t2(k, 1) Then
            Do Until t2(k, 1) = Int(CDate(t1(r, 8)))
                k = k + 1
                t2(k, 2) = t2(k - 1, 2)
            Loop
            t2(k, 2) = t1(r, 10)
        Else
            t2(k, 2) = t1(r, 10)
        End If
    Next r
    ' fill rest of empty days
    For r = k To UBound(t2, 1)
        If IsEmpty(t2(r, 2)) Then
            t2(r, 2) = t2(r - 1, 2)
        End If
    Next r
' r-square - rsq
    ReDim rsqX(1 To UBound(t2, 1))  ' the x-s and the y-s
    ReDim rsqY(1 To UBound(t2, 1))
    For r = 1 To UBound(t2, 1)
        rsqX(r) = t2(r, 1)
        rsqY(r) = t2(r, 2)
    Next r
    SV(s_rsq, 2) = WorksheetFunction.RSq(rsqX, rsqY)
End Sub
Private Sub GSPRM_Proc_Print_stats(ByRef hs As Worksheet, ByRef i As Integer)
    Dim r As Integer, c As Integer
    Dim hc As Range
    Dim z As Variant
    
    z = Array(3, 4, 5, 6, 7, 11, 12, 13, 14, 17)
    Set hc = hs.Cells
    If all_zeros Then
        For r = LBound(z) To UBound(z)
            SV(z(r), 2) = 0
        Next r
    End If
' print statistics names and values
    For r = LBound(SV, 1) To UBound(SV, 1)
        For c = LBound(SV, 2) To UBound(SV, 2)
            hc(r, c) = SV(r, c)
        Next c
    Next r
' print parameters
    hc(UBound(SV) + 2, 1) = "Параметры"
    For r = LBound(Par, 1) To UBound(Par, 1)
        For c = LBound(Par, 2) To UBound(Par, 2)
            hc(UBound(SV) + 2 + r, c) = Par(r, c)
        Next c
    Next r
' print "back to summary" link
    With hc(UBound(SV) + 2, 2)
        .Value = "результаты"
        .HorizontalAlignment = xlRight
    End With
    hs.Hyperlinks.Add anchor:=hc(UBound(SV) + 2, 2), Address:="", SubAddress:="'результаты'!A1"

    If all_zeros = False Then
    ' print tradelog-1
        ReDim Preserve t1(0 To UBound(t1, 1), 1 To 11)
        For r = LBound(t1, 1) To UBound(t1, 1)
            For c = LBound(t1, 2) To UBound(t1, 2)
                hc(r + 1, c + 2) = t1(r, c)
            Next c
        Next r
    ' print tradelog-2
        ReDim Preserve t2(0 To UBound(t2, 1), 1 To 2)
        For r = LBound(t2, 1) To UBound(t2, 1)
            For c = LBound(t2, 2) To UBound(t2, 2)
                hc(r + 1, c + UBound(t1, 2) + 2) = t2(r, c)
            Next c
        Next r
    End If
' apply formats
    For r = LBound(fm_date) To UBound(fm_date)
        hc(fm_date(r), 2).NumberFormat = "yyyy-mm-dd"
    Next r
    For r = LBound(fm_0p00) To UBound(fm_0p00)
        hc(fm_0p00(r), 2).NumberFormat = "0.00"
    Next r
    For r = LBound(fm_0p00pc) To UBound(fm_0p00pc)
        hc(fm_0p00pc(r), 2).NumberFormat = "0.00%"
    Next r
    For r = LBound(fm_clr) To UBound(fm_clr)
        For c = 1 To 2
            With hc(fm_clr(r), c)
                .Interior.Color = RGB(184, 204, 228)
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
            End With
        Next c
    Next r
' hyperlink
    hs.Hyperlinks.Add anchor:=hc(s_link, 2), Address:=sM(i, 0)
    hs.Activate
    hs.Range(Columns(1), Columns(2)).AutoFit
End Sub
Private Sub GSPRM_Overview_Summary_Extract_Print(ByRef os As Worksheet, ByRef ss As Worksheet)
    Dim r As Integer, c As Integer, hr As Integer
    Dim oc As Range, sc As Range
    Dim Rng As Range
    
    Set oc = os.Cells
    Set sc = ss.Cells
' extract from last report
    ov(s_ov_strat, 2) = SV(s_strat, 2)
    ov(s_ov_ins, 2) = SV(s_ins, 2)
    ov(s_ov_mns, 2) = SV(s_mns, 2)
    ov(s_ov_from, 2) = SV(s_date_begin, 2)
    ov(s_ov_to, 2) = SV(s_date_end, 2)
    ov(s_ov_created, 2) = Now
    ov(s_ov_macro_ver, 2) = macro_ver
' print OVERVIEW
    For r = LBound(ov, 1) To UBound(ov, 1)
        For c = LBound(ov, 2) To UBound(ov, 2)
            oc(r, c) = ov(r, c)
        Next c
    Next r
' apply formats
    oc(4, 2).NumberFormat = "0.00"
    oc(5, 2).NumberFormat = "yyyy-mm-dd"
    oc(6, 2).NumberFormat = "yyyy-mm-dd"
    oc(9, 2).NumberFormat = "yyyy-mm-dd, hh:mm"
    os.Columns(1).AutoFit
    os.Columns(2).AutoFit
' fill summary header
    sM(0, 0) = "html_link"
    sM(0, 1) = "№_ссылка"
    sM(0, 2) = "сделок_мес"
    sM(0, 3) = "год_прир"
    sM(0, 4) = "макс_прос"
    sM(0, 5) = "восст"
    sM(0, 6) = "r_кв"
    sM(0, 7) = "сред_сд_пп"
' print SUMMARY
    For r = LBound(sM, 1) To UBound(sM, 1)
        For c = 1 To UBound(sM, 2)
            sc(r + 1, c) = sM(r, c)
        Next c
    Next r
    ' parameters - head
    For c = LBound(par_sum_head) To UBound(par_sum_head)
        sc(1, c + UBound(sM, 2)) = par_sum_head(c)
    Next c
    ' parameters - values
    For r = LBound(par_sum, 1) To UBound(par_sum, 1)
        For c = LBound(par_sum, 2) To UBound(par_sum, 2)
            sc(r + 1, c + UBound(sM, 2)) = par_sum(r, c)
        Next c
    Next r
' apply formats
    ' 0.00%
    Set Rng = Range(sc(2, 3), sc(1 + UBound(sM, 1), 4))
    Rng.NumberFormat = "0.00%"
    ' 0.00
    Set Rng = Range(sc(2, 5), sc(1 + UBound(sM, 1), 7))
    Rng.NumberFormat = "0.00"
    Set Rng = Range(sc(1, 1), sc(1, UBound(sM, 2)))
    Rng.Font.Bold = True
' hyperlinks
    hr = UBound(SV) + 2
    For r = 1 To UBound(sM, 1)
        Select Case r
            Case Is < 10
                ss.Hyperlinks.Add anchor:=sc(r + 1, 1), Address:="", SubAddress:="'00" & r & "'!R" & hr & "C2"
            Case 10 To 99
                ss.Hyperlinks.Add anchor:=sc(r + 1, 1), Address:="", SubAddress:="'0" & r & "'!R" & hr & "C2"
            Case Else
                ss.Hyperlinks.Add anchor:=sc(r + 1, 1), Address:="", SubAddress:="'" & r & "'!R" & hr & "C2"
        End Select
    Next r
' add autofilter
    ss.Activate
    ss.Rows("1:1").AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub
Private Sub GSPRM_Save_To_Desktop()
'    Dim envir As String
    Dim temp_s As String, corenm As String
    Dim dbeg As String, dfin As String
    Dim s_yr As String, s_mn As String, s_dy As String
    Dim vers As Integer, j As Integer
    
    Application.StatusBar = "Сохраняюсь."
'    envir = Environ("UserProfile") & "\Desktop\"
' date begin
    s_yr = Year(SV(s_date_begin, 2))
    s_mn = Month(SV(s_date_begin, 2))
    If Len(s_mn) = 1 Then
        s_mn = "0" & s_mn
    End If
    s_dy = Day(SV(s_date_begin, 2))
    If Len(s_dy) = 1 Then
        s_dy = "0" & s_dy
    End If
    dbeg = s_yr & s_mn & s_dy
' date end
    s_yr = Year(SV(s_date_end, 2))
    s_mn = Month(SV(s_date_end, 2))
    If Len(s_mn) = 1 Then
        s_mn = "0" & s_mn
    End If
    s_dy = Day(SV(s_date_end, 2))
    If Len(s_dy) = 1 Then
        s_dy = "0" & s_dy
    End If
    dfin = s_yr & s_mn & s_dy
' core name (was - envir instead of folder_to_save)
    corenm = Folder_To_Save & SV(s_strat, 2) & "-" & Left(SV(s_ins, 2), 3) & Right(SV(s_ins, 2), 3) & "-" & dbeg & "-" & dfin & "-r" & ov(s_ov_htmls, 2)
    fNm = corenm & ".xlsx"
    If Dir(fNm) = "" Then
        mb.SaveAs fileName:=fNm, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Else
        fNm = corenm & "(2).xlsx"
        If Dir(fNm) <> "" Then
            j = InStr(1, fNm, "(", 1)
            temp_s = Right(fNm, Len(fNm) - j)
            j = InStr(1, temp_s, ")", 1)
            vers = Left(temp_s, j - 1)
            fNm = corenm & "(" & vers & ").xlsx"
            Do Until Dir(fNm) = ""
                vers = vers + 1
                fNm = corenm & "(" & vers & ").xlsx"
            Loop
        End If
        mb.SaveAs fileName:=fNm, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    End If
    addin_c(last_row_reports, 4) = fNm
'    MsgBox "Отчеты успешно обработаны. Файл сохранен на рабочем столе:" & vbNewLine & fnm, , "GetStats Pro"
    mb.Close savechanges:=False
    Application.StatusBar = False
End Sub
Private Sub GSPRM_Merge_Summaries()
'
' RIBBON > BUTTON "Recovery"
'
    Dim sel_count As Integer, i As Integer, lr As Integer
    Dim pos As Integer
    Dim tstr As String
    Dim wbA As Workbook, wbB As Workbook
    Dim s As Worksheet
    Dim Rng As Range
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "GetStats: Выбрать отчеты GetStats (максимум " & max_htmls & ")"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Отчеты надстройки GetStats", "*.xlsx"
        .ButtonName = "Вперед"
    End With
    If fd.Show = 0 Then
        MsgBox "Файлы не выбраны!"
        Exit Sub
    End If
    sel_count = fd.SelectedItems.count
    If sel_count > max_htmls Then
        MsgBox "GetStats не может обработать более " & max_htmls & " отчетов. Отмена."
        Exit Sub
    End If
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
        Application.StatusBar = "Adding sheet " & i & " (" & sel_count & ")."
        Set wbB = Workbooks.Open(fd.SelectedItems(i))
        tstr = wbB.Name
        pos = InStr(1, tstr, "-", 1)
        tstr = Right(Left(tstr, pos + 6), 6)
        If wbB.Sheets(2).Name = "results" Then
            wbB.Sheets("results").Copy after:=wbA.Sheets(wbA.Sheets.count)
            Set s = wbA.Sheets(wbA.Sheets.count)
            s.Name = i & "_" & tstr
            lr = s.Cells(1, 1).End(xlDown).Row
            Set Rng = s.Range(s.Cells(2, 1), s.Cells(lr, 1))
            Rng.Hyperlinks.Delete
            s.Rows(1).EntireRow.Insert
            s.Cells(1, 1) = "Открыть файл: " & wbB.Name
            s.Hyperlinks.Add anchor:=s.Cells(1, 1), Address:=wbB.path & "\" & wbB.Name
        End If
        wbB.Close savechanges:=False
    Next i
    Application.DisplayAlerts = False
    wbA.Sheets(1).Delete
    wbA.Sheets(1).Activate
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Call GSPR_Summary_of_summaries
    Application.ScreenUpdating = True
    MsgBox "Готово. Сохраните файл """ & wbA.Name & """ по вашему усмотрению.", , "GetStats Pro"
End Sub
Private Sub GSPR_Summary_of_summaries()
    Const rf_0 As Double = 0.5
    Const rf_1 As Double = 1
    Const rf_2 As Double = 1.5
    Const rf_3 As Double = 2
    Const rf_4 As Double = 3
    Dim i As Integer, lr As Integer
    Dim ws As Worksheet, iter_s As Worksheet
    Dim sc As Range, iter_c As Range, c_rng As Range, cell As Range
    
    Set ws = Sheets.Add(before:=Sheets(1))
    ws.Name = "коэфф_восст"
    Set sc = ws.Cells
    For i = 2 To Sheets.count
        Set iter_s = Sheets(i)
        Set iter_c = iter_s.Cells
        lr = iter_c(3, 1).End(xlDown).Row
        Set c_rng = iter_s.Range(iter_c(3, 5), iter_c(lr, 5))
        c_rng.Copy (sc(2, i - 1))
        Set c_rng = ws.Range(sc(2, i - 1), sc(lr - 1, i - 1))
        c_rng.Sort key1:=sc(2, i - 1), order1:=xlDescending, Header:=xlNo
        For Each cell In c_rng
            Select Case cell.Value
                Case rf_0 To rf_1
                    cell.Interior.Color = RGB(160, 255, 160)
                Case rf_1 To rf_2
                    cell.Interior.Color = RGB(0, 255, 0)
                Case rf_2 To rf_3
                    cell.Interior.Color = RGB(0, 220, 0)
                Case rf_3 To rf_4
                    cell.Interior.Color = RGB(0, 190, 0)
                Case Is > rf_4
                    cell.Interior.Color = RGB(0, 140, 0)
            End Select
        Next cell
        ws.Hyperlinks.Add anchor:=sc(1, i - 1), Address:="", SubAddress:="'" & Sheets(i).Name & "'!A1"
        With sc(1, i - 1)
            .Value = Right(iter_s.Name, 6)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
End Sub
Private Sub GSPR_Summary_of_summaries_Sharpe()
    Const sr_0 As Double = 0.075
    Const sr_1 As Double = 0.1
    Const sr_2 As Double = 1.5
    Const sr_3 As Double = 2
    Const sr_4 As Double = 3
    Dim i As Integer, lr As Integer
    Dim ws As Worksheet, iter_s As Worksheet
    Dim sc As Range, iter_c As Range, c_rng As Range, cell As Range
    Dim lastCol As Integer
    
    Set ws = Sheets.Add(before:=Sheets(1))
    ws.Name = "Sharpe_ratio"
    Set sc = ws.Cells
    For i = 2 To Sheets.count
        Set iter_s = Sheets(i)
        Set iter_c = iter_s.Cells
        lr = iter_c(3, 1).End(xlDown).Row
        lastCol = iter_c(2, 1).End(xlToRight).Column
        Set c_rng = iter_s.Range(iter_c(3, lastCol), iter_c(lr, lastCol))
        c_rng.Copy (sc(2, i - 1))
        Set c_rng = ws.Range(sc(2, i - 1), sc(lr - 1, i - 1))
        c_rng.Sort key1:=sc(2, i - 1), order1:=xlDescending, Header:=xlNo
        For Each cell In c_rng
            Select Case cell.Value
                Case sr_0 To sr_1
                    cell.Interior.Color = RGB(160, 255, 160)
                Case sr_1 To sr_2
                    cell.Interior.Color = RGB(0, 255, 0)
                Case sr_2 To sr_3
                    cell.Interior.Color = RGB(0, 220, 0)
                Case sr_3 To sr_4
                    cell.Interior.Color = RGB(0, 190, 0)
                Case Is > sr_4
                    cell.Interior.Color = RGB(0, 140, 0)
            End Select
        Next cell
        ws.Hyperlinks.Add anchor:=sc(1, i - 1), Address:="", SubAddress:="'" & Sheets(i).Name & "'!A1"
        With sc(1, i - 1)
            .Value = Right(iter_s.Name, 6)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
End Sub
Private Sub GSPR_Change_Folder_Link()

    Const err_msg As String = "Не похоже на книгу с отчетами GetStats."
    Dim ws As Worksheet
    Dim sc As Range
    Dim hyperlink_cell_row As Integer
    Dim address_string As String, report_name As String
    Dim len_subtract_current As Integer, len_subtract_new As Integer
    Dim new_prefix As String, new_hyperlink As String
    Dim i As Integer
    
    Application.ScreenUpdating = False
' sanity check
    If Sheets.count < 3 Then
        MsgBox err_msg, , "GetStats Pro"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Set ws = Sheets(3)
    Set sc = ws.Cells
    If sc.Find(what:="Размер отчета (МБ), ссылка") Is Nothing Then
        MsgBox err_msg, , "GetStats Pro"
        Application.ScreenUpdating = True
        Exit Sub
    End If
' end sanity check
    hyperlink_cell_row = sc.Find(what:="Размер отчета (МБ), ссылка").Row
    address_string = sc(hyperlink_cell_row, 2).Hyperlinks(1).Address
    len_subtract_current = Back_Slash_Pos(address_string)   ' call function
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "GetStats: Выбрать новый путь к отчетам"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Отчеты оптимизатора JForex", "*.html"
        .ButtonName = "Выбрать"
    End With
    If fd.Show = 0 Then
        MsgBox "Файл не выбран!", , "GetStats Pro"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    address_string = fd.SelectedItems(1)
    len_subtract_new = Back_Slash_Pos(address_string)   ' call function
    new_prefix = Left(address_string, len_subtract_new)
    For i = 3 To Sheets.count
        Set ws = Sheets(i)
        Set sc = ws.Cells
        address_string = sc(hyperlink_cell_row, 2).Hyperlinks(1).Address
        report_name = Right(address_string, Len(address_string) - len_subtract_current)
        new_hyperlink = new_prefix & report_name
        sc(hyperlink_cell_row, 2).Hyperlinks(1).Address = new_hyperlink
    Next i
    Application.ScreenUpdating = True
    MsgBox "Гиперссылки на html-отчеты обновлены (всего " & Sheets.count - 2 & ").", , "GetStats Pro"
End Sub
Private Sub GSPR_Check_Window()
    Dim win_start As Long
    Dim win_end As Long
    Dim ws As Worksheet
    Dim c As Range
    Dim rng_check As Range, rng_c As Range
    Dim i As Integer
    Dim html_count As Integer
    Dim add_c1 As Integer, add_c2 As Integer, add_c3 As Integer
    Dim dates_ok_counter As Integer

    Set addin_book = Workbooks(addin_file_name)
    Set addin_c = addin_book.Sheets("Settings").Cells
    win_start = addin_c(3, 2)
    win_end = addin_c(4, 2)
    html_count = addin_c(5, 2)
    
    add_c1 = Sheets(2).Cells(1, 1).End(xlToRight).Column + 1
    add_c2 = add_c1 + 1
    add_c3 = add_c2 + 1
    Sheets(2).Cells(1, add_c1) = "start"
    Sheets(2).Cells(1, add_c2) = "end"
    Sheets(2).Cells(1, add_c3) = "depo_ini"
    
    For i = 3 To Sheets.count
        Set c = Sheets(i).Cells
        ' check window start date
        If CLng(c(8, 2)) = win_start Then
            Sheets(2).Cells(i - 1, add_c1) = "ok"
        Else
            With Sheets(2).Cells(i - 1, add_c1)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
        ' check window end date
        If CLng(c(9, 2)) = win_end Then
            Sheets(2).Cells(i - 1, add_c2) = "ok"
'        ElseIf CLng(c(9, 2)) <> win_end And c(5, 2) > 0.9 Then
'            Sheets(2).Cells(i - 1, add_c2) = "ok"
        Else
            With Sheets(2).Cells(i - 1, add_c2)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
        ' check depo_ini
        If CDbl(c(16, 2)) = depo_ini_ok Then
            Sheets(2).Cells(i - 1, add_c3) = "ok"
        Else
            With Sheets(2).Cells(i - 1, add_c3)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
    Next i
    Sheets(2).Activate
    Sheets(2).Rows("1:1").AutoFilter
    Sheets(2).Rows("1:1").AutoFilter
    
    last_row_reports = addin_c(Workbooks(addin_file_name).Sheets("настройки").Rows.count, 4).End(xlUp).Row + 1
    
    Set rng_check = mb.Sheets(2).Range(Cells(2, add_c1), Cells(Sheets.count - 1, add_c2))
    ' result of checking dates, into addin
    For Each rng_c In rng_check
        If rng_c.Value <> "ok" Then
            With addin_c(last_row_reports, 5)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
            Exit For
        End If
    Next rng_c
    If addin_c(last_row_reports, 5).Value <> "error" Then
        addin_c(last_row_reports, 5).Value = "ok"
    End If
    ' result of checking sheets count, into addin
    If mb.Sheets.count - 2 = html_count Then
        addin_c(last_row_reports, 6) = "ok"
    Else
        With addin_c(last_row_reports, 6)
            .Value = "error"
            .Interior.Color = RGB(255, 0, 0)
        End With
    End If
    ' result of checking depo_ini, into addin
    Set rng_check = mb.Sheets(2).Range(Cells(2, add_c3), Cells(Sheets.count - 1, add_c3))
    For Each rng_c In rng_check
        If rng_c.Value <> "ok" Then
            With addin_c(last_row_reports, 7)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
            Exit For
        End If
    Next rng_c
    If addin_c(last_row_reports, 7).Value <> "error" Then
        addin_c(last_row_reports, 7).Value = "ok"
    End If
    ' add timestamp
    addin_c(last_row_reports, 8) = Now
End Sub
Sub Check_Window_Bulk()
' checks a selection of books
' window dates, number of reports
    Dim fd As FileDialog
    Dim win_start As Long
    Dim win_end As Long
    Dim ws As Worksheet
    Dim wbCheck As Workbook
    Dim wsSummary As Worksheet
    Dim c As Range
    Dim rng_check As Range, rng_c As Range
    Dim i As Integer, j As Integer
    Dim html_count As Integer
    Dim add_c1 As Integer, add_c2 As Integer, add_c3 As Integer
    Dim dates_ok_counter As Integer

    Application.ScreenUpdating = False
    
    Set addin_book = Workbooks(addin_file_name)
    Set addin_c = addin_book.Sheets("настройки").Cells
    win_start = addin_c(3, 2)
    win_end = addin_c(4, 2)
    html_count = addin_c(5, 2)
    
' SELECT BOOKS TO CHECK
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "GetStats: Выбрать XLSX для проверки"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Обработанные отчеты", "*.xlsx"
        .ButtonName = "Вперед"
    End With
    If fd.Show = 0 Then
        MsgBox "Файлы не выбраны!"
        Exit Sub
    End If
    
' MAIN LOOP
    For j = 1 To fd.SelectedItems.count
        Application.StatusBar = j & " (" & fd.SelectedItems.count & ")."
        Set wbCheck = Workbooks.Open(fd.SelectedItems(j))
        Set wsSummary = wbCheck.Sheets(2)
        add_c1 = wsSummary.Cells(1, 1).End(xlToRight).Column + 1
        add_c2 = add_c1 + 1
        add_c3 = add_c2 + 1
        wsSummary.Cells(1, add_c1) = "start"
        wsSummary.Cells(1, add_c2) = "end"
        wsSummary.Cells(1, add_c3) = "depo_ini"
        
        For i = 3 To wbCheck.Sheets.count
            Set c = wbCheck.Sheets(i).Cells
            ' check window start date
            If CLng(c(8, 2)) = win_start Then
                wsSummary.Cells(i - 1, add_c1) = "ok"
            Else
                With wsSummary.Cells(i - 1, add_c1)
                    .Value = "error"
                    .Interior.Color = RGB(255, 0, 0)
                End With
            End If
            ' check window end date
            If CLng(c(9, 2)) = win_end Then
                wsSummary.Cells(i - 1, add_c2) = "ok"
            Else
                With wsSummary.Cells(i - 1, add_c2)
                    .Value = "error"
                    .Interior.Color = RGB(255, 0, 0)
                End With
            End If
            ' check depo_ini
            If CDbl(c(16, 2)) = depo_ini_ok Then
                wsSummary.Cells(i - 1, add_c3) = "ok"
            Else
                With wsSummary.Cells(i - 1, add_c3)
                    .Value = "error"
                    .Interior.Color = RGB(255, 0, 0)
                End With
            End If
        Next i
        wsSummary.Activate
        wsSummary.Rows("1:1").AutoFilter
        wsSummary.Rows("1:1").AutoFilter
        
        last_row_reports = addin_c(Workbooks(addin_file_name).Sheets("настройки").Rows.count, 4).End(xlUp).Row + 1
        Set rng_check = wsSummary.Range(wsSummary.Cells(2, add_c1), wsSummary.Cells(wbCheck.Sheets.count - 1, add_c2))
    ' result of checking dates, into addin
        For Each rng_c In rng_check
            If rng_c.Value <> "ok" Then
                With addin_c(last_row_reports, 5)
                    .Value = "error"
                    .Interior.Color = RGB(255, 0, 0)
                End With
                Exit For
            End If
        Next rng_c
        If addin_c(last_row_reports, 5).Value <> "error" Then
            addin_c(last_row_reports, 5).Value = "ok"
        End If
    ' result of checking sheets count, into addin
        If wbCheck.Sheets.count - 2 = html_count Then
            addin_c(last_row_reports, 6) = "ok"
        Else
            With addin_c(last_row_reports, 6)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
    ' result of checking depo_ini, into addin
        Set rng_check = wsSummary.Range(wsSummary.Cells(2, add_c3), wsSummary.Cells(wbCheck.Sheets.count - 1, add_c3))
        For Each rng_c In rng_check
            If rng_c.Value <> "ok" Then
                With addin_c(last_row_reports, 7)
                    .Value = "error"
                    .Interior.Color = RGB(255, 0, 0)
                End With
                Exit For
            End If
        Next rng_c
        If addin_c(last_row_reports, 7).Value <> "error" Then
            addin_c(last_row_reports, 7).Value = "ok"
        End If
    ' add wb.Name and time stamp
        addin_c(last_row_reports, 4) = wbCheck.FullName
        addin_c(last_row_reports, 8) = Now
    ' close and save checked book
        wbCheck.Close savechanges:=True
    Next j
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
Sub Check_Window_Standalone()
'
' RIBBON > BUTTON "Окно"
'
    Dim win_start As Long
    Dim win_end As Long
    Dim ws As Worksheet
    Dim wbCheck As Workbook
    Dim wsSummary As Worksheet
    Dim c As Range
    Dim rng_check As Range, rng_c As Range
    Dim i As Integer
    Dim html_count As Integer
    Dim add_c1 As Integer, add_c2 As Integer
    Dim dates_ok_counter As Integer

    Application.ScreenUpdating = False
    Set wbCheck = ActiveWorkbook
    Set wsSummary = wbCheck.Sheets(2)
    
    Set addin_book = Workbooks(addin_file_name)
    Set addin_c = addin_book.Sheets("настройки").Cells
    win_start = addin_c(3, 2)
    win_end = addin_c(4, 2)
    html_count = addin_c(5, 2)
    
    add_c1 = wsSummary.Cells(1, 1).End(xlToRight).Column + 1
    add_c2 = add_c1 + 1
    wsSummary.Cells(1, add_c1) = "start"
    wsSummary.Cells(1, add_c2) = "end"
    
    For i = 3 To wbCheck.Sheets.count
        Set c = wbCheck.Sheets(i).Cells
        If CLng(c(8, 2)) = win_start Then
            wsSummary.Cells(i - 1, add_c1) = "ok"
        Else
            With wsSummary.Cells(i - 1, add_c1)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
        If CLng(c(9, 2)) = win_end Then
            wsSummary.Cells(i - 1, add_c2) = "ok"
        Else
            With wsSummary.Cells(i - 1, add_c2)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
    Next i
    wsSummary.Activate
    wsSummary.Rows("1:1").AutoFilter
    wsSummary.Rows("1:1").AutoFilter
    
    last_row_reports = addin_c(Workbooks(addin_file_name).Sheets("настройки").Rows.count, 4).End(xlUp).Row + 1
    Set rng_check = wsSummary.Range(wsSummary.Cells(2, add_c1), wsSummary.Cells(wbCheck.Sheets.count - 1, add_c2))
' result of checking dates, into addin
    For Each rng_c In rng_check
        If rng_c.Value <> "ok" Then
            addin_c(last_row_reports, 5) = "error"
            Exit For
        End If
    Next rng_c
    If addin_c(last_row_reports, 5).Value <> "error" Then
        addin_c(last_row_reports, 5).Value = "ok"
    End If
' result of checking sheets count, into addin
    If wbCheck.Sheets.count - 2 = html_count Then
        addin_c(last_row_reports, 6) = "ok"
    Else
        addin_c(last_row_reports, 6) = "error"
    End If
' add wb.Name and time stamp
    addin_c(last_row_reports, 4) = wbCheck.FullName
    addin_c(last_row_reports, 7) = Now
    Application.ScreenUpdating = True
End Sub
Sub Remove_Checks()
    Const lastSearchParam As String = "defaultInstrument"
    Dim fd As FileDialog
    Dim i As Integer
    Dim remCol0 As Integer, remCol9 As Integer
    Dim wbCheck As Workbook
    Dim ws As Worksheet
    Dim c As Range
    Dim Rng As Range
' SELECT BOOKS TO CHECK
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "GetStats: Выбрать XLSX для проверки"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Обработанные отчеты", "*.xlsx"
        .ButtonName = "Вперед"
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
' MAIN LOOP
    Application.ScreenUpdating = False
    For i = 1 To fd.SelectedItems.count
        Application.StatusBar = i & " (" & fd.SelectedItems.count & ")."
        Set wbCheck = Workbooks.Open(fd.SelectedItems(i))
        Set ws = wbCheck.Sheets(2)
        Set c = ws.Cells
        remCol0 = c.Find(what:=lastSearchParam, after:=c(1, 1), LookIn:=xlValues, LookAt _
            :=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:= _
            False, searchformat:=False).Column + 1
        If c(1, remCol0) <> "" Then
            remCol9 = c(1, remCol0).End(xlToRight).Column
            ws.Range(Columns(remCol0), Columns(remCol9)).EntireColumn.Delete
        End If
        wbCheck.Close savechanges:=True
    Next i
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub



' MODULE: Rep_Single
Option Explicit
Option Base 1
    Const rep_type As String = "GS_Pro_Single_Core"
    Dim ch_rep_type As Boolean
    Dim wb As Workbook, ws As Worksheet, wc As Range
    Dim rb As Workbook, rs As Worksheet, rc As Range
    Dim open_fail As Boolean, all_zeros As Boolean
    Dim rep_adr As String
    Dim ins_td_r As Integer
' ARRAYS
    Dim SV() As Variant
    Dim Par() As Variant
    Dim t1() As Variant
    Dim t2() As Variant
    Dim fm_date(1 To 2) As Integer, fm_0p00(1 To 10) As Integer, fm_0p00pc(1 To 3) As Integer, fm_clr(1 To 5) As Integer    ' count before changing
' OBJECTS
    Dim mb As Workbook
'---SHIFTS
'------ sn, sv
    Dim s_strat As Integer, s_ins As Integer
    Dim s_tpm As Integer, s_ar As Integer, s_mdd As Integer, s_rf As Integer, s_rsq As Integer
    Dim s_date_begin As Integer, s_date_end As Integer, s_mns As Integer
    Dim s_trades As Integer, s_win_pc As Integer, s_pips As Integer, s_avg_w2l As Integer, s_avg_pip As Integer
    Dim s_depo_ini As Integer, s_depo_fin As Integer, s_cmsn As Integer, s_link As Integer, s_rep_type As Integer
'------ ov_bas
    Dim s_ov_strat As Integer, s_ov_ins As Integer, s_ov_htmls As Integer
    Dim s_ov_mns As Integer, s_ov_from As Integer, s_ov_to As Integer
    Dim s_ov_params As Integer, s_ov_varied As Integer, s_ov_created
' separator variables
    Dim current_decimal As String
    Dim undo_sep As Boolean, undo_usesyst As Boolean
Private Sub GSPR_Single_Core()
'    Dim oTimer As clsTimer
'
' RIBBON > BUTTON "Main"
'
'    On Error Resume Next
'    Set oTimer = New clsTimer
    
    Application.ScreenUpdating = False
    Call GSPR_Separator_Auto_Switcher_Single_Core
    Call GSPR_Prepare_sv_fm
    Call GSPR_Open_Report
    If open_fail = True Then
        Call GSPR_Separator_Undo_Single_Core
        Exit Sub
    End If
    Call GSPR_Extract_stats
    Call GSPR_Print_stats
    If all_zeros = False Then
        Call GSPR_Build_Charts_Single_Report
    End If
    Call GSPR_Separator_Undo_Single_Core
    Application.ScreenUpdating = True
End Sub
Private Sub GSPR_Separator_Auto_Switcher_Single_Core()
    undo_sep = False
    undo_usesyst = False
    If Application.UseSystemSeparators Then     ' SYS - ON
        If Not Application.International(xlDecimalSeparator) = "." Then
            Application.UseSystemSeparators = False
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            undo_sep = True                     ' undo condition 1
            undo_usesyst = True                 ' undo condition 2
        End If
    Else                                        ' SYS - OFF
        If Not Application.DecimalSeparator = "." Then
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            undo_sep = True                     ' undo condition 1
            undo_usesyst = False                ' undo condition 2
        End If
    End If
End Sub
Private Sub GSPR_Separator_Undo_Single_Core()
    If undo_sep Then
        Application.DecimalSeparator = current_decimal
        If undo_usesyst Then
            Application.UseSystemSeparators = True
        End If
    End If
End Sub
Private Sub GSPR_Prepare_sv_fm()
    s_strat = 1
    s_ins = s_strat + 1
    s_tpm = s_ins + 1
    s_ar = s_tpm + 1
    s_mdd = s_ar + 1
    s_rf = s_mdd + 1
    s_rsq = s_rf + 1
    s_date_begin = s_rsq + 1
    s_date_end = s_date_begin + 1
    s_mns = s_date_end + 1
    s_trades = s_mns + 1
    s_win_pc = s_trades + 1
    s_pips = s_win_pc + 1
    s_avg_w2l = s_pips + 1
    s_avg_pip = s_avg_w2l + 1
    s_depo_ini = s_avg_pip + 1
    s_depo_fin = s_depo_ini + 1
    s_cmsn = s_depo_fin + 1
    s_link = s_cmsn + 1
    s_rep_type = s_link + 1
    ReDim SV(1 To s_rep_type, 1 To 2)
' SV
    SV(s_strat, 1) = "Strategy"
    SV(s_ins, 1) = "Instrument"
    SV(s_tpm, 1) = "Trades per month"
    SV(s_ar, 1) = "Annualized return, %"
    SV(s_mdd, 1) = "Maximum drawdown, %"
    SV(s_rf, 1) = "Recovery factor"
    SV(s_rsq, 1) = "R-squared"
    SV(s_date_begin, 1) = "Test begin date"
    SV(s_date_end, 1) = "Test end date"
    SV(s_mns, 1) = "Months"
    SV(s_trades, 1) = "Positions closed"
    SV(s_win_pc, 1) = "Winners, %"
    SV(s_pips, 1) = "Pips"
    SV(s_avg_w2l, 1) = "Avg. winner/loser, pips"
    SV(s_avg_pip, 1) = "Avg. trade, pips"
    SV(s_depo_ini, 1) = "Initial balance"
    SV(s_depo_fin, 1) = "End balance"
    SV(s_cmsn, 1) = "Commissions"
    SV(s_link, 1) = "Report size (MB), link"
    SV(s_rep_type, 1) = "Report type"
    SV(s_rep_type, 2) = rep_type
' formats
    ' "yyyy-mm-dd"
    fm_date(1) = s_date_begin
    fm_date(2) = s_date_end
    ' "0.00"
    fm_0p00(1) = s_tpm
    fm_0p00(2) = s_rf
    fm_0p00(3) = s_rsq
    fm_0p00(4) = s_mns
    fm_0p00(5) = s_avg_w2l
    fm_0p00(6) = s_avg_pip
    fm_0p00(7) = s_depo_ini
    fm_0p00(8) = s_depo_fin
    fm_0p00(9) = s_cmsn
    fm_0p00(10) = s_link
    ' "0.00%"
    fm_0p00pc(1) = s_ar
    fm_0p00pc(2) = s_mdd
    fm_0p00pc(3) = s_win_pc
    ' color, bold, center
    fm_clr(1) = s_tpm
    fm_clr(2) = s_ar
    fm_clr(3) = s_mdd
    fm_clr(4) = s_rf
    fm_clr(5) = s_rsq
End Sub
Private Sub GSPR_Open_Report()
    Const ctrl_str As String = "file:///"
    Dim xI As Integer
    
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
    Set wc = ws.Cells
    open_fail = False
    rep_adr = ActiveCell.Value
' encoding %20
    If InStr(1, rep_adr, "%20", 1) Then
        rep_adr = WorksheetFunction.Substitute(rep_adr, "%20", " ")
    End If ' check address validity
    If Left(rep_adr, 8) = ctrl_str Then
        rep_adr = Right(rep_adr, Len(rep_adr) - 8)
        open_fail = False
        Set rb = Workbooks.Open(rep_adr)
        Set rs = rb.Sheets(1)
        Set rc = rs.Cells
    ElseIf Right(rep_adr, 5) <> ".html" Then
        MsgBox "Error. html files are required", , "GetStats Pro"
        open_fail = True
        Exit Sub
    ElseIf Dir(rep_adr) = "" Then
        MsgBox "Error. Empty address", , "GetStats Pro"
        open_fail = True
        Exit Sub
    Else
        open_fail = True
        Exit Sub
    End If
End Sub
Private Sub GSPR_Extract_stats()
    Dim used_inss As Integer
    Dim j As Integer, k As Integer, l As Integer
    Dim p_fr As Integer, p_lr As Integer
    Dim s As String, s2 As String, ch As String
    Dim rc As Range
    
    Dim RowTbLast As Long
    Dim rgMerged As Range, rgPL As Range, cell As Range
    Dim hasMerged As Boolean
    
    all_zeros = False
    hasMerged = False
    Set rc = rb.Sheets(1).Cells
' strategy name
    s = rc(3, 1).Value
    j = InStr(1, s, " strategy report for", 1)
    SV(s_strat, 2) = Left(s, j - 1)
' get trades count
    k = InStr(j, s, " instrument(s) from", 1)
    s = Left(s, k - 1)
    s = Right(s, Len(s) - j - 20)
    s2 = Replace(s, ",", "", 1)
    used_inss = Len(s) - Len(s2) + 1
    ins_td_r = 10
    For j = 1 To used_inss
        ins_td_r = rc.Find(what:="Closed positions", _
            after:=rc(ins_td_r, 1), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            searchorder:=xlByColumns, _
            searchdirection:=xlNext).Row
        If rc(ins_td_r, 2) <> 0 Then
            Exit For        ' found instrument with trades
        End If
    Next j
' trades closed
    SV(s_trades, 2) = rc(ins_td_r, 2)
' get instrument
    s = rc(ins_td_r - 9, 1)
    j = InStr(1, s, " ", 1)
    s = Right(s, Len(s) - j)
    SV(s_ins, 2) = s
' get parameters
    Call GSPR_get_algo_params(rc, Par)  ' *!
' sort parameters alphabetically
    Call GSPR_Par_Bubblesort
' Test begin
    SV(s_date_begin, 2) = CDate(Int(rc(ins_td_r - 7, 2)))
' Test end
'    SV(s_date_end, 2) = Int(rc(ins_td_r - 4, 2))
    SV(s_date_end, 2) = CDate(Int(rc(ins_td_r - 4, 2)))
' Months
    SV(s_mns, 2) = (SV(s_date_end, 2) - SV(s_date_begin, 2)) * 12 / 365
' TPM
    SV(s_tpm, 2) = Round(SV(s_trades, 2) / SV(s_mns, 2), 2)
' Initial deposit
    SV(s_depo_ini, 2) = CDbl(Replace(rc(5, 2), "’", ""))
' Commissions
    SV(s_cmsn, 2) = CDbl(Replace(rc(8, 2), "’", ""))
' File size
    SV(s_link, 2) = Round(FileLen(rep_adr) / 1024 ^ 2, 2)
' Check for "MERGED"
' quick & dirty fix
    Set rgPL = rc.Find(what:="Profit/Loss in pips", _
        after:=rc(1, 6), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        searchorder:=xlByColumns, _
        searchdirection:=xlNext)
    RowTbLast = rc(rgPL.Row, rgPL.Column).End(xlDown).Row
    Set rgMerged = rs.Range(rc(rgPL.Row + 1, 5), rc(RowTbLast, 5))
    For Each cell In rgMerged
        If cell.Value = "MERGED" Or cell.Value = "ERROR" Then
            hasMerged = True
            Exit For
        End If
    Next cell
' move on
    If SV(s_trades, 2) = 0 Or hasMerged = True Then
        all_zeros = True
        SV(s_depo_fin, 2) = CDbl(Replace(rc(6, 2), "’", "")) ' Finish deposit
    Else
        Call GSPR_Fill_Tradelog
    End If
    rb.Close savechanges:=False
End Sub
Private Sub GSPR_get_algo_params(rc As Range, Par As Variant)
    Dim p_fr As Integer, p_lr As Integer
    Dim j As Integer, k As Integer
    
    p_fr = 12
    p_lr = rc(12, 1).End(xlDown).Row
    ReDim Par(1 To p_lr - p_fr + 1, 1 To 2)
    For j = LBound(Par, 1) To UBound(Par, 1)
        For k = LBound(Par, 2) To UBound(Par, 2)
            Par(j, k) = rc(p_fr - 1 + j, k)
        Next k
    Next j
End Sub
Private Sub GSPR_Par_Bubblesort()
    Dim sj As Integer, sk As Integer
    Dim tmp_c1 As Variant, tmp_c2 As Variant
    
    For sj = 1 To UBound(Par, 1) - 1
        For sk = sj + 1 To UBound(Par, 1)
            If Par(sj, 1) > Par(sk, 1) Then
                tmp_c1 = Par(sk, 1)
                tmp_c2 = Par(sk, 2)
                Par(sk, 1) = Par(sj, 1)
                Par(sk, 2) = Par(sj, 2)
                Par(sj, 1) = tmp_c1
                Par(sj, 2) = tmp_c2
            End If
        Next sk
    Next sj
End Sub
Private Sub GSPR_Fill_Tradelog()
    Dim r As Integer, c As Integer, k As Integer
    Dim oc_fr As Integer
    Dim oc_lr As Long, ro_d As Long
    Dim tl_r As Integer, cds As Integer
    Dim win_sum As Double, los_sum As Double
    Dim win_ct As Integer, los_ct As Integer
    Dim rsqX() As Double, rsqY() As Double
    Dim s As String
    Dim ArrCommis As Variant
    
' get trade log first row - header
    tl_r = rc.Find(what:="Closed orders:", after:=rc(ins_td_r, 1), LookIn:=xlValues, LookAt _
        :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
        False, searchformat:=False).Row + 2
' FILL T1
    ReDim t1(0 To SV(s_trades, 2), 1 To 13)
    t1(0, 10) = SV(s_depo_ini, 2)
    t1(0, 11) = "pc_chg"
    t1(0, 12) = "hwm"
    t1(0, 13) = "dd"
    SV(s_pips, 2) = 0
    win_sum = 0
    los_sum = 0
    win_ct = 0
    los_ct = 0
    For r = LBound(t1, 1) To UBound(t1, 1)
        For c = LBound(t1, 2) To UBound(t1, 2) - 4      ' to 9th column included
            If r > 0 Then
                Select Case c
                    Case Is = 1
                        t1(r, c) = Replace(rc(tl_r + r, c + 1), ",", ".", 1, 1, 1)
                    Case Is = 6
                        t1(r, c) = rc(tl_r + r, c + 1)
                        If t1(r, c) > 0 Then
                            win_ct = win_ct + 1
                            win_sum = win_sum + t1(r, c)
                        Else
                            los_ct = los_ct + 1
                            los_sum = los_sum + t1(r, c)
                        End If
                        SV(s_pips, 2) = SV(s_pips, 2) + t1(r, c)
                    Case Else
                        t1(r, c) = rc(tl_r + r, c + 1)
                End Select
            Else
                t1(r, c) = rc(tl_r + r, c + 1)
            End If
        Next c
    Next r
' winning trades, %
    SV(s_win_pc, 2) = win_ct / SV(s_trades, 2)
' average winner / average loser
    If win_ct = 0 Or los_ct = 0 Then
        SV(s_avg_w2l, 2) = 0
    Else
        SV(s_avg_w2l, 2) = Abs((win_sum / win_ct) / (los_sum / los_ct))
    End If
' average trade, pips
    SV(s_avg_pip, 2) = SV(s_pips, 2) / SV(s_trades, 2)
' FILL T2
    cds = SV(s_date_end, 2) - SV(s_date_begin, 2) + 2
    ReDim t2(0 To cds, 1 To 5)
' 1) date, 2) EOD fin.res., 3) trades closed this day, 4) sum cmsn, 5) avg cmsn
    t2(0, 1) = "date"
    t2(0, 2) = "EOD_fin_res"
    t2(0, 3) = "days_cmsn"
    t2(0, 4) = "tds_closed"
    For r = 1 To UBound(t2, 1)
        t2(r, 4) = 0
    Next r
    t2(0, 5) = "tds_avg_cmsn"
' fill dates
    t2(1, 1) = SV(s_date_begin, 2) - 1
    For r = 2 To UBound(t2, 1)
        t2(r, 1) = t2(r - 1, 1) + 1
    Next r
' fill returns
    t2(1, 2) = SV(s_depo_ini, 2)
' fill sum_cmsn
    oc_fr = rc.Find(what:="Event log:", after:=rc(ins_td_r + SV(s_trades, 2), 1), LookIn:=xlValues, LookAt _
        :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
        False, searchformat:=False).Row + 2 ' header row
    oc_lr = rc(oc_fr, 1).End(xlDown).Row
'
    ro_d = 1
    For r = 2 To UBound(t2, 1)
' compare dates
        If t2(r, 1) = Int(CDate(rc(oc_fr + ro_d, 1))) Then  ' *! cdate
            Do While t2(r, 1) = Int(CDate(rc(oc_fr + ro_d, 1))) ' *! cdate
                If Left(rc(oc_fr + ro_d, 2), 6) = "Commis" Then
                    s = rc(oc_fr + ro_d, 3)
                    ArrCommis = Split(s, " ")
                    s = ArrCommis(UBound(ArrCommis))
                    s = Replace(s, ".", ",", 1, 1, 1)
                    s = Replace(s, "[", "")
                    s = Replace(s, "]", "")
                    If CDbl(s) <> 0 Then
                        t2(r, 3) = CDbl(s)              ' enter cmsn
                    End If
                End If
                ' go to next trade in Trade log
                If ro_d + oc_fr < oc_lr Then
                    ro_d = ro_d + 1
                Else
                    Exit Do ' or exit do, when after last row processed
                End If
            Loop
        End If
    Next r
' calculate trades closed per each day
    c = 1   ' count trades for that day
    k = 1
    For r = 1 To UBound(t1, 1)
        ' find this day in t2
        Do Until t2(k, 1) = Int(CDate(t1(r, 8)))    ' *! cdate
            k = k + 1
        Loop
        t2(k, 4) = t2(k, 4) + 1
    Next r
' calculate average commission for a trade
    c = 1
    For r = 1 To UBound(t2, 1)
        If t2(r, 4) > 0 Then
            t2(r, 5) = 0
            ' sum cmsns
            For k = c To r
                t2(r, 5) = t2(r, 5) + t2(k, 3)
            Next k
            t2(r, 5) = t2(r, 5) / t2(r, 4)
            c = r + 1
        End If
    Next r
    
' fill t1 - equity curve
    k = 1
    For r = 1 To UBound(t1, 1)
        Do Until t2(k, 1) = Int(CDate(t1(r, 8)))    ' *! cdate
            k = k + 1
        Loop
        t1(r, 10) = t1(r - 1, 10) + t1(r, 5) - t2(k, 5)
        ' percent change
        t1(r, 11) = (t1(r, 10) - t1(r - 1, 10)) / t1(r - 1, 10)
    Next r

' hwm & mdd - high watermark & maximum drawdown
    t1(1, 12) = WorksheetFunction.Max(t1(0, 10), t1(1, 10))
    t1(1, 13) = (t1(1, 12) - t1(1, 10)) / t1(1, 12)
    SV(s_mdd, 2) = 0
    For r = 2 To UBound(t1, 1)
        t1(r, 12) = WorksheetFunction.Max(t1(r - 1, 12), t1(r, 10))
        t1(r, 13) = (t1(r, 12) - t1(r, 10)) / t1(r, 12)
        ' update mdd
        If t1(r, 13) > SV(s_mdd, 2) Then
            SV(s_mdd, 2) = t1(r, 13)
        End If
    Next r
' finish deposit
    SV(s_depo_fin, 2) = t1(UBound(t1, 1), 10)
' annual return = (1 + sv(i, s_depo_fin))
    If SV(s_depo_fin, 2) >= 0 Then
        SV(s_ar, 2) = (1 + (SV(s_depo_fin, 2) - SV(s_depo_ini, 2)) / SV(s_depo_ini, 2)) ^ (12 / SV(s_mns, 2)) - 1
    Else
        SV(s_ar, 2) = 0
    End If

' rf - recovery factor
    If SV(s_ar, 2) > 0 Then
        If SV(s_mdd, 2) > 0 Then
            SV(s_rf, 2) = SV(s_ar, 2) / SV(s_mdd, 2)
        Else
            SV(s_rf, 2) = 999
        End If
    Else
        SV(s_rf, 2) = 0
    End If
' fill t2 EOD fin_res
    k = 1
    For r = 1 To UBound(t1, 1)
        If Int(CDate(t1(r, 8))) > t2(k, 1) Then
            Do Until t2(k, 1) = Int(CDate(t1(r, 8)))
                k = k + 1
                t2(k, 2) = t2(k - 1, 2)
            Loop
            t2(k, 2) = t1(r, 10)
        Else
            t2(k, 2) = t1(r, 10)
        End If
    Next r
    ' fill rest of empty days
    For r = k To UBound(t2, 1)
        If IsEmpty(t2(r, 2)) Then
            t2(r, 2) = t2(r - 1, 2)
        End If
    Next r
' r-square - rsq
    ReDim rsqX(1 To UBound(t2, 1))     ' the x-s and the y-s
    ReDim rsqY(1 To UBound(t2, 1))
    For r = 1 To UBound(t2, 1)
        rsqX(r) = t2(r, 1)
        rsqY(r) = t2(r, 2)
    Next r
    SV(s_rsq, 2) = WorksheetFunction.RSq(rsqX, rsqY)
End Sub
Sub Print_2D_Array1(ByVal print_arr As Variant, ByVal is_inverted As Boolean, _
                   ByVal row_offset As Integer, ByVal col_offset As Integer, _
                   ByVal print_cells As Range)
' Procedure prints any 2-dimensional array in a new Workbook, sheet 1.
' Arguments:
'       1) 2-D array
'       2) rows-colums (is_inverted = False) or columns-rows (is_inverted = True)
    
'    Dim wb_print As Workbook
'    Dim c_print As Range
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

Private Sub GSPR_Print_stats()
    Dim r As Integer, c As Integer
    Dim z As Variant
    
    wc.Clear
    If all_zeros Then
        z = Array(s_tpm, s_ar, s_mdd, s_rf, s_rsq, s_trades, _
            s_win_pc, s_pips, s_avg_w2l, s_avg_pip, s_cmsn)
        For r = LBound(z) To UBound(z)
            SV(z(r), 2) = 0
        Next r
    End If
' print statistics names and values
    For r = LBound(SV, 1) To UBound(SV, 1)
        For c = LBound(SV, 2) To UBound(SV, 2)
            wc(r, c) = SV(r, c)
        Next c
    Next r
' print parameters
    wc(UBound(SV) + 2, 1) = "Parameters"
    For r = LBound(Par, 1) To UBound(Par, 1)
        For c = LBound(Par, 2) To UBound(Par, 2)
            wc(UBound(SV) + 2 + r, c) = Par(r, c)
        Next c
    Next r
    If all_zeros = False Then
' print tradelog-1
        ReDim Preserve t1(0 To UBound(t1, 1), 1 To 11)
        For r = LBound(t1, 1) To UBound(t1, 1)
            For c = LBound(t1, 2) To UBound(t1, 2)
                wc(r + 1, c + 2) = t1(r, c)
            Next c
        Next r
' print tradelog-2
        ReDim Preserve t2(0 To UBound(t2, 1), 1 To 2)
'        ReDim Preserve t2(0 To UBound(t2, 1), 1 To UBound(t2, 2))
        For r = LBound(t2, 1) To UBound(t2, 1)
            For c = LBound(t2, 2) To UBound(t2, 2)
                wc(r + 1, c + UBound(t1, 2) + 2) = t2(r, c)
            Next c
        Next r
    End If
' apply formats
    For r = LBound(fm_date) To UBound(fm_date)
        wc(fm_date(r), 2).NumberFormat = "yyyy-mm-dd"
    Next r
    For r = LBound(fm_0p00) To UBound(fm_0p00)
        wc(fm_0p00(r), 2).NumberFormat = "0.00"
    Next r
    For r = LBound(fm_0p00pc) To UBound(fm_0p00pc)
        wc(fm_0p00pc(r), 2).NumberFormat = "0.00%"
    Next r
    For r = LBound(fm_clr) To UBound(fm_clr)
        For c = 1 To 2
            With wc(fm_clr(r), c)
                .Interior.Color = RGB(184, 204, 228)
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
            End With
        Next c
    Next r
' hyperlink
    ws.Hyperlinks.Add anchor:=wc(s_link, 2), Address:=rep_adr
    ws.Range(Columns(1), Columns(2)).AutoFit
End Sub
Private Sub GSPR_Build_Chart_Check_Report_Type()
    ch_rep_type = False
    If ActiveSheet.Cells.Find(what:=rep_type) Is Nothing Then
        MsgBox "Error. Wrong report type", , "GetStats Pro"
        Exit Sub
    Else
        ch_rep_type = True
    End If
End Sub
Private Sub GSPR_Remove_Chart()
    Dim img As Shape
    
    For Each img In ActiveSheet.Shapes
        img.Delete
    Next
End Sub
Private Sub GSPR_Build_Charts_Singe_Button()

    If ActiveSheet.Cells(11, 2) > 0 Then
        Application.ScreenUpdating = False
        Call GSPR_Build_Charts_Single_Report
        Application.ScreenUpdating = True
    Else
        MsgBox "0 positions, unable to build chart", , "GetStats Pro"
    End If
End Sub
Private Sub GSPR_Build_Charts_Singe_Button_EN()

    If ActiveSheet.Cells(11, 2) > 0 Then
        Application.ScreenUpdating = False
        Call GSPR_Build_Charts_Single_Report_EN
        Application.ScreenUpdating = True
    Else
        MsgBox "0 positions, unable to build chart", , "GetStats Pro"
    End If
End Sub
Private Sub GSPR_Build_Charts_Single_Report()
    Const my_rnd As Integer = 100
    Dim ulr As Integer, ulc As Integer, chobj_n As Integer
    Dim rngX As Range, rngY As Range
    Dim ChTitle As String
    Dim MinVal As Long, maxVal As Long
    Dim sc As Range
    Dim lr As Integer
    
    Call GSPR_Build_Chart_Check_Report_Type
    If ch_rep_type = False Then
        Exit Sub
    End If
    If ActiveSheet.Shapes.count > 0 Then
        Call GSPR_Remove_Chart
        Exit Sub
    End If
    Set sc = ActiveSheet.Cells
    lr = sc(1, 15).End(xlDown).Row
' data to pass to chart-builder
    ulr = 1
    ulc = 3
    Set rngX = Range(sc(2, 14), sc(lr, 14))
    Set rngY = Range(sc(2, 15), sc(lr, 15))
    ChTitle = "Equity curve. Strategy '" & sc(1, 2) & "', " & sc(2, 2) & ", " & sc(8, 2) & "-" & sc(9, 2) & "."
    MinVal = WorksheetFunction.Min(rngY)
    maxVal = WorksheetFunction.Max(rngY)
    MinVal = my_rnd * Int(MinVal / my_rnd)
    maxVal = my_rnd * Int(maxVal / my_rnd) + my_rnd
    Call GSPR_Chart_Classic_wMinMax(sc, ulr, ulc, rngX, rngY, ChTitle, MinVal, maxVal)
' data to pass to chart-builder
    ulr = 23
    ulc = 3
    lr = sc(1, 8).End(xlDown).Row
    Set rngY = Range(sc(2, 8), sc(lr, 8))
    ChTitle = "Trade result, pips. Strategy '" & sc(1, 2) & "', " & sc(2, 2) & ", " & sc(8, 2) & "-" & sc(9, 2) & "."
    chobj_n = 2
    Call GSPR_Hist_Y(sc, ulr, ulc, rngY, ChTitle, chobj_n)
    sc(1, 1).Activate
End Sub
Private Sub GSPR_Build_Charts_Single_Report_EN()
    Const my_rnd As Integer = 100
    Dim ulr As Integer, ulc As Integer, chobj_n As Integer
    Dim rngX As Range, rngY As Range
    Dim ChTitle As String
    Dim MinVal As Long, maxVal As Long
    Dim sc As Range
    Dim lr As Integer
    
    Call GSPR_Build_Chart_Check_Report_Type
    If ch_rep_type = False Then
        Exit Sub
    End If
    If ActiveSheet.Shapes.count > 0 Then
        Call GSPR_Remove_Chart
        Exit Sub
    End If
    Set sc = ActiveSheet.Cells
    lr = sc(1, 15).End(xlDown).Row
' data to pass to chart-builder
    ulr = 1
    ulc = 3
    Set rngX = Range(sc(2, 14), sc(lr, 14))
    Set rngY = Range(sc(2, 15), sc(lr, 15))
    ChTitle = "Equity curve. Strategy '" & sc(1, 2) & "', " & sc(2, 2) & ", " & sc(8, 2) & "-" & sc(9, 2) & "."
    MinVal = WorksheetFunction.Min(rngY)
    maxVal = WorksheetFunction.Max(rngY)
    MinVal = my_rnd * Int(MinVal / my_rnd)
    maxVal = my_rnd * Int(maxVal / my_rnd) + my_rnd
    Call GSPR_Chart_Classic_wMinMax(sc, ulr, ulc, rngX, rngY, ChTitle, MinVal, maxVal)
' data to pass to chart-builder
    ulr = 23
    ulc = 3
    lr = sc(1, 8).End(xlDown).Row
    Set rngY = Range(sc(2, 8), sc(lr, 8))
    ChTitle = "Trade result, pips. Strategy '" & sc(1, 2) & "', " & sc(2, 2) & ", " & sc(8, 2) & "-" & sc(9, 2) & "."
    chobj_n = 2
    Call GSPR_Hist_Y(sc, ulr, ulc, rngY, ChTitle, chobj_n)
    sc(1, 1).Activate
End Sub
Private Sub GSPR_EN_Translate()
    Dim c As Range
    
    Set c = ActiveSheet.Cells
    c(1, 1) = "Strategy"
    c(2, 1) = "Instrument"
    c(3, 1) = "Trades per month"
    c(4, 1) = "Annualized return, %"
    c(5, 1) = "Maximum drawdown, %"
    c(6, 1) = "Recovery factor"
    c(7, 1) = "R-squared"
    c(8, 1) = "Test begin date"
    c(9, 1) = "Test end date"
    c(10, 1) = "Months"
    c(11, 1) = "Positions closed"
    c(12, 1) = "Winners, %"
    c(13, 1) = "Pips"
    c(14, 1) = "Avg. winner/loser, pips"
    c(15, 1) = "Avg. trade, pips"
    c(16, 1) = "Initial balance"
    c(17, 1) = "End balance"
    c(18, 1) = "Commissions"
    c(19, 1) = "Report size (MB), link"
    c(20, 1) = "Report type"
    c(22, 1) = "Parameters"
    c(22, 2) = "results"
End Sub
Private Sub GSPR_Chart_Classic_wMinMax(sc As Range, _
                ulr As Integer, _
                ulc As Integer, _
                rngX As Range, _
                rngY As Range, _
                ChTitle As String, _
                MinVal As Long, _
                maxVal As Long)
'    Dim chW As Integer, chH As Integer          ' chart width, chart height
    Dim chFontSize As Integer                   ' chart title font size
    Dim rng_to_cover As Range
    
'    chW = 624   ' standrad cell width = 48 pix
'    chH = 317   ' standard cell height = 15 pix. 330
    chFontSize = 12
    Set rng_to_cover = Range(sc(ulr, ulc), sc(22, 15))
' build chart
    rngY.Select
    ActiveSheet.Shapes.AddChart.Select
' adjust chart placement
    With ActiveSheet.ChartObjects(1)
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
'        .Width = chW
        .Width = rng_to_cover.Width
'        .Height = chH
        .Height = rng_to_cover.Height
'        .Placement = xlFreeFloating ' do not resize chart if cells resized
    End With
    With ActiveChart
        .SetSourceData Source:=Application.Union(rngX, rngY)
        .ChartType = xlLine
        .Legend.Delete
        .Axes(xlValue).MinimumScale = MinVal
        .Axes(xlValue).MaximumScale = maxVal
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
End Sub
Private Sub GSPR_Hist_Y(sc As Range, _
                        ulr As Integer, _
                        ulc As Integer, _
                        rngY As Range, _
                        ChTitle As String, _
                        chobj_n As Integer)
'    Dim chW As Integer, chH As Integer          ' chart width, chart height
    Dim chFontSize As Integer                   ' chart title font size
    Dim rng_to_cover As Range
    
'    chW = 624      ' chart width, pixels
'    chH = 270      ' chart height, pixels
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + 21, 15))
    chFontSize = 12
    rngY.Select
    ActiveSheet.Shapes.AddChart.Select        ' build chart
    With ActiveChart
        .SetSourceData Source:=rngY
        .ChartType = xlColumnClustered                  ' chart type - histogram
        .Legend.Delete
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow     ' axis position
    End With
    With ActiveSheet.ChartObjects(chobj_n)    ' adjust chart placement
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
'        .Width = chW
        .Width = rng_to_cover.Width
'        .Height = chH
        .Height = rng_to_cover.Height
'        .Placement = xlFreeFloating     ' do not resize chart if cells resized
    End With
End Sub



' MODULE: Tools
Option Explicit

Dim current_decimal As String
Dim undo_sep As Boolean, undo_usesyst As Boolean
Dim user_switched As Boolean
    
Sub GSPR_Remove_CommandBar()
    
    Dim i As Long
    
    On Error Resume Next
    
    For i = 1 To 9
        Application.CommandBars("GSPR-" & i).Delete
    Next i
    
End Sub

Sub GSPR_Create_CommandBar()
    
    Dim cBar1 As CommandBar
    Dim cBar2 As CommandBar
    Dim cBar3 As CommandBar
    Dim cBar4 As CommandBar
    Dim cBar5 As CommandBar
    Dim cBar6 As CommandBar
    Dim cBar7 As CommandBar
    Dim cBar8 As CommandBar
    Dim cBar9 As CommandBar
    Dim cControl As CommandBarControl
    
    Call GSPR_Remove_CommandBar
' Create toolbar
    Set cBar1 = Application.CommandBars.Add
    cBar1.Name = "GSPR-1"
    cBar1.Visible = True
' Create toolbar 2
    Set cBar2 = Application.CommandBars.Add
    cBar2.Name = "GSPR-2"
    cBar2.Visible = True
' Create toolbar 3
    Set cBar3 = Application.CommandBars.Add
    cBar3.Name = "GSPR-3"
    cBar3.Visible = True
' Create toolbar 4
    Set cBar4 = Application.CommandBars.Add
    cBar4.Name = "GSPR-4"
    cBar4.Visible = True
' Create toolbar 5
    Set cBar5 = Application.CommandBars.Add
    cBar5.Name = "GSPR-5"
    cBar5.Visible = True
' Create toolbar 6
    Set cBar6 = Application.CommandBars.Add
    cBar6.Name = "GSPR-6"
    cBar6.Visible = True
' Create toolbar 7
    Set cBar7 = Application.CommandBars.Add
    cBar7.Name = "GSPR-7"
    cBar7.Visible = True
' Create toolbar 8
    Set cBar8 = Application.CommandBars.Add
    cBar8.Name = "GSPR-8"
    cBar8.Visible = True
' Create toolbar 9
    Set cBar9 = Application.CommandBars.Add
    cBar9.Name = "GSPR-9"
    cBar9.Visible = True

' ROW 1 ===========
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 351
        .OnAction = "GSPR_Single_Core"
        .TooltipText = "Main report"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Main"
    End With
    
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 352 ' 352   126, 59, 630
        .OnAction = "GSPR_Single_Extra"
        .TooltipText = "Super duper extra report"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Extra"
    End With
    
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 418
        .OnAction = "GSPR_Build_Charts_Singe_Button"
        .TooltipText = "Build chart"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Chart"
    End With

' ROW 2 ===========
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 124
        .OnAction = "GSPR_show_sheet_index"
        .TooltipText = "Show this sheet's index"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ShIndex"
    End With
    
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 205
        .OnAction = "GSPR_Go_to_sheet_index"
        .TooltipText = "Go to sheet with your index"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ToIndex"
    End With

' ROW 3 ===========
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 585
        .OnAction = "GSPR_Copy_Sheet_Next"
        .TooltipText = "Copy sheet"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "CopySh"
    End With
    
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 478
        .OnAction = "GSPR_Delete_Sheet"
        .TooltipText = "Delete active sheet"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "DelSh"
    End With
    
    ' SEPARATOR
    user_switched = False
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 98
        .OnAction = "GSPR_Separator_Manual_Switch"
        .TooltipText = "Change decimal separator"
        .Control.Style = msoButtonIconAndCaption
'        .Caption = "Separator"
    End With

' ROW 4 ===========
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 688
        .OnAction = "GSPRM_Merge_Summaries"
        .TooltipText = "Merge on recovery factor"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "MergeRF"
    End With
    
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 688 ' 477
        .OnAction = "GSPRM_Merge_Sharpe"
        .TooltipText = "Merge reports on Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "MergeSR"
    End With
    
' ROW 5 ===========
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 279    ' 31, 279
        .OnAction = "GSPR_Mixer_Copy_Sheet_To_Book"
        .TooltipText = "Add this sheet to 'mixer'"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ToMix"
    End With

    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 645   ' 601
        .OnAction = "GSPR_robo_mixer"
        .TooltipText = "Magic - make the MIX"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "MIX"
    End With

    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 424   ' 601
        .OnAction = "GSPR_trades_to_days"
        .TooltipText = "Mix chart on calendar days"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "MixChart"
    End With

' ROW 6 ===========
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 283
        .OnAction = "CalcMore"
        .TooltipText = "Calculate rest of KPI"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "CalcMore"
    End With
    
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 424   ' 601
        .OnAction = "Stats_Chart_from_Joined_Windows"
        .TooltipText = "Chart for joined windows"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ChartJ"
    End With
    
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 435   ' 601
        .OnAction = "Calc_Sharpe_Ratio"
        .TooltipText = "Calculate Sharpe ratio for single sheet"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Sharpe"
    End With
' ROW 7 ===========
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 191
        .OnAction = "Params_To_Summary"
        .TooltipText = "Retrieve parameters/values to summary sheet"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ParamJ-Summary"
    End With
    
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 477   ' 601
        .OnAction = "Sharpe_to_all"
        .TooltipText = "Calculate Sharpe ratio on all sheets"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Sharpe all"
    End With

' ROW 8 ===========
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 430
        .OnAction = "Scatter_Sharpe"
        .TooltipText = "Build scatter plots based on Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ScatterPlots"
    End With
    
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 478
        .OnAction = "RemoveScatters"
        .TooltipText = "Remove all scatter plots"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "DelScatter"
    End With
    
' ROW 9 ===========
    Set cControl = cBar9.Controls.Add
    With cControl
        .FaceId = 143
        .OnAction = "SharpePivot"
        .TooltipText = "Merge summaries, calculate Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "SharpePvt"
    End With

End Sub

Private Sub GSPR_About_Link()
' unused
    
    ActiveWorkbook.FollowHyperlink Address:="https://vsatrader.ru/getstats/"

End Sub

Private Sub GSPR_Copy_Sheet_Next()
    
    Application.ScreenUpdating = False
    ActiveSheet.Copy after:=Sheets(ActiveSheet.Index)
    Application.ScreenUpdating = True

End Sub

Private Sub GSPR_Delete_Sheet()
    
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True

End Sub

Private Sub GSPR_Separator_Manual_Switch()
    
    Application.ScreenUpdating = False
    If user_switched Then
        Call GSPR_Separator_OFF
    Else
        Call GSPR_Separator_ON
    End If
    Application.ScreenUpdating = True

End Sub

Private Sub GSPR_Separator_ON()
    
    undo_sep = False
    undo_usesyst = False
    
    Const msg_rec As String = "Setting DOT as decimal separator." & vbNewLine & vbNewLine _
                & "Recommended for running GetStats." & vbNewLine & vbNewLine _
                & "To switch back to your separator press ""S"" again."
    Const msg_not As String = "Your separator is DOT. Optimal for GetStats."
    
    If Application.UseSystemSeparators Then     ' SYS - ON
        If Not Application.International(xlDecimalSeparator) = "." Then
            Application.UseSystemSeparators = False
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            MsgBox msg_rec, , "GetStats"
            undo_sep = True                     ' undo condition 1
            undo_usesyst = True                 ' undo condition 2
            user_switched = True
        Else
            undo_sep = False
            MsgBox msg_not, , "GetStats"
        End If
    Else                                        ' SYS - OFF
        If Not Application.DecimalSeparator = "." Then
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            MsgBox msg_rec, , "GetStats"
            undo_sep = True                     ' undo condition 1
            undo_usesyst = False                ' undo condition 2
            user_switched = True
        Else
            undo_sep = False
            MsgBox msg_not, , "GetStats"
        End If
    End If

End Sub

Private Sub GSPR_Separator_OFF()
    
    Const msg_not As String = "Separator not changed. Current: DOT. Optimal for GetStats."
    
    If undo_sep Then
        Application.DecimalSeparator = current_decimal
        If undo_usesyst Then
            Application.UseSystemSeparators = True
        End If
        MsgBox "Reverting back to user's decimal separator." & vbNewLine & vbNewLine _
        & "To switch back to recommended separator (DOT), press ""S"" again.", , "GetStats"
        user_switched = False
    Else
        MsgBox msg_not, , "GetStats"
    End If

End Sub



' MODULE: JFX_create
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



' MODULE: Join_intervals
Option Explicit

Const addInFName As String = "GetStats_BackTest_v1.21.xlsm"
Const joinShName As String = "join"
Const targetFdRow As Integer = 2
Const sourceFdFRow As Integer = 5

Dim positionTags As New Dictionary

Dim wsJ As Worksheet    ' worksheet "Join"
Dim cJ As Range         ' cells "Join"
Dim targetDateFrom As String, targetDateTo As String
Dim targetDateFromDt As Date, targetDateToDt As Date

Dim srcFdInfo() As Variant      ' source folders info
Dim matchFiles() As Variant     ' corresponding file lists

Private Sub Join_Intervals_Main()
    
    Dim i As Integer
    
    Application.ScreenUpdating = False
    Call InitPositionTags(positionTags)
    Call Init_sheet_cells
' sanity #1
    If Check_Target_Source = False Then
        MsgBox "Error. Target or source folders"
        Exit Sub
    End If
    srcFdInfo = Source_Folders_Info
' sanity #2 to 4
    For i = 2 To 4
        If Check_Column_Equal(srcFdInfo, i) = False Then
            MsgBox "Error. Files count, strategy name, or reports count"
            Exit Sub
        End If
    Next i
' matching files list - arr
    matchFiles = Matching_files
    Call Join_books
    Application.ScreenUpdating = True

End Sub

Private Sub Join_books()
    
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    Dim lastRow As Integer, lastRMatch As Integer
    Dim lastRowFull As Integer
    Dim nextRTarget As Integer
    Dim parMain() As Variant
    Dim parCompare() As Variant
    Dim wbs() As Variant
    Dim Rng As Range
    Dim rngFull As Range
    Dim targetWB As Workbook
    Dim wsTarget As Worksheet
    Dim cTarget As Range
    Dim rngMatch As Range
    Dim targetWBName As String
    Dim appSta As String
    Dim wsMain As Worksheet, wsSrch As Worksheet
    Dim cMain As Range, cSrch As Range
    
    targetDateFrom = Find_Extreme_Date(False, 5)
    targetDateTo = Find_Extreme_Date(True, 6)
    targetDateFromDt = Date_String_To_Date(targetDateFrom)
    targetDateToDt = Date_String_To_Date(targetDateTo)
    
    ReDim wbs(1 To UBound(srcFdInfo, 1))
    For i = LBound(matchFiles, 1) To UBound(matchFiles, 1)
        appSta = "File " & i & " (" & UBound(matchFiles, 1) & ")."
        Application.StatusBar = appSta
        For j = LBound(wbs) To UBound(wbs)
            Set wbs(j) = Workbooks.Open(matchFiles(i, j))
        Next j
        ' create target book
        Set targetWB = Workbooks.Add
        ' add sheets to targetWB
        Call Change_sheets_count(targetWB, wbs(1).Sheets.count)
        
' LOOP THROUGH ALL REPORTS
' FIND MATCHING PARAMETERS
' COPY TO TARGET BOOK
        For j = 3 To wbs(1).Sheets.count
            ' copy initial trades set to target book
            Set wsMain = wbs(1).Sheets(j)
            Set cMain = wsMain.Cells
            lastRow = cMain(wsMain.Rows.count, 3).End(xlUp).Row
            Set Rng = wsMain.Range(cMain(1, 3), cMain(lastRow, 13))
            Set wsTarget = targetWB.Sheets(j)
            Set cTarget = wsTarget.Cells
            Rng.Copy cTarget(1, 3)  ' copy trades
            lastRow = cMain(wsMain.Rows.count, 1).End(xlUp).Row
            Set Rng = wsMain.Range(cMain(23, 1), cMain(lastRow, 2))
            Call Remove_tag_from_parameters(Rng)
            Rng.Copy cTarget(23, 1) ' copy parameters
            ' move parameters to Arr
            Set Rng = wsMain.Range(cMain(23, 2), cMain(lastRow, 2))
            parMain = Parameters_to_arr(Rng, lastRow - 22)
            ' LOOP compare parMain to wsSrch / cSrch
            ' remove tags
            For k = 2 To UBound(wbs)
                For m = 3 To wbs(k).Sheets.count
                    Set wsSrch = wbs(k).Sheets(m)
                    Set cSrch = wsSrch.Cells
                    Set Rng = wsSrch.Range(cSrch(23, 1), cSrch(lastRow, 2))
                    Call Remove_tag_from_parameters(Rng)
                Next m
            Next k
            ' find matches, copy to target
            For k = 2 To UBound(wbs)
                For m = 3 To wbs(k).Sheets.count
                    Set wsSrch = wbs(k).Sheets(m)
                    Set cSrch = wsSrch.Cells
                    Set Rng = wsSrch.Range(cSrch(23, 2), cSrch(lastRow, 2))
                    parCompare = Parameters_to_arr(Rng, lastRow - 22)
                    If Parameters_Match(parMain, parCompare) Then
                        lastRMatch = cSrch(wsSrch.Rows.count, 3).End(xlUp).Row
                        Set rngMatch = wsSrch.Range(cSrch(2, 3), cSrch(lastRMatch, 13))
                        nextRTarget = cTarget(wsTarget.Rows.count, 3).End(xlUp).Row + 1
                        rngMatch.Copy cTarget(nextRTarget, 3)
                        ' fill some basic info: date from-to, trades count
                        If k = UBound(wbs) Then
                            Set rngMatch = wsSrch.Range(cSrch(1, 1), cSrch(2, 2))
                            rngMatch.Copy cTarget(1, 1)
                            Set rngMatch = wsSrch.Range(cSrch(3, 1), cSrch(22, 1))
                            rngMatch.Copy cTarget(3, 1)
                            cTarget(8, 2) = targetDateFromDt
                            cTarget(9, 2) = targetDateToDt
                            cTarget(11, 2) = cTarget(wsTarget.Rows.count, 3).End(xlUp).Row - 1
                        End If
                    End If
                Next m
            Next k
        Next j
        ' save & close all
        Application.StatusBar = appSta & " Saving target book " & i & "."
        targetWBName = Target_WB_Name(wbs(1).Name)
        targetWB.SaveAs fileName:=targetWBName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        targetWB.Close
        For j = LBound(wbs) To UBound(wbs)
            wbs(j).Close savechanges:=False
        Next j
    Next i
    Application.StatusBar = False

End Sub

Private Sub Remove_tag_from_parameters(ByRef Rng As Range)
    
    Dim c As Range
    
    For Each c In Rng
        If positionTags.Exists(c.Value) Then
            c.Offset(0, 1).Value = ""
            Exit For
        End If
    Next c

End Sub

Private Sub Change_sheets_count(ByRef someWB As Workbook, ByVal shCount As Integer)
' function returns a new workbook with specified number of sheets
    
    Const shNameOne As String = "summary"
    Const shNameTwo As String = "results"
    Dim i As Integer
    
    If someWB.Sheets.count > shCount Then
        Application.DisplayAlerts = False
        For i = 1 To someWB.Sheets.count - shCount
            someWB.Sheets(someWB.Sheets.count).Delete
        Next i
        Application.DisplayAlerts = True
    ElseIf someWB.Sheets.count < shCount Then
        For i = 1 To shCount - someWB.Sheets.count
            someWB.Sheets.Add after:=someWB.Sheets(someWB.Sheets.count)
        Next i
    End If
    someWB.Sheets(1).Name = shNameOne
    someWB.Sheets(2).Name = shNameTwo
' rename rest of sheets
    For i = 3 To someWB.Sheets.count
        someWB.Sheets(i).Name = i - 2
    Next i
    someWB.Sheets(3).Activate

End Sub

Private Sub Pick_target_folder()
' sub adds a folder path to cells(2, 1)
' in "Source folders" column (1)
    
    Dim fd As FileDialog
    
    Application.ScreenUpdating = False
    Call Init_sheet_cells
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Pick target folder"
'        .ButtonName = "OK"
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    cJ(targetFdRow, 1) = fd.SelectedItems(1)
    wsJ.Columns(1).AutoFit
    Application.ScreenUpdating = True

End Sub

Private Sub Add_source_folder()
' sub adds a folder path to next free row
' in "Source folders" column (1)
    
    Dim fd As FileDialog
    Dim nextFreeRow As Integer
    
    Application.ScreenUpdating = False
    Call Init_sheet_cells
    nextFreeRow = cJ(wsJ.Rows.count, 1).End(xlUp).Row + 1
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Pick folder with XLSX reports"
'        .ButtonName = "OK"
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    cJ(nextFreeRow, 1) = fd.SelectedItems(1)
    wsJ.Columns(1).AutoFit
    Application.ScreenUpdating = True

End Sub

Private Sub Clear_source_list()
' sub clears processing list (subfolders)
    
    Dim Rng As Range
    
    Application.ScreenUpdating = False
    Call Init_sheet_cells
    Set Rng = wsJ.Range(cJ(sourceFdFRow, 1), cJ(wsJ.Rows.count, 1))
    Rng.Clear
    Application.ScreenUpdating = True

End Sub

Private Sub Rename_source_files_no_postfix_dates()
    
    Dim lastRow As Integer
    Dim i As Integer, j As Integer
    Dim Rng As Range, c As Range
    Dim fList() As Variant
    Dim pFixes(1 To 4) As String
    Dim newFName As String, cutName As String
    Dim renameCounter As Integer
    Dim strategyName As String
    Dim instrumentName As String
    Dim dateFrom As String
    Dim dateTo As String
    
    Application.ScreenUpdating = False
    Call Init_sheet_cells
    pFixes(1) = "_mxu"
    pFixes(2) = "_mux"
    pFixes(3) = "_cxu"
    pFixes(4) = "_cux"
    lastRow = cJ(wsJ.Rows.count, 1).End(xlUp).Row
    Set Rng = wsJ.Range(cJ(sourceFdFRow, 1), cJ(lastRow, 1))
    For Each c In Rng
        fList = List_Files(c)
        ' rename
        For i = LBound(fList) To UBound(fList)
            newFName = Dir(fList(i))
' strategy name
            strategyName = Left(newFName, InStr(1, newFName, "-", vbTextCompare) - 1)
            cutName = Replace(newFName, strategyName & "-", "", 1, 1, vbTextCompare)
            ' remove postfix in strategy name
            For j = LBound(pFixes) To UBound(pFixes)
                If InStr(1, strategyName, pFixes(j), vbTextCompare) > 0 Then
                    strategyName = Replace(strategyName, pFixes(j), "", 1, 1, vbTextCompare)
                    Exit For
                End If
            Next j
' instrument name
            instrumentName = Left(cutName, 6)
            cutName = Replace(cutName, instrumentName & "-", "", 1, 1, vbTextCompare)
' date from
            dateFrom = Left(cutName, InStr(1, cutName, "-", vbTextCompare) - 1)
            cutName = Replace(cutName, dateFrom & "-", "", 1, 1, vbTextCompare)
            If Len(dateFrom) > 6 Then
                dateFrom = Right(dateFrom, 6)
            End If
' date to
            dateTo = Left(cutName, InStr(1, cutName, "-", vbTextCompare) - 1)
            cutName = Replace(cutName, dateTo, "", 1, 1, vbTextCompare)
            If Len(dateTo) > 6 Then
                dateTo = Right(dateTo, 6)
            End If
' compile full name anew
            newFName = c & "\" & strategyName & "-" & instrumentName & "-" & dateFrom & "-" & dateTo & cutName
            If fList(i) <> newFName Then
                Name fList(i) As newFName
                renameCounter = renameCounter + 1
            End If
        Next i
    Next c
    Application.ScreenUpdating = True
    MsgBox "Renamed " & renameCounter & " files"

End Sub

Private Sub Init_sheet_cells()
    
    Set wsJ = Workbooks(addInFName).Sheets(joinShName)
    Set cJ = wsJ.Cells

End Sub

Private Function Parameters_Match(ByVal pMain As Variant, ByVal pCompare As Variant) As Boolean
    
    Dim i As Integer
    
    For i = LBound(pMain) To UBound(pMain)
        If pMain(i) <> pCompare(i) Then
            Parameters_Match = False
            Exit Function
        End If
    Next i
    Parameters_Match = True

End Function

Private Function Parameters_to_arr(ByVal Rng As Range, ByVal ubnd As Integer) As Variant
    
    Dim arr() As Variant
    Dim i As Integer
    Dim c As Range
    
    ReDim arr(1 To ubnd)
    i = 0
    For Each c In Rng
        i = i + 1
        arr(i) = c
    Next c
    Parameters_to_arr = arr

End Function

Private Function Target_WB_Name(ByVal motherWBName As String) As String
    
    Dim j As Integer, vers As Integer
    Dim temp_s As String
    Dim coreName As String, finalName As String
    Dim currentIns As String
    
    currentIns = Extract_element_from_string(motherWBName, 2)
    coreName = cJ(targetFdRow, 1) & "\" & srcFdInfo(1, 3) & "-" & currentIns & _
            "-" & targetDateFrom & "-" & targetDateTo & "-" & srcFdInfo(1, 4)
    finalName = coreName & ".xlsx"
' check if exists
    If Dir(finalName) <> "" Then
        finalName = coreName & "(2).xlsx"
        If Dir(finalName) <> "" Then
            j = InStr(1, finalName, "(", 1)
            temp_s = Right(finalName, Len(finalName) - j)
            j = InStr(1, temp_s, ")", 1)
            vers = Left(temp_s, j - 1)
            finalName = coreName & "(" & vers & ").xlsx"
            Do Until Dir(finalName) = ""
                vers = vers + 1
                finalName = coreName & "(" & vers & ").xlsx"
            Loop
        End If
    End If
    Target_WB_Name = finalName

End Function

Private Function Find_Extreme_Date(ByVal searchMax As Boolean, ByVal colID As Integer) As String
    Dim i As Integer
    Dim xVal As Long
    Dim z As String
    
    If searchMax Then
        xVal = 0
        For i = LBound(srcFdInfo, 1) To UBound(srcFdInfo, 1)
            If Int(srcFdInfo(i, colID)) > xVal Then
                xVal = Int(srcFdInfo(i, colID))
                z = srcFdInfo(i, colID)
            End If
        Next i
    Else
        xVal = 999999
        For i = LBound(srcFdInfo, 1) To UBound(srcFdInfo, 1)
            If Int(srcFdInfo(i, colID)) < xVal Then
                xVal = Int(srcFdInfo(i, colID))
                z = srcFdInfo(i, colID)
            End If
        Next i
    End If
    Find_Extreme_Date = z

End Function

Private Function Date_String_To_Date(ByVal someDate As String) As Date
    
    Dim dtYear As Integer
    Dim dtMonth As Integer
    Dim dtDay As Integer
    
    dtYear = Left(someDate, 2)
    If dtYear <= 90 Then
        dtYear = 2000 + dtYear
    Else
        dtYear = 1900 + dtYear
    End If
    dtMonth = Left(Right(someDate, 4), 2)
    dtDay = Right(someDate, 2)
    Date_String_To_Date = CDate(dtDay & "." & dtMonth & "." & dtYear)

End Function

Private Function Matching_files() As Variant
    
    Dim arr() As Variant
    Dim fName As String, stratIns As String, matchPath As String
    Dim i As Integer, j As Integer

    ReDim arr(1 To srcFdInfo(1, 2), 1 To UBound(srcFdInfo, 1))
' 1st folder file list
    fName = Dir(srcFdInfo(1, 1) & "\")
    Do While fName <> ""
        i = i + 1
        arr(i, 1) = srcFdInfo(1, 1) & "\" & fName
'Debug.Print "i = " & i & ", val = " & arr(i, 1)
        fName = Dir()
    Loop
    For i = LBound(arr, 1) To UBound(arr, 1)    ' loop through 1st col file list, find matches
        fName = Dir(arr(i, 1))
        stratIns = Left(fName, Len(srcFdInfo(1, 3)) + 7)
'tmpS = "orig = " & arr(i, 1)
        For j = 2 To UBound(arr, 2)             ' columns = folders
            matchPath = srcFdInfo(j, 1) & "\" & stratIns & "-" & srcFdInfo(j, 5) _
                & "-" & srcFdInfo(j, 6) & "-" & srcFdInfo(j, 4) & ".xlsx"
            arr(i, j) = matchPath
'tmpS = tmpS & " - " & matchPath
        Next j
'Debug.Print tmpS
    Next i
    Matching_files = arr

End Function

Private Function Check_Column_Equal(ByVal arr As Variant, ByVal colID As Integer) As Boolean
    
    Dim s1 As String, s2 As String
    Dim i As Integer
    
    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        s1 = s1 & arr(i, colID)
        s2 = s2 & arr(i + 1, colID)
    Next i
    If s1 = s2 Then
        Check_Column_Equal = True
    Else
        Check_Column_Equal = False
    End If

End Function

Private Function Check_Target_Source() As Boolean
    
    Dim sourceFdLRow As Integer
    
    sourceFdLRow = cJ(wsJ.Rows.count, 1).End(xlUp).Row
' source folders count, must be > 1
' target folder must not be empty
    If sourceFdLRow > sourceFdFRow _
       And Not IsEmpty(cJ(targetFdRow, 1)) Then
        Check_Target_Source = True
    Else
        Check_Target_Source = False
    End If

End Function

Private Function Source_Folders_Info() As Variant
' creates a 2D array
' column 1: folder path
' column 2: files count in folder
' column 3: strategy name
' column 4: reports
' column 5: date from
' column 6: date to
    
    Dim arr() As Variant
    Dim lastRow As Integer
    Dim j As Integer
    Dim arrRow As Integer
    Dim randFileName As String
    
    lastRow = cJ(wsJ.Rows.count, 1).End(xlUp).Row
    ReDim arr(1 To lastRow - sourceFdFRow + 1, 1 To 6)
    For j = sourceFdFRow To lastRow
        arrRow = j - sourceFdFRow + 1
        ' 1. folder path
        arr(arrRow, 1) = cJ(j, 1)
        ' 2. files count
        arr(arrRow, 2) = Count_files(arr(arrRow, 1))
        ' 3. strategy name
        randFileName = Dir(arr(arrRow, 1) & "\")
        arr(arrRow, 3) = Extract_element_from_string(randFileName, 1)
        ' 4. reports
        arr(arrRow, 4) = Right(randFileName, Len(randFileName) - InStrRev(randFileName, "-", -1, vbTextCompare))
        arr(arrRow, 4) = Left(arr(arrRow, 4), Len(arr(arrRow, 4)) - 5)
        ' 5. date from
        arr(arrRow, 5) = Extract_element_from_string(randFileName, 3)
        ' 6. date to
        arr(arrRow, 6) = Extract_element_from_string(randFileName, 4)
    Next j
    Source_Folders_Info = arr

End Function

Private Function Extract_element_from_string(ByVal someString As String, _
                                     ByVal elemID As Integer) As String
    
    Dim outElem As String
    Dim cutName As String
    Dim i As Integer
    
    cutName = someString
    For i = 1 To elemID
        outElem = Left(cutName, InStr(1, cutName, "-", vbTextCompare) - 1)
        cutName = Replace(cutName, outElem & "-", "", 1, 1, vbTextCompare)
    Next i
    Extract_element_from_string = outElem

End Function

Private Function Count_files(ByVal folderPath As String)
    
    Dim fName As String
    Dim c As Integer
    
    fName = Dir(folderPath & "\*")
    Do While fName <> ""
        c = c + 1
        fName = Dir()
    Loop
    Count_files = c

End Function

Private Function List_Files(ByVal sPath As String) As Variant
' Function takes folder path
' returns files list in it as 1D array
    
    Dim vaArray() As Variant
    Dim i As Integer
    Dim oFile As Object
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFiles As Object
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    Set oFiles = oFolder.Files
    If oFiles.count = 0 Then Exit Function
    ReDim vaArray(1 To oFiles.count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = sPath & "\" & oFile.Name
        i = i + 1
    Next
    List_Files = vaArray

End Function

Sub Stats_Chart_from_Joined_Windows()
    
    Dim ws As Worksheet
    Dim Rng As Range, clr_rng As Range
    Dim ubnd As Long
    Dim lr_dates As Integer
    Dim tradesSet() As Variant
    Dim daysSet() As Variant
    
    Application.ScreenUpdating = False
    Set ws = ActiveSheet
    Set Rng = ws.Cells
    ubnd = Rng(ws.Rows.count, 3).End(xlUp).Row - 1
    If Rng(1, 15) <> "" Then
        Call GSPR_Remove_Chart2
        lr_dates = Rng(ws.Rows.count, 15).End(xlUp).Row
        Set clr_rng = ws.Range(Rng(1, 15), Rng(lr_dates, 16))
        clr_rng.Clear
    Else
        ' move to RAM
        tradesSet = Load_Slot_to_RAM2(Rng, ubnd)
        ' add Calendar x2 columns
        daysSet = Get_Calendar_Days_Equity2(tradesSet, Rng)
        ' print out
        Call Print_2D_Array2(daysSet, True, 0, 14, Rng)
        ' build chart
        Call WFA_Chart_Classic2(Rng, 1, 17)
    End If
    Application.ScreenUpdating = True

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
       
    chObj_idx = ActiveSheet.ChartObjects.count + 1
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
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
'    sc(1, first_col + 1) = chObj_idx
    sc(1, 15).Select

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

Private Sub GSPR_Remove_Chart2()
    
    Dim img As Shape
    
    For Each img In ActiveSheet.Shapes
        img.Delete
    Next

End Sub

Private Sub Add_Key_Stats()
    
    Dim ws As Worksheet
    Dim c As Range
    Dim i As Integer
    
    Application.ScreenUpdating = False
    For i = 3 To 4 ' ActiveWorkbook.Sheets.count
        Set ws = ActiveWorkbook.Sheets(i)
        Set c = ws.Cells
        ' TPM
        With c(3, 2)
            .Value = c(11, 2) / ((c(9, 2) - c(8, 2) + 1) / (365 / 12))
            .NumberFormat = "0.00"
        End With
    Next i
    Application.ScreenUpdating = True

End Sub

Sub Params_To_Summary()
    
    Const parFRow As Integer = 23
    
    Dim parLRow As Integer
    Dim i As Integer, j  As Integer, k As Integer, m As Integer
    Dim wsRes As Worksheet, ws As Worksheet
    Dim cRes As Range, c As Range
    Dim clz As Range
    Dim repNum As Integer
    
    Application.ScreenUpdating = False
' copy param names
    Set clz = Sheets(3).Cells
    Set wsRes = Sheets(2)
    Set cRes = wsRes.Cells
    
    parLRow = clz(parFRow, 1).End(xlDown).Row
    j = 2
    cRes(1, 1) = "#_link"
    For i = parFRow To parLRow
        cRes(1, j) = clz(i, 1)
        j = j + 1
    Next i
' copy parameters
    For i = 3 To Sheets.count
        repNum = i - 2
        j = i - 1
        m = 2
        Set ws = Sheets(i)
        Set c = ws.Cells
        cRes(i - 1, 1) = repNum
        For k = parFRow To parLRow
            cRes(j, m) = c(k, 2)
            m = m + 1
        Next k
        ' Add hyperlink to report sheet
        wsRes.Hyperlinks.Add anchor:=cRes(j, 1), Address:="", SubAddress:="'" & repNum & "'!R22C2"
        ' print "back to summary" link
        With c(22, 2)
            .Value = "results"
            .HorizontalAlignment = xlRight
        End With
        ws.Hyperlinks.Add anchor:=c(22, 2), Address:="", SubAddress:="'results'!A1"
    Next i
    wsRes.Activate
    cRes(2, 2).Activate
    wsRes.Rows("1:1").AutoFilter
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = True

End Sub

Sub CalcMore()
    
    Dim ws As Worksheet
    Dim Rng As Range
    Dim ubnd As Long
    Dim tradesSet() As Variant
    Dim daysSet() As Variant
    Dim pipsRng As Range
    
    Application.ScreenUpdating = False
    Set ws = ActiveSheet
    Set Rng = ws.Cells
    ubnd = Rng(ws.Rows.count, 3).End(xlUp).Row - 1
' move to RAM
    tradesSet = Load_Slot_to_RAM2(Rng, ubnd)
' add Calendar x2 columns
    daysSet = Get_Calendar_Days_Equity2(tradesSet, Rng)
' TPM
    With Rng(3, 2)
        .Value = pmTradesPerMonth(Rng(8, 2), Rng(9, 2), Rng(11, 2))
        .NumberFormat = "0.00"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With
' AR
    With Rng(4, 2)
        .Value = pmAR(daysSet, Rng(8, 2), Rng(9, 2))
        .NumberFormat = "0.00%"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With
' MDD
    With Rng(5, 2)
        .Value = pmMDD(tradesSet)
        .NumberFormat = "0.00%"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With
' RF
    With Rng(6, 2)
        .Value = Rng(4, 2).Value / Rng(5, 2).Value
        .NumberFormat = "0.00"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With
' RSQ
    With Rng(7, 2)
        .Value = pmRSQ(daysSet)
        .NumberFormat = "0.00"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With
' months
    With Rng(10, 2)
        .Value = (Rng(9, 2) - Rng(8, 2)) * 12 / 365
        .NumberFormat = "0.00"
    End With
' win ratio
    Set pipsRng = Range(Rng(2, 8), Rng(ubnd + 1, 8))
    With Rng(12, 2)
        .Value = pmWinRatio(pipsRng, ubnd)
        .NumberFormat = "0.00%"
    End With
' pips sum
    With Rng(13, 2)
        .Value = WorksheetFunction.Sum(pipsRng)
        .NumberFormat = "0.00"
    End With
' avg W/L ratio, pips
    With Rng(14, 2)
        .Value = Abs(WorksheetFunction.AverageIf(pipsRng, ">0") / WorksheetFunction.AverageIf(pipsRng, "<=0"))
        .NumberFormat = "0.00"
    End With
' avg trade, pips
    With Rng(15, 2)
        .Value = WorksheetFunction.Average(pipsRng)
        .NumberFormat = "0.00"
    End With
' depo ini
    With Rng(16, 2)
        .Value = 10000
        .NumberFormat = "0.00"
    End With
' depo finish
    With Rng(17, 2)
        .Value = Rng(16, 2).Value * daysSet(2, UBound(daysSet, 2))
        .NumberFormat = "0.00"
    End With
    Call Calc_Sharpe_Ratio
    ActiveSheet.Range(Columns(1), Columns(2)).AutoFit
    Application.ScreenUpdating = True

End Sub

Function pmTradesPerMonth(ByRef date0 As Date, _
        ByRef date9 As Date, _
        ByRef tradeCount As Long) As Double

    pmTradesPerMonth = tradeCount / ((date9 - date0 + 1) / 30.4)

End Function

Function pmAR(ByRef daysSet As Variant, _
        ByRef date0 As Date, _
        ByRef date9 As Date) As Double
            
    Dim finalEqCurvePoint As Double

' calc net return
    finalEqCurvePoint = daysSet(2, UBound(daysSet, 2))
    pmAR = finalEqCurvePoint ^ (365 / (date9 - date0 + 1)) - 1
    
End Function

Function pmMDD(ByRef tradesSet As Variant) As Double
' tradesSet:
'   INVERTED: COLUMNS, ROWS
'   1. open date
'   2. close date
'   3. return

' create x by ubound Array
'   1. equity curve
'   2. HWM
'   3. DD
    
    Dim arr() As Variant
    Dim maxDD As Double
    Dim i As Long, j As Long
    Dim tradesCount As Long
    
    tradesCount = UBound(tradesSet, 2)
    ReDim arr(1 To 3, 1 To tradesCount + 1)
    arr(1, 1) = 1   ' starting Equity curve
    arr(2, 1) = 1   ' starting HWM
    
    maxDD = 0   ' initialize MDD
    For i = 2 To UBound(arr, 2)
        j = i - 1
        arr(1, i) = arr(1, j) * (1 + tradesSet(3, j))   ' equity curve
        arr(2, i) = WorksheetFunction.Max(arr(2, j), arr(1, i)) ' HWM
        arr(3, i) = (arr(2, i) - arr(1, i)) / arr(2, i) ' DD
        If arr(3, i) > maxDD Then
            maxDD = arr(3, i)
        End If
    Next i
    pmMDD = maxDD
    
End Function

Function pmRSQ(ByRef daysSet As Variant) As Double

    Dim x() As Variant
    Dim y() As Variant
    Dim i As Long
    
    ReDim x(1 To UBound(daysSet, 2))
    ReDim y(1 To UBound(daysSet, 2))
    For i = LBound(daysSet, 2) To UBound(daysSet, 2)
        x(i) = i
        y(i) = daysSet(2, i)
    Next i
    pmRSQ = WorksheetFunction.RSq(x, y)
    
End Function

Function pmWinRatio(ByRef Rng As Range, ByRef tradesCount As Long) As Double

    Dim winners As Long
    Dim c As Range
    
    winners = 0
    For Each c In Rng
        If c.Value > 0 Then
            winners = winners + 1
        End If
    Next c
    pmWinRatio = winners / tradesCount
    
End Function



' MODULE: Sheet2
' empty


' MODULE: Sheet4
' empty


' MODULE: Sheet1


' MODULE: BackTest_Main_Multi
Option Explicit

Dim btWs As Worksheet, setWs As Worksheet
Dim btC As Range, setC As Range
Dim selectAll As Range, instrumentsList As Range

Dim stratFdRng As Range ' strategy folder cell
Dim stratNmRng As Range ' strategy name cell

Dim rdRepNameCol As Integer
Dim rdRepDateCol As Integer
Dim rdRepCountCol As Integer
Dim rdRepDepoIniCol As Integer
Dim rdRepRobotNameCol As Integer
Dim rdRepTimeFromCol As Integer
Dim rdRepTimeToCol As Integer
Dim rdRepLinkCol As Integer

Dim upperRow As Integer
Dim leftCol As Integer
Dim rightCol As Integer

Dim nextFreeRowBTreport As Integer

Dim activeInstrumentsList As Variant
Dim instrLotGroup As Variant
Dim dateFrom As Date, dateTo As Date
Dim dateFromStr As String, dateToStr As String
Dim stratFdPath As String
Dim stratNm As String
Dim htmlCount As Integer
Dim btNextFreeRow As Integer
Dim maxHtmlCount As Integer

Dim statusBarFolder As String
Dim oneFdFilesList As Variant

' separator variables
Dim currentDecimal As String
Dim undoSep As Boolean, undoUseSyst As Boolean

'---SHIFTS
'------ sn, sv
Dim s_strat As Integer, s_ins As Integer
Dim s_tpm As Integer, s_ar As Integer, s_mdd As Integer, s_rf As Integer, s_rsq As Integer
Dim s_date_begin As Integer, s_date_end As Integer, s_mns As Integer
Dim s_trades As Integer, s_win_pc As Integer, s_pips As Integer, s_avg_w2l As Integer, s_avg_pip As Integer
Dim s_depo_ini As Integer, s_depo_fin As Integer, s_cmsn As Integer, s_link As Integer, s_rep_type As Integer
'------ ov_bas
Dim s_ov_strat As Integer, s_ov_ins As Integer, s_ov_htmls As Integer
Dim s_ov_mns As Integer, s_ov_from As Integer, s_ov_to As Integer
Dim s_ov_params As Integer, s_ov_params_vbl As Integer, s_ov_created As Integer, s_ov_macro_ver As Integer

' BOOLEAN
Dim allZeros As Boolean    ' allZeros
Dim openFail As Boolean

' STRING
Dim repType As String
Dim folderToSave As String
Dim macroVer As String
Dim loopInstrument As String

' ARRAYS
Dim ov() As Variant
Dim SV() As Variant
Dim Par() As Variant, par_sum() As Variant, par_sum_head() As String
Dim sM() As Variant
Dim t1() As Variant
Dim t2() As Variant
Dim fm_date(1 To 2) As Integer, fm_0p00(1 To 10) As Integer, fm_0p00pc(1 To 3) As Integer, fm_clr(1 To 5) As Integer     ' count before changing

' OBJECTS
Dim mb As Workbook
Dim wbAddIn As Workbook

' DOUBLE
Dim depoIniCheck As Double

' DATE
Dim fdTimeStart As Date

Sub Process_Html_Folders()
' LOOP through folders
    ' LOOP through html files

' RETURNS:
' 1 file per each html folder
    Dim i As Integer
    Dim upperB As Integer
    
    Application.ScreenUpdating = False
    Call Init_Bt_Settings_Sheets(wbAddIn, setWs, btWs, btC, _
            activeInstrumentsList, instrLotGroup, stratFdPath, stratNm, _
            dateFrom, dateTo, htmlCount, _
            dateFromStr, dateToStr, btNextFreeRow, _
            maxHtmlCount, repType, macroVer, depoIniCheck, _
            rdRepNameCol, rdRepDateCol, rdRepCountCol, _
            rdRepDepoIniCol, rdRepRobotNameCol, rdRepTimeFromCol, _
            rdRepTimeToCol, rdRepLinkCol)
    If UBound(activeInstrumentsList) = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Instruments not selected."
        Exit Sub
    End If
    ' Separator - autoswitcher
    Call Separator_Auto_Switcher(currentDecimal, undoSep, undoUseSyst)
    upperB = UBound(activeInstrumentsList)
    ' LOOP THRU many FOLDERS
    For i = 1 To upperB
        loopInstrument = activeInstrumentsList(i)
        statusBarFolder = "Folders in queue: " & upperB - i + 1 & " (" & upperB & ")."
        Application.StatusBar = statusBarFolder
        oneFdFilesList = ListFiles(stratFdPath & "\" & activeInstrumentsList(i))
        ' LOOP THRU FILES IN ONE FOLDER
        openFail = False
        Call Loop_Thru_One_Folder
        If openFail Then
            Application.StatusBar = False
            Application.ScreenUpdating = True
            Exit For
        End If
    Next i
    Call Separator_Auto_Switcher_Undo(currentDecimal, undoSep, undoUseSyst)
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Beep
    
End Sub

Sub Loop_Thru_One_Folder()

    fdTimeStart = Now
    Call Prepare_sv_ov_fm
    Call Open_Reports
    If openFail Then
        Exit Sub
    End If
' Process one report and print
    Call Process_Each_Print
' check window
    Call Check_Window
' save
    Call Save_To_XLSX
    btNextFreeRow = btNextFreeRow + 1
    
End Sub

Sub Save_To_XLSX()

    Dim temp_s As String, corenm As String
    Dim vers As Integer, j As Integer
    Dim fNm As String
    Dim statusSaving As String
    
    statusSaving = statusBarFolder & " Saving..."
    Application.StatusBar = statusSaving
' core name
    corenm = folderToSave & stratNm & "-" & UCase(loopInstrument) & "-" _
        & dateFromStr & "-" & dateToStr & "-r" & ov(s_ov_htmls, 2)
    fNm = corenm & ".xlsx"
    If Dir(fNm) = "" Then
        mb.SaveAs fileName:=fNm, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Else
        fNm = corenm & "(2).xlsx"
        If Dir(fNm) <> "" Then
            j = InStr(1, fNm, "(", 1)
            temp_s = Right(fNm, Len(fNm) - j)
            j = InStr(1, temp_s, ")", 1)
            vers = Left(temp_s, j - 1)
            fNm = corenm & "(" & vers & ").xlsx"
            Do Until Dir(fNm) = ""
                vers = vers + 1
                fNm = corenm & "(" & vers & ").xlsx"
            Loop
        End If
        mb.SaveAs fileName:=fNm, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    End If
    btC(btNextFreeRow, rdRepNameCol) = fNm
    btC(btNextFreeRow, rdRepLinkCol) = "open"
    btWs.Hyperlinks.Add anchor:=btC(btNextFreeRow, rdRepLinkCol), Address:=fNm
    mb.Close savechanges:=False
    Application.StatusBar = False
    
End Sub

Sub Check_Window()

    Dim ws As Worksheet
    Dim c As Range
    Dim rng_check As Range, rng_c As Range
    Dim i As Integer
    Dim chk_row As Integer
    Dim add_c1 As Integer, add_c2 As Integer
    Dim add_c3 As Integer, add_c4 As Integer
    Dim dates_ok_counter As Integer
    Dim correctRobName As String
    
    add_c1 = Sheets(2).Cells(1, 1).End(xlToRight).Column + 1
    add_c2 = add_c1 + 1
    add_c3 = add_c2 + 1
    add_c4 = add_c3 + 1
    Sheets(2).Cells(1, add_c1) = "start"
    Sheets(2).Cells(1, add_c2) = "end"
    Sheets(2).Cells(1, add_c3) = "depo_ini"
    Sheets(2).Cells(1, add_c4) = "rob_name"
    
    For i = 3 To Sheets.count
        chk_row = i - 1
        Set c = Sheets(i).Cells
        ' check window start date
        If c(8, 2) = dateFrom Then
            Sheets(2).Cells(chk_row, add_c1) = "ok"
        Else
            With Sheets(2).Cells(chk_row, add_c1)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
        ' check window end date
        If c(9, 2) = dateTo Then
            Sheets(2).Cells(chk_row, add_c2) = "ok"
        Else
            With Sheets(2).Cells(chk_row, add_c2)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
        ' check depo_ini
        If CDbl(c(16, 2)) = depoIniCheck Then
            Sheets(2).Cells(chk_row, add_c3) = "ok"
        Else
            With Sheets(2).Cells(chk_row, add_c3)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
        ' check robotName
        Select Case setWs.Range("CodeSource")
            Case 1
                correctRobName = GetCorrectRobName(c(2, 2))
            Case 2
                correctRobName = btWs.Range("StrategyName")
        End Select
        If c(1, 2) = correctRobName Then
            Sheets(2).Cells(chk_row, add_c4) = "ok"
        Else
            With Sheets(2).Cells(chk_row, add_c4)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
    Next i
    
    Sheets(2).Activate
    Sheets(2).Rows("1:1").AutoFilter
    Sheets(2).Rows("1:1").AutoFilter
    
    Set rng_check = mb.Sheets(2).Range(Cells(2, add_c1), Cells(Sheets.count - 1, add_c2))
    ' result of checking dates, into addin
    For Each rng_c In rng_check
        If rng_c.Value <> "ok" Then
            With btC(btNextFreeRow, rdRepDateCol)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
            Exit For
        End If
    Next rng_c
    If btC(btNextFreeRow, rdRepDateCol).Value <> "error" Then
        btC(btNextFreeRow, rdRepDateCol).Value = "ok"
    End If
    
    ' result of checking sheets count, into addin
    If mb.Sheets.count - 2 = htmlCount Then
        btC(btNextFreeRow, rdRepCountCol) = "ok"
    Else
        With btC(btNextFreeRow, rdRepCountCol)
            .Value = "error"
            .Interior.Color = RGB(255, 0, 0)
        End With
    End If

    ' result of checking depo_ini, into addin
    Set rng_check = mb.Sheets(2).Range(Cells(2, add_c3), Cells(Sheets.count - 1, add_c3))
    For Each rng_c In rng_check
        If rng_c.Value <> "ok" Then
            With btC(btNextFreeRow, rdRepDepoIniCol)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
            Exit For
        End If
    Next rng_c
    
    If btC(btNextFreeRow, rdRepDepoIniCol).Value <> "error" Then
        btC(btNextFreeRow, rdRepDepoIniCol).Value = "ok"
    End If
    
    ' result of checking robot name, into addin
    Set rng_check = mb.Sheets(2).Range(Cells(2, add_c4), Cells(Sheets.count - 1, add_c4))
    For Each rng_c In rng_check
        If rng_c.Value <> "ok" Then
            With btC(btNextFreeRow, rdRepRobotNameCol)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
            Exit For
        End If
    Next rng_c
    If btC(btNextFreeRow, rdRepRobotNameCol).Value <> "error" Then
        btC(btNextFreeRow, rdRepRobotNameCol).Value = "ok"
    End If
    
    ' add timestamp
    btC(btNextFreeRow, rdRepTimeFromCol) = fdTimeStart
    btC(btNextFreeRow, rdRepTimeToCol) = Now
    
End Sub

Function GetCorrectRobName(ByRef currencyPair As String) As String
    ' stratNm
    ' instrLotGroup
    
    Dim postfix As String
    Dim i As Integer
    
    postfix = "not-found"
    For i = 1 To UBound(instrLotGroup, 1)
        If instrLotGroup(i, 1) = currencyPair Then
            postfix = instrLotGroup(i, 2)
            Exit For
        End If
    Next i
    GetCorrectRobName = stratNm + postfix
    
End Function

Sub Process_Each_Print()

    Dim i As Integer
    Dim rb As Workbook
    Dim os As Worksheet, hs As Worksheet   ' report sheet, overview sheet, html-processed sheet
    Dim ss As Worksheet                    ' summary sheet
    Dim time_started As Double
    Dim time_now As Double
    Const timer_step As Integer = 5
    Dim time_remaining As Double
    Dim rem_min As Integer, rem_sec As Integer, rem_sec_s As String
    Dim time_rem As String
    Dim counter_timer As Integer
    Dim sta As String
    
' create book
    Set mb = Workbooks.Add
    If mb.Sheets.count > 2 Then
        For i = 1 To mb.Sheets.count - 2
            Application.DisplayAlerts = False
                Sheets(mb.Sheets.count).Delete
            Application.DisplayAlerts = True
        Next i
    ElseIf mb.Sheets.count < 2 Then
        mb.Sheets.Add after:=mb.Sheets(mb.Sheets.count)
    End If
    Set os = mb.Sheets(1)
    os.Name = "summary"
    Set ss = mb.Sheets(2)
    ss.Name = "results"
' open and process each html-report
    time_started = Now
    counter_timer = 0
    For i = 1 To ov(s_ov_htmls, 2)
        counter_timer = counter_timer + 1
        If counter_timer = timer_step Then
            counter_timer = 0
            time_now = Now
            time_remaining = Round(((ov(s_ov_htmls, 2) - i) / i) * (time_now - time_started) * 1440, 2)
            rem_min = Int(time_remaining)
            rem_sec = Int((time_remaining - Int(time_remaining)) * 60)
            rem_sec_s = rem_sec
            If rem_sec < 10 Then
                rem_sec_s = "0" & rem_sec
            End If
            time_rem = rem_min & ":" & rem_sec_s
        End If
        If i < timer_step Then
            sta = "Processing report " & i & " (" & ov(s_ov_htmls, 2) & ")."
        Else
            sta = "Processing report " & i & " (" & ov(s_ov_htmls, 2) & "). Estimated time remaining " & time_rem
        End If
        Application.StatusBar = statusBarFolder & " " & sta
        Set rb = Workbooks.Open(oneFdFilesList(i))
        Set hs = mb.Sheets.Add(after:=mb.Sheets(mb.Sheets.count))
        Select Case i
            Case Is < 10
                hs.Name = "00" & i
            Case 10 To 99
                hs.Name = "0" & i
            Case Else
                hs.Name = i
        End Select
        ' extract statistics
        allZeros = False
        Call Proc_Extract_stats(rb, i)
        ' print statistics on sheet
        Call Proc_Print_stats(hs, i)
    Next i
    Application.StatusBar = False
' Overview: extract stats and print
    Call Overview_Summary_Extract_Print(os, ss)
    
End Sub

Sub Overview_Summary_Extract_Print(ByRef os As Worksheet, ByRef ss As Worksheet)

    Dim r As Integer, c As Integer, hr As Integer
    Dim oc As Range, sc As Range
    Dim Rng As Range
    
    Set oc = os.Cells
    Set sc = ss.Cells
' extract from last report
    ov(s_ov_strat, 2) = SV(s_strat, 2)
    ov(s_ov_ins, 2) = SV(s_ins, 2)
    ov(s_ov_mns, 2) = SV(s_mns, 2)
    ov(s_ov_from, 2) = SV(s_date_begin, 2)
    ov(s_ov_to, 2) = SV(s_date_end, 2)
    ov(s_ov_created, 2) = Now
    ov(s_ov_macro_ver, 2) = macroVer
' print OVERVIEW
    For r = LBound(ov, 1) To UBound(ov, 1)
        For c = LBound(ov, 2) To UBound(ov, 2)
            oc(r, c) = ov(r, c)
        Next c
    Next r
' apply formats
    oc(4, 2).NumberFormat = "0.00"
    oc(5, 2).NumberFormat = "yyyy-mm-dd"
    oc(6, 2).NumberFormat = "yyyy-mm-dd"
    oc(9, 2).NumberFormat = "yyyy-mm-dd, hh:mm"
    os.Columns(1).AutoFit
    os.Columns(2).AutoFit
' fill summary header
    sM(0, 0) = "html_link"
    sM(0, 1) = "#_link"
    sM(0, 2) = "tpm"
    sM(0, 3) = "ann_ret"
    sM(0, 4) = "mdd"
    sM(0, 5) = "rf"
    sM(0, 6) = "rsq"
    sM(0, 7) = "avg_tr_pips"
' print SUMMARY
    For r = LBound(sM, 1) To UBound(sM, 1)
        For c = 1 To UBound(sM, 2)
            sc(r + 1, c) = sM(r, c)
        Next c
    Next r
    ' parameters - head
    For c = LBound(par_sum_head) To UBound(par_sum_head)
        sc(1, c + UBound(sM, 2)) = par_sum_head(c)
    Next c
    ' parameters - values
    For r = LBound(par_sum, 1) To UBound(par_sum, 1)
        For c = LBound(par_sum, 2) To UBound(par_sum, 2)
            sc(r + 1, c + UBound(sM, 2)) = par_sum(r, c)
        Next c
    Next r
' apply formats
    ' 0.00%
    Set Rng = Range(sc(2, 3), sc(1 + UBound(sM, 1), 4))
    Rng.NumberFormat = "0.00%"
    ' 0.00
    Set Rng = Range(sc(2, 5), sc(1 + UBound(sM, 1), 7))
    Rng.NumberFormat = "0.00"
    Set Rng = Range(sc(1, 1), sc(1, UBound(sM, 2)))
    Rng.Font.Bold = True
' hyperlinks
    hr = UBound(SV) + 2
    For r = 1 To UBound(sM, 1)
        Select Case r
            Case Is < 10
                ss.Hyperlinks.Add anchor:=sc(r + 1, 1), Address:="", SubAddress:="'00" & r & "'!R" & hr & "C2"
            Case 10 To 99
                ss.Hyperlinks.Add anchor:=sc(r + 1, 1), Address:="", SubAddress:="'0" & r & "'!R" & hr & "C2"
            Case Else
                ss.Hyperlinks.Add anchor:=sc(r + 1, 1), Address:="", SubAddress:="'" & r & "'!R" & hr & "C2"
        End Select
    Next r
' add autofilter
    ss.Activate
    ss.Rows("1:1").AutoFilter
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    
End Sub

Sub Proc_Print_stats(ByRef hs As Worksheet, ByRef i As Integer)

    Dim r As Integer, c As Integer
    Dim hc As Range
    Dim z As Variant
    
    z = Array(3, 4, 5, 6, 7, 11, 12, 13, 14, 17)
    Set hc = hs.Cells
    If allZeros Then
        For r = LBound(z) To UBound(z)
            SV(z(r), 2) = 0
        Next r
    End If
' print statistics names and values
    For r = LBound(SV, 1) To UBound(SV, 1)
        For c = LBound(SV, 2) To UBound(SV, 2)
            hc(r, c) = SV(r, c)
        Next c
    Next r
' print parameters
    hc(UBound(SV) + 2, 1) = "Parameters"
    For r = LBound(Par, 1) To UBound(Par, 1)
        For c = LBound(Par, 2) To UBound(Par, 2)
            hc(UBound(SV) + 2 + r, c) = Par(r, c)
        Next c
    Next r
' print "back to summary" link
    With hc(UBound(SV) + 2, 2)
        .Value = "results"
        .HorizontalAlignment = xlRight
    End With
    hs.Hyperlinks.Add anchor:=hc(UBound(SV) + 2, 2), Address:="", SubAddress:="'results'!A1"

    If allZeros = False Then
    ' print tradelog-1
        ReDim Preserve t1(0 To UBound(t1, 1), 1 To 11)
        For r = LBound(t1, 1) To UBound(t1, 1)
            For c = LBound(t1, 2) To UBound(t1, 2)
                hc(r + 1, c + 2) = t1(r, c)
            Next c
        Next r
    ' print tradelog-2
        ReDim Preserve t2(0 To UBound(t2, 1), 1 To 2)
        For r = LBound(t2, 1) To UBound(t2, 1)
            For c = LBound(t2, 2) To UBound(t2, 2)
                hc(r + 1, c + UBound(t1, 2) + 2) = t2(r, c)
            Next c
        Next r
    End If
' apply formats
    For r = LBound(fm_date) To UBound(fm_date)
        hc(fm_date(r), 2).NumberFormat = "yyyy-mm-dd"
    Next r
    For r = LBound(fm_0p00) To UBound(fm_0p00)
        hc(fm_0p00(r), 2).NumberFormat = "0.00"
    Next r
    For r = LBound(fm_0p00pc) To UBound(fm_0p00pc)
        hc(fm_0p00pc(r), 2).NumberFormat = "0.00%"
    Next r
    For r = LBound(fm_clr) To UBound(fm_clr)
        For c = 1 To 2
            With hc(fm_clr(r), c)
                .Interior.Color = RGB(184, 204, 228)
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
            End With
        Next c
    Next r
' hyperlink
    hs.Hyperlinks.Add anchor:=hc(s_link, 2), Address:=sM(i, 0)
    hs.Activate
    hs.Range(Columns(1), Columns(2)).AutoFit
    
End Sub

Sub Proc_Extract_stats(ByRef rb As Workbook, ByRef i As Integer)

    Dim ins_td_r As Integer
    Dim varRow As Variant
    Dim j As Integer, k As Integer, l As Integer
    Dim p_fr As Integer, p_lr As Integer
    Dim s As String, ch As String
    Dim rc As Range
    Dim findStringInstrument As String
    
    Set rc = rb.Sheets(1).Cells
    s = rc(3, 1).Value
' Test end
    SV(s_date_end, 2) = CDate(Left(Right(s, 19), 10))
' get strategy name
    j = InStr(1, s, " strategy report for", 1)
    SV(s_strat, 2) = Left(s, j - 1)
' get trades count
    k = InStr(j, s, " instrument(s) from", 1)
' find relevant instrument, with trades; get "closed positions" row
    findStringInstrument = "Instrument " & Left(UCase(loopInstrument), 3) & "/" & Right(UCase(loopInstrument), 3)
    Set varRow = rc.Find(what:=findStringInstrument, after:=rc(10, 1), _
        LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByColumns, _
        searchdirection:=xlNext, MatchCase:=False, searchformat:=False)
    If varRow Is Nothing Then
        ins_td_r = rc.Find(what:="Closed positions", after:=rc(10, 1), _
        LookIn:=xlValues, LookAt:=xlWhole, searchorder:=xlByColumns, _
        searchdirection:=xlNext, MatchCase:=False, searchformat:=False).Row
    Else
        ins_td_r = varRow.Row + 9
    End If
' get parameters - Par()
    ' get parameters first & last row
    p_fr = 12
    p_lr = rc(12, 1).End(xlDown).Row
    ov(s_ov_params, 2) = p_lr - p_fr + 1
    ReDim Par(1 To ov(s_ov_params, 2), 1 To 2)
    ' fill Par
    For j = LBound(Par, 1) To UBound(Par, 1)
        For k = LBound(Par, 2) To UBound(Par, 2)
            Par(j, k) = rc(p_fr - 1 + j, k)
        Next k
    Next j
'     sort parameters alphabetically
    Call Par_Bubblesort
    If i = 1 Then
        ReDim par_sum(1 To ov(s_ov_htmls, 2), 1 To UBound(Par, 1))
        ReDim par_sum_head(1 To UBound(Par, 1))
        For j = LBound(par_sum_head) To UBound(par_sum_head)
            par_sum_head(j) = Par(j, 1)
        Next j
    End If
    For j = LBound(par_sum, 2) To UBound(par_sum, 2)
        par_sum(i, j) = Par(j, 2)
    Next j
' File size and link
    SV(s_link, 2) = Round(FileLen(oneFdFilesList(i)) / 1024 ^ 2, 2)
' trades closed
    SV(s_trades, 2) = rc(ins_td_r, 2)
' get instrument
    SV(s_ins, 2) = Left(UCase(loopInstrument), 3) & "/" & Right(UCase(loopInstrument), 3)
' Test begin
    SV(s_date_begin, 2) = CDate(Int(rc(ins_td_r - 7, 2))) ' *! new cdate
' Test end
    SV(s_date_end, 2) = CDate(Int(rc(ins_td_r - 4, 2).Value))    ' *! removed "-1"
' Months
    SV(s_mns, 2) = (SV(s_date_end, 2) - SV(s_date_begin, 2)) * 12 / 365
' TPM
    If rc(ins_td_r, 2) = 0 Then
        SV(s_tpm, 2) = 0
    Else
        SV(s_tpm, 2) = Round(SV(s_trades, 2) / SV(s_mns, 2), 2)
    End If
' Initial deposit
    SV(s_depo_ini, 2) = CDbl(Replace(rc(5, 2), "’", ""))
'' Finish deposit
'    SV(s_depo_fin, 2) = CDbl(Replace(rc(6, 2), "’", ""))
' Commissions
    SV(s_cmsn, 2) = CDbl(Replace(rc(8, 2), "’", ""))

' Check if has "MERGED" trades
' quick & dirty fix == BEGIN
    Dim RowTbLast As Long
    Dim rgMerged As Range, rgPL As Range, cell As Range
    Dim hasMerged As Boolean

    hasMerged = False

    Set rgPL = rc.Find(what:="Profit/Loss in pips", _
        after:=rc(1, 6), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        searchorder:=xlByColumns, _
        searchdirection:=xlNext)
    RowTbLast = rc(rgPL.Row, rgPL.Column).End(xlDown).Row
    Set rgMerged = rb.Sheets(1).Range(rc(rgPL.Row + 1, 5), rc(RowTbLast, 5))
    For Each cell In rgMerged
        If cell.Value = "MERGED" Or cell.Value = "ERROR" Then
            hasMerged = True
            Exit For
        End If
    Next cell

    If rc(ins_td_r, 2) = 0 Or hasMerged = True Then
' quick & dirty fix == END
        allZeros = True
        sM(i, 0) = oneFdFilesList(i)
        sM(i, 1) = i
        sM(i, 2) = 0
        sM(i, 3) = 0
        sM(i, 4) = 0
        sM(i, 5) = 0
        sM(i, 6) = 0
        sM(i, 7) = 0
        rb.Close savechanges:=False
        Exit Sub
    End If
    
    Call Fill_Tradelogs(rc, ins_td_r)
    
' fill summary stats
    sM(i, 0) = oneFdFilesList(i)
    sM(i, 1) = i
    sM(i, 2) = SV(s_tpm, 2)
    sM(i, 3) = SV(s_ar, 2)
    sM(i, 4) = SV(s_mdd, 2)
    sM(i, 5) = SV(s_rf, 2)
    sM(i, 6) = SV(s_rsq, 2)
    sM(i, 7) = SV(s_avg_pip, 2)
    rb.Close savechanges:=False
    
End Sub

Sub Fill_Tradelogs(ByRef rc As Range, ByRef ins_td_r As Integer)

    Dim r As Integer, c As Integer, k As Integer
    Dim oc_fr As Integer
    Dim oc_lr As Long, ro_d As Long
    Dim tl_r As Integer, cds As Integer
    Dim win_sum As Double, los_sum As Double
    Dim win_ct As Integer, los_ct As Integer
    Dim rsqX() As Double, rsqY() As Double
    Dim s As String
    Dim ArrCommis As Variant
    
' get trade log first row - header
    tl_r = rc.Find(what:="Closed orders:", after:=rc(ins_td_r, 1), LookIn:=xlValues, LookAt _
        :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
        False, searchformat:=False).Row + 2
' FILL T1
    ReDim t1(0 To SV(s_trades, 2), 1 To 13)
    t1(0, 10) = SV(s_depo_ini, 2)
    t1(0, 11) = "pc_chg"
    t1(0, 12) = "hwm"
    t1(0, 13) = "dd"
    SV(s_pips, 2) = 0
    win_sum = 0
    los_sum = 0
    win_ct = 0
    los_ct = 0
    For r = LBound(t1, 1) To UBound(t1, 1)
        For c = LBound(t1, 2) To UBound(t1, 2) - 4      ' to 9th column included
            If r > 0 Then
                Select Case c
                    Case Is = 1
                        t1(r, c) = Replace(rc(tl_r + r, c + 1), ",", ".", 1, 1, 1)
                    Case Is = 6
                        t1(r, c) = rc(tl_r + r, c + 1)
                        If t1(r, c) > 0 Then
                            win_ct = win_ct + 1
                            win_sum = win_sum + t1(r, c)
                        Else
                            los_ct = los_ct + 1
                            los_sum = los_sum + t1(r, c)
                        End If
                        SV(s_pips, 2) = SV(s_pips, 2) + t1(r, c)
                    Case Else
                        t1(r, c) = rc(tl_r + r, c + 1)
                End Select
            Else
                t1(r, c) = rc(tl_r + r, c + 1)
            End If
        Next c
    Next r
' winning trades, %
    SV(s_win_pc, 2) = win_ct / SV(s_trades, 2)
' average winner / average loser
    If win_ct = 0 Or los_ct = 0 Then
        SV(s_avg_w2l, 2) = 0
    Else
        SV(s_avg_w2l, 2) = Abs((win_sum / win_ct) / (los_sum / los_ct))
    End If
' average trade, pips
    SV(s_avg_pip, 2) = SV(s_pips, 2) / SV(s_trades, 2)
' FILL T2
    cds = SV(s_date_end, 2) - SV(s_date_begin, 2) + 2
    ReDim t2(0 To cds, 1 To 5)
' 1) date, 2) EOD fin.res., 3) trades closed this day, 4) sum cmsn, 5) avg cmsn
    t2(0, 1) = "date"
    t2(0, 2) = "EOD_fin_res"
    t2(0, 3) = "days_cmsn"
    t2(0, 4) = "tds_closed"
    For r = 1 To UBound(t2, 1)
        t2(r, 4) = 0
    Next r
    t2(0, 5) = "tds_avg_cmsn"
' fill dates
    t2(1, 1) = SV(s_date_begin, 2) - 1
    For r = 2 To UBound(t2, 1)
        t2(r, 1) = t2(r - 1, 1) + 1
    Next r
' fill returns
    t2(1, 2) = SV(s_depo_ini, 2)
' fill sum_cmsn
    oc_fr = rc.Find(what:="Event log:", _
        after:=rc(ins_td_r + SV(s_trades, 2), 1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        searchorder:=xlByColumns, _
        searchdirection:=xlNext, _
        MatchCase:=False, _
        searchformat:=False).Row + 2          ' header row
    oc_lr = rc(oc_fr, 1).End(xlDown).Row
'
    ro_d = 1
    For r = 2 To UBound(t2, 1)
' compare dates
        If t2(r, 1) = Int(CDate(rc(oc_fr + ro_d, 1))) Then
            Do While t2(r, 1) = Int(CDate(rc(oc_fr + ro_d, 1)))
                If Left(rc(oc_fr + ro_d, 2), 6) = "Commis" Then
                    s = rc(oc_fr + ro_d, 3)
                    ArrCommis = Split(s, " ")
                    s = ArrCommis(UBound(ArrCommis))
                    s = Replace(s, ".", ",")
                    s = Replace(s, "[", "")
                    s = Replace(s, "]", "")
                    If CDbl(s) <> 0 Then
                        t2(r, 3) = CDbl(s)              ' enter cmsn
                    End If
                End If
                ' go to next trade in Trade log
                If ro_d + oc_fr < oc_lr Then
                    ro_d = ro_d + 1
                Else
                    Exit Do ' or exit do, when after last row processed
                End If
            Loop
        End If
    Next r
' calculate trades closed per each day
    c = 1   ' count trades for that day
    k = 1
    For r = 1 To UBound(t1, 1)
        If t1(r, 8) = "ERROR" Then  ' clean ERROR data
            t1(r, 4) = t1(r, 3)
            t1(r, 5) = 0
            t1(r, 6) = 0
            t1(r, 8) = t1(r, 7)
            t1(r, 9) = "ERROR"
        End If
        ' find this day in t2
        Do Until Int(CDate(t2(k, 1))) = Int(CDate(t1(r, 8)))
            k = k + 1
        Loop
        t2(k, 4) = t2(k, 4) + 1
    Next r
' ============== AVERAGE COMMISSION: begin ==============================
' calculate average commission for a trade
    c = 1
    For r = 1 To UBound(t2, 1)
        If t2(r, 4) > 0 Then
            t2(r, 5) = 0
            ' sum cmsns
            For k = c To r
                t2(r, 5) = t2(r, 5) + t2(k, 3)
            Next k
            t2(r, 5) = t2(r, 5) / t2(r, 4)
            c = r + 1
        End If
    Next r
' ============== AVERAGE COMMISSION: end ==============================
' fill t1 - equity curve
    k = 1
    For r = 1 To UBound(t1, 1)
        Do Until t2(k, 1) = Int(CDate(t1(r, 8)))
            k = k + 1
        Loop
        t1(r, 10) = t1(r - 1, 10) + t1(r, 5) - t2(k, 5)
        ' percent change
        t1(r, 11) = (t1(r, 10) - t1(r - 1, 10)) / t1(r - 1, 10)
    Next r
' hwm & mdd - high watermark & maximum drawdown
    t1(1, 12) = WorksheetFunction.Max(t1(0, 10), t1(1, 10))
    t1(1, 13) = (t1(1, 12) - t1(1, 10)) / t1(1, 12)
    SV(s_mdd, 2) = 0
    For r = 2 To UBound(t1, 1)
        t1(r, 12) = WorksheetFunction.Max(t1(r - 1, 12), t1(r, 10))
        t1(r, 13) = (t1(r, 12) - t1(r, 10)) / t1(r, 12)
        ' update mdd
        If t1(r, 13) > SV(s_mdd, 2) Then
            SV(s_mdd, 2) = t1(r, 13)
        End If
    Next r
' finish deposit
    SV(s_depo_fin, 2) = t1(UBound(t1, 1), 10)
    If SV(s_depo_fin, 2) < 0 Then
        SV(s_depo_fin, 2) = 0
    End If
' annual return = (1 + sv(i, s_depo_fin))
    SV(s_ar, 2) = (1 + (SV(s_depo_fin, 2) - SV(s_depo_ini, 2)) / SV(s_depo_ini, 2)) ^ (12 / SV(s_mns, 2)) - 1
' rf - recovery factor: annualized return / mdd
    If SV(s_ar, 2) > 0 Then
        If SV(s_mdd, 2) > 0 Then
            SV(s_rf, 2) = SV(s_ar, 2) / SV(s_mdd, 2)
        Else
            SV(s_rf, 2) = 999
        End If
    Else
        SV(s_rf, 2) = 0
    End If
' fill t2 EOD fin_res
    k = 1
    For r = 1 To UBound(t1, 1)
        If Int(CDate(t1(r, 8))) > t2(k, 1) Then
            Do Until t2(k, 1) = Int(CDate(t1(r, 8)))
                k = k + 1
                t2(k, 2) = t2(k - 1, 2)
            Loop
            t2(k, 2) = t1(r, 10)
        Else
            t2(k, 2) = t1(r, 10)
        End If
    Next r
    ' fill rest of empty days
    For r = k To UBound(t2, 1)
        If IsEmpty(t2(r, 2)) Then
            t2(r, 2) = t2(r - 1, 2)
        End If
    Next r
' r-square - rsq
    ReDim rsqX(1 To UBound(t2, 1))  ' the x-s and the y-s
    ReDim rsqY(1 To UBound(t2, 1))
    For r = 1 To UBound(t2, 1)
        rsqX(r) = t2(r, 1)
        rsqY(r) = t2(r, 2)
    Next r
    SV(s_rsq, 2) = WorksheetFunction.RSq(rsqX, rsqY)
    
End Sub

Sub Par_Bubblesort()

    Dim sj As Integer, sk As Integer
    Dim tmp_c1 As Variant, tmp_c2 As Variant
    
    For sj = 1 To UBound(Par, 1) - 1
        For sk = sj + 1 To UBound(Par, 1)
            If Par(sj, 1) > Par(sk, 1) Then
                tmp_c1 = Par(sk, 1)
                tmp_c2 = Par(sk, 2)
                Par(sk, 1) = Par(sj, 1)
                Par(sk, 2) = Par(sj, 2)
                Par(sj, 1) = tmp_c1
                Par(sj, 2) = tmp_c2
            End If
        Next sk
    Next sj
    
End Sub

Sub Open_Reports()

    ov(s_ov_htmls, 2) = UBound(oneFdFilesList)
    If ov(s_ov_htmls, 2) > maxHtmlCount Then
        MsgBox "GetStats can not process more than " & maxHtmlCount & " reports. Cancelling."
        openFail = True
        Exit Sub
    End If
    folderToSave = GetSaveFolder(oneFdFilesList(1))
    ReDim sM(0 To ov(s_ov_htmls, 2), 0 To 7)
    
End Sub

Function GetSaveFolder(ByVal file_path As String) As String

    Dim q As Integer, i As Integer
    
    q = 0
    i = Len(file_path) + 1
    Do Until q = 2
        i = i - 1
        If Mid(file_path, i, 1) = "\" Then
            q = q + 1
        End If
    Loop
    GetSaveFolder = Left(file_path, i)
    
End Function

Sub Prepare_sv_ov_fm()

    s_strat = 1
    s_ins = s_strat + 1
    s_tpm = s_ins + 1
    s_ar = s_tpm + 1
    s_mdd = s_ar + 1
    s_rf = s_mdd + 1
    s_rsq = s_rf + 1
    s_date_begin = s_rsq + 1
    s_date_end = s_date_begin + 1
    s_mns = s_date_end + 1
    s_trades = s_mns + 1
    s_win_pc = s_trades + 1
    s_pips = s_win_pc + 1
    s_avg_w2l = s_pips + 1
    s_avg_pip = s_avg_w2l + 1
    s_depo_ini = s_avg_pip + 1
    s_depo_fin = s_depo_ini + 1
    s_cmsn = s_depo_fin + 1
    s_link = s_cmsn + 1
    s_rep_type = s_link + 1
    ReDim SV(1 To s_rep_type, 1 To 2)
' SN
    SV(s_strat, 1) = "Strategy"
    SV(s_ins, 1) = "Instrument"
    SV(s_tpm, 1) = "Trades per month"
    SV(s_ar, 1) = "Annualized return, %"
    SV(s_mdd, 1) = "Maximum drawdown, %"
    SV(s_rf, 1) = "Recovery factor"
    SV(s_rsq, 1) = "R-squared"
    SV(s_date_begin, 1) = "Test begin date"
    SV(s_date_end, 1) = "Test end date"
    SV(s_mns, 1) = "Months"
    SV(s_trades, 1) = "Positions closed"
    SV(s_win_pc, 1) = "Winners, %"
    SV(s_pips, 1) = "Pips"
    SV(s_avg_w2l, 1) = "Avg. winner/loser, pips"
    SV(s_avg_pip, 1) = "Avg. trade, pips"
    SV(s_depo_ini, 1) = "Initial balance"
    SV(s_depo_fin, 1) = "End balance"
    SV(s_cmsn, 1) = "Commissions"
    SV(s_link, 1) = "Report size (MB), link"
    SV(s_rep_type, 1) = "Report type"
    SV(s_rep_type, 2) = repType

' overview
    s_ov_strat = 1
    s_ov_ins = s_ov_strat + 1
    s_ov_htmls = s_ov_ins + 1
    s_ov_mns = s_ov_htmls + 1
    s_ov_from = s_ov_mns + 1
    s_ov_to = s_ov_from + 1
    s_ov_params = s_ov_to + 1
'    s_ov_params_vbl = s_ov_params + 1
    s_ov_created = s_ov_params + 1
    s_ov_macro_ver = s_ov_created + 1
    ReDim ov(1 To s_ov_macro_ver, 1 To 2)
    ov(s_ov_strat, 1) = "Strategy"
    ov(s_ov_ins, 1) = "Instrument"
    ov(s_ov_htmls, 1) = "Reports processed"
    ov(s_ov_mns, 1) = "Hist. window, months"
    ov(s_ov_from, 1) = "Test start date"
    ov(s_ov_to, 1) = "Test end date"
    ov(s_ov_params, 1) = "Parameters count"
'    ov(s_ov_params_vbl, 1) = "Parameters variable"
    ov(s_ov_created, 1) = "Report generated"
    ov(s_ov_macro_ver, 1) = "Version"
' formats
    ' "yyyy-mm-dd"
    fm_date(1) = s_date_begin
    fm_date(2) = s_date_end
    ' "0.00"
    fm_0p00(1) = s_tpm
    fm_0p00(2) = s_rf
    fm_0p00(3) = s_rsq
    fm_0p00(4) = s_mns
    fm_0p00(5) = s_avg_w2l
    fm_0p00(6) = s_avg_pip
    fm_0p00(7) = s_depo_ini
    fm_0p00(8) = s_depo_fin
    fm_0p00(9) = s_cmsn
    fm_0p00(10) = s_link
    ' "0.00%"
    fm_0p00pc(1) = s_ar
    fm_0p00pc(2) = s_mdd
    fm_0p00pc(3) = s_win_pc
    ' color, bold, center
    fm_clr(1) = s_tpm
    fm_clr(2) = s_ar
    fm_clr(3) = s_mdd
    fm_clr(4) = s_rf
    fm_clr(5) = s_rsq
    
End Sub

Function ListFiles(ByVal sPath As String) As Variant
' Arg: folder path
' Return: list of files in the folder
    
    Dim vaArray() As Variant
    Dim i As Integer
    Dim oFile As Object
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFiles As Object
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    Set oFiles = oFolder.Files
    If oFiles.count = 0 Then Exit Function
    ReDim vaArray(1 To oFiles.count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = sPath & "\" & oFile.Name
        i = i + 1
    Next
    ListFiles = vaArray
    
End Function

Sub Pick_Strategy_Folder()
' sheet "backtest"
' sub shows file dialog, lets user pick strategy folder

    Dim fd As FileDialog
    Dim stratName As String
    
    Call Init_Pick_Strategy_Folder(stratFdRng, stratNmRng)
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Pick strategy folder"
'        .ButtonName = "OK"
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    stratFdRng = fd.SelectedItems(1)
    stratName = fd.SelectedItems(1)
    stratName = Right(stratName, Len(stratName) - InStrRev(stratName, "\", -1, vbTextCompare))
    stratNmRng = stratName
    Columns(1).AutoFit
    
End Sub

Sub DeSelect_Instruments()

    Dim cell As Range
    
    Application.ScreenUpdating = False
    Call Init_DeSelect_Instruments(setWs, btWs, setC, btC, _
                selectAll, instrumentsList)
    If selectAll Then
        For Each cell In instrumentsList
            cell = True
        Next cell
    Else
        For Each cell In instrumentsList
            cell = False
        Next cell
    End If
    Application.ScreenUpdating = True
    
End Sub

Sub Clear_Ready_Reports()
    
    Dim lastRow As Integer
    Dim Rng As Range
    
    Call Init_Clear_Ready_Reports(btWs, btC, _
                upperRow, leftCol, rightCol)
    lastRow = btC(btWs.Rows.count, leftCol).End(xlUp).Row
    If lastRow = upperRow - 1 Then
        Exit Sub
    End If
    Set Rng = btWs.Range(btC(upperRow, leftCol), btC(lastRow, rightCol))
    Rng.Clear
    
End Sub



' MODULE: Inits
Option Explicit

Const addInFName As String = "GetStats_BackTest_v1.21.xlsm"
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



' MODULE: SharpeRatio
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
        MsgBox "No files picked!"
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
    
    MsgBox "Done. Please save """ & wbA.Name & """ as needed.", , "GetStats Pro"
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
            
'            cagr = (1 + net_return) ^ (365 / days_count) - 1
            On Error Resume Next
            cagr = (1 + net_return) ^ (365 / days_count) - 1
            If Err.Number = 5 Then
                cagr = 0
            End If
            On Error GoTo 0
            
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
    Dim msgAnswer As Variant
    
    msgAnswer = MsgBox("Separate window?", vbYesNoCancel)
    If msgAnswer = vbCancel Then
        Exit Sub
    End If
    
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
        If msgAnswer = vbNo Then
            Call Params_To_Summary
        End If
        Call Sharpe_to_all
        
        Set rg = ActiveCell.CurrentRegion
        rg.Copy tC(insertRow, 1)
        
        insertRow = tC(tC(tWs.Rows.count, 1).End(xlUp).Row, 1).Row + 2
    
        wb.Close savechanges:=False
    Next i
    
    Application.ScreenUpdating = True
    
End Sub



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
'    Call PrintModulesCode
'    Call PrintAllContainers
    
    MsgBox "Code exported.", , "GetStats"
    
End Sub

Sub DeleteAndMake()
    
    Dim childA As String
    Dim childB As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    parentFolder = ThisWorkbook.path
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
    
    If Dir(ThisWorkbook.path & newFile, vbDirectory) = vbNullString Then
        MkDir ThisWorkbook.path & newFile
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


' MODULE: clsTimer
Option Explicit

Public dtTime0 As Date
Public dtTime1 As Date
Public strSubOrFunc As String

Private Sub Class_Initialize(ByVal argSubOrFunc As String)
    dtTime0 = Timer
    strSubOrFunc = argSubOrFunc
End Sub

Private Sub Class_Terminate()
    dtTime1 = Timer
    Debug.Print "terminated " & suborfunc & " - " & (dtTime1 - dtTime0)
End Sub
'public property let strSubOrFunc as String


