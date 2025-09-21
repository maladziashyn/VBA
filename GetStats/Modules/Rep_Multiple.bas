Attribute VB_Name = "Rep_Multiple"
Option Explicit
Option Base 1
    Const addin_file_name As String = "GetStats_BackTest_v1.27.xlsm"
    Const rep_type As String = "GS_Pro_Single_Core"
    Const macro_ver As String = "GetStats Pro v1.27"
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

Private Sub ProcessManySelected()
    
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
'' check window
'    Call GSPR_Check_Window
'' save
'    Call GSPRM_Save_To_Desktop
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
Private Sub GSPRM_Open_Reports()
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "GetStats: Select HTML reports (max. " & max_htmls & ")"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "JForex back-test reports", "*.html"
'        .ButtonName = "Вперед"
    End With
    If fd.Show = 0 Then
        open_fail = True
        MsgBox "No files picked!"
        Exit Sub
    End If
    ov(s_ov_htmls, 2) = fd.SelectedItems.count
    If ov(s_ov_htmls, 2) > max_htmls Then
        MsgBox "GetStats cannot process more than " & max_htmls & " reports. Exiting."
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
            sta = "Working on report " & i & " (" & ov(s_ov_htmls, 2) & ")."
        Else
            sta = "Working on report " & i & " (" & ov(s_ov_htmls, 2) & "). Est. time remaining " & time_rem
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
'    SV(s_date_end, 2) = CDate(Int(rc(ins_td_r - 4, 2).Value))    ' *! removed "-1"
' Months
    SV(s_mns, 2) = (SV(s_date_end, 2) - SV(s_date_begin, 2)) * 12 / 365
' TPM
'    If rc(ins_td_r, 2) = 0 Then
'        SV(s_tpm, 2) = 0
'    Else
        SV(s_tpm, 2) = Round(SV(s_trades, 2) / SV(s_mns, 2), 2)
'    End If
' Initial deposit
    SV(s_depo_ini, 2) = CDbl(Replace(rc(5, 2), "’", ""))
'' Finish deposit
'    sv(s_depo_fin, 2) = CDbl(rc(6, 2))
' Commissions
    SV(s_cmsn, 2) = CDbl(Replace(rc(8, 2), "’", ""))
    
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
    sM(0, 1) = "№_link"
    sM(0, 2) = "trades_per_month"
    sM(0, 3) = "annualized_return"
    sM(0, 4) = "max_drawdown"
    sM(0, 5) = "recovery_factor"
    sM(0, 6) = "r_squared"
    sM(0, 7) = "avg_trade_pips"
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
    
    Application.StatusBar = "Saving..."
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


