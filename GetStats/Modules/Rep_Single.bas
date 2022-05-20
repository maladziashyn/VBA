Attribute VB_Name = "Rep_Single"
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
'
' RIBBON > BUTTON "Main"
'
'    On Error Resume Next
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
    
    all_zeros = False
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
    SV(s_date_begin, 2) = CDate(rc(ins_td_r - 7, 2))
' Test end
'    SV(s_date_end, 2) = Int(rc(ins_td_r - 4, 2))
    SV(s_date_end, 2) = CDate(Int(rc(ins_td_r - 4, 2)))
' Months
    SV(s_mns, 2) = (SV(s_date_end, 2) - SV(s_date_begin, 2)) * 12 / 365
' TPM
    SV(s_tpm, 2) = Round(SV(s_trades, 2) / SV(s_mns, 2), 2)
' Initial deposit
    SV(s_depo_ini, 2) = rc(5, 2)
' Commissions
    SV(s_cmsn, 2) = rc(8, 2)
' File size
    SV(s_link, 2) = Round(FileLen(rep_adr) / 1024 ^ 2, 2)
    If SV(s_trades, 2) = 0 Then
        all_zeros = True
        SV(s_depo_fin, 2) = CDbl(rc(6, 2)) ' Finish deposit
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
                If rc(oc_fr + ro_d, 2) = "Commissions" Then
                    s = rc(oc_fr + ro_d, 3)
                    s = Right(s, Len(s) - 13)
                    s = Left(s, Len(s) - 1)
                    s = Replace(s, ".", ",", 1, 1, 1)
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
    c(13, 1) = "Pips gained"
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
