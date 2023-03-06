Attribute VB_Name = "Rep_Extra"
Option Explicit
Option Base 1
    Const rep_type As String = "GS_Pro_Single_Core"
    Dim ch_rep_type As Boolean
' macro version
    Const macro_name As String = "GetStats Pro v1.22"
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


