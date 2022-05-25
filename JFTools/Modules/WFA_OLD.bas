Attribute VB_Name = "WFA_OLD"
Option Explicit
'    Const addInName As String = "GetStats_WFA_v0.18.xlsm"
'    Const wsSettingsName As String = "settings"
'    Const wsMergeName As String = "merge"
'
'    Dim timeStarted As Date, timeEnded As Date
    
    
    
' ========================================
' OLD VARIABLES
    Const addInName As String = "JFTools_0.01.xlsm"
    Const wsSettingsName As String = "WFA OLD"
    Const wsMergeName As String = "WFA Merge"
    
    Const tpm_col As Integer = 6
    Const rf_col As Integer = 8
    Const rsq_col As Integer = 10
    Const spans_col As Integer = 12
    Const init_row As Integer = 5
' Merge <
    Const srcFdRow As Integer = 2
    Const srcFdCol As Integer = 1
' Merge >
    Dim target_mdd As Double
    Dim target_fdm As Double
    Dim fract_mult As Double
    Dim param_combin() As Variant
    Dim spans_ar_rf_r2() As Variant
    Dim spans_ct As Integer
    Dim wbks_list() As Variant
    Dim mega_arr() As Variant
    Dim spans_table() As Integer
    Dim date_slots() As Variant
    Dim wbks_count As Integer
    Dim avail_start As Long
    Dim avail_end As Long
    Dim time_started As Date, time_ended As Date
    Dim highest_ar As Double, avg_ar As Double, pc_pos_os As Double
    Dim tg_chat_id As String
Private Sub WFA_Main_OLD()
    Application.ScreenUpdating = False
    time_started = Now
    Call WFA_Params_Spans_Combinations
    wbks_list = WFA_Grab_Wbks_List
    Call WFA_Scan_WB_List
    Call WFA_PrintOut_MegaArr
    Application.ScreenUpdating = True
End Sub
Private Sub Level_II()
    Const levTwoHpsFCol As Integer = 15
    Const levTwoHpsLCol As Integer = 17
    Const levTwoHpsFRow As Integer = 15
    Dim timeBegin As Date, timeEnd As Date
    Dim timeElapsed As String
    Dim tgMsg As String
    Dim levTwoHpsLRow As Integer
    Dim wbAddin As Workbook
    Dim wsAddin As Worksheet
    Dim cAddin As Range
    Dim z As Integer
'    Dim tgFileReady As Integer
    Dim tgAllFiles As Integer
    Dim hpsRng As Range
    Dim tgChatID As String
    
    Application.ScreenUpdating = False
    timeBegin = Now
    Set wbAddin = ActiveWorkbook
    Set wsAddin = wbAddin.Sheets(wsSettingsName)
    Set cAddin = wsAddin.Cells
    levTwoHpsLRow = cAddin(wsAddin.rows.Count, levTwoHpsFCol).End(xlUp).Row
    tgChatID = cAddin(8, 4)
'    tgFileReady = 1
    tgAllFiles = levTwoHpsLRow - levTwoHpsFRow + 1
    For z = levTwoHpsFRow To levTwoHpsLRow
        ' copy&paste hps
        Set hpsRng = wsAddin.Range(cAddin(z, levTwoHpsFCol), cAddin(z, levTwoHpsLCol))
        hpsRng.Copy cAddin(5, 12)
        
        'launch MAIN
        time_started = Now
        Call WFA_Params_Spans_Combinations
        wbks_list = WFA_Grab_Wbks_List
        Call WFA_Scan_WB_List
        Call WFA_PrintOut_MegaArr
'        ' Telegram message
'        tgMsg = "WFA " & tgFileReady & " (" & tgAllFiles & ") done"
'        Call Send_TG_Message(tgMsg, tgChatID)
'        tgFileReady = tgFileReady + 1
    Next z
    timeEnd = Now
    timeElapsed = Time_elapsed(timeBegin, timeEnd)
'    tgMsg = "Все HPs посчитаны" & vbNewLine & _
'            "Компьютер: " & cAddin(6, 4) & vbNewLine & _
'            "Начало: " & timeBegin & vbNewLine & _
'            "Окончание: " & timeEnd & vbNewLine & _
'            "Времени: " & timeElapsed & vbNewLine & _
'            "Файлов: " & tgAllFiles & vbNewLine & _
'            "Папка: " & cAddin(7, 4)
'    Call Send_TG_Message(tgMsg, tgChatID)
    Application.ScreenUpdating = True
End Sub
Private Sub WFA_Params_Spans_Combinations()
' Subroutine fills param_combin array with data
' from Excel sheet "combinations"
    Dim cur_tpm_row As Integer
    Dim cur_rf_row As Integer
    Dim cur_rsq_row As Integer
    Dim max_tpm_row As Integer
    Dim max_rf_row As Integer
    Dim max_rsq_row As Integer
    Dim ubnd As Integer
    Dim cc As Range
    
    Set cc = Workbooks(addInName).Sheets(wsSettingsName).Cells
    avail_start = cc(2, 4)
    avail_end = cc(3, 4)
    target_mdd = cc(4, 4)
    target_fdm = cc(5, 4)
    cur_tpm_row = init_row
    cur_rf_row = init_row
    cur_rsq_row = init_row
    max_tpm_row = cc(init_row - 1, tpm_col).End(xlDown).Row
    max_rf_row = cc(init_row - 1, rf_col).End(xlDown).Row
    max_rsq_row = cc(init_row - 1, rsq_col).End(xlDown).Row
    ubnd = 0
    param_combin = WFA_Init_Param_Combin_Array()
    Do Until cur_tpm_row = max_tpm_row And _
             cur_rf_row = max_rf_row And _
             cur_rsq_row = max_rsq_row
        Call WFA_Add_next_val(cc, ubnd, cur_tpm_row, cur_rf_row, cur_rsq_row)
        If cur_tpm_row = max_tpm_row Then
            cur_tpm_row = init_row
            If cur_rf_row = max_rf_row Then
                cur_rf_row = init_row
                cur_rsq_row = cur_rsq_row + 1
            Else
                cur_rf_row = cur_rf_row + 1
            End If
        Else
            cur_tpm_row = cur_tpm_row + 1
        End If
    Loop
    Call WFA_Add_next_val(cc, ubnd, cur_tpm_row, cur_rf_row, cur_rsq_row)
' spans & metrics array
    spans_ar_rf_r2 = WFA_Init_Spans_Metrics_Arr(cc, UBound(param_combin, 2))
End Sub
Private Function WFA_Init_Param_Combin_Array() As Variant
    Dim A() As Variant
    Dim i As Integer
    
    ReDim A(1 To 6, 0 To 0)
    A(1, 0) = "TPM min"
    A(2, 0) = "TPM max"
    A(3, 0) = "RF min"
    A(4, 0) = "RF max"
    A(5, 0) = "R2 min"
    A(6, 0) = "R2 max"
    WFA_Init_Param_Combin_Array = A
End Function
Private Sub WFA_Add_next_val(ByVal cc As Range, ByRef ubnd As Integer, _
                 ByVal cur_tpm_row As Integer, ByVal cur_rf_row As Integer, ByVal cur_rsq_row As Integer)
    ubnd = ubnd + 1
    ReDim Preserve param_combin(1 To UBound(param_combin, 1), 0 To ubnd)
    param_combin(1, ubnd) = cc(cur_tpm_row, tpm_col)
    param_combin(2, ubnd) = cc(cur_tpm_row, tpm_col + 1)
    param_combin(3, ubnd) = cc(cur_rf_row, rf_col)
    param_combin(4, ubnd) = cc(cur_rf_row, rf_col + 1)
    param_combin(5, ubnd) = cc(cur_rsq_row, rsq_col)
    param_combin(6, ubnd) = cc(cur_rsq_row, rsq_col + 1)
End Sub
Private Function WFA_Init_Spans_Metrics_Arr(ByVal cc As Range, _
                                    ByVal combinations_count As Integer) As Variant
' function takes arguments:
' 1. Cells with parameters
' 2. Combinations count (upper boundary of "param_combin"(dim2 - rows) array)
    Dim arr(1 To 3) As Variant
        ' 1. AR - Annualized Return
        ' 2. RF - Recovery Factor
        ' 3. R2 - R-Squared
    Dim tmp_arr() As Variant
    Dim i As Integer, j As Integer
    Dim rows_ct As Integer
    
    spans_ct = cc(3, spans_col).End(xlDown).Row - init_row + 1
    ReDim tmp_arr(1 To spans_ct, 0 To combinations_count)
    ReDim spans_table(1 To 2, 1 To spans_ct)
    For i = LBound(tmp_arr, 1) To UBound(tmp_arr, 1)
        j = init_row - 1 + i
        tmp_arr(i, 0) = cc(j, spans_col) & "/" & cc(j, spans_col + 1)
        spans_table(1, i) = cc(j, spans_col)
        spans_table(2, i) = cc(j, spans_col + 1)
    Next i
    date_slots = WFA_Generate_Date_Slots
    For i = 1 To 3
        arr(i) = tmp_arr
    Next i
    WFA_Init_Spans_Metrics_Arr = arr
End Function
Private Function WFA_Generate_Date_Slots() As Variant
    Dim i As Integer
    Dim arr() As Variant
    
    ReDim arr(1 To UBound(spans_table, 2))
    For i = LBound(arr) To UBound(arr)
        arr(i) = WFA_Generate_Four_Dates(spans_table(1, i), spans_table(2, i))
    Next i
    WFA_Generate_Date_Slots = arr
End Function
Private Function WFA_Generate_Four_Dates(ByVal is_wks As Integer, _
                                ByVal os_wks As Integer) As Variant
' function returns (1 to 4, 1 to Rows) array of dates:
' col 1-2: IS from/to, col 3-4: OS from/to
    Dim arr() As Variant
    Dim i As Integer
    
    ReDim arr(1 To 4, 1 To 1)
    i = 1
    arr(1, i) = avail_start
    arr(2, i) = arr(1, i) + 7 * is_wks - 1
    arr(3, i) = arr(2, i) + 1
    arr(4, i) = arr(3, i) + 7 * os_wks - 1
    Do While arr(2, i) + 7 * os_wks < avail_end
        i = i + 1
        ReDim Preserve arr(1 To 4, 1 To i)
        arr(1, i) = arr(1, i - 1) + 7 * os_wks
        arr(2, i) = arr(1, i) + 7 * is_wks - 1
        arr(3, i) = arr(2, i) + 1
        arr(4, i) = arr(3, i) + 7 * os_wks - 1
    Loop
    WFA_Generate_Four_Dates = arr
End Function
Private Sub WFA_Scan_WB_List()
    Dim this_wbk As Integer, this_sheet As Integer
    Dim this_param As Integer, this_span As Integer, this_date_slot As Integer
    Dim files_count As Integer, files_remaining As Integer
    Dim wb As Workbook
    Dim is_single_tradeset() As Variant
    Dim is_single_ts_pfl() As Variant
    Dim os_single_tradeset() As Variant
    Dim x() As Variant
    Dim status_str As String
    Dim bt_date_slot_winner_id_shnm As Variant
    Dim pfl_date_1 As Date
'    Dim sh_count As Integer
    Dim bt_winner_ShName As String
    Dim set_string As String
    Dim sheet_in_ram() As Variant
    ' reinit
    Dim raw_sets_pfls(1 To 2, 0 To 0) As Variant
    
    files_count = UBound(wbks_list)
    files_remaining = files_count
    
    mega_arr = WFA_Init_Mega_Array
    For this_wbk = LBound(wbks_list) To UBound(wbks_list)
        status_str = "Files remaining: " & files_remaining & " (" & files_count & ")."
        Application.StatusBar = status_str
        Set wb = Workbooks.Open(wbks_list(this_wbk))

' ### init parts of mega_arr
        For this_sheet = 3 To wb.Sheets.Count
            Application.StatusBar = status_str & " Sheet: " & this_sheet & " (" & wb.Sheets.Count & ")."
            sheet_in_ram = WFA_Load_Report_to_RAM(wb.Sheets(this_sheet))
            For this_param = 1 To UBound(param_combin, 2)
                For this_span = LBound(spans_table, 2) To UBound(spans_table, 2)
                    For this_date_slot = LBound(date_slots(this_span), 2) To UBound(date_slots(this_span), 2)
                        is_single_tradeset = WFA_Get_Tradeset_From_RAM(sheet_in_ram, date_slots(this_span)(1, this_date_slot), date_slots(this_span)(2, this_date_slot), wb.Name, wb.Sheets(this_sheet).Name)
                        is_single_ts_pfl = WFA_Get_Pfl_From_TSet(is_single_tradeset, date_slots(this_span)(1, this_date_slot), date_slots(this_span)(2, this_date_slot), wb.Sheets(this_sheet).Name)
                        mega_arr(this_span, this_param)(1)(this_date_slot)(1) = WFA_Update_Raw_Pfls(mega_arr(this_span, this_param)(1)(this_date_slot)(1), is_single_ts_pfl)
                    Next this_date_slot
                Next this_span
            Next this_param
        Next this_sheet

        ' move best set to BT best_arr = append
        For this_param = 1 To UBound(param_combin, 2)
            For this_span = LBound(spans_table, 2) To UBound(spans_table, 2)
                For this_date_slot = LBound(date_slots(this_span), 2) To UBound(date_slots(this_span), 2)
                    bt_winner_ShName = WFA_Get_Winner_ShName(mega_arr(this_span, this_param)(1)(this_date_slot)(1), this_param)
                    If bt_winner_ShName <> "0" Then
                        sheet_in_ram = WFA_Load_Report_to_RAM(wb.Sheets(bt_winner_ShName))
                        is_single_tradeset = WFA_Get_Tradeset_From_RAM(sheet_in_ram, date_slots(this_span)(1, this_date_slot), date_slots(this_span)(2, this_date_slot), wb.Name, bt_winner_ShName)
                        mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1) = WFA_Add_To_Top_BT_FD_Sets(mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1), is_single_tradeset)
                        os_single_tradeset = WFA_Get_Tradeset_From_Sheet(wb.Sheets(bt_winner_ShName), date_slots(this_span)(3, this_date_slot), date_slots(this_span)(4, this_date_slot), wb.Name)
                        mega_arr(this_span, this_param)(2)(this_date_slot)(1) = WFA_Add_To_Top_BT_FD_Sets(mega_arr(this_span, this_param)(2)(this_date_slot)(1), os_single_tradeset)
                        mega_arr(this_span, this_param)(1)(this_date_slot)(1) = WFA_Init_2D_Arr(1, 7, 0, 0)   ' (1 to 7, 0 to 0)
                    End If
                Next this_date_slot
            Next this_span
        Next this_param

        wb.Close savechanges:=False
        files_remaining = files_remaining - 1
    Next this_wbk
    Application.StatusBar = False
    
    For this_param = 1 To UBound(param_combin, 2)
        Application.StatusBar = "Param_combin: " & this_param & " (" & UBound(param_combin, 2) & ")."
        For this_span = LBound(spans_table, 2) To UBound(spans_table, 2)
            For this_date_slot = LBound(date_slots(this_span), 2) To UBound(date_slots(this_span), 2)
            ' IS: bubble sort, add headers/cumsum, profile
                set_string = "BT_par-" & this_param & "_spa-" & this_span & "_slo-" & this_date_slot
                If UBound(mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1), 2) = 0 Then
                    mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1) = WFA_Add_Header_CumSum(mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1), False)
                Else
                    mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1) = WFA_Sort_CloseDate(mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1))
            ' ### BEGIN Alter Fraction to Target MDD
                    ' find multiplier
                    fract_mult = WFA_Get_Fraction_Multiplier(mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1))
                    ' rewrite fraction (pc_chg column)
                    mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1) = WFA_Alter_Fraction(mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1), fract_mult)
            ' ### END Alter
                    mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1) = WFA_Add_Header_CumSum(mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1), True)
                End If
                mega_arr(this_span, this_param)(1)(this_date_slot)(2)(2) = WFA_Get_Pfl_From_TSet(mega_arr(this_span, this_param)(1)(this_date_slot)(2)(1), date_slots(this_span)(1, this_date_slot), date_slots(this_span)(2, this_date_slot), set_string)
            ' OS: bubble sort, profile
                set_string = "FD_par-" & this_param & "_spa-" & this_span & "_slo-" & this_date_slot
                If UBound(mega_arr(this_span, this_param)(2)(this_date_slot)(1), 2) = 0 Then
                    mega_arr(this_span, this_param)(2)(this_date_slot)(1) = WFA_Add_Header_CumSum(mega_arr(this_span, this_param)(2)(this_date_slot)(1), False)
                Else
                    mega_arr(this_span, this_param)(2)(this_date_slot)(1) = WFA_Sort_CloseDate(mega_arr(this_span, this_param)(2)(this_date_slot)(1))
            ' ### BEGIN Alter Fraction, using Backtest multiplier
                    mega_arr(this_span, this_param)(2)(this_date_slot)(1) = WFA_Alter_Fraction(mega_arr(this_span, this_param)(2)(this_date_slot)(1), fract_mult)
            ' ### END Alter Fraction, using Backtest multiplier
                    mega_arr(this_span, this_param)(2)(this_date_slot)(1) = WFA_Add_Header_CumSum(mega_arr(this_span, this_param)(2)(this_date_slot)(1), True)
                End If
'                mega_arr(this_span, this_param)(2)(this_date_slot)(2) = WFA_Get_Pfl_From_TSet(mega_arr(this_span, this_param)(2)(this_date_slot)(1), date_slots(this_span)(3, this_date_slot), date_slots(this_span)(4, this_date_slot), set_string)
                '1
                If date_slots(this_span)(4, this_date_slot) > avail_end Then
                    pfl_date_1 = avail_end
                Else
                    pfl_date_1 = date_slots(this_span)(4, this_date_slot)
                End If
                mega_arr(this_span, this_param)(2)(this_date_slot)(2) = WFA_Get_Pfl_From_TSet(mega_arr(this_span, this_param)(2)(this_date_slot)(1), date_slots(this_span)(3, this_date_slot), pfl_date_1, set_string)
                ' add to Compiled
                mega_arr(this_span, this_param)(3)(1) = WFA_Add_To_Top_BT_FD_Sets(mega_arr(this_span, this_param)(3)(1), mega_arr(this_span, this_param)(2)(this_date_slot)(1))
            Next this_date_slot
            ' COMPILED: bubble sort, add header, equity curve & calc stats
            mega_arr(this_span, this_param)(3)(1) = WFA_Sort_CloseDate(mega_arr(this_span, this_param)(3)(1))
            If UBound(mega_arr(this_span, this_param)(3)(1), 2) = 0 Then
                mega_arr(this_span, this_param)(3)(1) = WFA_Add_Header_CumSum(mega_arr(this_span, this_param)(3)(1), False)
            Else
                mega_arr(this_span, this_param)(3)(1) = WFA_Add_Header_CumSum(mega_arr(this_span, this_param)(3)(1), True)
            End If
            ' calc COMPILED stats - pfl
            set_string = "CU_par-" & this_param & "_spa-" & this_span
'            mega_arr(this_span, this_param)(3)(2) = WFA_Get_Pfl_From_TSet(mega_arr(this_span, this_param)(3)(1), date_slots(this_span)(3, 1), date_slots(this_span)(4, UBound(date_slots(this_span), 2)), set_string)
            mega_arr(this_span, this_param)(3)(2) = WFA_Get_Pfl_From_TSet(mega_arr(this_span, this_param)(3)(1), date_slots(this_span)(3, 1), avail_end, set_string)
        Next this_span
    Next this_param
    Application.StatusBar = False
End Sub
Private Function WFA_Get_Fraction_Multiplier(ByVal arr As Variant) As Double
' arr(1 to 4, 0 to trades)
    Const init_lower_mult As Double = 0
    Const init_upper_mult As Double = 10
    Dim returns() As Variant
    Dim i As Integer
    Dim lower_mult As Double, upper_mult As Double, mid_mult As Double
    Dim lower_mdd As Double, upper_mdd As Double, mid_mdd As Double
    Dim mdd_delta As Double

' Collect returns
    ReDim returns(0 To UBound(arr, 2))
    For i = 1 To UBound(returns)
        returns(i) = arr(4, i)
    Next i
' GET Upper & Lower multiplicators
    lower_mult = init_lower_mult
    upper_mult = init_upper_mult
    Do Until WFA_Calc_MDD_Only(returns, upper_mult) > target_mdd
        lower_mult = upper_mult
        upper_mult = upper_mult * 2
    Loop
    mid_mult = (lower_mult + upper_mult) / 2
' NARROW search
    mdd_delta = target_fdm * 2  ' initialize delta
    Do Until mdd_delta <= target_fdm
        mid_mdd = WFA_Calc_MDD_Only(returns, mid_mult)
        mdd_delta = Abs(mid_mdd - target_mdd)
        If mdd_delta <= target_fdm Then
            Exit Do
        Else
            If mid_mdd > target_mdd Then
                upper_mult = mid_mult
            ElseIf mid_mdd < target_mdd Then
                lower_mult = mid_mult
            Else
                Exit Do
            End If
            mid_mult = (lower_mult + upper_mult) / 2
        End If
    Loop
    WFA_Get_Fraction_Multiplier = mid_mult
End Function
Private Function WFA_Calc_MDD_Only(ByVal returns As Variant, _
                           ByVal multiplier As Double) As Double
    Dim eh() As Variant
    Dim dd() As Variant
    Dim i As Integer
    
    ReDim eh(1 To 2, 0 To UBound(returns))
    eh(1, 0) = 1   ' equity
    eh(2, 0) = 1   ' hwm
    ReDim dd(0 To UBound(returns))
    dd(0) = 0   ' dd
    For i = 1 To UBound(eh, 2)
        eh(1, i) = eh(1, i - 1) * (1 + multiplier * returns(i))     ' equity
        eh(2, i) = WorksheetFunction.Max(eh(2, i - 1), eh(1, i))
        dd(i) = (eh(2, i) - eh(1, i)) / eh(2, i)
    Next i
    WFA_Calc_MDD_Only = WorksheetFunction.Max(dd)
End Function
Private Function WFA_Alter_Fraction(ByVal arr As Variant, _
                            ByVal multiplier As Double) As Variant
    Dim result_arr() As Variant
    Dim i As Integer
    
    result_arr = arr
    For i = 1 To UBound(result_arr, 2)
        result_arr(4, i) = result_arr(4, i) * multiplier
    Next i
    WFA_Alter_Fraction = result_arr
End Function
Private Function WFA_Init_1D_Arr(ByVal d1_1 As Integer, ByVal d1_2 As Integer) As Variant
    Dim A() As Variant
    
    ReDim A(d1_1 To d1_2)
    WFA_Init_1D_Arr = A
End Function
Private Function WFA_Init_2D_Arr(ByVal d1_1 As Integer, ByVal d1_2 As Integer, ByVal d2_1 As Integer, ByVal d2_2 As Integer) As Variant
    Dim A() As Variant
    
    ReDim A(d1_1 To d1_2, d2_1 To d2_2)
    WFA_Init_2D_Arr = A
End Function
Private Function WFA_Init_Mega_Array() As Variant
    Dim arr() As Variant
    Dim this_span As Integer, this_param As Integer
    Dim i As Integer
    
    ReDim arr(1 To UBound(spans_table, 2), 1 To UBound(param_combin, 2))
    For this_span = LBound(arr, 1) To UBound(arr, 1)
        For this_param = LBound(arr, 2) To UBound(arr, 2)
            arr(this_span, this_param) = WFA_Init_1D_Arr(1, 3)
            ' 1. IS date_slots
            arr(this_span, this_param)(1) = WFA_Init_1D_Arr(1, UBound(date_slots(this_span), 2))
            For i = LBound(arr(this_span, this_param)(1)) To UBound(arr(this_span, this_param)(1))
                arr(this_span, this_param)(1)(i) = WFA_Init_1D_Arr(1, 2)
                arr(this_span, this_param)(1)(i)(1) = WFA_Init_2D_Arr(1, 7, 0, 0)   ' raw pfls
                arr(this_span, this_param)(1)(i)(2) = WFA_Init_1D_Arr(1, 2)
                arr(this_span, this_param)(1)(i)(2)(1) = WFA_Init_2D_Arr(1, 4, 0, 0)
            Next i
            ' 2. OS date_slots
            arr(this_span, this_param)(2) = WFA_Init_1D_Arr(1, UBound(date_slots(this_span), 2))
            For i = LBound(arr(this_span, this_param)(2)) To UBound(arr(this_span, this_param)(2))
                arr(this_span, this_param)(2)(i) = WFA_Init_1D_Arr(1, 2)
                arr(this_span, this_param)(2)(i)(1) = WFA_Init_2D_Arr(1, 4, 0, 0)
            Next i
            ' 3. OS compiled
            arr(this_span, this_param)(3) = WFA_Init_1D_Arr(1, 2)
            arr(this_span, this_param)(3)(1) = WFA_Init_2D_Arr(1, 4, 0, 0)
        Next this_param
    Next this_span
    WFA_Init_Mega_Array = arr
End Function
Private Function WFA_Load_Report_to_RAM(ByVal ws As Worksheet) As Variant
' Function loads html report from sheet to RAM
' Returns (1 To 3, 0 To trades_count) array
    Dim arr() As Variant
    Dim last_row As Integer
    Dim wsC As Range
    Dim i As Integer, j As Integer
    
    Set wsC = ws.Cells
    last_row = wsC(ws.rows.Count, 4).End(xlUp).Row
    ReDim arr(1 To 3, 0 To last_row - 1)
    For i = 1 To last_row
        j = i - 1
        arr(1, j) = wsC(i, 9)   ' open date
        arr(2, j) = wsC(i, 10)  ' close date
        arr(3, j) = wsC(i, 13)  ' return
    Next i
    WFA_Load_Report_to_RAM = arr
End Function
Private Function WFA_Get_Tradeset_From_RAM(ByVal sheet_ram As Variant, _
                                   ByVal date_0 As Date, _
                                   ByVal date_1 As Date, _
                                   ByVal book_name As String, _
                                   ByVal sheet_name As String) As Variant
' Function returns 4-column set of trades:
' 1. open date
' 2. close date
' 3. comment (new: bookName_sheetName)
' 4. return
    Dim result_arr() As Variant
    Dim ubnd As Integer, new_ubnd As Integer
    Dim comment_str As String
    Dim i As Integer
    
    comment_str = Left(book_name, Len(book_name) - 5) & "_" & sheet_name
    ubnd = UBound(sheet_ram, 2)
    ReDim result_arr(1 To 4, 0 To 0)
    If ubnd = 0 Then
        WFA_Get_Tradeset_From_RAM = result_arr
        Exit Function
    End If
    i = 1
    ' 1 - open date, 2 - close date !
    Do While Int(sheet_ram(2, i)) <= date_1 And i <= ubnd   ' while open date <= last_date
        If Int(sheet_ram(1, i)) >= date_0 And Int(sheet_ram(2, i)) <= date_1 Then
            new_ubnd = UBound(result_arr, 2) + 1
            ReDim Preserve result_arr(1 To 4, 0 To new_ubnd)
            result_arr(1, new_ubnd) = sheet_ram(1, i)
            result_arr(2, new_ubnd) = sheet_ram(2, i)
            result_arr(3, new_ubnd) = comment_str
            result_arr(4, new_ubnd) = sheet_ram(3, i)
        End If
        If i = ubnd Then
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    WFA_Get_Tradeset_From_RAM = result_arr
End Function
Private Function WFA_Get_Pfl_From_TSet(ByVal trades_set As Variant, _
                               ByVal date_0 As Date, _
                               ByVal date_1 As Date, _
                               ByVal ws_name As String) As Variant
    Dim ubnd As Long
    Dim perf_profile_arr(1 To 7) As Variant
            ' 1. sheet name ("005")
            ' 2. tpm - trades per month
            ' 3. ar - annualized return
            ' 4. mdd - maximum drawdown
            ' 5. rf - recovery factor
            ' 6. r2 - R-squared
            ' 7. usable = True or False
    Dim clnd_days() As Variant
    Dim daily_eq() As Variant
    Dim equity_hwm_arr() As Double
    Dim dd_arr() As Double
    Dim cd_days As Integer
    Dim net_return As Double
    Dim i As Long

    ubnd = UBound(trades_set, 2)
    If ubnd = 0 Then
        perf_profile_arr(7) = False
        WFA_Get_Pfl_From_TSet = perf_profile_arr
        Exit Function
    End If
' build equity curve
    ReDim clnd_days(0 To 0)
    ReDim daily_eq(0 To 0)
    ReDim equity_hwm_arr(0 To ubnd, 1 To 2) ' 1. equity; 2. HWM
    ReDim dd_arr(1 To ubnd)
    equity_hwm_arr(0, 1) = 1
    equity_hwm_arr(0, 2) = 1
    For i = 1 To ubnd
        equity_hwm_arr(i, 1) = equity_hwm_arr(i - 1, 1) * (1 + trades_set(4, i))
        equity_hwm_arr(i, 2) = WorksheetFunction.Max(equity_hwm_arr(i, 1), equity_hwm_arr(i - 1, 2))
        dd_arr(i) = (equity_hwm_arr(i, 2) - equity_hwm_arr(i, 1)) / equity_hwm_arr(i, 2)
    Next i
    ' 1. sheet name
    perf_profile_arr(1) = ws_name
    ' 2. tpm
    cd_days = date_1 - date_0 + 1
    perf_profile_arr(2) = UBound(trades_set, 2) * 30.417 / cd_days
    ' 3. ar
    net_return = equity_hwm_arr(ubnd, 1) - 1
    If net_return < -1 Then
        net_return = -1
    End If
    perf_profile_arr(3) = (1 + net_return) ^ (365 / cd_days) - 1
    ' 4. mdd
    perf_profile_arr(4) = WorksheetFunction.Max(dd_arr)
    ' 5. rf
    If perf_profile_arr(3) > 0 And perf_profile_arr(4) > 0 Then
        perf_profile_arr(5) = perf_profile_arr(3) / perf_profile_arr(4)
    ElseIf perf_profile_arr(3) = 0 Then
        perf_profile_arr(5) = 999
    Else
        perf_profile_arr(5) = -1
    End If
    ' 6. r2
    perf_profile_arr(6) = WFA_Calculate_Rsq(trades_set, date_0, date_1)
    ' 7. usable
    perf_profile_arr(7) = True
    WFA_Get_Pfl_From_TSet = perf_profile_arr
End Function
Private Function WFA_Calculate_Rsq(ByVal trades_arr As Variant, ByVal date_0 As Date, ByVal date_1 As Date) As Double
    Dim i As Integer
    Dim j As Long
    Dim calend_arr() As Double
    Dim equity_arr() As Double
    Dim calendar_days As Integer
    
    calendar_days = date_1 - date_0 + 2
    ReDim calend_arr(1 To calendar_days)
    ReDim equity_arr(1 To calendar_days)
    calend_arr(1) = date_0 - 1
    equity_arr(1) = 1
    j = 1
    For i = 2 To UBound(calend_arr)
        calend_arr(i) = calend_arr(i - 1) + 1
        equity_arr(i) = equity_arr(i - 1)
        If calend_arr(i) = Int(trades_arr(2, j)) Then
            Do While calend_arr(i) = Int(trades_arr(2, j)) ' And j <= UBound(trades_arr, 2)
                equity_arr(i) = equity_arr(i) * (1 + trades_arr(4, j))
                If j < UBound(trades_arr, 2) Then
                    j = j + 1
                ElseIf j = UBound(trades_arr, 2) Then
                    Exit Do
                End If
            Loop
        End If
    Next i
    WFA_Calculate_Rsq = WorksheetFunction.RSq(calend_arr, equity_arr)
End Function
Private Function WFA_Update_Raw_Pfls(ByVal pfls_arr As Variant, _
                             ByVal is_single_ts_pfl As Variant) As Variant
    Dim ubnd As Long
    Dim i As Integer
    
    ubnd = UBound(pfls_arr, 2) + 1
    ReDim Preserve pfls_arr(1 To 7, 0 To ubnd)
    For i = LBound(pfls_arr, 1) To UBound(pfls_arr, 1)
        pfls_arr(i, ubnd) = is_single_ts_pfl(i)
    Next i
    WFA_Update_Raw_Pfls = pfls_arr
End Function
Private Function WFA_Get_Winner_ShName(ByVal raw_pfls As Variant, _
                               ByVal this_param As Integer) As String
' raw_pfls(1 To 7, 0 To sheets_in_one_book)
' 1. sheet name ("005")
' 2. tpm - trades per month
' 3. ar - annualized return
' 4. mdd - maximum drawdown
' 5. rf - recovery factor
' 6. r2 - R-squared
' 7. usable = True or False
    Dim filtered_indexes() As Variant
    Dim winner_sheet_id As Long
    Dim i As Long
    Dim ubnd As Long

    ReDim filtered_indexes(0 To UBound(raw_pfls, 2))
    For i = 1 To UBound(filtered_indexes)
        filtered_indexes(i) = i
    Next i
' FILTER 1. RF
    filtered_indexes = WFA_Apply_Param_Filter(raw_pfls, filtered_indexes, _
                       param_combin(3, this_param), param_combin(4, this_param), 5)
' FILTER 2. TPM
    If UBound(filtered_indexes) > 0 Then
        filtered_indexes = WFA_Apply_Param_Filter(raw_pfls, filtered_indexes, _
                           param_combin(1, this_param), param_combin(2, this_param), 2)
' FILTER 3. R2
        If UBound(filtered_indexes) > 0 Then
            filtered_indexes = WFA_Apply_Param_Filter(raw_pfls, filtered_indexes, _
                               param_combin(5, this_param), param_combin(6, this_param), 6)
' 4. MAXIMIZE RECOVERY FACTOR
            If UBound(filtered_indexes) > 0 Then
                winner_sheet_id = WFA_Maximize_by_RF(filtered_indexes, raw_pfls)
                WFA_Get_Winner_ShName = raw_pfls(1, winner_sheet_id)     ' SHEET NAME
            Else
                WFA_Get_Winner_ShName = "0"
            End If
        Else
            WFA_Get_Winner_ShName = "0"
        End If
    Else
        WFA_Get_Winner_ShName = "0"
    End If
End Function
Private Function WFA_Apply_Param_Filter(ByVal raw_pfls As Variant, _
                            ByVal filtered_indexes As Variant, _
                            ByVal req_min As Double, _
                            ByVal req_max As Double, _
                            ByVal param_id As Integer) As Variant
' Function returns an UPDATED array of filtered indexes
    Dim i As Long, j As Long
    Dim return_arr() As Variant
    Dim ubnd As Long
    
    ReDim return_arr(0 To 0)
    For i = 1 To UBound(filtered_indexes)
        j = filtered_indexes(i)
        If raw_pfls(param_id, j) > req_min And raw_pfls(param_id, j) <= req_max Then
            ubnd = UBound(return_arr) + 1
            ReDim Preserve return_arr(0 To ubnd)
            return_arr(ubnd) = j
        End If
    Next i
    WFA_Apply_Param_Filter = return_arr
End Function
Private Function WFA_Maximize_by_RF(ByVal filtered_indexes As Variant, _
                            ByVal raw_pfls As Variant) As Long
' Function returns index of WINNING report for FWD test (in one file - one instrument)
' uses Bubble Sort
    Dim index_vs_rf() As Variant
    Dim i As Integer, j As Integer
    Dim tmp_1 As Variant, tmp_2 As Variant
    
    ReDim index_vs_rf(1 To UBound(filtered_indexes), 1 To 2)
    For i = 1 To UBound(index_vs_rf, 1)
        index_vs_rf(i, 1) = filtered_indexes(i)
        index_vs_rf(i, 2) = raw_pfls(5, filtered_indexes(i))
    Next i
    ' bubble sort
    If UBound(index_vs_rf, 1) > 1 Then
        For i = 1 To UBound(index_vs_rf, 1) - 1
            For j = i + 1 To UBound(index_vs_rf, 1)
                If index_vs_rf(i, 2) < index_vs_rf(j, 2) Then
                    tmp_1 = index_vs_rf(j, 1)
                    tmp_2 = index_vs_rf(j, 2)
                    index_vs_rf(j, 1) = index_vs_rf(i, 1)
                    index_vs_rf(j, 2) = index_vs_rf(i, 2)
                    index_vs_rf(i, 1) = tmp_1
                    index_vs_rf(i, 2) = tmp_2
                End If
            Next j
        Next i
    End If
    WFA_Maximize_by_RF = index_vs_rf(1, 1)
End Function
Private Function WFA_Add_To_Top_BT_FD_Sets(ByVal original_full_arr As Variant, _
                               ByVal single_arr As Variant) As Variant
' Function appends trade set to big backtest/forwardtest array
' original_full_arr (1 To 4, 0 to trades_count)
    Dim r As Integer, full_row As Long
    Dim orig_ubnd As Long
    Dim c As Integer
    
    orig_ubnd = UBound(original_full_arr, 2)
    ReDim Preserve original_full_arr(1 To 4, 0 To orig_ubnd + UBound(single_arr, 2))
    For r = 1 To UBound(single_arr, 2)
        full_row = orig_ubnd + r
        For c = 1 To 4
        'For c = LBound(original_full_arr, 1) To UBound(original_full_arr, 1)
            original_full_arr(c, full_row) = single_arr(c, r)
        Next c
    Next r
    WFA_Add_To_Top_BT_FD_Sets = original_full_arr
End Function
Private Function WFA_Get_Tradeset_From_Sheet(ByVal ws As Worksheet, _
                                     ByVal date_0 As Date, _
                                     ByVal date_1 As Date, _
                                     ByVal book_name As String) As Variant
' Function returns 4-column set of trades.
    Dim result_arr() As Variant
    Dim last_row As Integer
    Dim comment_str As String
    Dim i As Integer
    Dim ubnd As Integer
    Dim wsC As Range
    
    comment_str = Left(book_name, Len(book_name) - 5) & "_" & ws.Name
    Set wsC = ws.Cells
    ReDim result_arr(1 To 4, 0 To 0)
    If wsC(11, 2) = 0 Then '1 JOIN
        WFA_Get_Tradeset_From_Sheet = result_arr
        Exit Function
    End If
    last_row = wsC(1, 9).End(xlDown).Row
    i = 2
    Do While Int(wsC(i, 9)) <= date_1 And i <= last_row
        If Int(wsC(i, 9)) >= date_0 And Int(wsC(i, 10)) <= date_1 Then
            ubnd = UBound(result_arr, 2) + 1
            ReDim Preserve result_arr(1 To 4, 0 To ubnd)
            result_arr(1, ubnd) = wsC(i, 9)
            result_arr(2, ubnd) = wsC(i, 10)
            result_arr(3, ubnd) = comment_str
            result_arr(4, ubnd) = wsC(i, 13)
        End If
        i = i + 1
    Loop
    WFA_Get_Tradeset_From_Sheet = result_arr
End Function
Private Function WFA_Sort_CloseDate(ByVal nd_arr As Variant) As Variant
' bubble sort trades set on "Close date" - 2nd col
    
    Dim i As Long, j As Long, k As Long
    Dim ubnd As Long
    Dim tmp_arr() As Variant
    
    ubnd = UBound(nd_arr, 1)
    ReDim tmp_arr(1 To ubnd)
    ' bubble sort
    If UBound(nd_arr, 2) > 1 Then   ' if more than 1 row
        For i = 1 To UBound(nd_arr, 2) - 1
            For j = i + 1 To UBound(nd_arr, 2)
                If nd_arr(2, i) > nd_arr(2, j) Then
                    For k = LBound(nd_arr, 1) To UBound(nd_arr, 1)
                        tmp_arr(k) = nd_arr(k, j)
                    Next k
                    For k = LBound(nd_arr, 1) To UBound(nd_arr, 1)
                        nd_arr(k, j) = nd_arr(k, i)
                    Next k
                    For k = LBound(nd_arr, 1) To UBound(nd_arr, 1)
                        nd_arr(k, i) = tmp_arr(k)
                    Next k
                End If
            Next j
        Next i
    End If
    WFA_Sort_CloseDate = nd_arr
End Function
Private Function WFA_Add_Header_CumSum(ByVal trade_set As Variant, ByVal add_cumsum As Boolean) As Variant
    Dim result_arr() As Variant
    Dim i As Long, j As Long
    
    ReDim result_arr(1 To 5, 0 To UBound(trade_set, 2))
    result_arr(1, 0) = "Open date"
    result_arr(2, 0) = "Close date"
    result_arr(3, 0) = "Comment"
    result_arr(4, 0) = "return"
    result_arr(5, 0) = 1
    If add_cumsum = False Then
        WFA_Add_Header_CumSum = result_arr
        Exit Function
    End If
    For i = 1 To UBound(trade_set, 2)   ' rows
        For j = LBound(trade_set, 1) To UBound(trade_set, 1)
            result_arr(j, i) = trade_set(j, i)
        Next j
        result_arr(5, i) = result_arr(5, i - 1) * (1 + trade_set(4, i))
    Next i
    WFA_Add_Header_CumSum = result_arr
End Function
Private Sub WFA_PrintOut_MegaArr()
' Subroutine prints out whole results:
' sheet(1) - summary
' sheets 2 to N+1 - single forward tests
    Const additional_sheets As Integer = 2
    Dim wb_po As Workbook
    Dim ws As Worksheet, ws_summary As Worksheet, ws_hps As Worksheet
    Dim print_cells As Range, summary_cells As Range, hps_cells As Range
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    Dim this_span As Integer, this_param As Integer, this_sheet As Integer, this_slot As Integer
    Dim sheets_required As Integer
    Dim col_offset As Integer
    Dim print_col As Integer, print_row As Integer
    Dim core_name As String, final_name As String
    Dim temp_s As String
    Dim vers As Integer
    Dim sum_ar As Double
    Dim ar_count As Integer, pos_ar_count As Integer

    Application.StatusBar = "Printing out"
    Set wb_po = Workbooks.Add
' add empty sheets
    sheets_required = UBound(spans_ar_rf_r2(1), 1) * UBound(param_combin, 2) + additional_sheets
    If wb_po.Sheets.Count < sheets_required Then
        For i = wb_po.Sheets.Count + 1 To sheets_required
            wb_po.Sheets.Add after:=wb_po.Sheets(wb_po.Sheets.Count)
        Next i
    ElseIf wb_po.Sheets.Count > sheets_required Then
        For i = sheets_required + 1 To wb_po.Sheets.Count
            Application.DisplayAlerts = False
            wb_po.Sheets(i).Delete
            Application.DisplayAlerts = True
        Next i
    End If
    For i = additional_sheets + 1 To Sheets.Count ' NAME SHEETS
        wb_po.Sheets(i).Name = i - additional_sheets
    Next i
    wb_po.Sheets(2).Activate
' hyperparameters - hps
    Set ws_hps = wb_po.Sheets(1)
    ws_hps.Name = "hps"
    Set hps_cells = ws_hps.Cells
    Call WFA_PrintOut_Hyperparameters(hps_cells)
' summary
    Set ws_summary = wb_po.Sheets(2)
    ws_summary.Name = "summary"
    Set summary_cells = ws_summary.Cells
    Call WFA_PrintOut_Summary(summary_cells)
' single sheets: 3 to end
    this_sheet = additional_sheets + 1   ' initialize sheet index
    For this_span = LBound(spans_ar_rf_r2(1), 1) To UBound(spans_ar_rf_r2(1), 1)    ' from 1 to spans_count
        For this_param = 1 To UBound(param_combin, 2)
            Set print_cells = wb_po.Sheets(this_sheet).Cells
            print_cells(1, 1) = "name " & wb_po.Sheets(this_sheet).Name   ' DUMMY
            ' print hyperparameters
            print_cells(2, 1) = "Parameters"
            For m = LBound(param_combin, 1) To UBound(param_combin, 1)
                print_cells(2 + m, 1) = param_combin(m, 0)
                print_cells(2 + m, 2) = param_combin(m, this_param)
            Next m
            print_row = 4 + UBound(param_combin, 1)
            
            print_cells(print_row, 1) = "IS wks"
            print_cells(print_row, 2) = spans_table(1, this_span)
            print_row = print_row + 1
            print_cells(print_row, 1) = "OS wks"
            print_cells(print_row, 2) = spans_table(2, this_span)
            
            ' print date slots
            print_row = print_row + 2
            print_cells(print_row, 1) = "Date from"
            print_cells(print_row, 2) = "Date to"
            print_cells(print_row, 3) = "type"
            print_cells(print_row, 4) = "TPM"
            print_cells(print_row, 5) = "AR"
            print_cells(print_row, 6) = "MDD"
            print_cells(print_row, 7) = "RF"
            print_cells(print_row, 8) = "R2"
            print_cells(print_row, 9) = "Usable"
            print_row = print_row + 1
            
            For this_slot = LBound(date_slots(this_span), 2) To UBound(date_slots(this_span), 2)
                print_cells(print_row, 1) = CDate(date_slots(this_span)(1, this_slot)) ' IS date from
                print_cells(print_row, 2) = CDate(date_slots(this_span)(2, this_slot)) ' IS date to
                print_cells(print_row, 3) = "is_single"
                col_offset = 3
                ' print stats for this slot: IS
                For k = 2 To UBound(mega_arr(this_span, this_param)(1)(this_slot)(2)(2))
                    print_col = col_offset + k - 1
                    With print_cells(print_row, print_col)
                        .Value = mega_arr(this_span, this_param)(1)(this_slot)(2)(2)(k)
                        .Interior.Color = RGB(197, 217, 241)
                        If k = 3 Or k = 4 Then
                            .NumberFormat = "0.0%"
                        Else
                            .NumberFormat = "0.00"
                        End If
                    End With
                Next k
                
                print_row = print_row + 1
                print_cells(print_row, 1) = CDate(date_slots(this_span)(3, this_slot)) ' OS date from
                print_cells(print_row, 2) = CDate(date_slots(this_span)(4, this_slot)) ' OS date to
                print_cells(print_row, 3) = "os_single"
                col_offset = 3
                ' print stats for this slot: OS
                For k = 2 To UBound(mega_arr(this_span, this_param)(2)(this_slot)(2))
                    print_col = col_offset + k - 1
                    With print_cells(print_row, print_col)
                        .Value = mega_arr(this_span, this_param)(2)(this_slot)(2)(k)
                        .Interior.Color = RGB(253, 233, 217)
                        If k = 3 Or k = 4 Then
                            .NumberFormat = "0.0%"
                        Else
                            .NumberFormat = "0.00"
                        End If
                    End With
                    ' highlight AR > 0
                    If k = 3 Then
                        If print_cells(print_row, print_col).Value > 0 Then
                            print_cells(print_row, print_col).Font.Color = RGB(0, 180, 80)
                        Else
                            print_cells(print_row, print_col).Font.Color = RGB(255, 0, 0)
                        End If
                    End If
                Next k
                print_row = print_row + 1
            Next this_slot
            print_cells(print_row, 1) = CDate(date_slots(this_span)(3, 1)) ' IS date from
            print_cells(print_row, 2) = CDate(date_slots(this_span)(4, UBound(date_slots(this_span), 2))) ' IS date to
            print_cells(print_row, 3) = "fwd_full"
            col_offset = 3
            ' print stats for this slot
            For k = 2 To UBound(mega_arr(this_span, this_param)(3)(2))
                print_col = col_offset + k - 1
                With print_cells(print_row, print_col)
                    .Value = mega_arr(this_span, this_param)(3)(2)(k)
                    .Interior.Color = RGB(146, 208, 80)
                    If k = 3 Or k = 4 Then
                        .NumberFormat = "0.0%"
                    Else
                        .NumberFormat = "0.00"
                    End If
                End With
            Next k

' COMPILED OS TRADESET
            print_cells(1, 11) = "Forward compiled"
            Call WFA_Print_2D_Array(mega_arr(this_span, this_param)(3)(1), True, 1, 10, print_cells)
            
' Print single IS & OS
            col_offset = 10 + UBound(mega_arr(this_span, this_param)(3)(1), 1) + 5
            For this_slot = LBound(mega_arr(this_span, this_param)(1)) To UBound(mega_arr(this_span, this_param)(1))
                ' print IS
                print_cells(1, col_offset + 1) = "IS slot " & this_slot
                Call WFA_Print_2D_Array(mega_arr(this_span, this_param)(1)(this_slot)(2)(1), True, 1, col_offset, print_cells)
                
                ' print OS
                col_offset = col_offset + UBound(mega_arr(this_span, this_param)(1)(this_slot)(2)(1), 1) + 5
                print_cells(1, col_offset + 1) = "OS slot " & this_slot
                Call WFA_Print_2D_Array(mega_arr(this_span, this_param)(2)(this_slot)(1), True, 1, col_offset, print_cells)
                
                col_offset = col_offset + UBound(mega_arr(this_span, this_param)(2)(this_slot)(1), 1) + 5
            Next this_slot

            this_sheet = this_sheet + 1
        Next this_param
    Next this_span
    
    highest_ar = -1000
    sum_ar = 0
    ar_count = 0
    pos_ar_count = 0
' Print out 3 metrics VS spans
    col_offset = 1 + UBound(param_combin, 1)
    For this_span = LBound(mega_arr, 1) To UBound(mega_arr, 1)
        print_col = col_offset + this_span
        For this_param = LBound(mega_arr, 2) To UBound(mega_arr, 2)
            ' update highest_ar, sum_ar
            If mega_arr(this_span, this_param)(3)(2)(3) > highest_ar Then
                highest_ar = mega_arr(this_span, this_param)(3)(2)(3)
            End If
            sum_ar = sum_ar + mega_arr(this_span, this_param)(3)(2)(3)
            ar_count = ar_count + 1
            If mega_arr(this_span, this_param)(3)(2)(3) > 0 Then
                pos_ar_count = pos_ar_count + 1
            End If
            '
            print_row = 2 + this_param
            With summary_cells(print_row, print_col)
                .Value = mega_arr(this_span, this_param)(3)(2)(3)  ' AR
                .NumberFormat = "0.0%"
            End With
            If summary_cells(print_row, print_col).Value > 0 Then
                summary_cells(print_row, print_col).Interior.Color = RGB(70, 255, 70)
                If summary_cells(print_row, print_col).Value > 0.3 Then
                    summary_cells(print_row, print_col).Font.Bold = True
                End If
            ElseIf summary_cells(print_row, print_col).Value < 0 Then
                summary_cells(print_row, print_col).Interior.Color = RGB(255, 70, 0)
            End If
            print_col = print_col + UBound(mega_arr, 1)
            With summary_cells(print_row, print_col)
                .Value = mega_arr(this_span, this_param)(3)(2)(5)  ' RF
                .NumberFormat = "0.00"
                If mega_arr(this_span, this_param)(3)(2)(3) > 0 Then
                    .Interior.Color = RGB(190, 255, 190)
                End If
            End With
            print_col = print_col + UBound(mega_arr, 1)
            With summary_cells(print_row, print_col)
                .Value = mega_arr(this_span, this_param)(3)(2)(6)  ' R2
                .NumberFormat = "0.00"
                If mega_arr(this_span, this_param)(3)(2)(3) > 0 Then
                    .Interior.Color = RGB(100, 255, 100)
                End If
            End With
            print_col = col_offset + this_span
        Next this_param
    Next this_span

' print highest_ar, avg_ar, pc_pos_os
    avg_ar = sum_ar / ar_count
    pc_pos_os = pos_ar_count / ar_count
    
' hps_cells
    hps_cells(11, 3) = "AR max"
    hps_cells(12, 3) = "AR avg"
    hps_cells(13, 3) = "AR positive, %"
    With hps_cells(11, 4)
        .Value = highest_ar
        .NumberFormat = "0.0%"
    End With
    With hps_cells(12, 4)
        .Value = avg_ar
        .NumberFormat = "0.0%"
    End With
    With hps_cells(13, 4)
        .Value = pc_pos_os
        .NumberFormat = "0.0%"
    End With

' save book
    core_name = WFA_Generate_BookName(hps_cells)
    final_name = core_name & ".xlsx"
    If Dir(final_name) = "" Then
        wb_po.SaveAs fileName:=final_name, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Else
        final_name = core_name & "(2).xlsx"
        If Dir(final_name) <> "" Then
            j = InStr(1, final_name, "(", 1)
            temp_s = Right(final_name, Len(final_name) - j)
            j = InStr(1, temp_s, ")", 1)
            vers = Left(temp_s, j - 1)
            final_name = core_name & "(" & vers & ").xlsx"
            Do Until Dir(final_name) = ""
                vers = vers + 1
                final_name = core_name & "(" & vers & ").xlsx"
            Loop
        End If
        wb_po.SaveAs fileName:=final_name, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    End If
    wb_po.Close
    Application.StatusBar = False
End Sub
Private Function WFA_Generate_BookName(ByVal hps_cells As Range) As String
    Dim dbeg As String, dfin As String
    Dim s_yr As String, s_mn As String, s_dy As String
    Dim lr As Integer
    Dim userTarget As String
    
    lr = hps_cells(1000, 1).End(xlUp).Row - 1
' date begin
    s_yr = Right(Year(hps_cells(2, 4)), 2)
    s_mn = Month(hps_cells(2, 4))
    If Len(s_mn) = 1 Then
        s_mn = "0" & s_mn
    End If
    s_dy = Day(hps_cells(2, 4))
    If Len(s_dy) = 1 Then
        s_dy = "0" & s_dy
    End If
    dbeg = s_yr & s_mn & s_dy
' date end
    s_yr = Right(Year(hps_cells(3, 4)), 2)
    s_mn = Month(hps_cells(3, 4))
    If Len(s_mn) = 1 Then
        s_mn = "0" & s_mn
    End If
    s_dy = Day(hps_cells(3, 4))
    If Len(s_dy) = 1 Then
        s_dy = "0" & s_dy
    End If
    dfin = s_yr & s_mn & s_dy
    userTarget = hps_cells(7, 4)
    If InStr(Len(userTarget) - 1, userTarget, "\", vbTextCompare) > 0 Then
        userTarget = Left(userTarget, Len(userTarget) - 1)
    End If
    WFA_Generate_BookName = userTarget & "\_WFA-" & hps_cells(5, 14) & "-" & lr & "-" & dbeg & "-" & dfin
End Function
Private Sub WFA_PrintOut_Hyperparameters(ByVal print_cells As Range)
    Dim wb_addin As Workbook
    Dim ws_settings As Worksheet ', ws_folders As Worksheet
    Dim settings_cells As Range ', folders_cells As Range
    Dim this_last_row As Integer
    Dim i As Integer, j As Integer
    Dim copy_rng As Range

    Set wb_addin = Workbooks(addInName)
    Set ws_settings = wb_addin.Sheets(wsSettingsName)
    Set settings_cells = ws_settings.Cells
' copy folders, settings, hyperparameters
    this_last_row = 4
    For i = 1 To 14
        j = settings_cells(ws_settings.rows.Count, i).End(xlUp).Row
        If j > this_last_row Then
            this_last_row = j
        End If
    Next i
    ' folders
    Set copy_rng = ws_settings.Range(settings_cells(1, 1), settings_cells(this_last_row, 1))
    copy_rng.Copy print_cells(1, 1)
    ' settings & hps
    Set copy_rng = ws_settings.Range(settings_cells(1, 3), settings_cells(this_last_row, 14))
    copy_rng.Copy print_cells(1, 3)
' print time
    print_cells(9, 3) = "Time started"
    print_cells(10, 3) = "Time ended"
    print_cells(9, 4) = time_started
    time_ended = Now
    print_cells(10, 4) = time_ended
' copy column widths
    For i = 1 To 14
        print_cells(5, i).columns.ColumnWidth = ws_settings.columns(i).ColumnWidth
    Next i
End Sub
Private Sub WFA_PrintOut_Summary(ByVal print_cells As Range)
    Dim i As Integer
    Dim print_col_offset As Integer
    Dim output_tables(1 To 3) As String

    output_tables(1) = "Annualized Return"
    output_tables(2) = "Recovery Factor"
    output_tables(3) = "R-Squared"
' indexes
    For i = 1 To UBound(param_combin, 2)
        print_cells(i + 2, 1) = i
'        ws.Hyperlinks.Add anchor:=print_cells(i + 2, 1), Address:="", SubAddress:="'" & i & "'!R1C1"
    Next i
    
    print_col_offset = 1
    Call WFA_Print_2D_Array(param_combin, True, 1, print_col_offset, print_cells)

    print_col_offset = print_col_offset + UBound(param_combin, 1)
    For i = 1 To 3
        print_cells(1, print_col_offset + 1) = output_tables(i)
        Call WFA_Print_2D_Array(spans_ar_rf_r2(i), True, 1, print_col_offset, print_cells)
        print_col_offset = print_col_offset + UBound(spans_ar_rf_r2(i), 1)
    Next i
End Sub
Private Function WFA_ListFiles(ByVal sPath As String)
    Dim vaArray As Variant
    Dim i As Integer
    Dim oFile As Object
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFiles As Object
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(sPath)
    Set oFiles = oFolder.files
    If oFiles.Count = 0 Then Exit Function
    ReDim vaArray(1 To oFiles.Count)
    i = 1
    For Each oFile In oFiles
        vaArray(i) = oFile.Name
        i = i + 1
    Next
    WFA_ListFiles = vaArray
End Function
Private Function WFA_Generate_Wbks_List(ByVal folders_arr As Variant) As Variant
    Dim arr() As Variant
    Dim this_folder As Integer
    Dim this_folder_filelist() As Variant
    Dim i As Integer, cur_pos As Integer
    Dim ubnd As Integer
    
    cur_pos = 1
    For this_folder = LBound(folders_arr) To UBound(folders_arr)
        this_folder_filelist = WFA_ListFiles(folders_arr(this_folder))
        If this_folder = 1 Then
            ubnd = UBound(this_folder_filelist)
        Else
            ubnd = UBound(arr) + UBound(this_folder_filelist)
        End If
        ReDim Preserve arr(1 To ubnd)
        For i = LBound(this_folder_filelist) To UBound(this_folder_filelist)
            arr(cur_pos) = folders_arr(this_folder) & "\" & this_folder_filelist(i)
            cur_pos = cur_pos + 1
        Next i
    Next this_folder
    WFA_Generate_Wbks_List = arr
End Function
Private Function WFA_Grab_Wbks_List() As Variant
    Dim arr() As Variant
    Dim cc As Range
    Dim wsf As Worksheet
    Dim last_row As Integer
    Dim i As Integer
    
    Set wsf = Workbooks(addInName).Sheets(wsSettingsName)
    Set cc = wsf.Cells
    last_row = cc(wsf.rows.Count, 1).End(xlUp).Row
    ReDim arr(1 To last_row - 1)
    For i = 2 To last_row
        arr(i - 1) = cc(i, 1)
    Next i
    WFA_Grab_Wbks_List = WFA_Generate_Wbks_List(arr)
End Function
Private Sub WFA_Print_2D_Array(ByVal print_arr As Variant, ByVal is_inverted As Boolean, _
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
Private Function WFA_Time_Consumed() As String
    Dim hrs As Long
    Dim mins As Long
    Dim secs As Long
    Dim st_hrs As String, st_mins As String, st_secs As String
    
    secs = Int((time_ended - time_started) * 86400)
    hrs = Int(secs / 3600)
    st_hrs = hrs
    If Len(st_hrs) = 1 Then
        st_hrs = 0 & st_hrs
    End If
    mins = Int((secs - hrs * 3600) / 60)
    st_mins = mins
    If Len(st_mins) = 1 Then
        st_mins = 0 & st_mins
    End If
    secs = Int(secs - hrs * 3600 - mins * 60)
    st_secs = secs
    If Len(st_secs) = 1 Then
        st_secs = 0 & st_secs
    End If
    WFA_Time_Consumed = st_hrs & ":" & st_mins & ":" & st_secs
End Function
Private Function Time_elapsed(ByVal timeBegin As Date, ByVal timeEnd As Date) As String
    Dim hrs As Long
    Dim mins As Long
    Dim secs As Long
    Dim st_hrs As String, st_mins As String, st_secs As String
    
    secs = Int((timeEnd - timeBegin) * 86400)
    hrs = Int(secs / 3600)
    st_hrs = hrs
    If Len(st_hrs) = 1 Then
        st_hrs = 0 & st_hrs
    End If
    mins = Int((secs - hrs * 3600) / 60)
    st_mins = mins
    If Len(st_mins) = 1 Then
        st_mins = 0 & st_mins
    End If
    secs = Int(secs - hrs * 3600 - mins * 60)
    st_secs = secs
    If Len(st_secs) = 1 Then
        st_secs = 0 & st_secs
    End If
    Time_elapsed = st_hrs & ":" & st_mins & ":" & st_secs
End Function
Private Sub WFA_Navigate_to_sheet()
    Dim paramCombCount As Integer
    Dim spansCount As Integer
    Dim arCol As Integer, rfCol As Integer, rsqCol As Integer
    Dim lastRow As Integer, lastCol As Integer
    Dim aRow As Integer, aCol As Integer
    Dim curSpan As Integer
    Dim shIndex As String
    Dim ws As Worksheet
    Dim c As Range
    
'    Application.ScreenUpdating = False
    If ActiveSheet.Name = "summary" Then
        Set ws = ActiveSheet
        Set c = ws.Cells
        lastRow = c(ws.rows.Count, 1).End(xlUp).Row
        paramCombCount = c(lastRow, 1).Value
        arCol = ws.Cells.Find(what:="Annualized Return", _
                after:=ws.Cells(1, 7), searchorder:=xlByRows).Column
        rfCol = ws.Cells.Find(what:="Recovery Factor", _
                after:=ws.Cells(1, 7), searchorder:=xlByRows).Column
        rsqCol = ws.Cells.Find(what:="R-Squared", _
                after:=ws.Cells(1, 7), searchorder:=xlByRows).Column
        spansCount = rfCol - arCol + 1
        lastCol = rsqCol + spansCount - 1
        aRow = ActiveCell.Row
        aCol = ActiveCell.Column
        If aRow > 2 And aRow <= lastRow And aCol >= arCol And aCol <= lastCol Then
            If aCol >= rsqCol Then
                curSpan = aCol - rsqCol + 1
            ElseIf aCol >= rfCol Then
                curSpan = aCol - rfCol + 1
            Else
                curSpan = aCol - arCol + 1
            End If
            shIndex = (curSpan - 1) * paramCombCount + aRow - 2
            Sheets(shIndex).Activate
        End If
    ElseIf Sheets(2).Name = "summary" Then
        Sheets("summary").Activate
    End If
'    Application.ScreenUpdating = True
End Sub
Private Sub Send_TG_Message(ByVal tgMsg As String, ByVal tgChatID As String)
' https://github.com/jenizar/Microsoft-Excel-Send-Message-to-Telegram/blob/master/README.md
'    Const strChatId As String = "your-chat-ID"
    Const bot_token As String = "your-bot-token"
    Dim objRequest As Object
    Dim strPostData As String
    Dim time_consumed As String
    Dim fsize As Double
    
    strPostData = "chat_id=" & tgChatID & "&text=" & tgMsg
    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    With objRequest
        .Open "POST", "https://api.telegram.org/bot" & bot_token & "/sendMessage?"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send (strPostData)
    End With
End Sub
Private Sub Merge_pick_source_folder()
' sheet "backtest"
' sub shows file dialog, lets user pick strategy folder
    Dim fd As FileDialog
    
    Application.ScreenUpdating = False
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "GetStats: Выбрать папку"
        .ButtonName = "Выбрать"
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    Workbooks(addInName).Sheets(wsMergeName).Cells(srcFdRow, srcFdCol) = fd.SelectedItems(1)
    Application.ScreenUpdating = True
End Sub
Private Sub Merge_summaries()
    Const zeroShCol As Integer = 7
'    Const everyNFiles As Integer = 5
    Dim pasteCol As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim pasteWSIndex As Integer
    Dim sheetsReq As Integer
    Dim nextFreeRow As Integer
    Dim lastHPRow As Integer
    Dim iterWSLRow As Long
    Dim iterWSLCol As Integer
    Dim m As Long
    Dim avgVal As Double
    Dim fileList() As Variant
    Dim targetWb As Workbook, sourceWb As Workbook
    Dim saveName As String
    Dim sourceWBPath As String
    Dim sourceFolder As String
    Dim iterCells As Range
    Dim Rng As Range
    Dim x As Integer
    Dim iterWS As Worksheet
'    Dim initSaved As Boolean
    
    Application.ScreenUpdating = False
    sourceFolder = Workbooks(addInName).Sheets(wsMergeName).Cells(srcFdRow, srcFdCol)
    fileList = WFA_ListFiles(sourceFolder)
    sheetsReq = 2
    Set targetWb = Workbooks.Add
    Call Change_sheets_count_names(targetWb, sheetsReq)
    x = UBound(fileList)
    nextFreeRow = 5
'    initSaved = False
    For i = LBound(fileList) To UBound(fileList)
        Application.StatusBar = "File " & i & " (" & x & ")."
        sourceWBPath = sourceFolder & "\" & fileList(i)
        Set sourceWb = Workbooks.Open(sourceWBPath)
        If i = 1 Then
            ' copy info from summary sheet
            Set iterCells = sourceWb.Sheets(1).Cells
            Set Rng = Range(iterCells(1, 1), iterCells(7, 4))
            Rng.Copy targetWb.Sheets(1).Cells(1, 1)
            Set Rng = Range(iterCells(1, 6), iterCells(9, 14))
            Rng.Copy targetWb.Sheets(1).Cells(1, 6)
            Set iterCells = sourceWb.Sheets(2).Cells
            lastHPRow = iterCells(sourceWb.Sheets(2).rows.Count, 1).End(xlUp).Row
            Set Rng = Range(iterCells(1, 1), iterCells(lastHPRow, 7))
            Rng.Copy targetWb.Sheets(2).Cells(1, 1)
            ' headings
            targetWb.Sheets(2).Cells(1, 8) = "Annualized Return"
            targetWb.Sheets(2).Cells(1, zeroShCol + x + 1) = "Recovery Factor"
            targetWb.Sheets(2).Cells(1, zeroShCol + 2 * x + 1) = "R-Squared"
        End If
        ' copy hps
        Set iterCells = sourceWb.Sheets(1).Cells
        Set Rng = Range(iterCells(5, 12), iterCells(5, 14))
        Rng.Copy targetWb.Sheets(1).Cells(nextFreeRow, 12)
        nextFreeRow = nextFreeRow + 1
        
        Set iterCells = sourceWb.Sheets(2).Cells
        ' AR
        Set Rng = Range(iterCells(2, 8), iterCells(lastHPRow, 8))
        pasteCol = zeroShCol + i
        Rng.Copy targetWb.Sheets(2).Cells(2, pasteCol)
            
        avgVal = WorksheetFunction.Average(Rng)
        ' highlight max
        Set Rng = targetWb.Sheets(2).Range(targetWb.Sheets(2).Cells(2, pasteCol), targetWb.Sheets(2).Cells(lastHPRow, pasteCol))
        Call HighlightMaxValueInRng(Rng)
        Set Rng = targetWb.Sheets(2).Cells(lastHPRow + 1, pasteCol)
        With Rng
            .Value = avgVal ' calc Avg AR
            .NumberFormat = "0.0%"
        End With
        If Rng.Value > 0 Then
            Rng.Interior.Color = RGB(150, 255, 150)
        Else
            Rng.Interior.Color = RGB(255, 150, 150)
        End If
            
            ' hyperlink
        targetWb.Sheets(2).Cells(lastHPRow + 2, pasteCol) = "open"
        targetWb.Sheets(2).Hyperlinks.Add Anchor:=targetWb.Sheets(2).Cells(lastHPRow + 2, pasteCol), _
            Address:=sourceWBPath
        
        ' RF
        Set Rng = Range(iterCells(2, 9), iterCells(lastHPRow, 9))
        pasteCol = zeroShCol + x + i
        Rng.Copy targetWb.Sheets(2).Cells(2, pasteCol)
            ' hyperlink
        targetWb.Sheets(2).Cells(lastHPRow + 2, pasteCol) = "open"
        targetWb.Sheets(2).Hyperlinks.Add Anchor:=targetWb.Sheets(2).Cells(lastHPRow + 2, pasteCol), _
            Address:=sourceWBPath
        ' RSQ
        Set Rng = Range(iterCells(2, 10), iterCells(lastHPRow, 10))
        pasteCol = zeroShCol + 2 * x + i
        Rng.Copy targetWb.Sheets(2).Cells(2, pasteCol)
            ' hyperlink
        targetWb.Sheets(2).Cells(lastHPRow + 2, pasteCol) = "open"
        targetWb.Sheets(2).Hyperlinks.Add Anchor:=targetWb.Sheets(2).Cells(lastHPRow + 2, pasteCol), _
            Address:=sourceWBPath
        sourceWb.Close savechanges:=False
    Next i
    ' save & close
    Application.StatusBar = "Saving target workbook."
    saveName = Generate_save_name(sourceFolder, fileList(1), x)
    targetWb.SaveAs fileName:=saveName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    targetWb.Close
    '
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Beep
End Sub
Private Sub Merge_main()
    Const zeroShCol As Integer = 7
    Dim pasteCol As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim pasteWSIndex As Integer
    Dim sheetsReq As Integer
    Dim nextFreeRow As Integer
    Dim iterWSLRow As Long
    Dim iterWSLCol As Integer
    Dim hyperParamsCount As Integer
    Dim m As Long
    Dim fileList() As Variant
    Dim targetWb As Workbook, sourceWb As Workbook
    Dim saveName As String
    Dim sourceWBPath As String
    Dim sourceFolder As String
    Dim iterCells As Range
    Dim Rng As Range
    Dim x As Integer
    Dim iterWS As Worksheet
    Dim avgVal As Double
    
    Application.ScreenUpdating = False
    sourceFolder = Workbooks(addInName).Sheets(wsMergeName).Cells(srcFdRow, srcFdCol)
    fileList = WFA_ListFiles(sourceFolder)
    x = UBound(fileList)
    nextFreeRow = 5
    For i = LBound(fileList) To UBound(fileList)
        Application.StatusBar = "File " & i & " (" & x & ")."
        sourceWBPath = sourceFolder & "\" & fileList(i)
        Set sourceWb = Workbooks.Open(sourceWBPath)
        If i = 1 Then
            Set iterCells = sourceWb.Sheets(2).Cells
            hyperParamsCount = iterCells(iterCells(3, 1).End(xlDown).Row, 1).Value
            sheetsReq = UBound(fileList) * hyperParamsCount + 2
            Set targetWb = Workbooks.Add
            Call Change_sheets_count_names(targetWb, sheetsReq)
            Set iterCells = sourceWb.Sheets(1).Cells
            Set Rng = Range(iterCells(1, 1), iterCells(7, 4))
            Rng.Copy targetWb.Sheets(1).Cells(1, 1)
            Set Rng = Range(iterCells(1, 6), iterCells(9, 14))
            Rng.Copy targetWb.Sheets(1).Cells(1, 6)
            Set iterCells = sourceWb.Sheets(2).Cells
            Set Rng = Range(iterCells(1, 1), iterCells(hyperParamsCount + 2, 7))
            Rng.Copy targetWb.Sheets(2).Cells(1, 1)
            ' headings
            targetWb.Sheets(2).Cells(1, 8) = "Annualized Return"
            targetWb.Sheets(2).Cells(1, zeroShCol + x + 1) = "Recovery Factor"
            targetWb.Sheets(2).Cells(1, zeroShCol + 2 * x + 1) = "R-Squared"
        End If
        ' copy hps
        Set iterCells = sourceWb.Sheets(1).Cells
        Set Rng = Range(iterCells(5, 12), iterCells(5, 14))
        Rng.Copy targetWb.Sheets(1).Cells(nextFreeRow, 12)
        nextFreeRow = nextFreeRow + 1
        ' AR
        Set iterCells = sourceWb.Sheets(2).Cells
        Set Rng = Range(iterCells(2, 8), iterCells(52, 8))
        pasteCol = zeroShCol + i
        Rng.Copy targetWb.Sheets(2).Cells(2, pasteCol)
        avgVal = WorksheetFunction.Average(Rng)
        
        ' highlight max
        Set Rng = targetWb.Sheets(2).Range(targetWb.Sheets(2).Cells(2, pasteCol), targetWb.Sheets(2).Cells(hyperParamsCount + 2, pasteCol))
        Call HighlightMaxValueInRng(Rng)
        Set Rng = targetWb.Sheets(2).Cells(hyperParamsCount + 3, pasteCol)
        With Rng
            .Value = avgVal ' calc Avg AR
            .NumberFormat = "0.0%"
        End With
        If Rng.Value > 0 Then
            Rng.Interior.Color = RGB(150, 255, 150)
        Else
            Rng.Interior.Color = RGB(255, 150, 150)
        End If
        
        ' RF
        Set Rng = Range(iterCells(2, 9), iterCells(hyperParamsCount + 2, 9))
        pasteCol = zeroShCol + x + i
        Rng.Copy targetWb.Sheets(2).Cells(2, pasteCol)
        ' RSQ
        Set Rng = Range(iterCells(2, 10), iterCells(hyperParamsCount + 2, 10))
        pasteCol = zeroShCol + 2 * x + i
        Rng.Copy targetWb.Sheets(2).Cells(2, pasteCol)
        ' copy sheets
        For j = 3 To sourceWb.Sheets.Count
            Set iterWS = sourceWb.Sheets(j)
            Set iterCells = iterWS.Cells
            iterWSLCol = iterCells(2, iterWS.columns.Count).End(xlToLeft).Column
            ' find last row on iter sheet
            iterWSLRow = 1
            For k = 1 To iterWSLCol
                m = iterCells(iterWS.rows.Count, k).End(xlUp).Row
                If m > iterWSLRow Then
                    iterWSLRow = m
                End If
            Next k
            ' copy range
            Set Rng = Range(iterCells(1, 1), iterCells(iterWSLRow, iterWSLCol))
            pasteWSIndex = j + hyperParamsCount * (i - 1)
            Rng.Copy targetWb.Sheets(pasteWSIndex).Cells(1, 1)
        Next j
        sourceWb.Close savechanges:=False
    Next i
    targetWb.Sheets(2).Cells(3, 8).Activate
    ActiveWindow.FreezePanes = True
    ' save & close
    Application.StatusBar = "Saving target workbook."
    saveName = Generate_save_name(sourceFolder, fileList(1), UBound(fileList))
    targetWb.SaveAs fileName:=saveName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    targetWb.Close
    '
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Beep
End Sub
Private Sub HighlightMaxValueInRng(ByRef someRng As Range)
    Dim cell As Range
    Dim maxVal As Double
    
    maxVal = WorksheetFunction.Max(someRng)
    For Each cell In someRng
        If cell.Value = maxVal Then
            With cell.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = 5
            End With
            With cell.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = 5
            End With
            With cell.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = 5
            End With
            With cell.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = 5
            End With
            Exit For
        End If
    Next cell
End Sub
Private Function Generate_save_name(ByVal targetFolder As String, _
                                    ByVal oneFileName As String, _
                                    ByVal filesCount As Integer) As String
' function generates save file name
' checks if exists
' adds index in parentheses, if exists
    Dim j As Integer
    Dim vers As Integer
    Dim coreName As String
    Dim temp_s As String
    Dim stratName As String
    Dim finalName As String
    Dim dateFrom As String, dateTo As String
    Dim newTargetFolder As String
    
    j = InStrRev(targetFolder, "\", , vbTextCompare)
    stratName = Right(targetFolder, Len(targetFolder) - j)
    newTargetFolder = Left(targetFolder, j)
    dateFrom = Left(Right(oneFileName, 18), 6)
    dateTo = Left(Right(oneFileName, 11), 6)
'    coreName = newTargetFolder & "WFA-hps" & filesCount & "-" & stratName & "-" & dateFrom & "-" & dateTo
    coreName = newTargetFolder & "WFA-" & stratName & "-hps" & filesCount & "-" & dateFrom & "-" & dateTo
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
    Generate_save_name = finalName
End Function
Private Sub Change_sheets_count_names(ByRef someWB As Workbook, ByVal shCount As Integer)
' function returns a new workbook with specified number of sheets
    Const shNameOne As String = "hps"
    Const shNameTwo As String = "summary"
    Dim i As Integer
    
    If someWB.Sheets.Count > shCount Then
        Application.DisplayAlerts = False
        For i = 1 To someWB.Sheets.Count - shCount
            someWB.Sheets(someWB.Sheets.Count).Delete
        Next i
        Application.DisplayAlerts = True
    ElseIf someWB.Sheets.Count < shCount Then
        For i = 1 To shCount - someWB.Sheets.Count
            someWB.Sheets.Add after:=someWB.Sheets(someWB.Sheets.Count)
        Next i
    End If
    someWB.Sheets(1).Name = shNameOne
    someWB.Sheets(2).Name = shNameTwo
' rename rest of sheets
    For i = 3 To someWB.Sheets.Count
        someWB.Sheets(i).Name = i - 2
    Next i
    someWB.Sheets(2).Activate
End Sub
