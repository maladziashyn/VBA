

' MODULE: ThisWorkbook
Option Explicit
Private Sub Workbook_Open()
    Call CreateCommandBars
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call RemoveCommandBars
End Sub

' MODULE: WFA_OLD
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

' MODULE: WFA_Tools_old
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


' MODULE: Command_Bars
Option Explicit

    Const cBarsCount As Integer = 3

Sub RemoveCommandBars()
    
    On Error Resume Next
    Dim i As Integer
    Dim addInName As String
    
    Call CommandBars_Inits(addInName)
    
    For i = 1 To cBarsCount
        Application.CommandBars(CommandBarName(addInName, i)).Delete
    Next i

End Sub

Sub CreateCommandBars()
    
    Dim cBar1 As CommandBar
    Dim cBar2 As CommandBar
    Dim cBar3 As CommandBar
    Dim cControl As CommandBarControl
    Dim addInName As String
    
    Call CommandBars_Inits(addInName)
    Call RemoveCommandBars
' Create toolbar 1
    Set cBar1 = Application.CommandBars.Add
    cBar1.Name = CommandBarName(addInName, 1)
    cBar1.Visible = True
' Create toolbar 2
    Set cBar2 = Application.CommandBars.Add
    cBar2.Name = CommandBarName(addInName, 2)
    cBar2.Visible = True
' Create toolbar 3
    Set cBar3 = Application.CommandBars.Add
    cBar3.Name = CommandBarName(addInName, 3)
    cBar3.Visible = True

' Row 1
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 424
        .OnAction = "ChartForTradeList"
        .TooltipText = "Chart for Trade List"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Chart"
    End With

    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 435
        .OnAction = "WfaPreviews"
        .TooltipText = "Make Previews"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Previews"
    End With
    
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 458
        .OnAction = "WfaWinnersRemoveDuplicates"
        .TooltipText = "Select Winners from IS/OS"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "WfaSlotFilter"
    End With

' Row 2
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 2937
        .OnAction = "OpenWfaSource"
        .TooltipText = "Open WFA Source Sheet"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "OpenSrc"
    End With
    
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 1248
        .OnAction = "ManuallyApplyDateFilter"
        .TooltipText = "Date Filter, KPIs"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "DtFilterKPIs"
    End With

    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 435
        .OnAction = "WfaDateSlotPreviews"
        .TooltipText = "Date slot previews"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "DtSlotPreviews"
    End With

' Row 3
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 601
        .OnAction = "DescriptionFilterChart"
        .TooltipText = "Statement filter and chart"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "StatementChart"
    End With

    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 620
        .OnAction = "SortSheetsAlphabetically"
        .TooltipText = "Sort Sheets Alphabetically"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "SortSheetsAsc"
    End With

End Sub

' MODULE: Inits
Option Explicit

    Const addinFileName As String = "JFTools"
    Const addinVersion As String = "0.22"
    
    Const btGroupShNm As String = "Back-test"
    Dim btGroupWs As Worksheet
    Dim btGroupC As Range
    
' **** WFA Main *** '
    Const wfaMainShNm As String = "WFA"
    Const testDate0Row As Integer = 2
    Const testDate9Row As Integer = 3
    Const targetMDDRow As Integer = 4
    Const mddFreedomRow As Integer = 5
    Const targetDirRow As Integer = 7
    Const sourceZeroRow As Integer = 8
    
    Const stgKeyCol As Integer = 29
    Const stgValueCol As Integer = 30
'    Const wfaMergeShNm As String = "WFA Merge"
'    Dim wfaMergeWs As Worksheet
'    Dim wfaMergeC As Range
        
' **** Hidden Settings *** '
    Const hiddenSetShNm As String = "Hidden Settings"
    Const scanModeRow As Integer = 18
    Const scanModeCol As Integer = 8
    
    Const activeKPIsFRow As Integer = 3
    Const activeKPIsFCol As Integer = 7

' **** Sorting *** '
    Const sortShNm As String = "Sorting"
    Const sortColID As Integer = 2

' **** other *** '
    Const dialTitleTargetDir As String = "Locate Target Directory"
    Const dialTitleSourceDir As String = "Locate Source Directory"
    Const dialPickParentDir As String = "Locate Parent Directory"
    Const okButtonName As String = "Okey Dokey"
    
    Const scanTableRowOffset As Integer = 3
    Const scanTableColOffset As Integer = 33
    
    Const windowsFirstRow As Integer = 1
    Const windowsFirstCol As Integer = 25
    
    Const permutZeroRow As Integer = 14
    Const permutFirstCol As Integer = 3
    
' **** WFA Chart **** '
    Const wfaIsLogScale As Boolean = True

' *** Back-test *** '
    Const stratFdRow As Integer = 2 ' parent directory row
    Const stratFdCol As Integer = 2 ' parent directory column
    Const stratNmRow As Integer = 3 ' strategy name row
    Const stratNmCol As Integer = 2 ' strategy name column

' **** STATEMENT *** '
    Const statementShNm As String = "Statement"
    Const fundsHistoryRow As Integer = 2
    Const portfolioSummaryRow As Integer = 3
    Const positionsCloseRow As Integer = 4
    Const targetDirectoryRow As Integer = 7
    Const srcInsertCol As Integer = 3
    Const dialTitleFundsHistory As String = "Pick Funds History directory"
    Const dialTitlePortfolioSummary As String = "Pick Portfolio Summary directory"
    Const dialTitlePositionsClose As String = "Pick Positions Close directory"
    Const dialTitleTargetDirectory As String = "Pick Statement Target directory"
    Const dialTitleStatementRoot As String = "Pick Root Statement directory"

Sub CommandBars_Inits(ByRef addInName As String)
    
    addInName = addinFileName

End Sub

Sub InitWfaChart(ByRef isChartLogScale As Boolean)

    isChartLogScale = wfaIsLogScale
    
End Sub

Sub Init_Parameters(ByRef param As Dictionary, _
            ByRef exitOnError As Boolean, _
            ByRef errorMsg As String, _
            ByRef sortWs As Worksheet, _
            ByRef sortC As Range, _
            ByRef sortBubbleColID As Integer, _
            ByRef kpiFormatting As Dictionary)
' Create WFA parameters dictionary.

    Dim kpisAndPermutations As Variant
    Dim mainWs As Worksheet
    Dim mainC As Range
    Dim stgWs As Worksheet
    Dim stgC As Range
    Dim addinWbNm As String
    Dim lastSrcDirRow As Integer
    
    addinWbNm = AddInFullFileName(addinFileName, addinVersion)
    Set mainWs = Workbooks(addinWbNm).Sheets(wfaMainShNm)
    Set mainC = mainWs.Cells
    Set stgWs = Workbooks(addinWbNm).Sheets(hiddenSetShNm)
    Set stgC = stgWs.Cells
    Set sortWs = Workbooks(addinWbNm).Sheets(sortShNm)
    Set sortC = sortWs.Cells
    sortBubbleColID = sortColID
    ' "KPI formatting"
    Set kpiFormatting = GetKPIFormatting   ' as Dictionary
    
    Set param = New Dictionary
    
' "Scan mode"
' Integer
    param.Add "Scan mode", stgC(scanModeRow, scanModeCol)
    
' "Date start", "Date end"
' Date
    param.Add "Date start", mainC(testDate0Row, stgValueCol)
    param.Add "Date end", mainC(testDate9Row, stgValueCol)
    
' "Target MDD"
' Double
    param.Add "Target MDD", mainC(targetMDDRow, stgValueCol)
    
' "MDD freedom"
' Double
    param.Add "MDD freedom", mainC(mddFreedomRow, stgValueCol)
    
' "Target directory"
' String
    param.Add "Target directory", StringRemoveBackslash(mainC(targetDirRow, stgKeyCol))
    
' "Source directories"
' Variant, 1D array, 1-based
    param.Add "Source directories", GetSourceDirectories(mainWs, mainC, sourceZeroRow, stgKeyCol)
    
' "Scan table"
' Variant, 2D array
' Not Inverted
' ROWS: 0-based. Header = strategies, rows - currencies
' COLUMNS: 0-based. Index column = currencies, columns - strategies
    param.Add "Scan table", GetScanTable(param("Source directories"), mainWs, mainC, scanTableRowOffset, scanTableColOffset)
    
' "IS/OS windows"
' Variant
' from range on mainC
' 1-based 3-column array of IS and OS weeks with their codes
' NOT INVERTED
    param.Add "IS/OS windows", GetIsOsWindows(mainWs, mainC, windowsFirstRow, windowsFirstCol)

' "MaxiMinimize"
' Variant
' 1D array (1 to 2)
' arr(1) = "maximize" or "minimize" or "none"
' arr(2) = "Sharpe Ratio" or "none"
    param.Add "MaxiMinimize", GetMaxiMinimize(stgWs, stgC, _
            activeKPIsFRow, activeKPIsFCol)

' KPI ranges
' Variant
' 2D array
' not inverted
' ROWS: 0-based, 0 is KPI names (skip 1 col), 1 is header "min/max"
'    param.Add "KPI ranges", GetKpiRanges()
    kpisAndPermutations = GetPermutations(mainWs, mainC, _
            permutZeroRow, permutFirstCol, stgWs, stgC, _
            activeKPIsFRow, activeKPIsFCol, param("MaxiMinimize"))
    param.Add "KPI ranges", kpisAndPermutations(1)

' "Permutations"
' Variant
' 2D array
' not inverted
' ROWS: 0-based, header-1 - KPIs, header-2 - "min/max"
' COLUMNS: 0-based, index column - index of KPI starting with "1"
    param.Add "Permutations", kpisAndPermutations(2)
'    param.Add "Permutations", GetPermutations(mainWs, mainC, _
'            permutZeroRow, permutFirstCol, stgWs, stgC, _
'            activeKPIsFRow, activeKPIsFCol)



'    param.Add "Scan sequences", GetScanSequences(param("Source directories"), _
'                param("Scan table"), _
'                param("Scan mode"), _
'                param("Target directory"))
    
' param("Scan table")

End Sub

Sub Click_Copy_Dates_From_Selection_Inits(ByRef mainWs As Worksheet, _
            ByRef mainC As Range, _
            ByRef date0Row As Integer, _
            ByRef date9Row As Integer, _
            ByRef datesCol As Integer, _
            ByRef dirsCol As Integer, _
            ByRef srcDirsZeroRow As Integer)
' sanity check
' not empty, row after "Source directories", right column

    Call Main_Sheet_Cells_Inits(mainWs, mainC)
    date0Row = testDate0Row
    date9Row = testDate9Row
    datesCol = stgValueCol
    dirsCol = stgKeyCol
    srcDirsZeroRow = sourceZeroRow
    
End Sub
Sub Main_Sheet_Cells_Inits(ByRef mainWs As Worksheet, _
            ByRef mainC As Range)
            
    Dim addinWbNm As String
    addinWbNm = AddInFullFileName(addinFileName, addinVersion)
    Set mainWs = Workbooks(addinWbNm).Sheets(wfaMainShNm)
    Set mainC = mainWs.Cells
    
End Sub
Sub Click_Locate_Target_Inits(ByRef ws As Worksheet, _
            ByRef fillRng As Range, _
            ByRef fillRow As Integer, _
            ByRef fillCol As Integer, _
            ByRef dialTitle As String, _
            ByRef okBtnName As String, _
            ByRef addSource As Boolean)
' Sub initiates worksheet, filedialog variables
' for "Locate Target Directory" button
    Call Main_Sheet_Cells_Inits(ws, fillRng)
    fillCol = stgKeyCol
    If addSource Then
        fillRow = fillRng(ws.rows.Count, fillCol).End(xlUp).Row + 1
        dialTitle = dialTitleSourceDir
    Else
        fillRow = targetDirRow
        dialTitle = dialTitleTargetDir
    End If
    okBtnName = okButtonName
End Sub
Sub Click_Clear_Sources_Inits(ByRef ws As Worksheet, _
            ByRef clrRng As Range)
    Dim c As Range
    Dim lastRow As Integer
    Set ws = Workbooks(AddInFullFileName(addinFileName, addinVersion)).Sheets(wfaMainShNm)
    Set c = ws.Cells
    lastRow = c(ws.rows.Count, stgKeyCol).End(xlUp).Row
    If lastRow = sourceZeroRow Then
        lastRow = lastRow + 1
    End If
    Set clrRng = ws.Range(c(sourceZeroRow + 1, stgKeyCol), c(lastRow, stgKeyCol))
End Sub
Sub DeSelect_KPIs_Inits(ByRef iHiddenSetWs As Worksheet, _
            ByRef iHiddenSetC As Range, _
            ByRef iCheckAll As Range, _
            ByRef iKpisList As Range)
    Dim addinWbNm As String
    
    addinWbNm = AddInFullFileName(addinFileName, addinVersion)
    Set iHiddenSetWs = Workbooks(addinWbNm).Sheets(hiddenSetShNm)
    Set iHiddenSetC = iHiddenSetWs.Cells
    Set iCheckAll = iHiddenSetC(2, 8)
    Set iKpisList = iHiddenSetWs.Range(iHiddenSetC(3, 8), iHiddenSetC(12, 8))
End Sub
Sub GenerateIsOsCodes_Inits(ByRef rg As Range, _
            ByRef exitError As Boolean, _
            ByRef addInName As String)
    Dim mainWs As Worksheet
    Dim mainC As Range
    Call Main_Sheet_Cells_Inits(mainWs, mainC)
    addInName = addinFileName
    Set rg = mainC(windowsFirstRow, windowsFirstCol).CurrentRegion
    If rg.rows.Count = 2 Then
        exitError = True
        Exit Sub
    End If
    Set rg = rg.Offset(2, 0).Resize(rg.rows.Count - 2)
End Sub
Sub MergeSummaries_Inits(ByRef initDirPath As String, _
            ByRef okBtnName As String)
    Dim addinWbNm As String
    Dim mainC As Range
    addinWbNm = AddInFullFileName(addinFileName, addinVersion)
    Set mainC = Workbooks(addinWbNm).Sheets(wfaMainShNm).Cells
    initDirPath = StringRemoveBackslash(mainC(targetDirRow, stgKeyCol).Value)
    okBtnName = okButtonName
End Sub
Sub DeSelect_Instruments_Inits(ByRef setWs As Worksheet, _
            ByRef btWs As Worksheet, _
            ByRef setC As Range, _
            ByRef btC As Range, _
            ByRef selectAll As Range, _
            ByRef instrumentsList As Range)
    Dim addinWbNm As String
    addinWbNm = AddInFullFileName(addinFileName, addinVersion)
    Set setWs = Workbooks(addinWbNm).Sheets(hiddenSetShNm)
    Set setC = setWs.Cells
    Set btWs = Workbooks(addinWbNm).Sheets(btGroupShNm)
    Set btC = btWs.Cells
    Set selectAll = setC(1, 2)
    Set instrumentsList = setWs.Range(setC(2, 2), setC(31, 2))
End Sub
Sub LocateParentDirectory_Inits(ByRef parentDirRg As Range, _
            ByRef stratNmRg As Range, _
            ByRef fdTitle As String, _
            ByRef fdButton As String)
    Dim addinWbNm As String
    addinWbNm = AddInFullFileName(addinFileName, addinVersion)
    Set parentDirRg = Workbooks(addinWbNm).Sheets(btGroupShNm).Cells(stratFdRow, stratFdCol)
    Set stratNmRg = Workbooks(addinWbNm).Sheets(btGroupShNm).Cells(stratNmRow, stratNmCol)
    fdTitle = dialPickParentDir
    fdButton = okButtonName
End Sub
Sub StatementClickLocate_Inits(ByRef mainC As Range, _
            ByRef insertRow As Integer, _
            ByRef insertCol As Integer, _
            ByVal sourceType As String, _
            ByRef dialogTitle As String, _
            ByRef okBtnName As String)
    Dim addinWbNm As String
    addinWbNm = AddInFullFileName(addinFileName, addinVersion)
    Set mainC = Workbooks(addinWbNm).Sheets(statementShNm).Cells
    Select Case sourceType
        Case Is = "Funds History"
            insertRow = fundsHistoryRow
            dialogTitle = dialTitleFundsHistory
        Case Is = "Portfolio Summary"
            insertRow = portfolioSummaryRow
            dialogTitle = dialTitlePortfolioSummary
        Case Is = "Positions Close"
            insertRow = positionsCloseRow
            dialogTitle = dialTitlePositionsClose
        Case Is = "Target Directory"
            insertRow = targetDirectoryRow
            dialogTitle = dialTitleTargetDirectory
        Case Is = "Root"
            insertRow = 0
            dialogTitle = dialTitleStatementRoot
    End Select
    insertCol = srcInsertCol
    okBtnName = okButtonName
End Sub
Sub Statement_Init_Parameters(ByRef param As Dictionary)
    Dim addinWbNm As String
    Dim mainC As Range
    addinWbNm = AddInFullFileName(addinFileName, addinVersion)
    Set mainC = Workbooks(addinWbNm).Sheets(statementShNm).Cells
    Set param = New Dictionary
    param.Add "Funds History", mainC(fundsHistoryRow, srcInsertCol)
    param.Add "Portfolio Summary", mainC(portfolioSummaryRow, srcInsertCol)
    param.Add "Positions Close", mainC(positionsCloseRow, srcInsertCol)
    param.Add "Target Directory", mainC(targetDirectoryRow, srcInsertCol)
End Sub
Sub ClickLocateRoot_InitsPart2(ByRef fundsRow As Integer, _
            ByRef portfolioRow As Integer, _
            ByRef positionsRow As Integer, _
            ByRef targetDirRow As Integer)
    fundsRow = fundsHistoryRow
    portfolioRow = portfolioSummaryRow
    positionsRow = positionsCloseRow
    targetDirRow = targetDirectoryRow
End Sub

' MODULE: Sheet3
Option Explicit


' MODULE: Sheet100
Option Explicit


' MODULE: bt_BackTest_Main_Multi
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
'    Dim addin_book As Workbook

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
    Call Init_Bt_Settings_Sheets(btWs, btC, _
            activeInstrumentsList, instrLotGroup, stratFdPath, stratNm, _
            dateFrom, dateTo, htmlCount, _
            dateFromStr, dateToStr, btNextFreeRow, _
            maxHtmlCount, repType, macroVer, depoIniCheck, _
            rdRepNameCol, rdRepDateCol, rdRepCountCol, _
            rdRepDepoIniCol, rdRepRobotNameCol, rdRepTimeFromCol, _
            rdRepTimeToCol, rdRepLinkCol)
    If UBound(activeInstrumentsList) = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Не выбраны инструменты."
        Exit Sub
    End If
    ' Separator - autoswitcher
    Call Separator_Auto_Switcher(currentDecimal, undoSep, undoUseSyst)
    upperB = UBound(activeInstrumentsList)
    ' LOOP THRU many FOLDERS
    For i = 1 To upperB
        loopInstrument = activeInstrumentsList(i)
        statusBarFolder = "Папок в очереди: " & upperB - i + 1 & " (" & upperB & ")."
        Application.StatusBar = statusBarFolder
        oneFdFilesList = ListFiles(stratFdPath & "\" & activeInstrumentsList(i))
        ' LOOP THRU FILES IN ONE FOLDER
        openFail = False
        Call Loop_Thru_One_Folder
        If openFail Then
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
    
    statusSaving = statusBarFolder & " Сохраняюсь..."
    Application.StatusBar = statusSaving
' core name
    corenm = folderToSave & stratNm & "-" & UCase(loopInstrument) & "-" & dateFromStr & "-" & dateToStr & "-r" & ov(s_ov_htmls, 2)
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
    btC(btNextFreeRow, rdRepLinkCol) = "открыть"
    btWs.Hyperlinks.Add Anchor:=btC(btNextFreeRow, rdRepLinkCol), Address:=fNm
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
    
    For i = 3 To Sheets.Count
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
        correctRobName = GetCorrectRobName(c(2, 2))
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
    Sheets(2).rows("1:1").AutoFilter
    Sheets(2).rows("1:1").AutoFilter
    
    Set rng_check = mb.Sheets(2).Range(Cells(2, add_c1), Cells(Sheets.Count - 1, add_c2))
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
    If mb.Sheets.Count - 2 = htmlCount Then
        btC(btNextFreeRow, rdRepCountCol) = "ok"
    Else
        With btC(btNextFreeRow, rdRepCountCol)
            .Value = "error"
            .Interior.Color = RGB(255, 0, 0)
        End With
    End If

    ' result of checking depo_ini, into addin
    Set rng_check = mb.Sheets(2).Range(Cells(2, add_c3), Cells(Sheets.Count - 1, add_c3))
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
    Set rng_check = mb.Sheets(2).Range(Cells(2, add_c4), Cells(Sheets.Count - 1, add_c4))
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
    If mb.Sheets.Count > 2 Then
        For i = 1 To mb.Sheets.Count - 2
            Application.DisplayAlerts = False
                Sheets(mb.Sheets.Count).Delete
            Application.DisplayAlerts = True
        Next i
    ElseIf mb.Sheets.Count < 2 Then
        mb.Sheets.Add after:=mb.Sheets(mb.Sheets.Count)
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
        Application.StatusBar = statusBarFolder & " " & sta
        Set rb = Workbooks.Open(oneFdFilesList(i))
        Set hs = mb.Sheets.Add(after:=mb.Sheets(mb.Sheets.Count))
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
    os.columns(1).AutoFit
    os.columns(2).AutoFit
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
                ss.Hyperlinks.Add Anchor:=sc(r + 1, 1), Address:="", SubAddress:="'00" & r & "'!R" & hr & "C2"
            Case 10 To 99
                ss.Hyperlinks.Add Anchor:=sc(r + 1, 1), Address:="", SubAddress:="'0" & r & "'!R" & hr & "C2"
            Case Else
                ss.Hyperlinks.Add Anchor:=sc(r + 1, 1), Address:="", SubAddress:="'" & r & "'!R" & hr & "C2"
        End Select
    Next r
' add autofilter
    ss.Activate
    ss.rows("1:1").AutoFilter
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
    hs.Hyperlinks.Add Anchor:=hc(UBound(SV) + 2, 2), Address:="", SubAddress:="'результаты'!A1"

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
    hs.Hyperlinks.Add Anchor:=hc(s_link, 2), Address:=sM(i, 0)
    hs.Activate
    hs.Range(columns(1), columns(2)).AutoFit
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
        LookIn:=xlValues, lookat:=xlWhole, searchorder:=xlByColumns, _
        searchdirection:=xlNext, MatchCase:=False, searchformat:=False)
    If varRow Is Nothing Then
        ins_td_r = rc.Find(what:="Closed positions", after:=rc(10, 1), _
        LookIn:=xlValues, lookat:=xlWhole, searchorder:=xlByColumns, _
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
    SV(s_depo_ini, 2) = rc(5, 2)
'' Finish deposit
'    sv(s_depo_fin, 2) = CDbl(rc(6, 2))
' Commissions
    SV(s_cmsn, 2) = rc(8, 2)
    
    If rc(ins_td_r, 2) = 0 Then
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
    
' get trade log first row - header
    tl_r = rc.Find(what:="Closed orders:", after:=rc(ins_td_r, 1), LookIn:=xlValues, lookat _
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
    oc_fr = rc.Find(what:="Event log:", after:=rc(ins_td_r + SV(s_trades, 2), 1), LookIn:=xlValues, lookat _
        :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
        False, searchformat:=False).Row + 2 ' header row
    oc_lr = rc(oc_fr, 1).End(xlDown).Row
'
    ro_d = 1
    For r = 2 To UBound(t2, 1)
' compare dates
        If t2(r, 1) = Int(CDate(rc(oc_fr + ro_d, 1))) Then
            Do While t2(r, 1) = Int(CDate(rc(oc_fr + ro_d, 1)))
                If rc(oc_fr + ro_d, 2) = "Commissions" Then
                    s = rc(oc_fr + ro_d, 3)
                    s = Right(s, Len(s) - 13)
                    s = Left(s, Len(s) - 1)
                    s = Replace(s, ".", ",", 1, 1, 1)   ' cmsn extracted
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
        MsgBox "GetStats не может обработать более " & maxHtmlCount & " отчетов. Отмена."
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
    Set oFiles = oFolder.files
    If oFiles.Count = 0 Then Exit Function
    ReDim vaArray(1 To oFiles.Count)
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
        .Title = "Выбрать папку стратегии"
        .ButtonName = "Выбрать"
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    stratFdRng = fd.SelectedItems(1)
    stratName = fd.SelectedItems(1)
    stratName = Right(stratName, Len(stratName) - InStrRev(stratName, "\", -1, vbTextCompare))
    stratNmRng = stratName
    columns(1).AutoFit
End Sub

Sub Clear_Ready_Reports()
    Dim lastRow As Integer
    Dim Rng As Range
    
    Call Init_Clear_Ready_Reports(btWs, btC, _
                upperRow, leftCol, rightCol)
    lastRow = btC(btWs.rows.Count, leftCol).End(xlUp).Row
    If lastRow = upperRow - 1 Then
        Exit Sub
    End If
    Set Rng = btWs.Range(btC(upperRow, leftCol), btC(lastRow, rightCol))
    Rng.Clear
End Sub

' MODULE: bt_Inits
Option Explicit

    Const addinFName As String = "GetStats_BackTest_v1.11.xlsm"
    Const settingsSheetName As String = "hSettings"
    Const backSheetName As String = "Бэктест"

    Const maxHtmls As Integer = 999
    Const reportType As String = "GS_Pro_Single_Core"
    Const depoIniOK As Double = 10000

    Const stratFdRow As Integer = 2 ' strategy folder row
    Const stratFdCol As Integer = 1 ' strategy folder column
    Const stratNmRow As Integer = 7 ' strategy name row
    Const stratNmCol As Integer = 1 ' strategy name column

    Const instrFRow As Integer = 2
    Const instrLRow As Integer = 31
    Const instrCol As Integer = 2
    Const instrGrpFRow As Integer = 2
    Const instrGrpLRow As Integer = 31
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
Sub Init_Bt_Settings_Sheets(ByRef btWs As Worksheet, _
            ByRef btC As Range, _
            ByRef activeInstrumentsList As Variant, _
            ByRef instrumentLotGroup As Variant, _
            ByRef stratFdPath As String, _
            ByRef stratNm As String, _
            ByRef dateFrom As Date, _
            ByRef dateTo As Date, _
            ByRef htmlCount As Integer, _
            ByRef dateFromStr As String, _
            ByRef dateToStr As String, _
            ByRef btNextFreeRow As Integer, _
            ByRef maxHtmlCount As Integer, _
            ByRef repType As String, _
            ByRef macroVer As String, _
            ByRef depoIniCheck As Double, _
            ByRef rdRepNameCol As Integer, _
            ByRef rdRepDateCol As Integer, _
            ByRef rdRepCountCol As Integer, _
            ByRef rdRepDepoIniCol As Integer, _
            ByRef rdRepRobotNameCol As Integer, _
            ByRef rdRepTimeFromCol As Integer, _
            ByRef rdRepTimeToCol As Integer, _
            ByRef rdRepLinkCol As Integer)
    Dim setWs As Worksheet
    Dim setC As Range
    Dim instrumentsList As Range
    Dim lastCh As String
    
    Set btWs = Workbooks(addinFName).Sheets(backSheetName)
    Set btC = btWs.Cells
    Set setWs = Workbooks(addinFName).Sheets(settingsSheetName)
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
    btNextFreeRow = btC(btWs.rows.Count, readyRepFCol).End(xlUp).Row + 1
    maxHtmlCount = maxHtmls
    repType = reportType
    macroVer = addinFName
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
    Dim A() As Variant
    Dim i As Integer, j As Integer
    Dim ubndRows As Integer
    ubndRows = lastRow - firstRow + 1
    ReDim A(1 To ubndRows, 1 To 2)
    For i = firstRow To lastRow
        j = i - 1
        A(j, 1) = Rng(i, firstCol)
        A(j, 2) = Rng(i, lastCol)
    Next i
    GetInstrumentLotGroups = A
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
    Set stratFdRng = Workbooks(addinFName).Sheets(backSheetName).Cells(stratFdRow, stratFdCol)
    Set stratNmRng = Workbooks(addinFName).Sheets(backSheetName).Cells(stratNmRow, stratNmCol)
End Sub

Sub Init_Clear_Ready_Reports(ByRef btWs As Worksheet, _
            ByRef btC As Range, _
            ByRef upperRow As Integer, _
            ByRef leftCol As Integer, _
            ByRef rightCol As Integer)
    Set btWs = Workbooks(addinFName).Sheets(backSheetName)
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

' MODULE: bt_JFX_create
Option Explicit
    Const myFraction As Double = 0.0067   ' 0.0067 = 0.67%
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
'
' RIBBON > BUTTON "JFX"
'
' USER GUIDE: COPY FROM Public class to last row with parameters
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
Private Function Index_in_array(ByVal objArr As Variant, ByVal objStr As String) As Integer
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
    first_row = Rng.rows(1).Row
    this_col = Rng.columns(1).Column
    last_row = first_row + Rng.rows.Count - 1
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

' MODULE: bt_Join_intervals
Option Explicit
    Const addinFName As String = "GetStats_BackTest_v1.11.xlsm"
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
'
' SHEET "join" > BUTTON "GO"
'
    Dim i As Integer
    
    Application.ScreenUpdating = False
    Call InitPositionTags(positionTags)
    Call Init_sheet_cells
' sanity #1
    If Check_Target_Source = False Then
        MsgBox "Error. Target or source folders"
        Exit Sub
    End If
'
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
    Dim targetWb As Workbook
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
        Set targetWb = Workbooks.Add
        ' add sheets to targetWB
        Call Change_sheets_count(targetWb, wbs(1).Sheets.Count)
        
' LOOP THROUGH ALL REPORTS
' FIND MATCHING PARAMETERS
' COPY TO TARGET BOOK
        For j = 3 To wbs(1).Sheets.Count
            ' copy initial trades set to target book
            Set wsMain = wbs(1).Sheets(j)
            Set cMain = wsMain.Cells
            lastRow = cMain(wsMain.rows.Count, 3).End(xlUp).Row
            Set Rng = wsMain.Range(cMain(1, 3), cMain(lastRow, 13))
            Set wsTarget = targetWb.Sheets(j)
            Set cTarget = wsTarget.Cells
            Rng.Copy cTarget(1, 3)  ' copy trades
            lastRow = cMain(wsMain.rows.Count, 1).End(xlUp).Row
            Set Rng = wsMain.Range(cMain(23, 1), cMain(lastRow, 2))
            Call Remove_tag_from_parameters(Rng)
            Rng.Copy cTarget(23, 1) ' copy parameters
            ' move parameters to Arr
            Set Rng = wsMain.Range(cMain(23, 2), cMain(lastRow, 2))
            parMain = Parameters_to_arr(Rng, lastRow - 22)
            ' LOOP compare parMain to wsSrch / cSrch
            ' remove tags
            For k = 2 To UBound(wbs)
                For m = 3 To wbs(k).Sheets.Count
                    Set wsSrch = wbs(k).Sheets(m)
                    Set cSrch = wsSrch.Cells
                    Set Rng = wsSrch.Range(cSrch(23, 1), cSrch(lastRow, 2))
                    Call Remove_tag_from_parameters(Rng)
                Next m
            Next k
            ' find matches, copy to target
            For k = 2 To UBound(wbs)
                For m = 3 To wbs(k).Sheets.Count
                    Set wsSrch = wbs(k).Sheets(m)
                    Set cSrch = wsSrch.Cells
                    Set Rng = wsSrch.Range(cSrch(23, 2), cSrch(lastRow, 2))
                    parCompare = Parameters_to_arr(Rng, lastRow - 22)
                    If Parameters_Match(parMain, parCompare) Then
                        lastRMatch = cSrch(wsSrch.rows.Count, 3).End(xlUp).Row
                        Set rngMatch = wsSrch.Range(cSrch(2, 3), cSrch(lastRMatch, 13))
                        nextRTarget = cTarget(wsTarget.rows.Count, 3).End(xlUp).Row + 1
                        rngMatch.Copy cTarget(nextRTarget, 3)
                        ' fill some basic info: date from-to, trades count
                        If k = UBound(wbs) Then
                            Set rngMatch = wsSrch.Range(cSrch(1, 1), cSrch(2, 2))
                            rngMatch.Copy cTarget(1, 1)
                            Set rngMatch = wsSrch.Range(cSrch(3, 1), cSrch(22, 1))
                            rngMatch.Copy cTarget(3, 1)
                            cTarget(8, 2) = targetDateFromDt
                            cTarget(9, 2) = targetDateToDt
                            cTarget(11, 2) = cTarget(wsTarget.rows.Count, 3).End(xlUp).Row - 1
                        End If
                    End If
                Next m
            Next k
        Next j
        ' save & close all
        Application.StatusBar = appSta & " Saving target book " & i & "."
        targetWBName = Target_WB_Name(wbs(1).Name)
        targetWb.SaveAs fileName:=targetWBName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        targetWb.Close
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
    Const shNameOne As String = "сводка"
    Const shNameTwo As String = "результаты"
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
    someWB.Sheets(3).Activate
End Sub
Private Sub Pick_target_folder()
'
' SHEET "join" > BUTTON "Выбрать целевую папку"
'
' sub adds a folder path to cells(2, 1)
' in "Source folders" column (1)
    Dim fd As FileDialog
    
    Application.ScreenUpdating = False
    Call Init_sheet_cells
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "GetStats: Выбрать целевую папку"
        .ButtonName = "Выбрать"
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    cJ(targetFdRow, 1) = fd.SelectedItems(1)
    wsJ.columns(1).AutoFit
    Application.ScreenUpdating = True
End Sub
Private Sub Add_source_folder()
'
' SHEET "join" > BUTTON "Добавить источник"
'
' sub adds a folder path to next free row
' in "Source folders" column (1)
    Dim fd As FileDialog
    Dim nextFreeRow As Integer
    
    Application.ScreenUpdating = False
    Call Init_sheet_cells
    nextFreeRow = cJ(wsJ.rows.Count, 1).End(xlUp).Row + 1
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "GetStats: Выбрать папку с XLSX отчетами"
        .ButtonName = "Выбрать"
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    cJ(nextFreeRow, 1) = fd.SelectedItems(1)
    wsJ.columns(1).AutoFit
    Application.ScreenUpdating = True
End Sub
Private Sub Clear_source_list()
'
' SHEET "join" > BUTTON "Очистить"
'
' sub clears processing list (subfolders)
    Dim Rng As Range
    
    Application.ScreenUpdating = False
    Call Init_sheet_cells
    Set Rng = wsJ.Range(cJ(sourceFdFRow, 1), cJ(wsJ.rows.Count, 1))
    Rng.Clear
    Application.ScreenUpdating = True
End Sub
Private Sub Rename_source_files_no_postfix_dates()
'
' SHEET "join" > BUTTON "Переименовать (legacy)"
'
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
    lastRow = cJ(wsJ.rows.Count, 1).End(xlUp).Row
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
    Set wsJ = Workbooks(addinFName).Sheets(joinShName)
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
            matchPath = srcFdInfo(j, 1) & "\" & stratIns & "-" & srcFdInfo(j, 5) & "-" & srcFdInfo(j, 6) & "-" & srcFdInfo(j, 4) & ".xlsx"
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
    
    sourceFdLRow = cJ(wsJ.rows.Count, 1).End(xlUp).Row
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
    
    lastRow = cJ(wsJ.rows.Count, 1).End(xlUp).Row
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
    Set oFiles = oFolder.files
    If oFiles.Count = 0 Then Exit Function
    ReDim vaArray(1 To oFiles.Count)
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
    ubnd = Rng(ws.rows.Count, 3).End(xlUp).Row - 1
    If Rng(1, 15) <> "" Then
        Call GSPR_Remove_Chart2
        lr_dates = Rng(ws.rows.Count, 15).End(xlUp).Row
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
Function Get_Calendar_Days_Equity2(ByVal tset As Variant, _
                                   ByVal wc As Range) As Variant
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
Private Sub Print_2D_Array2(ByVal print_arr As Variant, ByVal is_inverted As Boolean, _
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
    cRes(1, 1) = "№_ссылка"
    For i = parFRow To parLRow
        cRes(1, j) = clz(i, 1)
        j = j + 1
    Next i
' copy parameters
    For i = 3 To Sheets.Count
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
        wsRes.Hyperlinks.Add Anchor:=cRes(j, 1), Address:="", SubAddress:="'" & repNum & "'!R22C2"
        ' print "back to summary" link
        With c(22, 2)
            .Value = "результаты"
            .HorizontalAlignment = xlRight
        End With
        ws.Hyperlinks.Add Anchor:=c(22, 2), Address:="", SubAddress:="'результаты'!A1"
    Next i
    wsRes.Activate
    cRes(2, 2).Activate
    wsRes.rows("1:1").AutoFilter
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = True
End Sub

' MODULE: bt_Mixer
Option Explicit
Option Base 1
    Const close_date_col As Integer = 8
    Const depo_ini As Integer = 10000
Private Sub GSPR_show_sheet_index()
'
' RIBBON > BUTTON "Индекс"
'
    Const msg As String = "Лист номер "
    
    On Error Resume Next
    MsgBox msg & ActiveSheet.Index & "."
End Sub
Private Sub GSPR_Go_to_sheet_index()
'
' RIBBON > BUTTON "К листу"
'
    Dim sh_idx As Integer

    On Error Resume Next
    sh_idx = InputBox("Введите номер листа:")
    Sheets(sh_idx).Activate
End Sub
Private Sub GSPR_robo_mixer()
'
' RIBBON > BUTTON "МИКС"
'
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
    Const rf_lower As Double = 0      ' МИНИМАЛЬНЫЙ ФАКТОР ВОССТАНОВЛЕНИЯ
    Const rf_upper As Double = 990    ' МАКСИМАЛЬНЫЙ ФАКТОР ВОССТАНОВЛЕНИЯ
    Const max_tpm As Double = 99      ' МАКСИМАЛЬНОЕ К-ВО СДЕЛОК В МЕСЯЦ
' +++++++++++++++++++++++++++++++++++++
    
    sh_ini = InputBox("Введите номер первого листа для объединения списков сделок:")
    sh_fin = InputBox("Введите номер последнего листа для объединения списков сделок:")
    Application.ScreenUpdating = False
' create / assign "mix" sheet
    If Sheets(Sheets.Count).Name = mix_sheet_name Then
        Set mix_ws = Sheets(Sheets.Count)
        mix_ws.Cells.Clear
    Else
        Set mix_ws = Sheets.Add(after:=Sheets(Sheets.Count))
        mix_ws.Name = mix_sheet_name
    End If
    Set mix_c = mix_ws.Cells
' copy trades to "mix" sheet
    j = 1
    For i = sh_ini To sh_fin ' Sheets.Count - 1
'        Application.StatusBar = i
        Set ws = Sheets(i)
        Set wc = ws.Cells
        If wc(6, 2) >= rf_lower And wc(6, 2) <= rf_upper And wc(3, 2) <= max_tpm Then
            last_row = wc(1, 3).End(xlDown).Row
            If j = 1 Then
                first_row = 1
                next_empty_row = 3
            Else
                first_row = 2
                next_empty_row = mix_c(mix_ws.rows.Count, 1).End(xlUp).Row + 1
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
    mix_c(1, 3) = "р-тов"
    mix_c(2, 3) = algos_mixed
' add autofilter
    mix_ws.Activate
    mix_ws.rows("3:3").AutoFilter
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 3
        .FreezePanes = True
    End With
' sort close date ascending
    last_row = mix_c(3, 1).End(xlDown).Row
    Set Rng = mix_ws.Range(mix_c(3, 1), mix_c(last_row, 11))
    Rng.Sort Key1:=mix_c(3, close_date_col), Order1:=xlAscending, Header:=xlYes
' calculate winning/losing trades
    mix_c(1, 5) = "плюс"
    mix_c(2, 5) = "минус"
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
    mix_c(1, 8) = "дней"
    mix_c(2, 8) = backtest_days
    With mix_c(1, 9)
        .Value = 0.01
        .NumberFormat = "0.00%"
    End With
'    mix_c(2, 10) = "множ"
'    mix_c(2, 11) = 1
    mix_c(1, 11) = "нач.кап."
    mix_c(1, 12) = depo_ini
    mix_c(3, 12) = depo_ini
    mix_c(3, 13) = depo_ini
' calculate trade-to-trade equity curve, hwm
    mix_c(3, 14) = "dd"
    For i = 4 To last_row
        If i Mod 100 = 0 Then
            Application.StatusBar = "Adding formula " & i & " (" & last_row & ")."
        End If
        mix_c(i, 12).FormulaR1C1 = "=R[-1]C*(1+RC[-1]*R1C9*100)"
        mix_c(i, 13).FormulaR1C1 = "=MAX(RC[-1],R[-1]C)"
        With mix_c(i, 14)
            .FormulaR1C1 = "=(RC[-1]-RC[-2])/RC[-1]"
            .NumberFormat = "0.00%"
        End With
    Next i
    mix_c(2, 11) = "кон.кап."
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
    mix_c(1, 17) = "Восст"
    With mix_c(2, 17)
        .FormulaR1C1 = "=R2C16/R2C14"
        .NumberFormat = "0.00"
    End With
    mix_c(1, 18) = "Сделок"
    mix_c(2, 18).FormulaR1C1 = "=COUNT(R4C8:R" & last_row & "C8)"
    mix_c(1, 19) = "В мес"
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
'    mix_ws.Name = mix_sheet_name & "_" & algos_mixed & "_" & max_tpm & "_" & rf_lower & "-" & rf_upper
    mix_ws.Name = mix_sheet_name & "_" & algos_mixed & "_" & Sheets.Count
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Готово"
End Sub
Private Sub GSPR_trades_to_days()
'
' RIBBON > BUTTON "График М"
'
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
    ch_title = "Все роботы (" & wc(2, 3).Value & "), год=" & Round(wc(2, 16).Value * 100, 0) & "%, фин.=" & Round(day_fr(UBound(day_fr)), 0) & " usd. Лог. шкала."
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
Private Sub Merged_Chart_Classic_wMinMax(ulr As Integer, ulc As Integer, _
                                    rngX As Range, rngY As Range, _
                                    ChTitle As String, _
                                    MinVal As Long, maxVal As Currency)
    Dim chW As Integer, chH As Integer          ' chart width, chart height
    Dim chFontSize As Integer                   ' chart title font size
    
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
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
End Sub
Private Sub GSPR_Mixer_Copy_Sheet_To_Book()
'
' RIBBON > BUTTON "В микс"
'
' Copy sheet to mixer.xlsx book
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
    sh_copy.Copy after:=wb_to.Sheets(wb_to.Sheets.Count)
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

' MODULE: bt_Rep_Extra
Option Explicit
Option Base 1
    Const rep_type As String = "GS_Pro_Single_Core"
    Dim ch_rep_type As Boolean
' macro version
    Const macro_name As String = "GetStats Pro v1.11"
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
'
' RIBBON > BUTTON "Экстра"
'
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.StatusBar = "Создаю отчет ""Экстра""."
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
        MsgBox "Неправильный формат адреса. Скопируйте из браузера.", 48, "GetStats Pro"
        Exit Sub
' 2. wrong address
    ElseIf Left(rep_adr, 8) <> ctrl_str Then
        open_fail = True
        MsgBox "Неправильный формат адреса. Скопируйте из браузера.", 48, "GetStats Pro"
        Exit Sub
    ElseIf Right(rep_adr, 5) <> ".html" Then
        open_fail = True
        MsgBox "Неправильный формат адреса. Скопируйте из браузера.", 48, "GetStats Pro"
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
        MsgBox "Неправильный формат адреса. Скопируйте из браузера.", 48, "GetStats Pro"
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
        ins_td_r = rc.Find(what:="Closed positions", after:=rc(ins_td_r, 1), LookIn:=xlValues, lookat _
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
    tl_r = rc.Find(what:="Closed orders:", after:=rc(ins_td_r, 1), LookIn:=xlValues, lookat _
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
    ' 10. кривая доходности
    ' 11. сумма пунктов
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
    Tlog_head(10) = "Кривая доходности"
    Tlog_head(11) = "Сумма пунктов"
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
    oc_fr = rc.Find(what:="Event log:", after:=rc(ins_td_r + SV(r_tds_closed), 1), LookIn:=xlValues, lookat _
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
    SN(r_s_report) = "ОТЧЕТ"
    SN(r_name) = "Макрос"
    SN(r_type) = "Тип отчета"
    SN(r_date_gen) = "Дата получения"
    SN(r_time_gen) = "Время получения"
    SN(r_file) = "Отчет Dukascopy, ссылка"
    SN(r_s_basic) = "ОСНОВНЫЕ ДАННЫЕ"
    SN(r_strat) = "Название стратегии"
    SN(r_ac) = "Валюта счета (в/с)"
    SN(r_ins) = "Инструмент"
    SN(r_init_depo) = "Начальный депозит"
    SN(r_fin_depo) = "Конечный депозит"
    SN(r_s_return) = "ДОХОДНОСТЬ"
    SN(r_net_pc) = "Чистая, %"
    SN(r_net_ac) = "Чистая, в/с"
    SN(r_mon_won) = "Сумма прибыльных сделок"
    SN(r_mon_lost) = "Сумма убыточных сделок"
    SN(r_ann_ret) = "Годовой прирост, %"
    SN(r_mn_ret) = "Месячный прирост, %"
    SN(r_s_pips) = "ПУНКТЫ"
    SN(r_net_pp) = "Сумма"
    SN(r_won_pp) = "В прибыльных сделках"
    SN(r_lost_pp) = "В убыточных сделках"
    SN(r_per_yr_pp) = "В год, сред."
    SN(r_per_mn_pp) = "В месяц, сред."
    SN(r_per_w_pp) = "В неделю, сред."
    SN(r_s_rsq) = "R-КВАДРАТ"
    SN(r_rsq_tr_cve) = "R-кв по кривой сделок (пп)"
    SN(r_rsq_eq_cve) = "R-кв по кривой капитала"
    SN(r_s_pf) = "ПРОФИТ-ФАКТОР"
    SN(r_pf_ac) = "В в/с"
    SN(r_pf_pp) = "В пунктах"
    SN(r_s_rf) = "КОЭФ. ВОССТАНОВЛЕНИЯ"
    SN(r_rf_ac) = "В в/с"
    SN(r_rf_pp) = "В пунктах"
    SN(r_s_avgs_pp) = "СРЕД.ЗНАЧ. В ПУНКТАХ"
    SN(r_avg_td_pp) = "Сделка"
    SN(r_avg_win_pp) = "Прибыльная"
    SN(r_avg_los_pp) = "Убыточная"
    SN(r_avg_win_los_pp) = "Приб/Убыт"
    SN(r_s_avgs_ac) = "СРЕД.ЗНАЧ. В В/С"
    SN(r_avg_td_ac) = "Сделка"
    SN(r_avg_win_ac) = "Прибыльная"
    SN(r_avg_los_ac) = "Убыточная"
    SN(r_avg_win_los_ac) = "Приб/Убыт"
    SN(r_s_intvl) = "ВРЕМЕННЫЕ ИНТЕРВАЛЫ"
    SN(r_mn_win) = "Месяцев прибыльных"
    SN(r_mn_los) = "Месяцев убыточных"
    SN(r_mn_no_tds) = "Месяцев без сделок"
    SN(r_mn_win_los) = "Мес. приб/убыт"
    SN(r_w_win) = "Недель прибыльных"
    SN(r_w_los) = "Недель убыточных"
    SN(r_w_no_tds) = "Недель без сделок"
    SN(r_w_win_los) = "Нед. приб/убыт"
    SN(r_d_win) = "Дней прибыльных"
    SN(r_d_los) = "Дней убыточных"
    SN(r_d_no_tds) = "Дней без сделок"
    SN(r_d_win_los) = "Дней приб/убыт"
    SN(r_s_act_intvl) = "АКТИВНЫЕ ИНТЕРВАЛЫ"
    SN(r_mn_act) = "Месяцев активных"
    SN(r_mn_act_all) = "Мес. акт/все"
    SN(r_w_act) = "Недель активных"
    SN(r_w_act_all) = "Нед. акт/все"
    SN(r_d_act) = "Дней активных"
    SN(r_d_act_all) = "Дней акт/все"
    SN(r_s_std) = "СТАНДАРТНЫЕ ОТКЛОН."
    SN(r_std_tds_pp) = "Сделки (пп)"
    SN(r_std_tds_ac) = "Сделки (в/с)"
'
    SN(r_s_time) = "ИСТОРИЧЕСКОЕ ОКНО"
    SN(r_dt_begin) = "Начало теста"
    SN(r_dt_end) = "Конец теста"
    SN(r_yrs) = "Лет"
    SN(r_mns) = "Месяцев"
    SN(r_wks) = "Недель"
    SN(r_cds) = "Календарных дней"
    SN(r_s_cmsn) = "КОМИССИИ"
    SN(r_cmsn_amnt_ac) = "Сумма в в/с"
    SN(r_cmsn_avg_per_d) = "Средняя в день в в/с"
    SN(r_s_mdd_ac) = "MDD, MFE В В/С"
    SN(r_mdd_ec_ac) = "MDD от кривой капитала"
    SN(r_mfe_ec_ac) = "MFE от кривой капитала"
    SN(r_abs_hi_ac) = "Абс. максимум"
    SN(r_abs_lo_ac) = "Абс. минимум"
    SN(r_s_mdd_pp) = "MDD, MFE В ПП"
    SN(r_mdd_ec_pp) = "MDD от кривой суммы пп."
    SN(r_mfe_ec_pp) = "MFE от кривой суммы пп."
    SN(r_abs_hi_pp) = "Абс. максимум"
    SN(r_abs_lo_pp) = "Абс. минимум"
    SN(r_s_trades) = "ПОЗИЦИИ"
    SN(r_tds_closed) = "Закрыто"             ' trades closed
    SN(r_tds_per_yr) = "В год"
    SN(r_tds_per_mn) = "В месяц"
    SN(r_tds_per_w) = "В неделю"
    SN(r_tds_max_per_d) = "Максимум в день"
    SN(r_tds_win_count) = "Прибыльных"
    SN(r_tds_los_count) = "Убыточных"
    SN(r_tds_win_pc) = "Прибыльных, %"
    SN(r_tds_lg) = "Лонг"
    SN(r_tds_sh) = "Шорт"
    SN(r_tds_lg_sh) = "Лонг/Шорт"
    SN(r_tds_lg_win_pc) = "Лонг, прибыльные, %"
    SN(r_tds_sh_win_pc) = "Шорт, прибыльные, %"
    SN(r_s_dur) = "ПРОДОЛЖ-ТЬ ПОЗИЦИЙ"
    SN(r_avg_dur) = "Средняя, дней"
    SN(r_avg_win_dur) = "Средняя прибыльная, дней"
    SN(r_avg_los_dur) = "Средняя убыточная, дней"
    SN(r_avg_dur_win_los) = "Средняя приб/убыт"
    SN(r_s_stks) = "СЕРИИ"
    SN(r_stk_win_tds) = "Max прибыльная, поз."
    SN(r_stk_los_tds) = "Max убыточная, поз."
    SN(r_stk_win_mns) = "Max прибыльная, мес."
    SN(r_stk_los_mns) = "Max убыточная, мес."
    SN(r_stk_win_wks) = "Max прибыльная, нед."
    SN(r_stk_los_wks) = "Max убыточная, нед."
    SN(r_stk_win_ds) = "Max прибыльная, дн."
    SN(r_stk_los_ds) = "Max убыточная, дн."
    SN(r_runs_tds) = "Серий, позиций"
    SN(r_zscore_tds) = "Z-оценка, позиции"
    SN(r_runs_wks) = "Серий, нед."
    SN(r_zscore_wks) = "Z-оценка, нед."
    SN(r_s_over) = "ОВЕРНАЙТЫ"
    SN(r_over_amnt_pp) = "Сумма, пп"
    SN(r_ds_over) = "Дней с овернайтами"
    SN(r_dwo_per_mn) = "Дней с оверн. в мес."
    SN(r_s_expo) = "ЭКСПОЗИЦИЯ"
    SN(r_tm_in_tds) = "Дней в позициях"
    SN(r_tm_in_win_tds) = "Дней в прибыльных поз."
    SN(r_tm_in_los_tds) = "Дней в убыточных поз."
    SN(r_tm_win_los) = "Дней в приб/убыт"
    SN(r_s_orders) = "ОРДЕРА"
    SN(r_ord_sent) = "Ордеров отправлено"
    SN(r_ord_tds) = "Ордеров/Позиций"
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
    ns.columns(1).AutoFit
    ns.columns(3).AutoFit
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
    ns.Hyperlinks.Add Anchor:=nc(r_file, 2), Address:=rep_adr
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
    Dlog_head(1) = "Конец дня"
    Dlog_head(2) = "Доход"
    Dlog_head(3) = "Открытие"
    Dlog_head(4) = "Хай"
    Dlog_head(5) = "Лоу"
    Dlog_head(6) = "Закрытие"
    Dlog_head(7) = "Свопы, пп"
    Dlog_head(8) = "Комиссия"
    Dlog_head(9) = "Доход"
    Dlog_head(10) = "Открытие"
    Dlog_head(11) = "Хай"
    Dlog_head(12) = "Лоу"
    Dlog_head(13) = "Закрытие"
' weekly log head
    Wlog_head(1) = "Конец недели"
    Wlog_head(2) = "Доход"
    Wlog_head(3) = "Открытие"
    Wlog_head(4) = "Хай"
    Wlog_head(5) = "Лоу"
    Wlog_head(6) = "Закрытие"
    Wlog_head(7) = "Свопы, пп"
    Wlog_head(8) = "Комиссия"
    Wlog_head(9) = "Доход"
    Wlog_head(10) = "Открытие"
    Wlog_head(11) = "Хай"
    Wlog_head(12) = "Лоу"
    Wlog_head(13) = "Закрытие"
' monthly log head
    Mlog_head(1) = "Конец месяца"
    Mlog_head(2) = "Доход"
    Mlog_head(3) = "Открытие"
    Mlog_head(4) = "Хай"
    Mlog_head(5) = "Лоу"
    Mlog_head(6) = "Закрытие"
    Mlog_head(7) = "Свопы, пп"
    Mlog_head(8) = "Комиссия"
    Mlog_head(9) = "Доход"
    Mlog_head(10) = "Открытие"
    Mlog_head(11) = "Хай"
    Mlog_head(12) = "Лоу"
    Mlog_head(13) = "Закрытие"
End Sub
Private Sub GSPR_Show_All_Logs()
    Dim rI As Integer, cI As Integer, tI_co As Integer
' PARAMETERS
    nc(di_r, di_c) = "ПАРАМЕТРЫ"
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
    nc(di_r, tI_co) = "Журнал позиций"
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
    nc(di_r, tI_co) = "Журнал позиций по календарным дням"
    nc(di_r, tI_co + 1) = "В валюте счета"
    nc(di_r, tI_co + 8) = "В пунктах"
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
    nc(di_r, tI_co) = "Журнал позиций по неделям"
    nc(di_r, tI_co + 1) = "В валюте счета"
    nc(di_r, tI_co + 8) = "В пунктах"
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
    nc(di_r, tI_co) = "Журнал позиций по месяцам"
    nc(di_r, tI_co + 1) = "В валюте счета"
    nc(di_r, tI_co + 8) = "В пунктах"
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
    ChTitle = "СУММА ПУНКТОВ. Стратегия: " & SV(r_strat) & ". Инструмент: " & SV(r_ins) & ". Даты: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & ". Итог = " & SV(r_net_pp) & " пп."
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
    ChTitle = "РЕЗУЛЬТАТ ПОЗИЦИЙ В ПП. Стратегия: " & SV(r_strat) & ". Инструмент: " & SV(r_ins) & ". Даты: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & "."
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
    ChTitle = "КРИВАЯ КАПИТАЛА ПО ДНЯМ. Стратегия: " & SV(r_strat) & ". Инструмент: " & SV(r_ins) & ". Даты: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & ". Итог = " & SV(r_fin_depo) & " " & SV(r_ac) & "."
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
    ChTitle = "ИТОГ ДНЯ В " & SV(r_ac) & ". Стратегия: " & SV(r_strat) & ". Инструмент: " & SV(r_ins) & ". Даты: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & "."
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
    ChTitle = "КРИВАЯ КАПИТАЛА ПО НЕДЕЛЯМ, OHLC. Стратегия: " & SV(r_strat) & ". Инструмент: " & SV(r_ins) & ". Даты: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & ". Итог = " & SV(r_fin_depo) & " " & SV(r_ac) & "."
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
    ChTitle = "ИТОГ НЕДЕЛИ В " & SV(r_ac) & ". Стратегия: " & SV(r_strat) & ". Инструмент: " & SV(r_ins) & ". Даты: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & "."
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
    ChTitle = "КРИВАЯ КАПИТАЛА ПО МЕСЯЦАМ, OHLC. Стратегия: " & SV(r_strat) & ". Инструмент: " & SV(r_ins) & ". Даты: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & ". Итог = " & SV(r_fin_depo) & " " & SV(r_ac) & "."
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
    ChTitle = "ИТОГ МЕСЯЦА В " & SV(r_ac) & ". Стратегия: " & SV(r_strat) & ". Инструмент: " & SV(r_ins) & ". Даты: " & SV(r_dt_begin) & "-" & Int(SV(r_dt_end)) & "."
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
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
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
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
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
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
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
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
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

' MODULE: bt_Rep_Multiple
Option Explicit
Option Base 1
    Const addin_file_name As String = "GetStats_BackTest_v1.11.xlsm"
    Const rep_type As String = "GS_Pro_Single_Core"
    Const macro_ver As String = "GetStats Pro v1.11"
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
        .Title = "GetStats: Выбрать отчеты GetStats, объединение по Sharpe Ratio"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Отчеты надстройки GetStats", "*.xlsx"
        .ButtonName = "Вперед"
    End With
    If fd.Show = 0 Then
        MsgBox "Файлы не выбраны!"
        Exit Sub
    End If
    wbksSelected = fd.SelectedItems.Count
    ' Create new workbook with 1 sheet
    Call Create_WB_N_Sheets(wbA, 1)
    
    Application.ScreenUpdating = False
    For i = 1 To wbksSelected
        Application.StatusBar = "Добавляю лист " & i & " (" & wbksSelected & ")."
        Set wbB = Workbooks.Open(fd.SelectedItems(i))
        ' Add Parameters to summary sheet
        Call Params_To_Summary_Sharpe(wbB)
        ' Calculate Sharpe ratios
        Call SharpeBeforeMerge(wbB)
        
        
        tstr = wbB.Name
        pos = InStr(1, tstr, "-", 1)
        tstr = Right(Left(tstr, pos + 6), 6)
        If wbB.Sheets(2).Name = "результаты" Then
            wbB.Sheets("результаты").Copy after:=wbA.Sheets(wbA.Sheets.Count)
            Set s = wbA.Sheets(wbA.Sheets.Count)
            s.Name = i & "_" & tstr
            lr = s.Cells(1, 1).End(xlDown).Row
            Set rg = s.Range(s.Cells(2, 1), s.Cells(lr, 1))
            rg.Hyperlinks.Delete
            s.rows(1).EntireRow.Insert
            s.Cells(1, 1) = "Открыть файл: " & wbB.Name
            s.Hyperlinks.Add Anchor:=s.Cells(1, 1), Address:=wbB.Path & "\" & wbB.Name
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
    MsgBox "Готово. Сохраните файл """ & wbA.Name & """ по вашему усмотрению.", , "GetStats Pro"
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
    cRes(1, 1) = "№_ссылка"
    For i = parFRow To parLRow
        cRes(1, j) = clz(i, 1)
        j = j + 1
    Next i
' copy parameters
    For i = 3 To Sheets.Count
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
        wsRes.Hyperlinks.Add Anchor:=cRes(j, 1), Address:="", SubAddress:="'" & repNum & "'!R22C2"
        ' print "back to summary" link
        With c(22, 2)
            .Value = "результаты"
            .HorizontalAlignment = xlRight
        End With
        ws.Hyperlinks.Add Anchor:=c(22, 2), Address:="", SubAddress:="'результаты'!A" & j
    Next i
    wsRes.Activate
    cRes(2, 2).Activate
    wsRes.rows("1:1").AutoFilter
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
    For i = 3 To wb.Sheets.Count
        wb.Sheets(i).Activate
        Set cSh = wb.Sheets(i).Cells
        Call Calc_Sharpe_Ratio_Sheet
        
        With c(i - 1, new_col)
            .Value = cSh(21, 2)
            .NumberFormat = "0.00"
        End With
    Next i
    wb.Sheets(2).Activate
    rows(1).AutoFilter
    rows(1).AutoFilter
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
        last_row = Cells(rows.Count, 13).End(xlUp).Row
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
            cagr = (1 + net_return) ^ (365 / days_count) - 1
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
        MsgBox "Файлы не выбраны!"
        Exit Sub
    End If
    ov(s_ov_htmls, 2) = fd.SelectedItems.Count
    If ov(s_ov_htmls, 2) > max_htmls Then
        MsgBox "GetStats не может обработать более " & max_htmls & " отчетов. Отмена."
        open_fail = True
        Exit Sub
    End If
    ov(s_ov_htmls, 2) = fd.SelectedItems.Count
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
    If mb.Sheets.Count > 2 Then
        For i = 1 To mb.Sheets.Count - 2
            Application.DisplayAlerts = False
                Sheets(mb.Sheets.Count).Delete
            Application.DisplayAlerts = True
        Next i
    ElseIf mb.Sheets.Count < 2 Then
        mb.Sheets.Add after:=mb.Sheets(mb.Sheets.Count)
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
        Set hs = mb.Sheets.Add(after:=mb.Sheets(mb.Sheets.Count))
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
        ins_td_r = rc.Find(what:="Closed positions", after:=rc(ins_td_r, 1), LookIn:=xlValues, lookat _
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
    
' get trade log first row - header
    tl_r = rc.Find(what:="Closed orders:", after:=rc(ins_td_r, 1), LookIn:=xlValues, lookat _
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
    oc_fr = rc.Find(what:="Event log:", after:=rc(ins_td_r + SV(s_trades, 2), 1), LookIn:=xlValues, lookat _
        :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
        False, searchformat:=False).Row + 2 ' header row
    oc_lr = rc(oc_fr, 1).End(xlDown).Row
'
    ro_d = 1
    For r = 2 To UBound(t2, 1)
' compare dates
        If t2(r, 1) = Int(CDate(rc(oc_fr + ro_d, 1))) Then
            Do While t2(r, 1) = Int(CDate(rc(oc_fr + ro_d, 1)))
                If rc(oc_fr + ro_d, 2) = "Commissions" Then
                    s = rc(oc_fr + ro_d, 3)
                    s = Right(s, Len(s) - 13)
                    s = Left(s, Len(s) - 1)
                    s = Replace(s, ".", ",", 1, 1, 1)   ' cmsn extracted
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
    hs.Hyperlinks.Add Anchor:=hc(UBound(SV) + 2, 2), Address:="", SubAddress:="'результаты'!A1"

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
    hs.Hyperlinks.Add Anchor:=hc(s_link, 2), Address:=sM(i, 0)
    hs.Activate
    hs.Range(columns(1), columns(2)).AutoFit
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
    os.columns(1).AutoFit
    os.columns(2).AutoFit
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
                ss.Hyperlinks.Add Anchor:=sc(r + 1, 1), Address:="", SubAddress:="'00" & r & "'!R" & hr & "C2"
            Case 10 To 99
                ss.Hyperlinks.Add Anchor:=sc(r + 1, 1), Address:="", SubAddress:="'0" & r & "'!R" & hr & "C2"
            Case Else
                ss.Hyperlinks.Add Anchor:=sc(r + 1, 1), Address:="", SubAddress:="'" & r & "'!R" & hr & "C2"
        End Select
    Next r
' add autofilter
    ss.Activate
    ss.rows("1:1").AutoFilter
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
    sel_count = fd.SelectedItems.Count
    If sel_count > max_htmls Then
        MsgBox "GetStats не может обработать более " & max_htmls & " отчетов. Отмена."
        Exit Sub
    End If
    Set wbA = Workbooks.Add
    Application.ScreenUpdating = False
    If wbA.Sheets.Count > 1 Then
        Do Until wbA.Sheets.Count = 1
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
            wbB.Sheets("результаты").Copy after:=wbA.Sheets(wbA.Sheets.Count)
            Set s = wbA.Sheets(wbA.Sheets.Count)
            s.Name = i & "_" & tstr
            lr = s.Cells(1, 1).End(xlDown).Row
            Set Rng = s.Range(s.Cells(2, 1), s.Cells(lr, 1))
            Rng.Hyperlinks.Delete
            s.rows(1).EntireRow.Insert
            s.Cells(1, 1) = "Открыть файл: " & wbB.Name
            s.Hyperlinks.Add Anchor:=s.Cells(1, 1), Address:=wbB.Path & "\" & wbB.Name
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
    For i = 2 To Sheets.Count
        Set iter_s = Sheets(i)
        Set iter_c = iter_s.Cells
        lr = iter_c(3, 1).End(xlDown).Row
        Set c_rng = iter_s.Range(iter_c(3, 5), iter_c(lr, 5))
        c_rng.Copy (sc(2, i - 1))
        Set c_rng = ws.Range(sc(2, i - 1), sc(lr - 1, i - 1))
        c_rng.Sort Key1:=sc(2, i - 1), Order1:=xlDescending, Header:=xlNo
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
        ws.Hyperlinks.Add Anchor:=sc(1, i - 1), Address:="", SubAddress:="'" & Sheets(i).Name & "'!A1"
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
    For i = 2 To Sheets.Count
        Set iter_s = Sheets(i)
        Set iter_c = iter_s.Cells
        lr = iter_c(3, 1).End(xlDown).Row
        lastCol = iter_c(2, 1).End(xlToRight).Column
        Set c_rng = iter_s.Range(iter_c(3, lastCol), iter_c(lr, lastCol))
        c_rng.Copy (sc(2, i - 1))
        Set c_rng = ws.Range(sc(2, i - 1), sc(lr - 1, i - 1))
        c_rng.Sort Key1:=sc(2, i - 1), Order1:=xlDescending, Header:=xlNo
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
        ws.Hyperlinks.Add Anchor:=sc(1, i - 1), Address:="", SubAddress:="'" & Sheets(i).Name & "'!A1"
        With sc(1, i - 1)
            .Value = Right(iter_s.Name, 6)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
End Sub
Private Sub GSPR_Change_Folder_Link()
'
' RIBBON > BUTTON "Ссылки"
'
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
    If Sheets.Count < 3 Then
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
    For i = 3 To Sheets.Count
        Set ws = Sheets(i)
        Set sc = ws.Cells
        address_string = sc(hyperlink_cell_row, 2).Hyperlinks(1).Address
        report_name = Right(address_string, Len(address_string) - len_subtract_current)
        new_hyperlink = new_prefix & report_name
        sc(hyperlink_cell_row, 2).Hyperlinks(1).Address = new_hyperlink
    Next i
    Application.ScreenUpdating = True
    MsgBox "Гиперссылки на html-отчеты обновлены (всего " & Sheets.Count - 2 & ").", , "GetStats Pro"
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
    Set addin_c = addin_book.Sheets("настройки").Cells
    win_start = addin_c(3, 2)
    win_end = addin_c(4, 2)
    html_count = addin_c(5, 2)
    
    add_c1 = Sheets(2).Cells(1, 1).End(xlToRight).Column + 1
    add_c2 = add_c1 + 1
    add_c3 = add_c2 + 1
    Sheets(2).Cells(1, add_c1) = "start"
    Sheets(2).Cells(1, add_c2) = "end"
    Sheets(2).Cells(1, add_c3) = "depo_ini"
    
    For i = 3 To Sheets.Count
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
    Sheets(2).rows("1:1").AutoFilter
    Sheets(2).rows("1:1").AutoFilter
    
    last_row_reports = addin_c(Workbooks(addin_file_name).Sheets("настройки").rows.Count, 4).End(xlUp).Row + 1
    
    Set rng_check = mb.Sheets(2).Range(Cells(2, add_c1), Cells(Sheets.Count - 1, add_c2))
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
    If mb.Sheets.Count - 2 = html_count Then
        addin_c(last_row_reports, 6) = "ok"
    Else
        With addin_c(last_row_reports, 6)
            .Value = "error"
            .Interior.Color = RGB(255, 0, 0)
        End With
    End If
    ' result of checking depo_ini, into addin
    Set rng_check = mb.Sheets(2).Range(Cells(2, add_c3), Cells(Sheets.Count - 1, add_c3))
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
    For j = 1 To fd.SelectedItems.Count
        Application.StatusBar = j & " (" & fd.SelectedItems.Count & ")."
        Set wbCheck = Workbooks.Open(fd.SelectedItems(j))
        Set wsSummary = wbCheck.Sheets(2)
        add_c1 = wsSummary.Cells(1, 1).End(xlToRight).Column + 1
        add_c2 = add_c1 + 1
        add_c3 = add_c2 + 1
        wsSummary.Cells(1, add_c1) = "start"
        wsSummary.Cells(1, add_c2) = "end"
        wsSummary.Cells(1, add_c3) = "depo_ini"
        
        For i = 3 To wbCheck.Sheets.Count
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
        wsSummary.rows("1:1").AutoFilter
        wsSummary.rows("1:1").AutoFilter
        
        last_row_reports = addin_c(Workbooks(addin_file_name).Sheets("настройки").rows.Count, 4).End(xlUp).Row + 1
        Set rng_check = wsSummary.Range(wsSummary.Cells(2, add_c1), wsSummary.Cells(wbCheck.Sheets.Count - 1, add_c2))
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
        If wbCheck.Sheets.Count - 2 = html_count Then
            addin_c(last_row_reports, 6) = "ok"
        Else
            With addin_c(last_row_reports, 6)
                .Value = "error"
                .Interior.Color = RGB(255, 0, 0)
            End With
        End If
    ' result of checking depo_ini, into addin
        Set rng_check = wsSummary.Range(wsSummary.Cells(2, add_c3), wsSummary.Cells(wbCheck.Sheets.Count - 1, add_c3))
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
    
    For i = 3 To wbCheck.Sheets.Count
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
    wsSummary.rows("1:1").AutoFilter
    wsSummary.rows("1:1").AutoFilter
    
    last_row_reports = addin_c(Workbooks(addin_file_name).Sheets("настройки").rows.Count, 4).End(xlUp).Row + 1
    Set rng_check = wsSummary.Range(wsSummary.Cells(2, add_c1), wsSummary.Cells(wbCheck.Sheets.Count - 1, add_c2))
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
    If wbCheck.Sheets.Count - 2 = html_count Then
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
    For i = 1 To fd.SelectedItems.Count
        Application.StatusBar = i & " (" & fd.SelectedItems.Count & ")."
        Set wbCheck = Workbooks.Open(fd.SelectedItems(i))
        Set ws = wbCheck.Sheets(2)
        Set c = ws.Cells
        remCol0 = c.Find(what:=lastSearchParam, after:=c(1, 1), LookIn:=xlValues, lookat _
            :=xlWhole, searchorder:=xlByRows, searchdirection:=xlNext, MatchCase:= _
            False, searchformat:=False).Column + 1
        If c(1, remCol0) <> "" Then
            remCol9 = c(1, remCol0).End(xlToRight).Column
            ws.Range(columns(remCol0), columns(remCol9)).EntireColumn.Delete
        End If
        wbCheck.Close savechanges:=True
    Next i
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

' MODULE: bt_Rep_Single
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
' RIBBON > BUTTON "Основной"
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
        MsgBox "Ошибка. Неверный формат. Выберите "".html"".", , "GetStats Pro"
        open_fail = True
        Exit Sub
    ElseIf Dir(rep_adr) = "" Then
        MsgBox "Ошибка. Файл не найден.", , "GetStats Pro"
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
        ins_td_r = rc.Find(what:="Closed positions", after:=rc(ins_td_r, 1), LookIn:=xlValues, lookat _
            :=xlWhole, searchorder:=xlByColumns, searchdirection:=xlNext, MatchCase:= _
            False, searchformat:=False).Row
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
    tl_r = rc.Find(what:="Closed orders:", after:=rc(ins_td_r, 1), LookIn:=xlValues, lookat _
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
    oc_fr = rc.Find(what:="Event log:", after:=rc(ins_td_r + SV(s_trades, 2), 1), LookIn:=xlValues, lookat _
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
    wc(UBound(SV) + 2, 1) = "Параметры робота"
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
    ws.Hyperlinks.Add Anchor:=wc(s_link, 2), Address:=rep_adr
    ws.Range(columns(1), columns(2)).AutoFit
End Sub
Private Sub GSPR_Build_Chart_Check_Report_Type()
    ch_rep_type = False
    If ActiveSheet.Cells.Find(what:=rep_type) Is Nothing Then
        MsgBox "Ошибка. Неподходящий тип отчета.", , "GetStats Pro"
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
'
' RIBBON > BUTTON "График"
'
    If ActiveSheet.Cells(11, 2) > 0 Then
        Application.ScreenUpdating = False
        Call GSPR_Build_Charts_Single_Report
        Application.ScreenUpdating = True
    Else
        MsgBox "0 сделок, построение графиков невозможно.", , "GetStats Pro"
    End If
End Sub
Private Sub GSPR_Build_Charts_Singe_Button_EN()
'
' RIBBON > BUTTON "EN"
'
    If ActiveSheet.Cells(11, 2) > 0 Then
        Application.ScreenUpdating = False
        Call GSPR_Build_Charts_Single_Report_EN
        Application.ScreenUpdating = True
    Else
        MsgBox "0 сделок, построение графиков невозможно.", , "GetStats Pro"
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
    If ActiveSheet.Shapes.Count > 0 Then
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
    ChTitle = "Кривая капитала. Стратегия '" & sc(1, 2) & "', " & sc(2, 2) & ", " & sc(8, 2) & "-" & sc(9, 2) & "."
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
    ChTitle = "Результат сделок, в пп. Стратегия '" & sc(1, 2) & "', " & sc(2, 2) & ", " & sc(8, 2) & "-" & sc(9, 2) & "."
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
    If ActiveSheet.Shapes.Count > 0 Then
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
'    ChTitle = "Кривая капитала. Стратегия '" & sc(1, 2) & "', " & sc(2, 2) & ", " & sc(8, 2) & "-" & sc(9, 2) & "."
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
'    ChTitle = "Результат сделок, в пп. Стратегия '" & sc(1, 2) & "', " & sc(2, 2) & ", " & sc(8, 2) & "-" & sc(9, 2) & "."
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
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
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
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
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

' MODULE: bt_SharpeRatio
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
    lastRow = c(wsResult.rows.Count, 1).End(xlUp).Row
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
        MsgBox "Файлы не выбраны!"
        Exit Sub
    End If
    sel_count = fd.SelectedItems.Count
    
    Set wbA = Workbooks.Add
    Application.ScreenUpdating = False
    If wbA.Sheets.Count > 1 Then
        Do Until wbA.Sheets.Count = 1
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
            wbB.Sheets("результаты").Copy after:=wbA.Sheets(wbA.Sheets.Count)
            Set s = wbA.Sheets(wbA.Sheets.Count)
            s.Name = i & "_" & tstr
            ' remove hyperlinks
            lr = s.Cells(1, 1).End(xlDown).Row
            Set Rng = s.Range(s.Cells(2, 1), s.Cells(lr, 1))
            Rng.Hyperlinks.Delete
            ' Add hyperlink to original book into cell "A1"
            s.Hyperlinks.Add Anchor:=s.Cells(1, 1), Address:=wbB.Path & "\" & wbB.Name
        End If
        wbB.Close savechanges:=False
    Next i
    Application.DisplayAlerts = False
    wbA.Sheets(1).Delete
    wbA.Sheets(1).Activate
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Готово. Сохраните файл """ & wbA.Name & """ по вашему усмотрению.", , "GetStats Pro"
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
    shCopy.Copy after:=wbTo.Sheets(wbTo.Sheets.Count)
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
        rngStep = (rngMax - rngMin) / (listVals.Count - 1)
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
'    RngX.Select
    chsht.Shapes.AddChart.Select        ' build chart
    With ActiveChart
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = rngX
        .SeriesCollection(1).Values = rngY
        .ChartType = xlXYScatter
        .Legend.Delete
        .SetElement (msoElementChartTitleAboveChart)    ' chart title position
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
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
    For i = 1 To userSelection.Areas.Count
        aFirstCol = userSelection.Areas.item(i).Column
        aColCount = userSelection.Areas.item(i).columns.Count
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
    If Cells(21, 1) <> "" Then
        Set Rng = Range(Cells(21, 1), Cells(21, 2))
        Rng.Clear
    Else
        last_row = Cells(rows.Count, 13).End(xlUp).Row
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
            cagr = (1 + net_return) ^ (365 / days_count) - 1
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
    For i = 3 To Sheets.Count
        Sheets(i).Activate
        Call Calc_Sharpe_Ratio
        With c(i - 1, new_col)
            .Value = Cells(21, 2)
            .NumberFormat = "0.00"
        End With
    Next i
    Sheets(2).Activate
    rows(1).AutoFilter
    rows(1).AutoFilter
    Application.ScreenUpdating = True
End Sub


' MODULE: bt_Tools
Option Explicit
    Dim current_decimal As String
    Dim undo_sep As Boolean, undo_usesyst As Boolean
    Dim user_switched As Boolean
Sub GSPR_Remove_CommandBar()
    On Error Resume Next
    Application.CommandBars("GSPR-1").Delete
    Application.CommandBars("GSPR-2").Delete
    Application.CommandBars("GSPR-3").Delete
    Application.CommandBars("GSPR-4").Delete
    Application.CommandBars("GSPR-5").Delete
    Application.CommandBars("GSPR-6").Delete
    Application.CommandBars("GSPR-7").Delete
    Application.CommandBars("GSPR-8").Delete
    Application.CommandBars("GSPR-9").Delete
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
' Single report - core
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 351
        .OnAction = "GSPR_Single_Core"
        .TooltipText = "Обработать отчет и вывести основную статистику"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Основной"
    End With
' Single report - extra
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 352 ' 352   126, 59, 630
        .OnAction = "GSPR_Single_Extra"
        .TooltipText = "Получить полную статистику по одному отчету"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Экстра"
    End With
' ROW 2 ===========
' Multiple
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 688
        .OnAction = "GSPRM_Multiple_Main"
        .TooltipText = "Обработать группу отчетов (выбрать из папки)"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Группа"
    End With
' Build charts, single core
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 418
        .OnAction = "GSPR_Build_Charts_Singe_Button"
        .TooltipText = "Построить/удалить графики к отчету типа ""основной"""
        .Control.Style = msoButtonIconAndCaption
        .Caption = "График"
    End With
' Build charts, single core
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 418
        .OnAction = "GSPR_Build_Charts_Singe_Button_EN"
        .TooltipText = "Построить/удалить графики к отчету типа ""основной"""
        .Control.Style = msoButtonIconAndCaption
        .Caption = "EN"
    End With
' ROW 3 ===========
' Copy sheet
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 585
        .OnAction = "GSPR_Copy_Sheet_Next"
        .TooltipText = "Скопировать лист и вставить рядом"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Копия"
    End With
' Delete sheet
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 478
        .OnAction = "GSPR_Delete_Sheet"
        .TooltipText = "Удалить активный лист без возможности восстановления"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Удалить"
    End With
' Separator
    user_switched = False       ' SEPARATOR
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 98
        .OnAction = "GSPR_Separator_Manual_Switch"
        .TooltipText = "Установить рекомендуемый разделитель (точку) или вернуть пользовательские настройки"
        .Control.Style = msoButtonIconAndCaption
'        .Caption = "Разделитель"
    End With
' GSen_Translate
    user_switched = False       ' SEPARATOR
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 84
        .OnAction = "GSPR_EN_Translate"
        .TooltipText = "Translate into EN"
        .Control.Style = msoButtonIconAndCaption
'        .Caption = "Разделитель"
    End With
' ROW 4 ===========
' GSPRM_Merge_Summaries
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 688
        .OnAction = "GSPRM_Merge_Summaries"
        .TooltipText = "Объединить листы ""результаты"" в одну книгу (merged)"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Recovery"
    End With
' GSPR_Change_Folder_Link
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 1576
        .OnAction = "GSPR_Change_Folder_Link"
        .TooltipText = "Обновить гиперссылки на отчеты, если их папка была перемещена или переименована"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Ссылки"
    End With
' GSPR_Mixer_Copy_Sheet_To_Book
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 279    ' 31, 279
        .OnAction = "GSPR_Mixer_Copy_Sheet_To_Book"
        .TooltipText = "Добавить лист в книгу 'mixer'"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "В микс"
    End With
' ROW 5 ===========
'' About
'    Set cControl = cBar5.Controls.Add
'    With cControl
'        .FaceId = 487
'        .OnAction = "GSPR_About_Link"
'        .TooltipText = "Справка о программе ""GetStats Pro"" на VSAtrader.ru"
'        .Control.Style = msoButtonIconAndCaption
'        .Caption = "О GetStats"
'    End With
' Показать индекс листа
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 124
        .OnAction = "GSPR_show_sheet_index"
        .TooltipText = "Индекс листа"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Индекс"
    End With
' Перейти к листу по номеру
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 205
        .OnAction = "GSPR_Go_to_sheet_index"
        .TooltipText = "Перейти на лист по номеру"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "К листу"
    End With
' GSPR_robo_mixer
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 645   ' 601
        .OnAction = "GSPR_robo_mixer"
        .TooltipText = "Объединить списки сделок и выдать статистику"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "МИКС"
    End With
' ROW 6 ===========
' GSPR_trades_to_days
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 424   ' 601
        .OnAction = "GSPR_trades_to_days"
        .TooltipText = "График эквити по объединенному списку сделок"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "График М"
    End With
' Check_Window
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 601   ' 601
        .OnAction = "Check_Window_Bulk"
        .TooltipText = "Проверка окон, счета, кол-ва html"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Проверка"
    End With
' Create_JFX_file_Main
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 28    ' 7, 28, 159, 176
        .OnAction = "Create_JFX_file_Main"
        .TooltipText = "Создать JFX"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "JFX"
    End With
'' Stats_And_Chart - WFA
'    Set cControl = cBar6.Controls.Add
'    With cControl
'        .FaceId = 424
'        .OnAction = "Stats_And_Chart"
'        .TooltipText = "WFA stats&chart"
'        .Control.Style = msoButtonIconAndCaption
'        .Caption = "wfa"
'    End With
' ROW 7 ===========

' Settings_To_Launch_Log
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 7
        .OnAction = "Settings_To_Launch_Log"
        .TooltipText = "Настройки робота из java в журнал"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "java-log"
    End With
' GSPR_trades_to_days
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 424   ' 601
        .OnAction = "Stats_Chart_from_Joined_Windows"
        .TooltipText = "График эквити по объединенным окнам"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "График J"
    End With
' Calc_Sharpe_Ratio
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 435   ' 601
        .OnAction = "Calc_Sharpe_Ratio"
        .TooltipText = "Коэффициент Шарпа"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Sharpe"
    End With
' Params_To_Summary
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 191
        .OnAction = "Params_To_Summary"
        .TooltipText = "Параметры Joined в Результаты"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ParamJ-Summary"
    End With
' Sharpe_to_all
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 477   ' 601
        .OnAction = "Sharpe_to_all"
        .TooltipText = "Коэффициент Шарпа, вся книга"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Sharpe all"
    End With
' Scatter_Sharpe
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 430
        .OnAction = "Scatter_Sharpe"
        .TooltipText = "Параметры и Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ScatterPlots"
    End With
' RemoveScatters
    Set cControl = cBar9.Controls.Add
    With cControl
        .FaceId = 478
        .OnAction = "RemoveScatters"
        .TooltipText = "Удалить все графики"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Удал. граф."
    End With
' GSPRM_Merge_Sharpe
    Set cControl = cBar9.Controls.Add
    With cControl
        .FaceId = 477
        .OnAction = "GSPRM_Merge_Sharpe"
        .TooltipText = "Объединить по Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "SharpeMerge"
    End With
End Sub
Private Sub GSPR_About_Link()
' unused
    ActiveWorkbook.FollowHyperlink Address:="https://vsatrader.ru/getstats/"
End Sub
Private Sub GSPR_Copy_Sheet_Next()
'
' RIBBON > BUTTON "Копия"
'
    Application.ScreenUpdating = False
    ActiveSheet.Copy after:=Sheets(ActiveSheet.Index)
    Application.ScreenUpdating = True
End Sub
Private Sub GSPR_Delete_Sheet()
'
' RIBBON > BUTTON "Удалить"
'
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End Sub
Private Sub GSPR_Separator_Manual_Switch()
'
' RIBBON > BUTTON "S"
'
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
    Const msg_rec As String = "Устанавливаю точку как разделитель целой и дробной части." & vbNewLine & vbNewLine _
                & "Правильный выбор!" & vbNewLine & vbNewLine _
                & "Рекомендовано для быстрой работы GetStats." & vbNewLine & vbNewLine _
                & "Чтобы вернуть ваш разделитель, нажмите ""S"" еще раз."
    Const msg_not As String = "Ваш разделитель - точка. Оставляем, как есть. Это оптимальное решение для быстрой работы GetStats."
    If Application.UseSystemSeparators Then     ' SYS - ON
        If Not Application.International(xlDecimalSeparator) = "." Then
            Application.UseSystemSeparators = False
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            MsgBox msg_rec, , "GetStats Pro"
            undo_sep = True                     ' undo condition 1
            undo_usesyst = True                 ' undo condition 2
            user_switched = True
        Else
            undo_sep = False
            MsgBox msg_not, , "GetStats Pro"
        End If
    Else                                        ' SYS - OFF
        If Not Application.DecimalSeparator = "." Then
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            MsgBox msg_rec, , "GetStats Pro"
            undo_sep = True                     ' undo condition 1
            undo_usesyst = False                ' undo condition 2
            user_switched = True
        Else
            undo_sep = False
            MsgBox msg_not, , "GetStats Pro"
        End If
    End If
End Sub
Private Sub GSPR_Separator_OFF()
    Const msg_not As String = "Ваш разделитель не был изменен. Сейчас это точка. Оставляем, как есть. Это оптимальное решение для быстрой работы GetStats."
    If undo_sep Then
        Application.DecimalSeparator = current_decimal
        If undo_usesyst Then
            Application.UseSystemSeparators = True
        End If
        MsgBox "Возвращаю пользовательский разделитель целой и дробной части." & vbNewLine & vbNewLine _
        & "Чтобы задать рекомендованный разделитель (точку) и ускорить работу программы, снова нажмите ""S"".", , "GetStats Pro"
        user_switched = False
    Else
        MsgBox msg_not, , "GetStats Pro"
    End If
End Sub

' MODULE: WFA_Main
' ИДЕИ:
' ОСТАНАВЛИВАТЬ НА ОС ПЛОХИХ РОБОТОВ
' - ОБСЧЕТ ДЛЯ ПОРТФЕЛЯ - СРАЗУ ВСЕ ПАПКИ - ОДИН СПИСОК ФАЙЛОВ (А НЕ 3-5)
'
' превьюшка для победителей ИС или ОС
'
Option Explicit
    Dim A As Variant            ' RESULT ARRAY for each windowSet
    Dim param As Dictionary
    Dim datesISOS As Variant
    Dim fwdCalDays As Variant
    Dim fwdCalDaysLong As Variant
    Dim scanWb As Workbook
    Dim scanWs As Worksheet
    Dim scanC As Range
    Dim resultArr As Variant
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim targetC As Range
    Dim kpiFormatting As Dictionary
    
    Dim iDir As Integer
    Dim iFile As Integer
    Dim iSheet As Integer
    Dim iWindowSet As Integer
    Dim iPermutation As Integer, permutationID As Integer
    Dim iDateSlot As Integer
    
    Dim newDirPath As String
    Dim sortWs As Worksheet
    Dim sortC As Range
    Dim sortColumnId As Integer

Sub WFA_Run()
' Main sub, runs Walk-Forward Analysis.
    Dim t0 As Double, t9 As Double
    Dim time0 As Double, time9 As Double
    Dim exitOnError As Boolean
    Dim errorMsg As String
    
    Dim filesArr As Variant
    Dim tradeList As Variant
    Dim kpisDict As Dictionary
    Dim critDict As Dictionary
    Dim candidateWinner As Variant
    
    Dim reportRam As Variant
    Dim wbBaseName As String

    Dim dirsCount As Integer, winSetsCount As Integer, filesCount As Integer, sheetsCount As Integer
    Dim sDir As String, sWindowSet As String, sFile As String, sSheet As String
    Dim srcStr As String
    
    Dim startIS As Date, endIS As Date
    Dim startOS As Date, endOS As Date
    Dim maxPermutations As Integer
    Dim fractionMultiplier As Double


    ' debug
    Dim wbD As Workbook
    Dim wsD As Worksheet
    Dim cD As Range
    Dim wfaMainCells As Range
    Set wfaMainCells = ActiveSheet.Cells
'    time0 = Timer

    Application.ScreenUpdating = False
    exitOnError = False
    Call Init_Parameters(param, exitOnError, errorMsg, sortWs, sortC, sortColumnId, kpiFormatting)
    
'=================================================================================
    
    ' Loop thru directories
    dirsCount = GetNewDirsCount(param("Scan table"), param("Scan mode"))
    For iDir = 1 To dirsCount
        sDir = "Directory " & iDir & "/" & dirsCount & ". "
        newDirPath = GetNewDirPath( _
                param("Scan table"), _
                param("Scan mode"), _
                iDir, _
                param("Target directory"))
        ' Create new dir
        MkDir newDirPath

        ' Prepare files array, to scan
        filesArr = GetFilesArr( _
                param("Source directories"), _
                param("Scan mode"), _
                param("Scan table"), _
                iDir)

        ' Loop thru window sets
        ' 1 RESULT FILE = 1 WINDOW SET
        winSetsCount = UBound(param("IS/OS windows"), 1) ' 260 / 104
        For iWindowSet = 1 To winSetsCount
            sWindowSet = "WindowSet " & iWindowSet & "/" & winSetsCount & ". "
            ' Debug.Print param("IS/OS windows")(1, 3)
            
            ' Get 4 dates - rows
            ' INVERTED
            datesISOS = GetFourDates( _
                    param("Date start"), _
                    param("Date end"), _
                    param("IS/OS windows")(iWindowSet, 1), _
                    param("IS/OS windows")(iWindowSet, 2))
            
            fwdCalDays = GenerateCalendarDays(datesISOS(3, 1), param("Date end"))
            fwdCalDaysLong = GenerateLongDays(UBound(fwdCalDays))

' INITIALIZE RESULT ARRAY
            A = InitializeResultArray(param("Permutations"), datesISOS, param("MaxiMinimize")(1))

            ' Loop thru files list
            filesCount = UBound(filesArr)
            For iFile = LBound(filesArr) To filesCount
                
                sFile = "File " & iFile & "/" & filesCount & ". "
                Set scanWb = Workbooks.Open(filesArr(iFile))
                
                wbBaseName = Left(scanWb.Name, Len(scanWb.Name) - 5) ' Source String - trade comment

                ' Loop thru each sheet
                sheetsCount = Sheets.Count
                
'                ' debug
'                t0 = Timer
'                Call CreateNewWorkbookSheetsCountNames(wbD, sheetsCount)
                
' ***** SHEET ********************************************************
'
                For iSheet = 3 To sheetsCount
                
'                    ' debug
'                    Set cD = wbD.Sheets(iSheet - 2).Cells
                    
                    ' Update status bar
                    sSheet = "Sheet " & iSheet & "/" & sheetsCount & "."
                    Application.StatusBar = sDir & sWindowSet & sFile & sSheet

                    Set scanWs = scanWb.Sheets(iSheet)
                    Set scanC = scanWs.Cells
                    
                    ' Load report to RAM
                    srcStr = wbBaseName & "_" & scanWs.Name ' Source String - trade comment
                    reportRam = LoadReportToRAM(scanWs, srcStr)

                    ' Loop thru each permutation
                    For iPermutation = 2 To UBound(param("Permutations"), 1)
                        permutationID = iPermutation - 1
                        
                        Set critDict = CreateThisCritDict(param("Permutations"), iPermutation)
'                        ' debug:
'                        Call Print_2D_Array(param("Permutations"), False, 0, 0, Cells)

                        ' Loop thru each IS date-slot
                        For iDateSlot = LBound(datesISOS, 2) To UBound(datesISOS, 2)
                            startIS = datesISOS(1, iDateSlot)
                            endIS = datesISOS(2, iDateSlot)
                            startOS = datesISOS(3, iDateSlot)
                            endOS = datesISOS(4, iDateSlot)
                            ' Get TRADELIST
                            tradeList = ApplyDateFilter(reportRam, startIS, endIS) ' Fastest function
                            
                            ' Calculate KPIs for IN-SAMPLE
                            Set kpisDict = CalcKPIs(tradeList, startIS, endIS, _
                                datesISOS(5, iDateSlot), datesISOS(7, iDateSlot))
                            
                            If PassesCriteria(kpisDict, critDict) Then
                                If param("MaxiMinimize")(1) = "none" Then
                                    ' Extend IS Winners UNITED Trade List
                                    A(permutationID)(iDateSlot)(1)(1) = ExtendTradeList( _
                                            A(permutationID)(iDateSlot)(1)(1), _
                                            tradeList)
                                    ' Extend OS Winners UNITED Trade List
                                    A(permutationID)(iDateSlot)(2)(1) = ExtendTradeList( _
                                            A(permutationID)(iDateSlot)(2)(1), _
                                            ApplyDateFilter(reportRam, startOS, endOS))
                                Else
                                    ' Append to candidates list:
                                    ' - IS tradelist
                                    ' - IS KPIs dictionary
                                    ' - OS tradelist
                                    ' if no Maxi/Minimization
                                    A(permutationID)(iDateSlot)(0) = AppendCandidate( _
                                            A(permutationID)(iDateSlot)(0), _
                                            tradeList, _
                                            kpisDict, _
                                            ApplyDateFilter(reportRam, startOS, endOS))
                                End If
                            End If
'                            ' debug
'                            Call Print_2D_Array(tradeList, True, 0, 0, cD)
                        Next iDateSlot
                    Next iPermutation
                Next iSheet
                
'                ' debug
'                t9 = Timer
'                Debug.Print iDir & "-" & iWindowSet & "-" & iFile & ". Time: " & Round(t9 - t0, 5)
' ***** SHEET ********************************************************
                
                scanWb.Close savechanges:=False
                
                ' IF MaxiMinimize
                If param("MaxiMinimize")(1) <> "none" Then
                    
                    ' MAXI/MINIMIZATION IS HERE
                    For iPermutation = 2 To UBound(param("Permutations"), 1)
                        permutationID = iPermutation - 1
                        
                        For iDateSlot = LBound(datesISOS, 2) To UBound(datesISOS, 2)
                            endIS = datesISOS(2, iDateSlot)
                            startOS = datesISOS(3, iDateSlot)
                            startIS = datesISOS(1, iDateSlot)
                            endOS = datesISOS(4, iDateSlot)

                            ' FIND WINNER through maximizing/minimizing
                            candidateWinner = DefineWinner( _
                                    A(permutationID)(iDateSlot)(0), _
                                    param("MaxiMinimize")(1), _
                                    param("MaxiMinimize")(2))

                            ' Push candidate winner to IS array - Winners UNITED Trade List
                            A(permutationID)(iDateSlot)(1)(1) = ExtendTradeList( _
                                    A(permutationID)(iDateSlot)(1)(1), _
                                    candidateWinner(1))
                            
                            ' Push candidate winner to OS array - Winners UNITED Trade List
                            A(permutationID)(iDateSlot)(2)(1) = ExtendTradeList( _
                                    A(permutationID)(iDateSlot)(2)(1), _
                                    candidateWinner(2))
                            
                            ' IMPORTANTE!!!
                            ' Reinitialize Candidates Array here
                            ' to avoid multiplying trade lists
                            A(permutationID)(iDateSlot)(0) = InitCandidatesArray
                            
                        Next iDateSlot
                    Next iPermutation
                    
                End If
            Next iFile
            
            ' Loop thru result array, sort IS & OS trades, calculate their KPIs
            
' =================
            
' Once file is done,
' sort winners, create Forward Compiled
            maxPermutations = UBound(param("Permutations"), 1) - 1
            For iPermutation = 2 To UBound(param("Permutations"), 1)
                permutationID = iPermutation - 1
                
'                Debug.Print "permutationID: " & permutationID
                
                For iDateSlot = LBound(datesISOS, 2) To UBound(datesISOS, 2)
                    Application.StatusBar = sDir & sWindowSet & "Calculations: Permutation " _
                            & permutationID & "/" & maxPermutations & ", DateSlot " & iDateSlot _
                            & "/" & UBound(datesISOS, 2) & "."
'                    Debug.Print "iDateSlot: " & iDateSlot
                    startIS = datesISOS(1, iDateSlot)
                    endIS = datesISOS(2, iDateSlot)
                    startOS = datesISOS(3, iDateSlot)
                    endOS = datesISOS(4, iDateSlot)

                    ' Bubble sort IS array winners
                    A(permutationID)(iDateSlot)(1)(1) = BubbleSort2DArray( _
                            A(permutationID)(iDateSlot)(1)(1), True, True, True, sortColumnId, sortWs, sortC)
                    ' IS - Alter Fraction to Target MDD
                    fractionMultiplier = GetFractionMultiplier( _
                            A(permutationID)(iDateSlot)(1)(1), _
                            param("MDD freedom"), _
                            param("Target MDD"))
                    ' Apply fraction multiplier to "Return" column
                    A(permutationID)(iDateSlot)(1)(1) = ApplyFractionMultiplier( _
                            A(permutationID)(iDateSlot)(1)(1), _
                            fractionMultiplier)
                    ' Calculate KPIs for IS
                    Set A(permutationID)(iDateSlot)(1)(2) = CalcKPIs( _
                            A(permutationID)(iDateSlot)(1)(1), _
                            startIS, endIS, datesISOS(5, iDateSlot), datesISOS(7, iDateSlot))

                    ' Bubble sort OS array winners
                    A(permutationID)(iDateSlot)(2)(1) = BubbleSort2DArray( _
                            A(permutationID)(iDateSlot)(2)(1), True, True, True, sortColumnId, sortWs, sortC)
                    ' Apply fraction multiplier to "Return" column
                    
                    ' debug
'                    Call Print_2D_Array(A(permutationID)(iDateSlot)(2)(1), True, 0, 0, sortC)
                    
                    A(permutationID)(iDateSlot)(2)(1) = ApplyFractionMultiplier( _
                            A(permutationID)(iDateSlot)(2)(1), _
                            fractionMultiplier)
                    
                    ' debug
'                    Call Print_2D_Array(A(permutationID)(iDateSlot)(2)(1), True, 0, 0, sortC)
                    
                    ' Calculate KPIs for OS
                    Set A(permutationID)(iDateSlot)(2)(2) = CalcKPIs( _
                            A(permutationID)(iDateSlot)(2)(1), _
                            startOS, endOS, datesISOS(6, iDateSlot), datesISOS(8, iDateSlot))
                    
                    ' PUSH OS trade list to Compiled Forward
                    A(permutationID)(0)(1) = ExtendTradeList( _
                            A(permutationID)(0)(1), _
                            A(permutationID)(iDateSlot)(2)(1))
                    
                Next iDateSlot
                
                ' BubbleSort Compiled Forward for this permutation
                A(permutationID)(0)(1) = BubbleSort2DArray( _
                        A(permutationID)(0)(1), True, True, True, sortColumnId, sortWs, sortC)
                ' Calculate KPIs for Compiled Forward
                Set A(permutationID)(0)(2) = CalcKPIs( _
                        A(permutationID)(0)(1), _
                        datesISOS(3, 1), _
                        param("Date end"), _
                        fwdCalDays, _
                        fwdCalDaysLong)

            Next iPermutation


' =================

            ' Add new WB for results
            Application.StatusBar = sDir & sWindowSet & "Print, save, close..."

            ' -- sheets = permutations count + summary
            Call CreateNewWorkbookSheetsCountNames(targetWb, UBound(param("Permutations"), 1))

            ' Print the results
            Call PrintResultArraySaveClose
            
            ' Purge result array from memory
            Set A = Nothing

        Next iWindowSet
    Next iDir
    
'=================================================================================

'    time9 = Timer
'    wfaMainCells(20, 1) = Round(time9 - time0, 5)
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub



Sub PrintResultArraySaveClose()
    Const sr_0 As Double = 0.075
    Const sr_1 As Double = 0.1
    Const sr_2 As Double = 1.5
    Const sr_3 As Double = 2
    Const sr_4 As Double = 3
    
    
    Const permutZeroRow As Integer = 9
    Const descrColSpan As Integer = 4
    Const tradeListColSpan As Integer = 7
    Dim datesKPIsColSpan As Integer
    Dim zeroCol As Integer
    Dim printRow As Long, printCol As Integer
    Dim tableRow As Long, tableCol As Long
    Dim printDict As Dictionary
    Dim rg As Range, cell As Range
    Dim hyperLinkLastRow As Integer
    Dim refCol As Integer

' SUMMARY SHEET
    Set targetWs = targetWb.Sheets(1)
    Set targetC = targetWs.Cells
    Call PrintWindowInfo
    
    ' print permutations
    Call Print_2D_Array(param("Permutations"), False, 8, 0, targetC)

    ' print FWD KPIs
    zeroCol = UBound(param("Permutations"), 2) + 1
    
    ' KPIs
    Set printDict = A(1)(0)(2) ' only for header
    printRow = 9
    For tableCol = 0 To printDict.Count - 1
        printCol = zeroCol + tableCol + 1
        targetC(printRow, printCol) = printDict.Keys(tableCol)
        ' window set
        targetC(printRow + 1, printCol) = param("IS/OS windows")(iWindowSet, 1) & "/" & param("IS/OS windows")(iWindowSet, 2)
    Next tableCol
    For tableRow = 2 To UBound(param("Permutations"), 1)
        printRow = tableRow + 9
        ' add hyperlink to index column
        targetWs.Hyperlinks.Add Anchor:=targetC(printRow, 1), Address:="", SubAddress:="'" & tableRow - 1 & "'!A1"

        Set printDict = A(tableRow - 1)(0)(2)
        For tableCol = 0 To printDict.Count - 1
            printCol = zeroCol + tableCol + 1
            ' apply format
            With targetC(printRow, printCol)
                .Value = printDict.Items(tableCol)
                .NumberFormat = kpiFormatting(kpiFormatting.Keys(tableCol))
                If IsNumeric(printDict.Items(tableCol)) Then
                    Select Case tableCol
                        Case Is = 0 ' Sharpe Ratio
                            Select Case printDict.Items(tableCol)
                                Case sr_0 To sr_1
                                    .Interior.Color = RGB(160, 255, 160)
                                Case sr_1 To sr_2
                                    .Interior.Color = RGB(0, 255, 0)
                                Case sr_2 To sr_3
                                    .Interior.Color = RGB(0, 220, 0)
                                Case sr_3 To sr_4
                                    .Interior.Color = RGB(0, 190, 0)
                                Case Is > sr_4
                                    .Interior.Color = RGB(0, 140, 0)
                            End Select
                        Case Is = 2 ' Annualized Return
                            If printDict.Items(tableCol) > 0 Then
                                .Interior.Color = RGB(70, 255, 70)
                                ' highlight R-squared
                                ' GRADIENT
                                
                                
                                ' 0.6
                                ' 0.8
                                ' 0.9
' !!!!!!!!!!!!!!!!!1

                                If targetC(printRow, printCol - 1) >= 0.75 Then
                                    targetC(printRow, printCol - 1).Interior.Color = RGB(0, 255, 0)
                                End If
                            Else
                                .Interior.Color = RGB(255, 70, 0)
                            End If
                    End Select
                End If
            End With
        Next tableCol
    Next tableRow
' Permutations
    printRow = printRow + 3
    With targetC(printRow, 2)
        .Value = "KPI ranges"
        .Font.Bold = True
    End With
    Call Print_2D_Array(param("KPI ranges"), True, printRow, 1, targetC)
    targetWs.columns(1).AutoFit
    
' =================================
' SINGLE REPORTS
    Set printDict = A(1)(0)(2)
    datesKPIsColSpan = printDict.Count + 4
    For iSheet = 2 To targetWb.Sheets.Count
        Set targetWs = targetWb.Sheets(iSheet)
        Set targetC = targetWs.Cells
        permutationID = iSheet - 1
        
' DESCRIPTION
        Call PrintWindowInfo
        
'        ' Maxi/Minimizing
'        targetC(6, 1) = "Maxi/Minimize"
'        targetC(7, 1) = "KPI"
'        targetC(6, 2) = param("MaxiMinimize")(2)
'        targetC(7, 2) = param("MaxiMinimize")(1)
        
        ' Permutation info
        With targetC(9, 1)
            .Value = "Permutation " & permutationID
            .Font.Bold = True
        End With
        ' permutation index columns
        For tableCol = 1 To UBound(param("Permutations"), 2)
            printRow = tableCol + permutZeroRow
            For tableRow = 0 To 1
                printCol = tableRow + 1
                targetC(printRow, printCol) = param("Permutations")(tableRow, tableCol)
            Next tableRow
        Next tableCol
        ' permutation contents
        printCol = 3
        For tableCol = 1 To UBound(param("Permutations"), 2)
            printRow = tableCol + permutZeroRow
            targetC(printRow, printCol) = param("Permutations")(permutationID + 1, tableCol)
        Next tableCol
        targetWs.columns(1).AutoFit
        
' DATES, KPIs
        ' Date slots
        ' header row
        targetC(1, descrColSpan + 1) = "Date from"
        targetC(1, descrColSpan + 2) = "Date to"
        targetC(1, descrColSpan + 3) = "Type"
        ' KPI names
        Set printDict = A(permutationID)(1)(2)(2) ' only for header
        printRow = 1
        For tableCol = 0 To printDict.Count - 1
            printCol = descrColSpan + 4 + tableCol
            targetC(printRow, printCol) = printDict.Keys(tableCol)
        Next tableCol
        
        For tableRow = LBound(datesISOS, 2) To UBound(datesISOS, 2)
            ' IS
            printRow = tableRow * 2
            targetC(printRow, descrColSpan + 1) = CDate(datesISOS(1, tableRow))
            targetC(printRow, descrColSpan + 2) = CDate(datesISOS(2, tableRow))
            
' -------------------------------
            targetC(printRow, descrColSpan + 3) = "IS-" & tableRow
            ' KPIs
            Set printDict = A(permutationID)(tableRow)(1)(2)
            For tableCol = 0 To printDict.Count - 1
                printCol = descrColSpan + 4 + tableCol
                With targetC(printRow, printCol)
                    .Value = printDict.Items(tableCol)
                    .NumberFormat = kpiFormatting(kpiFormatting.Keys(tableCol))
                    .Interior.Color = RGB(197, 217, 241)
                    If IsNumeric(printDict.Items(tableCol)) Then
                        Select Case tableCol
                            Case Is = 2 ' Annualized Return
                                If printDict.Items(tableCol) > 0 Then
                                    .Font.Color = RGB(0, 180, 80)
                                Else
                                    .Font.Color = RGB(255, 0, 0)
                                End If
                        End Select
                    End If
                End With
            Next tableCol
            
            ' OS
            printRow = printRow + 1
            targetC(printRow, descrColSpan + 1) = CDate(datesISOS(3, tableRow))
            targetC(printRow, descrColSpan + 2) = CDate(datesISOS(4, tableRow))
            
' -------------------------------
            targetC(printRow, descrColSpan + 3) = "OS-" & tableRow
            ' KPIs
            Set printDict = A(permutationID)(tableRow)(2)(2)
            For tableCol = 0 To printDict.Count - 1
                printCol = descrColSpan + 4 + tableCol
                With targetC(printRow, printCol)
                    .Value = printDict.Items(tableCol)
                    .NumberFormat = kpiFormatting(kpiFormatting.Keys(tableCol))
                    .Interior.Color = RGB(253, 233, 217)
                    If IsNumeric(printDict.Items(tableCol)) Then
                        Select Case tableCol
                            Case Is = 2 ' Annualized Return
                                If printDict.Items(tableCol) > 0 Then
                                    .Font.Color = RGB(0, 180, 80)
                                Else
                                    .Font.Color = RGB(255, 0, 0)
                                End If
                        End Select
                    End If
                End With
            Next tableCol
        Next tableRow
        printRow = printRow + 1
        targetC(printRow, descrColSpan + 1) = CDate(datesISOS(3, 1))
        targetC(printRow, descrColSpan + 2) = CDate(datesISOS(4, UBound(datesISOS, 2)))
        
' -------------------------------
        targetC(printRow, descrColSpan + 3) = "Forward Compiled"
        ' KPIs
        Set printDict = A(permutationID)(0)(2)
        For tableCol = 0 To printDict.Count - 1
            printCol = descrColSpan + 4 + tableCol
            With targetC(printRow, printCol)
                .Value = printDict.Items(tableCol)
                .NumberFormat = kpiFormatting(kpiFormatting.Keys(tableCol))
                .Interior.Color = RGB(146, 208, 80)
            End With
        Next tableCol
        
' TRADE LISTS
        ' Forward complied
        zeroCol = descrColSpan + datesKPIsColSpan
        targetC(1, zeroCol + 1) = "Forward Compiled"
        Call Print_2D_Array(A(permutationID)(0)(1), True, 1, zeroCol, targetC)
        
        ' Single IS & OS trade lists
        For tableCol = 1 To UBound(A(permutationID))
            zeroCol = zeroCol + tradeListColSpan
            targetC(1, zeroCol + 1) = "IS-" & tableCol
            Call Print_2D_Array(A(permutationID)(tableCol)(1)(1), True, 1, zeroCol, targetC)
            zeroCol = zeroCol + tradeListColSpan
            targetC(1, zeroCol + 1) = "OS-" & tableCol
            Call Print_2D_Array(A(permutationID)(tableCol)(2)(1), True, 1, zeroCol, targetC)
        Next tableCol
        ' hyperlinks from KPIs to IS/OS tradelists
        hyperLinkLastRow = targetC(1, 7).End(xlDown).Row
        Set rg = targetWs.Range(targetC(2, 7), targetC(hyperLinkLastRow, 7))
        For Each cell In rg
            refCol = targetC.Find(what:=cell.Value, after:=targetC(1, 1), _
                    searchorder:=xlByRows).Column
            targetWs.Hyperlinks.Add Anchor:=cell, Address:="", _
                    SubAddress:="'" & targetWs.Name & "'!R1C" _
                    & refCol
            targetWs.Hyperlinks.Add Anchor:=targetC(1, refCol), Address:="", _
                    SubAddress:="'" & targetWs.Name & "'!R" & cell.Row & "C7"
        Next cell
    Next iSheet
' =================================

    ' Save & close target book
    targetWb.SaveAs fileName:=GetTargetWBSaveName( _
            newDirPath, _
            param("IS/OS windows")(iWindowSet, 3), _
            GetBasenameForTargetWb(newDirPath), _
            param("Date start"), _
            param("Date end"))
    targetWb.Close
End Sub
Sub PrintWindowInfo()
    ' Window info
    With targetC(1, 1)
        .Value = "Window set"
        .Font.Bold = True
    End With
    targetC(2, 1) = "In-Sample"
    targetC(3, 1) = "Out-of-Sample"
    targetC(4, 1) = "code"
    targetC(1, 2) = "weeks"
    targetC(1, 3) = "years"
    
    targetC(2, 2) = param("IS/OS windows")(iWindowSet, 1)
    targetC(3, 2) = param("IS/OS windows")(iWindowSet, 2)
    targetC(2, 3) = Round(param("IS/OS windows")(iWindowSet, 1) / 52, 1)
    targetC(3, 3) = Round(param("IS/OS windows")(iWindowSet, 2) / 52, 1)
    targetC(4, 2) = param("IS/OS windows")(iWindowSet, 3)


' Maxi/Minimizing
    targetC(6, 1) = "Maxi/Minimize"
    targetC(7, 1) = "KPI"
    targetC(6, 2) = param("MaxiMinimize")(1)
    targetC(7, 2) = param("MaxiMinimize")(2)
End Sub
Sub ChartForTradeList()
    Const datesFirstCol As Integer = 5
    Dim tset() As Variant
    Dim dset() As Variant
    Dim wc As Range
    Dim ws As Worksheet
    Dim this_col As Integer
    Dim ch_obj_id As Integer
    Dim first_row As Long, last_row As Long
    Dim first_col As Integer, last_col As Integer
    Dim datesStartEnd As Variant
    Dim wfaInSampleLog As Boolean
    Dim finRes As Double
    
    Application.ScreenUpdating = False
    Call InitWfaChart(wfaInSampleLog)
    
    ' sanity check
    Set ws = ActiveSheet
    Set wc = ws.Cells
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
    last_col = first_col + 3
    ch_obj_id = Cells(1, first_col + 1)
    If ch_obj_id > 0 Then
        Call CleanDaysAndChart(ws, wc, ch_obj_id, first_col, last_col)
        wc(1, first_col).Select
        Application.ScreenUpdating = True
        Exit Sub
    End If
    ' move to RAM
    tset = LoadSlotToRAM(wc, first_row, last_row, first_col)
    ' add Calendar x2 columns
    datesStartEnd = GetDateStartEndForChart(wc, first_col, datesFirstCol)
    dset = GetCalendarDaysEquity(tset, datesStartEnd(1), datesStartEnd(2))
    finRes = Round(dset(2, UBound(dset, 2)), 2)
    ' print out
    ws.Range(columns(last_col + 1), columns(last_col + 2)).ColumnWidth = 30
    Call Print_2D_Array(dset, True, 1, first_col + 3, wc)
    ' build chart
    Call WFAChartClassic(wc, 3, first_col, wfaInSampleLog, finRes)
    Application.ScreenUpdating = True
End Sub
Sub ChartForTradeListPreview()
    Const datesFirstCol As Integer = 5
    Dim tset() As Variant
    Dim dset() As Variant
    Dim wc As Range
    Dim ws As Worksheet
    Dim this_col As Integer
    Dim ch_obj_id As Integer
    Dim first_row As Long, last_row As Long
    Dim first_col As Integer, last_col As Integer
    Dim datesStartEnd As Variant
    Dim wfaInSampleLog As Boolean
    Dim finRes As Double
    
    Call InitWfaChart(wfaInSampleLog)
    
    ' sanity check
    Set ws = ActiveSheet
    Set wc = ws.Cells
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
    last_col = first_col + 3
    ch_obj_id = Cells(1, first_col + 1)
    If ch_obj_id > 0 Then
        Call CleanDaysAndChart(ws, wc, ch_obj_id, first_col, last_col)
        wc(1, 1).Select
'        Application.ScreenUpdating = True
        Exit Sub
    End If
    ' move to RAM
    tset = LoadSlotToRAM(wc, first_row, last_row, first_col)
    ' add Calendar x2 columns
    datesStartEnd = GetDateStartEndForChart(wc, first_col, datesFirstCol)
    dset = GetCalendarDaysEquity(tset, datesStartEnd(1), datesStartEnd(2))
    finRes = Round(dset(2, UBound(dset, 2)), 2)
    ' print out
    ws.Range(columns(last_col + 1), columns(last_col + 2)).ColumnWidth = 30
    Call Print_2D_Array(dset, True, 1, first_col + 3, wc)
    ' build chart
    Call WFAChartClassic(wc, 3, first_col, wfaInSampleLog, finRes)
End Sub

Function GetDateStartEndForChart(ByVal wc As Range, _
            ByVal firstCol As Integer, _
            ByVal datesFirstCol As Integer) As Variant
    Dim arr(1 To 2) As Variant
    Dim tradeListType As String
    Dim i As Integer
    Dim lastDateRow As Integer
    
    lastDateRow = wc(1, datesFirstCol).End(xlDown).Row
    tradeListType = wc(1, firstCol)
    For i = 2 To lastDateRow
        If wc(i, datesFirstCol + 2) = tradeListType Then
            arr(1) = wc(i, datesFirstCol)
            arr(2) = wc(i, datesFirstCol + 1)
            Exit For
        End If
    Next i
    GetDateStartEndForChart = arr
End Function
Sub CleanDaysAndChart(ByRef ws As Worksheet, _
            ByRef wc As Range, _
            ByVal ch_obj_id As Integer, _
            ByVal first_col As Integer, _
            ByVal last_col As Integer)
    Dim Rng As Range
    Dim days_last_row As Integer
    
    ActiveSheet.ChartObjects(ch_obj_id).Delete
    Cells(1, first_col + 1).Clear
    Call DecreaseChIndex(ws, wc, ch_obj_id)
    days_last_row = Cells(2, last_col + 1).End(xlDown).Row
    Set Rng = Range(Cells(2, last_col + 1), Cells(days_last_row, last_col + 2))
    Rng.Clear
    ws.Range(columns(last_col + 1), columns(last_col + 2)).ColumnWidth = 8.43
End Sub
Sub DecreaseChIndex(ByRef ws As Worksheet, _
            ByRef wc As Range, _
            ByVal ch_obj_id As Integer)
    Dim i As Integer
    Dim the_last_col As Integer
    Dim the_first_col As Integer
    the_first_col = wc.Find(what:="Forward Compiled", after:=wc(1, 1), _
            searchorder:=xlByRows).Column + 1
    the_last_col = wc(1, ws.columns.Count).End(xlToLeft).Column
    For i = the_first_col To the_last_col + 1 Step 7
        If wc(1, i).Value > ch_obj_id Then
            wc(1, i).Value = wc(1, i).Value - 1
        End If
    Next i
End Sub
Function LoadSlotToRAM(ByVal wc As Range, _
            ByVal first_row As Long, _
            ByVal last_row As Long, _
            ByVal first_col As Integer) As Variant
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
    LoadSlotToRAM = arr
End Function
Function GetCalendarDaysEquity(ByVal tset As Variant, _
            ByVal date_0 As Long, _
            ByVal date_1 As Long) As Variant
    Dim i As Integer, j As Integer
    Dim arr() As Variant
    Dim calendar_days As Long
    
'    date_0 = Int(tset(1, 1))
'    date_1 = Int(tset(2, UBound(tset, 2)))
    calendar_days = date_1 - date_0 + 1
    ReDim arr(1 To 2, 1 To calendar_days)
        ' 1. calendar days
        ' 2. equity curve
    arr(1, 1) = CDate(date_0 - 1)
    arr(2, 1) = 1
    j = 1
    For i = 2 To UBound(arr, 2)
        arr(1, i) = CDate(arr(1, i - 1) + 1)   ' populate with dates
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
    GetCalendarDaysEquity = arr
End Function
Sub NavigateToWfaSheet()
    Dim shIndex As String
    Dim ws As Worksheet
    Dim c As Range
    Application.ScreenUpdating = False
    If ActiveSheet.Name = "Summary" Then
        Set ws = ActiveSheet
        Set c = ws.Cells
        shIndex = c(ActiveCell.Row, 1).Value
        If shIndex <> "" Then
            If IsNumeric(shIndex) Then
                Sheets(shIndex).Activate
            End If
        End If
    ElseIf Sheets(1).Name = "Summary" Then
        Sheets("Summary").Activate
    End If
    Application.ScreenUpdating = True
End Sub
Sub MergeWfaSummaries()
    Dim fd As FileDialog
    Dim initDirPath As String
    Dim userDirPath As String
    Dim okButtonName As String
    Dim filesList As Variant
    Dim fullPath As String
    Dim iFile As Integer
    Dim sourceWb As Workbook
    Dim targetWb As Workbook
    Dim targetWbPath As String
    Dim sourceWs As Worksheet
    Dim sourceC As Range
    Dim targetWs As Worksheet
    Dim targetC As Range
    Dim pasteInitialData As Boolean
    Dim pasteOneKpiName As Boolean
    Dim initLastCol As Integer
    Dim kpisLastCol As Integer
    Dim lastRow As Integer
    Dim rg As Range
    Dim cell As Range
    Dim iCol As Integer
    Dim pasteCol As Integer
    Dim filesCount As Integer
    
'    Dim dirPath As String
'    Dim ws As Worksheet
'    Dim fillC As Range
'    Dim fillRow As Integer
'    Dim fillCol As Integer
'    Dim dialTitle As String
'    Dim okBtnName As String

    Application.ScreenUpdating = False
    
    Call MergeSummaries_Inits(initDirPath, okButtonName)
'
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Merge Summaries - Pick Source"
        .InitialFileName = initDirPath
        .ButtonName = okButtonName
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    userDirPath = CStr(fd.SelectedItems(1))
    filesList = DirectoryFilesList(userDirPath, False, False)
    targetWbPath = GetParentDirectory(userDirPath) & "\wfa-merged-" & GetBasename(userDirPath)
    targetWbPath = PathIncrementIndex(targetWbPath, True)
    
    Call NewWorkbookSheetsCount(targetWb, 1)
    
    Set targetWs = targetWb.Sheets(1)
    Set targetC = targetWs.Cells
    pasteInitialData = True
    pasteOneKpiName = True
    For iFile = LBound(filesList) To UBound(filesList)
        Application.StatusBar = "File " & iFile & "/" & UBound(filesList) & "."
        fullPath = userDirPath & "\" & filesList(iFile)
        Set sourceWb = Workbooks.Open(fullPath)
        Set sourceWs = sourceWb.Sheets(1)
        Set sourceC = sourceWs.Cells
        
        If pasteInitialData Then
            kpisLastCol = sourceC(9, sourceWs.columns.Count).End(xlToLeft).Column
            initLastCol = kpisLastCol - 10
            
            lastRow = sourceC(9, 1).End(xlDown).Row
            
'            lastRow = sourceC(sourceWs.rows.Count, 1).End(xlUp).Row
            Set rg = sourceWs.Range(sourceC(6, 1), sourceC(lastRow, initLastCol))
            rg.Copy targetC(1, 1)
            ' clear hyperlinks
            Set rg = targetWs.Range(targetC(6, 1), targetC(lastRow - 5, 1))
            For Each cell In rg
                cell.Hyperlinks.Delete
            Next cell
            filesCount = UBound(filesList)
        End If
        pasteCol = initLastCol + iFile
        For iCol = initLastCol + 1 To kpisLastCol
            Set rg = sourceWs.Range(sourceC(10, iCol), sourceC(lastRow, iCol))
            rg.Copy targetC(5, pasteCol)
            targetC(lastRow - 4, pasteCol) = "open" ' hyperlink
            targetWs.Hyperlinks.Add Anchor:=targetC(lastRow - 4, pasteCol), _
                    Address:=fullPath
            If pasteOneKpiName Then
                With targetC(4, pasteCol)
                    .Value = sourceC(9, iCol)
                    .Font.Bold = True
                End With
            End If
            pasteCol = pasteCol + UBound(filesList)
        Next iCol
        sourceWb.Close savechanges:=False
        
        If pasteInitialData Then
            pasteOneKpiName = False
            pasteInitialData = False
        End If
    Next iFile
    Application.StatusBar = "Saving target book..."
    targetWb.SaveAs fileName:=targetWbPath
'    targetWb.Close
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
Sub WfaPreviews()
    Const previewHeight As Integer = 50
    Dim iSheet As Integer
    Dim ws As Worksheet
    Dim wc As Range
    Dim doInits As Boolean
    Dim fwdCol As Integer
    Dim exportDir As String
    Dim objChrt As ChartObject
    Dim myChart As Chart
    Dim myFileName As String
    Dim tWs As Worksheet    ' target worksheet
    Dim tC As Range         ' target cells
    Dim insertRow As Integer
    Dim insertCol As Integer
    Dim rg As Range
    Dim img As Shape
    Dim lastRow As Integer
    Dim i As Integer

    Application.ScreenUpdating = False
    Set tWs = ActiveWorkbook.Sheets(1)
    Set tC = tWs.Cells
    If tC(1, 4) = "PREVIEWS" Then
        For Each img In tWs.Shapes
            img.Delete
        Next
        Set rg = tC(9, 1).CurrentRegion
        lastRow = rg.rows.Count + 8
        For i = 11 To lastRow
            With tWs.rows(i)
                .RowHeight = 15
                .VerticalAlignment = xlBottom
            End With
        Next i
        tC(1, 4).Clear
        Application.ScreenUpdating = True
        Exit Sub
    End If
    exportDir = GetParentDirectory(ActiveWorkbook.Path) & "\tmpImgExport"
    If Dir(exportDir, vbDirectory) = "" Then
        MkDir exportDir
    End If
    doInits = True
    For iSheet = 2 To Sheets.Count
        Application.StatusBar = "Making previews: " & iSheet - 1 & "/" & Sheets.Count - 1 & "."
        Set ws = Sheets(iSheet)
        Set wc = ws.Cells
        If doInits Then
            fwdCol = wc.Find(what:="Forward Compiled", after:=wc(1, 1), _
                    searchorder:=xlByRows).Column
            insertRow = 11
            insertCol = tC(insertRow, tWs.columns.Count).End(xlToLeft).Column + 1
            myFileName = exportDir & "\excelImg.gif"
            doInits = False
        End If
        ws.Activate
        wc(1, fwdCol).Activate
        Call ChartForTradeListPreview
        
        If ws.ChartObjects.Count > 0 Then
            With ws.ChartObjects(1).Chart
                .chartTitle.Delete
                .Axes(xlValue).Delete
                .Axes(xlCategory).Delete
                .Axes(xlValue).MajorGridlines.Delete
                .Export fileName:=myFileName, Filtername:="GIF"
            End With
    '        myChart.Export Filename:=myFileName, Filtername:="GIF"
    '        ws.ChartObjects(1).Delete
            ' remove days & chart object number
            Call ChartForTradeListPreview
            
            ' insert as preview
            With tWs.rows(insertRow)
                .RowHeight = previewHeight
                .VerticalAlignment = xlCenter
            End With
            With tWs.Pictures.Insert(myFileName)
                With .ShapeRange
                    .LockAspectRatio = msoTrue
    '                .Width = 50
                    .Height = previewHeight
                End With
                .Left = tC(insertRow, insertCol).Left
                .Top = tC(insertRow, insertCol).Top
                .Placement = 1
                .PrintObject = True
            End With
            ' remove file
            Kill myFileName
        End If
        ' update insert row
        insertRow = insertRow + 1
    Next iSheet
    tC(1, 4) = "PREVIEWS"
'    Kill exportDir & "\*.*"
    RmDir exportDir
    tWs.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
Sub WfaDateSlotPreviews()
    
' Open workbook and activate sheet with WFA source.
' Boolean
    Dim doInits As Boolean
' FileDialog
    Dim fd As FileDialog
' Integer
    Const previewHeight As Integer = 50
    Dim slotRow As Integer
    Dim uscorePos As Integer
' Range
    Dim rg As Range
    Dim cell As Range
    Dim actC As Range
' Shape
    Dim img As Shape
' String
    Dim dirPath As String
    Dim exportDir As String
    Dim filePath As String
    Dim myFileName As String
    Dim sheetName As String
' Workbook
    Dim wb As Workbook
' Worksheet
    Dim actSh As Worksheet
' from 1.11
    Dim ws As Worksheet
    Dim Rng As Range, clr_rng As Range
    Dim ubnd As Long
    Dim lr_dates As Integer
    Dim tradesSet() As Variant
    Dim daysSet() As Variant

    If IsEmpty(ActiveCell) Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Set rg = ActiveCell.CurrentRegion
    Set actSh = ActiveSheet
    Set actC = actSh.Cells
    
    If actC(rg.Row, rg.Column - 2) = "PREVIEWS" Then
        For Each img In actSh.Shapes
            img.Delete
        Next
        For Each cell In rg
            With actSh.rows(cell.Row)
                .RowHeight = 15
                .VerticalAlignment = xlBottom
            End With
        Next cell
        actC(rg.Row, rg.Column - 2).Clear
        Application.ScreenUpdating = True
        Exit Sub
    End If
    dirPath = rg.Cells(1, 1)
    If Not LooksLikeDirectory(dirPath) Then
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        With fd
            .Title = "Pick source folder"
            .AllowMultiSelect = False
            .ButtonName = "Okey Dokey"
        End With
        If fd.Show = 0 Then
            Exit Sub
            Application.ScreenUpdating = True
        End If
        dirPath = CStr(fd.SelectedItems(1))
        rg.Cells(0, 1) = dirPath
        Set rg = ActiveCell.CurrentRegion
    End If

    exportDir = GetParentDirectory(ActiveWorkbook.Path) & "\tmpImgExport"
    myFileName = exportDir & "\excelImg.gif"
    If Dir(exportDir, vbDirectory) = "" Then
        MkDir exportDir
    End If

    doInits = True
    For Each cell In rg
        Application.StatusBar = "Making previews: " & cell.Row - rg.Row & "/" & rg.rows.Count - 1 & "."
        If doInits = True Then
            doInits = False
        Else
            uscorePos = InStrRev(cell, "_", -1, vbTextCompare)
            filePath = dirPath & "\" & Left(cell.Value, uscorePos - 1) & ".xlsx"
            sheetName = Right(cell, Len(cell) - uscorePos)
            Set wb = Workbooks.Open(filePath)
            wb.Sheets(sheetName).Activate
        ' Build chart
            Set ws = ActiveSheet
            Set Rng = ws.Cells
            ubnd = Rng(ws.rows.Count, 3).End(xlUp).Row - 1

            Call GSPR_Remove_Chart3
            If Rng(1, 14) = "date" _
                Or (Rng(1, 14) = vbEmpty And Rng(1, 15) <> vbEmpty) Then
                lr_dates = Rng(ws.rows.Count, 15).End(xlUp).Row
                Set clr_rng = ws.Range(Rng(1, 14), Rng(lr_dates, 16))
                clr_rng.Clear
            End If
            ' move to RAM
            tradesSet = Load_Slot_to_RAM3(Rng, ubnd)
            ' add Calendar x2 columns
            daysSet = Get_Calendar_Days_Equity3(tradesSet, Rng)
            ' print out
            Call Print_2D_Array3(daysSet, True, 0, 14, Rng)
            ' build & export chart
            Call WFA_Chart_Classic3(Rng, 1, 17, myFileName)

            wb.Close savechanges:=False
            ' insert as preview
            actSh.Activate
            With actSh.rows(cell.Row)
                .RowHeight = previewHeight
                .VerticalAlignment = xlCenter
            End With
            With actSh.Pictures.Insert(myFileName)
                With .ShapeRange
                    .LockAspectRatio = msoTrue
    '                .Width = 50
                    .Height = previewHeight
                End With
                .Left = actC(cell.Row, rg.Column - 2).Left
                .Top = actC(cell.Row, rg.Column - 2).Top
                .Placement = 1
                .PrintObject = True
            End With
            ' remove file
            Kill myFileName
        End If
    Next cell
    
    actC(rg.Row, rg.Column - 2) = "PREVIEWS"
'    Kill exportDir & "\*.*"
    RmDir exportDir

    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
Private Sub GSPR_Remove_Chart3()
    Dim img As Shape
    For Each img In ActiveSheet.Shapes
        img.Delete
    Next
End Sub
Function Load_Slot_to_RAM3(ByVal wc As Range, _
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
    Load_Slot_to_RAM3 = arr
End Function
Function Get_Calendar_Days_Equity3(ByVal tset As Variant, _
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
    Get_Calendar_Days_Equity3 = arr
End Function
Private Sub Print_2D_Array3(ByVal print_arr As Variant, ByVal is_inverted As Boolean, _
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
Sub WFA_Chart_Classic3(sc As Range, _
                ulr As Integer, _
                ulc As Integer, _
                ByVal myFileName As String)
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

    With ActiveSheet.ChartObjects(1).Chart
        .chartTitle.Delete
        .Axes(xlValue).Delete
        .Axes(xlCategory).Delete
        .Axes(xlValue).MajorGridlines.Delete
        .Export fileName:=myFileName, Filtername:="GIF"
    End With
'    sc(1, first_col + 1) = chObj_idx
    sc(1, 15).Select
End Sub


Sub WfaDateSlotPreviews_SOURCE()
    Const previewHeight As Integer = 80
    Dim iSheet As Integer
    Dim ws As Worksheet
    Dim wc As Range
    Dim doInits As Boolean
    Dim fwdCol As Integer
    Dim exportDir As String
    Dim objChrt As ChartObject
    Dim myChart As Chart
    Dim myFileName As String
    Dim tWs As Worksheet    ' target worksheet
    Dim tC As Range         ' target cells
    Dim insertRow As Integer
    Dim insertCol As Integer
    Dim rg As Range
    Dim img As Shape
    Dim lastRow As Integer
    Dim i As Integer

    Application.ScreenUpdating = False
    Set tWs = ActiveWorkbook.Sheets(1)
    Set tC = tWs.Cells
    If tC(1, 4) = "PREVIEWS" Then
        For Each img In tWs.Shapes
            img.Delete
        Next
        Set rg = tC(9, 1).CurrentRegion
        lastRow = rg.rows.Count + 8
        For i = 11 To lastRow
            With tWs.rows(i)
                .RowHeight = 15
                .VerticalAlignment = xlBottom
            End With
        Next i
        tC(1, 4).Clear
        Application.ScreenUpdating = True
        Exit Sub
    End If
    exportDir = GetParentDirectory(ActiveWorkbook.Path) & "\tmpImgExport"
    If Dir(exportDir, vbDirectory) = "" Then
        MkDir exportDir
    End If
    doInits = True
    For iSheet = 2 To Sheets.Count
        Application.StatusBar = "Making previews: " & iSheet - 1 & "/" & Sheets.Count - 1 & "."
        Set ws = Sheets(iSheet)
        Set wc = ws.Cells
        If doInits Then
            fwdCol = wc.Find(what:="Forward Compiled", after:=wc(1, 1), _
                    searchorder:=xlByRows).Column
            insertRow = 11
            insertCol = tC(insertRow, tWs.columns.Count).End(xlToLeft).Column + 1
            myFileName = exportDir & "\excelImg.gif"
            doInits = False
        End If
        ws.Activate
        wc(1, fwdCol).Activate
        Call ChartForTradeListPreview
        
        If ws.ChartObjects.Count > 0 Then
            With ws.ChartObjects(1).Chart
                .chartTitle.Delete
                .Axes(xlValue).Delete
                .Axes(xlCategory).Delete
                .Axes(xlValue).MajorGridlines.Delete
                .Export fileName:=myFileName, Filtername:="GIF"
            End With
    '        myChart.Export Filename:=myFileName, Filtername:="GIF"
    '        ws.ChartObjects(1).Delete
            ' remove days & chart object number
            Call ChartForTradeListPreview
            
            ' insert as preview
            With tWs.rows(insertRow)
                .RowHeight = previewHeight
                .VerticalAlignment = xlCenter
            End With
            With tWs.Pictures.Insert(myFileName)
                With .ShapeRange
                    .LockAspectRatio = msoTrue
    '                .Width = 50
                    .Height = previewHeight
                End With
                .Left = tC(insertRow, insertCol).Left
                .Top = tC(insertRow, insertCol).Top
                .Placement = 1
                .PrintObject = True
            End With
            ' remove file
            Kill myFileName
        End If
        ' update insert row
        insertRow = insertRow + 1
    Next iSheet
    tC(1, 4) = "PREVIEWS"
'    Kill exportDir & "\*.*"
    RmDir exportDir
    tWs.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub


' MODULE: Sheet1
Option Explicit


' MODULE: Backtest_Group
Option Explicit

' Select / Deselect instruments
    Dim hiddenSetWs As Worksheet
    Dim btWs As Worksheet
    Dim hiddenSetC As Range
    Dim btC As Range
    Dim selectAll As Range
    Dim instrumentsList As Range

' ProcessHTMLs
    Dim activeInstrumentsList As Variant
    Dim instrLotGroup As Variant
    Dim dateFrom As Date, dateTo As Date
    Dim dateFromStr As String, dateToStr As String
    Dim stratFdPath As String
    Dim stratNm As String
    Dim htmlCount As Integer
    Dim btNextFreeRow As Integer
    Dim maxHtmlCount As Integer


Sub ProcessHTMLs()
' LOOP through folders
    ' LOOP through html files

' RETURNS:
' 1 file per each html folder
    Dim i As Integer
    Dim upperB As Integer
    
    Application.ScreenUpdating = False
    Call Init_Bt_Settings_Sheets(btWs, btC, _
            activeInstrumentsList, instrLotGroup, stratFdPath, stratNm, _
            dateFrom, dateTo, htmlCount, _
            dateFromStr, dateToStr, btNextFreeRow, _
            maxHtmlCount, repType, macroVer, depoIniCheck, _
            rdRepNameCol, rdRepDateCol, rdRepCountCol, _
            rdRepDepoIniCol, rdRepRobotNameCol, rdRepTimeFromCol, _
            rdRepTimeToCol, rdRepLinkCol)
    If UBound(activeInstrumentsList) = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Не выбраны инструменты."
        Exit Sub
    End If
    ' Separator - autoswitcher
    Call Separator_Auto_Switcher(currentDecimal, undoSep, undoUseSyst)
    upperB = UBound(activeInstrumentsList)
    ' LOOP THRU many FOLDERS
    For i = 1 To upperB
        loopInstrument = activeInstrumentsList(i)
        statusBarFolder = "Папок в очереди: " & upperB - i + 1 & " (" & upperB & ")."
        Application.StatusBar = statusBarFolder
        oneFdFilesList = ListFiles(stratFdPath & "\" & activeInstrumentsList(i))
        ' LOOP THRU FILES IN ONE FOLDER
        openFail = False
        Call Loop_Thru_One_Folder
        If openFail Then
            Exit For
        End If
    Next i
    Call Separator_Auto_Switcher_Undo(currentDecimal, undoSep, undoUseSyst)
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Beep
End Sub







Sub DeSelect_Instruments()
    Dim cell As Range
    Application.ScreenUpdating = False
    Call DeSelect_Instruments_Inits( _
            hiddenSetWs, _
            btWs, _
            hiddenSetC, _
            btC, _
            selectAll, _
            instrumentsList)
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
Sub LocateParentDirectory()
' sheet "backtest"
' sub shows file dialog, lets user pick strategy folder
    Dim fd As FileDialog
    Dim parentDirRg As Range ' strategy folder cell
    Dim stratNmRg As Range ' strategy name cell
    Dim fdTitle As String
    Dim fdButton As String
    Call LocateParentDirectory_Inits(parentDirRg, stratNmRg, fdTitle, fdButton)
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = fdTitle
        .ButtonName = fdButton
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    parentDirRg = fd.SelectedItems(1)
    stratNmRg = GetBasename(fd.SelectedItems(1))
    columns(parentDirRg.Column).AutoFit
End Sub

' MODULE: Tools
Option Explicit
Sub Click_Locate_Target()
    Application.ScreenUpdating = False
    Call Click_LocateTarget_AddSource(False)
    Application.ScreenUpdating = True
End Sub
Sub Click_Add_Source()
    Application.ScreenUpdating = False
    Call Click_LocateTarget_AddSource(True)
    Application.ScreenUpdating = True
End Sub
Sub Click_Clear_Sources()
    Dim ws As Worksheet
    Dim clrRng As Range
    
    Application.ScreenUpdating = False
    Call Click_Clear_Sources_Inits(ws, clrRng)
    clrRng.Clear
    Application.ScreenUpdating = True
End Sub
Sub Click_LocateTarget_AddSource(ByRef addSource As Boolean)
' Sub pastes directory path to 'A6'
' callable by other subs above, ignore screen updating
    Dim fd As FileDialog
    Dim dirPath As String
    Dim ws As Worksheet
    Dim fillC As Range
    Dim fillRow As Integer
    Dim fillCol As Integer
    Dim dialTitle As String
    Dim okBtnName As String
    
    Call Click_Locate_Target_Inits(ws, fillC, fillRow, fillCol, _
            dialTitle, okBtnName, addSource)
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = dialTitle
        .AllowMultiSelect = True
        .ButtonName = okBtnName
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    dirPath = CStr(fd.SelectedItems(1))
    fillC(fillRow, fillCol) = dirPath
    ws.columns(fillCol).AutoFit
End Sub
Sub Click_Copy_Dates_From_Selection()
    Dim mainWs As Worksheet
    Dim mainC As Range
    Dim date0Row As Integer
    Dim date9Row As Integer
    Dim datesCol As Integer
    Dim dirsCol As Integer
    Dim srcDirsZeroRow As Integer
    Dim date0 As Date
    Dim date9 As Date
    Dim selectedFilePath As String
    Dim checkCellValue As String
    
    Call Click_Copy_Dates_From_Selection_Inits(mainWs, mainC, _
            date0Row, date9Row, datesCol, dirsCol, srcDirsZeroRow)
' sanity check
' not empty, row after "Source directories", right column
    If Selection.Row > srcDirsZeroRow _
        And Selection.Column = dirsCol _
        And mainC(Selection.Row, Selection.Column) <> "" Then
' highlight
        mainC(date0Row, datesCol).Interior.Color = RGB(255, 0, 0)
        mainC(date9Row, datesCol).Interior.Color = RGB(255, 0, 0)
    
        Application.ScreenUpdating = False
        selectedFilePath = FirstFileFromSelectedDir(mainC, Selection, dirsCol)
        Call ExtractTestDates(selectedFilePath, date0, date9)
        mainC(date0Row, datesCol) = date0
        mainC(date9Row, datesCol) = date9
' remove highlight
        mainC(date0Row, datesCol).Interior.Pattern = xlNone
        mainC(date9Row, datesCol).Interior.Pattern = xlNone
    Else
        MsgBox "Select one of Source Directories.", vbCritical, "Error"
    End If
    Application.ScreenUpdating = True
End Sub
Sub DeSelect_KPIs()
    Dim stgWs As Worksheet
    Dim stgC As Range
    Dim cell As Range
    Dim checkAll As Range
    Dim kpisList As Range
    Application.ScreenUpdating = False
    Call DeSelect_KPIs_Inits(stgWs, stgC, checkAll, kpisList)
    If checkAll Then
        For Each cell In kpisList
            cell = True
        Next cell
    Else
        For Each cell In kpisList
            cell = False
        Next cell
    End If
    Application.ScreenUpdating = True
End Sub
Sub ExtractTestDates(ByRef filePath As String, _
            ByRef date0 As Date, _
            ByRef date9 As Date)
    Dim wbDates As Workbook
    Set wbDates = Workbooks.Open(filePath)
    date0 = wbDates.Sheets(3).Cells(8, 2)
    date9 = wbDates.Sheets(3).Cells(9, 2)
    wbDates.Close savechanges:=False
End Sub

Function FirstFileFromSelectedDir(ByVal mainCells As Range, _
            ByVal userSelection As Range, _
            ByVal srcCol As Integer) As String
    Dim srcRow As Integer
    Dim dirPath As String
    Dim filePath As String
    srcRow = userSelection.Areas.item(1).Row
    dirPath = mainCells(srcRow, srcCol).Value
    filePath = DirectoryFilesList(dirPath, True, False)(1)
    If Right(filePath, 4) = "xlsx" Then
        FirstFileFromSelectedDir = dirPath & "\" & filePath
    Else
        FirstFileFromSelectedDir = ""
    End If
End Function
Sub Print_1D_Array(ByVal print_arr As Variant, _
            ByVal col_offset As Integer, _
            ByVal print_cells As Range)
' Procedure prints any 1-dimensional array in a new Workbook, sheet 1.
' Arguments:
'       1) 1-D array
    
'    Dim wb_print As Workbook
    Dim r As Integer
    Dim print_row As Integer, print_col As Integer
    Dim add_rows As Integer

    If LBound(print_arr) = 0 Then
        add_rows = 1
    Else
        add_rows = 0
    End If
'    Set wb_print = Workbooks.Add
'    Set c_print = wb_print.Sheets(1).cells
    print_col = 1 + col_offset
    For r = LBound(print_arr) To UBound(print_arr)
        print_row = r + add_rows
        print_cells(print_row, print_col) = print_arr(r)
    Next r
End Sub
Sub Print_2D_Array(ByVal print_arr As Variant, _
            ByVal is_inverted As Boolean, _
            ByVal row_offset As Integer, _
            ByVal col_offset As Integer, _
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
Sub GenerateIsOsCodes()
    Dim i As Integer
    Dim rg As Range
    Dim exitError As Boolean
    Dim addInName As String
    exitError = False
    Call GenerateIsOsCodes_Inits(rg, exitError, addInName)
    If exitError = True Then
        MsgBox "Please, provide In-Sample & Out-of-Sample windows.", vbCritical, addInName
        Exit Sub
    End If
    For i = 1 To rg.rows.Count
        rg(i, 3) = "i" & Int(rg(i, 1) / 52) & "o" & Int(rg(i, 2) / 52)
    Next i
End Sub
Sub WFAChartClassic(ByRef sc As Range, _
                ByVal ulr As Integer, _
                ByVal ulc As Integer, _
                ByVal logScale As Boolean, _
                ByVal finRes As Double)
' Build chart for IS/OS/Forward date slots.
'
' ** CONSTANTS **
    Const ch_hght_cells As Integer = 20
    Const ch_wdth_cells As Integer = 5
    Const my_rnd = 0.1
' ** VARIABLES **
' Double
    Dim maxVal As Double
    Dim MinVal As Double
' Integer
    Dim chFontSize As Integer
    Dim chObj_idx As Integer
    Dim last_date_row As Integer
' Range
    Dim rng_to_cover As Range
    Dim rngX As Range
    Dim rngY As Range
' String
    Dim ChTitle As String
    
    chObj_idx = ActiveSheet.ChartObjects.Count + 1
    ChTitle = sc(1, ulc) & ", FinRes = " & finRes
    If Left(sc(1, ulc), 2) = "IS" And logScale Then
        ChTitle = ChTitle & ", log scale"         ' log scale
    End If
    last_date_row = sc(2, ulc + 4).End(xlDown).Row
    chFontSize = 12
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + ch_hght_cells, ulc + ch_wdth_cells))
    Set rngX = Range(sc(2, ulc + 4), sc(last_date_row, ulc + 4))
    Set rngY = Range(sc(2, ulc + 5), sc(last_date_row, ulc + 5))
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
        If Left(sc(1, ulc), 2) = "IS" And logScale Then
            .Axes(xlValue).ScaleType = xlScaleLogarithmic   ' log scale
        End If
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
        .chartTitle.Text = ChTitle
        .chartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
    sc(1, ulc + 1) = chObj_idx
    sc(1, ulc).Select
End Sub
Sub ChartRangesXandY(ByVal rngX As Range, _
            ByVal rngY As Range, _
            ByRef sc As Range, _
            ByVal ulr As Integer, _
            ByVal ulc As Integer, _
            ByVal logScale As Boolean, _
            ByVal selRow As Integer)
' Build chart for IS/OS/Forward date slots.
'
' ** CONSTANTS **
    Const ch_hght_cells As Integer = 19
    Const ch_wdth_cells As Integer = 9
    Const my_rnd = 0.1
' ** VARIABLES **
' Double
    Dim maxVal As Double
    Dim MinVal As Double
' Integer
    Dim chFontSize As Integer
    Dim chObj_idx As Integer
' Range
    Dim rng_to_cover As Range
'    Dim rngX As Range
'    Dim rngY As Range
' String
    
    chObj_idx = ActiveSheet.ChartObjects.Count + 1
    chFontSize = 12
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + ch_hght_cells, ulc + ch_wdth_cells))
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
        If logScale Then
            .Axes(xlValue).ScaleType = xlScaleLogarithmic   ' log scale
        End If
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
        .chartTitle.Delete
'        .ChartTitle.Text = ChTitle
'        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
    sc(selRow, ulc).Select
End Sub
Sub StatementChartRangesXandY(ByVal rngX As Range, _
            ByVal rngY As Range, _
            ByRef sc As Range, _
            ByVal ulr As Integer, _
            ByVal ulc As Integer, _
            ByVal logScale As Boolean, _
            ByVal selRow As Integer, _
            ByVal chartTitle As String)
' Build chart for IS/OS/Forward date slots.
'
' ** CONSTANTS **
    Const ch_hght_cells As Integer = 19
    Const ch_wdth_cells As Integer = 9
    Const my_rnd = 0.02
' ** VARIABLES **
' Double
    Dim maxVal As Double
    Dim MinVal As Double
' Integer
    Dim chFontSize As Integer
    Dim chObj_idx As Integer
' Range
    Dim rng_to_cover As Range
    
    chObj_idx = ActiveSheet.ChartObjects.Count + 1
    chFontSize = 12
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + ch_hght_cells, ulc + ch_wdth_cells))
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
        If logScale Then
            .Axes(xlValue).ScaleType = xlScaleLogarithmic   ' log scale
        End If
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
'        .chartTitle.Delete
        .chartTitle.Text = chartTitle
'        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
    sc(selRow, ulc - 2).Select
End Sub

Sub WfaWinnersRemoveDuplicates()
' WFA single report, any date slot: IS-n, OS-n, Forward Compiled.
' Take source column, copy below, remove duplicats, sort ascending.
' Integer
    Dim firstCol As Integer
    Dim srcCol As Integer
    Dim thisCol As Integer
' Long
    Dim firstRow As Long
    Dim lastRow As Long
    Dim lastRowAlt As Long
' Range
    Dim rg As Range
    
    If IsEmpty(ActiveCell) Then  ' Sanity check
        Exit Sub
    End If
    Application.ScreenUpdating = False
    thisCol = ActiveCell.Column
    If Not IsEmpty(ActiveCell.Offset(0, -1)) Then
        firstCol = Cells(2, thisCol).End(xlToLeft).Column
    Else
        firstCol = thisCol
    End If
    lastRow = Cells(ActiveSheet.rows.Count, firstCol).End(xlUp).Row
    If lastRow < 3 Then  ' Sanity check
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Set rg = ActiveCell.CurrentRegion
    If rg.columns.Count = 1 Then
        ' Remove existing copy #1
        rg.Clear
'        Cells(1, thisCol - 2).Select
    Else
        ' Remove existing copy #2
        srcCol = firstCol + 2
        lastRowAlt = Cells(ActiveSheet.rows.Count, srcCol).End(xlUp).Row
        If lastRow <> lastRowAlt Then
            Set rg = Cells(lastRowAlt, thisCol).CurrentRegion
            rg.Clear
            Application.ScreenUpdating = True
            Exit Sub
        End If
        ' Copy, remove duplicates, sort ascending
        Set rg = Range(Cells(3, srcCol), Cells(lastRow, srcCol))
        rg.Copy Cells(lastRow + 3, srcCol)
        firstRow = lastRow + 3
        lastRow = Cells(ActiveSheet.rows.Count, srcCol).End(xlUp).Row
        Set rg = Range(Cells(firstRow, srcCol), Cells(lastRow, srcCol))
        rg.RemoveDuplicates columns:=1, Header:=xlNo
        Set rg = Cells(firstRow, srcCol).CurrentRegion
        rg.Sort Key1:=rg.Cells(1), Order1:=xlAscending, Header:=xlNo
        rg.Cells(0, 1).Activate
    End If
    Application.ScreenUpdating = True
End Sub
Sub OpenWfaSource()
' Open workbook and activate sheet with WFA source.
' FileDialog
    Dim fd As FileDialog
' Integer
    Dim uscorePos As Integer
' Range
    Dim rg As Range
' String
    Dim dirPath As String
    Dim filePath As String
    Dim sheetName As String
' Workbook
    Dim wb As Workbook
    
    If IsEmpty(ActiveCell) Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Set rg = ActiveCell.CurrentRegion
    dirPath = rg.Cells(1, 1)
    If Not LooksLikeDirectory(dirPath) Then
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        With fd
            .Title = "Pick source folder"
            .AllowMultiSelect = False
            .ButtonName = "Okey Dokey"
        End With
        If fd.Show = 0 Then
            Exit Sub
            Application.ScreenUpdating = True
        End If
        dirPath = CStr(fd.SelectedItems(1))
        rg.Cells(0, 1) = dirPath
    End If
    uscorePos = InStrRev(ActiveCell, "_", -1, vbTextCompare)
    filePath = dirPath & "\" & Left(ActiveCell.Value, uscorePos - 1) & ".xlsx"
    sheetName = Right(ActiveCell, Len(ActiveCell) - uscorePos)
    Set wb = Workbooks.Open(filePath)
    wb.Sheets(sheetName).Activate
    Application.ScreenUpdating = True
End Sub
Function LooksLikeDirectory(ByRef dirPath As String) As Boolean
' Find out if a suggested path is a directory and it exists.
' Boolean
    Dim result As Boolean
    result = False
    If InStr(1, dirPath, "\", vbTextCompare) > 0 Then
        dirPath = StringRemoveBackslash(dirPath)
        result = True
    End If
    LooksLikeDirectory = result
End Function
Sub ManuallyApplyDateFilter()
    Dim i As Integer
    Dim kpis As Dictionary
    Dim calDays As Variant
    Dim calDaysLong As Variant
    Dim tradeList As Variant
    Dim date0 As Variant
    Dim date9 As Variant
    Dim sRow As Integer
    Dim sCol As Integer
    Dim rg As Range
    Dim printRow As Integer
    Dim printCol As Integer
    Dim daysAndEquity As Variant
    Dim someShape As Shape
    Dim rangeX As Range
    Dim rangeY As Range
    Const rowShift As Integer = 5
    Const shiftPrintRow As Integer = 2
    Application.ScreenUpdating = False
    sRow = Selection.Row
    sCol = Selection.Column
    date0 = Cells(sRow, sCol)
    date9 = Cells(sRow, sCol + 1)
    tradeList = GetTradeListFromSheet(ActiveSheet, date0, date9, ActiveWorkbook.Name)
    If Not IsEmpty(Cells(sRow + rowShift + 1, sCol)) Then
        For Each someShape In ActiveSheet.Shapes
            someShape.Delete
        Next someShape
        Set rg = Cells(sRow + rowShift + 1, sCol).CurrentRegion
        rg.Clear
        Set rg = Cells(sRow + shiftPrintRow + 1, sCol).CurrentRegion
        rg.Clear
    End If
    Call Print_2D_Array(tradeList, True, sRow + rowShift, sCol - 1, Cells)
    ' calculate KPIs
    calDays = GenerateCalendarDays(date0, date9)
    calDaysLong = GenerateLongDays(UBound(calDays))
    Set kpis = CalcKPIs(tradeList, date0, date9, calDays, calDaysLong)
    Call PrintDictionary(kpis, True, sRow + shiftPrintRow, sCol - 1, Cells)
    daysAndEquity = GetDailyEquityFromTradeSet(tradeList, date0, date9)
    Call Print_2D_Array(daysAndEquity, False, sRow + rowShift, sCol + 3, Cells)
    Set rangeX = Cells(sRow + rowShift + 1, sCol).CurrentRegion
    Set rangeX = rangeX.Offset(0, 4).Resize(rangeX.rows.Count, rangeX.columns.Count - 5)
    Set rangeY = Cells(sRow + rowShift + 1, sCol).CurrentRegion
    Set rangeY = rangeY.Offset(0, 5).Resize(rangeY.rows.Count, rangeY.columns.Count - 5)
    Call ChartRangesXandY(rangeX, rangeY, Cells, sRow + rowShift + 2, sCol, False, sRow)
    Application.ScreenUpdating = True
End Sub
Sub PrintDictionary(ByVal pDict As Dictionary, _
            ByVal arrHorizontal As Boolean, _
            ByVal rowOffset As Integer, _
            ByVal colOffset As Integer, _
            ByVal pCells As Range)
    Dim prArr As Variant
    prArr = DictionaryToArray(pDict, arrHorizontal)
    Call Print_2D_Array(prArr, False, rowOffset, colOffset, pCells)
End Sub
Sub CreateNewWorkbookSheetsCountNames(ByRef targetWb As Workbook, _
            ByVal newSheetsCount As Integer)
    Dim i As Integer
    Call NewWorkbookSheetsCount(targetWb, newSheetsCount)
    targetWb.Sheets(1).Name = "Summary"
    For i = 2 To targetWb.Sheets.Count
        targetWb.Sheets(i).Name = CStr(i - 1)
    Next i
End Sub
Sub NewWorkbookSheetsCount(ByRef theWb As Workbook, _
            ByVal newShCount As Integer)
    Dim oldSheetsCount As Integer
    oldSheetsCount = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = newShCount
    Set theWb = Workbooks.Add
    Application.SheetsInNewWorkbook = oldSheetsCount
End Sub
Sub StatementChartRangesXYZ(ByVal rngX As Range, _
            ByVal rngY As Range, _
            ByVal rngZ As Range, _
            ByRef sc As Range, _
            ByVal ulr As Integer, _
            ByVal ulc As Integer)
' Build chart for IS/OS/Forward date slots.
'
' ** CONSTANTS **
    Const ch_hght_cells As Integer = 24
    Const ch_wdth_cells As Integer = 13
    Const my_rnd = 0.1
    Const my_rnd2 = 0.05
' ** VARIABLES **
' Double
    Dim minValY As Double
    Dim maxValY As Double
    Dim minValZ As Double
    Dim maxValZ As Double
' Integer
    Const chFontSize As Integer = 12
    Dim chObj_idx As Integer
' Range
    Dim rng_to_cover As Range
' String
    
    chObj_idx = ActiveSheet.ChartObjects.Count + 1
    Set rng_to_cover = Range(sc(ulr, ulc), sc(ulr + ch_hght_cells, ulc + ch_wdth_cells))
    
    minValY = WorksheetFunction.Min(rngY)
    maxValY = WorksheetFunction.Max(rngY)
    minValY = my_rnd * Int(minValY / my_rnd2)
    maxValY = my_rnd * Int(maxValY / my_rnd2) + my_rnd2
    
    minValZ = WorksheetFunction.Min(rngZ)
    maxValZ = WorksheetFunction.Max(rngZ)
    minValZ = my_rnd * Int(minValZ / my_rnd)
    maxValZ = my_rnd * Int(maxValZ / my_rnd) + my_rnd
    rngZ.Select
    ActiveSheet.Shapes.AddChart.Select
    With ActiveSheet.ChartObjects(chObj_idx)
        .Left = Cells(ulr, ulc).Left
        .Top = Cells(ulr, ulc).Top
        .Width = rng_to_cover.Width
        .Height = rng_to_cover.Height
'        .Placement = xlFreeFloating ' do not resize chart if cells resized
    End With
    With ActiveChart
        .SetSourceData Source:=Application.Union(rngX, rngY, rngZ)
        .ChartType = xlLine
'        .Legend.Delete
        .Axes(xlValue).MinimumScale = minValZ
        .Axes(xlValue).MaximumScale = maxValZ
        .SetElement (msoElementChartTitleAboveChart) ' chart title position
'        .ChartTitle.Delete
        .chartTitle.Text = "Cumulative Returns"
        .chartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
        .SeriesCollection(1).AxisGroup = 2
        .SeriesCollection(1).ChartType = xlColumnClustered
        .Axes(xlValue).TickLabels.NumberFormat = "0.0%"
        .Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0.0%"
        .SeriesCollection(1).Name = "DayReturn"
        .SeriesCollection(2).Name = "ReturnCurve"
        .Axes(xlValue, xlSecondary).MinimumScale = minValY
        .Axes(xlValue, xlSecondary).MaximumScale = maxValY
    End With
    sc(1, 1).Select
End Sub
Sub SortSheetsAlphabetically()
' Sort sheets in alphabetical order
    Dim i As Integer
    Dim areSorted As Boolean
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim sortPoints As Integer
    Dim iterCount As Long
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    areSorted = False
    iterCount = 0
    Do While areSorted = False
        Application.StatusBar = iterCount
        sortPoints = 0
        For i = 1 To Sheets.Count - 1
            Set ws1 = Sheets(i)
            Set ws2 = Sheets(i + 1)
            If LCase(ws2.Name) > LCase(ws1.Name) Then
                sortPoints = sortPoints + 1
            End If
        Next i
        If sortPoints = Sheets.Count - 1 Then
            areSorted = True
            Exit Do
        Else
            For i = 1 To Sheets.Count - 1
                Set ws1 = Sheets(i)
                Set ws2 = Sheets(i + 1)
                If LCase(ws1.Name) > LCase(ws2.Name) Then
                    Sheets(ws2.Index).Move before:=Sheets(ws1.Index)
                End If
            Next i
        End If
        iterCount = iterCount + 1
    Loop
    Sheets(1).Activate
    
    With Application
        .StatusBar = False
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    MsgBox "Done. Iterations: " & iterCount
End Sub

' MODULE: WFA_Functions
Option Explicit

Function AddInFullFileName(ByVal addinFileName As String, _
                           ByVal addinVersion As String) As String
' Return add-in full file name

    AddInFullFileName = addinFileName & "_" & addinVersion & ".xlsm"

End Function

Function CommandBarName(ByVal addInName As String, _
                        ByVal ordNum As Integer) As String
' Return command bar name
    
    CommandBarName = addInName & "-" & ordNum

End Function

Function DirectoryFilesList(ByVal dirPath As String, _
                            ByVal asCollection As Boolean, _
                            ByVal attachDirPath As Boolean)
' Return files list in a directory. Returns Base-1 array.
'
' Parameters:
' dirPath (String): directory to scan for files
' asCollection (Boolean): True - collection, False - array
' attachDirPath (Boolean): True to attach directory path to files names
    
    Dim myList As New Collection
    Dim vaArray As Variant
    Dim i As Integer
    Dim oFile As Object
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFiles As Object
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(dirPath)
    Set oFiles = oFolder.files
    
    If oFiles.Count = 0 Then
        Exit Function
    End If
    
    If asCollection Then
        If attachDirPath Then
            For Each oFile In oFiles
                myList.Add dirPath & "\" & oFile.Name
            Next
        Else
            For Each oFile In oFiles
                myList.Add oFile.Name
            Next
        End If
        Set DirectoryFilesList = myList
    Else
        ReDim vaArray(1 To oFiles.Count)
        i = 1
        
        If attachDirPath Then
            For Each oFile In oFiles
                vaArray(i) = dirPath & "\" & oFile.Name
                i = i + 1
            Next
        Else
            For Each oFile In oFiles
                vaArray(i) = oFile.Name
                i = i + 1
            Next
        End If
        DirectoryFilesList = vaArray
    End If

End Function

Function GetSourceDirectories(ByVal mainWs As Worksheet, _
                              ByVal mainC As Range, _
                              ByVal zeroRow As Integer, _
                              ByVal dataCol As Integer) As Variant
' Return WFA source directories as 1D array. Range to array.

    Dim arr As Variant
    Dim lastSrcDirRow As Integer
    Dim srcDirsRng As Range
    Dim i As Range
    
    lastSrcDirRow = mainC(mainWs.rows.Count, dataCol).End(xlUp).Row
    If lastSrcDirRow > zeroRow Then
        Set srcDirsRng = mainWs.Range(mainC(zeroRow + 1, dataCol), mainC(lastSrcDirRow, dataCol))
        For Each i In srcDirsRng
            i.Value = StringRemoveBackslash(CStr(i.Value))
        Next i
        GetSourceDirectories = RngToArray(srcDirsRng)
    Else
        Set GetSourceDirectories = Nothing
    End If
    
End Function

Function StringRemoveBackslash(ByVal someString As String) As String
' Remove backslash from string end.
    
    If Right(someString, 1) = "\" Then someString = Left(someString, Len(someString) - 1)
    StringRemoveBackslash = someString

End Function

Function RngToCollection(ByVal srcRng As Range) As Collection
' Moves range values to collection.
    
    Dim coll As New Collection
    Dim cell As Range
    
    For Each cell In srcRng
        coll.Add cell.Value
    Next cell
    Set RngToCollection = coll

End Function

Function RngToArray(ByVal srcRng As Range) As Variant
' Convert a range into an array.
    
    Dim arr As Variant
    Dim rRow As Long
    Dim rCol As Integer
    Dim rows As Long
    Dim columns As Integer
    
    rows = srcRng.rows.Count
    columns = srcRng.columns.Count
    
    If columns > 1 Then
        ReDim arr(1 To rows, 1 To columns)
        For rRow = LBound(arr, 1) To UBound(arr, 1)
            For rCol = LBound(arr, 2) To UBound(arr, 2)
                arr(rRow, rCol) = srcRng.item(rRow, rCol)
            Next rCol
        Next rRow
    Else
        ReDim arr(1 To rows)
        For rRow = LBound(arr) To UBound(arr)
            arr(rRow) = srcRng.item(rRow, 1)
        Next rRow
    End If
    RngToArray = arr

End Function

Function GetIsOsWindows(ByVal mainWs As Worksheet, _
                        ByVal mainC As Range, _
                        ByVal windowsFirstRow As Integer, _
                        ByVal windowsFirstCol As Integer) As Variant
' Return 1-based 3 column array of IS and OS weeks with their codes.
' NOT INVERTED.
    
    Dim arr As Variant
    Dim rg As Range
    
    Set rg = mainC(windowsFirstRow, windowsFirstCol).CurrentRegion
    Set rg = rg.Offset(2).Resize(rg.rows.Count - 2)
    arr = rg    ' NOT INVERTED, rows 1 to 20, columns 1 to 3
    GetIsOsWindows = arr

End Function

' ***************************************************************************************************

Function GetScanTable(ByVal srcDirsList As Variant, _
            ByRef printWs As Worksheet, _
            ByRef printCells As Range, _
            ByVal rowOffset As Integer, _
            ByVal colOffset As Integer)
' Returns file lists in the passed folders (dict).
    Dim arr As Variant
    Dim item As Variant
    Dim filesInDir As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim maxFiles As Integer
    Dim currDict As New Dictionary
    Dim currArr As Variant
    Dim insName As String
    Dim dirContents As Variant
    Dim thisInstr As String
    Dim workRng As Range
    Dim clearLastRow As Integer, clearLastCol As Integer
    
    maxFiles = 0
    For i = LBound(srcDirsList) To UBound(srcDirsList)
        filesInDir = DirectoryFilesList(srcDirsList(i), False, False)
        For j = LBound(filesInDir) To UBound(filesInDir)
            insName = InsNameFromReportName(filesInDir(j))
            If Not currDict.Exists(insName) Then
                currDict.Add insName, Nothing
            End If
        Next j
        If UBound(filesInDir) > maxFiles Then
            maxFiles = UBound(filesInDir)
        End If
    Next i
' sort all available instruments alphabetically (bubblesort)
    ' instruments from dictionary to array
    ReDim currArr(1 To currDict.Count)
    i = 1
    For Each item In currDict.Keys
         currArr(i) = item
         i = i + 1
    Next item
    currArr = BubbleSort1DArray(currArr, True)

' create  "INSTRUMENTS BY STRATEGIES" table
' plus index column (currencies) & header (strategies)
    ReDim arr(0 To UBound(currArr), 0 To UBound(srcDirsList))
    ' fill zero-zero cell
    arr(0, 0) = "Scan Table"
    ' fill header (strategies)
    For i = 1 To UBound(arr, 2)
        arr(0, i) = GetBasename(srcDirsList(i))
    Next i
    ' fill index column (currencies)
    For i = 1 To UBound(arr, 1)
        arr(i, 0) = currArr(i)
    Next i
    ' fill paths
    For i = 1 To UBound(arr, 2) ' columns first
        dirContents = DirectoryFilesList(srcDirsList(i), False, False)
        For j = 1 To UBound(arr, 1) ' rows then
            thisInstr = arr(j, 0)
            For k = 1 To UBound(dirContents)
                If InStr(1, dirContents(k), thisInstr, vbTextCompare) > 0 Then
                    arr(j, i) = dirContents(k)
                    Exit For
                End If
            Next k
        Next j
    Next i
'' Print Scan Table
'    ' clear cells for scan table
'    clearLastRow = printCells(printWs.rows.Count, colOffset + 1).End(xlUp).Row
'    clearLastCol = printCells(rowOffset + 1, printWs.columns.Count).End(xlToLeft).Column
'    Set workRng = printWs.Range(printCells(rowOffset + 1, colOffset + 1), printCells(clearLastRow, clearLastCol))
'    workRng.Clear
'    Call Print_2D_Array(arr, False, rowOffset, colOffset, printCells)
'    ' bold index column and header row
'    Set workRng = printWs.Range(printCells(rowOffset + 1, colOffset + 1), printCells(rowOffset + UBound(arr, 1) + 1, colOffset + 1))
'    workRng.Font.Bold = True
'    Set workRng = printWs.Range(printCells(rowOffset + 1, colOffset + 2), printCells(rowOffset + 1, colOffset + UBound(arr, 2) + 1))
'    workRng.Font.Bold = True
    GetScanTable = arr
End Function
Function BubbleSort1DArray(ByVal unsortedArr As Variant, _
            ByVal sortAscending As Boolean) As Variant
' Sorts 1-dimensional array, ascending/alphabetically.
' Base 1 or whatever.
    Dim i As Integer, j As Integer
    Dim tmp As Variant
    If sortAscending Then
        For i = LBound(unsortedArr) To UBound(unsortedArr) - 1
            For j = i + 1 To UBound(unsortedArr)
                If unsortedArr(i) > unsortedArr(j) Then ' ASC
                    tmp = unsortedArr(j)
                    unsortedArr(j) = unsortedArr(i)
                    unsortedArr(i) = tmp
                End If
            Next j
        Next i
    Else
        For i = LBound(unsortedArr) To UBound(unsortedArr) - 1
            For j = i + 1 To UBound(unsortedArr)
                If unsortedArr(i) < unsortedArr(j) Then ' DESC
                    tmp = unsortedArr(j)
                    unsortedArr(j) = unsortedArr(i)
                    unsortedArr(i) = tmp
                End If
            Next j
        Next i
    End If
    BubbleSort1DArray = unsortedArr
End Function
Function InsNameFromReportName(ByVal rptName As String) As String
    InsNameFromReportName = Split(rptName, "-", , vbTextCompare)(1)
End Function
Function GetBasename(ByVal somePath As String) As String
' Return base name from a path string.
    Dim arr As Variant
    somePath = StringRemoveBackslash(somePath)
    arr = Split(somePath, "\", , vbTextCompare)
    GetBasename = arr(UBound(arr))
End Function
Function GetBasenameForTargetWb(ByVal somePath As String) As String
    Dim arr As Variant
    Dim baseName As String
    somePath = StringRemoveBackslash(somePath)
    arr = Split(somePath, "\", , vbTextCompare)
    baseName = arr(UBound(arr))
    If InStr(1, baseName, "(", vbTextCompare) > 0 Then
        arr = Split(baseName, "(", , vbTextCompare)
        baseName = arr(0)
    End If
    GetBasenameForTargetWb = baseName
End Function
Function GetParentDirectory(ByVal somePath As String) As String
    Dim i As Integer
    somePath = StringRemoveBackslash(somePath)
    i = InStrRev(somePath, "\", -1, vbTextCompare)
    GetParentDirectory = Left(somePath, i - 1)
End Function

Function GetNewDirsCount(ByVal scanTable As Variant, _
            ByVal scanMode As Integer) As Integer
' Return count of new directories to be created:
' scanMode 1 - for rows, or instruments
' scanMode 2 - for columns, or strategies
    Select Case scanMode
        Case Is = 1
            ' Loop thru "Scan table" columns (header row)
            ' New folder created for each strategy
            GetNewDirsCount = UBound(scanTable, 2)
        Case Is = 2
            ' Loop thru "Scan table" rows (index column)
            ' New folder created for each instrument
            GetNewDirsCount = UBound(scanTable, 1)
    End Select
End Function
Function GetNewDirPath(ByVal scanTable As Variant, _
            ByVal scanMode As Integer, _
            ByVal dirIndex As Integer, _
            ByVal targetDir As String) As String
' Create path for directory that doesn't exist.
' If directory exists, increment its version in brackets: e.g. (2) -> (3).
    Dim newDirPath As String
    Dim currVersion As String
    Dim versionPath As String
    
    Select Case scanMode
        Case Is = 1
            newDirPath = targetDir & "\" & scanTable(0, dirIndex)
        Case Is = 2
            newDirPath = targetDir & "\" & scanTable(dirIndex, 0)
    End Select
    If Dir(newDirPath, vbDirectory) <> "" Then
        currVersion = 2
        versionPath = newDirPath & "(" & currVersion & ")"
        If Dir(versionPath, vbDirectory) <> "" Then
            Do Until Dir(versionPath, vbDirectory) = ""
                currVersion = currVersion + 1
                versionPath = newDirPath & "(" & currVersion & ")"
            Loop
        End If
        newDirPath = versionPath
    End If
    GetNewDirPath = newDirPath
End Function
Function PathIncrementIndex(ByVal someName As String, _
            ByVal isFile As Boolean) As String
' isFile True means file, False means directory
' someName is passed without ".xlsx"
' increments file/directory index in brackets by 1
' if file/directory exists
    Dim currVersion As Integer
    Dim versionPath As String
    Dim finalName As String
    Dim j As Integer
    Dim temp_s As String
    
    If isFile Then
        versionPath = someName & ".xlsx"
        If Dir(versionPath) <> "" Then
            currVersion = 2
            versionPath = someName & "(" & currVersion & ").xlsx"
            If Dir(versionPath) <> "" Then
                Do Until Dir(versionPath) = ""
                    currVersion = currVersion + 1
                    versionPath = someName & "(" & currVersion & ").xlsx"
                Loop
            End If
        End If
    Else
        If Dir(someName, vbDirectory) <> "" Then
            currVersion = 2
            versionPath = someName & "(" & currVersion & ")"
            If Dir(versionPath, vbDirectory) <> "" Then
                Do Until Dir(versionPath, vbDirectory) = ""
                    currVersion = currVersion + 1
                    versionPath = someName & "(" & currVersion & ")"
                Loop
            End If
        End If
    End If
    PathIncrementIndex = versionPath
End Function

Function GetFilesArr(ByVal srcDirs As Variant, _
            ByVal scanMode As Integer, _
            ByVal scanTable As Variant, _
            ByVal dirIndex As Integer) As Variant
' Return files array for one directory,
' depending on scan mode.
' 1 for scanning within strategy folder,
' 2 for scanning one instrument across many strategy folders.
    Dim scanTableSubset As Variant
    Dim i As Integer, j As Integer
    Dim ubnd As Integer
    
    ' Create subset of file names that includes empty values.
    ReDim scanTableSubset(1 To 1)
    
    j = 0
    For i = 1 To UBound(scanTable, scanMode)
        Select Case scanMode
            Case Is = 1
                ' dirIndex is column in scanTable
                If scanTable(i, dirIndex) <> "" Then
                    j = j + 1
                    ReDim Preserve scanTableSubset(1 To j)
                    scanTableSubset(j) = srcDirs(dirIndex) & "\" & scanTable(i, dirIndex)
                End If
            Case Is = 2
                ' dirIndex is row in scanTable
                If scanTable(dirIndex, i) <> "" Then
                    j = j + 1
                    ReDim Preserve scanTableSubset(1 To j)
                    scanTableSubset(j) = srcDirs(i) & "\" & scanTable(dirIndex, i)
                End If
        End Select
    Next i
    GetFilesArr = scanTableSubset
End Function
Function GetFourDates(ByVal availStart As Long, _
            ByVal availEnd As Long, _
            ByVal weeksIS As Integer, _
            ByVal weeksOS As Integer) As Variant
' function returns (1 to 4, 1 to Rows) array of dates:
' col 1-2: IS from/to, col 3-4: OS from/to
' col 5-6: is & os calendar days as date
' col 7-8: is & os days from 1 to N as Long
' INVERTED: COLUMNS, ROWS
    Dim arr() As Variant
    Dim i As Integer
    Dim j As Long
    Dim rowsCount As Integer
    Dim calendarDays As Integer
    
    ReDim arr(1 To 8, 1 To 1)
    i = 1
    arr(1, i) = availStart
    arr(2, i) = arr(1, i) + 7 * weeksIS - 1
    arr(3, i) = arr(2, i) + 1
    arr(4, i) = arr(3, i) + 7 * weeksOS - 1
    Do While arr(2, i) + 7 * weeksOS < availEnd
        i = i + 1
        ReDim Preserve arr(1 To 8, 1 To i)
        arr(1, i) = arr(1, i - 1) + 7 * weeksOS
        arr(2, i) = arr(1, i) + 7 * weeksIS - 1
        arr(3, i) = arr(2, i) + 1
        arr(4, i) = arr(3, i) + 7 * weeksOS - 1
    Loop
' Adjust last date according to available end date
    If arr(4, UBound(arr, 2)) > availEnd Then
        arr(4, UBound(arr, 2)) = availEnd
    End If

' Add "Calendar days" array for future calculations: R-sq
    For i = LBound(arr, 2) To UBound(arr, 2)
        ' Calendar days IS
        ' 1D array
        arr(5, i) = GenerateCalendarDays(arr(1, i), arr(2, i))
        ' Calendar days OS
        ' 1D array
        arr(6, i) = GenerateCalendarDays(arr(3, i), arr(4, i))
        ' Long, range from 1, IS
        arr(7, i) = GenerateLongDays(UBound(arr(5, i)))
        ' Long, range from 1, OS
        arr(8, i) = GenerateLongDays(UBound(arr(6, i)))
    Next i
    GetFourDates = arr
End Function
Function GenerateLongDays(ByVal ubnd As Long) As Variant
    Dim arr() As Variant
    Dim i As Long
    ReDim arr(1 To ubnd)
    For i = 1 To ubnd
        arr(i) = i
    Next i
    GenerateLongDays = arr
End Function
Function GenerateCalendarDays(ByVal dateStart As Date, _
            ByVal dateEnd As Date) As Variant
    Dim arr As Variant
    Dim cDays As Integer
    Dim i As Integer
    cDays = dateEnd - dateStart + 2
    ReDim arr(1 To cDays)
    arr(1) = dateStart - 1
    For i = 2 To UBound(arr)
        arr(i) = arr(i - 1) + 1
    Next i
    GenerateCalendarDays = arr
End Function
Function Init1DArr(ByVal d1_1 As Integer, _
            ByVal d1_2 As Integer) As Variant
    Dim arr As Variant
    ReDim arr(d1_1 To d1_2)
    Init1DArr = arr
End Function
Function Init2DArr(ByVal d1_1 As Integer, _
            ByVal d1_2 As Integer, _
            ByVal d2_1 As Integer, _
            ByVal d2_2 As Integer) As Variant
    Dim arr As Variant
    ReDim arr(d1_1 To d1_2, d2_1 To d2_2)
    Init2DArr = arr
End Function
Function KPIsDictColumns() As Dictionary
    Dim dict As New Dictionary
    dict.Add "Sharpe Ratio", 3
    dict.Add "R-squared", 5
    dict.Add "Annualized Return", 7
    dict.Add "MDD", 9
    dict.Add "Recovery Factor", 11
    dict.Add "Trades per Month", 13
    dict.Add "Win Ratio", 15
    dict.Add "Avg Winner/Loser", 17
    dict.Add "Avg Trade", 19
    dict.Add "Profit Factor", 21
    Set KPIsDictColumns = dict
End Function
Function GetPermutations(ByVal mainWs As Worksheet, _
            ByVal mainC As Range, _
            ByVal firstRow As Integer, _
            ByVal firstCol As Integer, _
            ByVal stgWs As Worksheet, _
            ByVal stgC As Range, _
            ByVal activeKPIsFRow As Integer, _
            ByVal activeKPIsFCol As Integer, _
            ByVal maxiMinimizing As Variant) As Variant
' Return 2D array of permutations
' not inverted
' ROWS: 0-based, header-1 - KPIs, header-2 - "min/max"
' COLUMNS: 0-based, index column - index of KPI starting with "1"
    Dim arr(1 To 2) As Variant
    Dim colVsKPI As New Dictionary
    Dim activeKPIsDict As New Dictionary
    Dim activeKPIsUpdDict As New Dictionary
    Dim maxRowsCount As Variant
    Dim rg As Range
    Dim i As Integer
    Dim activeKPIsRg As Range
    Dim cell As Range
    Dim activeKPIsLRow As Integer
    Dim activeKPIs As Integer
    Dim thisKPI As String
' on "Hidden Settings" sheet
    activeKPIsLRow = stgC(activeKPIsFRow, activeKPIsFCol).End(xlDown).Row
    Set activeKPIsRg = stgWs.Range(stgC(activeKPIsFRow, activeKPIsFCol), _
            stgC(activeKPIsLRow, activeKPIsFCol))
' fill in dictionary: "Sharpe Ratio", True
    For Each cell In activeKPIsRg
        activeKPIsDict.Add cell.Value, cell.Offset(0, 1).Value
    Next cell
' on "WFA Main" sheet
' create dict of columns vs KPIs: 3, "Sharpe Ratio"
    colVsKPI.Add firstCol, "Sharpe Ratio"
    colVsKPI.Add firstCol + 2, "R-squared"
    colVsKPI.Add firstCol + 4, "Annualized Return"
    colVsKPI.Add firstCol + 6, "MDD"
    colVsKPI.Add firstCol + 8, "Recovery Factor"
    colVsKPI.Add firstCol + 10, "Trades per Month"
    colVsKPI.Add firstCol + 12, "Win Ratio"
    colVsKPI.Add firstCol + 14, "Avg Winner/Loser"
    colVsKPI.Add firstCol + 16, "Avg Trade"
    colVsKPI.Add firstCol + 18, "Profit Factor"
' Select min/max user input
    Set rg = mainC(firstRow, firstCol).CurrentRegion
    Set rg = rg.Offset(1, 0).Resize(rg.rows.Count - 1)
' Get permutations count
    ReDim maxRowsCount(1 To 1)
    activeKPIs = 0
    For i = 1 To rg.columns.Count Step 2
        thisKPI = colVsKPI(i + firstCol - 1)
        ' if range not empty and KPI is active then
        ' expand list of values, for product later
        If rg(1, i) <> "" And activeKPIsDict(thisKPI) = True Then
            activeKPIs = activeKPIs + 1
            ReDim Preserve maxRowsCount(1 To activeKPIs)
            maxRowsCount(activeKPIs) = rg(0, i).End(xlDown).Row - firstRow
            ' Update dictionary of Active KPI vs 2D array of its min/max values
            activeKPIsUpdDict.Add thisKPI, GetMinMaxValsForKPI(mainWs, mainC, _
                    firstRow + 1, i + firstCol - 1)
        End If
    Next i
' DEBUG
'    For i = 0 To activeKPIsUpdDict.Count - 1
'        Debug.Print activeKPIsUpdDict.Keys(i)
'        For j = LBound(activeKPIsUpdDict.Items(i), 1) To UBound(activeKPIsUpdDict.Items(i), 1)
'            For k = LBound(activeKPIsUpdDict.Items(i), 2) To UBound(activeKPIsUpdDict.Items(i), 2)
'                Debug.Print activeKPIsUpdDict.Items(i)(j, k)
'            Next k
'        Next j
'    Next i

    arr(1) = KpiRangesToArray(activeKPIsUpdDict, maxiMinimizing)
'Debug
'Call Print_2D_Array(arr(1), True, 25, 0, mainC)

' PERMUTATIONS COUNT = WorksheetFunction.Product(maxRowsCount)
    arr(2) = GetPermutationsTable(WorksheetFunction.Product(maxRowsCount), _
            activeKPIs, _
            activeKPIsUpdDict, _
            maxRowsCount)
    GetPermutations = arr
End Function
Function KpiRangesToArray(ByVal origDict As Dictionary, _
            ByVal maxiMinimizing As Variant) As Variant
' INVERTES: columns, rows
    Dim arr As Variant
    Dim tmpArr As Variant
    Dim i As Integer
    Dim j As Integer
    Dim kpiName As String
    Dim arrRow As Integer
    Dim arrCol As Integer
    ReDim arr(1 To origDict.Count * 2, 0 To 1)
    For i = 0 To origDict.Count - 1
        arrCol = (i + 1) * 2 - 1
        kpiName = origDict.Keys(i)
        arr(arrCol, 0) = kpiName
        If kpiName = maxiMinimizing(2) Then
            arr(arrCol + 1, 0) = maxiMinimizing(1)
        End If
        arr(arrCol, 1) = "min"
        arr(arrCol + 1, 1) = "max"
        tmpArr = origDict(kpiName)
        If UBound(arr, 2) - 1 < UBound(tmpArr, 1) Then
            ReDim Preserve arr(1 To UBound(arr, 1), 0 To UBound(tmpArr, 1) + 1)
        End If
        For j = LBound(tmpArr, 1) To UBound(tmpArr, 1)
            arrRow = j + 1
            arr(arrCol, arrRow) = tmpArr(j, 1)
            arr(arrCol + 1, arrRow) = tmpArr(j, 2)
        Next j
    Next i
    KpiRangesToArray = arr
End Function

Function GetPermutationsTable(ByVal permCount As Integer, _
            ByVal activeKPIsCount As Integer, _
            ByVal sourceDict As Dictionary, _
            ByVal maxRowsCount As Variant) As Variant
' Return 2D array of all permutations, not inverted, columns by rows.
    Dim arr As Variant
    Dim pointers As Variant
    Dim i As Integer
    Dim fillRow As Integer
    Dim pointRow As Integer, pointCol As Integer
' row 0: KPI name >> header 1
' row 1: min, max, min, max, ... >> header 2
' row 2: value min, value max
' column 0: Index - starting from 2nd row, value 1
    
'    Dim t1 As Variant
'    t1 = sourceDict("Sharpe Ratio")
''Debug.Print sourceDict("Sharpe Ratio")

    ReDim arr(0 To permCount + 1, 0 To activeKPIsCount * 2)
' Fill header 1
    arr(0, 0) = "KPI"
    For i = 0 To sourceDict.Count - 1
        arr(0, i * 2 + 1) = sourceDict.Keys(i)
    Next i
' Fill header 2: min, max
    arr(1, 0) = "index"
    For i = 1 To UBound(arr, 2) - 1 Step 2
        arr(1, i) = "min"
        arr(1, i + 1) = "max"
    Next i
' Fill indices
    For i = 2 To UBound(arr, 1)
        arr(i, 0) = i - 1
    Next i
' Fill the "meat"
    ' fill pointers' starting values
    ReDim pointers(1 To UBound(maxRowsCount))
    For i = LBound(pointers) To UBound(pointers)
        pointers(i) = 1
    Next i
    pointers(1) = 0
    ' loop, using pointers
    fillRow = 1
    Do Until SeriesAreEqual(pointers, maxRowsCount)
        If pointers(1) = maxRowsCount(1) Then
            pointers = RecursivelyUpdate(pointers, maxRowsCount, 1)
        Else
            pointers(1) = pointers(1) + 1
        End If
        ' DEBUG Call pointersDebug(pointers)
        ' fill arr here
        fillRow = fillRow + 1
        For pointCol = LBound(pointers) To UBound(pointers)
            pointRow = pointers(pointCol)
            arr(fillRow, pointCol * 2 - 1) = sourceDict(sourceDict.Keys(pointCol - 1))(pointRow, 1)
            arr(fillRow, pointCol * 2) = sourceDict(sourceDict.Keys(pointCol - 1))(pointRow, 2)
        Next pointCol
    Loop
    ' debug - print 2d
'    Call Print_2D_Array(arr, False, 20, 1, Cells)
    GetPermutationsTable = arr
End Function
Function RecursivelyUpdate(ByRef currentPointers As Variant, _
            ByVal referencePointers As Variant, _
            ByVal thisCol As Integer) As Variant
' return 1d array (series) with updated pointers
    currentPointers(thisCol) = 1
    If currentPointers(thisCol + 1) < referencePointers(thisCol + 1) Then   ' max reached
        currentPointers(thisCol + 1) = currentPointers(thisCol + 1) + 1
    Else
        currentPointers = RecursivelyUpdate(currentPointers, referencePointers, thisCol + 1)
    End If
    RecursivelyUpdate = currentPointers
End Function
Function SeriesAreEqual(ByVal arr1 As Variant, _
            ByVal arr2 As Variant) As Boolean
    Dim i As Integer
    Dim counter As Integer
    Dim limitUp As Integer
    
    If UBound(arr1) = UBound(arr2) Then
        limitUp = UBound(arr1)
        counter = 0
        For i = LBound(arr1) To UBound(arr1)
            If arr1(i) = arr2(i) Then
                counter = counter + 1
            Else
                SeriesAreEqual = False
                Exit For
            End If
        Next i
        If counter = limitUp Then
            SeriesAreEqual = True
        Else
            SeriesAreEqual = False
        End If
    Else
        SeriesAreEqual = False
    End If
End Function
Function GetMinMaxValsForKPI(ByVal mainWs As Worksheet, _
            ByVal mainC As Range, _
            ByVal firstRow As Integer, _
            ByVal firstCol As Integer) As Variant
' Return 2D array of min & max KPI values, 1 based, rows by columns.
' Not inverted
    Dim rg As Range
    Dim arr As Variant
    Dim lastRow As Integer
    lastRow = mainC(firstRow - 1, firstCol).End(xlDown).Row
    Set rg = mainWs.Range(mainC(firstRow, firstCol), mainC(lastRow, firstCol + 1))
    arr = rg
    GetMinMaxValsForKPI = arr
End Function
Function GetTargetWBSaveName(ByVal targetDir As String, _
            ByVal windowCode As String, _
            ByVal stratOrInstrumentName As String, _
            ByVal dateBegin As Date, _
            ByVal dateEnd As Date) As String
    Dim dtBeginString As String
    Dim dtEndString As String
    dtBeginString = GetDateAsString(dateBegin)
    dtEndString = GetDateAsString(dateEnd)
    GetTargetWBSaveName = targetDir & "\wfa-" & windowCode & "-" & stratOrInstrumentName _
        & "-" & dtBeginString & "-" & dtEndString & ".xlsx"
End Function
Function GetDateAsString(ByVal someDate As Date) As String
    Dim sYear As String, sMonth As String, sDay As String
    sYear = Right(CStr(Year(someDate)), 2)
    sMonth = CStr(Month(someDate))
    If Len(sMonth) = 1 Then
        sMonth = "0" & sMonth
    End If
    sDay = CStr(Day(someDate))
    If Len(sDay) = 1 Then
        sDay = "0" & sDay
    End If
    GetDateAsString = sYear & sMonth & sDay
End Function
Function GetTradeListFromSheetAF(ByVal ws As Worksheet, _
            ByVal wsC As Range, _
            ByVal dateFrom As Long, _
            ByVal dateTo As Long, _
            ByVal wbName As String) As Variant
' Return 2D array of trades
' INVERTED
' Include header row
' Columns:
    ' Open date
    ' Close date
    ' Source
    ' Return
' Source = filename & postfix "_5" where 5 is sheet number for quick access
    Const insertRow As Integer = 5
    Const insertCol As Integer = 15
    Dim arr As Variant
    Dim dbRg As Range, dbSmall As Range, critRg As Range, cell As Range
    Dim lastRow As Long, i As Long
    Dim srcStr As String
    
    ReDim arr(1 To 4, 0 To 0)
    If wsC(11, 2) = 0 Then
        GetTradeListFromSheetAF = arr
        Exit Function
    End If
    lastRow = wsC(ws.rows.Count, 9).End(xlUp).Row
    Set dbRg = ws.Range(wsC(1, 3), wsC(lastRow, 13))
    
    ' criteria range
    wsC(1, 15) = "Open date"
    wsC(1, 16) = "Close date"
    wsC(2, 15) = ">=" & dateFrom
    wsC(2, 16) = "<" & dateTo
    Set critRg = wsC(1, 15).CurrentRegion
    
    ' Advanced filter
    dbRg.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=critRg, _
        CopyToRange:=wsC(insertRow, insertCol), Unique:=False
    
    ' Select columns: open date, close date, source, return
    Set dbRg = wsC(insertRow, insertCol).CurrentRegion
    If dbRg.rows.Count = 1 Then
        GetTradeListFromSheetAF = arr
        Exit Function
    End If
    Set dbSmall = dbRg.Offset(1, 6).Resize(dbRg.rows.Count - 1, 1)
    ReDim arr(1 To 4, 0 To dbSmall.rows.Count)
    srcStr = Left(wbName, Len(wbName) - 5) & "_" & ws.Name
    i = 0
    arr(1, i) = "Open date"
    arr(2, i) = "Close date"
    arr(3, i) = "Source"
    arr(4, i) = "Return"
    For Each cell In dbSmall
        i = i + 1
        arr(1, i) = cell
        arr(2, i) = cell.Offset(0, 1)
        arr(3, i) = srcStr
        arr(4, i) = cell.Offset(0, 4)
    Next cell
    critRg.Clear
    dbRg.Clear
    GetTradeListFromSheetAF = arr
End Function
Function GetTradeListFromSheet(ByVal ws As Worksheet, _
            ByVal date_0 As Date, _
            ByVal date_1 As Date, _
            ByVal book_name As String)
' Return 2D array of trades
' INVERTED
' Include header row
' Columns:
    ' Open date
    ' Close date
    ' Source
    ' Return
' Source = filename & postfix "_5" where 5 is sheet number for quick access
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
        GetTradeListFromSheet = result_arr
        Exit Function
    End If
    last_row = wsC(1, 9).End(xlDown).Row
    i = 2
    result_arr(1, 0) = "Open date"
    result_arr(2, 0) = "Close date"
    result_arr(3, 0) = "Source"
    result_arr(4, 0) = "Return"
    Do While Int(wsC(i, 9)) < date_1 And i <= last_row
        If Int(wsC(i, 9)) >= date_0 And Int(wsC(i, 10)) < date_1 Then
            ubnd = UBound(result_arr, 2) + 1
            ReDim Preserve result_arr(1 To 4, 0 To ubnd)
            result_arr(1, ubnd) = wsC(i, 9)
            result_arr(2, ubnd) = wsC(i, 10)
            result_arr(3, ubnd) = comment_str
            result_arr(4, ubnd) = wsC(i, 13)
        End If
        i = i + 1
    Loop
    GetTradeListFromSheet = result_arr
End Function
Function LoadReportToRAM(ByVal ws As Worksheet, _
            ByVal srcString As String) As Variant
' Function loads html report from sheet to RAM
' Returns (1 To 4, 0 To trades_count) array
    Dim arr() As Variant
    Dim lastRow As Long
    Dim wsC As Range
    Dim i As Long, j As Long
    
    Set wsC = ws.Cells
    lastRow = wsC(ws.rows.Count, 4).End(xlUp).Row
    ReDim arr(1 To 4, 0 To 0)
    arr(1, 0) = "Open date"
    arr(2, 0) = "Close date"
    arr(3, 0) = "Source"
    arr(4, 0) = "Return"
    If lastRow = 1 Then
        LoadReportToRAM = arr
        Exit Function
    End If
    
    ReDim Preserve arr(1 To 4, 0 To lastRow - 1)
    For i = 2 To lastRow
        j = i - 1
        arr(1, j) = wsC(i, 9)   ' open date
        arr(2, j) = wsC(i, 10)  ' close date
        arr(3, j) = srcString   ' source
        arr(4, j) = wsC(i, 13)  ' return
    Next i
    LoadReportToRAM = arr
End Function
Function LoadRngToRAM(ByVal rg As Range, _
            ByVal isInverted As Boolean) As Variant
    Dim arr As Variant
    Dim rowsCount As Long
    Dim columnsCount As Long
    Dim thisRow As Long, thisCol As Integer
    Dim rgRow As Long, rgCol As Long
    rowsCount = rg.rows.Count
    columnsCount = rg.columns.Count
    If isInverted Then
        ReDim arr(1 To columnsCount, 0 To rowsCount - 1) ' "-1" because 0-based
        For rgRow = 1 To rowsCount
            For rgCol = 1 To columnsCount
                thisRow = rgRow - 1
                arr(rgCol, thisRow) = rg(rgRow, rgCol)
            Next rgCol
        Next rgRow
    Else
        ReDim arr(0 To rowsCount - 1, 1 To columnsCount) ' "-1" because 0-based
        For rgRow = 1 To rowsCount
            For rgCol = 1 To columnsCount
                thisRow = rgRow - 1
'                arr(thisRow, thisCol) = rg(rgRow, rgCol)
                arr(thisRow, rgCol) = rg(rgRow, rgCol)
            Next rgCol
        Next rgRow
    End If
    LoadRngToRAM = arr
End Function
Function ApplyDateFilter(ByVal reportRam, _
            ByVal startDate As Long, _
            ByVal endDate As Long) As Variant
' Returns 1 to 4, 0 to N Trade List
' INVERTED
' 0th = header row
    Dim arr As Variant
    Dim i As Long
    Dim ubnd As Long
    
    ReDim arr(1 To 4, 0 To 0)
    arr(1, 0) = "Open date"
    arr(2, 0) = "Close date"
    arr(3, 0) = "Source"
    arr(4, 0) = "Return"
    If UBound(reportRam, 2) = 0 Then
        ApplyDateFilter = arr
        Exit Function
    End If
    i = 1
    Do While reportRam(1, i) < endDate
                                        ' And i <= UBound(reportRam, 2)
        If reportRam(1, i) >= startDate And reportRam(2, i) < endDate Then
            ubnd = UBound(arr, 2) + 1
            ReDim Preserve arr(1 To 4, 0 To ubnd)
            arr(1, ubnd) = reportRam(1, i)
            arr(2, ubnd) = reportRam(2, i)
            arr(3, ubnd) = reportRam(3, i)
            arr(4, ubnd) = reportRam(4, i)
        End If
        i = i + 1
        If i > UBound(reportRam, 2) Then
            Exit Do
        End If
    Loop
    ApplyDateFilter = arr
End Function
Function GetMaxiMinimize(ByVal stgSheet As Worksheet, _
            ByVal stgCells As Range, _
            ByVal kpiFRow As Integer, _
            ByVal kpiFCol As Integer) As Variant
    Dim arr(1 To 2) As Variant
    Dim lastRow As Integer
    Dim rg As Range
    Dim cell As Range
    lastRow = stgCells(kpiFRow, kpiFCol).End(xlDown).Row
    Set rg = stgSheet.Range(stgCells(kpiFRow, kpiFCol), _
            stgCells(lastRow, kpiFCol))
    arr(1) = "none"
    arr(2) = "none"
    For Each cell In rg
        If cell.Offset(0, 1).Value = True Then
            If cell.Offset(0, 2).Value = True Then
                arr(2) = cell.Value
                arr(1) = "maximize"
                Exit For
            End If
            If cell.Offset(0, 3).Value = True Then
                arr(2) = cell.Value
                arr(1) = "minimize"
                Exit For
            End If
        End If
    Next cell
    GetMaxiMinimize = arr
End Function
Function GetDailyEquityFromTradeSet(ByVal tradeSet As Variant, _
            ByVal dateStart As Date, _
            ByVal dateEnd As Date) As Variant
' Not Inverted
' arr(1 to days, 1 to 2)
' column 1 - calendar days
' column 2 - daily equity
    Dim arr As Variant
    Dim calendarDays As Long
    Dim i As Long
    Dim j As Long
' day-by-day equity (including weekends)
    calendarDays = dateEnd - dateStart + 2
    ReDim arr(1 To calendarDays, 1 To 2)
    arr(1, 1) = dateStart - 1
    arr(1, 2) = 1
    j = 1
    For i = 2 To UBound(arr, 1)
        arr(i, 1) = arr(i - 1, 1) + 1
        arr(i, 2) = arr(i - 1, 2)
        If CLng(arr(i, 1)) = CLng(tradeSet(2, j)) Then
            Do While CLng(arr(i, 1)) = CLng(tradeSet(2, j))
                arr(i, 2) = arr(i, 2) * (1 + tradeSet(4, j))
                If j < UBound(tradeSet, 2) Then
                    j = j + 1
                ElseIf j = UBound(tradeSet, 2) Then
                    Exit Do
                End If
            Loop
        End If
    Next i
    GetDailyEquityFromTradeSet = arr
End Function

Function CalcKPIs(ByVal tradeSet As Variant, _
            ByVal dateStart As Date, _
            ByVal dateEnd As Date, _
            ByVal calDays As Variant, _
            ByVal calDaysLong As Variant) As Dictionary
    Dim resultDict As Dictionary
    Dim i As Long, j As Long
    Dim tradeEq As Variant
    Dim hwmArr As Variant
    Dim ddArr As Variant
    Dim tradeReturnOnly As Variant
    Dim servDict As Dictionary
    Dim dailyEq As Variant
    
    Set resultDict = InitKPIsDict
    If UBound(tradeSet, 2) = 0 Then
        Set CalcKPIs = resultDict
        Exit Function
    End If
' day-by-day equity (including weekends)
    ReDim dailyEq(1 To UBound(calDays))
    dailyEq(1) = 1
    j = 1
    For i = 2 To UBound(dailyEq)
        dailyEq(i) = dailyEq(i - 1)
        If CLng(calDays(i)) = CLng(tradeSet(2, j)) Then
            Do While CLng(calDays(i)) = CLng(tradeSet(2, j)) ' And j <= UBound(trades_arr, 2)
                dailyEq(i) = dailyEq(i) * (1 + tradeSet(4, j))
                If j < UBound(tradeSet, 2) Then
                    j = j + 1
                ElseIf j = UBound(tradeSet, 2) Then
                    Exit Do
                End If
            Loop
        End If
    Next i

' trade return only, trade-by-trade equity, hwm, dd
    ReDim tradeReturnOnly(1 To UBound(tradeSet, 2))
    ReDim tradeEq(0 To UBound(tradeSet, 2))
    ReDim hwmArr(0 To UBound(tradeSet, 2))
    ReDim ddArr(0 To UBound(tradeSet, 2))
    tradeEq(0) = 1
    hwmArr(0) = 1
    For i = LBound(tradeReturnOnly) To UBound(tradeReturnOnly)
        tradeReturnOnly(i) = tradeSet(4, i)
        tradeEq(i) = tradeEq(i - 1) * (1 + tradeReturnOnly(i))
        hwmArr(i) = WorksheetFunction.Max(hwmArr(i - 1), tradeEq(i))
        ddArr(i) = (hwmArr(i) - tradeEq(i)) / hwmArr(i)
    Next i
' KPIs
'    Call CreateServiceDict(servDict, tradeReturnOnly)
    Set servDict = CreateServiceDict(tradeReturnOnly)
    resultDict("R-squared") = WorksheetFunction.RSq(calDaysLong, dailyEq)
    resultDict("Annualized Return") = dailyEq(UBound(dailyEq)) ^ (365 / (UBound(calDays) - 1)) - 1
    resultDict("Sharpe Ratio") = CalcKPIs_SharpeRatio(tradeReturnOnly, resultDict("Annualized Return"))
    resultDict("MDD") = WorksheetFunction.Max(ddArr)
    resultDict("Recovery Factor") = CalcKPIs_RecoveryFactor(resultDict("Annualized Return"), resultDict("MDD"))
    resultDict("Trades per Month") = UBound(tradeSet, 2) / ((dateEnd - dateStart + 1) * 12 / 365)
    resultDict("Win Ratio") = servDict("Winners Count") / UBound(tradeReturnOnly)
    resultDict("Avg Winner/Loser") = CalcKPIs_AvgWinnerToLoser(servDict)
    resultDict("Avg Trade") = WorksheetFunction.Average(tradeReturnOnly)
    resultDict("Profit Factor") = CalcKPIs_ProfitFactor(servDict("Winners Sum"), servDict("Losers Sum"))
'    ' debug - choose clean sheet
'    Cells(1, 1) = dateStart
'    Cells(2, 1) = dateEnd
'    Call Print_2D_Array(tradeSet, True, 0, 1, Cells)
'    Call Print_1D_Array(tradeEq, 5, Cells)
'    Call Print_1D_Array(calDays, 6, Cells)
'    Call Print_1D_Array(dailyEq, 7, Cells)
'
'    For i = 0 To dict.Count - 1
'        Cells(i + 1, 9) = dict.Keys(i)
'        Cells(i + 1, 10) = dict(dict.Keys(i))
'    Next i
'' DEBUG print all dict to immediate window
'    For i = 0 To resultDict.Count - 1
'        Debug.Print resultDict.Keys(i)
'        Debug.Print resultDict(resultDict.Keys(i))
'    Next i
' end DEBUG
'    Debug.Print resultDict("MDD")
    Set CalcKPIs = resultDict
End Function
Function CreateServiceDict(ByVal tradeReturns As Variant) As Dictionary
    Dim i As Long
    Dim winCount As Long
    Dim losCount As Long
    Dim sumWinners As Double
    Dim sumLosers As Double
    Dim servDict As New Dictionary
    
    For i = LBound(tradeReturns) To UBound(tradeReturns)
        If tradeReturns(i) > 0 Then
            winCount = winCount + 1
            sumWinners = sumWinners + tradeReturns(i)
        Else
            sumLosers = sumLosers + tradeReturns(i)
        End If
    Next i
    losCount = UBound(tradeReturns) - winCount
    servDict.Add "Winners Count", winCount
    servDict.Add "Losers Count", losCount
    servDict.Add "Winners Sum", sumWinners
    servDict.Add "Losers Sum", sumLosers
    Set CreateServiceDict = servDict
End Function
Function CreateThisCritDict(ByVal permTable As Variant, _
            ByVal iPermutation As Integer) As Dictionary
    Dim iCol As Integer

    Dim minMaxArr As Variant
    Dim critDict As New Dictionary
    For iCol = 1 To UBound(permTable, 2) Step 2
        minMaxArr = Init1DArr(1, 2)
        minMaxArr(1) = permTable(iPermutation, iCol)
        minMaxArr(2) = permTable(iPermutation, iCol + 1)
        critDict.Add permTable(0, iCol), minMaxArr
    Next iCol
'' debug
'    Dim iKey As Integer
'    For iKey = 0 To critDict.Count - 1
'        Debug.Print "KPI = " & critDict.Keys(iKey)
'        Debug.Print "min = " & critDict(critDict.Keys(iKey))(1)
'        Debug.Print "max = " & critDict(critDict.Keys(iKey))(2)
'    Next iKey
'' end debug
    Set CreateThisCritDict = critDict
End Function
Function CalcKPIs_ProfitFactor(ByVal winnersSum As Double, _
            ByVal losersSum As Double) As Double
    If losersSum = 0 Then
        CalcKPIs_ProfitFactor = 999
    Else
        CalcKPIs_ProfitFactor = Abs(winnersSum / losersSum)
    End If
End Function
Function CalcKPIs_RecoveryFactor(ByVal annReturn As Double, _
            ByVal maxDD As Double) As Double
    If maxDD = 0 Then
        CalcKPIs_RecoveryFactor = 999
    Else
        CalcKPIs_RecoveryFactor = annReturn / maxDD
    End If
End Function
Function CalcKPIs_AvgWinnerToLoser(ByVal servDict As Dictionary) As Double
    Dim result As Double
    If servDict("Winners Count") = 0 Then
        CalcKPIs_AvgWinnerToLoser = -999
    ElseIf servDict("Losers Count") = 0 _
            Or servDict("Losers Sum") = 0 Then
        CalcKPIs_AvgWinnerToLoser = 999
    Else
        CalcKPIs_AvgWinnerToLoser = Abs((servDict("Winners Sum") / _
            servDict("Winners Count")) / (servDict("Losers Sum") / _
            servDict("Losers Count")))
    End If
End Function
Function CalcKPIs_SharpeRatio(ByVal tradeReturnOnly As Variant, _
            ByVal annReturn As Double) As Variant
    Dim annStd As Variant
    If UBound(tradeReturnOnly) = 1 Then
        annStd = "N/A"
    Else
        annStd = WorksheetFunction.StDev(tradeReturnOnly) * Sqr(365)
    End If
    If annStd = "N/A" Then
        CalcKPIs_SharpeRatio = "N/A"
    Else
        CalcKPIs_SharpeRatio = annReturn / annStd
    End If
End Function
Function PassesCriteria(ByVal kpisDict As Dictionary, _
            ByVal critDict As Dictionary) As Boolean
' Return True if Trade List passes criteria from "Criteria Dictionary"
    Dim i As Integer
    Dim kpiName As String
    Dim kpiMin As Double
    Dim kpiMax As Double
    Dim passPoints As Integer
    passPoints = 0
    For i = 0 To critDict.Count - 1
        kpiName = critDict.Keys(i)
        kpiMin = critDict(kpiName)(1)
        kpiMax = critDict(kpiName)(2)
        If kpisDict(kpiName) >= kpiMin And kpisDict(kpiName) < kpiMax Then
            passPoints = passPoints + 1
        End If
    Next i
    If passPoints = critDict.Count Then
        PassesCriteria = True
    Else
        PassesCriteria = False
    End If
End Function
Function InitializeResultArray(ByVal permArr As Variant, _
            ByVal datesISOS As Variant, _
            ByVal maximization As String) As Variant
' INIT RESULT ARRAY
' A(1 to permutations count)
' permCount = param("Permutations")(UBound(param("Permutations"), 1), 0)
    Dim A As Variant
    Dim permID As Integer
    Dim dateSlotID As Integer
    Dim sampleID As Integer
    
    A = Init1DArr(1, permArr(UBound(permArr, 1), 0)) ' WHERE
            ' 1, 2, ... , N - are permutations
    For permID = LBound(A) To UBound(A)
        
        A(permID) = Init1DArr(0, UBound(datesISOS, 2)) ' WHERE
                ' 1, 2, ... , N - are date slots (IS+OS)
                ' 0 - is forward compiled
        ' init forward compiled arr
        A(permID)(0) = Init1DArr(1, 2) ' WHERE
                ' 1 OS United tradeList
                ' 2 OS United KPIs

        ' init forward compiled tradeList - INVERTED
        A(permID)(0)(1) = InitEmptyTradeList
'        A(permID)(0)(1) = Init2DArr(1, 4, 0, 0) ' WHERE 0th row is Header
'        A(permID)(0)(1)(1, 0) = "Open date" ' fill header row
'        A(permID)(0)(1)(2, 0) = "Close date"
'        A(permID)(0)(1)(3, 0) = "Source"
'        A(permID)(0)(1)(4, 0) = "Return"
        
        ' init empty KPIs dict for forward compiled (or try Nothing instead of empty dict)
        Set A(permID)(0)(2) = InitKPIsDict
        
        For dateSlotID = 1 To UBound(A(permID))
            
            ' If maximization is ON
            If maximization = "none" Then
                A(permID)(dateSlotID) = Init1DArr(1, 2) ' WHERE
                        ' 1 is IS array
                        ' 2 is OS array
            Else
                A(permID)(dateSlotID) = Init1DArr(0, 2) ' WHERE
                        ' 0 is Candidates array (1 to 2, 1 to N)
                        ' 1 is IS array
                        ' 2 is OS array
                
                ' init candidates array
                A(permID)(dateSlotID)(0) = InitCandidatesArray
            End If
            
            ' init winners array
            For sampleID = 1 To 2 ' IS winners, OS winners (both - tradeLists & KPIs)
                A(permID)(dateSlotID)(sampleID) = Init1DArr(1, 2) ' WHERE
                        ' 1 is tradeList
                        ' 2 is KPIs
                
                A(permID)(dateSlotID)(sampleID)(1) = InitEmptyTradeList
'                A(permID)(dateSlotID)(sampleID)(1) = Init2DArr(1, 4, 0, 0) ' init arr
'                            ' INVERTED
'                            ' for winners tradeList: dtOpen, dtClose, src, return
'                A(permID)(dateSlotID)(sampleID)(1)(1, 0) = "Open date"
'                A(permID)(dateSlotID)(sampleID)(1)(2, 0) = "Close date"
'                A(permID)(dateSlotID)(sampleID)(1)(3, 0) = "Source"
'                A(permID)(dateSlotID)(sampleID)(1)(4, 0) = "Return"
                
                ' init empty KPIs dict
                Set A(permID)(dateSlotID)(sampleID)(2) = InitKPIsDict
            Next sampleID
        Next dateSlotID
    Next permID
    InitializeResultArray = A
End Function
Function InitCandidatesArray() As Variant
    Dim arr As Variant
    arr = Init2DArr(1, 3, 0, 0) ' WHERE
            ' INVERTED
            ' column 1 is IS Trade Lists
            ' column 2 is IS KPIs dictionaries
            ' column 3 is OS Trade Lists
            ' use ReDim Preserve when adding new rows
    arr(1, 0) = "IS trade lists"
    arr(2, 0) = "IS KPIs dictionaries"
    arr(3, 0) = "OS trade lists"
    InitCandidatesArray = arr
End Function
Function InitEmptyTradeList() As Variant
' INVERTED
    Dim arr As Variant
    arr = Init2DArr(1, 4, 0, 0) ' init arr
    arr(1, 0) = "Open date"
    arr(2, 0) = "Close date"
    arr(3, 0) = "Source"
    arr(4, 0) = "Return"
    InitEmptyTradeList = arr
End Function
Function InitKPIsDict()
    Dim dict As New Dictionary
    Dim setVal As Variant
'    setVal = 0
    setVal = "N/A"
'    Set setVal = Nothing
    With dict
        .Add "Sharpe Ratio", setVal
        .Add "R-squared", setVal
        .Add "Annualized Return", setVal
        .Add "MDD", setVal
        .Add "Recovery Factor", setVal
        .Add "Trades per Month", setVal
        .Add "Win Ratio", setVal
        .Add "Avg Winner/Loser", setVal
        .Add "Avg Trade", setVal
        .Add "Profit Factor", setVal
    End With
    Set InitKPIsDict = dict
End Function
Function ExtendTradeList(ByVal originalList As Variant, _
            ByVal newList As Variant) As Variant
' Function appends trade list with new trades
' originalList (1 To 4, 0 to trades_count)
    Dim extendedList As Variant
    Dim r As Long
    Dim rowInExtended As Long
    Dim origUbnd As Long
    Dim c As Integer

    origUbnd = UBound(originalList, 2)
    extendedList = originalList
    ReDim Preserve extendedList(1 To 4, 0 To origUbnd + UBound(newList, 2))
    For r = 1 To UBound(newList, 2)
        rowInExtended = origUbnd + r
        For c = LBound(newList, 1) To UBound(newList, 1)
            extendedList(c, rowInExtended) = newList(c, r)
        Next c
    Next r
    ExtendTradeList = extendedList
End Function
Function AppendCandidate(ByVal originalSet As Variant, _
            ByVal isTradeList As Variant, _
            ByVal isKPIs As Dictionary, _
            ByVal osTradeList As Variant) As Variant
' Appends 3 items to original set (inverted)
' into a new row
' Cols = 1 To 3: 1) IS trade lists, 2) IS KPIs dicts, 3) OS trade lists
' Rows = 0 to N
' INVERTED
    Dim arr As Variant
    Dim newUbnd As Integer
    arr = originalSet
    newUbnd = UBound(originalSet, 2) + 1
    ReDim Preserve arr(1 To 3, 0 To newUbnd)
    arr(1, newUbnd) = isTradeList
    Set arr(2, newUbnd) = isKPIs
    arr(3, newUbnd) = osTradeList
    AppendCandidate = arr
'    Debug.Print arr(2, 1)("Sharpe Ratio")
End Function
Function BubbleSortRange(ByRef sortWorkSheet As Worksheet, _
            ByRef sortCells As Range, _
            ByVal sortColID As Integer, _
            ByVal sortAscending As Boolean) As Variant
    Dim arr As Variant
    Dim sortRg As Range
    Set sortRg = sortCells(1, 1).CurrentRegion
    If sortAscending Then
        sortRg.Sort Key1:=sortCells(1, sortColID), _
            Order1:=xlAscending, _
            Header:=xlYes
    Else
        sortRg.Sort Key1:=sortCells(1, sortColID), _
            Order1:=xlDescending, _
            Header:=xlYes
    End If
    arr = sortRg
    sortRg.Clear
End Function
Function BubbleSort2DArray(ByVal origArr As Variant, _
            ByVal isInverted As Boolean, _
            ByVal hasHeaderRow As Boolean, _
            ByVal sortAscending As Boolean, _
            ByVal sortColID As Integer, _
            ByRef sortWorkSheet As Worksheet, _
            ByRef sortCells As Range) As Variant
    Dim arr As Variant
    Dim sortRg As Range
    Dim colDimension As Integer
    Dim rowDimension As Integer
    Dim startPosition As Long
    Dim i As Long, j As Long
    Dim k As Integer
    Dim tmp As Variant

    arr = origArr
    If isInverted Then
        colDimension = 1
        rowDimension = 2
    Else
        rowDimension = 1
        colDimension = 2
    End If
    If hasHeaderRow Then
        startPosition = LBound(arr, rowDimension) + 1
    Else
        startPosition = LBound(arr, rowDimension)
    End If
    If UBound(arr, rowDimension) > startPosition Then
'        If UBound(arr, rowDimension) > 10 Then
            Call Print_2D_Array(origArr, isInverted, 0, 0, sortCells)
            Set sortRg = sortCells(1, 1).CurrentRegion
            If sortAscending Then
                sortRg.Sort Key1:=sortCells(1, sortColID), _
                    Order1:=xlAscending, _
                    Header:=xlYes
            Else
                sortRg.Sort Key1:=sortCells(1, sortColID), _
                    Order1:=xlDescending, _
                    Header:=xlYes
            End If
            arr = LoadRngToRAM(sortRg, isInverted)
''            debug
'            Call Print_2D_Array(arr, isInverted, 0, 4, sortCells)
            sortRg.Clear
            Set sortRg = Nothing
'        Else
'            ' HARD CORE bubble sort for small arrays
'            If sortAscending Then
'                ' Sort Ascending
'                If rowDimension = 1 Then
'                    ' Not inverted array
'                    For i = startPosition To UBound(arr, rowDimension) - 1
'                        For j = i + 1 To UBound(arr, rowDimension)
'                            If arr(i, sortColId) > arr(j, sortColId) Then
'                                For k = LBound(arr, colDimension) To UBound(arr, colDimension)
'                                    tmp = arr(j, k)
'                                    arr(j, k) = arr(i, k)
'                                    arr(i, k) = tmp
'                                Next k
'                            End If
'                        Next j
'                    Next i
'                Else
'                    ' Inverted array
'                    For i = startPosition To UBound(arr, rowDimension) - 1
'                        For j = i + 1 To UBound(arr, rowDimension)
'                            If arr(sortColId, i) > arr(sortColId, j) Then
'                                For k = LBound(arr, colDimension) To UBound(arr, colDimension)
'                                    tmp = arr(k, j)
'                                    arr(k, j) = arr(k, i)
'                                    arr(k, i) = tmp
'                                Next k
'                            End If
'                        Next j
'                    Next i
'                End If
'            Else
'                ' Sort Descending
'                If rowDimension = 1 Then
'                    ' Not inverted array
'                    For i = startPosition To UBound(arr, rowDimension) - 1
'                        For j = i + 1 To UBound(arr, rowDimension)
'                            If arr(i, sortColId) < arr(j, sortColId) Then
'                                For k = LBound(arr, colDimension) To UBound(arr, colDimension)
'                                    tmp = arr(j, k)
'                                    arr(j, k) = arr(i, k)
'                                    arr(i, k) = tmp
'                                Next k
'                            End If
'                        Next j
'                    Next i
'                Else
'                    ' Inverted array
'                    For i = startPosition To UBound(arr, rowDimension) - 1
'                        For j = i + 1 To UBound(arr, rowDimension)
'                            If arr(sortColId, i) < arr(sortColId, j) Then
'                                For k = LBound(arr, colDimension) To UBound(arr, colDimension)
'                                    tmp = arr(k, j)
'                                    arr(k, j) = arr(k, i)
'                                    arr(k, i) = tmp
'                                Next k
'                            End If
'                        Next j
'                    Next i
'                End If ' end inverted / not inverted
'            End If ' end sort ascending/descending
'        End If ' END hard core bubble sort
        
    End If
    BubbleSort2DArray = arr
End Function
Function DefineWinner(ByVal candidatesArr As Variant, _
            ByVal maxiMiniType As String, _
            ByVal kpiName As String) As Variant
' candidatesArr: 1 to 3, 1 to N candidates
' Where:    column 1 - IS trade lists
'           column 2 - IS KPI dictionaries
'           column 3 - OS trade lists
' Return 1 To 2 array.
' arr(1) - IS trade list
' arr(2) - OS trade list
' each 1 To 4, 0 to N
    Dim arr(1 To 2) As Variant
    Dim i As Integer, bestPointer As Integer
    Dim pointVal As Variant
    
    If UBound(candidatesArr, 2) > 1 Then
        pointVal = candidatesArr(2, 1)(kpiName)
        bestPointer = 1
        If maxiMiniType = "maximize" Then
            For i = 1 To UBound(candidatesArr, 2) ' 0-based, has header row
'                Debug.Print "current kpi " & candidatesArr(2, i)(kpiName)
'                Debug.Print "point val " & pointVal
'                Debug.Print "best pointer " & bestPointer
                If candidatesArr(2, i)(kpiName) > pointVal Then
                    pointVal = candidatesArr(2, i)(kpiName)
                    bestPointer = i
                End If
            Next i
        ElseIf maxiMiniType = "minimize" Then
            For i = 1 To UBound(candidatesArr, 2) ' 0-based, has header row
                If candidatesArr(2, i)(kpiName) < pointVal Then
                    pointVal = candidatesArr(2, i)(kpiName)
                    bestPointer = i
                End If
            Next i
        Else
            MsgBox "Error. Should be maximizing or minimizing instead of none."
            arr(1) = candidatesArr(1, bestPointer)
            arr(2) = candidatesArr(3, bestPointer)
        End If
        arr(1) = candidatesArr(1, bestPointer)
        arr(2) = candidatesArr(3, bestPointer)
'        arr(1) = candidatesArr(1, 1) - ERROR )
'        arr(2) = candidatesArr(3, 1) - ERROR )
    Else
        For i = 1 To 2
            ' init empty trade list: 1 to 4, 0 to 0, with header row
            arr(i) = Init2DArr(1, 4, 0, 0)
            arr(i)(1, 0) = "Open date"
            arr(i)(2, 0) = "Close date"
            arr(i)(3, 0) = "Source date"
            arr(i)(4, 0) = "Return"
        Next i
    End If
    DefineWinner = arr
End Function
Function GetFractionMultiplier(ByVal origTradeList As Variant, _
            ByVal mddFreedom As Double, _
            ByVal targetMDD As Double) As Double
' origTradeSet(1 to 4, 0 to trades): Open date, Close date, Source, Return
    Const init_lower_mult As Double = 0
    Const init_upper_mult As Double = 10
    Dim returns() As Variant
    Dim i As Long
    Dim lower_mult As Double, upper_mult As Double, mid_mult As Double
    Dim lower_mdd As Double, upper_mdd As Double, mid_mdd As Double
    Dim mdd_delta As Double
    Dim allPositive As Boolean

' Sanity check
    If UBound(origTradeList, 2) = 0 Then
        GetFractionMultiplier = 1
        Exit Function
    End If

' Collect returns into Series
    ReDim returns(0 To UBound(origTradeList, 2))
    For i = 1 To UBound(returns)
        returns(i) = origTradeList(4, i)
    Next i
    
' Sanity check #2
' If all returns are positive, leave multiplier as 1
    allPositive = True
    For i = 1 To UBound(returns)
        If returns(i) < 0 Then
            allPositive = False
            Exit For
        End If
    Next i
    If allPositive Then
        GetFractionMultiplier = 1
        Exit Function
    End If
    
' GET Upper & Lower multiplicators
    lower_mult = init_lower_mult
    upper_mult = init_upper_mult
    Do Until GetFractionMultiplier_CalcMDDOnly(returns, upper_mult) > targetMDD
        lower_mult = upper_mult
        upper_mult = upper_mult * 2
    Loop
    mid_mult = (lower_mult + upper_mult) / 2
' NARROW search
    mdd_delta = mddFreedom * 2  ' initialize delta
    Do Until mdd_delta <= mddFreedom
        mid_mdd = GetFractionMultiplier_CalcMDDOnly(returns, mid_mult)
        mdd_delta = Abs(mid_mdd - targetMDD)
        If mdd_delta <= mddFreedom Then
            Exit Do
        Else
            If mid_mdd > targetMDD Then
                upper_mult = mid_mult
            ElseIf mid_mdd < targetMDD Then
                lower_mult = mid_mult
            Else
                Exit Do
            End If
            mid_mult = (lower_mult + upper_mult) / 2
        End If
    Loop
    GetFractionMultiplier = mid_mult
End Function
Function GetFractionMultiplier_CalcMDDOnly(ByVal returns As Variant, _
            ByVal multiplier As Double) As Double
    Dim eh() As Variant ' Equity & HWM
    Dim dd() As Variant ' Drawdown
    Dim i As Long
    
    ReDim eh(1 To 2, 0 To UBound(returns))
    eh(1, 0) = 1   ' equity
    eh(2, 0) = 1   ' hwm
    ReDim dd(0 To UBound(returns))
    dd(0) = 0   ' dd
    For i = 1 To UBound(eh, 2)
        eh(1, i) = eh(1, i - 1) * (1 + multiplier * returns(i))     ' Equity
        eh(2, i) = WorksheetFunction.Max(eh(2, i - 1), eh(1, i))    ' HWM
        dd(i) = (eh(2, i) - eh(1, i)) / eh(2, i)                    ' Drawdown
    Next i
    GetFractionMultiplier_CalcMDDOnly = WorksheetFunction.Max(dd)
End Function
Function ApplyFractionMultiplier(ByVal arr As Variant, _
            ByVal multiplier As Double) As Variant
    Dim result_arr() As Variant
    Dim i As Long
    result_arr = arr
    If UBound(result_arr, 2) = 0 Then
        ApplyFractionMultiplier = result_arr
        Exit Function
    End If
    For i = 1 To UBound(result_arr, 2)
        result_arr(4, i) = result_arr(4, i) * multiplier
    Next i
    ApplyFractionMultiplier = result_arr
End Function
Function GetKPIFormatting() As Dictionary
    Dim dict As New Dictionary
    With dict
        .Add "Sharpe Ratio", "0.00"
        .Add "R-squared", "0.00"
        .Add "Annualized Return", "0.0%"
        .Add "MDD", "0.0%"
        .Add "Recovery Factor", "0.00"
        .Add "Trades per Month", "0.00"
        .Add "Win Ratio", "0.00%"
        .Add "Avg Winner/Loser", "0.000"
        .Add "Avg Trade", "0.00%"
        .Add "Profit Factor", "0.00"
    End With
    Set GetKPIFormatting = dict
End Function
Function DictionaryToArray(ByVal origDict As Dictionary, _
            ByVal arrHorizontal As Boolean) As Variant
    Dim arr As Variant
    Dim i As Integer
' convert dictionary to 2D array, not inverted
' columns as dictionary keys
' rows as values
    If arrHorizontal Then
        ReDim arr(1 To 2, 1 To origDict.Count) ' keys as columns
        For i = 0 To origDict.Count - 1
            arr(1, i + 1) = origDict.Keys(i)
            arr(2, i + 1) = origDict(origDict.Keys(i))
        Next i
    Else    ' inverted
        ReDim arr(1 To origDict.Count, 1 To 2) ' keys as rows
        For i = 0 To origDict.Count - 1
            arr(i + 1, 1) = origDict.Keys(i)
            arr(i + 1, 2) = origDict(origDict.Keys(i))
        Next i
    End If
    DictionaryToArray = arr
End Function
Function GetDirectoryPathFolderPicker(ByVal dialTitle As String, _
            ByVal okBtnName As String) As String
' STATEMENT sheet
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = dialTitle
        .AllowMultiSelect = True
        .ButtonName = okBtnName
    End With
    If fd.Show = 0 Then
        GetDirectoryPathFolderPicker = ""
    Else
        GetDirectoryPathFolderPicker = CStr(fd.SelectedItems(1))
    End If
End Function
Function StatementTargetWbSaveName(ByVal saveDir As String) As String
    Dim saveName As String
    saveName = saveDir & "\statement"
    saveName = PathIncrementIndex(saveName, True)
    StatementTargetWbSaveName = saveName
End Function
Function DateRangesDict() As Dictionary
' for STATEMENT
    Dim dict As New Dictionary
    dict.Add "date", Nothing
    dict.Add "openDate", Nothing
    dict.Add "closeDate", Nothing
    Set DateRangesDict = dict
End Function
Function SortDateRangesDict() As Dictionary
' for STATEMENT
    Dim dict As New Dictionary
    dict.Add "date", Nothing
    dict.Add "closeDate", Nothing
    Set SortDateRangesDict = dict
End Function
'''' OLD BUBBLESORT
'Function BubbleSort2DArray(ByVal origArr As Variant, _
'            ByVal isInverted As Boolean, _
'            ByVal hasHeaderRow As Boolean, _
'            ByVal sortAscending As Boolean, _
'            ByVal sortColId As Integer, _
'            ByRef sortWorkSheet As Worksheet, _
'            ByRef sortCells As Range) As Variant
'    Dim arr As Variant
'    Dim colDimension As Integer
'    Dim rowDimension As Integer
'    Dim startPosition As Long
'    Dim i As Long, j As Long
'    Dim k As Integer
'    Dim tmp As Variant
'
'    If isInverted Then
'        colDimension = 1
'        rowDimension = 2
'    Else
'        rowDimension = 1
'        colDimension = 2
'    End If
'    If hasHeaderRow Then
'        startPosition = LBound(arr, rowDimension) + 1
'    Else
'        startPosition = LBound(arr, rowDimension)
'    End If
'    If UBound(arr, rowDimension) > startPosition Then
'        If sortAscending Then
'            ' Sort Ascending
'            If rowDimension = 1 Then
'                ' Not inverted array
'                For i = startPosition To UBound(arr, rowDimension) - 1
'                    For j = i + 1 To UBound(arr, rowDimension)
'                        If arr(i, sortColId) > arr(j, sortColId) Then
'                            For k = LBound(arr, colDimension) To UBound(arr, colDimension)
'                                tmp = arr(j, k)
'                                arr(j, k) = arr(i, k)
'                                arr(i, k) = tmp
'                            Next k
'                        End If
'                    Next j
'                Next i
'            Else
'                ' Inverted array
'                For i = startPosition To UBound(arr, rowDimension) - 1
'                    For j = i + 1 To UBound(arr, rowDimension)
'                        If arr(sortColId, i) > arr(sortColId, j) Then
'                            For k = LBound(arr, colDimension) To UBound(arr, colDimension)
'                                tmp = arr(k, j)
'                                arr(k, j) = arr(k, i)
'                                arr(k, i) = tmp
'                            Next k
'                        End If
'                    Next j
'                Next i
'            End If
'        Else
'            ' Sort Descending
'            If rowDimension = 1 Then
'                ' Not inverted array
'                For i = startPosition To UBound(arr, rowDimension) - 1
'                    For j = i + 1 To UBound(arr, rowDimension)
'                        If arr(i, sortColId) < arr(j, sortColId) Then
'                            For k = LBound(arr, colDimension) To UBound(arr, colDimension)
'                                tmp = arr(j, k)
'                                arr(j, k) = arr(i, k)
'                                arr(i, k) = tmp
'                            Next k
'                        End If
'                    Next j
'                Next i
'            Else
'                ' Inverted array
'                For i = startPosition To UBound(arr, rowDimension) - 1
'                    For j = i + 1 To UBound(arr, rowDimension)
'                        If arr(sortColId, i) < arr(sortColId, j) Then
'                            For k = LBound(arr, colDimension) To UBound(arr, colDimension)
'                                tmp = arr(k, j)
'                                arr(k, j) = arr(k, i)
'                                arr(k, i) = tmp
'                            Next k
'                        End If
'                    Next j
'                Next i
'            End If
'        End If
'    End If
'    BubbleSort2DArray = arr
' End Function

' MODULE: Sheet4
Option Explicit


' MODULE: Sheet5
Option Explicit


' MODULE: Sheet2
Option Explicit


' MODULE: Statement
Option Explicit
' Dictionary
    Dim param As Dictionary
' Range
    Dim mainC As Range
' Workbook
    Dim targetWb As Workbook

Sub ProcessStatementsMain()
' Dictionary
    Dim datesDict As Dictionary
    Dim sortDatesDict As Dictionary
' Integer
    Dim descrCol As Integer
    Dim iCsv As Integer
    Dim iDateCol As Integer
    Dim iDir As Integer
    Dim nextFreeRow As Integer
' Long
    Dim lastRow As Long
' Range
    Dim cell As Range
    Dim rg As Range
    Dim wsC As Range
' String
    Dim dateColValue As String
    Dim qtName As String
    Dim reportType As String
' Variant
    Dim dateCol As Variant
    Dim filesList As Variant
' Worksheet
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Call Statement_Init_Parameters(param)
    Call NewWorkbookSheetsCount(targetWb, 3)
    Set datesDict = DateRangesDict
    Set sortDatesDict = SortDateRangesDict
    For iDir = 0 To param.Count - 2
        reportType = param.Keys(iDir)
        filesList = DirectoryFilesList(param(reportType), False, True)
        Set ws = targetWb.Sheets(iDir + 1)
        ws.Activate
        ws.Name = reportType
        Application.StatusBar = "Importing " & reportType & "."
        Set wsC = ws.Cells
        nextFreeRow = 1
        For iCsv = LBound(filesList) To UBound(filesList)
            qtName = GetBasename(filesList(iCsv))
            qtName = Left(qtName, Len(qtName) - 4)
            With ws.QueryTables.Add( _
                Connection:="TEXT;" & filesList(iCsv), _
                Destination:=wsC(nextFreeRow, 1))
                .Name = qtName
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
'                .TextFilePlatform = 866
                .TextFilePlatform = 65001 ' Unicode UTF-8
                .TextFileStartRow = 1
                .TextFileParseType = xlDelimited
                .TextFileTextQualifier = xlTextQualifierDoubleQuote
                .TextFileConsecutiveDelimiter = False
                .TextFileTabDelimiter = True
                .TextFileSemicolonDelimiter = False
                .TextFileCommaDelimiter = True
                .TextFileSpaceDelimiter = False
                .TextFileColumnDataTypes = Array(1, 1, 1, 1)
                .TextFileTrailingMinusNumbers = True
                .Refresh BackgroundQuery:=False
            End With
            If nextFreeRow > 1 Then ' remove header of newly added CSV
                ws.rows(nextFreeRow).EntireRow.Delete
            End If
            nextFreeRow = wsC(ws.rows.Count, 1).End(xlUp).Row + 1
            If wsC(nextFreeRow - 1, 1) = "TOTAL" Then
                ws.rows(nextFreeRow - 1).EntireRow.Delete
                nextFreeRow = nextFreeRow - 1
            End If
        Next iCsv
        ' edit dates, remove duplicates, sort ascending
        For iDateCol = 0 To datesDict.Count - 1
            dateColValue = datesDict.Keys(iDateCol)
            If Not wsC.Find(what:=dateColValue, _
                    after:=wsC(1, 1), _
                    searchorder:=xlByRows, _
                    lookat:=xlWhole) Is Nothing Then
                dateCol = wsC.Find(what:=dateColValue, _
                    after:=wsC(1, 1), _
                    searchorder:=xlByRows, _
                    lookat:=xlWhole).Column
                ' edit dates
                lastRow = wsC(ws.rows.Count, dateCol).End(xlUp).Row
                Set rg = ws.Range(wsC(2, dateCol), wsC(lastRow, dateCol))
                For Each cell In rg
                    cell.Value = Replace(cell.Value, "T", " ", 1, -1, vbTextCompare)
                    If ws.Name = "Portfolio Summary" Then
                        cell.Value = Replace(cell.Value, "00:00:00Z", " ", 1, -1, vbTextCompare)
                    End If
                    cell.Value = Replace(cell.Value, "Z", "", 1, -1, vbTextCompare)
                    cell.Value = Replace(cell.Value, "+00:00", "", 1, -1, vbTextCompare)
                    cell.Value = Replace(cell.Value, "-", ".", 1, -1, vbTextCompare)
                    cell.Value = CDate(cell.Value)
                Next cell
                ' add "_n/a_"
                If ws.Name = "Positions Close" Then
                    descrCol = wsC.Find(what:="description", _
                        after:=wsC(1, 1), _
                        searchorder:=xlByRows, _
                        lookat:=xlWhole).Column
                    Set rg = ws.Range(wsC(2, descrCol), wsC(lastRow, descrCol))
                    For Each cell In rg
                        If cell.Value = "" Then
                            cell.Value = "_n/a_"
                        End If
                    Next cell
                End If
                wsC.EntireColumn.AutoFit
            End If
        Next iDateCol
        ws.rows("1:1").AutoFilter
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
            .FreezePanes = True
        End With
        wsC.columns.AutoFit
        ' Remove duplicates
        Set rg = wsC(1, 1).CurrentRegion
        Select Case ws.Name
            Case Is = "Funds History"
                rg.RemoveDuplicates columns:=1, Header:=xlYes
            Case Is = "Portfolio Summary"
                rg.RemoveDuplicates columns:=2, Header:=xlYes
            Case Is = "Positions Close"
                rg.RemoveDuplicates columns:=4, Header:=xlYes
        End Select
        ' Sort ascending
        For iDateCol = 0 To sortDatesDict.Count - 1
            dateColValue = sortDatesDict.Keys(iDateCol)
            If Not wsC.Find(what:=dateColValue, _
                    after:=wsC(1, 1), _
                    searchorder:=xlByRows, _
                    lookat:=xlWhole) Is Nothing Then
                dateCol = wsC.Find(what:=dateColValue, _
                    after:=wsC(1, 1), _
                    searchorder:=xlByRows, _
                    lookat:=xlWhole).Column
                Set rg = wsC(1, 1).CurrentRegion
                Set rg = rg.Offset(1).Resize(rg.rows.Count - 1)
                With ws.Sort
                    .SortFields.Clear
                    .SortFields.Add Key:=wsC(1, dateCol), SortOn:=xlSortOnValues, Order:=xlAscending
                    .SetRange rg
                    .Header = xlNo
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                wsC(1, 1).Select
                Exit For
            End If
        Next iDateCol
    Next iDir
    Set ws = targetWb.Sheets("Portfolio Summary")
    Set wsC = ws.Cells
    Application.StatusBar = "Chart, calculations..."
    Call PortfolioSummaryComputations(ws, wsC)
    Call PositionsCloseComputations
    Call PositionDescriptionsInstruments("description", "Descriptions")
    Call PositionDescriptionsInstruments("instrument", "Instruments")
    Application.StatusBar = "Saving target book..."
    targetWb.SaveAs fileName:=PathIncrementIndex(param("Target Directory") & "\statement", True)
'    targetWb.Close
    targetWb.Sheets(2).Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub
Sub PositionDescriptionsInstruments(ByVal desValue As String, _
            ByVal newSheetName As String)
    Dim desWs As Worksheet
    Dim desC As Range
    Dim desD As Dictionary
    Dim desCol As Integer
    Dim desItemValue As Range
    Dim posWs As Worksheet
    Dim posC As Range
    Dim rg As Range
    Dim lastRow As Integer
    Dim columnsDict As Dictionary
    Dim posCloseRow As Long
    Dim valUnits As Double
    
    Dim winnersCount As Long
    Dim losersCount As Long
    Dim winnersSum As Double
    Dim losersSum As Double
    Dim i As Integer
    
    Set posWs = targetWb.Sheets("Positions Close")
    Set posC = posWs.Cells
    desCol = posC.Find(what:=desValue, _
            after:=posC(1, 1), _
            searchorder:=xlByRows, _
            lookat:=xlWhole).Column
    
    Set columnsDict = GetColumnsDictionary(posC)
    
    lastRow = posC(posWs.rows.Count, 1).End(xlUp).Row
    
    Set rg = posWs.Range(posC(2, desCol), posC(lastRow, desCol))

    Set desWs = targetWb.Sheets.Add(after:=Sheets(targetWb.Sheets.Count))
    desWs.Name = newSheetName
    desWs.Activate
    Set desC = desWs.Cells
    
    Set desD = DescriptionInstrumentDictionary
    
    desC(1, 1) = desValue
    Set rg = posWs.Range(posC(2, desCol), posC(lastRow, desCol))
    rg.Copy desC(2, 1)
    Set rg = desC(1, 1).CurrentRegion
    rg.RemoveDuplicates columns:=1, Header:=xlYes
    Set rg = desC(1, 1).CurrentRegion
    rg.Sort Key1:=desC(1, 1), _
            Order1:=xlAscending, _
            Header:=xlYes
    Set rg = rg.Offset(1, 0).Resize(rg.rows.Count - 1)
    
    ' print header
    For i = 0 To desD.Count - 1
        desC(1, i + 2) = desD.Keys(i)
    Next i
    
    For Each desItemValue In rg
        winnersCount = 0
        losersCount = 0
        winnersSum = 0
        losersSum = 0
        For posCloseRow = 2 To lastRow
            If desItemValue = posC(posCloseRow, desCol) Then
                ' positionsCount
                desD("positionsCount") = desD("positionsCount") + 1
                If posC(posCloseRow, columnsDict("side")) = "LONG" Then
                ' longPositions
                    desD("longPositions") = desD("longPositions") + 1
                    ' valueInUnits
                    valUnits = posC(posCloseRow, columnsDict("closePrice")) - posC(posCloseRow, columnsDict("openPrice"))
                Else
                ' shortPositions
                    desD("shortPositions") = desD("shortPositions") + 1
                    ' valueInUnits
                    valUnits = posC(posCloseRow, columnsDict("openPrice")) - posC(posCloseRow, columnsDict("closePrice"))
                End If
                desD("valueUnits") = desD("valueUnits") + valUnits
                ' preparatory computations for winRation & avgWLValUnits
                If valUnits > 0 Then
                    winnersCount = winnersCount + 1
                    winnersSum = winnersSum + valUnits
                Else
                    losersCount = losersCount + 1
                    losersSum = losersSum + Abs(valUnits)
                End If
                ' netPl
                desD("netPL") = desD("netPL") + posC(posCloseRow, columnsDict("netPl"))
                ' grossPl
                desD("grossPL") = desD("grossPL") + posC(posCloseRow, columnsDict("grossPl"))
                ' swaps
                desD("swaps") = desD("swaps") + posC(posCloseRow, columnsDict("swap"))
                ' commissions
                desD("commissions") = desD("commissions") + posC(posCloseRow, columnsDict("commission"))
                ' amountTraded
                desD("amountTraded") = desD("amountTraded") + Abs(posC(posCloseRow, columnsDict("amount")))
                ' approxReturns
                desD("approxReturns") = desD("approxReturns") + posC(posCloseRow, columnsDict("_approxReturn"))
            End If
        Next posCloseRow
        ' winRatio
        desD("winRatio") = winnersCount / desD("positionsCount")
        ' avgWLValUnits
        If winnersCount = 0 And losersCount > 0 Then
            desD("avgWLValUnits") = -999
        ElseIf losersSum = 0 Or losersCount = 0 Then
            desD("avgWLValUnits") = 999
        Else
            desD("avgWLValUnits") = (winnersSum / winnersCount) / (losersSum / losersCount)
        End If
        ' PRINT DICTIONARY VALUES
        For i = 0 To desD.Count - 1
            With desC(desItemValue.Row, i + 2)
                .Value = desD(desD.Keys(i))
                If desD.Keys(i) = "approxReturns" Then
                    .NumberFormat = "0.00%"
                ElseIf desD.Keys(i) = "winRatio" Then
                    .NumberFormat = "0.0%"
                ElseIf desD.Keys(i) = "avgWLValUnits" Then
                    If desD(desD.Keys(i)) <> 999 Or desD(desD.Keys(i)) <> -999 Then
                        .NumberFormat = "0.00"
                    End If
                End If
            End With
        Next i
        ' reinitialize dictionary
        Set desD = DescriptionInstrumentDictionary
    Next desItemValue
    desWs.rows("1:1").AutoFilter
    desWs.columns.AutoFit
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub
Function GetColumnsDictionary(ByVal wsC As Range) As Dictionary
    Dim dict As New Dictionary
    dict.Add "side", wsC.Find(what:="side", after:=wsC(1, 1), _
            searchorder:=xlByRows, lookat:=xlWhole).Column
    dict.Add "amount", wsC.Find(what:="amount", after:=wsC(1, 1), _
            searchorder:=xlByRows, lookat:=xlWhole).Column
    dict.Add "openPrice", wsC.Find(what:="openPrice", after:=wsC(1, 1), _
            searchorder:=xlByRows, lookat:=xlWhole).Column
    dict.Add "closePrice", wsC.Find(what:="closePrice", after:=wsC(1, 1), _
            searchorder:=xlByRows, lookat:=xlWhole).Column
    dict.Add "swap", wsC.Find(what:="swap", after:=wsC(1, 1), _
            searchorder:=xlByRows, lookat:=xlWhole).Column
    dict.Add "commission", wsC.Find(what:="commission", after:=wsC(1, 1), _
            searchorder:=xlByRows, lookat:=xlWhole).Column
    dict.Add "netPl", wsC.Find(what:="netPl", after:=wsC(1, 1), _
            searchorder:=xlByRows, lookat:=xlWhole).Column
    dict.Add "grossPl", wsC.Find(what:="grossPl", after:=wsC(1, 1), _
            searchorder:=xlByRows, lookat:=xlWhole).Column
    dict.Add "_approxReturn", wsC.Find(what:="_approxReturn", after:=wsC(1, 1), _
            searchorder:=xlByRows, lookat:=xlWhole).Column
    Set GetColumnsDictionary = dict
End Function

Function DescriptionInstrumentDictionary() As Dictionary
    Const defVal As Variant = 0
    Dim dict As New Dictionary
    dict.Add "positionsCount", defVal
    dict.Add "longPositions", defVal
    dict.Add "shortPositions", defVal
    dict.Add "winRatio", defVal
    dict.Add "avgWLValUnits", defVal
    dict.Add "netPL", defVal
    dict.Add "grossPL", defVal
    dict.Add "swaps", defVal
    dict.Add "commissions", defVal
    dict.Add "valueUnits", defVal
    dict.Add "amountTraded", defVal
    dict.Add "approxReturns", defVal
    Set DescriptionInstrumentDictionary = dict
End Function
Sub PositionsCloseComputations()
    Dim wsPos As Worksheet
    Dim cPos As Range
    Dim wsPort As Worksheet
    Dim cPort As Range
    Dim prevSettl As Double
    Dim approxRetCol As Integer
    Dim netPlCol As Integer
    Dim tradeDate As Long
    Dim rg As Range
    Dim cell As Range
    Dim sumRg As Range
    Dim sumCell As Range
    Dim lastRow As Integer
    Dim balanceCol As Integer
    
    Set wsPos = targetWb.Sheets("Positions Close")
    Set cPos = wsPos.Cells
    Set wsPort = targetWb.Sheets("Portfolio Summary")
    Set cPort = wsPort.Cells
    
    ' add appoximate percentage return
    wsPos.Activate
    wsPos.rows("1:1").AutoFilter
    approxRetCol = cPos(1, wsPos.columns.Count).End(xlToLeft).Column + 1
    cPos(1, approxRetCol) = "_approxReturn"
    netPlCol = cPos.Find(what:="netPl", _
            after:=cPos(1, 1), _
            searchorder:=xlByRows, _
            lookat:=xlWhole).Column
    lastRow = cPos(wsPos.rows.Count, 1).End(xlUp).Row
    Set rg = wsPos.Range(cPos(2, netPlCol), cPos(lastRow, netPlCol))
    lastRow = cPort(wsPort.rows.Count, 1).End(xlUp).Row
    Set sumRg = wsPort.Range(cPort(3, 2), cPort(lastRow, 2))
    balanceCol = cPort.Find(what:="balance", _
            after:=cPort(1, 1), _
            searchorder:=xlByRows, _
            lookat:=xlWhole).Column
    For Each cell In rg
        tradeDate = Int(cPos(cell.Row, 1).Value)
        For Each sumCell In sumRg
            If Int(sumCell.Value) = tradeDate Then
                prevSettl = cPort(sumCell.Row - 1, balanceCol)
'                cPos(cell.Row, approxRetCol) = cPos(cell.Row, netPlCol) / prevSettl
                Exit For
            End If
        Next sumCell
        With cPos(cell.Row, approxRetCol)
            .Value = cPos(cell.Row, netPlCol) / prevSettl
            .NumberFormat = "0.00%"
        End With
    Next cell
    wsPos.rows("1:1").AutoFilter
    wsPos.columns.AutoFit
End Sub


Sub tportcomp()
    Dim ws As Worksheet
    Dim wsC As Range
    Application.ScreenUpdating = False
    Set ws = ActiveSheet
    Set wsC = ws.Cells
    Call PortfolioSummaryComputations(ws, wsC)
    Application.ScreenUpdating = True
End Sub
Sub PortfolioSummaryComputations(ByRef ws As Worksheet, _
            ByRef wsC As Range)
    Dim lastCol As Integer
    Dim i As Integer
    Dim rg As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim dayReturnCol As Integer
    Dim returnCurveCol As Integer
    Dim resVal As Double
    Dim rngX As Range
    Dim rngY As Range
    Dim rngZ As Range
    Dim rngCol As Integer
    
    ws.Activate
    ws.rows("1:1").AutoFilter
    ws.rows(2).EntireRow.Insert
    lastCol = wsC(1, ws.columns.Count).End(xlToLeft).Column
    Set rg = ws.Range(wsC(3, 1), wsC(3, lastCol))
    
    For Each cell In rg
        Select Case True
            Case Application.IsText(cell)
                ' CellType = "Text"
                If cell.Value = "SUMMARY" Then
                    cell.Offset(-1, 0) = "_DayZero_"
                Else
                    cell.Offset(-1, 0) = cell.Value
                End If
            Case IsDate(cell)
                ' CellType = "Date"
                With cell.Offset(-1, 0)
                    .Value = cell.Value - 1
                    .NumberFormat = "DD.MM.YY"
                End With
            Case IsNumeric(cell)
                ' CellType = "Value"
                cell.Offset(-1, 0) = 0
        End Select
    Next cell
    dayReturnCol = lastCol + 1
    wsC(1, dayReturnCol) = "_dayReturn"
    wsC(2, dayReturnCol) = 0
    returnCurveCol = dayReturnCol + 1
    wsC(1, returnCurveCol) = "_returnCurve"
    wsC(2, returnCurveCol) = 0
    lastRow = wsC(ws.rows.Count, 1).End(xlUp).Row
    Set rg = ws.Range(wsC(3, dayReturnCol), wsC(lastRow, dayReturnCol))
    For Each cell In rg
        ' daily return
        If cell.Offset(-1, -2) = 0 Then
            resVal = 0
        Else
            resVal = (cell.Offset(0, -3) - cell.Offset(0, -5) - cell.Offset(0, -4)) / cell.Offset(-1, -2)
        End If
        With cell
            .Value = resVal
            .NumberFormat = "0.00%"
        End With
        ' return curve
        With cell.Offset(0, 1)
            .Value = (cell.Offset(-1, 1) + 1) * (1 + cell.Value) - 1
            .NumberFormat = "0.0%"
        End With
    Next cell
    ws.rows("1:1").AutoFilter
    ws.columns.AutoFit
' Chart
    rngCol = wsC.Find(what:="date", _
            after:=wsC(1, 1), _
            searchorder:=xlByRows, _
            lookat:=xlWhole).Column
    Set rngX = ws.Range(wsC(1, rngCol), wsC(lastRow, rngCol))
    Set rngY = ws.Range(wsC(1, dayReturnCol), wsC(lastRow, dayReturnCol))
    Set rngZ = ws.Range(wsC(1, returnCurveCol), wsC(lastRow, returnCurveCol))
    Call StatementChartRangesXYZ(rngX, rngY, rngZ, wsC, 2, returnCurveCol + 1)
End Sub
Sub ClickLocateFundsHistory()
    Call ClickLocateADirectory("Funds History")
End Sub
Sub ClickLocatePortfolioSummary()
    Call ClickLocateADirectory("Portfolio Summary")
End Sub
Sub ClickLocatePositionsClose()
    Call ClickLocateADirectory("Positions Close")
End Sub
Sub ClickLocateStatementTargetDirectory()
    Call ClickLocateADirectory("Target Directory")
End Sub
Sub ClickLocateRoot()
    Dim dialogTitle As String
    Dim okButtonName As String
    Dim insertRow As Integer
    Dim insertCol As Integer
    Dim fundsRow As Integer
    Dim portfolioRow As Integer
    Dim positionsRow As Integer
    Dim targetDirRow As Integer
    Dim dirPath As String
    Application.ScreenUpdating = False
    Call StatementClickLocate_Inits(mainC, insertRow, insertCol, "Root", dialogTitle, okButtonName)
    Call ClickLocateRoot_InitsPart2(fundsRow, portfolioRow, positionsRow, targetDirRow)
    dirPath = GetDirectoryPathFolderPicker(dialogTitle, okButtonName)
    mainC(fundsRow, insertCol) = dirPath & "\funds-history"
    mainC(portfolioRow, insertCol) = dirPath & "\portfolio-summary"
    mainC(positionsRow, insertCol) = dirPath & "\positions-close"
    mainC(targetDirRow, insertCol) = dirPath
    Application.ScreenUpdating = True
End Sub
Sub ClickLocateADirectory(ByVal sourceType As String)
    Dim dialogTitle As String
    Dim okButtonName As String
    Dim insertRow As Integer
    Dim insertCol As Integer
    Dim cellVal As String
    Application.ScreenUpdating = False
    Call StatementClickLocate_Inits(mainC, insertRow, insertCol, sourceType, dialogTitle, okButtonName)
    cellVal = GetDirectoryPathFolderPicker(dialogTitle, okButtonName)
    If cellVal = "" Then
        Application.ScreenUpdating = True
        Exit Sub
    Else
        mainC(insertRow, insertCol) = cellVal
        columns(insertCol).AutoFit
        Application.ScreenUpdating = True
    End If
End Sub
Sub DescriptionFilterChart()
' Dictionary
    Dim kpis As Dictionary
' Integer
    Dim desCol As Integer
    Dim j As Integer
    Dim lastCol As Integer
' Long
    Dim i As Long
    Dim lastRow As Long
' Range
    Dim posC As Range
    Dim rangeX As Range
    Dim rangeY As Range
    Dim tarC As Range
    Dim tmpC As Range
' String
    Dim desType As String
' Variant
    Dim calDays As Variant
    Dim calDaysLong As Variant
    Dim dailyEquity As Variant
    Dim dateEnd As Variant
    Dim dateStart As Variant
    Dim desVal As Variant
    Dim onlyCloseDates As Variant
    Dim onlyOpenDates As Variant
    Dim preTradesList As Variant    ' INVERTED
    Dim tradesList As Variant
' Worksheet
    Dim posWs As Worksheet
    Dim tarWs As Worksheet
    Dim tmpWs As Worksheet

    Application.ScreenUpdating = False
    desVal = ActiveCell.Value
    desType = Cells(1, ActiveCell.Column)
    
' Collect trades
    Set posWs = Sheets("Positions Close")
    Set posC = posWs.Cells
    
    If Not posC.Find(what:=desType, after:=posC(1, 1), _
            searchorder:=xlByRows, _
            lookat:=xlWhole) Is Nothing Then
        desCol = posC.Find(what:=desType, after:=posC(1, 1), _
            searchorder:=xlByRows, _
            lookat:=xlWhole).Column
    Else
        Application.ScreenUpdating = True
        Exit Sub
    End If
    lastRow = posC(posWs.rows.Count, 1).End(xlUp).Row
    lastCol = posC(1, 1).End(xlToRight).Column
    
    ReDim preTradesList(1 To lastCol, 0 To 0)
    For j = 1 To lastCol
        preTradesList(j, 0) = posC(1, j)
    Next j
    
    For i = 2 To lastRow
        If posC(i, desCol) = desVal Then
            ReDim Preserve preTradesList(1 To lastCol, 0 To UBound(preTradesList, 2) + 1)
            For j = 1 To lastCol
                preTradesList(j, UBound(preTradesList, 2)) = posC(i, j)
            Next j
        End If
    Next i
' get trades list
    ReDim tradesList(1 To 4, 0 To UBound(preTradesList, 2)) ' INVERTED
    ReDim onlyOpenDates(1 To UBound(tradesList, 2))
    ReDim onlyCloseDates(1 To UBound(tradesList, 2))
    For i = LBound(tradesList, 2) To UBound(tradesList, 2)
        tradesList(1, i) = preTradesList(1, i)
        tradesList(2, i) = preTradesList(2, i)
        tradesList(3, i) = preTradesList(18, i)
        tradesList(4, i) = preTradesList(19, i)
        If i > 0 Then
            onlyOpenDates(i) = CLng(tradesList(1, i))
            onlyCloseDates(i) = CLng(tradesList(2, i))
        End If
    Next i
    dateStart = WorksheetFunction.Min(onlyOpenDates)
    dateEnd = WorksheetFunction.Max(onlyCloseDates)
    ' sort tradeslist
    Set tmpWs = Worksheets.Add(after:=Sheets(ActiveWorkbook.Sheets.Count))
    Set tmpC = tmpWs.Cells
    tradesList = BubbleSort2DArray(tradesList, True, True, True, 2, tmpWs, tmpC)
    Application.DisplayAlerts = False
    tmpWs.Delete
    Application.DisplayAlerts = True
    
    dailyEquity = GetDailyEquityFromTradeSet(tradesList, dateStart, dateEnd)
    calDays = GenerateCalendarDays(dateStart, dateEnd)
    calDaysLong = GenerateLongDays(UBound(calDays))
    Set kpis = CalcKPIs(tradesList, dateStart, dateEnd, calDays, calDaysLong)
' Print to new sheet
    Set tarWs = Worksheets.Add(after:=Sheets(ActiveWorkbook.Sheets.Count))
    desVal = Replace(desVal, "/", "", 1, -1, vbTextCompare)
    tarWs.Name = tarWs.Index & "_" & desVal
    Set tarC = tarWs.Cells
    
    Call Print_2D_Array(preTradesList, True, 0, 0, tarC)
    Call Print_2D_Array(dailyEquity, False, 3, 20, tarC)
    Call PrintDictionary(kpis, True, 0, 20, tarC)
    
    Set rangeX = tarC(4, 21).CurrentRegion
    Set rangeX = rangeX.Resize(rangeX.rows.Count, rangeX.columns.Count - 1)
    Set rangeY = tarC(4, 21).CurrentRegion
    Set rangeY = rangeY.Offset(0, 1).Resize(rangeY.rows.Count, rangeY.columns.Count - 1)
    Call StatementChartRangesXandY(rangeX, rangeY, tarC, 4, 23, False, 1, "Equity curve, filter=" & desVal)
    Application.ScreenUpdating = True
End Sub

' MODULE: VersionControl
Option Explicit

' Run GitSave() to export code and modules.
'
' Source:
' https://github.com/Vitosh/VBA_personal/blob/master/VBE/GitSave.vb
' The code below is slightly modified to include a list of
' modules you want to ignore.

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
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        If Not IsStringInList(item.Name, ignoreList) Then
            
            lineToPrint = vbNewLine & "' MODULE: " & item.CodeModule.Name & vbNewLine
            If item.CodeModule.CountOfLines > 0 Then
                lineToPrint = lineToPrint & item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
            Else
                lineToPrint = lineToPrint & "' empty" & vbNewLine
            End If
'            Debug.Print lineToPrint
            textToPrint = textToPrint & vbCrLf & lineToPrint
            
        End If
    Next item
    
    Dim pathToExport As String: pathToExport = parentFolder & dirNameCode
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    SaveTextToFile textToPrint, pathToExport & "\all_code.vb"
    
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
            
            
            
            If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
            Call SaveTextToFile(lineToPrint, pathToExport & "\" & item.CodeModule.Name & "_code.vb")
        
        End If
    Next item

End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        If Not IsStringInList(lineToPrint, ignoreList) Then
'            Debug.Print lineToPrint
            textToPrint = textToPrint & vbCrLf & lineToPrint
        End If
    Next item
    
    Dim pathToExport As String: pathToExport = parentFolder & dirNameCode
    SaveTextToFile textToPrint, pathToExport & "\all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = parentFolder & dirNameModules
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

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
'                Debug.Print unitsCount & " exporting " & filePath
                component.Export pathToExport & "\" & filePath
            End If
            
        End If
        
    Next

'    Debug.Print "Exported at " & pathToExport
    
End Sub

Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub

Function IsStringInList(ByVal whatString As String, whatList As Variant) As Boolean
' True if string is found in the list.
' Pass the list as Array.

    IsStringInList = Not (IsError(Application.Match(whatString, whatList, 0)))

End Function
