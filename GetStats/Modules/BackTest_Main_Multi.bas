Attribute VB_Name = "BackTest_Main_Multi"
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
