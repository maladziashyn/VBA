Attribute VB_Name = "WFA_Main"
Option Explicit

Dim A               As Variant          ' RESULT ARRAY for each windowSet
Dim param           As Dictionary
Dim datesISOS       As Variant
Dim fwdCalDays      As Variant
Dim fwdCalDaysLong  As Variant
Dim scanWb          As Workbook
Dim scanWs          As Worksheet
Dim scanC           As Range
Dim resultArr       As Variant
Dim targetWb        As Workbook
Dim targetWs        As Worksheet
Dim targetC         As Range
Dim kpiFormatting   As Dictionary

Dim iDir            As Integer
Dim iFile           As Integer
Dim iSheet          As Integer
Dim iWindowSet      As Integer
Dim iPermutation    As Integer
Dim permutationID   As Integer
Dim iDateSlot       As Integer

Dim newDirPath      As String
Dim sortWs          As Worksheet
Dim sortC           As Range
Dim sortColumnId    As Integer

Sub WFA_Run()

'Dim timer0 As Single
'timer0 = Timer

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

    Application.ScreenUpdating = False
    exitOnError = False
    Call Init_Parameters(param, exitOnError, errorMsg, sortWs, sortC, sortColumnId, kpiFormatting)
    
    If VarType(param("IS/OS windows")) = 11 Then
        MsgBox "No suitable IS/OS windows." & vbNewLine _
            & "Please add relevant values."
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
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

' ***** SHEET ********************************************************
                For iSheet = 3 To sheetsCount
                    
                    If iSheet Mod 50 = 0 Then
                        DoEvents
                    End If
                    
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

                        Next iDateSlot
                    Next iPermutation
                Next iSheet
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
' Once file is done,
' sort winners, create Forward Compiled
            maxPermutations = UBound(param("Permutations"), 1) - 1
            For iPermutation = 2 To UBound(param("Permutations"), 1)
                permutationID = iPermutation - 1
                
                For iDateSlot = LBound(datesISOS, 2) To UBound(datesISOS, 2)
                    Application.StatusBar = sDir & sWindowSet & "Calculations: Permutation " _
                            & permutationID & "/" & maxPermutations & ", DateSlot " & iDateSlot _
                            & "/" & UBound(datesISOS, 2) & "."

                    startIS = datesISOS(1, iDateSlot)
                    endIS = datesISOS(2, iDateSlot)
                    startOS = datesISOS(3, iDateSlot)
                    endOS = datesISOS(4, iDateSlot)

                    ' Bubble-sort IS array winners
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

                    ' Bubble-sort OS array winners
                    A(permutationID)(iDateSlot)(2)(1) = BubbleSort2DArray( _
                            A(permutationID)(iDateSlot)(2)(1), True, True, True, sortColumnId, sortWs, sortC)
                    
                    ' Apply fraction multiplier to "Return" column
                    A(permutationID)(iDateSlot)(2)(1) = ApplyFractionMultiplier( _
                            A(permutationID)(iDateSlot)(2)(1), _
                            fractionMultiplier)

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

            ' Add new WB for results
            Application.StatusBar = sDir & sWindowSet & "Print, save, close..."

            ' sheets = permutations count + summary
            Call CreateNewWorkbookSheetsCountNames(targetWb, UBound(param("Permutations"), 1))

            ' Print the results
            Call PrintResultArraySaveClose
            
            ' Purge result array from memory
            Set A = Nothing

        Next iWindowSet
    Next iDir

    Application.StatusBar = False
    Application.ScreenUpdating = True

'Debug.Print Round(Timer - timer0, 2)

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
    Dim C As Range
    Application.ScreenUpdating = False
    If ActiveSheet.Name = "Summary" Then
        Set ws = ActiveSheet
        Set C = ws.Cells
        shIndex = C(ActiveCell.Row, 1).Value
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
    exportDir = GetParentDirectory(ActiveWorkbook.path) & "\tmpImgExport"
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

    exportDir = GetParentDirectory(ActiveWorkbook.path) & "\tmpImgExport"
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
    Dim C As Integer
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
        For C = LBound(print_arr, col_dim) To UBound(print_arr, col_dim)
            print_col = C + add_cols + col_offset
            If is_inverted Then
                print_cells(print_row, print_col) = print_arr(C, r)
            Else
                print_cells(print_row, print_col) = print_arr(r, C)
            End If
        Next C
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
    exportDir = GetParentDirectory(ActiveWorkbook.path) & "\tmpImgExport"
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
