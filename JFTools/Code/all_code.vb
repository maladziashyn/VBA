

' MODULE: ThisWorkbook
Option Explicit

Private Sub Workbook_Open()
    
    Call CreateCommandBars

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    Call RemoveCommandBars

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
Const addinVersion As String = "0.23"

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


' MODULE: Sheet100



' MODULE: WFA_Main
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


' MODULE: Sheet5


' MODULE: Sheet2



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
