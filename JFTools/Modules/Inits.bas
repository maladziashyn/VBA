Attribute VB_Name = "Inits"
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
