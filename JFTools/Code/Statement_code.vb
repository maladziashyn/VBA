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
