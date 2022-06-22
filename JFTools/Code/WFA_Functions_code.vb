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
    
    ' If account goes below zero DUE TO ERROR ON BACK-TEST
    ' This will essentially eliminate bad back-tests from WFA,
    ' but will not help spot & correct the error on back-test.
    If dailyEq(UBound(dailyEq)) > 0 Then
        resultDict("Annualized Return") = dailyEq(UBound(dailyEq)) ^ (365 / (UBound(calDays) - 1)) - 1
    Else
        resultDict("Annualized Return") = -0.99999
    End If
    
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
