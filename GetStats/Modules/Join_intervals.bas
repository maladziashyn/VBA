Attribute VB_Name = "Join_intervals"
Option Explicit

Const addInFName As String = "GetStats_BackTest_v1.24.xlsm"
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
    
    Dim i As Integer
    
    Application.ScreenUpdating = False
    Call InitPositionTags(positionTags)
    Call Init_sheet_cells
' sanity #1
    If Check_Target_Source = False Then
        MsgBox "Error. Target or source folders"
        Exit Sub
    End If
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
    Dim targetWB As Workbook
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
        Set targetWB = Workbooks.Add
        ' add sheets to targetWB
        Call Change_sheets_count(targetWB, wbs(1).Sheets.count)
        
' LOOP THROUGH ALL REPORTS
' FIND MATCHING PARAMETERS
' COPY TO TARGET BOOK
        For j = 3 To wbs(1).Sheets.count
            ' copy initial trades set to target book
            Set wsMain = wbs(1).Sheets(j)
            Set cMain = wsMain.Cells
            lastRow = cMain(wsMain.Rows.count, 3).End(xlUp).Row
            Set Rng = wsMain.Range(cMain(1, 3), cMain(lastRow, 13))
            Set wsTarget = targetWB.Sheets(j)
            Set cTarget = wsTarget.Cells
            Rng.Copy cTarget(1, 3)  ' copy trades
            lastRow = cMain(wsMain.Rows.count, 1).End(xlUp).Row
            Set Rng = wsMain.Range(cMain(23, 1), cMain(lastRow, 2))
            Call Remove_tag_from_parameters(Rng)
            Rng.Copy cTarget(23, 1) ' copy parameters
            ' move parameters to Arr
            Set Rng = wsMain.Range(cMain(23, 2), cMain(lastRow, 2))
            parMain = Parameters_to_arr(Rng, lastRow - 22)
            ' LOOP compare parMain to wsSrch / cSrch
            ' remove tags
            For k = 2 To UBound(wbs)
                For m = 3 To wbs(k).Sheets.count
                    Set wsSrch = wbs(k).Sheets(m)
                    Set cSrch = wsSrch.Cells
                    Set Rng = wsSrch.Range(cSrch(23, 1), cSrch(lastRow, 2))
                    Call Remove_tag_from_parameters(Rng)
                Next m
            Next k
            ' find matches, copy to target
            For k = 2 To UBound(wbs)
                For m = 3 To wbs(k).Sheets.count
                    Set wsSrch = wbs(k).Sheets(m)
                    Set cSrch = wsSrch.Cells
                    Set Rng = wsSrch.Range(cSrch(23, 2), cSrch(lastRow, 2))
                    parCompare = Parameters_to_arr(Rng, lastRow - 22)
                    If Parameters_Match(parMain, parCompare) Then
                        lastRMatch = cSrch(wsSrch.Rows.count, 3).End(xlUp).Row
                        Set rngMatch = wsSrch.Range(cSrch(2, 3), cSrch(lastRMatch, 13))
                        nextRTarget = cTarget(wsTarget.Rows.count, 3).End(xlUp).Row + 1
                        rngMatch.Copy cTarget(nextRTarget, 3)
                        ' fill some basic info: date from-to, trades count
                        If k = UBound(wbs) Then
                            Set rngMatch = wsSrch.Range(cSrch(1, 1), cSrch(2, 2))
                            rngMatch.Copy cTarget(1, 1)
                            Set rngMatch = wsSrch.Range(cSrch(3, 1), cSrch(22, 1))
                            rngMatch.Copy cTarget(3, 1)
                            cTarget(8, 2) = targetDateFromDt
                            cTarget(9, 2) = targetDateToDt
                            cTarget(11, 2) = cTarget(wsTarget.Rows.count, 3).End(xlUp).Row - 1
                        End If
                    End If
                Next m
            Next k
        Next j
        ' save & close all
        Application.StatusBar = appSta & " Saving target book " & i & "."
        targetWBName = Target_WB_Name(wbs(1).Name)
        targetWB.SaveAs fileName:=targetWBName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        targetWB.Close
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
    
    Const shNameOne As String = "summary"
    Const shNameTwo As String = "results"
    Dim i As Integer
    
    If someWB.Sheets.count > shCount Then
        Application.DisplayAlerts = False
        For i = 1 To someWB.Sheets.count - shCount
            someWB.Sheets(someWB.Sheets.count).Delete
        Next i
        Application.DisplayAlerts = True
    ElseIf someWB.Sheets.count < shCount Then
        For i = 1 To shCount - someWB.Sheets.count
            someWB.Sheets.Add after:=someWB.Sheets(someWB.Sheets.count)
        Next i
    End If
    someWB.Sheets(1).Name = shNameOne
    someWB.Sheets(2).Name = shNameTwo
' rename rest of sheets
    For i = 3 To someWB.Sheets.count
        someWB.Sheets(i).Name = i - 2
    Next i
    someWB.Sheets(3).Activate

End Sub

Private Sub Pick_target_folder()
' sub adds a folder path to cells(2, 1)
' in "Source folders" column (1)
    
    Dim fd As FileDialog
    
    Application.ScreenUpdating = False
    Call Init_sheet_cells
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Pick target folder"
'        .ButtonName = "OK"
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    cJ(targetFdRow, 1) = fd.SelectedItems(1)
    wsJ.Columns(1).AutoFit
    Application.ScreenUpdating = True

End Sub

Private Sub Add_source_folder()
' sub adds a folder path to next free row
' in "Source folders" column (1)
    
    Dim fd As FileDialog
    Dim nextFreeRow As Integer
    
    Application.ScreenUpdating = False
    Call Init_sheet_cells
    nextFreeRow = cJ(wsJ.Rows.count, 1).End(xlUp).Row + 1
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Pick folder with XLSX reports"
'        .ButtonName = "OK"
    End With
    If fd.Show = 0 Then
        Exit Sub
    End If
    cJ(nextFreeRow, 1) = fd.SelectedItems(1)
    wsJ.Columns(1).AutoFit
    Application.ScreenUpdating = True

End Sub

Private Sub Clear_source_list()
' sub clears processing list (subfolders)
    
    Dim Rng As Range
    
    Application.ScreenUpdating = False
    Call Init_sheet_cells
    Set Rng = wsJ.Range(cJ(sourceFdFRow, 1), cJ(wsJ.Rows.count, 1))
    Rng.Clear
    Application.ScreenUpdating = True

End Sub

Private Sub Rename_source_files_no_postfix_dates()
    
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
    lastRow = cJ(wsJ.Rows.count, 1).End(xlUp).Row
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
    
    Set wsJ = Workbooks(addInFName).Sheets(joinShName)
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
            matchPath = srcFdInfo(j, 1) & "\" & stratIns & "-" & srcFdInfo(j, 5) _
                & "-" & srcFdInfo(j, 6) & "-" & srcFdInfo(j, 4) & ".xlsx"
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
    
    sourceFdLRow = cJ(wsJ.Rows.count, 1).End(xlUp).Row
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
    
    lastRow = cJ(wsJ.Rows.count, 1).End(xlUp).Row
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
    Set oFiles = oFolder.Files
    If oFiles.count = 0 Then Exit Function
    ReDim vaArray(1 To oFiles.count)
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
    ubnd = Rng(ws.Rows.count, 3).End(xlUp).Row - 1
    If Rng(1, 15) <> "" Then
        Call GSPR_Remove_Chart2
        lr_dates = Rng(ws.Rows.count, 15).End(xlUp).Row
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
       
    chObj_idx = ActiveSheet.ChartObjects.count + 1
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
        .ChartTitle.Text = ChTitle
        .ChartTitle.Characters.Font.Size = chFontSize
        .Axes(xlCategory).TickLabelPosition = xlLow
    End With
'    sc(1, first_col + 1) = chObj_idx
    sc(1, 15).Select

End Sub

Function Get_Calendar_Days_Equity2(ByVal tset As Variant, _
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

Private Sub Print_2D_Array2(ByVal print_arr As Variant, _
        ByVal is_inverted As Boolean, _
        ByVal row_offset As Integer, _
        ByVal col_offset As Integer, _
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
    cRes(1, 1) = "#_link"
    For i = parFRow To parLRow
        cRes(1, j) = clz(i, 1)
        j = j + 1
    Next i
' copy parameters
    For i = 3 To Sheets.count
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
        wsRes.Hyperlinks.Add anchor:=cRes(j, 1), Address:="", SubAddress:="'" & repNum & "'!R22C2"
        ' print "back to summary" link
        With c(22, 2)
            .Value = "results"
            .HorizontalAlignment = xlRight
        End With
        ws.Hyperlinks.Add anchor:=c(22, 2), Address:="", SubAddress:="'results'!A1"
    Next i
    wsRes.Activate
    cRes(2, 2).Activate
    wsRes.Rows("1:1").AutoFilter
    ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = True

End Sub

Sub CalcMore()
    
    Dim ws As Worksheet
    Dim Rng As Range
    Dim ubnd As Long
    Dim tradesSet() As Variant
    Dim daysSet() As Variant
    Dim pipsRng As Range
    
    Application.ScreenUpdating = False
    Set ws = ActiveSheet
    Set Rng = ws.Cells
    ubnd = Rng(ws.Rows.count, 3).End(xlUp).Row - 1
' move to RAM
    tradesSet = Load_Slot_to_RAM2(Rng, ubnd)
' add Calendar x2 columns
    daysSet = Get_Calendar_Days_Equity2(tradesSet, Rng)
' TPM
    With Rng(3, 2)
        .Value = pmTradesPerMonth(Rng(8, 2), Rng(9, 2), Rng(11, 2))
        .NumberFormat = "0.00"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With
' AR
    With Rng(4, 2)
        .Value = pmAR(daysSet, Rng(8, 2), Rng(9, 2))
        .NumberFormat = "0.00%"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With
' MDD
    With Rng(5, 2)
        .Value = pmMDD(tradesSet)
        .NumberFormat = "0.00%"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With
' RF
    With Rng(6, 2)
        .Value = Rng(4, 2).Value / Rng(5, 2).Value
        .NumberFormat = "0.00"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With
' RSQ
    With Rng(7, 2)
        .Value = pmRSQ(daysSet)
        .NumberFormat = "0.00"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With
' months
    With Rng(10, 2)
        .Value = (Rng(9, 2) - Rng(8, 2)) * 12 / 365
        .NumberFormat = "0.00"
    End With
' win ratio
    Set pipsRng = Range(Rng(2, 8), Rng(ubnd + 1, 8))
    With Rng(12, 2)
        .Value = pmWinRatio(pipsRng, ubnd)
        .NumberFormat = "0.00%"
    End With
' pips sum
    With Rng(13, 2)
        .Value = WorksheetFunction.Sum(pipsRng)
        .NumberFormat = "0.00"
    End With
' avg W/L ratio, pips
    With Rng(14, 2)
        .Value = Abs(WorksheetFunction.AverageIf(pipsRng, ">0") / WorksheetFunction.AverageIf(pipsRng, "<=0"))
        .NumberFormat = "0.00"
    End With
' avg trade, pips
    With Rng(15, 2)
        .Value = WorksheetFunction.Average(pipsRng)
        .NumberFormat = "0.00"
    End With
' depo ini
    With Rng(16, 2)
        .Value = 10000
        .NumberFormat = "0.00"
    End With
' depo finish
    With Rng(17, 2)
        .Value = Rng(16, 2).Value * daysSet(2, UBound(daysSet, 2))
        .NumberFormat = "0.00"
    End With
    Call Calc_Sharpe_Ratio
    ActiveSheet.Range(Columns(1), Columns(2)).AutoFit
    Application.ScreenUpdating = True

End Sub

Function pmTradesPerMonth(ByRef date0 As Date, _
        ByRef date9 As Date, _
        ByRef tradeCount As Long) As Double

    pmTradesPerMonth = tradeCount / ((date9 - date0 + 1) / 30.4)

End Function

Function pmAR(ByRef daysSet As Variant, _
        ByRef date0 As Date, _
        ByRef date9 As Date) As Double
            
    Dim finalEqCurvePoint As Double

' calc net return
    finalEqCurvePoint = daysSet(2, UBound(daysSet, 2))
    pmAR = finalEqCurvePoint ^ (365 / (date9 - date0 + 1)) - 1
    
End Function

Function pmMDD(ByRef tradesSet As Variant) As Double
' tradesSet:
'   INVERTED: COLUMNS, ROWS
'   1. open date
'   2. close date
'   3. return

' create x by ubound Array
'   1. equity curve
'   2. HWM
'   3. DD
    
    Dim arr() As Variant
    Dim maxDD As Double
    Dim i As Long, j As Long
    Dim tradesCount As Long
    
    tradesCount = UBound(tradesSet, 2)
    ReDim arr(1 To 3, 1 To tradesCount + 1)
    arr(1, 1) = 1   ' starting Equity curve
    arr(2, 1) = 1   ' starting HWM
    
    maxDD = 0   ' initialize MDD
    For i = 2 To UBound(arr, 2)
        j = i - 1
        arr(1, i) = arr(1, j) * (1 + tradesSet(3, j))   ' equity curve
        arr(2, i) = WorksheetFunction.Max(arr(2, j), arr(1, i)) ' HWM
        arr(3, i) = (arr(2, i) - arr(1, i)) / arr(2, i) ' DD
        If arr(3, i) > maxDD Then
            maxDD = arr(3, i)
        End If
    Next i
    pmMDD = maxDD
    
End Function

Function pmRSQ(ByRef daysSet As Variant) As Double

    Dim x() As Variant
    Dim y() As Variant
    Dim i As Long
    
    ReDim x(1 To UBound(daysSet, 2))
    ReDim y(1 To UBound(daysSet, 2))
    For i = LBound(daysSet, 2) To UBound(daysSet, 2)
        x(i) = i
        y(i) = daysSet(2, i)
    Next i
    pmRSQ = WorksheetFunction.RSq(x, y)
    
End Function

Function pmWinRatio(ByRef Rng As Range, ByRef tradesCount As Long) As Double

    Dim winners As Long
    Dim c As Range
    
    winners = 0
    For Each c In Rng
        If c.Value > 0 Then
            winners = winners + 1
        End If
    Next c
    pmWinRatio = winners / tradesCount
    
End Function


