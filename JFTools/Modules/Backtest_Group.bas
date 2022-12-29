Attribute VB_Name = "Backtest_Group"
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
