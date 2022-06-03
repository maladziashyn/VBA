Sub Snippets()

' Insert a column before colNumber
    Columns(colNumber).Insert

' Apply autofilter to rowNumber
    Rows(rowNumber).AutoFilter

' Set column width
    Columns(colNumber).ColumnWidth = 10

' Remove cell color
    With ActiveSheet.Columns("A:A").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

' Enable error handling, go to the following statement
    On Error Resume Next

' Disable error handling
    On Error GoTo 0

' Copy worksheet to another workbook
    shToCopy.Copy after:=wbToInsert.Sheets(wbToInsert.Worksheets.Count)

' Save as XLSB
    ActiveWorkbook.SaveAs fileName:=yourFileName, FileFormat:=xlExcel12

' Add hyperlink to cell
    ActiveSheet.Hyperlinks.Add anchor:=yourCell, Address:="", SubAddress:="Sheet1!A1"
    ' , TextToDisplay:="MyHyperlink"

' Make a directory
    MkDir yourPath

' Find row or column
    f = yourSheet.Cells.Find(what:="FindString", _
        after:=yourSheet.Cells(1, 1), _
        LookIn:=xlValues, _
        lookat:=xlWhole, _
        searchorder:=xlByRows, _
        searchdirection:=xlNext, _
        MatchCase:=False, _
        searchformat:=False).Column

' A simpler find
    f = Cells.Find(what:="what").Column

' Text property
    Set c = Worksheets(1).Range("B14")
    c.Value = 1198.3
    c.NumberFormat = "$#,##0_);($#,##0)"
    MsgBox c.Value
    MsgBox c.Text

' Count cells in range
    MsgBox Range("A1:C3").Count

' Get range address
    MsgBox Range(Cells(1, 1), Cells(5, 5)).Address

End Sub
