Option Explicit
Private Sub Workbook_Open()
    Call CreateCommandBars
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call RemoveCommandBars
End Sub
