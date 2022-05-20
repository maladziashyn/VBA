Option Explicit
Private Sub Workbook_Open()
    Call GSPR_Create_CommandBar
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call GSPR_Remove_CommandBar
End Sub
