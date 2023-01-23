Attribute VB_Name = "GlobalsConstantsHelpers"
Option Explicit

Const MsgBoxTitle As String = "JFTools"

Sub OnStart()

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
        .Cursor = xlWait
    End With

End Sub

Sub OnExit()

    Dim Message As String

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .StatusBar = False
        .DisplayAlerts = True
        .Cursor = xlDefault
    End With

    If Err.Number > 0 Then
        Message = Message & "Error " & Err.Number & ". " _
            & vbNewLine & Err.Description
        MsgBox Message, vbCritical, MsgBoxTitle
        Err.Clear
    End If

End Sub

