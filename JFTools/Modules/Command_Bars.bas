Attribute VB_Name = "Command_Bars"
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
