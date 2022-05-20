Attribute VB_Name = "Tools"
Option Explicit

Dim current_decimal As String
Dim undo_sep As Boolean, undo_usesyst As Boolean
Dim user_switched As Boolean
    
Sub GSPR_Remove_CommandBar()
    
    On Error Resume Next
    
    Application.CommandBars("GSPR-1").Delete
    Application.CommandBars("GSPR-2").Delete
    Application.CommandBars("GSPR-3").Delete
    Application.CommandBars("GSPR-4").Delete
    Application.CommandBars("GSPR-5").Delete
    Application.CommandBars("GSPR-6").Delete
    Application.CommandBars("GSPR-7").Delete
    Application.CommandBars("GSPR-8").Delete
    Application.CommandBars("GSPR-9").Delete
    Application.CommandBars("GSPR-10").Delete

End Sub

Sub GSPR_Create_CommandBar()
    
    Dim cBar1 As CommandBar
    Dim cBar2 As CommandBar
    Dim cBar3 As CommandBar
    Dim cBar4 As CommandBar
    Dim cBar5 As CommandBar
    Dim cBar6 As CommandBar
    Dim cBar7 As CommandBar
    Dim cBar8 As CommandBar
    Dim cBar9 As CommandBar
    Dim cBar10 As CommandBar
    Dim cControl As CommandBarControl
    
    Call GSPR_Remove_CommandBar
' Create toolbar
    Set cBar1 = Application.CommandBars.Add
    cBar1.Name = "GSPR-1"
    cBar1.Visible = True
' Create toolbar 2
    Set cBar2 = Application.CommandBars.Add
    cBar2.Name = "GSPR-2"
    cBar2.Visible = True
' Create toolbar 3
    Set cBar3 = Application.CommandBars.Add
    cBar3.Name = "GSPR-3"
    cBar3.Visible = True
' Create toolbar 4
    Set cBar4 = Application.CommandBars.Add
    cBar4.Name = "GSPR-4"
    cBar4.Visible = True
' Create toolbar 5
    Set cBar5 = Application.CommandBars.Add
    cBar5.Name = "GSPR-5"
    cBar5.Visible = True
' Create toolbar 6
    Set cBar6 = Application.CommandBars.Add
    cBar6.Name = "GSPR-6"
    cBar6.Visible = True
' Create toolbar 7
    Set cBar7 = Application.CommandBars.Add
    cBar7.Name = "GSPR-7"
    cBar7.Visible = True
' Create toolbar 8
    Set cBar8 = Application.CommandBars.Add
    cBar8.Name = "GSPR-8"
    cBar8.Visible = True
' Create toolbar 9
    Set cBar9 = Application.CommandBars.Add
    cBar9.Name = "GSPR-9"
    cBar9.Visible = True
' Create toolbar 10
    Set cBar10 = Application.CommandBars.Add
    cBar10.Name = "GSPR-10"
    cBar10.Visible = True

' ROW 1 ===========

    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 351
        .OnAction = "GSPR_Single_Core"
        .TooltipText = "Main report"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Main"
    End With
    
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 352 ' 352   126, 59, 630
        .OnAction = "GSPR_Single_Extra"
        .TooltipText = "Super duper extra report"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Extra"
    End With

' ROW 2 ===========

    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 688
        .OnAction = "GSPRM_Multiple_Main"
        .TooltipText = "Process a group of reports"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Group"
    End With

    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 418
        .OnAction = "GSPR_Build_Charts_Singe_Button"
        .TooltipText = "Build chart"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Chart"
    End With

' ROW 3 ===========
    
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 585
        .OnAction = "GSPR_Copy_Sheet_Next"
        .TooltipText = "Copy sheet"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "CopySh"
    End With
    
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 478
        .OnAction = "GSPR_Delete_Sheet"
        .TooltipText = "Delete active sheet"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "DelSh"
    End With
    
    ' SEPARATOR
    
    user_switched = False
    
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 98
        .OnAction = "GSPR_Separator_Manual_Switch"
        .TooltipText = "Change decimal separator"
        .Control.Style = msoButtonIconAndCaption
'        .Caption = "Separator"
    End With

' ROW 4 ===========
    
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 688
        .OnAction = "GSPRM_Merge_Summaries"
        .TooltipText = "Merge on recovery factor"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Recovery"
    End With
    
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 1576
        .OnAction = "GSPR_Change_Folder_Link"
        .TooltipText = "Refresh hyperlinks"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "HLinks"
    End With
    
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 279    ' 31, 279
        .OnAction = "GSPR_Mixer_Copy_Sheet_To_Book"
        .TooltipText = "Add this sheet to 'mixer'"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ToMix"
    End With

' ROW 5 ===========
    
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 124
        .OnAction = "GSPR_show_sheet_index"
        .TooltipText = "Show this sheet's index"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ShIndex"
    End With
    
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 205
        .OnAction = "GSPR_Go_to_sheet_index"
        .TooltipText = "Go to sheet with your index"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ToIndex"
    End With
    
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 645   ' 601
        .OnAction = "GSPR_robo_mixer"
        .TooltipText = "Magic - make the MIX"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "MIX"
    End With

' ROW 6 ===========
    
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 424   ' 601
        .OnAction = "GSPR_trades_to_days"
        .TooltipText = "Mix chart on calendar days"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "MixChart"
    End With
    
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 601   ' 601
        .OnAction = "Check_Window_Bulk"
        .TooltipText = "Check errors"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "CheckErrs"
    End With
    
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 28    ' 7, 28, 159, 176
        .OnAction = "Create_JFX_file_Main"
        .TooltipText = "Create code snippet for JFX"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "JFX"
    End With

' ROW 7 ===========

    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 7
        .OnAction = "Settings_To_Launch_Log"
        .TooltipText = "Настройки робота из java в журнал"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "java-log"
    End With
    
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 424   ' 601
        .OnAction = "Stats_Chart_from_Joined_Windows"
        .TooltipText = "Chart for joined windows"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ChartJ"
    End With
    
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 435   ' 601
        .OnAction = "Calc_Sharpe_Ratio"
        .TooltipText = "Calculate Sharpe ratio for single sheet"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Sharpe"
    End With
    
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 191
        .OnAction = "Params_To_Summary"
        .TooltipText = "Retrieve parameters/values to summary sheet"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ParamJ-Summary"
    End With
    
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 477   ' 601
        .OnAction = "Sharpe_to_all"
        .TooltipText = "Calculate Sharpe ratio on all sheets"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Sharpe all"
    End With
    
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 430
        .OnAction = "Scatter_Sharpe"
        .TooltipText = "Build scatter plots based on Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ScatterPlots"
    End With
    
    Set cControl = cBar9.Controls.Add
    With cControl
        .FaceId = 478
        .OnAction = "RemoveScatters"
        .TooltipText = "Remove all scatter plots"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "DelScatter"
    End With
    
    Set cControl = cBar9.Controls.Add
    With cControl
        .FaceId = 477
        .OnAction = "GSPRM_Merge_Sharpe"
        .TooltipText = "Merge reports on Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "SharpeMerge"
    End With
    
    Set cControl = cBar9.Controls.Add
    With cControl
        .FaceId = 283
        .OnAction = "CalcMore"
        .TooltipText = "Calculate rest of KPI"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "CalcMore"
    End With
    
    Set cControl = cBar10.Controls.Add
    With cControl
        .FaceId = 143
        .OnAction = "SharpePivot"
        .TooltipText = "Merge summaries, calculate Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "SharpePvt"
    End With

End Sub

Private Sub GSPR_About_Link()
' unused
    
    ActiveWorkbook.FollowHyperlink Address:="https://vsatrader.ru/getstats/"

End Sub

Private Sub GSPR_Copy_Sheet_Next()
    
    Application.ScreenUpdating = False
    ActiveSheet.Copy after:=Sheets(ActiveSheet.Index)
    Application.ScreenUpdating = True

End Sub

Private Sub GSPR_Delete_Sheet()
    
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True

End Sub

Private Sub GSPR_Separator_Manual_Switch()
    
    Application.ScreenUpdating = False
    If user_switched Then
        Call GSPR_Separator_OFF
    Else
        Call GSPR_Separator_ON
    End If
    Application.ScreenUpdating = True

End Sub

Private Sub GSPR_Separator_ON()
    
    undo_sep = False
    undo_usesyst = False
    
    Const msg_rec As String = "Setting DOT as decimal separator." & vbNewLine & vbNewLine _
                & "Recommended for running GetStats." & vbNewLine & vbNewLine _
                & "To switch back to your separator press ""S"" again."
    Const msg_not As String = "Your separator is DOT. Optimal for GetStats."
    
    If Application.UseSystemSeparators Then     ' SYS - ON
        If Not Application.International(xlDecimalSeparator) = "." Then
            Application.UseSystemSeparators = False
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            MsgBox msg_rec, , "GetStats"
            undo_sep = True                     ' undo condition 1
            undo_usesyst = True                 ' undo condition 2
            user_switched = True
        Else
            undo_sep = False
            MsgBox msg_not, , "GetStats"
        End If
    Else                                        ' SYS - OFF
        If Not Application.DecimalSeparator = "." Then
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            MsgBox msg_rec, , "GetStats"
            undo_sep = True                     ' undo condition 1
            undo_usesyst = False                ' undo condition 2
            user_switched = True
        Else
            undo_sep = False
            MsgBox msg_not, , "GetStats"
        End If
    End If

End Sub

Private Sub GSPR_Separator_OFF()
    
    Const msg_not As String = "Separator not changed. Current: DOT. Optimal for GetStats."
    
    If undo_sep Then
        Application.DecimalSeparator = current_decimal
        If undo_usesyst Then
            Application.UseSystemSeparators = True
        End If
        MsgBox "Reverting back to user's decimal separator." & vbNewLine & vbNewLine _
        & "To switch back to recommended separator (DOT), press ""S"" again.", , "GetStats"
        user_switched = False
    Else
        MsgBox msg_not, , "GetStats"
    End If

End Sub
