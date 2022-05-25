Attribute VB_Name = "bt_Tools"
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
' ROW 1 ===========
' Single report - core
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 351
        .OnAction = "GSPR_Single_Core"
        .TooltipText = "���������� ����� � ������� �������� ����������"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "��������"
    End With
' Single report - extra
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 352 ' 352   126, 59, 630
        .OnAction = "GSPR_Single_Extra"
        .TooltipText = "�������� ������ ���������� �� ������ ������"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "������"
    End With
' ROW 2 ===========
' Multiple
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 688
        .OnAction = "GSPRM_Multiple_Main"
        .TooltipText = "���������� ������ ������� (������� �� �����)"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "������"
    End With
' Build charts, single core
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 418
        .OnAction = "GSPR_Build_Charts_Singe_Button"
        .TooltipText = "���������/������� ������� � ������ ���� ""��������"""
        .Control.Style = msoButtonIconAndCaption
        .Caption = "������"
    End With
' Build charts, single core
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 418
        .OnAction = "GSPR_Build_Charts_Singe_Button_EN"
        .TooltipText = "���������/������� ������� � ������ ���� ""��������"""
        .Control.Style = msoButtonIconAndCaption
        .Caption = "EN"
    End With
' ROW 3 ===========
' Copy sheet
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 585
        .OnAction = "GSPR_Copy_Sheet_Next"
        .TooltipText = "����������� ���� � �������� �����"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "�����"
    End With
' Delete sheet
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 478
        .OnAction = "GSPR_Delete_Sheet"
        .TooltipText = "������� �������� ���� ��� ����������� ��������������"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "�������"
    End With
' Separator
    user_switched = False       ' SEPARATOR
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 98
        .OnAction = "GSPR_Separator_Manual_Switch"
        .TooltipText = "���������� ������������� ����������� (�����) ��� ������� ���������������� ���������"
        .Control.Style = msoButtonIconAndCaption
'        .Caption = "�����������"
    End With
' GSen_Translate
    user_switched = False       ' SEPARATOR
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 84
        .OnAction = "GSPR_EN_Translate"
        .TooltipText = "Translate into EN"
        .Control.Style = msoButtonIconAndCaption
'        .Caption = "�����������"
    End With
' ROW 4 ===========
' GSPRM_Merge_Summaries
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 688
        .OnAction = "GSPRM_Merge_Summaries"
        .TooltipText = "���������� ����� ""����������"" � ���� ����� (merged)"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Recovery"
    End With
' GSPR_Change_Folder_Link
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 1576
        .OnAction = "GSPR_Change_Folder_Link"
        .TooltipText = "�������� ����������� �� ������, ���� �� ����� ���� ���������� ��� �������������"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "������"
    End With
' GSPR_Mixer_Copy_Sheet_To_Book
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 279    ' 31, 279
        .OnAction = "GSPR_Mixer_Copy_Sheet_To_Book"
        .TooltipText = "�������� ���� � ����� 'mixer'"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "� ����"
    End With
' ROW 5 ===========
'' About
'    Set cControl = cBar5.Controls.Add
'    With cControl
'        .FaceId = 487
'        .OnAction = "GSPR_About_Link"
'        .TooltipText = "������� � ��������� ""GetStats Pro"" �� VSAtrader.ru"
'        .Control.Style = msoButtonIconAndCaption
'        .Caption = "� GetStats"
'    End With
' �������� ������ �����
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 124
        .OnAction = "GSPR_show_sheet_index"
        .TooltipText = "������ �����"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "������"
    End With
' ������� � ����� �� ������
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 205
        .OnAction = "GSPR_Go_to_sheet_index"
        .TooltipText = "������� �� ���� �� ������"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "� �����"
    End With
' GSPR_robo_mixer
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 645   ' 601
        .OnAction = "GSPR_robo_mixer"
        .TooltipText = "���������� ������ ������ � ������ ����������"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "����"
    End With
' ROW 6 ===========
' GSPR_trades_to_days
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 424   ' 601
        .OnAction = "GSPR_trades_to_days"
        .TooltipText = "������ ������ �� ������������� ������ ������"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "������ �"
    End With
' Check_Window
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 601   ' 601
        .OnAction = "Check_Window_Bulk"
        .TooltipText = "�������� ����, �����, ���-�� html"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "��������"
    End With
' Create_JFX_file_Main
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 28    ' 7, 28, 159, 176
        .OnAction = "Create_JFX_file_Main"
        .TooltipText = "������� JFX"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "JFX"
    End With
'' Stats_And_Chart - WFA
'    Set cControl = cBar6.Controls.Add
'    With cControl
'        .FaceId = 424
'        .OnAction = "Stats_And_Chart"
'        .TooltipText = "WFA stats&chart"
'        .Control.Style = msoButtonIconAndCaption
'        .Caption = "wfa"
'    End With
' ROW 7 ===========

' Settings_To_Launch_Log
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 7
        .OnAction = "Settings_To_Launch_Log"
        .TooltipText = "��������� ������ �� java � ������"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "java-log"
    End With
' GSPR_trades_to_days
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 424   ' 601
        .OnAction = "Stats_Chart_from_Joined_Windows"
        .TooltipText = "������ ������ �� ������������ �����"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "������ J"
    End With
' Calc_Sharpe_Ratio
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 435   ' 601
        .OnAction = "Calc_Sharpe_Ratio"
        .TooltipText = "����������� �����"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Sharpe"
    End With
' Params_To_Summary
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 191
        .OnAction = "Params_To_Summary"
        .TooltipText = "��������� Joined � ����������"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ParamJ-Summary"
    End With
' Sharpe_to_all
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 477   ' 601
        .OnAction = "Sharpe_to_all"
        .TooltipText = "����������� �����, ��� �����"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Sharpe all"
    End With
' Scatter_Sharpe
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 430
        .OnAction = "Scatter_Sharpe"
        .TooltipText = "��������� � Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ScatterPlots"
    End With
' RemoveScatters
    Set cControl = cBar9.Controls.Add
    With cControl
        .FaceId = 478
        .OnAction = "RemoveScatters"
        .TooltipText = "������� ��� �������"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "����. ����."
    End With
' GSPRM_Merge_Sharpe
    Set cControl = cBar9.Controls.Add
    With cControl
        .FaceId = 477
        .OnAction = "GSPRM_Merge_Sharpe"
        .TooltipText = "���������� �� Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "SharpeMerge"
    End With
End Sub
Private Sub GSPR_About_Link()
' unused
    ActiveWorkbook.FollowHyperlink Address:="https://vsatrader.ru/getstats/"
End Sub
Private Sub GSPR_Copy_Sheet_Next()
'
' RIBBON > BUTTON "�����"
'
    Application.ScreenUpdating = False
    ActiveSheet.Copy after:=Sheets(ActiveSheet.Index)
    Application.ScreenUpdating = True
End Sub
Private Sub GSPR_Delete_Sheet()
'
' RIBBON > BUTTON "�������"
'
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End Sub
Private Sub GSPR_Separator_Manual_Switch()
'
' RIBBON > BUTTON "S"
'
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
    Const msg_rec As String = "������������ ����� ��� ����������� ����� � ������� �����." & vbNewLine & vbNewLine _
                & "���������� �����!" & vbNewLine & vbNewLine _
                & "������������� ��� ������� ������ GetStats." & vbNewLine & vbNewLine _
                & "����� ������� ��� �����������, ������� ""S"" ��� ���."
    Const msg_not As String = "��� ����������� - �����. ���������, ��� ����. ��� ����������� ������� ��� ������� ������ GetStats."
    If Application.UseSystemSeparators Then     ' SYS - ON
        If Not Application.International(xlDecimalSeparator) = "." Then
            Application.UseSystemSeparators = False
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            MsgBox msg_rec, , "GetStats Pro"
            undo_sep = True                     ' undo condition 1
            undo_usesyst = True                 ' undo condition 2
            user_switched = True
        Else
            undo_sep = False
            MsgBox msg_not, , "GetStats Pro"
        End If
    Else                                        ' SYS - OFF
        If Not Application.DecimalSeparator = "." Then
            current_decimal = Application.DecimalSeparator
            Application.DecimalSeparator = "."
            MsgBox msg_rec, , "GetStats Pro"
            undo_sep = True                     ' undo condition 1
            undo_usesyst = False                ' undo condition 2
            user_switched = True
        Else
            undo_sep = False
            MsgBox msg_not, , "GetStats Pro"
        End If
    End If
End Sub
Private Sub GSPR_Separator_OFF()
    Const msg_not As String = "��� ����������� �� ��� �������. ������ ��� �����. ���������, ��� ����. ��� ����������� ������� ��� ������� ������ GetStats."
    If undo_sep Then
        Application.DecimalSeparator = current_decimal
        If undo_usesyst Then
            Application.UseSystemSeparators = True
        End If
        MsgBox "��������� ���������������� ����������� ����� � ������� �����." & vbNewLine & vbNewLine _
        & "����� ������ ��������������� ����������� (�����) � �������� ������ ���������, ����� ������� ""S"".", , "GetStats Pro"
        user_switched = False
    Else
        MsgBox msg_not, , "GetStats Pro"
    End If
End Sub
