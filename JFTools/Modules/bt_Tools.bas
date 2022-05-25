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
        .TooltipText = "Обработать отчет и вывести основную статистику"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Основной"
    End With
' Single report - extra
    Set cControl = cBar1.Controls.Add
    With cControl
        .FaceId = 352 ' 352   126, 59, 630
        .OnAction = "GSPR_Single_Extra"
        .TooltipText = "Получить полную статистику по одному отчету"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Экстра"
    End With
' ROW 2 ===========
' Multiple
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 688
        .OnAction = "GSPRM_Multiple_Main"
        .TooltipText = "Обработать группу отчетов (выбрать из папки)"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Группа"
    End With
' Build charts, single core
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 418
        .OnAction = "GSPR_Build_Charts_Singe_Button"
        .TooltipText = "Построить/удалить графики к отчету типа ""основной"""
        .Control.Style = msoButtonIconAndCaption
        .Caption = "График"
    End With
' Build charts, single core
    Set cControl = cBar2.Controls.Add
    With cControl
        .FaceId = 418
        .OnAction = "GSPR_Build_Charts_Singe_Button_EN"
        .TooltipText = "Построить/удалить графики к отчету типа ""основной"""
        .Control.Style = msoButtonIconAndCaption
        .Caption = "EN"
    End With
' ROW 3 ===========
' Copy sheet
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 585
        .OnAction = "GSPR_Copy_Sheet_Next"
        .TooltipText = "Скопировать лист и вставить рядом"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Копия"
    End With
' Delete sheet
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 478
        .OnAction = "GSPR_Delete_Sheet"
        .TooltipText = "Удалить активный лист без возможности восстановления"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Удалить"
    End With
' Separator
    user_switched = False       ' SEPARATOR
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 98
        .OnAction = "GSPR_Separator_Manual_Switch"
        .TooltipText = "Установить рекомендуемый разделитель (точку) или вернуть пользовательские настройки"
        .Control.Style = msoButtonIconAndCaption
'        .Caption = "Разделитель"
    End With
' GSen_Translate
    user_switched = False       ' SEPARATOR
    Set cControl = cBar3.Controls.Add
    With cControl
        .FaceId = 84
        .OnAction = "GSPR_EN_Translate"
        .TooltipText = "Translate into EN"
        .Control.Style = msoButtonIconAndCaption
'        .Caption = "Разделитель"
    End With
' ROW 4 ===========
' GSPRM_Merge_Summaries
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 688
        .OnAction = "GSPRM_Merge_Summaries"
        .TooltipText = "Объединить листы ""результаты"" в одну книгу (merged)"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Recovery"
    End With
' GSPR_Change_Folder_Link
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 1576
        .OnAction = "GSPR_Change_Folder_Link"
        .TooltipText = "Обновить гиперссылки на отчеты, если их папка была перемещена или переименована"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Ссылки"
    End With
' GSPR_Mixer_Copy_Sheet_To_Book
    Set cControl = cBar4.Controls.Add
    With cControl
        .FaceId = 279    ' 31, 279
        .OnAction = "GSPR_Mixer_Copy_Sheet_To_Book"
        .TooltipText = "Добавить лист в книгу 'mixer'"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "В микс"
    End With
' ROW 5 ===========
'' About
'    Set cControl = cBar5.Controls.Add
'    With cControl
'        .FaceId = 487
'        .OnAction = "GSPR_About_Link"
'        .TooltipText = "Справка о программе ""GetStats Pro"" на VSAtrader.ru"
'        .Control.Style = msoButtonIconAndCaption
'        .Caption = "О GetStats"
'    End With
' Показать индекс листа
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 124
        .OnAction = "GSPR_show_sheet_index"
        .TooltipText = "Индекс листа"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Индекс"
    End With
' Перейти к листу по номеру
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 205
        .OnAction = "GSPR_Go_to_sheet_index"
        .TooltipText = "Перейти на лист по номеру"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "К листу"
    End With
' GSPR_robo_mixer
    Set cControl = cBar5.Controls.Add
    With cControl
        .FaceId = 645   ' 601
        .OnAction = "GSPR_robo_mixer"
        .TooltipText = "Объединить списки сделок и выдать статистику"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "МИКС"
    End With
' ROW 6 ===========
' GSPR_trades_to_days
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 424   ' 601
        .OnAction = "GSPR_trades_to_days"
        .TooltipText = "График эквити по объединенному списку сделок"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "График М"
    End With
' Check_Window
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 601   ' 601
        .OnAction = "Check_Window_Bulk"
        .TooltipText = "Проверка окон, счета, кол-ва html"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Проверка"
    End With
' Create_JFX_file_Main
    Set cControl = cBar6.Controls.Add
    With cControl
        .FaceId = 28    ' 7, 28, 159, 176
        .OnAction = "Create_JFX_file_Main"
        .TooltipText = "Создать JFX"
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
        .TooltipText = "Настройки робота из java в журнал"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "java-log"
    End With
' GSPR_trades_to_days
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 424   ' 601
        .OnAction = "Stats_Chart_from_Joined_Windows"
        .TooltipText = "График эквити по объединенным окнам"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "График J"
    End With
' Calc_Sharpe_Ratio
    Set cControl = cBar7.Controls.Add
    With cControl
        .FaceId = 435   ' 601
        .OnAction = "Calc_Sharpe_Ratio"
        .TooltipText = "Коэффициент Шарпа"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Sharpe"
    End With
' Params_To_Summary
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 191
        .OnAction = "Params_To_Summary"
        .TooltipText = "Параметры Joined в Результаты"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ParamJ-Summary"
    End With
' Sharpe_to_all
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 477   ' 601
        .OnAction = "Sharpe_to_all"
        .TooltipText = "Коэффициент Шарпа, вся книга"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Sharpe all"
    End With
' Scatter_Sharpe
    Set cControl = cBar8.Controls.Add
    With cControl
        .FaceId = 430
        .OnAction = "Scatter_Sharpe"
        .TooltipText = "Параметры и Sharpe"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "ScatterPlots"
    End With
' RemoveScatters
    Set cControl = cBar9.Controls.Add
    With cControl
        .FaceId = 478
        .OnAction = "RemoveScatters"
        .TooltipText = "Удалить все графики"
        .Control.Style = msoButtonIconAndCaption
        .Caption = "Удал. граф."
    End With
' GSPRM_Merge_Sharpe
    Set cControl = cBar9.Controls.Add
    With cControl
        .FaceId = 477
        .OnAction = "GSPRM_Merge_Sharpe"
        .TooltipText = "Объединить по Sharpe"
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
' RIBBON > BUTTON "Копия"
'
    Application.ScreenUpdating = False
    ActiveSheet.Copy after:=Sheets(ActiveSheet.Index)
    Application.ScreenUpdating = True
End Sub
Private Sub GSPR_Delete_Sheet()
'
' RIBBON > BUTTON "Удалить"
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
    Const msg_rec As String = "Устанавливаю точку как разделитель целой и дробной части." & vbNewLine & vbNewLine _
                & "Правильный выбор!" & vbNewLine & vbNewLine _
                & "Рекомендовано для быстрой работы GetStats." & vbNewLine & vbNewLine _
                & "Чтобы вернуть ваш разделитель, нажмите ""S"" еще раз."
    Const msg_not As String = "Ваш разделитель - точка. Оставляем, как есть. Это оптимальное решение для быстрой работы GetStats."
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
    Const msg_not As String = "Ваш разделитель не был изменен. Сейчас это точка. Оставляем, как есть. Это оптимальное решение для быстрой работы GetStats."
    If undo_sep Then
        Application.DecimalSeparator = current_decimal
        If undo_usesyst Then
            Application.UseSystemSeparators = True
        End If
        MsgBox "Возвращаю пользовательский разделитель целой и дробной части." & vbNewLine & vbNewLine _
        & "Чтобы задать рекомендованный разделитель (точку) и ускорить работу программы, снова нажмите ""S"".", , "GetStats Pro"
        user_switched = False
    Else
        MsgBox msg_not, , "GetStats Pro"
    End If
End Sub
