Attribute VB_Name = "ModAppSettings"
Option Private Module

Sub TurnOffApps()
'// This sub will turn off apps while a procedure is running.
'// Created by AL on 11/8/2024.
   
    With ThisWorkbook
        .Application.ScreenUpdating = False
        .Application.DisplayAlerts = False
        .Application.EnableEvents = False
        .ActiveSheet.DisplayPageBreaks = False
    End With
    
End Sub

Sub TurnOnApps()
'// This sub will turn on apps afetr a procedure is done running.
'// Created by AL on 11/8/2024.

    With ThisWorkbook
        .Application.ScreenUpdating = True
        .Application.DisplayAlerts = True
        .Application.EnableEvents = True
    End With
    
End Sub


