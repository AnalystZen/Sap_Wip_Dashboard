Attribute VB_Name = "ModActionManager"
Option Private Module

Sub SpGetData()
'// This sub will let the user select the date of data they would like to import.
'// Created by AL on 11/09/2024.

    '// Declare variables.
    Dim MsgResult As VbMsgBoxResult
    Dim DateEntry As Date
    
    '// Assign value.
    DateEntry = Range("DateEntry").Value
    MsgResult = MsgBox("Would you like to get the production information for the selected date?" _
                , vbExclamation + vbYesNo _
                , "Import Production Data")
      
    '// Conditional check for procedure to run.
    If DateEntry = Empty Then
        MsgBox Prompt:="Please enter the date." _
        , Buttons:=vbOKCancel + vbExclamation _
        , Title:="Enter Date"
        Exit Sub
    End If
    
    '// Handle errors.
    On Error GoTo ErrHandler
    
    '// Turn off screen update.
    TurnOffApps
    
    '// Exit or run sub based on user.
    Select Case MsgResult
        Case vbYes
            SpChainAllImports
            SpDataTable
        Case vbNo
            Exit Sub
    End Select
    
    '// Unprotect dahboard for pivot chart updates.
    ShDashboard.Unprotect LCase("control")
    
    '// Update all pivot tables in workbook and refresh chart visuals.
    Application.ScreenUpdating = True
    ThisWorkbook.RefreshAll
    
    '// Protect dashboard after refresh update.
    ShDashboard.Protect Password:="control", AllowUsingPivotTables:=True
    
    '// Turn on screen update.
    TurnOnApps
    
    '// User update of import success.
    MsgBox "Data was imported successfully." _
            , vbInformation _
            , "Success"
    
    '// Clean exit
    Exit Sub

ErrHandler:
    '// hide sheets.
    With ShUsage
        .Visible = False
    End With
    
    '// Go to table.
    ShTable.Select

    '// User update of failure.
    MsgBox "Something went wrong. Verify a session of SAP is open and a valid date has been input." _
            , vbCritical _
            , "Data Import Failed"
            
    '// Turn on screen update.
    TurnOnApps

End Sub

