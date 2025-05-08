Attribute VB_Name = "ModAddDataToTable"
Option Private Module

Sub SpDataTable()
'// This procedure will collect data that was imported and add it to the data table.
'// Created by AL on 11/10/2024.

    '// Declare variables.
    Dim WbWip As Workbook
    Dim ShData As Worksheet
    Dim RngToAdd As Range
    Dim RngtoCalclulate As Range
    
    '// Assign values.
    Set WbWip = ThisWorkbook
    Set ShData = ShTable
    Set RngToAdd = ShData.ListObjects("WipTable").DataBodyRange(, EdaDate).End(xlDown).Offset(1)
    Set RngtoCalclulate = ShData.Cells(Rows.Count, EdaDate).End(xlUp)
    DateEntry = Range("DateEntry").Value
    
    '// Start adding imported data from coid sheet to table.
    With ShCoid
        .Activate
        .Range(Cells(EimCoid, EcoStart), Cells(Rows.Count, EcoStart).End(xlUp)).Copy RngToAdd
        .Range(Cells(EimCoid, EcoOrder), Cells(Rows.Count, EcoOrder).End(xlUp)).Copy RngToAdd.Offset(, EdaDate)
        .Range(Cells(EimCoid, EcoMaterial), Cells(Rows.Count, EcoMaterial).End(xlUp)).Copy RngToAdd.Offset(, EdaProcessOrder)
        .Range(Cells(EimCoid, EcoDescription), Cells(Rows.Count, EcoDescription).End(xlUp)).Copy RngToAdd.Offset(, EdaMaterialNumber)
        .Range(Cells(EimCoid, EcoBatch), Cells(Rows.Count, EcoBatch).End(xlUp)).Copy RngToAdd.Offset(, EdaMaterialDescription)
        .Range(Cells(EimCoid, EcoTarget), Cells(Rows.Count, EcoTarget).End(xlUp)).Copy RngToAdd.Offset(, EdaBatchNumber)
        .Range(Cells(EimCoid, EcoDelivered), Cells(Rows.Count, EcoDelivered).End(xlUp)).Copy RngToAdd.Offset(, EdaTgtQuantity)
        .Range(Cells(EimCoid, EcoConfirmed), Cells(Rows.Count, EcoConfirmed).End(xlUp)).Copy RngToAdd.Offset(, EdaDeliveredQuantity)
    End With

    '// Select table sheet.
    ShData.Activate
    
    '// Assign variables for loop.
    Dim RngTgtMix As Range
    Dim RngTgtCalculation As Range
    Dim RowAdvance As Long
    
    '// Set variables for quantity withdrawn loop.
    Set RngTgtMix = Range(RngToAdd.Offset(, EdaConfirmedQuantity), RngToAdd.Offset(, EdaConfirmedQuantity).End(xlDown))
    X = 0
     
    '// Sumif caclulation for table quantity withdrawn data.
    For Each RngTgtCalculation In RngTgtMix
        RngTgtCalculation = WorksheetFunction.SumIf(ShUsage.Range("D:D"), RngToAdd.Offset(X, EdaMaterialDescription), ShUsage.Range("$H:$H"))
        If Not IsNumeric(RngToAdd.Offset(X, EdaMaterialDescription).Value) Then RngTgtCalculation = 0
         X = X + 1
    Next RngTgtCalculation
    
    '// Set variables for target mix loop
    Set RngTgtMix = Range(RngToAdd.Offset(, EdaWithdrawnQuantity), RngToAdd.Offset(, EdaWithdrawnQuantity).End(xlDown))
    X = 0
     
    '// Sumif caclulation for table target mix data.
    For Each RngTgtCalculation In RngTgtMix
        RngTgtCalculation = WorksheetFunction.SumIf(ShMixes.Range("B:B"), RngToAdd.Offset(X, EdaDate), ShMixes.Range("$E:$E"))
         X = X + 1
    Next RngTgtCalculation
    
    '// Set variables for actual mix loop.
    Set RngTgtMix = Range(RngToAdd.Offset(, EdaTargetMixes), RngToAdd.Offset(, EdaTargetMixes).End(xlDown))
    X = 0
     
    '// Sumif caclulation for table actual mix data.
    For Each RngTgtCalculation In RngTgtMix
        RngTgtCalculation = WorksheetFunction.SumIf(ShMixes.Range("B:B"), RngToAdd.Offset(X, EdaDate), ShMixes.Range("$D:$D"))
         X = X + 1
    Next RngTgtCalculation
    
End Sub
