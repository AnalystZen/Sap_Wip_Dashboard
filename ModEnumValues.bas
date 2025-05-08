Attribute VB_Name = "ModEnumValues"
Option Private Module
'// List all emun values or global ranges here.

    '// Enum list for data table.
    Enum EDataTable
        EdaDate = 1
        EdaProcessOrder
        EdaMaterialNumber
        EdaMaterialDescription
        EdaBatchNumber
        EdaTgtQuantity
        EdaDeliveredQuantity
        EdaConfirmedQuantity
        EdaWithdrawnQuantity
        EdaTargetMixes
        EdaActualMixes
    End Enum

    '// Enum list for Coid Copy data.
    Enum ECoid
        EcoOrder = 2
        EcoMaterial = 3
        EcoDescription = 4
        EcoBatch = 5
        EcoStart = 8
        EcoTarget = 9
        EcoDelivered = 10
        EcoConfirmed = 11
    End Enum
        
    '// Enum list for data import rows
    Enum EImportRow
        EimCoid = 4
    End Enum
