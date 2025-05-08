Attribute VB_Name = "ModImportData"
Option Private Module

Sub SpChainAllImports()
'// This sub will chain all the import procedures for data together. Created to avoid persistent gui confirmations.
'// Created by AL.

    '// This sub will import the milano order information for the selected user date range.
    '// Created by AL on 11/8/2024.
    
    '// Declare variables.
    Dim DateEntry As Date
    Dim Search2 As Date
    
    '// Assign values from named ranges.
    DateEntry = Range("DateEntry").Value
    Search2 = Range("Search2").Value
    
    '// Conditional check for procedure to run.
    If DateEntry = Empty Then
        MsgBox Prompt:="Please enter the date.", Buttons:=vbOKCancel + vbExclamation, Title:="Enter Date"
        Exit Sub
    End If
    
    '// If no second date is selected use first date.
    If Search2 = "12:00:00AM" Then
        Search2 = Range("DateEntry").Value
    End If
    
    '// Establish sap connectiion.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
    
    '// Select tcode for sap.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtP_PROFID").Text = "000001"
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = "/AL COID"
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "400140050421"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "400140050496"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "400140050497"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = "4014"
'    session.findById("wnd[0]/usr/ctxtP_SYST2").Text = "clsd"
'    session.findById("wnd[0]/usr/chkP_KZ_E2").Selected = True
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").Text = DateEntry
    session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").Text = Search2
    session.findById("wnd[0]").sendVKey 8
    '// Export data to clipboard.
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    
    '// Paste data into coid sheet.
    With ShCoid
        .Visible = True
        .Activate
        .Columns("A:Z").ClearContents
        Range("A1").Select
        ActiveSheet.Paste
        
        '// Format imported data.
        Columns("A:A").Select
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
        
        '// Hide Coid sheet.
        .Visible = False
    End With
    
    '// Select table sheet.
    ShTable.Select
    
    
    '// This macro will import phase 20 mixes from sap by the user selected date range.
    '// Created by AL on 11/11/2024.
    
    '// Start sap tcode.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    
    session.findById("wnd[0]/usr/radREP_OPER").Select
    session.findById("wnd[0]/usr/radREP_OPER").SetFocus
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtP_PROFID").Text = "000001"
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = "/ALMIXCOMMIT"
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "400140050421"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "400140050496"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "400140050497"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtS_CWERK-LOW").Text = "4014"
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").Text = DateEntry
    session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").Text = Search2
    session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").SetFocus
    session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").caretPosition = 10
    session.findById("wnd[0]").sendVKey 8
    '// Export data to clipboard.
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    
    '// Paste and format data.
    With ShMixes
        .Visible = True
        .Activate
        .Columns("A:Z").ClearContents
        Range("A1").Select
        ActiveSheet.Paste
        Columns("A:A").Select
        
        '// Format data in sheet.
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
        
        '// hide mix sheet.
        .Visible = False
    End With
    
    '// Select table sheet.
    ShTable.Select
    
    '// This sub will import the material activity report based on the range(ingredient).
    '// Created by AL on 8/4/2024. R-09/07/2024.

    '// Declare variables.
    Dim Ingredient As String
    Dim RngDynamic As Range
    Dim MaterialCheckOne As Long
    Dim MaterialCheckTwo As Long
    Dim MaterialCheckThree As Long
    Dim AlreadyRan As Boolean
    Dim AlreadyRanTwo As Boolean
    
     On Error Resume Next
    '// Assign values from named ranges.
    DateEntry = Range("DateEntry").Value - 1
    Search2 = Range("Search2").Value
    Set RngDynamic = ShUsage.Range("A1")
    MaterialCheckOne = WorksheetFunction.Match(400140050421#, ShCoid.Range("C:C"), 0)
    MaterialCheckTwo = WorksheetFunction.Match(400140050496#, ShCoid.Range("C:C"), 0)
    MaterialCheckThree = WorksheetFunction.Match(400140050497#, ShCoid.Range("C:C"), 0)
    On Error GoTo 0

    '// Give value to variable search2
    If Search2 = "12:00:00AM" Then
        Search2 = Range("DateEntry").Value + 1
    ElseIf Search2 <> "12:00:00AM" Then
        Search2 = Range("Search2").Value + 1
    End If
    
    '// assign value to ingredient variable.
    If MaterialCheckOne > 0 Then
        Ingredient = "400140050421"
    ElseIf MaterialCheckTwo > 0 Then
        Ingredient = "400140050496"
        AlreadyRan = True
    ElseIf MaterialCheckThree > 0 Then
        Ingredient = "400140050497"
        AlreadyRanTwo = True
    Else
        GoTo DontRun
    End If
    
    '// Clear usage sheets.
    ShUsage.Cells.ClearContents
'    ShUnique.Cells.ClearContents
    
LoopStart:
    '// Start transaction ZPP_MATOVER
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "ZPP_MATOVER"

    '// Handle missing MOA data.
    On Error GoTo ErrRestart:
    session.findById("wnd[0]/usr/ctxtP_WERKS").Text = "4014"
    session.findById("wnd[0]/usr/ctxtP_LGNUM").Text = "406"
    session.findById("wnd[0]/usr/ctxtP_MATNR").Text = Ingredient
    session.findById("wnd[0]/usr/ctxtS_BUDAT-LOW").Text = DateEntry
    session.findById("wnd[0]/usr/ctxtS_BUDAT-HIGH").Text = Search2
    session.findById("wnd[0]").sendVKey 8
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = -1
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BWART"
    session.findById("wnd[0]/tbar[1]/btn[29]").press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "261"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[45]").press

    '// Export the material Report
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    '// Unhide sheets and paste data.
    With ShUsage
        .Visible = True
        .Activate
    End With

    '// Paste dynamic range
    RngDynamic.Select
    RngDynamic.PasteSpecial

ErrRestart:
    
    '// Loop for all milano usage data if applicable.
    If Ingredient <> "400140050496" And MaterialCheckTwo > 0 And AlreadyRan = False Then
        Ingredient = "400140050496"
        Set RngDynamic = Cells(Rows.Count, 1).End(xlUp).Offset(1)
        AlreadyRan = True
        GoTo LoopStart
    ElseIf Ingredient <> "400140050497" And MaterialCheckThree > 0 And AlreadyRanTwo = False Then
        Ingredient = "400140050497"
        Set RngDynamic = Cells(Rows.Count, 1).End(xlUp).Offset(1)
        AlreadyRanTwo = True
        GoTo LoopStart
    End If
    
    '// Install data into columns.
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
       
    '// hide sheets and paste data.
'    With ShUnique
'        .Visible = True
'    End With
    
    '// Filter Unique materila documents and copy.
'    Columns("C:C").Select
'    Range("C1:C50000").AdvancedFilter Action:=xlFilterInPlace, Unique:=True
'    Cells.Copy ShUnique.Range("A1")
'    ShUsage.ShowAllData
    
     '// hide sheets and paste data.
'    With ShUnique
'        .Visible = False
'    End With


DontRun:
    '// hide sheets and paste data.
    With ShUsage
        .Visible = False
    End With
    
    '// Go to table.
    ShTable.Select
    
End Sub

Sub SpImportWipinfo()
'// This sub will import the milano order information for the selected user date range.
'// Created by AL on 11/8/2024.
    
    '// Declare variables.
    Dim DateEntry As Date
    Dim Search2 As Date
    
    '// Assign values from named ranges.
    DateEntry = Range("DateEntry").Value
    Search2 = Range("Search2").Value
    
    '// Conditional check for procedure to run.
    If DateEntry = Empty Then
        MsgBox Prompt:="Please enter the date.", Buttons:=vbOKCancel + vbExclamation, Title:="Enter Date"
        Exit Sub
    End If
    
    '// If no second date is selected use first date.
    If Search2 = "12:00:00AM" Then
        Search2 = Range("DateEntry").Value
    End If
    
    '// Establish sap connectiion.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
    
    '// Select tcode for sap.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtP_PROFID").Text = "000001"
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = "/AL COID"
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "400140050421"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "400140050496"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "400140050497"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = "4014"
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").Text = DateEntry
    session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").Text = Search2
    session.findById("wnd[0]").sendVKey 8
    '// Export data to clipboard.
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    
    '// Paste data into coid sheet.
    With ShCoid
        .Visible = True
        .Activate
        .Columns("A:Z").ClearContents
        Range("A1").Select
        ActiveSheet.Paste
        
        '// Format imported data.
        Columns("A:A").Select
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
        
        '// Hide Coid sheet.
        .Visible = False
    End With
    
    '// Select table sheet.
    ShTable.Select
    
End Sub

Sub SpImportMatUsage()
'// This sub will import the material activity report based on the range(ingredient).
'// Created by AL on 8/4/2024. R-09/07/2024.

    '// Declare variables.
    Dim DateEntry As Date
    Dim Search2 As Date
    Dim Ingredient As String
    Dim RngDynamic As Range
    Dim MaterialCheckOne As Long
    Dim MaterialCheckTwo As Long
    Dim MaterialCheckThree As Long
    
     On Error Resume Next
    '// Assign values from named ranges.
    DateEntry = Range("DateEntry").Value - 1
    Search2 = Range("Search2").Value
    Set RngDynamic = ShUsage.Range("A1")
    MaterialCheckOne = WorksheetFunction.Match(400140050421#, ShCoid.Range("C:C"), 0)
    MaterialCheckTwo = WorksheetFunction.Match(400140050496#, ShCoid.Range("C:C"), 0)
    MaterialCheckThree = WorksheetFunction.Match(400140050497#, ShCoid.Range("C:C"), 0)
    On Error GoTo 0
    
    '// Conditional checks to run macro and place value to variable search2.
    If DateEntry = Empty Then
        MsgBox Prompt:="Please enter a valid date selction and try again.", Buttons:=vbExclamation + vbOKOnly, Title:="Date Selection"
        Exit Sub
    End If

    '// Give value to variable search2
    If Search2 = "12:00:00AM" Then
        Search2 = Range("DateEntry").Value + 1
    ElseIf Search2 <> "12:00:00AM" Then
        Search2 = Range("Search2").Value + 1
    End If
    
    '// assign value to ingredient variable.
    If MaterialCheckOne > 0 Then
        Ingredient = "400140050421"
    ElseIf MaterialCheckTwo > 0 Then
        Ingredient = "400140050496"
    ElseIf MaterialCheckThree > 0 Then
        Ingredient = "400140050497"
    Else
        GoTo DontRun
    End If
    
    '// Clear usage sheet.
    ShUsage.Cells.ClearContents

    '// SAP connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)

    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If

LoopStart:
    '// Start transaction ZPP_MATOVER
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "ZPP_MATOVER"

    session.findById("wnd[0]/usr/ctxtP_WERKS").Text = "4014"
    session.findById("wnd[0]/usr/ctxtP_LGNUM").Text = "406"
    session.findById("wnd[0]/usr/ctxtP_MATNR").Text = Ingredient
    session.findById("wnd[0]/usr/ctxtS_BUDAT-LOW").Text = DateEntry
    session.findById("wnd[0]/usr/ctxtS_BUDAT-HIGH").Text = Search2
    session.findById("wnd[0]").sendVKey 8
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = -1
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BWART"
    session.findById("wnd[0]/tbar[1]/btn[29]").press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "261"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[45]").press

    '// Export the material Report
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    '// Unhide sheets and paste data.
    With ShUsage
        .Visible = True
        .Activate
    End With

    '// Paste dynamic range
    RngDynamic.Select
    RngDynamic.PasteSpecial

    Dim AlreadyRan As Boolean
    Dim AlreadyRanTwo As Boolean
    
    '// Loop for all milano usage data if applicable.
    If Ingredient <> "400140050496" And MaterialCheckTwo > 0 And AlreadyRan = False Then
        Ingredient = "400140050496"
        Set RngDynamic = Cells(Rows.Count, 1).End(xlUp).Offset(1)
        AlreadyRan = True
        GoTo LoopStart
    ElseIf Ingredient <> "400140050497" And MaterialCheckThree > 0 And AlreadyRanTwo = False Then
        Ingredient = "400140050497"
        Set RngDynamic = Cells(Rows.Count, 1).End(xlUp).Offset(1)
        AlreadyRan = True
        GoTo LoopStart
    End If
    
    '// Install data into columns.
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
    
    '// hide sheets and paste data.
    With ShUsage
        .Visible = False
    End With
    
DontRun:
    
    '// Go to table.
    ShTable.Select

End Sub

Sub SpImportSapMixes()
'// This macro will import phase 20 mixes from sap by the user selected date range.
'// Created by AL on 11/11/2024.

    '// Declare variables.
    Dim DateEntry As Variant
    Dim Search2 As Date

    '// Assign values.
    DateEntry = Range("DateEntry").Value
    Search2 = Range("Search2").Value
    
    '// Conditional statements to run sub
    If DateEntry = "" Then
        MsgBox "Please enter the date.", vbExclamation + vbOKCancel, "Enter Date"
        Exit Sub
    End If
    
    If Search2 = "12:00:00AM" Then
        Search2 = Range("DateEntry").Value
    End If
    
    '// Establish sap connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject Application, "on"
    End If
    
    '// Start sap tcode.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    
    session.findById("wnd[0]/usr/radREP_OPER").Select
    session.findById("wnd[0]/usr/radREP_OPER").SetFocus
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtP_PROFID").Text = "000001"
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = "/ALMIXCOMMIT"
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "400140050421"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "400140050496"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "400140050497"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtS_CWERK-LOW").Text = "4014"
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").Text = DateEntry
    session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").Text = Search2
    session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").SetFocus
    session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").caretPosition = 10
    session.findById("wnd[0]").sendVKey 8
    '// Export data to clipboard.
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    
    '// Paste and format data.
    With ShMixes
        .Visible = True
        .Activate
        .Columns("A:Z").ClearContents
        Range("A1").Select
        ActiveSheet.Paste
        Columns("A:A").Select
        
        '// Format data in sheet.
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
        
        '// hide mix sheet.
        .Visible = False
    End With
    
    '// Select table sheet.
    ShTable.Select

End Sub

Sub SpImportShiftReport()
'// This macro imports the shift report based on the date.
'// Created by AL on 5/1/2024. R-08/25/2024

    '// Declare variables.
    Dim DateEntry As String
    Dim FileDate As String
    Dim WrkBook As String
    Dim SheetDate As String
    Dim Sheetpath As String
    Dim Trackerbk As String
    
    '// Assign values
    DateEntry = Range("DateEntry")
    FileDate = Format(DateEntry, "yyyy-mm-dd")
    WrkBook = "G:\Reports\Cookie Daily Reports Archive\CookieReports\Cookie Daily Reports\" + (FileDate) + ".xlsx"
    SheetDate = Format(DateEntry, "yyyymmdd")
    Sheetpath = SheetDate + "Data"
    Trackerbk = Format(DateEntry, "MMMM YYYY") + " Cookie"
        
    '// Verification check of macro.
    If DateEntry = "" Then
        MsgBox Prompt:="Please Insert The Date", Buttons:=vbCritical + vbOKCancel, Title:="Date Entry"
        Exit Sub
    End If
    
    '// Clears old data in worksheet.
    With ShReport
        .Columns("A:Z").ClearContents
    End With
    
    '// Opens source workbook and copies data.
    Workbooks.Open Filename:=WrkBook, UpdateLinks:=3, ReadOnly:=True
    Sheets("NO").Select
    Range("A1:G30").Select
    Selection.Copy
    '// Opens destination workbook and pastes data.
    ShReport.Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
    '// Opens source workbook and copies data.
    Windows(FileDate + ".xlsx").Activate
    Sheets("AM").Select
    Range("A1:G30").Select
    Application.CutCopyMode = False
    Selection.Copy
    '// Opens destination workbook and pastes data.
    ShReport.Activate
    Range("A31").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
    '// Opens source workbook and copies data.
    Windows(FileDate + ".xlsx").Activate
    Sheets("PM").Select
    Range("A1:G30").Select
    Application.CutCopyMode = False
    Selection.Copy
    '// Opens destination workbook and pastes data.
    ShReport.Activate
    Range("A62").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    '// Closes the source workbook.
    Windows(FileDate + ".xlsx").Close
    Range("A1").Select
    
    '// Select Data table sheet
    ShTable.Select
    
End Sub

Sub SpImportMexReport()
' This macro will let the user select what mexvision report to import as the data source.
' Created by AL on 6/12/2024.
    
    '// Declare variables.
    Dim FileToOpen As Variant
    Dim SelectedBook As Workbook
    ChDrive "G:"
    ChDir "G:\Control Room\A Lassalle\MexVision"

    '// Let user select file to import.
    FileToOpen = Application.GetOpenFilename(filefilter:="Excel Files(*.xls*),*xls*", Title:="PLEASE SELECT THE CORRECT FILE")
    Application.ScreenUpdating = False
    
    '// If user did not select a workbook then exit sub.
    If FileToOpen <> False Then
        Set SelectedBook = Application.Workbooks.Open(FileToOpen, , ReadOnly:=True)
        ElseIf FileToOpen = False Then
        Exit Sub
    End If
    
    '// Select mm import sheet and import.
    ShMex.Visible = True
    ShMex.Cells.ClearContents
    
    '//  Cop data from selected book and close.
    SelectedBook.Sheets(1).Cells.Copy Destination:=ShMex.Range("A1")
    SelectedBook.Close Savechanges:=False
    
    '// Hide mex rpeort sheet.
    ShMex.Visible = False
    
End Sub
