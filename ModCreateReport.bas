Attribute VB_Name = "ModCreateReport"
Option Private Module

Sub SpCreateReport()
'// This procedure will create a weekly yeild report for the user. It creates the report based in user selected date range.
'// Created by AL on 1/7/2024.

    '// Declare variables.
    Dim SelectedDate As Date
    Dim ReportRange As Date
    Dim ReportBook As Workbook
    Dim WipBook As Workbook
    Dim WipOne As String
    Dim WipTwo As String
    Dim WipThree As String
    Dim DataRange As Range
    Dim MaterialInfo As Range
    
    '// Assign variables.
    On Error Resume Next
    SelectedDate = InputBox(Prompt:="Please input the date for the desired report inquiry" _
                            , Title:="Please Select A Date" _
                            , Default:="MM/DD/YY")
    ReportRange = SelectedDate - Weekday(SelectedDate, vbMonday) + 1
    Set ReportBook = Application.Workbooks.Add
    Set WipBook = Application.ThisWorkbook
    WipOne = "400140050421"
    WipTwo = "400140050496"
    WipThree = "400140050497"
    WipBook.Activate
    Set DataRange = ShTable.Range(Cells(4, EdaDate), Cells(Rows.Count, EdaDate).End(xlUp))
    On Error GoTo 0

    '// End procedure if no date or cancel was selected from input box.
    If SelectedDate = #12:00:00 AM# Or ReportRange = #12/25/1899# Then Exit Sub
    
    '// Start loop though data for report.
    Dim I As Long
    I = 1
    For X = ReportRange To SelectedDate
        For Each MaterialInfo In DataRange
            '// Check if line 1 has data.
            If MaterialInfo = X And MaterialInfo.Offset(, 2) = WipOne Then
                Cells(MaterialInfo.Row, 1).Resize(, 16).Copy
                ReportBook.Worksheets("Sheet1").Range("A6").Offset(I).PasteSpecial (xlPasteValuesAndNumberFormats)
            End If
            
            '// Check if line 3 has data.
            If MaterialInfo = X And MaterialInfo.Offset(, 2) = WipTwo Then
                Cells(MaterialInfo.Row, 1).Resize(, 16).Copy
                ReportBook.Worksheets("Sheet1").Range("A19").Offset(I).PasteSpecial (xlPasteValuesAndNumberFormats)
            End If
            
            '// Check if line 4 has data.
            If MaterialInfo = X And MaterialInfo.Offset(, 2) = WipThree Then
                Cells(MaterialInfo.Row, 1).Resize(, 16).Copy
                ReportBook.Worksheets("Sheet1").Range("A31").Offset(I).PasteSpecial (xlPasteValuesAndNumberFormats)
            End If
        Next MaterialInfo
        I = I + 1
    Next X

    '// Copy headers from data table.
    ShTable.Range("A3:P3").Copy
    ReportBook.Worksheets("Sheet1").Range("A6").PasteSpecial (xlPasteValues)
    
    '// Activate report book for formatting.
    ReportBook.Activate

    '// Format the created report header.
    With Range("A2:P2")
        .Merge
        .Value = "Daily Wip Yield Report"
    End With
    
    '// Select header and make bold.
    Range("A2:P2").Select
    Selection.Font.Bold = True
    
    '// Format report header.
    With Selection.Font
        .Name = "Calibri"
        .Size = 22
        .Underline = xlUnderlineStyleNone
        .ThemeFont = xlThemeFontMinor
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
    End With
    
    '// Add double border to header
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
    End With
    
   '// Align header text.
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    '// Format data headers.
    With Range("A6:P6")
        .Font.Bold = True
    End With
    
    '// Format line assignment on report.
    With Range("A5")
        .Font.Bold = True
        .Font.ThemeColor = xlThemeColorAccent1
        .Font.TintAndShade = -0.499984740745262
    End With

    '// Install formulas into report.
    Range("F14").FormulaR1C1 = "=SUM(R[-7]C:R[-1]C)"
    Range("F14").AutoFill Destination:=Range("F14:O14"), Type:=xlFillDefault
    Range("P14").FormulaR1C1 = "=AVERAGE(R[-7]C:R[-1]C)"
    Range("E14").FormulaR1C1 = "Totals:"

    '// Install shade into total row for report.
    With Range("E14:P14")
        .Font.Bold = True
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorDark1
        .Interior.TintAndShade = -0.249977111117893
        .Interior.PatternTintAndShade = 0
    End With
    
    '// Format column widths and hide un-needed columns.
    Columns("A:Q").EntireColumn.AutoFit
    Columns("L:N").EntireColumn.Hidden = True

    '// Adjust printer setting for report.
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    
End Sub
