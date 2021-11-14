Attribute VB_Name = "Module4"
Sub HSBC_Bank_Statementv2()
'   This macro formats the bank statement to be copied in the Daybook.
'   Version 2 - 15/11/2021
'
start_sheet = Range("a1").Address
end_sheet = Range("a1").End(xlDown).Address

'   Format HSBC
    Range(start_sheet, end_sheet).Select
    Selection.AutoFilter
    With Selection
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Columns("A:R").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight

'   This section extracts the client's name from the narrative
'   It runs from the 'TRN type' column to identify which formula to use
'   Edit formulas
start_point = Range("d2").Address
end_point = Range("d2").End(xlDown).Address
Range(start_point, end_point).Select
Dim cell As Range
For Each cell In ActiveSheet.Range(start_point, end_point)
    If cell.Value = "FBP     " Then
        cell.Cells(1, -1).Select
        Selection.NumberFormat = "General"
        ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],FIND("" FP0"",RC[-1])-1)"
    Else
        cell.Cells(1, -1).Select
        Selection.NumberFormat = "General"
        ActiveCell.FormulaR1C1 = "=TRIM(SUBSTITUTE(RC[-1],RC[1],"""",1))"
    End If
Next cell

End Sub
