' Macro for formating the daily checking report
' Format the daily checking report
'
Sub format_daily_report()
'
' Delete the first row
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Selection.RowHeight = 50
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
'
' Set text alignment
    Cells.Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
'
' Format the column width
    Range("A1").Select
    Columns("A:A").ColumnWidth = 8
    Columns("B:B").ColumnWidth = 8
    Columns("C:C").ColumnWidth = 19.5
    Columns("D:D").ColumnWidth = 23.5
    Columns("E:E").ColumnWidth = 8
    Columns("F:F").ColumnWidth = 7.57
    Columns("G:G").ColumnWidth = 7.57
    Columns("H:H").ColumnWidth = 16
    Columns("I:I").ColumnWidth = 13.29
    Columns("J:J").ColumnWidth = 14
    Columns("K:K").ColumnWidth = 12
    Columns("L:L").ColumnWidth = 12
    Columns("M:M").ColumnWidth = 12
    Columns("N:N").ColumnWidth = 12
    Columns("O:O").ColumnWidth = 12
    Columns("P:P").ColumnWidth = 12
    Columns("Q:Q").ColumnWidth = 12
    Columns("R:R").ColumnWidth = 12
    Columns("S:S").ColumnWidth = 9
    Columns("T:T").ColumnWidth = 9.57
    Columns("U:U").ColumnWidth = 8.57
    Columns("V:V").ColumnWidth = 8.57
    Columns("W:W").ColumnWidth = 10
'
' Text to columns
    Columns("C:C").Select
    Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
'
' Sort the audit ID
    Range("C2").Select
    lastRow = ActiveSheet.UsedRange.Rows.Count
    lastColumn = ActiveSheet.UsedRange.Columns.Count
    ActiveWorkbook.Worksheets("Actions").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Actions").Sort.SortFields.Add2 Key:=Range( _
        Selection, Selection.End(xlDown)), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Actions").Sort
        .SetRange Range(Cells(2, 1), Cells(lastRow, lastColumn))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'
' Freeze top row
    Range("A1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
'
' Filter the closed CAPs
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.UsedRange.AutoFilter Field:=1, Criteria1:= _
        "=Ongoing", Operator:=xlOr, Criteria2:="=Resolved"
'
' Goto home A1 position
    Range("A1").Select

End Sub
'
' End of macro program
' ====================
