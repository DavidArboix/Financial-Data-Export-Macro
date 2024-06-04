# Financial-Data-Export-Macro

Sub Macro20()
'
' Macro20 Macro
'
' Acceso directo: CTRL+w
'
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:G").Select
    ActiveSheet.Range("$A$1:$G$55").RemoveDuplicates Columns:=1, Header:=xlNo
    Columns("A:A").Select
    Selection.Replace What:=" ", Replacement:="_", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="%", Replacement:="Percentage", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(", Replacement:="I", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=")", Replacement:="I", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="/", Replacement:="Per", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=":", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("A:G").Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=",", Replacement:=".", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("B1:F1").Select
    Selection.NumberFormat = "General"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "2019"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "2020"
    Range("B1:C1").Select
    Selection.AutoFill Destination:=Range("B1:F1"), Type:=xlFillDefault
    Range("B1:F1").Select
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.Copy
    Columns("B:B").Select
    ActiveSheet.Paste
    Range("B3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC1&""_""&R1C[1]"
    Range("B3").Select
    Selection.AutoFill Destination:=Range("B3:B75"), Type:=xlFillDefault
    Range("B3:B75").Select
    ActiveWindow.SmallScroll Down:=-60
    Columns("B:B").Select
    ActiveSheet.Range("$B$1:$B$75").RemoveDuplicates Columns:=1, Header:=xlNo
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("B:B").Select
    Selection.Copy
    Columns("D:D").Select
    ActiveSheet.Paste
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("D:D").Select
    Selection.Copy
    Columns("F:F").Select
    ActiveSheet.Paste
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("F:F").Select
    Selection.Copy
    Columns("H:H").Select
    ActiveSheet.Paste
    Columns("J:J").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Copy
    Columns("J:J").Select
    ActiveSheet.Paste
    Columns("L:L").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("J:J").Select
    Selection.Copy
    Columns("L:L").Select
    ActiveSheet.Paste
    Range("B3:C75").Select
    Application.CutCopyMode = False
    Selection.CreateNames Top:=False, Left:=True, Bottom:=False, Right:= _
        False
    Range("D3:E75").Select
    Selection.CreateNames Top:=False, Left:=True, Bottom:=False, Right:= _
        False
    Range("F3:G75").Select
    Selection.CreateNames Top:=False, Left:=True, Bottom:=False, Right:= _
        False
    Range("H3:I75").Select
    Selection.CreateNames Top:=False, Left:=True, Bottom:=False, Right:= _
        False
    Range("J3:K75").Select
    Selection.CreateNames Top:=False, Left:=True, Bottom:=False, Right:= _
        False
    Range("L3:M75").Select
    Selection.CreateNames Top:=False, Left:=True, Bottom:=False, Right:= _
        False
    Range("A1").Select
End Sub
