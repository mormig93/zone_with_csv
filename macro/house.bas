Attribute VB_Name = "AmazonMacro"
Sub AmazonMacro()

' Limpia las celdas vacias
    Columns("E:E").Select
    Selection.AutoFilter
    ActiveSheet.Range("$E$1:$E$829").AutoFilter Field:=1, Criteria1:="="
    Range("D830").Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.EntireRow.Delete
    Columns("E:E").Select
    Selection.AutoFilter

' Deletes all null elements col E
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A3:J3").Select
    Selection.Copy
    Range("A2").Select
    ActiveSheet.Paste
    Range("E2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "null"
    Range("E1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$1122").AutoFilter Field:=5, Criteria1:="null"
    ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveCell.Offset(-1, 4).Range("A1").Select
    Selection.AutoFilter
    Range("Z2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-24],'[Categorias Amazon Cupones.xlsx]Sheet1'!C4:C6,3,0)"
    Range("Z2").Select
    Selection.Copy
    ActiveCell.Offset(0, -16).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 16).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Columns("Z:Z").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Z1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Category"
    Range("Z2").Select
    Range("k2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]&"".""&RC[-1]"
    Range("K2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("K2").Select
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-9],5,4)"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-9],5,5)"
    Range("L2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlUp).Select 
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/100"
    Range("M2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlUp).Select
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]-(RC[-3]*RC[-1])"
    Range("N2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-12],26)"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-12],26)"
    Range("O2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=TRIM(RC[-1])"
    Range("P2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-2],4)"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-2],6)"
    Range("Q2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=TRIM(RC[-1])"
    Range("R2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("S1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="&", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("T:T").Select
    Selection.ClearContents
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]&""&tag=housedeals.us-20"""
    Range("T2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("T1").Select
    Range("F2").Select
    Selection.Copy
    Columns("F:F").Select
    Selection.Replace What:="._SR400,400_", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("F1").Select
    ActiveWindow.ScrollColumn = 1
    Columns("A:AD").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("T2").Select
    Application.CutCopyMode = False
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A:A,B:B").Select
    Range("B1").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Range("D:D,E:E,F:F,G:G").Select
    Range("G1").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Regular"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Discount"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Special Pr"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Expires"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Aff Link"
    Range("H2").Select
    Columns("F:F").Select
    Selection.NumberFormat = "0.00"
    Range("F2").Select
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],27,10)"
    Range("I2").Select
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.End(xlUp).Select
    Selection.Copy
            ActiveCell.Offset(0, -1).Range("A1").Select
            Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
            Selection.End(xlToLeft).Select
            ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
            Application.CutCopyMode = False
    Selection.End(xlUp).Select
    ActiveCell.Columns("A:A").EntireColumn.Select
            ActiveCell.Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll Down:=-9
    ActiveCell.Select
    Application.CutCopyMode = False
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "ASIN"
    Columns("G:G").Select
    Selection.Replace What:=" w", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormat = "m/d/yyyy"
     Columns("J:M").Select
    Selection.Delete Shift:=xlToLeft

'
' MacroPromoCode Macro
'
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "=COUNTA(C[-10])"
    x = Range("K1").Value

    Range("M1").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(RIGHT(RC[-12],LEN(RC[-12])-(FIND(""code "",RC[-12],1))-LEN(""code "")+1),"""")"
    Range("M1:M"&x).Select
    Selection.FillDown
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[1],8)"
    Range("L1:L"&x).Select
    Selection.FillDown
    ActiveCell.FormulaR1C1 = "=LEFT(RC[1],8)"
    Selection.Copy
    Range("K1:K"&x).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Promocode"
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft

    MsgBox "Preparacion Completa"
End Sub