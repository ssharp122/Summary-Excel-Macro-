# Summary-Excel-Macro
Excel Macro used to create a new column entry with relevant financial information to automate cost sheet making process

Sub Summary()

'

' Summary Macro

'

' Keyboard Shortcut: Ctrl+p

'

    ActiveWorkbook.Save

    Sheets("Sheet3").Select

    Range("A1").Select

    Range("D37").Select

    Selection.Copy

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("E37").Select

    Application.CutCopyMode = False

    Selection.Copy

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("F37").Select

    Application.CutCopyMode = False

    Selection.Copy

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("G37").Select

    Application.CutCopyMode = False

    Selection.Copy

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("H37").Select

    Application.CutCopyMode = False

    Selection.Copy

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("I37").Select

    Application.CutCopyMode = False

    Selection.Copy

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("J37").Select

    Application.CutCopyMode = False

    Selection.Copy

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("L38").Select

    Range("D23:J40").Select

    Selection.Copy

    Range("C23").Select

    ActiveSheet.Paste

    Range("I5").Select

    Selection.Copy

    Range("J25").Select

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("C3").Select

    Application.CutCopyMode = False

    Selection.Copy

    Range("J26").Select

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("C15").Select

    Application.CutCopyMode = False

    Selection.Copy

    Range("J27").Select

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("C17").Select

    Application.CutCopyMode = False

    Selection.Copy

    Range("J28").Select

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("J29").Select

    ActiveCell.FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"

    Range("K21").Select

    Application.CutCopyMode = False

    Selection.Copy

    Range("J31").Select

   Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("H9").Select

    Application.CutCopyMode = False

    Selection.Copy

    Range("J33").Select

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("H10").Select

    Application.CutCopyMode = False

    Selection.Copy

    Range("J35").Select

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("J37").Select

    Application.CutCopyMode = False

    Range("J37").Select

    ActiveCell.FormulaR1C1 = "=+R[-4]C-R8C4"

    Range("J1").Select

    Selection.Copy

    Range("J39").Select

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("J23").Select

    Application.CutCopyMode = False

    ActiveCell.FormulaR1C1 = "=TODAY()-1"

    Range("J23").Select

    Selection.Copy

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("A1").Select

    Sheets("Sheet2").Select

    Range("F1").Select

    ActiveCell.FormulaR1C1 = "=TODAY()-1"

    Range("F1").Select

    Selection.Copy

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Sheets("Sheet3").Select

    Range("A1").Select

    Range("K25").Select

    Application.CutCopyMode = False

    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-8]:RC[-1])"

    Range("K26").Select

    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-8]:RC[-1])"

    Range("K27").Select

    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-8]:RC[-1])"

    Range("K28").Select

    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-8]:RC[-1])"

    Range("K29").Select

    ActiveCell.FormulaR1C1 = "=R[-2]C+R[-1]C+R[-3]C"

    Range("K31").Select

    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-8]:RC[-1])"

    Range("K33").Select

    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-8]:RC[-1])"

    Range("K35").Select

    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-8]:RC[-1])"

    Range("K37").Select

    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-8]:RC10)"

    Range("K39").Select

   ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-8]:RC[-1])"

    Range("A1").Select

    Range("C23:J23").Select

    Selection.NumberFormat = "m/d/yy;@"

    Range("A1").Select

    Sheets("Sheet1").Select

    Range("B6").Select

    Selection.Copy

    Range("A1").Select

    Sheets("Sheet3").Select

    Range("J40").Select

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _

        :=False, Transpose:=False

    Range("A1").Select

    Sheets("Sheet1").Select

    Range("A1").Select

    Application.CutCopyMode = False

    Range("A1").Select

    Sheets("Sheet3").Select

    Range("A1").Select
