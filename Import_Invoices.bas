Attribute VB_Name = "Module2"
Sub Invoice_data()
Attribute Invoice_data.VB_Description = "Import invoice data"
Attribute Invoice_data.VB_ProcData.VB_Invoke_Func = "i\n14"

Application.ScreenUpdating = False

'
' Invoice_data Macro
' Import invoice data
'
' Keyboard Shortcut: Ctrl+i
'
    Sheets("Sales Sheet").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Invoice").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B2").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Sheets("Sales Sheet").Select
    Range("M7").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Invoice").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C9").Select
    Sheets("Sales Sheet").Select
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Invoice").Select
    Range("E11").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("C2").Select
    Sheets("Sales Sheet").Select
    Selection.Copy
    Sheets("Invoice").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D3").Select
    Sheets("Sales Sheet").Select
    Range("M6").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Invoice").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Sales Sheet").Select
    Range("M10").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Invoice").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Sales Sheet").Select
    Range("M8").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("D2").Select
    Sheets("Sales Sheet").Select
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Invoice").Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Sales Sheet").Select
    Range("M10").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Invoice").Select
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Sales Sheet").Select
    Range("L7").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Invoice").Select
    Range("H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("I7").Select
    
    
    ThisWorkbook.RefreshAll
    
    Application.ScreenUpdating = True

    
End Sub
