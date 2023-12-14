Attribute VB_Name = "LastFourDigitRight"
Sub FourDigitRight()
Attribute FourDigitRight.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' FourDigitRight Macro
'
' Keyboard Shortcut: Ctrl+Shift+S
'
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.FormulaR1C1 = "=""***-**-""&RIGHT(RC[-2],4)"
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A698")
    ActiveCell.Range("A1:A698").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -2).Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 2).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveCell.Activate
    Selection.ClearContents
    ActiveCell.Offset(0, -2).Range("A1").Select
End Sub

Sub TimesNewRoman()
'
' TimesNewRoman Macro
'
' Keyboard Shortcut: Ctrl+Shift+T
'
    'Application.Left = 5.2
    'Application.Top = 7.6
    'Application.Width = 620.4
    'Application.Height = 606.6
    Cells.Select
    Selection.Font.Name = "Times New Roman"
        
End Sub

//////////// Activesheet Name ///////////
Public Sub Test123()

Sheets("Conservative1").Select
ActiveSheet.Name = "Conservative"
Sheets("Balanced1").Select
ActiveSheet.Name = "Balanced"

End Sub
/////////////////////////////////////////

