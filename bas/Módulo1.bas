Attribute VB_Name = "M�dulo1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

    Range("C71").Select
    Selection.Copy
    Range("C72").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    Range("C86").Select
    Selection.Copy
    Range("C87").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub


Sub Macro2()

Dim ws As Worksheet
Set ws = Worksheets("BANCOS")

ws.EnableCalculation = False
ws.EnableCalculation = True

Set ws = Nothing

End Sub
