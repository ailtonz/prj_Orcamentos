Attribute VB_Name = "M�dulo1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

    '' PRE�O VENDA GROSS MARKETING
    Range("C71").Select
    Selection.Copy
    Range("C72").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    '' PRE�O VENDA GROSS COMPRAS
    Range("C86").Select
    Selection.Copy
    Range("C87").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub



