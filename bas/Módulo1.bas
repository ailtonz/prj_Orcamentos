Attribute VB_Name = "Módulo1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

    '' PREÇO VENDA GROSS MARKETING
    Range("C71").Select
    Selection.Copy
    Range("C72").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    '' PREÇO VENDA GROSS COMPRAS
    Range("C86").Select
    Selection.Copy
    Range("C87").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub



