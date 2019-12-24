Attribute VB_Name = "Módulo2"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("C65").Select
    ActiveCell.FormulaR1C1 = "22345.21"
    Range("C66").Select
    Application.CutCopyMode = False
    Range("C65").Select
    ActiveCell.FormulaR1C1 = "24380.21"
    Range("C66").Select
    ActiveWorkbook.Save
    ChDir "C:\Users\AILTON\Desktop\WORK\__SPRINGER\SPRINGER"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\AILTON\Desktop\WORK\__SPRINGER\SPRINGER\Orcamentos_140825-1509.xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    ActiveWindow.SmallScroll Down:=9
    Range("C65").Select
    ActiveWindow.SmallScroll Down:=-3
    ActiveCell.FormulaR1C1 = "15000"
    Range("C66").Select
    ActiveWindow.SmallScroll Down:=15
    Range("C80").Select
    ActiveCell.FormulaR1C1 = "12300"
    Range("C81").Select
    ActiveWindow.SmallScroll Down:=-3
    Range("C78").Select
    ActiveCell.FormulaR1C1 = "3%"
    Range("D78").Select
    Selection.Copy
    Range("B78").Select
    ActiveSheet.Paste
    Range("C78").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("C78").Select
    ActiveCell.FormulaR1C1 = "3%"
    Range("C80").Select
    ActiveCell.FormulaR1C1 = "14600"
    Range("C78").Select
    ActiveCell.FormulaR1C1 = "1%"
    Range("C80").Select
    ActiveCell.FormulaR1C1 = "14900"
    Range("C78").Select
    ActiveCell.FormulaR1C1 = "5%"
    Range("C80").Select
    ActiveCell.FormulaR1C1 = "14300"
    Range("C86").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-63]C=""PRODUTO"",IF(ISERROR(+PrecoDeCompras_01/PRODUTO),"""",+PrecoDeCompras_01/PRODUTO),IF(ISERROR(+PrecoDeCompras_01/SERVICO),"""",+PrecoDeCompras_01/SERVICO))"
    Range("C86").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-63]C=""PRODUTO"",IF(ISERROR(+PrecoDeCompras_01/PRODUTO),"""",+PrecoDeCompras_01/PRODUTO),IF(ISERROR(+PrecoDeCompras_01/SERVICO),"""",+PrecoDeCompras_01/SERVICO))"
    Range("D86").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-63]C=""PRODUTO"",IF(ISERROR(+PrecoDeCompras_02/PRODUTO),"""",+PrecoDeCompras_02/PRODUTO),IF(ISERROR(+PrecoDeCompras_02/SERVICO),"""",+PrecoDeCompras_02/SERVICO))"
    Range("E86").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-63]C=""PRODUTO"",IF(ISERROR(+PrecoDeCompras_03/PRODUTO),"""",+PrecoDeCompras_03/PRODUTO),IF(ISERROR(+PrecoDeCompras_03/SERVICO),"""",+PrecoDeCompras_03/SERVICO))"
    Range("F86").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-63]C=""PRODUTO"",IF(ISERROR(+PrecoDeCompras_04/PRODUTO),"""",+PrecoDeCompras_04/PRODUTO),IF(ISERROR(+PrecoDeCompras_04/SERVICO),"""",+PrecoDeCompras_04/SERVICO))"
    Range("G86").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-63]C=""PRODUTO"",IF(ISERROR(+PrecoDeCompras_05/PRODUTO),"""",+PrecoDeCompras_05/PRODUTO),IF(ISERROR(+PrecoDeCompras_05/SERVICO),"""",+PrecoDeCompras_05/SERVICO))"
    Range("H86").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-63]C=""PRODUTO"",IF(ISERROR(+PrecoDeCompras_06/PRODUTO),"""",+PrecoDeCompras_06/PRODUTO),IF(ISERROR(+PrecoDeCompras_06/SERVICO),"""",+PrecoDeCompras_06/SERVICO))"
    Range("I86").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-63]C=""PRODUTO"",IF(ISERROR(+PrecoDeCompras_07/PRODUTO),"""",+PrecoDeCompras_07/PRODUTO),IF(ISERROR(+PrecoDeCompras_07/SERVICO),"""",+PrecoDeCompras_07/SERVICO))"
    Range("J86").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-63]C=""PRODUTO"",IF(ISERROR(+PrecoDeCompras_08/PRODUTO),"""",+PrecoDeCompras_08/PRODUTO),IF(ISERROR(+PrecoDeCompras_08/SERVICO),"""",+PrecoDeCompras_08/SERVICO))"
    Range("J86").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(R[-63]C=""PRODUTO"",IF(ISERROR(+PrecoDeCompras_08/PRODUTO),"""",+PrecoDeCompras_08/PRODUTO),IF(ISERROR(+PrecoDeCompras_08/SERVICO),"""",+PrecoDeCompras_08/SERVICO))"
    Range("C87").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("C72").Select
    ActiveWindow.SmallScroll Down:=15
    Range("D87").Select
    Selection.Copy
    Range("C87").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWorkbook.Names.add Name:="ARREDONDAMENTO_COMPRAS_01", RefersToR1C1:= _
        "='140823-1230'!R87C3"
    ActiveWorkbook.Names("ARREDONDAMENTO_COMPRAS_01").Comment = ""
    Range("D87").Select
    ActiveWorkbook.Names.add Name:="ARREDONDAMENTO_COMPRAS_02", RefersToR1C1:= _
        "='140823-1230'!R87C4"
    ActiveWorkbook.Names("ARREDONDAMENTO_COMPRAS_02").Comment = ""
    Range("E87").Select
    ActiveWorkbook.Names.add Name:="ARREDONDAMENTO_COMPRAS_03", RefersToR1C1:= _
        "='140823-1230'!R87C5"
    ActiveWorkbook.Names("ARREDONDAMENTO_COMPRAS_03").Comment = ""
    Range("F87").Select
    ActiveWorkbook.Names.add Name:="ARREDONDAMENTO_COMPRAS_04", RefersToR1C1:= _
        "='140823-1230'!R87C6"
    ActiveWorkbook.Names("ARREDONDAMENTO_COMPRAS_04").Comment = ""
    Range("G87").Select
    ActiveWorkbook.Names.add Name:="ARREDONDAMENTO_COMPRAS_05", RefersToR1C1:= _
        "='140823-1230'!R87C7"
    ActiveWorkbook.Names("ARREDONDAMENTO_COMPRAS_05").Comment = ""
    Range("H87").Select
    ActiveWorkbook.Names.add Name:="ARREDONDAMENTO_COMPRAS_06", RefersToR1C1:= _
        "='140823-1230'!R87C8"
    ActiveWorkbook.Names("ARREDONDAMENTO_COMPRAS_06").Comment = ""
    Range("I87").Select
    ActiveWorkbook.Names.add Name:="ARREDONDAMENTO_COMPRAS_07", RefersToR1C1:= _
        "='140823-1230'!R87C9"
    ActiveWorkbook.Names("ARREDONDAMENTO_COMPRAS_07").Comment = ""
    Range("J87").Select
    ActiveWorkbook.Names.add Name:="ARREDONDAMENTO_COMPRAS_08", RefersToR1C1:= _
        "='140823-1230'!R87C10"
    ActiveWorkbook.Names("ARREDONDAMENTO_COMPRAS_08").Comment = ""
    Range("C88").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(+ARREDONDAMENTO_COMPRAS_01/TIRAGEM_01),""0"",+ARREDONDAMENTO_COMPRAS_01/TIRAGEM_01)"
    Range("C80").Select
    ActiveCell.FormulaR1C1 = "14300"
    Range("C88").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(+ARREDONDAMENTO_COMPRAS_01/TIRAGEM_01),""0"",+ARREDONDAMENTO_COMPRAS_01/TIRAGEM_01)"
    Range("D88").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(+ARREDONDAMENTO_COMPRAS_02/TIRAGEM_02),""0"",+ARREDONDAMENTO_COMPRAS_02/TIRAGEM_02)"
    Range("E88").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(+ARREDONDAMENTO_COMPRAS_03/TIRAGEM_03),""0"",+ARREDONDAMENTO_COMPRAS_03/TIRAGEM_03)"
    Range("F88").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(+ARREDONDAMENTO_COMPRAS_04/TIRAGEM_04),""0"",+ARREDONDAMENTO_COMPRAS_04/TIRAGEM_04)"
    Range("G88").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(+ARREDONDAMENTO_COMPRAS_05/TIRAGEM_05),""0"",+ARREDONDAMENTO_COMPRAS_05/TIRAGEM_05)"
    Range("H88").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(+ARREDONDAMENTO_COMPRAS_06/TIRAGEM_06),""0"",+ARREDONDAMENTO_COMPRAS_06/TIRAGEM_06)"
    Range("I88").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(+ARREDONDAMENTO_COMPRAS_07/TIRAGEM_07),""0"",+ARREDONDAMENTO_COMPRAS_07/TIRAGEM_07)"
    Range("J88").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(+ARREDONDAMENTO_COMPRAS_08/TIRAGEM_08),""0"",+ARREDONDAMENTO_COMPRAS_08/TIRAGEM_08)"
    Range("C90").Select
    ActiveWindow.SmallScroll Down:=-6
    Range("C86").Select
    ActiveWorkbook.Names.add Name:="PRECO_VENDA_COMPRAS_01", RefersToR1C1:= _
        "='140823-1230'!R86C3"
    ActiveWorkbook.Names("PRECO_VENDA_COMPRAS_01").Comment = ""
    Range("D86").Select
    ActiveWorkbook.Names.add Name:="PRECO_VENDA_COMPRAS_02", RefersToR1C1:= _
        "='140823-1230'!R86C4"
    Range("E86").Select
    ActiveWorkbook.Names.add Name:="PRECO_VENDA_COMPRAS_03", RefersToR1C1:= _
        "='140823-1230'!R86C5"
    Range("F86").Select
    ActiveWorkbook.Names.add Name:="PRECO_VENDA_COMPRAS_04", RefersToR1C1:= _
        "='140823-1230'!R86C6"
    Range("G86").Select
    ActiveWorkbook.Names.add Name:="PRECO_VENDA_COMPRAS_05", RefersToR1C1:= _
        "='140823-1230'!R86C7"
    Range("H86").Select
    ActiveWorkbook.Names.add Name:="PRECO_VENDA_COMPRAS_06", RefersToR1C1:= _
        "='140823-1230'!R86C8"
    Range("I86").Select
    ActiveWorkbook.Names.add Name:="PRECO_VENDA_COMPRAS_07", RefersToR1C1:= _
        "='140823-1230'!R86C9"
    Range("J86").Select
    ActiveWorkbook.Names.add Name:="PRECO_VENDA_COMPRAS_08", RefersToR1C1:= _
        "='140823-1230'!R86C10"
    Range("C90").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(PRECO_VENDA_COMPRAS_01*FASCICULOS_01),"""",PRECO_VENDA_COMPRAS_01*FASCICULOS_01)"
    Range("D90").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(PRECO_VENDA_COMPRAS_02*FASCICULOS_02),"""",PRECO_VENDA_COMPRAS_02*FASCICULOS_02)"
    Range("E90").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(PRECO_VENDA_COMPRAS_03*FASCICULOS_03),"""",PRECO_VENDA_COMPRAS_03*FASCICULOS_03)"
    Range("F90").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(PRECO_VENDA_COMPRAS_04*FASCICULOS_04),"""",PRECO_VENDA_COMPRAS_04*FASCICULOS_04)"
    Range("G90").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(PRECO_VENDA_COMPRAS_05*FASCICULOS_05),"""",PRECO_VENDA_COMPRAS_05*FASCICULOS_05)"
    Range("H90").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(PRECO_VENDA_COMPRAS_06*FASCICULOS_06),"""",PRECO_VENDA_COMPRAS_06*FASCICULOS_06)"
    Range("I90").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(PRECO_VENDA_COMPRAS_07*FASCICULOS_07),"""",PRECO_VENDA_COMPRAS_07*FASCICULOS_07)"
    Range("J90").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(PRECO_VENDA_COMPRAS_08*FASCICULOS_08),"""",PRECO_VENDA_COMPRAS_08*FASCICULOS_08)"
    Range("J90").Select
    ActiveWindow.SmallScroll Down:=-180
    ActiveWindow.SelectedSheets.PrintPreview
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
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
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.SmallScroll ToRight:=-2
    Range("AB1:AP4").Select
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.SmallScroll Down:=21
    Range("AB1:AP32").Select
    ActiveWindow.SmallScroll Down:=-27
    ActiveSheet.PageSetup.PrintArea = "$AB$1:$AP$32"
    With ActiveSheet.PageSetup
        .LeftHeader = "Orçamento Nº &A"
        .CenterHeader = ""
        .RightHeader = "&F"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "&D"
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0.393700787401575)
        .BottomMargin = Application.InchesToPoints(0.354330708661417)
        .HeaderMargin = Application.InchesToPoints(0.236220472440945)
        .FooterMargin = Application.InchesToPoints(0.15748031496063)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = False
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.14)
        .RightMargin = Application.InchesToPoints()
        .TopMargin = Application.InchesToPoints()
        .BottomMargin = Application.InchesToPoints()
        .HeaderMargin = Application.InchesToPoints()
        .FooterMargin = Application.InchesToPoints()
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints()
        .RightMargin = Application.InchesToPoints(0.18)
        .TopMargin = Application.InchesToPoints()
        .BottomMargin = Application.InchesToPoints()
        .HeaderMargin = Application.InchesToPoints()
        .FooterMargin = Application.InchesToPoints()
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ActiveWindow.SelectedSheets.PrintPreview
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    ActiveWindow.SmallScroll ToRight:=14
    Range("AP10").Select
    ActiveCell.FormulaR1C1 = "5171.55"
    Range("AG5").Select
    Selection.Copy
    Range("AP11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AC7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AC6").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("AG6").Select
    Selection.Copy
    Range("AP16").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AG7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AP17").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWindow.SelectedSheets.PrintPreview
    ActiveWindow.SelectedSheets.PrintOut Copies:=1
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=30
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
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
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.SmallScroll Down:=-24
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Windows("Pasta1").Activate
    ActiveWindow.Close
End Sub
