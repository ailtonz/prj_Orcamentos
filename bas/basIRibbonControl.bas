Attribute VB_Name = "basIRibbonControl"
Sub Pesquisar(ByVal control As IRibbonControl)
    frmPesquisar.Show
End Sub

Sub cadastro(ByVal control As IRibbonControl)
'    If Range(GerenteDeContas) <> "" Then
'        frmCadastro.Show
'    End If
    
    frmDadosOrcamento.Show
    
End Sub

Sub AnexosArquivos(ByVal control As IRibbonControl)
    If Range(GerenteDeContas) <> "" Then
        frmAnexosArquivos.Show
    End If
End Sub

Sub EnviarReceber(ByVal control As IRibbonControl)
    frmEnviarReceber.Show
End Sub

Sub Indices(ByVal control As IRibbonControl)
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)
Dim strMsg As String
Dim strTitulo As String

If Range(GerenteDeContas) <> "" Then

    If LiberarIndice(strBanco, strUsuario) = False Then
        strMsg = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
        strMsg = strMsg & "Você não tem permissão para acessar este conteúdo. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Indices de calculos!"

        MsgBox strMsg, vbInformation + vbOKOnly, strTitulo
    Else

        LiberarIndice strBanco, strUsuario
        frmIndices.Show

    End If

End If

End Sub

Sub ENVIAR(ByVal control As IRibbonControl)
    frmEnviar.Show
End Sub

Sub nmMODELO(ByVal control As IRibbonControl)
'Dim strUsuario As String: strUsuario = Range(NomeUsuario)
'
'    If ActiveSheet.Name = strUsuario Then
'        Unload frmProjetosModelo
'        Exit Sub
'    Else
'        frmProjetosModelo.Show
'    End If
End Sub

Sub nmPROJETOS(ByVal control As IRibbonControl)
'Dim strUsuario As String: strUsuario = Range(NomeUsuario)
'
'    If ActiveSheet.Name = strUsuario Then
'        Unload frmProjetos
'        Exit Sub
'    Else
'        frmProjetos.Show
'    End If
    
    frmProjetosGuia.Show
    
End Sub

Sub nmIMPRESSOES(ByVal control As IRibbonControl)
'Dim strUsuario As String: strUsuario = Range(NomeUsuario)
'
'    If ActiveSheet.Name = strUsuario Then
'        Unload frmProjetos
'        Exit Sub
'    Else
'        frmProjetos.Show
'    End If
    
    frmImpressoes.Show
    
End Sub

Sub nmACABAMENTO(ByVal control As IRibbonControl)
'Dim strUsuario As String: strUsuario = Range(NomeUsuario)
'
'    If ActiveSheet.Name = strUsuario Then
'        Unload frmProjetos
'        Exit Sub
'    Else
'        frmProjetos.Show
'    End If
    
    frmAcabamento.Show
    
End Sub

Sub nmPROPOSTAS(ByVal control As IRibbonControl)
'Dim strUsuario As String: strUsuario = Range(NomeUsuario)
'
'    If ActiveSheet.Name = strUsuario Then
'        Unload frmProposta
'        MsgBox "ATENÇÃO: É necessário abrir um orçamento para imprimir uma proposta", vbInformation + vbOKOnly, "Imprimir proposta."
'        Exit Sub
'    Else
'        frmProposta.Show
'    End If
    
    frmProposta.Show
    
End Sub

Sub desbloqueio(ByVal control As IRibbonControl)

DesbloqueioDeGuia SenhaBloqueio

End Sub



Sub SimuladorCustos(ByVal control As IRibbonControl)
    MsgBox "EM TESTES", vbInformation + vbOKOnly, "Simulador de custos."
End Sub

Sub ControleGrand(ByVal control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

    If ActiveSheet.Name = strUsuario Then
        Unload frmGrand
        Exit Sub
    Else
        frmGrand.Show
    End If
    
End Sub

Sub EnviarDados(ByVal control As IRibbonControl)
'    '' CARREGAR BANCOS
'    loadBancos
'
'    '' CARREGAR ORÇAMENTO
'    loadOrcamento Sheets(ActiveSheet.Name).Range(GerenteDeContas), ActiveSheet.Name, Sheets(ActiveSheet.Name).Range(NomeUsuario)
'
'    Transferencia "ENVIAR", Departamento(banco(1), orcamento), orcamento, banco(1), banco(0)
      
    
    MsgBox "EM TESTES", vbInformation + vbOKOnly, "Enviar Dados."
End Sub

Sub ReceberDados(ByVal control As IRibbonControl)
'    '' CARREGAR BANCOS
'    loadBancos
'
'    '' CARREGAR ORÇAMENTO
'    loadOrcamento Sheets(ActiveSheet.Name).Range(GerenteDeContas), ActiveSheet.Name, Sheets(ActiveSheet.Name).Range(NomeUsuario)
'
'    Transferencia "RECEBER", Departamento(banco(1), orcamento), orcamento, banco(0), banco(1)
'
'    carregarOrcamentoSelecionado

    MsgBox "EM TESTES", vbInformation + vbOKOnly, "Receber Dados."
End Sub


Sub carregarOrcamentoSelecionado()

'        Application.ScreenUpdating = False
'
''        admOrcamentoFormulariosLimpar
'
'        carregarOrcamento Sheets(ActiveSheet.Name).Range(BancoLocal), ActiveSheet.Name, Sheets(ActiveSheet.Name).Range(GerenteDeContas)
'
'        admIntervalosDeEdicaoControle Sheets(ActiveSheet.Name).Range(BancoLocal), ActiveSheet.Name, Sheets(ActiveSheet.Name).Range(GerenteDeContas)
'
'        admOrcamentoFormulariosLiberar Sheets(ActiveSheet.Name).Range(NomeUsuario)
'
'        Sheets(ActiveSheet.Name).Range(InicioCursor).Select
'
'        Application.ScreenUpdating = True

End Sub


Sub modelo_teste(ByVal control As IRibbonControl)
   
'Dim strBanco As String: strBanco = Range(BancoLocal)
'Dim strControle As String
'Dim strUsuario As String
'
'    strControle = InputBox("Informe o numero de controle:", "Numero de controle", "082-14")
'    strUsuario = InputBox("Informe o nome do vendedor:", "Nome do vendedor", "azs")
'
'    admLimparAnexos
'
'    DesbloqueioDeGuia SenhaBloqueio
'
'    CarregarAnexoLinha strBanco, strControle, strUsuario, 3, 12
'    CarregarAnexoMoeda strBanco, strControle, strUsuario, 3, 16
'    CarregarAnexoVenda strBanco, strControle, strUsuario, 3, 19
'    CarregarAnexoDesconto strBanco, strControle, strUsuario, 3, 22
'
'    CarregarAnexoTraducao strBanco, strControle, strUsuario, 3, 29
'    CarregarAnexoRevisao strBanco, strControle, strUsuario, 3, 32
'    CarregarAnexoDiagramacao strBanco, strControle, strUsuario, 3, 35

'MarcaTexto InputBox("Informe a seleção:", "seleção", "")
    
End Sub

Sub SelecaoDeArea(ByVal control As IRibbonControl)
'Dim marcacao As String: marcacao = InputBox("Informe a seleção:", "seleção", "")
'
'    DesbloqueioDeGuia SenhaBloqueio
'
'    If marcacao <> "" Then
'        MarcaSelecao marcacao
'    End If
    
End Sub

Sub MenuChoice(control As IRibbonControl)

'DesbloqueioDeGuia SenhaBloqueio
'
'admIntervalosDeEdicaoLimparSelecao Range(BancoLocal)
'
'Select Case control.ID
'
'    Case "menuHistorico"
''        MarcaSelecao ""
'    Case "menuDesconto"
''        MarcaSelecao ""
'    Case "menuReCusto"
'        '' CUSTOS
'        MarcaSelecao "C37:J57"
'    Case "menuCancelado"
''        MarcaSelecao ""
'    Case "menuExcluido"
''        MarcaSelecao ""
'    Case "menuVendido"
''        MarcaSelecao ""
'    Case "menuNovo"
'
'        '' ORÇAMENTO
'        MarcaSelecao "C4,C5,G3:G4,C6,C8:J10,C13:J15,C17:J23,C61:J61,C23:J23"
'
'        '' ROYALTY PERCENTUAL
'        MarcaSelecao "C21:J21"
'
'        '' ROYALTY ESPECIE
'        MarcaSelecao "C22:J22"
'
'        '' IMPRESSÃO
'        MarcaSelecao "C25:J29,B31:J34"
'
'        '' DESCONTOS
'        MarcaSelecao "C61:J61"
'
'        '' PREÇO MKT
'        MarcaSelecao "C73:J73"
'
'        '' PREÇO COMPRA
'        MarcaSelecao "C80:J80"
'
'        '' DESCONTO COMPRA
'        MarcaSelecao "C79:J79"
'
'    Case "menuCusto"
'
'        '' CUSTOS
'        MarcaSelecao "C37:J57"
'
'    Case "menuOrcamento"
'
'        '' CUSTOS
'        MarcaSelecao "C37:J57"
'
'    Case "menuPreco"
'
'        '' ORÇAMENTO
'        MarcaSelecao "C4,C5,G3:G4,C6,C8:J10,C13:J15,C17:J23,C61:J61,C23:J23"
'
'        '' ROYALTY PERCENTUAL
'        MarcaSelecao "C21:J21"
'
'        '' ROYALTY ESPECIE
'        MarcaSelecao "C22:J22"
'
'        '' IMPRESSÃO
'        MarcaSelecao "C25:J29,B31:J34"
'
'        '' DESCONTOS
'        MarcaSelecao "C61:J61"
'
'        '' PREÇO MKT
'        MarcaSelecao "C73:J73"
'
'        '' PREÇO COMPRA
'        MarcaSelecao "C80:J80"
'
'        '' DESCONTO COMPRA
'        MarcaSelecao "C79:J79"
'
'End Select

End Sub

Sub Administracao(ByVal control As IRibbonControl)

'    frmProjetos.Show

'    frmProjetosModelo.Show


    frmPropostas.Show

End Sub

Sub teste_formatos()

Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

For Each cLoc In ws.Range("Projetos")
    If cLoc <> "" Then
        MsgBox cLoc
    End If
Next cLoc

End Sub
