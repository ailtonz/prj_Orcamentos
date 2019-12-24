VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIR 
   Caption         =   "IMPOSTO DE RENDA"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   OleObjectBlob   =   "frmIR.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    listarRegistros
    limparCampos
End Sub

Private Sub txtDescricao_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtDescricao.value = UCase(Me.txtDescricao.value)
End Sub

Private Sub cmdCancelar_Click()
    limparCampos
End Sub

Private Sub cmdSalvar_Click()
    salvarRegistro
End Sub

Private Sub lstRegistros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '' CARREGAR REGISTRO
    If Not IsNull(Me.lstRegistros.value) Then
        carregarCampos
        Me.cmdSalvar.Caption = "SALVAR"
    End If
    
End Sub

Private Sub lstRegistros_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '' EXCLUIR REGISTRO
    If KeyCode = vbKeyDelete Then
        If Not IsNull(Me.lstRegistros.value) Then
            carregarCampos
            Me.cmdSalvar.Caption = "EXCLUIR"
            Me.cmdSalvar.SetFocus
        End If
    End If
End Sub

Private Sub salvarRegistro()

Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

Dim obj As clsIR

Set obj = New clsIR

carregarBanco

    With obj
        .ID = Me.txtId.value
        .Ano = Me.txtAno.value
        .Descricao = Me.txtDescricao.value
        .FaixaInicial = Me.txtFaixaInicial.value
        .FaixaFinal = Me.txtFaixaFinal.value
        .Aliquota = Me.txtAliquota.value
        .ParcelaDeduzir = Me.txtParcelaDeducao.value
        
        .add obj
    End With

    If obj.ID = "" Then
        If (obj.Insert(Bnc, obj) = True) Then
            MsgBox "Cadastro realizado com sucesso!", vbInformation + vbOKOnly, "Cadastro"
        Else
            MsgBox "Não foi possivel realizar o cadastro!", vbCritical + vbOKOnly, "Cadastro - ERRO!"
        End If

    Else

        If Me.cmdSalvar.Caption = "SALVAR" Then
            If (obj.Update(Bnc, obj) = True) Then
                MsgBox "Alteração realizada com sucesso!", vbInformation + vbOKOnly, "Alteração"
            Else
                MsgBox "Não foi possivel realizar alteração!", vbCritical + vbOKOnly, "Alteração - ERRO!"
            End If
        Else
            If mostrarRegistro = vbYes Then
                If (obj.Delete(Bnc, obj) = True) Then
                    MsgBox "Exclusão realizada com sucesso!", vbInformation + vbOKOnly, "Exclusão"
                Else
                    MsgBox "Não foi possivel realizar Exclusão!", vbCritical + vbOKOnly, "Exclusão - ERRO!"
                End If
            End If

        End If

    End If

    listarRegistros
    limparCampos

Set obj = Nothing
Set Bnc = Nothing

End Sub

Private Sub listarRegistros()
Dim Prf As clsIR
Dim col As clsIR

carregarBanco

Set Prf = New clsIR

Set col = Prf.getIR(Bnc)

With Me.lstRegistros
    .Clear
    .ColumnCount = 7
    .ColumnWidths = "0;60;90;0;0;0;0"
    
    For Each Prf In col.Itens
        .AddItem Prf.ID
        .List(.ListCount - 1, 1) = Prf.Ano
        .List(.ListCount - 1, 2) = Prf.Descricao
        .List(.ListCount - 1, 3) = Prf.FaixaInicial
        .List(.ListCount - 1, 4) = Prf.FaixaFinal
        .List(.ListCount - 1, 5) = Prf.Aliquota
        .List(.ListCount - 1, 6) = Prf.ParcelaDeduzir
    
    Next Prf

End With


End Sub

Private Sub limparCampos()
    
    Me.txtAno.Enabled = True

    Me.txtId.value = ""
    Me.txtAno.value = ""
    Me.txtDescricao.value = ""
    Me.txtFaixaInicial.value = ""
    Me.txtFaixaFinal.value = ""
    Me.txtAliquota.value = ""
    Me.txtParcelaDeducao.value = ""
    
    Me.cmdSalvar.Caption = "NOVO"
    Me.txtAno.SetFocus
    
End Sub

Private Sub carregarCampos()

    Me.txtAno.Enabled = False

    Me.txtId.value = Me.lstRegistros.value
    Me.txtAno.value = Me.lstRegistros.Column(1)
    Me.txtDescricao.value = Me.lstRegistros.Column(2)
    Me.txtFaixaInicial.value = Me.lstRegistros.Column(3)
    Me.txtFaixaFinal.value = Me.lstRegistros.Column(4)
    Me.txtAliquota.value = Me.lstRegistros.Column(5)
    Me.txtParcelaDeducao.value = Me.lstRegistros.Column(6)
        
End Sub

Private Function mostrarRegistro() As Variant
Dim retVal As Variant
Dim strAno As String: strAno = IIf(Not IsNull(Me.lstRegistros.Column(1)), Me.lstRegistros.Column(1), 0)
Dim strDescricao As String: strDescricao = IIf((Me.lstRegistros.Column(2)) <> "", Me.lstRegistros.Column(2), 0)
Dim strFaixaInicial As String: strFaixaInicial = IIf((Me.lstRegistros.Column(3)) <> "", Me.lstRegistros.Column(3), 0)
Dim strFaixaFinal As String: strFaixaFinal = IIf((Me.lstRegistros.Column(4)) <> "", Me.lstRegistros.Column(4), 0)
Dim strAliquota As String: strAliquota = IIf((Me.lstRegistros.Column(5)) <> "", Me.lstRegistros.Column(5), 0)
Dim strDeducao As String: strDeducao = IIf((Me.lstRegistros.Column(6)) <> "", Me.lstRegistros.Column(6), 0)


    retVal = MsgBox("Você deseja realmente EXCLUIR o registro abaixo:" & vbNewLine & _
            vbNewLine & _
            "ANO: " & strAno & vbNewLine & _
            "DESCRIÇÃO : " & strDescricao & vbNewLine & _
            "FAIXA INICIAL  : " & FormatCurrency(strFaixaInicial) & vbNewLine & _
            "FAIXA FINAL  : " & FormatCurrency(strFaixaFinal) & vbNewLine & _
            "ALIQUOTA  : " & strAliquota & vbNewLine & _
            "PARCELA DEDUÇÃO  : " & strDeducao & vbNewLine & _
            vbNewLine, vbCritical + vbYesNo, "EXCLUSÃO DE REGISTRO!")
            
    mostrarRegistro = retVal
            
Set retVal = Nothing

End Function

