VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjetosGrands 
   Caption         =   "GRANT DE PROJETO"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6435
   OleObjectBlob   =   "frmProjetosGrands.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProjetosGrands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
'    listarGrands

    If objOrc.NumProjeto = 1 Then
        carregarGrands "54"
    ElseIf objOrc.NumProjeto = 3 Then
        carregarGrands "57"
    End If
    
    listarProfissoes
    limparCampos
End Sub

Private Sub cmdCancelar_Click()
    limparCampos
End Sub

Private Sub cmdSalvar_Click()
    salvarRegistro
End Sub

Private Sub lstRegistros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNull(Me.lstRegistros.value) Then
        carregarCampos
        Me.cmdSalvar.Caption = "SALVAR"
    End If
    
End Sub

Private Sub lstRegistros_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Not IsNull(Me.lstRegistros.value) Then
            carregarCampos
            Me.cmdSalvar.Caption = "EXCLUIR"
        End If
    End If
End Sub

Private Sub txtValorLiquido_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtValorLiquido.value <> "" Then
        Me.txtValorLiquido.value = FormatCurrency(Me.txtValorLiquido.value)
    End If
    
End Sub

Private Sub salvarRegistro()

Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

Dim Grd As clsGrands
Set Grd = New clsGrands

carregarBanco
            
    With Grd
        .ID = Me.txtId.value
        
        .NumProjeto = objOrc.NumProjeto
        .NumControle = objOrc.Controle
        .Vendedor = objOrc.Vendedor
        
        .Profissao = Me.cboProfissao.value
        .ValorLiquido = Me.txtValorLiquido.value
'        .CustoMedico = ws.Range("C45").value
'        .CustoEditorFee = ws.Range("C55").value
        .add Grd
    End With
                      
    If Grd.ID = "" Then
        If (Grd.Insert(Bnc, Grd) = True) Then
            MsgBox "Cadastro realisado com sucesso!", vbInformation + vbOKOnly, "Cadastro"
        Else
            MsgBox "Não foi possivel realizar o cadastro!", vbCritical + vbOKOnly, "Cadastro - ERRO!"
        End If
        
    Else
    
        If Me.cmdSalvar.Caption = "SALVAR" Then
            If (Grd.Update(Bnc, Grd) = True) Then
                MsgBox "Alteração realizada com sucesso!", vbInformation + vbOKOnly, "Alteração"
            Else
                MsgBox "Não foi possivel realizar alteração!", vbCritical + vbOKOnly, "Alteração - ERRO!"
            End If
        Else
            If mostrarRegistro = vbYes Then
                If (Grd.Delete(Bnc, Grd) = True) Then
                    MsgBox "Exclusão realizada com sucesso!", vbInformation + vbOKOnly, "Exclusão"
                Else
                    MsgBox "Não foi possivel realizar Exclusão!", vbCritical + vbOKOnly, "Exclusão - ERRO!"
                End If
            End If
            
        End If
        
    End If
    
    listarGrands
    limparCampos

Set Grd = Nothing
Set Bnc = Nothing

End Sub

Private Function mostrarRegistro() As Variant
Dim retVal As Variant

    retVal = MsgBox("Você deseja realmente EXCLUIR o registro abaixo:" & vbNewLine & _
            vbNewLine & _
            "NOME: " & Me.lstRegistros.Column(1) & vbNewLine & _
            "PROFISSÃO: " & Me.lstRegistros.Column(2) & vbNewLine & _
            "VALOR : " & FormatCurrency(Me.lstRegistros.Column(3)) & vbNewLine, vbCritical + vbYesNo, "EXCLUSÃO DE REGISTRO!")
            
    mostrarRegistro = retVal
            
Set retVal = Nothing

End Function

Private Sub carregarCampos()

limparCampos
   
Me.txtId.value = Me.lstRegistros.Column(0)
        
Me.cboProfissao.SetFocus
Me.cboProfissao.SelText = Me.lstRegistros.Column(1)
    
Me.txtValorLiquido.value = FormatCurrency(Me.lstRegistros.Column(2))
    
Me.cboProfissao.SetFocus

                
End Sub


'Private Sub carregarCampos()
'Dim obj As clsGrands
'Dim col As clsGrands
'Dim i As Long: i = Me.lstRegistros.value
'
'carregarBanco
'limparCampos
'
'Set obj = New clsGrands
'
'Set col = obj.getGrandsIndex(Bnc, i)
'Set registro = obj.getGrandsIndex(Bnc, i)
'
'For Each obj In col.Itens
'
'    Me.txtId.value = obj.ID
'
'    Me.cboProfissao.SetFocus
'    Me.cboProfissao.SelText = obj.Profissao
'
'    Me.txtValorLiquido.value = FormatCurrency(obj.ValorLiquido)
'
'    Me.cboProfissao.SetFocus
'
'Next obj
'
'End Sub
    
Private Sub listarProfissoes()
Dim ws As Worksheet
Set ws = Worksheets("Apoio")

ComboBoxUpdate ws.Name, "PROFISSOES", Me.cboProfissao

Set ws = Nothing

End Sub

'Private Sub listarProfissoes()
'
'carregarBanco
'
'Dim Prf As clsProfissoes
'Set Prf = New clsProfissoes
'
'Dim col As clsProfissoes
'Set col = Prf.getProfissoes(Bnc)
'
'With Me.cboProfissao
'    .Clear
'    .Clear
'    .ColumnCount = 1
'    .ColumnWidths = "60"
'
'    For Each Prf In col.Itens
'        .AddItem Prf.Profissao
'    Next Prf
'
'End With
'
'End Sub

Private Sub listarGrands()
'Dim Prf As clsGrands
'Dim col As clsGrands
'
'carregarBanco
'
'Set Prf = New clsGrands
'Set col = Prf.getGrands(Bnc, objOrc)
'
'With Me.lstRegistros
'    .Clear
'    .ColumnCount = 4
'    .ColumnWidths = "0;200;60;0"
'
'    For Each Prf In col.Itens
'        .AddItem Prf.ID
'        .List(.ListCount - 1, 1) = Prf.Profissao
'        .List(.ListCount - 1, 2) = FormatCurrency(Prf.ValorLiquido)
'    Next Prf
'
'End With

'limparGrandsGuia
'listarGrandsGuia

End Sub

Private Sub limparCampos()

    Me.txtId.value = ""
    Me.cboProfissao.value = ""
    Me.txtValorLiquido.value = ""
    
    Me.cmdSalvar.Caption = "NOVO"
    
End Sub


Private Sub carregarGrands(c As Integer)
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

With Me.lstRegistros
    .Clear
    .ColumnCount = 3
    .ColumnWidths = "0;200;60"
    
    For x = 2 To 3
        .AddItem x
        .List(.ListCount - 1, 1) = ws.Cells(x, c).value
        .List(.ListCount - 1, 2) = FormatCurrency(ws.Cells(x, c + 1).value)
    Next x

End With

End Sub


