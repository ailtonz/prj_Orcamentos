VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjetos 
   Caption         =   "PROJETOS"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11595
   OleObjectBlob   =   "frmProjetos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProjetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private registro As New clsProjeto
Private impressao As New clsProjetoImpressao
Private objGrand As New clsGrands

Private Sub UserForm_Activate()
    listarRegistros
    listarLinhas
    listarVendas
    listarIdiomas
    listarMoedas
    listarNumProjetos
End Sub

Private Sub cmdClone_Click()
    clonarRegistro
End Sub

Private Sub cmdCloneGrand_Click()
    clonarGrand
End Sub

Private Sub cmdCloneImpressao_Click()
    clonarImpressao
End Sub

Private Sub cmdCancelar_Click()
    limparCampos
End Sub

Private Sub cmdSalvar_Click()
Dim strMsg As String
Dim strTitulo As String

    If Me.cboProjeto.value = "" Or IsNull(Me.cboProjeto.value) Then
        strMsg = "ATENÇÃO: Por favor selecione o numero do projeto. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "ERRO!"
        
        MsgBox strMsg, vbInformation + vbOKOnly, strTitulo
        Me.cboProjeto.SetFocus
        Me.cboProjeto.DropDown
    Else
        If Me.cboLinha.value = "" Or IsNull(Me.cboLinha.value) Then
            strMsg = "ATENÇÃO: Por favor selecione uma linha de produto. " & Chr(10) & Chr(13) & Chr(13)
            strTitulo = "ERRO!"
            
            MsgBox strMsg, vbInformation + vbOKOnly, strTitulo
            Me.cboLinha.SetFocus
            Me.cboLinha.DropDown
        Else
            salvarRegistro
        End If
    End If
    
End Sub

Private Sub cmdGrands_Click()
    frmProjetosGrands.Show
End Sub

Private Sub cmdImpressao_Click()
    frmProjetosImpressao.Show
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
            Me.cmdSalvar.SetFocus
        End If
    End If
End Sub

Private Sub salvarRegistro()
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

Dim obj As clsProjeto
Dim orc As clsOrcamentos
       
Set orc = New clsOrcamentos
Set obj = New clsProjeto

carregarBanco
        
    With orc
        .Controle = ActiveSheet.Name
        .Vendedor = Range(GerenteDeContas)
        .add orc
    End With
    
    With obj
        .ID = Me.txtId.value
        
        .NumControle = ActiveSheet.Name
        .Vendedor = Range(GerenteDeContas)
        .NumProjeto = Me.cboProjeto.value
        
        .Linha = Me.cboLinha.Column(1)
        .Fasciculos = Me.txtFasciculos.value
        .Venda = Me.cboVendas.value
        .Idioma = Me.cboIdiomas.value
        .Tiragem = Me.txtTiragem.value
        
        .Especificacao = Me.txtEspecificacao.value
        .Moeda = Me.cboMoeda.value
        .RoyaltyPercentual = Me.txtRoyalty_Percentual.value
        .RoyaltyValor = Me.txtRoyalty_Valor.value
        .ReImpressao = Me.txtReImpressao.value
                
        .Vendido = IIf(Me.chkVendido.value = True, "x", "")
        
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


Set orc = Nothing
Set obj = Nothing
Set Bnc = Nothing

End Sub

Private Sub carregarProjeto()

With objOrc
    .Controle = ActiveSheet.Name
    .Vendedor = Range(GerenteDeContas)
    .NumProjeto = Me.cboProjeto.value
    .add objOrc
End With

End Sub

Private Sub clonarRegistro()
Dim obj As clsProjeto
Set obj = New clsProjeto

carregarBanco
            
If (obj.Clone(Bnc, registro.Itens(1), 9) = True) Then
    MsgBox "Cadastro clonado com sucesso!", vbInformation + vbOKOnly, "clone"
Else
    MsgBox "Não foi possivel realizar o clone do cadastro!", vbCritical + vbOKOnly, "clone - ERRO!"
End If
    
listarRegistros
limparCampos

Set obj = Nothing

End Sub

Private Sub clonarImpressao()
Dim obj As clsProjetoImpressao
Set obj = New clsProjetoImpressao

carregarProjeto
carregarBanco
       
Set impressao = obj.getImpressaoProjeto(Bnc, objOrc)

If (obj.Clone(Bnc, impressao.Itens(1), 9) = True) Then
    MsgBox "Cadastro clonado com sucesso!", vbInformation + vbOKOnly, "clone"
Else
    MsgBox "Não foi possivel realizar o clone do cadastro!", vbCritical + vbOKOnly, "clone - ERRO!"
End If
        
    listarRegistros
    limparCampos

Set obj = Nothing

End Sub


Private Sub clonarGrand()
Dim obj As clsGrands
Set obj = New clsGrands

carregarProjeto
carregarBanco
       
Set objGrand = obj.getGrands(Bnc, objOrc)

If (obj.Clone(Bnc, objGrand.Itens(1), 9) = True) Then
    MsgBox "Cadastro clonado com sucesso!", vbInformation + vbOKOnly, "clone"
Else
    MsgBox "Não foi possivel realizar o clone do cadastro!", vbCritical + vbOKOnly, "clone - ERRO!"
End If
        
    listarRegistros
    limparCampos

Set obj = Nothing

End Sub

Private Function mostrarRegistro() As Variant
Dim retVal As Variant
Dim strMsg As String

carregarCampos
    
strMsg = "PROJETO: " & vbTab & registro.Item(1).NumProjeto & vbNewLine
strMsg = strMsg & "LINHA: " & vbTab & vbTab & registro.Item(1).Linha & vbNewLine
strMsg = strMsg & "FASCICULOS : " & vbTab & registro.Item(1).Fasciculos & vbNewLine
strMsg = strMsg & "VENDA : " & vbTab & vbTab & registro.Item(1).Venda & vbNewLine
strMsg = strMsg & "IDIOMA : " & vbTab & vbTab & registro.Item(1).Idioma & vbNewLine
strMsg = strMsg & "TIRAGEM : " & vbTab & registro.Item(1).Tiragem & vbNewLine
    
retVal = MsgBox("Você deseja realmente EXCLUIR o registro abaixo:" & vbNewLine & _
        vbNewLine & strMsg, vbCritical + vbYesNo, "EXCLUSÃO DE REGISTRO!")
        
mostrarRegistro = retVal
            
Set retVal = Nothing

End Function

Private Sub limparCampos()

    Me.txtId.value = ""
    
    Me.cboProjeto.value = ""
    
    Me.cboLinha.value = ""
    Me.txtFasciculos.value = ""
    Me.cboVendas.value = ""
    Me.cboIdiomas.value = ""
    Me.txtTiragem.value = ""
    Me.txtEspecificacao.value = "NÃO"
    
    Me.cboMoeda.value = ""
    Me.txtRoyalty_Percentual.value = "0,00"
    Me.txtRoyalty_Valor.value = "0,00"
    Me.txtReImpressao.value = "NÃO"
       
    Me.chkVendido.value = False
    
    Me.cboProjeto.SetFocus
    
    Me.cmdSalvar.Caption = "NOVO"
    
    Me.cmdImpressao.Enabled = False
    Me.cmdGrands.Enabled = False
        
End Sub

Private Sub carregarCampos()
Dim obj As clsProjeto
Dim col As clsProjeto
Dim i As Long: i = Me.lstRegistros.value

carregarBanco
limparCampos

Set obj = New clsProjeto

Set col = obj.getProjetoIndex(Bnc, i)
Set registro = obj.getProjetoIndex(Bnc, i)

For Each obj In col.Itens
   
    Me.txtId.value = obj.ID
    
    Me.cboProjeto.value = obj.NumProjeto
    
    Me.cboLinha.SetFocus
    Me.cboLinha.SelText = obj.Linha
    
    Me.txtFasciculos.value = obj.Fasciculos
    
    Me.cboVendas.SetFocus
    Me.cboVendas.SelText = obj.Venda
    
    Me.cboIdiomas.SetFocus
    Me.cboIdiomas.SelText = obj.Idioma
        
    Me.txtTiragem.value = obj.Tiragem
    
    Me.txtEspecificacao.value = obj.Especificacao
    
    Me.cboMoeda.SetFocus
    Me.cboMoeda.SelText = obj.Moeda
    
    Me.txtRoyalty_Percentual.value = obj.RoyaltyPercentual
    Me.txtRoyalty_Valor.value = obj.RoyaltyValor
    Me.txtReImpressao.value = obj.ReImpressao
        
    Me.chkVendido.value = IIf(obj.Vendido <> "", True, False)
    
    Me.cmdImpressao.Enabled = True
    Me.cmdGrands.Enabled = True
    
    Me.cboProjeto.SetFocus

Next obj
        
With objOrc
    .Controle = ActiveSheet.Name
    .Vendedor = Range(GerenteDeContas)
    .NumProjeto = Me.cboProjeto.value
    .add objOrc
End With
        
End Sub

Private Sub listarRegistros()
Dim Prf As clsProjeto
Dim col As clsProjeto
Dim orc As clsOrcamentos

carregarBanco

Set orc = New clsOrcamentos
Set Prf = New clsProjeto

With orc
    .Controle = ActiveSheet.Name
    .Vendedor = Range(GerenteDeContas)
    .add orc
End With

Set col = Prf.getProjetosOrcamentos(Bnc, orc)

With Me.lstRegistros
    .Clear
    .ColumnCount = 3
    .ColumnWidths = "0;20;200;"
    
    For Each Prf In col.Itens
        .AddItem Prf.ID
        .List(.ListCount - 1, 1) = Prf.NumProjeto
        .List(.ListCount - 1, 2) = Prf.Linha

    Next Prf

End With

End Sub

Private Sub listarLinhas()
Dim Prf As clsLinhaProdutos
Dim col As clsLinhaProdutos

carregarBanco

Set Prf = New clsLinhaProdutos

Set col = Prf.getLinhas(Bnc)

With Me.cboLinha
    .Clear
    .ColumnCount = 2
    .ColumnWidths = "0;40"
    
    For Each Prf In col.Itens
        .AddItem Prf.ID
        .List(.ListCount - 1, 1) = Prf.Linha
    Next Prf

End With

End Sub

Private Sub listarVendas()
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

ComboBoxUpdate ws.Name, "VENDAS", Me.cboVendas

Set ws = Nothing

End Sub

Private Sub listarIdiomas()
Dim ws As Worksheet
Set ws = Worksheets("Apoio")

ComboBoxUpdate ws.Name, "IDIOMAS", Me.cboIdiomas

Set ws = Nothing

End Sub

Private Sub listarMoedas()
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

ComboBoxUpdate ws.Name, "MOEDA", Me.cboMoeda

Set ws = Nothing

End Sub


Private Sub listarNumProjetos()

Me.cboProjeto.List = Array("1", "2", "3", "4", "5", "6", "7", "8")

End Sub
