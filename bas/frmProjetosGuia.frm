VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjetosGuia 
   Caption         =   "PROJETOS"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11385
   OleObjectBlob   =   "frmProjetosGuia.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProjetosGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private registro As New clsProjeto

Private Sub cmdGrands_Click()
    carregarOrcamento
    frmProjetosGrands.Show
End Sub

Private Sub UserForm_Activate()
    carregarDados
    carregarListagens
    carregarProjetos
End Sub

Private Sub cmdSalvar_Click()
    carregarProjeto Me.lstRegistros.Column(1)
    Call cmdCancelar_Click
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub lstRegistros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(Me.lstRegistros.value) Then
        carregarCampos Me.lstRegistros.Column(1)
        
        If Me.lstRegistros.Column(0) = 1 Or Me.lstRegistros.Column(0) = 3 Then
            Me.cmdGrands.Enabled = True
        Else
            Me.cmdGrands.Enabled = False
        End If
        
        Me.cmdSalvar.Caption = "SALVAR"
        Me.cmdSalvar.Enabled = True
    End If
End Sub

Private Sub carregarProjeto(c As Integer)
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

With Me
    
    ws.Cells(12, c).value = IIf(.chkVendido.value = True, "X", "")
    ws.Cells(13, c).value = .cboLinha.value
    ws.Cells(14, c).value = .txtFasciculos.value
    ws.Cells(15, c).value = .cboVendas.value
    ws.Cells(17, c).value = .cboIdiomas.value
    ws.Cells(18, c).value = .txtTiragem.value
    ws.Cells(19, c).value = .txtEspecificacao.value
    ws.Cells(20, c).value = .cboMoeda.value
    ws.Cells(21, c).value = .txtRoyalty_Percentual.value
    ws.Cells(22, c).value = .txtRoyalty_Valor.value
    ws.Cells(23, c).value = .txtReImpressao.value

End With

End Sub

Private Sub carregarProjetos()
Dim col As clsProjeto

With Me.lstRegistros
    .Clear
    .ColumnCount = 4
    .ColumnWidths = "20;0;70;90"

    For Each col In registro.Itens
        .AddItem col.ID
        .List(.ListCount - 1, 1) = col.ColunaExcel
        .List(.ListCount - 1, 2) = col.Tiragem
        .List(.ListCount - 1, 3) = col.Idioma
    Next col

End With

End Sub


Private Sub carregarCampos(c As Integer)

With Me
    
    .txtNumProjeto.value = Me.lstRegistros.Column(0)
    .chkVendido.value = IIf(Cells(12, c).value <> "", True, False)
    .cboLinha.value = Cells(13, c).value
    .txtFasciculos.value = Cells(14, c).value
    .cboVendas.value = Cells(15, c).value
    .cboIdiomas.value = Cells(17, c).value
    .txtTiragem.value = Cells(18, c).value
    .txtEspecificacao.value = Cells(19, c).value
    .cboMoeda.value = Cells(20, c).value
    .txtRoyalty_Percentual.value = Cells(21, c).value
    .txtRoyalty_Valor.value = Cells(22, c).value
    .txtReImpressao.value = Cells(23, c).value

End With

End Sub

Private Sub carregarDados()
Dim col As New clsProjeto

lCol = 8
c = 3
For x = 1 To lCol
    Set col = New clsProjeto
    With col
        .ID = x
        
        .Vendido = Cells(12, c).value
        
        .Linha = Cells(13, c).value
        
        .Fasciculos = Cells(14, c).value
        
        .Venda = Cells(15, c).value
        
        .Idioma = Cells(17, c).value
        
        .Tiragem = Cells(18, c).value
        
        .Especificacao = Cells(19, c).value
        
        .Moeda = Cells(20, c).value
        
        .RoyaltyPercentual = Cells(21, c).value
        
        .RoyaltyValor = Cells(22, c).value
        
        .ReImpressao = Cells(23, c).value
        
        .ColunaExcel = c
        
        c = c + 1
        
        registro.add col
    End With
Next x

End Sub

Private Sub carregarListagens()
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

ComboBoxUpdate "apoio", "IDIOMAS", Me.cboIdiomas
ComboBoxUpdate ws.Name, "VENDAS", Me.cboVendas
ComboBoxUpdate ws.Name, "MOEDA", Me.cboMoeda
ComboBoxUpdate ws.Name, "Linha", Me.cboLinha

End Sub


Private Sub carregarOrcamento()

With objOrc
    .Controle = ActiveSheet.Name
    .Vendedor = Range(GerenteDeContas)
    .NumProjeto = Me.txtNumProjeto.value
    .add objOrc
End With

End Sub
