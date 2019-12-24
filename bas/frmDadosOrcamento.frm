VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDadosOrcamento 
   Caption         =   "DADOS DO ORÇAMENTO"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10725
   OleObjectBlob   =   "frmDadosOrcamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDadosOrcamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private registro As New clsOrcamentos

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    carregarOrcamento
    Call cmdCancelar_Click
End Sub

Private Sub UserForm_Activate()

    DesbloqueioDeGuia SenhaBloqueio
    
    carregarListagens
    
    carregarDados
    carregarCampos
    
End Sub

Private Sub carregarDados()
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

With registro
    .Controle = ActiveSheet.Name
    .Vendedor = UCase(ws.Range("C3").value)
    .Cliente = UCase(ws.Range("C4").value)
    .Responsavel = UCase(ws.Range("C5").value)
    .Projeto = UCase(ws.Range("C6").value)
    .Publisher = UCase(ws.Range("C8").value)
    .Journal = UCase(ws.Range("C9").value)
    .Citacao = UCase(ws.Range("C10").value)
    .DataAbertura = ws.Range("G3").value
    .DataVenda = ws.Range("G4").value
    .add registro
End With

End Sub

Private Sub carregarOrcamento()
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)
    
With Me
    ws.Range("C4").value = .cboCliente.value
    ws.Range("C5").value = .txtResponsavel.value
    ws.Range("C6").value = .txtTitulo.value
    ws.Range("C8").value = .cboPublisher.value
    ws.Range("C9").value = .cboJournal.value
    ws.Range("C10").value = .txtCitacao.value
    ws.Range("G3").value = .txtDataAbertura.value
    ws.Range("G4").value = .txtDataVenda.value
End With

End Sub

Private Sub carregarCampos()

With Me
    .txtControle.value = registro.Controle
    .txtGerente.value = registro.Vendedor
    .cboCliente.value = registro.Cliente
    .txtResponsavel.value = registro.Responsavel
    .txtTitulo.value = registro.Projeto
    .cboPublisher.value = registro.Publisher
    .cboJournal.value = registro.Journal
    .txtCitacao.value = registro.Citacao
    .txtDataAbertura.value = registro.DataAbertura
    .txtDataVenda.value = registro.DataVenda
End With

End Sub

Private Sub carregarListagens()

ComboBoxUpdate "apoio", "Clientes", Me.cboCliente
ComboBoxUpdate "apoio", "Publisher", Me.cboPublisher
ComboBoxUpdate "apoio", "Journal", Me.cboJournal

End Sub
