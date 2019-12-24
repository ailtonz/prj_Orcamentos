VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProposta 
   Caption         =   "IMPRESSÃO DE PROPOSTA"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11400
   OleObjectBlob   =   "frmProposta.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private registro As New clsProposta
Private objPrj As New clsProjeto

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub UserForm_Activate()
    carregarRegistro
    carregarCampos
    carregarListaProjetos
End Sub

Private Sub cmdImprimir_Click()
    carregarRegistro
    impressaoPropostas
End Sub

Public Sub impressaoPropostas()
Dim wsPrj As Worksheet
Set wsPrj = Worksheets(ActiveSheet.Name)

Dim obj As clsProposta
Set obj = New clsProposta

If (obj.GerarProposta(registro, objPrj) = True) Then
    MsgBox "Proposta impressa com sucesso!" & vbNewLine & vbNewLine & "LOCAL DO ARQUIVO: " & tmpProposta, vbInformation + vbOKOnly, "Impressão de proposta."
Else
    MsgBox "Não foi possivel realizar a impressão de proposta !", vbCritical + vbOKOnly, "Impressão de proposta - ERRO!"
End If

                      
Set obj = Nothing
Set objPrj = Nothing
Set Bnc = Nothing
End Sub

Private Sub carregarRegistro()
Dim wsBanco As Worksheet
Set wsBanco = Worksheets("BANCOS")

Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

Dim NumPaginas As Integer
Dim TotalTiragem As Long
Dim TotalGeral As Currency

'' SOMAR O NUMERO DE PAGINAS
ws.Range("C27:J27").Select
NumPaginas = Application.WorksheetFunction.Sum(Selection)


'' SOMAR O Total Tiragem
ws.Range("C18:J18").Select
TotalTiragem = Application.WorksheetFunction.Sum(Selection)


'' SOMAR O Total Geral
ws.Range("C75:J75").Select
TotalGeral = Application.WorksheetFunction.Sum(Selection)


With registro
    '' CARREGAR MODELO DE PROPOSTA
    .ArqCaminho = wsBanco.Range("O2").value
    .ArqNome = wsBanco.Range("O3").value

    '' CARREGAR TITULO DA PROPOSTA
    .Controle = ActiveSheet.Name
    .Cliente = ws.Range("C4").value
    .Responsavel = ws.Range("C5").value
    .Projeto = ws.Range("C6").value
    .Publisher = ws.Range("C8").value
    .Journal = ws.Range("C9").value
    
    .NumPaginas = NumPaginas
    .TotalTiragem = TotalTiragem
    .TotalGeral = TotalGeral
    
    '' CARREGAR GERENTE DE VENDAS
    .GerenteNome = wsBanco.Range("L2").value
    .GerenteTelefone = wsBanco.Range("L3").value
    .GerenteCelular01 = wsBanco.Range("L4").value
    .GerenteCelular02 = wsBanco.Range("L5").value
    .GerenteEmail = wsBanco.Range("L6").value
    
    .add registro
End With

Set ws = Nothing
Set wsBanco = Nothing

End Sub

Private Sub carregarCampos()
    With Me
        .txtControle.value = registro.Controle
        .txtCliente.value = registro.Cliente
        .txtResponsavel.value = registro.Responsavel
        .txtProjeto.value = registro.Projeto
        .txtJournal.value = registro.Journal
        .txtPublisher.value = registro.Publisher
        
        .txtGerenteNome.value = registro.GerenteNome
        .txtGerenteTelefone.value = registro.GerenteTelefone
        .txtGerenteCelular01.value = registro.GerenteCelular01
        .txtGerenteCelular02.value = registro.GerenteCelular02
        .txtGerenteEmail.value = registro.GerenteEmail
    End With
End Sub

Private Sub carregarListaProjetos()
Dim col As New clsProjeto
Dim c As Integer, l As Integer

Dim wsPrj As Worksheet
Set wsPrj = Worksheets(ActiveSheet.Name)

'' CONTAR LINHAS DE PRODUTOS
wsPrj.Range("C13").Select
lRow = wsPrj.Range(Selection, Selection.End(xlToRight)).Columns.count

With Me.lstRegistros
    .Clear
    .ColumnCount = 2
    .ColumnWidths = "100;200"
    
    l = 17
    c = 3
    For x = 1 To lRow
        '' TIRAGEM
        .AddItem Cells(l + 1, c).value
        '' IDIOMA
        .List(.ListCount - 1, 1) = Cells(l, c).value
        c = c + 1
    Next x
    
End With

c = 3
For x = 1 To lRow
    Set col = New clsProjeto
    With col
        .ID = x
        
        '' TIRAGEM
        .Tiragem = Cells(18, c).value
        
        '' FASCICULOS
        .Opcao = Cells(14, c).value
        
        '' IDIOMA
        .Idioma = Cells(17, c).value
        
        .PrcVendas = Cells(73, c).value
        .PrcTotal = Cells(75, c).value
        
        c = c + 1
        
        objPrj.add col
    End With
Next x

    
End Sub

