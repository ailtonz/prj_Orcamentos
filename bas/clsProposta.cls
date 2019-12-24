VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String

Private pArqNome As String
Private pArqCaminho As String

Private pControle As String
Private pCliente As String
Private pResponsavel As String
Private pProjeto As String
Private pJournal As String
Private pAutor As String
Private pPublisher As String

Private pNumPaginas As String
Private pTotalGeral As String
Private pTotalTiragem As String

Private pGerenteNome As String
Private pGerenteTelefone As String
Private pGerenteCelular01 As String
Private pGerenteCelular02 As String
Private pGerenteEmail As String




''---------------
'' propriedades
''---------------

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get ArqCaminho() As String
    ArqCaminho = pArqCaminho
End Property

Public Property Let ArqCaminho(ByVal vArqCaminho As String)
    pArqCaminho = vArqCaminho
End Property

Public Property Get ArqNome() As String
    ArqNome = pArqNome
End Property

Public Property Let ArqNome(ByVal vArqNome As String)
    pArqNome = vArqNome
End Property

Public Property Get Controle() As String
    Controle = pControle
End Property

Public Property Let Controle(ByVal vControle As String)
    pControle = vControle
End Property

Public Property Get Cliente() As String
    Cliente = pCliente
End Property

Public Property Let Cliente(ByVal vCliente As String)
    pCliente = vCliente
End Property

Public Property Get Responsavel() As String
    Responsavel = pResponsavel
End Property

Public Property Let Responsavel(ByVal vResponsavel As String)
    pResponsavel = vResponsavel
End Property

Public Property Get Projeto() As String
    Projeto = pProjeto
End Property

Public Property Let Projeto(ByVal vProjeto As String)
    pProjeto = vProjeto
End Property

Public Property Get Journal() As String
    Journal = pJournal
End Property

Public Property Let Journal(ByVal vJournal As String)
    pJournal = vJournal
End Property

Public Property Get Autor() As String
    Autor = pAutor
End Property

Public Property Let Autor(ByVal vAutor As String)
    pAutor = vAutor
End Property

Public Property Get Publisher() As String
    Publisher = pPublisher
End Property

Public Property Let Publisher(ByVal vPublisher As String)
    pPublisher = vPublisher
End Property

Public Property Get NumPaginas() As String
    NumPaginas = pNumPaginas
End Property

Public Property Let NumPaginas(ByVal vNumPaginas As String)
    pNumPaginas = vNumPaginas
End Property

Public Property Get TotalGeral() As String
    TotalGeral = pTotalGeral
End Property

Public Property Let TotalGeral(ByVal vTotalGeral As String)
    pTotalGeral = vTotalGeral
End Property

Public Property Get TotalTiragem() As String
    TotalTiragem = pTotalTiragem
End Property

Public Property Let TotalTiragem(ByVal vTotalTiragem As String)
    pTotalTiragem = vTotalTiragem
End Property

Public Property Get GerenteNome() As String
    GerenteNome = pGerenteNome
End Property

Public Property Let GerenteNome(ByVal vGerenteNome As String)
    pGerenteNome = vGerenteNome
End Property


Public Property Get GerenteTelefone() As String
    GerenteTelefone = pGerenteTelefone
End Property

Public Property Let GerenteTelefone(ByVal vGerenteTelefone As String)
    pGerenteTelefone = vGerenteTelefone
End Property


Public Property Get GerenteCelular01() As String
    GerenteCelular01 = pGerenteCelular01
End Property

Public Property Let GerenteCelular01(ByVal vGerenteCelular01 As String)
    pGerenteCelular01 = vGerenteCelular01
End Property


Public Property Get GerenteCelular02() As String
    GerenteCelular02 = pGerenteCelular02
End Property

Public Property Let GerenteCelular02(ByVal vGerenteCelular02 As String)
    pGerenteCelular02 = vGerenteCelular02
End Property


Public Property Get GerenteEmail() As String
    GerenteEmail = pGerenteEmail
End Property

Public Property Let GerenteEmail(ByVal vGerenteEmail As String)
    pGerenteEmail = vGerenteEmail
End Property

''---------------
'' administração
''---------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set pCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCol = Nothing
End Sub

Public Sub add(ByVal rec As clsProposta, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsProposta)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsProposta(ByVal i As Variant) As clsProposta
    Set clsProposta = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsProposta
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function GerarProposta(ByVal obj As clsProposta, ByVal prj As clsProjeto) As Boolean
On Error GoTo Err_GerarProposta
Dim col As New clsProjeto, x As Integer: x = 1

Dim wdApp As Word.Application, wdDoc As Word.Document


On Error Resume Next

Set wdApp = GetObject(, "Word.Application")

'Word isn't already running
If Err.Number <> 0 Then
    Set wdApp = CreateObject("Word.Application")
End If

On Error GoTo 0

Set wdDoc = wdApp.Documents.Open(obj.ArqCaminho & obj.ArqNome, ReadOnly:=True)

wdApp.Visible = True

With wdDoc
    '' PROPOSTA
    .Bookmarks("N_CONTROLE").Range.Text = obj.Controle
    .Bookmarks("CLIENTE").Range.Text = obj.Cliente
    .Bookmarks("RESPONSAVEL").Range.Text = obj.Responsavel
    .Bookmarks("PROJETO").Range.Text = obj.Projeto
    .Bookmarks("JOURNAL").Range.Text = obj.Journal
    .Bookmarks("AUTOR").Range.Text = obj.Autor
    .Bookmarks("PUBLISHER").Range.Text = obj.Publisher
    
    '' PROJETOS
    For Each col In prj.Itens

        .Bookmarks("OPCAO_0" & x).Range.Text = col.Opcao
        .Bookmarks("IDIOMA_0" & x).Range.Text = col.Idioma
        .Bookmarks("VOLUME_0" & x).Range.Text = Format(col.Tiragem, "#,##0")
        .Bookmarks("PRC_VENDA_0" & x).Range.Text = Format(col.PrcVendas, "#,##0.00")
        .Bookmarks("PRC_TOTAL_0" & x).Range.Text = Format(col.PrcTotal, "#,##0.00")
        
        x = x + 1

    Next col
    
    .Bookmarks("TOTAL_VOLUME").Range.Text = Format(obj.TotalTiragem, "#,##0")
    .Bookmarks("TOTAL_GERAL").Range.Text = Format(obj.TotalGeral, "#,##0.00")
    .Bookmarks("N_PAGINAS").Range.Text = Format(obj.NumPaginas, "#,##0")
    
    '' GERENTE DE CONTAS
    .Bookmarks("G_CONTAS").Range.Text = obj.GerenteNome
    .Bookmarks("TELEFONE").Range.Text = obj.GerenteTelefone
    .Bookmarks("CELULAR_01").Range.Text = obj.GerenteCelular01
    .Bookmarks("CELULAR_02").Range.Text = obj.GerenteCelular02
    .Bookmarks("EMAIL").Range.Text = obj.GerenteEmail
End With

tmpProposta = pathDesktopAddress & Format(CDate(Now()), "dd-mm-yy_hh-mm") & "_" & obj.ArqNome

wdDoc.SaveAs tmpProposta


Exit_GerarProposta:
    wdDoc.Close
    wdApp.Application.Quit
    GerarProposta = True
    
    Exit Function

Err_GerarProposta:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: GerarProposta()" & vbNewLine & Err.Description)
    GerarProposta = False
    Resume Exit_GerarProposta

End Function

Private Function pathDesktopAddress() As String
    pathDesktopAddress = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
End Function

Sub UpdateBookmark(BookmarkToUpdate As String, TextToUse As String)
    Dim BMRange As Range
    Set BMRange = ActiveDocument.Bookmarks(BookmarkToUpdate).Range
    BMRange.Text = TextToUse
    ActiveDocument.Bookmarks.add BookmarkToUpdate, BMRange
End Sub

