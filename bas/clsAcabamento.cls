VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAcabamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pColunaExcel As String

Private pNumProjeto As String
Private pNumControle As String
Private pVendedor As String

Private pAcabamento As String


'' PROPRIEDADES
Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get ColunaExcel() As String
    ColunaExcel = pColunaExcel
End Property

Public Property Let ColunaExcel(ByVal vColunaExcel As String)
    pColunaExcel = vColunaExcel
End Property

Public Property Get NumProjeto() As String
    NumProjeto = pNumProjeto
End Property

Public Property Let NumProjeto(ByVal vNumProjeto As String)
    pNumProjeto = vNumProjeto
End Property

Public Property Get NumControle() As String
    NumControle = pNumControle
End Property

Public Property Let NumControle(ByVal vNumControle As String)
    pNumControle = vNumControle
End Property

Public Property Get Vendedor() As String
    Vendedor = pVendedor
End Property

Public Property Let Vendedor(ByVal vVendedor As String)
    pVendedor = vVendedor
End Property

Public Property Get Acabamento() As String
    Acabamento = pAcabamento
End Property

Public Property Let Acabamento(ByVal vAcabamento As String)
    pAcabamento = vAcabamento
End Property



'' ADMINISTRAÇÃO
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

Public Sub add(ByVal rec As clsAcabamento, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsAcabamento)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsAcabamento(ByVal i As Variant) As clsAcabamento
    Set clsAcabamento = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsAcabamento
    Set Item = pCol(i)
End Property
