VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLinhaProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pLinha As String
Private pMaximo As Currency
Private pMinimo As Currency
Private pEstilo As String

''---------------
'' propriedades
''---------------

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get Linha() As String
    Linha = pLinha
End Property

Public Property Let Linha(ByVal vLinha As String)
    pLinha = vLinha
End Property

Public Property Get Maximo() As Currency
    Maximo = pMaximo
End Property

Public Property Let Maximo(ByVal vMaximo As Currency)
    pMaximo = vMaximo
End Property

Public Property Get Minimo() As Currency
    Minimo = pMinimo
End Property

Public Property Let Minimo(ByVal vMinimo As Currency)
    pMinimo = vMinimo
End Property

Public Property Get Estilo() As String
    Estilo = pEstilo
End Property

Public Property Let Estilo(ByVal vEstilo As String)
    pEstilo = vEstilo
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

Public Sub add(ByVal rec As clsLinhaProdutos, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsLinhaProdutos)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsLinhaProdutos(ByVal i As Variant) As clsLinhaProdutos
    Set clsLinhaProdutos = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsLinhaProdutos
    Set Item = pCol(i)
End Property

''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As clsBancos, obj As clsLinhaProdutos) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admLinhaProdutoNew"
        .CommandType = adCmdStoredProc
    
        .Parameters.Append .CreateParameter("@NM_DESCRICAO", adVarChar, adParamInput, 50, obj.Linha)
        .Parameters.Append .CreateParameter("@NM_MAXIMO", adVarChar, adParamInput, 50, obj.Maximo)
        .Parameters.Append .CreateParameter("@NM_MINIMO", adVarChar, adParamInput, 50, obj.Minimo)
        .Parameters.Append .CreateParameter("@NM_ESTILO", adVarChar, adParamInput, 50, obj.Estilo)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
      Insert = False
    Else
      'Return a good return code
      Insert = True
    End If
  
Exit_Insert:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Insert:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Insert()" & vbNewLine & Err.Description)
  Insert = False
  Resume Exit_Insert

End Function

Public Function Update(banco As clsBancos, obj As clsLinhaProdutos) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admLinhaProdutoUpd"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("@NM_DESCRICAO", adVarChar, adParamInput, 50, obj.Linha)
        .Parameters.Append .CreateParameter("@NM_MAXIMO", adVarChar, adParamInput, 50, obj.Maximo)
        .Parameters.Append .CreateParameter("@NM_MINIMO", adVarChar, adParamInput, 50, obj.Minimo)
        .Parameters.Append .CreateParameter("@NM_ESTILO", adVarChar, adParamInput, 50, obj.Estilo)
        
        .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
      Update = False
    Else
      'Return a good return code
      Update = True
    End If
  
Exit_Update:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Update:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Update()" & vbNewLine & Err.Description)
  Update = False
  Resume Exit_Update

End Function

Public Function Delete(banco As clsBancos, obj As clsLinhaProdutos) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admLinhaProdutoDel"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
      Delete = False
    Else
      'Return a good return code
      Delete = True
    End If
  
Exit_Delete:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Delete:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Delete()" & vbNewLine & Err.Description)
  Delete = False
  Resume Exit_Delete

End Function

Public Function getLinhas(banco As clsBancos) As clsLinhaProdutos
On Error GoTo Err_getLinhas
Dim colObj As New clsLinhaProdutos
Dim recObj As New clsLinhaProdutos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryIndices_LinhaDeProdutos"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsLinhaProdutos
    With recObj
        .ID = rst.Fields("codCategoria")
        .Linha = rst.Fields("Descricao")
        .Maximo = rst.Fields("PORCENTAGEM_MAXIMA")
        .Minimo = rst.Fields("PORCENTAGEM_MINIMA")
        .Estilo = IIf(Not IsNull(rst.Fields("Estilo")), rst.Fields("Estilo"), "")
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getLinhas:
    Set getLinhas = colObj
    Set recObj = Nothing
    Set colObj = Nothing
    
    Exit Function

Err_getLinhas:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getLinhas()" & vbNewLine & Err.Description)
    Resume Exit_getLinhas
End Function

