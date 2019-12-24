VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustoProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pPaginas As String
Private pValor As Currency
Private pTipo As String
Private pEstilo As String
Private pSubTipo As String

''---------------
'' propriedades
''---------------

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get Paginas() As String
    Paginas = pPaginas
End Property

Public Property Let Paginas(ByVal vPaginas As String)
    pPaginas = vPaginas
End Property

Public Property Get Valor() As Currency
    Valor = pValor
End Property

Public Property Let Valor(ByVal vValor As Currency)
    pValor = vValor
End Property

Public Property Get Tipo() As String
    Tipo = pTipo
End Property

Public Property Let Tipo(ByVal vTipo As String)
    pTipo = vTipo
End Property

Public Property Get Estilo() As String
    Estilo = pEstilo
End Property

Public Property Let Estilo(ByVal vEstilo As String)
    pEstilo = vEstilo
End Property

Public Property Get SubTipo() As String
    SubTipo = pSubTipo
End Property

Public Property Let SubTipo(ByVal vSubTipo As String)
    pSubTipo = vSubTipo
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

Public Sub add(ByVal rec As clsCustoProducao, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsCustoProducao)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsCustoProducao(ByVal i As Variant) As clsCustoProducao
    Set clsCustoProducao = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsCustoProducao
    Set Item = pCol(i)
End Property

''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As clsBancos, obj As clsCustoProducao) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admCustosProducaoNew"
        .CommandType = adCmdStoredProc
    
        .Parameters.Append .CreateParameter("@NM_PAGINAS", adVarChar, adParamInput, 50, obj.Paginas)
        .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, obj.Valor)
        .Parameters.Append .CreateParameter("@NM_TIPO", adVarChar, adParamInput, 50, obj.Tipo)
        .Parameters.Append .CreateParameter("@NM_ESTILO", adVarChar, adParamInput, 50, obj.Estilo)
        .Parameters.Append .CreateParameter("@NM_SUBTIPO", adVarChar, adParamInput, 50, obj.SubTipo)
        
    
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

Public Function Update(banco As clsBancos, obj As clsCustoProducao) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admCustosProducaoUpd"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("@NM_PAGINAS", adVarChar, adParamInput, 50, obj.Paginas)
        .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, obj.Valor)
        .Parameters.Append .CreateParameter("@NM_TIPO", adVarChar, adParamInput, 50, obj.Tipo)
        .Parameters.Append .CreateParameter("@NM_ESTILO", adVarChar, adParamInput, 50, obj.Estilo)
        .Parameters.Append .CreateParameter("@NM_SUBTIPO", adVarChar, adParamInput, 50, obj.SubTipo)
        
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

Public Function Delete(banco As clsBancos, obj As clsCustoProducao) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admCustosProducaoDel"
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

Public Function getCustosProducao(banco As clsBancos) As clsCustoProducao
On Error GoTo Err_getCustosProducao
Dim colObj As New clsCustoProducao
Dim recObj As New clsCustoProducao

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryCustosProducao"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsCustoProducao
    With recObj
        .ID = rst.Fields("codCategoria")
        .Paginas = rst.Fields("Paginas")
        .Valor = rst.Fields("Valor")
        .Tipo = rst.Fields("Tipo")
        .Estilo = rst.Fields("Estilo")
        .SubTipo = IIf(Not IsNull(rst.Fields("SubTipo")), rst.Fields("SubTipo"), "")
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getCustosProducao:
    Set getCustosProducao = colObj
    Set recObj = Nothing
    Set colObj = Nothing
    
    Exit Function

Err_getCustosProducao:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getCustosProducao()" & vbNewLine & Err.Description)
    Resume Exit_getCustosProducao

End Function
