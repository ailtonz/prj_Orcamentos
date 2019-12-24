VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProjetoImpressao"
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

Private pTipo As String
Private pPapel As String
Private pNumPaginas As String
Private pImpressao As String
Private pFormato As String


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

Public Property Get Tipo() As String
    Tipo = pTipo
End Property

Public Property Let Tipo(ByVal vTipo As String)
    pTipo = vTipo
End Property

Public Property Get Papel() As String
    Papel = pPapel
End Property

Public Property Let Papel(ByVal vPapel As String)
    pPapel = vPapel
End Property

Public Property Get NumPaginas() As String
    NumPaginas = pNumPaginas
End Property

Public Property Let NumPaginas(ByVal vNumPaginas As String)
    pNumPaginas = vNumPaginas
End Property

Public Property Get impressao() As String
    impressao = pImpressao
End Property

Public Property Let impressao(ByVal vImpressao As String)
    pImpressao = vImpressao
End Property

Public Property Get Formato() As String
    Formato = pFormato
End Property

Public Property Let Formato(ByVal vFormato As String)
    pFormato = vFormato
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

Public Sub add(ByVal rec As clsProjetoImpressao, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsProjetoImpressao)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsProjetoImpressao(ByVal i As Variant) As clsProjetoImpressao
    Set clsProjetoImpressao = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsProjetoImpressao
    Set Item = pCol(i)
End Property


'' FUNÇÕES
Public Function Insert(banco As clsBancos, obj As clsProjetoImpressao) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjImpressaoNew"
        .CommandType = adCmdStoredProc
    
        .Parameters.Append .CreateParameter("@NUMERO_CONTROLE", adVarChar, adParamInput, 50, obj.NumControle)
        .Parameters.Append .CreateParameter("@NOME_VENDEDOR", adVarChar, adParamInput, 50, obj.Vendedor)
        
        .Parameters.Append .CreateParameter("@NM_PROJETO", adVarChar, adParamInput, 50, obj.NumProjeto)
        .Parameters.Append .CreateParameter("@NM_TIPO", adVarChar, adParamInput, 50, obj.Tipo)
        
        .Parameters.Append .CreateParameter("@NM_PAPEL", adVarChar, adParamInput, 50, obj.Papel)
        .Parameters.Append .CreateParameter("@NM_NUM_PAGINAS", adVarChar, adParamInput, 50, obj.NumPaginas)
        .Parameters.Append .CreateParameter("@NM_IMPRESSAO", adVarChar, adParamInput, 50, obj.impressao)
        .Parameters.Append .CreateParameter("@NM_FORMATO", adVarChar, adParamInput, 50, obj.Formato)
        
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

Public Function Update(banco As clsBancos, obj As clsProjetoImpressao) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjImpressaoUpd"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("@NM_TIPO", adVarChar, adParamInput, 50, obj.Tipo)
        
        .Parameters.Append .CreateParameter("@NM_PAPEL", adVarChar, adParamInput, 50, obj.Papel)
        .Parameters.Append .CreateParameter("@NM_NUM_PAGINAS", adVarChar, adParamInput, 50, obj.NumPaginas)
        .Parameters.Append .CreateParameter("@NM_IMPRESSAO", adVarChar, adParamInput, 50, obj.impressao)
        .Parameters.Append .CreateParameter("@NM_FORMATO", adVarChar, adParamInput, 50, obj.Formato)
        
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

Public Function Delete(banco As clsBancos, obj As clsProjetoImpressao) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjImpressaoDel"
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


Public Function Clone(banco As clsBancos, obj As clsProjetoImpressao, i As Long) As Boolean
On Error GoTo Err_Clone
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjImpressaoNew"
        .CommandType = adCmdStoredProc
    
        .Parameters.Append .CreateParameter("@NUMERO_CONTROLE", adVarChar, adParamInput, 50, obj.NumControle)
        .Parameters.Append .CreateParameter("@NOME_VENDEDOR", adVarChar, adParamInput, 50, obj.Vendedor)
        .Parameters.Append .CreateParameter("@NM_PROJETO", adVarChar, adParamInput, 50, i)
        
        .Parameters.Append .CreateParameter("@NM_TIPO", adVarChar, adParamInput, 50, obj.Tipo)
        .Parameters.Append .CreateParameter("@NM_PAPEL", adVarChar, adParamInput, 50, obj.Papel)
        .Parameters.Append .CreateParameter("@NM_NUM_PAGINAS", adVarChar, adParamInput, 50, obj.NumPaginas)
        .Parameters.Append .CreateParameter("@NM_IMPRESSAO", adVarChar, adParamInput, 50, obj.impressao)
        .Parameters.Append .CreateParameter("@NM_FORMATO", adVarChar, adParamInput, 50, obj.Formato)
        
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
      Clone = False
    Else
      'Return a good return code
      Clone = True
    End If
  
Exit_Clone:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Clone:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Clone()" & vbNewLine & Err.Description)
  Clone = False
  Resume Exit_Clone

End Function

Public Function getImpressao(banco As clsBancos) As clsProjetoImpressao
Dim colObj As New clsProjetoImpressao
Dim recObj As New clsProjetoImpressao

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProjetosImpressao"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsProjetoImpressao
    With recObj
        .ID = rst.Fields("ID")
        .NumProjeto = rst.Fields("NumPROJETO")
        .NumControle = rst.Fields("CONTROLE")
        .Vendedor = rst.Fields("VENDEDOR")
        .Tipo = rst.Fields("TIPO")
        .Papel = rst.Fields("PAPEL")
        .NumPaginas = rst.Fields("NumPaginas")
        .impressao = rst.Fields("Impressao")
        .Formato = rst.Fields("Formato")
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop

Set getImpressao = colObj

Set recObj = Nothing
Set colObj = Nothing

End Function

Public Function getImpressaoProjeto(banco As clsBancos, obj As clsOrcamentos) As clsProjetoImpressao
Dim colObj As New clsProjetoImpressao
Dim recObj As New clsProjetoImpressao

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProjetosImpressao where controle = '" & obj.Controle & "' and vendedor = '" & obj.Vendedor & "' and NumPROJETO = '" & obj.NumProjeto & "'"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsProjetoImpressao
    With recObj
        .ID = rst.Fields("ID")
        .NumProjeto = rst.Fields("NumPROJETO")
        .NumControle = rst.Fields("CONTROLE")
        .Vendedor = rst.Fields("VENDEDOR")
        .Tipo = rst.Fields("TIPO")
        .Papel = rst.Fields("PAPEL")
        .NumPaginas = rst.Fields("NumPaginas")
        .impressao = rst.Fields("Impressao")
        .Formato = rst.Fields("Formato")
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop

Set getImpressaoProjeto = colObj

Set recObj = Nothing
Set colObj = Nothing

End Function

Public Function getImpressaoIndex(banco As clsBancos, i As Long) As clsProjetoImpressao
Dim colObj As New clsProjetoImpressao
Dim recObj As New clsProjetoImpressao

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProjetosImpressao where id = " & i & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsProjetoImpressao
    With recObj
        .ID = rst.Fields("ID")
        .NumProjeto = rst.Fields("NumPROJETO")
        .NumControle = rst.Fields("CONTROLE")
        .Vendedor = rst.Fields("VENDEDOR")
        .Tipo = rst.Fields("TIPO")
        .Papel = rst.Fields("PAPEL")
        .NumPaginas = rst.Fields("NumPaginas")
        .impressao = rst.Fields("Impressao")
        .Formato = rst.Fields("Formato")
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop

Set getImpressaoIndex = colObj

Set recObj = Nothing
Set colObj = Nothing

End Function
