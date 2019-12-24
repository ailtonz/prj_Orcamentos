VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGrands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String

Private pNumProjeto As String
Private pNumControle As String
Private pVendedor As String

Private pNome As String
Private pProfissao As String
Private pValorLiquido As Currency
Private pCustoMedico As String
Private pCustoEditorFee As String

''---------------
'' propriedades
''---------------

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
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

Public Property Get Nome() As String
    Nome = pNome
End Property

Public Property Let Nome(ByVal vNome As String)
    pNome = vNome
End Property

Public Property Get Profissao() As String
    Profissao = pProfissao
End Property

Public Property Let Profissao(ByVal vProfissao As String)
    pProfissao = vProfissao
End Property

Public Property Get ValorLiquido() As Currency
    ValorLiquido = pValorLiquido
End Property

Public Property Let ValorLiquido(ByVal vValorLiquido As Currency)
    pValorLiquido = vValorLiquido
End Property

Public Property Get CustoMedico() As String
    CustoMedico = pCustoMedico
End Property

Public Property Let CustoMedico(ByVal vCustoMedico As String)
    pCustoMedico = vCustoMedico
End Property

Public Property Get CustoEditorFee() As String
    CustoEditorFee = pCustoEditorFee
End Property

Public Property Let CustoEditorFee(ByVal vCustoEditorFee As String)
    pCustoEditorFee = vCustoEditorFee
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

Public Sub add(ByVal rec As clsGrands, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsGrands)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Integer
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsGrand(ByVal index As Variant) As clsGrands
    Set clsGrand = pCol.Item(index)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsGrands
    Set Item = pCol(i)
End Property

''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As clsBancos, obj As clsGrands) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admGrandNew"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, obj.NumControle)
        .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, obj.Vendedor)
        .Parameters.Append .CreateParameter("@NumProjeto", adVarChar, adParamInput, 50, obj.NumProjeto)
    
        .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, obj.Profissao)
        .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, obj.Nome)
        .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, obj.ValorLiquido)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
        Insert = False
    Else
'        '' ATUALIZAÇÃO DE CUSTOS
'        UpdateCustos banco, Orcamento, Grand
        
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

Public Function Update(banco As clsBancos, obj As clsGrands) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admGrandUpd"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, obj.Profissao)
        .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, obj.Nome)
        .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, obj.ValorLiquido)
        
        .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
        Update = False
    Else
'        '' ATUALIZAÇÃO DE CUSTOS
'        UpdateCustos banco, Orcamento, Grand
    
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

Public Function Delete(banco As clsBancos, obj As clsGrands) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admGrandDel"
        .CommandType = adCmdStoredProc
    
        .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
        Delete = False
    Else
'        '' ATUALIZAÇÃO DE CUSTOS
'        UpdateCustos banco, Orcamento, Grand
        
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

Public Function UpdateCustos(banco As clsBancos, Orcamento As clsOrcamentos, grand As clsGrands) As Boolean
On Error GoTo Err_UpdateCustos
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admGrandsUpd"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, Orcamento.Controle)
        .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, Orcamento.Vendedor)
    
        .Parameters.Append .CreateParameter("@MEDICO", adVarChar, adParamInput, 50, grand.CustoMedico)
        .Parameters.Append .CreateParameter("@EDITOR_FEE", adVarChar, adParamInput, 50, grand.CustoEditorFee)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
      UpdateCustos = False
    Else
      'Return a good return code
      UpdateCustos = True
    End If
  
Exit_UpdateCustos:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_UpdateCustos:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: UpdateCustos() " & vbNewLine & Err.Description)
  UpdateCustos = False
  Resume Exit_UpdateCustos

End Function


Public Function Clone(banco As clsBancos, obj As clsGrands, i As Long) As Boolean
On Error GoTo Err_Clone
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admGrandNew"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, obj.NumControle)
        .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, obj.Vendedor)
        .Parameters.Append .CreateParameter("@NumProjeto", adVarChar, adParamInput, 50, i)
    
        .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, obj.Profissao)
        .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, obj.Nome)
        .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, obj.ValorLiquido)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
        Clone = False
    Else
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


Public Function getGrands(banco As clsBancos, orc As clsOrcamentos) As clsGrands
On Error GoTo Err_getGrands
Dim colObj As New clsGrands
Dim recObj As New clsGrands

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

Dim strTmp As String: strTmp = "select * from qryProjetosGrands where vendedor = '" & orc.Vendedor & "' and controle = '" & orc.Controle & "' and NumProjeto = '" & orc.NumProjeto & "' order by nome"

With cd
    .ActiveConnection = cnn
    .CommandText = strTmp
    .CommandType = adCmdText
    Set rst = .Execute
End With

'Saida strTmp, "getGrands.log"

Do While Not rst.EOF
    Set recObj = New clsGrands
    With recObj
        .ID = rst.Fields("id")
        .NumControle = rst.Fields("Controle")
        .Vendedor = rst.Fields("Vendedor")
        .Nome = rst.Fields("nome")
        .Profissao = rst.Fields("Profissao")
        .ValorLiquido = rst.Fields("Valor_Liquido")
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getGrands:
    Set getGrands = colObj
    Set recObj = Nothing
    Set colObj = Nothing
    
    Exit Function

Err_getGrands:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getGrands()" & vbNewLine & Err.Description)
    Resume Exit_getGrands

End Function

Public Function getGrandsIndex(banco As clsBancos, i As Long) As clsGrands
On Error GoTo Err_getGrandsIndex
Dim colObj As New clsGrands
Dim recObj As New clsGrands

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProjetosGrands where id = " & i & " order by nome"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsGrands
    With recObj
        .ID = rst.Fields("id")
        .Nome = rst.Fields("nome")
        .Profissao = rst.Fields("Profissao")
        .ValorLiquido = rst.Fields("Valor_Liquido")
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getGrandsIndex:
    Set getGrandsIndex = colObj
    Set recObj = Nothing
    Set colObj = Nothing
    
    Exit Function

Err_getGrandsIndex:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getGrandsIndex()" & vbNewLine & Err.Description)
    Resume Exit_getGrandsIndex

End Function

