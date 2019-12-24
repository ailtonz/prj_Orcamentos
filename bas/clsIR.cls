VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pAno As String
Private pDescricao As String
Private pFaixaInicial As String
Private pFaixaFinal As String
Private pAliquota As String
Private pParcelaDeduzir As String

''---------------
'' propriedades
''---------------

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get Ano() As String
    Ano = pAno
End Property

Public Property Let Ano(ByVal vAno As String)
    pAno = vAno
End Property

Public Property Get Descricao() As String
    Descricao = pDescricao
End Property

Public Property Let Descricao(ByVal vDescricao As String)
    pDescricao = vDescricao
End Property

Public Property Get FaixaInicial() As String
    FaixaInicial = pFaixaInicial
End Property

Public Property Let FaixaInicial(ByVal vFaixaInicial As String)
    pFaixaInicial = vFaixaInicial
End Property

Public Property Get FaixaFinal() As String
    FaixaFinal = pFaixaFinal
End Property

Public Property Let FaixaFinal(ByVal vFaixaFinal As String)
    pFaixaFinal = vFaixaFinal
End Property

Public Property Get Aliquota() As String
    Aliquota = pAliquota
End Property

Public Property Let Aliquota(ByVal vAliquota As String)
    pAliquota = vAliquota
End Property

Public Property Get ParcelaDeduzir() As String
    ParcelaDeduzir = pParcelaDeduzir
End Property

Public Property Let ParcelaDeduzir(ByVal vParcelaDeduzir As String)
    pParcelaDeduzir = vParcelaDeduzir
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

Public Sub add(ByVal rec As clsIR, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsIR)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsIR(ByVal i As Variant) As clsIR
    Set clsIR = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsIR
    Set Item = pCol(i)
End Property

''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As clsBancos, obj As clsIR) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admIRNew"
        .CommandType = adCmdStoredProc
    
        .Parameters.Append .CreateParameter("@NM_ANO", adVarChar, adParamInput, 50, obj.Ano)
        .Parameters.Append .CreateParameter("@NM_DESCRICAO", adVarChar, adParamInput, 50, obj.Descricao)
        .Parameters.Append .CreateParameter("@NM_FAIXA_INICIAL", adVarChar, adParamInput, 50, obj.FaixaInicial)
        .Parameters.Append .CreateParameter("@NM_FAIXA_FINAL", adVarChar, adParamInput, 50, obj.FaixaFinal)
        .Parameters.Append .CreateParameter("@NM_ALIQUOTA", adVarChar, adParamInput, 50, obj.Aliquota)
        .Parameters.Append .CreateParameter("@NM_PARCELA_DEDUZIR", adVarChar, adParamInput, 50, obj.ParcelaDeduzir)
    
        
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

Public Function Update(banco As clsBancos, obj As clsIR) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admIRUpd"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("@NM_DESCRICAO", adVarChar, adParamInput, 50, obj.Descricao)
        .Parameters.Append .CreateParameter("@NM_FAIXA_INICIAL", adVarChar, adParamInput, 50, obj.FaixaInicial)
        .Parameters.Append .CreateParameter("@NM_FAIXA_FINAL", adVarChar, adParamInput, 50, obj.FaixaFinal)
        .Parameters.Append .CreateParameter("@NM_ALIQUOTA", adVarChar, adParamInput, 50, obj.Aliquota)
        .Parameters.Append .CreateParameter("@NM_PARCELA_DEDUZIR", adVarChar, adParamInput, 50, obj.ParcelaDeduzir)
        
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

Public Function Delete(banco As clsBancos, obj As clsIR) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long


    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admIRDel"
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

Public Function getIR(banco As clsBancos) As clsIR
On Error GoTo Err_getIR
Dim colObj As New clsIR
Dim recObj As New clsIR

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryIR"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsIR
    With recObj
        .ID = rst.Fields("codCategoria")
        .Ano = rst.Fields("ANO_IR")
        .Descricao = rst.Fields("Descricao")
        .FaixaInicial = rst.Fields("Faixa_Inicial")
        .FaixaFinal = rst.Fields("Faixa_Final")
        .Aliquota = rst.Fields("Aliquota")
        .ParcelaDeduzir = rst.Fields("Parcela_Deduzir")
        colObj.all recObj
    End With
    rst.MoveNext
Loop



Exit_getIR:
    Set getIR = colObj
    Set recObj = Nothing
    Set colObj = Nothing
    
    Exit Function

Err_getIR:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getIR()" & vbNewLine & Err.Description)
    Resume Exit_getIR
End Function


