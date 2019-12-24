VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pTipo As String

''---------------
'' propriedades
''---------------

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get Tipo() As String
    Tipo = pTipo
End Property

Public Property Let Tipo(ByVal vTipo As String)
    pTipo = vTipo
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

Public Sub add(ByVal rec As clsTipos, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsTipos)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsTipos(ByVal i As Variant) As clsTipos
    Set clsTipos = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsTipos
    Set Item = pCol(i)
End Property

''---------------
'' FUNÇÕES
''---------------

'Public Function Insert(banco As clsBancos, obj As clsTipos) As Boolean
'On Error GoTo Err_Insert
'Dim Connection As New ADODB.Connection
'Set Connection = OpenConnectionNEW(banco)
'Dim rst As ADODB.Recordset
'Dim cd As ADODB.Command
'Dim lRecordsAffected As Long
'
'    Set cd = New ADODB.Command
'    With cd
'        .ActiveConnection = Connection
'        .CommandText = "admEstiloNew"
'        .CommandType = adCmdStoredProc
'
'        .Parameters.Append .CreateParameter("@NM_ESTILO", adVarChar, adParamInput, 50, obj.Estilo)
'
'        Set rst = .Execute(lRecordsAffected)
'    End With
'    Connection.Close
'
'    If lRecordsAffected = 0 Then
'      Insert = False
'    Else
'      'Return a good return code
'      Insert = True
'    End If
'
'Exit_Insert:
'  'Clean up the connection to the database
'  Set cd = Nothing
'  Set rst = Nothing
'
'  Exit Function
'
'Err_Insert:
'  Call MsgBox("Class: " & TypeName(Me) & " , Function: Insert()"& vbNewLine & Err.Description)
'  Insert = False
'  Resume Exit_Insert
'
'End Function

'Public Function Update(banco As clsBancos, obj As clsTipos) As Boolean
'On Error GoTo Err_Update
'Dim Connection As New ADODB.Connection
'Set Connection = OpenConnectionNEW(banco)
'Dim rst As ADODB.Recordset
'Dim cd As ADODB.Command
'Dim lRecordsAffected As Long
'
'    Set cd = New ADODB.Command
'    With cd
'        .ActiveConnection = Connection
'        .CommandText = "admEstiloUpd"
'        .CommandType = adCmdStoredProc
'        .Parameters.Append .CreateParameter("@NM_ESTILO", adVarChar, adParamInput, 50, obj.Estilo)
'        .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)
'
'        Set rst = .Execute(lRecordsAffected)
'    End With
'    Connection.Close
'
'    If lRecordsAffected = 0 Then
'      Update = False
'    Else
'      'Return a good return code
'      Update = True
'    End If
'
'Exit_Update:
'  'Clean up the connection to the database
'  Set cd = Nothing
'  Set rst = Nothing
'
'  Exit Function
'
'Err_Update:
'  Call MsgBox("Class: " & TypeName(Me) & " , Function: Update()"& vbNewLine & Err.Description)
'  Update = False
'  Resume Exit_Update
'
'End Function

'Public Function Delete(banco As clsBancos, obj As clsTipos) As Boolean
'On Error GoTo Err_Delete
'Dim Connection As New ADODB.Connection
'Set Connection = OpenConnectionNEW(banco)
'Dim rst As ADODB.Recordset
'Dim cd As ADODB.Command
'Dim lRecordsAffected As Long
'
'    Set cd = New ADODB.Command
'    With cd
'        .ActiveConnection = Connection
'        .CommandText = "admEstiloDel"
'        .CommandType = adCmdStoredProc
'
'        .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)
'
'        Set rst = .Execute(lRecordsAffected)
'    End With
'    Connection.Close
'
'    If lRecordsAffected = 0 Then
'      Delete = False
'    Else
'      'Return a good return code
'      Delete = True
'    End If
'
'Exit_Delete:
'  'Clean up the connection to the database
'  Set cd = Nothing
'  Set rst = Nothing
'
'  Exit Function
'
'Err_Delete:
'  Call MsgBox("Class: " & TypeName(Me) & " , Function: Delete()"& vbNewLine & Err.Description)
'  Delete = False
'  Resume Exit_Delete
'
'End Function


Public Function getTipos(banco As clsBancos) As clsTipos
On Error GoTo Err_getTipos
Dim colObj As New clsTipos
Dim recObj As New clsTipos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryCustosProducaoTipos"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsTipos
    With recObj
        .Tipo = rst.Fields("Tipo")
        colObj.add recObj
    End With
    rst.MoveNext
Loop


Exit_getTipos:
    Set getTipos = colObj
    Set recObj = Nothing
    Set colObj = Nothing
    
    Exit Function

Err_getTipos:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getTipos()" & vbNewLine & Err.Description)
    Resume Exit_getTipos

End Function
