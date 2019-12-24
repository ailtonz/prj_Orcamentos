VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProjetoCustos"
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

Private pIndexacao As String
Private pTraducao As String
Private pRevisaoOrtografica As String
Private pRevisaoMedica As String
Private pCriacao As String
Private pIlustracao As String
Private pRevisao  As String
Private pDiagramacao  As String
Private pMedico   As String
Private pGrafica  As String
Private pMidia    As String
Private pCorreio  As String
Private pUltimaCapa   As String
Private pImport   As String
Private pTransporteNacional   As String
Private pTransporteInternacional  As String
Private pSeguros  As String
Private pExtras   As String
Private pEditorFee    As String
Private pDespesaViagem    As String
Private pOutros As String

'' PROPRIEDADES
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

Public Property Get Indexacao() As String
    Indexacao = pIndexacao
End Property

Public Property Let Indexacao(ByVal vIndexacao As String)
    pIndexacao = vIndexacao
End Property

Public Property Get Traducao() As String
    Traducao = pTraducao
End Property

Public Property Let Traducao(ByVal vTraducao As String)
    pTraducao = vTraducao
End Property

Public Property Get RevisaoOrtografica() As String
    RevisaoOrtografica = pRevisaoOrtografica
End Property

Public Property Let RevisaoOrtografica(ByVal vRevisaoOrtografica As String)
    pRevisaoOrtografica = vRevisaoOrtografica
End Property

Public Property Get RevisaoMedica() As String
    RevisaoMedica = pRevisaoMedica
End Property

Public Property Let RevisaoMedica(ByVal vRevisaoMedica As String)
    pRevisaoMedica = vRevisaoMedica
End Property

Public Property Get Criacao() As String
    Criacao = pCriacao
End Property

Public Property Let Criacao(ByVal vCriacao As String)
    pCriacao = vCriacao
End Property

Public Property Get Ilustracao() As String
    Ilustracao = pIlustracao
End Property

Public Property Let Ilustracao(ByVal vIlustracao As String)
    pIlustracao = vIlustracao
End Property

Public Property Get Revisao() As String
    Revisao = pRevisao
End Property

Public Property Let Revisao(ByVal vRevisao As String)
    pRevisao = vRevisao
End Property

Public Property Get Diagramacao() As String
    Diagramacao = pDiagramacao
End Property

Public Property Let Diagramacao(ByVal vDiagramacao As String)
    pDiagramacao = vDiagramacao
End Property

Public Property Get Medico() As String
    Medico = pMedico
End Property

Public Property Let Medico(ByVal vMedico As String)
    pMedico = vMedico
End Property

Public Property Get Grafica() As String
    Grafica = pGrafica
End Property

Public Property Let Grafica(ByVal vGrafica As String)
    pGrafica = vGrafica
End Property

Public Property Get Midia() As String
    Midia = pMidia
End Property

Public Property Let Midia(ByVal vMidia As String)
    pMidia = vMidia
End Property

Public Property Get Correio() As String
    Correio = pCorreio
End Property

Public Property Let Correio(ByVal vCorreio As String)
    pCorreio = vCorreio
End Property

Public Property Get UltimaCapa() As String
    UltimaCapa = pUltimaCapa
End Property

Public Property Let UltimaCapa(ByVal vUltimaCapa As String)
    pUltimaCapa = vUltimaCapa
End Property

Public Property Get Import() As String
    Import = pImport
End Property

Public Property Let Import(ByVal vImport As String)
    pImport = vImport
End Property

Public Property Get TransporteNacional() As String
    TransporteNacional = pTransporteNacional
End Property

Public Property Let TransporteNacional(ByVal vTransporteNacional As String)
    pTransporteNacional = vTransporteNacional
End Property

Public Property Get TransporteInternacional() As String
    TransporteInternacional = pTransporteInternacional
End Property

Public Property Let TransporteInternacional(ByVal vTransporteInternacional As String)
    pTransporteInternacional = vTransporteInternacional
End Property

Public Property Get Seguros() As String
    Seguros = pSeguros
End Property

Public Property Let Seguros(ByVal vSeguros As String)
    pSeguros = vSeguros
End Property

Public Property Get Extras() As String
    Extras = pExtras
End Property

Public Property Let Extras(ByVal vExtras As String)
    pExtras = vExtras
End Property

Public Property Get EditorFee() As String
    EditorFee = pEditorFee
End Property

Public Property Let EditorFee(ByVal vEditorFee As String)
    pEditorFee = vEditorFee
End Property

Public Property Get DespesaViagem() As String
    DespesaViagem = pDespesaViagem
End Property

Public Property Let DespesaViagem(ByVal vDespesaViagem As String)
    pDespesaViagem = vDespesaViagem
End Property

Public Property Get Outros() As String
    Outros = pOutros
End Property

Public Property Let Outros(ByVal vOutros As String)
    pOutros = vOutros
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

Public Sub add(ByVal rec As clsProjetoCustos, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsProjetoCustos)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsProjetoCustos(ByVal i As Variant) As clsProjetoCustos
    Set clsProjetoCustos = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsProjetoCustos
    Set Item = pCol(i)
End Property


'' FUNÇÕES
Public Function Insert(banco As clsBancos, obj As clsProjetoCustos) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjCustosNew"
        .CommandType = adCmdStoredProc
    
        .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, obj.NumControle)
        .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, obj.Vendedor)
        .Parameters.Append .CreateParameter("@NumProjeto", adVarChar, adParamInput, 50, obj.NumProjeto)
        
        .Parameters.Append .CreateParameter("@NM_INDEXACAO", adVarChar, adParamInput, 50, obj.Indexacao)
        .Parameters.Append .CreateParameter("@NM_TRADUCAO", adVarChar, adParamInput, 50, obj.Traducao)
        .Parameters.Append .CreateParameter("@NM_REVISAO_ORTOGRAFICA", adVarChar, adParamInput, 50, obj.RevisaoOrtografica)
        .Parameters.Append .CreateParameter("@NM_REVISAO_MEDICA", adVarChar, adParamInput, 50, obj.RevisaoMedica)
        .Parameters.Append .CreateParameter("@NM_CRIACAO", adVarChar, adParamInput, 50, obj.Criacao)
        .Parameters.Append .CreateParameter("@NM_ILUSTRACAO", adVarChar, adParamInput, 50, obj.Ilustracao)
        .Parameters.Append .CreateParameter("@NM_REVISAO", adVarChar, adParamInput, 50, obj.Revisao)
        .Parameters.Append .CreateParameter("@NM_DIAGRAMACAO", adVarChar, adParamInput, 50, obj.Diagramacao)
        .Parameters.Append .CreateParameter("@NM_MEDICO", adVarChar, adParamInput, 50, obj.Medico)
        .Parameters.Append .CreateParameter("@NM_GRAFICA", adVarChar, adParamInput, 50, obj.Grafica)
        .Parameters.Append .CreateParameter("@NM_MIDIA", adVarChar, adParamInput, 50, obj.Midia)
        .Parameters.Append .CreateParameter("@NM_CORREIO", adVarChar, adParamInput, 50, obj.Correio)
        .Parameters.Append .CreateParameter("@NM_ULTIMA_CAPA", adVarChar, adParamInput, 50, obj.UltimaCapa)
        .Parameters.Append .CreateParameter("@NM_IMPORT", adVarChar, adParamInput, 50, obj.Import)
        .Parameters.Append .CreateParameter("@NM_TRANSPORTE_NACIONAL", adVarChar, adParamInput, 50, obj.TransporteNacional)
        .Parameters.Append .CreateParameter("@NM_TRANSPORTE_INTERNACIONAL", adVarChar, adParamInput, 50, obj.TransporteInternacional)
        .Parameters.Append .CreateParameter("@NM_SEGUROS", adVarChar, adParamInput, 50, obj.Seguros)
        .Parameters.Append .CreateParameter("@NM_EXTRAS", adVarChar, adParamInput, 50, obj.Extras)
        .Parameters.Append .CreateParameter("@NM_EDITOR_FEE", adVarChar, adParamInput, 50, obj.EditorFee)
        .Parameters.Append .CreateParameter("@NM_DESP_VIAGEM", adVarChar, adParamInput, 50, obj.DespesaViagem)
        .Parameters.Append .CreateParameter("@NM_OUTROS", adVarChar, adParamInput, 50, obj.Outros)

        .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)
        
    
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

Public Function Update(banco As clsBancos, obj As clsProjetoCustos) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjCustosUpd"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("@NM_INDEXACAO", adVarChar, adParamInput, 50, obj.Indexacao)
        .Parameters.Append .CreateParameter("@NM_TRADUCAO", adVarChar, adParamInput, 50, obj.Traducao)
        .Parameters.Append .CreateParameter("@NM_REVISAO_ORTOGRAFICA", adVarChar, adParamInput, 50, obj.RevisaoOrtografica)
        .Parameters.Append .CreateParameter("@NM_REVISAO_MEDICA", adVarChar, adParamInput, 50, obj.RevisaoMedica)
        .Parameters.Append .CreateParameter("@NM_CRIACAO", adVarChar, adParamInput, 50, obj.Criacao)
        .Parameters.Append .CreateParameter("@NM_ILUSTRACAO", adVarChar, adParamInput, 50, obj.Ilustracao)
        .Parameters.Append .CreateParameter("@NM_REVISAO", adVarChar, adParamInput, 50, obj.Revisao)
        .Parameters.Append .CreateParameter("@NM_DIAGRAMACAO", adVarChar, adParamInput, 50, obj.Diagramacao)
        .Parameters.Append .CreateParameter("@NM_MEDICO", adVarChar, adParamInput, 50, obj.Medico)
        .Parameters.Append .CreateParameter("@NM_GRAFICA", adVarChar, adParamInput, 50, obj.Grafica)
        .Parameters.Append .CreateParameter("@NM_MIDIA", adVarChar, adParamInput, 50, obj.Midia)
        .Parameters.Append .CreateParameter("@NM_CORREIO", adVarChar, adParamInput, 50, obj.Correio)
        .Parameters.Append .CreateParameter("@NM_ULTIMA_CAPA", adVarChar, adParamInput, 50, obj.UltimaCapa)
        .Parameters.Append .CreateParameter("@NM_IMPORT", adVarChar, adParamInput, 50, obj.Import)
        .Parameters.Append .CreateParameter("@NM_TRANSPORTE_NACIONAL", adVarChar, adParamInput, 50, obj.TransporteNacional)
        .Parameters.Append .CreateParameter("@NM_TRANSPORTE_INTERNACIONAL", adVarChar, adParamInput, 50, obj.TransporteInternacional)
        .Parameters.Append .CreateParameter("@NM_SEGUROS", adVarChar, adParamInput, 50, obj.Seguros)
        .Parameters.Append .CreateParameter("@NM_EXTRAS", adVarChar, adParamInput, 50, obj.Extras)
        .Parameters.Append .CreateParameter("@NM_EDITOR_FEE", adVarChar, adParamInput, 50, obj.EditorFee)
        .Parameters.Append .CreateParameter("@NM_DESP_VIAGEM", adVarChar, adParamInput, 50, obj.DespesaViagem)
        .Parameters.Append .CreateParameter("@NM_OUTROS", adVarChar, adParamInput, 50, obj.Outros)

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

Public Function Delete(banco As clsBancos, obj As clsProjetoCustos) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjCustosDel"
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


Public Function getCustos(banco As clsBancos) As clsProjetoCustos
Dim colObj As New clsProjetoCustos
Dim recObj As New clsProjetoCustos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProjetosCustos"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsProjetoCustos
    With recObj
        .ID = rst.Fields("ID")
        .NumControle = rst.Fields("CONTROLE")
        .Vendedor = rst.Fields("VENDEDOR")
        .NumProjeto = rst.Fields("NumPROJETO")
        
        .Indexacao = rst.Fields("INDEXACAO")
        .Traducao = rst.Fields("TRADUCAO")
        .RevisaoOrtografica = rst.Fields("REVISAO_ORTOGRAFICA")
        .RevisaoMedica = rst.Fields("REVISAO_MEDICA")
        .Criacao = rst.Fields("CRIACAO")
        .Ilustracao = rst.Fields("ILUSTRACAO")
        .Revisao = rst.Fields("REVISAO")
        .Diagramacao = rst.Fields("DIAGRAMACAO")
        .Medico = rst.Fields("MEDICO")
        .Grafica = rst.Fields("GRAFICA")
        .Midia = rst.Fields("MIDIA")
        .Correio = rst.Fields("CORREIO")
        .UltimaCapa = rst.Fields("ULTIMA_CAPA")
        .Import = rst.Fields("IMPORT")
        .TransporteNacional = rst.Fields("TRANSPORTE_NACIONAL")
        .TransporteInternacional = rst.Fields("TRANSPORTE_INTERNACIONAL")
        .Seguros = rst.Fields("SEGUROS")
        .Extras = rst.Fields("EXTRAS")
        .EditorFee = rst.Fields("EDITOR_FEE")
        .DespesaViagem = rst.Fields("DESP_VIAGEM")
        .Outros = rst.Fields("OUTROS")
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop

Set getCustos = colObj

Set recObj = Nothing
Set colObj = Nothing

End Function


Public Function getCustosProjeto(banco As clsBancos, obj As clsProjeto) As clsProjetoCustos
Dim colObj As New clsProjetoCustos
Dim recObj As New clsProjetoCustos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProjetosCustos  where controle = '" & obj.NumControle & "' and vendedor = '" & obj.Vendedor & "' and NumProjeto = '" & obj.NumProjeto & "'  order by NumPROJETO,id"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsProjetoCustos
    With recObj
        .ID = rst.Fields("ID")
        .NumControle = rst.Fields("CONTROLE")
        .Vendedor = rst.Fields("VENDEDOR")
        .NumProjeto = rst.Fields("NumPROJETO")
        
        .Indexacao = rst.Fields("INDEXACAO")
        .Traducao = rst.Fields("TRADUCAO")
        .RevisaoOrtografica = rst.Fields("REVISAO_ORTOGRAFICA")
        .RevisaoMedica = rst.Fields("REVISAO_MEDICA")
        .Criacao = rst.Fields("CRIACAO")
        .Ilustracao = rst.Fields("ILUSTRACAO")
        .Revisao = rst.Fields("REVISAO")
        .Diagramacao = rst.Fields("DIAGRAMACAO")
        .Medico = rst.Fields("MEDICO")
        .Grafica = rst.Fields("GRAFICA")
        .Midia = rst.Fields("MIDIA")
        .Correio = rst.Fields("CORREIO")
        .UltimaCapa = rst.Fields("ULTIMA_CAPA")
        .Import = rst.Fields("IMPORT")
        .TransporteNacional = rst.Fields("TRANSPORTE_NACIONAL")
        .TransporteInternacional = rst.Fields("TRANSPORTE_INTERNACIONAL")
        .Seguros = rst.Fields("SEGUROS")
        .Extras = rst.Fields("EXTRAS")
        .EditorFee = rst.Fields("EDITOR_FEE")
        .DespesaViagem = rst.Fields("DESP_VIAGEM")
        .Outros = rst.Fields("OUTROS")
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop

Set getCustosProjeto = colObj

Set recObj = Nothing
Set colObj = Nothing

End Function


Public Function getCustosProjetoIndex(banco As clsBancos, i As Long) As clsProjetoCustos
Dim colObj As New clsProjetoCustos
Dim recObj As New clsProjetoCustos

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProjetosCustos where id = " & i & " order by ID"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsProjetoCustos
    With recObj
        .ID = rst.Fields("ID")
        .NumControle = rst.Fields("CONTROLE")
        .Vendedor = rst.Fields("VENDEDOR")
        .NumProjeto = rst.Fields("NumPROJETO")
        
        .Indexacao = rst.Fields("INDEXACAO")
        .Traducao = rst.Fields("TRADUCAO")
        .RevisaoOrtografica = rst.Fields("REVISAO_ORTOGRAFICA")
        .RevisaoMedica = rst.Fields("REVISAO_MEDICA")
        .Criacao = rst.Fields("CRIACAO")
        .Ilustracao = rst.Fields("ILUSTRACAO")
        .Revisao = rst.Fields("REVISAO")
        .Diagramacao = rst.Fields("DIAGRAMACAO")
        .Medico = rst.Fields("MEDICO")
        .Grafica = rst.Fields("GRAFICA")
        .Midia = rst.Fields("MIDIA")
        .Correio = rst.Fields("CORREIO")
        .UltimaCapa = rst.Fields("ULTIMA_CAPA")
        .Import = rst.Fields("IMPORT")
        .TransporteNacional = rst.Fields("TRANSPORTE_NACIONAL")
        .TransporteInternacional = rst.Fields("TRANSPORTE_INTERNACIONAL")
        .Seguros = rst.Fields("SEGUROS")
        .Extras = rst.Fields("EXTRAS")
        .EditorFee = rst.Fields("EDITOR_FEE")
        .DespesaViagem = rst.Fields("DESP_VIAGEM")
        .Outros = rst.Fields("OUTROS")
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop

Set getCustosProjetoIndex = colObj

Set recObj = Nothing
Set colObj = Nothing

End Function


