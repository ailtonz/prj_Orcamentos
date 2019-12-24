VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pColunaExcel As String
Private pNumProjeto As String
Private pVendido As String
Private pLinha As String
Private pOpcao As String
Private pFasciculos As String
Private pVenda As String
Private pIdioma As String
Private pVolume As String
Private pTiragem As String
Private pEspecificacao As String
Private pMoeda As String
Private pRoyaltyPercentual As String
Private pRoyaltyValor As String
Private pReImpressao As String
Private pPrecoMKT As String
Private pDescontoPadrao As String
Private pPrecoTotalMKT As String
Private pArredondamentoGross As String
Private pArredondamentoMKT As String
Private pPrcVendas As String ' UNITARIO GROSS
Private pPrcTotal As String ' PREÇO TOTAL (FASCICULOS)



Private pNumControle As String
Private pVendedor As String

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

Public Property Get Vendido() As String
    Vendido = pVendido
End Property

Public Property Let Vendido(ByVal vVendido As String)
    pVendido = vVendido
End Property

Public Property Get Linha() As String
    Linha = pLinha
End Property

Public Property Let Linha(ByVal vLinha As String)
    pLinha = vLinha
End Property

Public Property Get Opcao() As String
    Opcao = pOpcao
End Property

Public Property Let Opcao(ByVal vOpcao As String)
    pOpcao = vOpcao
End Property

Public Property Get Fasciculos() As String
    Fasciculos = pFasciculos
End Property

Public Property Let Fasciculos(ByVal vFasciculos As String)
    pFasciculos = vFasciculos
End Property

Public Property Get Venda() As String
    Venda = pVenda
End Property

Public Property Let Venda(ByVal vVenda As String)
    pVenda = vVenda
End Property

Public Property Get Idioma() As String
    Idioma = pIdioma
End Property

Public Property Let Idioma(ByVal vIdioma As String)
    pIdioma = vIdioma
End Property

Public Property Get Volume() As String
    Volume = pVolume
End Property

Public Property Let Volume(ByVal vVolume As String)
    pVolume = vVolume
End Property

Public Property Get Tiragem() As String
    Tiragem = pTiragem
End Property

Public Property Let Tiragem(ByVal vTiragem As String)
    pTiragem = vTiragem
End Property

Public Property Get Especificacao() As String
    Especificacao = pEspecificacao
End Property

Public Property Let Especificacao(ByVal vEspecificacao As String)
    pEspecificacao = vEspecificacao
End Property

Public Property Get Moeda() As String
    Moeda = pMoeda
End Property

Public Property Let Moeda(ByVal vMoeda As String)
    pMoeda = vMoeda
End Property

Public Property Get RoyaltyPercentual() As String
    RoyaltyPercentual = pRoyaltyPercentual
End Property

Public Property Let RoyaltyPercentual(ByVal vRoyaltyPercentual As String)
    pRoyaltyPercentual = vRoyaltyPercentual
End Property

Public Property Get RoyaltyValor() As String
    RoyaltyValor = pRoyaltyValor
End Property

Public Property Let RoyaltyValor(ByVal vRoyaltyValor As String)
    pRoyaltyValor = vRoyaltyValor
End Property

Public Property Get ReImpressao() As String
    ReImpressao = pReImpressao
End Property

Public Property Let ReImpressao(ByVal vReImpressao As String)
    pReImpressao = vReImpressao
End Property

Public Property Get PrecoMKT() As String
    PrecoMKT = pPrecoMKT
End Property

Public Property Let PrecoMKT(ByVal vPrecoMKT As String)
    pPrecoMKT = vPrecoMKT
End Property

Public Property Get DescontoPadrao() As String
    DescontoPadrao = pDescontoPadrao
End Property

Public Property Let DescontoPadrao(ByVal vDescontoPadrao As String)
    pDescontoPadrao = vDescontoPadrao
End Property

Public Property Get PrecoTotalMKT() As String
    PrecoTotalMKT = pPrecoTotalMKT
End Property

Public Property Let PrecoTotalMKT(ByVal vPrecoTotalMKT As String)
    pPrecoTotalMKT = vPrecoTotalMKT
End Property

Public Property Get ArredondamentoGross() As String
    ArredondamentoGross = pArredondamentoGross
End Property

Public Property Let ArredondamentoGross(ByVal vArredondamentoGross As String)
    pArredondamentoGross = vArredondamentoGross
End Property

Public Property Get ArredondamentoMkt() As String
    ArredondamentoMkt = pArredondamentoMKT
End Property

Public Property Let ArredondamentoMkt(ByVal vArredondamentoMkt As String)
    pArredondamentoMKT = vArredondamentoMkt
End Property

Public Property Get PrcVendas() As String
    PrcVendas = pPrcVendas
End Property

Public Property Let PrcVendas(ByVal vPrcVendas As String)
    pPrcVendas = vPrcVendas
End Property

Public Property Get PrcTotal() As String
    PrcTotal = pPrcTotal
End Property

Public Property Let PrcTotal(ByVal vPrcTotal As String)
    pPrcTotal = vPrcTotal
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

Public Sub add(ByVal rec As clsProjeto, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsProjeto)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsProjeto(ByVal i As Variant) As clsProjeto
    Set clsProjeto = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsProjeto
    Set Item = pCol(i)
End Property


'' FUNÇÕES
Public Function Insert(banco As clsBancos, obj As clsProjeto) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjetoNew"
        .CommandType = adCmdStoredProc
    
        .Parameters.Append .CreateParameter("@NUMERO_CONTROLE", adVarChar, adParamInput, 50, obj.NumControle)
        .Parameters.Append .CreateParameter("@NOME_VENDEDOR", adVarChar, adParamInput, 50, obj.Vendedor)
        
        .Parameters.Append .CreateParameter("@NM_PROJETO", adVarChar, adParamInput, 50, obj.NumProjeto)
        .Parameters.Append .CreateParameter("@NM_LINHA", adVarChar, adParamInput, 50, obj.Linha)
        
        .Parameters.Append .CreateParameter("@NM_FASCICULOS", adVarChar, adParamInput, 50, obj.Fasciculos)
        .Parameters.Append .CreateParameter("@NM_VENDA", adVarChar, adParamInput, 50, obj.Venda)
        .Parameters.Append .CreateParameter("@NM_IDIOMA", adVarChar, adParamInput, 50, obj.Idioma)
        .Parameters.Append .CreateParameter("@NM_TIRAGEM", adVarChar, adParamInput, 50, obj.Tiragem)
        
        .Parameters.Append .CreateParameter("@NM_ESPECIFICACAO", adVarChar, adParamInput, 50, obj.Especificacao)
        .Parameters.Append .CreateParameter("@NM_MOEDA", adVarChar, adParamInput, 50, obj.Moeda)
        .Parameters.Append .CreateParameter("@NM_ROYALTY_PERCENTUAL", adVarChar, adParamInput, 50, obj.RoyaltyPercentual)
        .Parameters.Append .CreateParameter("@NM_ROYALTY_VALOR", adVarChar, adParamInput, 50, obj.RoyaltyValor)
        .Parameters.Append .CreateParameter("@NM_REIMPRESSAO", adVarChar, adParamInput, 50, obj.ReImpressao)
        
        .Parameters.Append .CreateParameter("@NM_PRECO_MKT", adVarChar, adParamInput, 50, obj.PrecoMKT)
        .Parameters.Append .CreateParameter("@NM_DESCONTO_PADRAO", adVarChar, adParamInput, 50, obj.DescontoPadrao)
        .Parameters.Append .CreateParameter("@NM_PRECO_TOTAL", adVarChar, adParamInput, 50, obj.PrecoTotalMKT)
        .Parameters.Append .CreateParameter("@NM_ARREDONDAMENTO_GROSS", adVarChar, adParamInput, 50, obj.ArredondamentoGross)
        .Parameters.Append .CreateParameter("@NM_ARREDONDAMENTO_MKT", adVarChar, adParamInput, 50, obj.ArredondamentoMkt)
        .Parameters.Append .CreateParameter("@NM_VENDIDO", adVarChar, adParamInput, 50, obj.Vendido)
        
    
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

Public Function Update(banco As clsBancos, obj As clsProjeto) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjetoUpd"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("@NM_LINHA", adVarChar, adParamInput, 50, obj.Linha)
        
        .Parameters.Append .CreateParameter("@NM_FASCICULOS", adVarChar, adParamInput, 50, obj.Fasciculos)
        .Parameters.Append .CreateParameter("@NM_VENDA", adVarChar, adParamInput, 50, obj.Venda)
        .Parameters.Append .CreateParameter("@NM_IDIOMA", adVarChar, adParamInput, 50, obj.Idioma)
        .Parameters.Append .CreateParameter("@NM_TIRAGEM", adVarChar, adParamInput, 50, obj.Tiragem)
        
        .Parameters.Append .CreateParameter("@NM_ESPECIFICACAO", adVarChar, adParamInput, 50, obj.Especificacao)
        .Parameters.Append .CreateParameter("@NM_MOEDA", adVarChar, adParamInput, 50, obj.Moeda)
        .Parameters.Append .CreateParameter("@NM_ROYALTY_PERCENTUAL", adVarChar, adParamInput, 50, obj.RoyaltyPercentual)
        .Parameters.Append .CreateParameter("@NM_ROYALTY_VALOR", adVarChar, adParamInput, 50, obj.RoyaltyValor)
        .Parameters.Append .CreateParameter("@NM_REIMPRESSAO", adVarChar, adParamInput, 50, obj.ReImpressao)
        
        .Parameters.Append .CreateParameter("@NM_PRECO_MKT", adVarChar, adParamInput, 50, obj.PrecoMKT)
        .Parameters.Append .CreateParameter("@NM_DESCONTO_PADRAO", adVarChar, adParamInput, 50, obj.DescontoPadrao)
        .Parameters.Append .CreateParameter("@NM_PRECO_TOTAL", adVarChar, adParamInput, 50, obj.PrecoTotalMKT)
        .Parameters.Append .CreateParameter("@NM_ARREDONDAMENTO_GROSS", adVarChar, adParamInput, 50, obj.ArredondamentoGross)
        .Parameters.Append .CreateParameter("@NM_ARREDONDAMENTO_MKT", adVarChar, adParamInput, 50, obj.ArredondamentoMkt)
        .Parameters.Append .CreateParameter("@NM_VENDIDO", adVarChar, adParamInput, 50, obj.Vendido)
        
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

Public Function Delete(banco As clsBancos, obj As clsProjeto) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjetoDel"
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

Public Function Clone(banco As clsBancos, obj As clsProjeto, i As Long) As Boolean
On Error GoTo Err_Clone
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admProjetoNew"
        .CommandType = adCmdStoredProc
    
        .Parameters.Append .CreateParameter("@NUMERO_CONTROLE", adVarChar, adParamInput, 50, obj.NumControle)
        .Parameters.Append .CreateParameter("@NOME_VENDEDOR", adVarChar, adParamInput, 50, obj.Vendedor)
        
        .Parameters.Append .CreateParameter("@NM_PROJETO", adVarChar, adParamInput, 50, i)
        .Parameters.Append .CreateParameter("@NM_LINHA", adVarChar, adParamInput, 50, obj.Linha)
        
        .Parameters.Append .CreateParameter("@NM_FASCICULOS", adVarChar, adParamInput, 50, obj.Fasciculos)
        .Parameters.Append .CreateParameter("@NM_VENDA", adVarChar, adParamInput, 50, obj.Venda)
        .Parameters.Append .CreateParameter("@NM_IDIOMA", adVarChar, adParamInput, 50, obj.Idioma)
        .Parameters.Append .CreateParameter("@NM_TIRAGEM", adVarChar, adParamInput, 50, obj.Tiragem)
        
        .Parameters.Append .CreateParameter("@NM_ESPECIFICACAO", adVarChar, adParamInput, 50, obj.Especificacao)
        .Parameters.Append .CreateParameter("@NM_MOEDA", adVarChar, adParamInput, 50, obj.Moeda)
        .Parameters.Append .CreateParameter("@NM_ROYALTY_PERCENTUAL", adVarChar, adParamInput, 50, obj.RoyaltyPercentual)
        .Parameters.Append .CreateParameter("@NM_ROYALTY_VALOR", adVarChar, adParamInput, 50, obj.RoyaltyValor)
        .Parameters.Append .CreateParameter("@NM_REIMPRESSAO", adVarChar, adParamInput, 50, obj.ReImpressao)
        
        .Parameters.Append .CreateParameter("@NM_PRECO_MKT", adVarChar, adParamInput, 50, obj.PrecoMKT)
        .Parameters.Append .CreateParameter("@NM_DESCONTO_PADRAO", adVarChar, adParamInput, 50, obj.DescontoPadrao)
        .Parameters.Append .CreateParameter("@NM_PRECO_TOTAL", adVarChar, adParamInput, 50, obj.PrecoTotalMKT)
        .Parameters.Append .CreateParameter("@NM_ARREDONDAMENTO_GROSS", adVarChar, adParamInput, 50, obj.ArredondamentoGross)
        .Parameters.Append .CreateParameter("@NM_ARREDONDAMENTO_MKT", adVarChar, adParamInput, 50, obj.ArredondamentoMkt)
        .Parameters.Append .CreateParameter("@NM_VENDIDO", adVarChar, adParamInput, 50, obj.Vendido)
        
    
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


Public Function getProjetos(banco As clsBancos) As clsProjeto
Dim colObj As New clsProjeto
Dim recObj As New clsProjeto

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProjetos"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsProjeto
    With recObj
        .ID = rst.Fields("ID")
        .NumControle = rst.Fields("CONTROLE")
        .Vendedor = rst.Fields("VENDEDOR")
        .NumProjeto = rst.Fields("NumPROJETO")
        .Linha = rst.Fields("LINHA")
        .Fasciculos = rst.Fields("FASCICULOS")
        .Venda = rst.Fields("VENDA")
        .Idioma = rst.Fields("IDIOMA")
        .Tiragem = rst.Fields("TIRAGEM")
        .Especificacao = rst.Fields("ESPECIFICACAO")
        .Moeda = rst.Fields("MOEDA")
        .RoyaltyPercentual = rst.Fields("ROYALTY_PERCENTUAL")
        .RoyaltyValor = rst.Fields("ROYALTY_VALOR")
        .ReImpressao = rst.Fields("RE_IMPRESSAO")
        .PrecoMKT = rst.Fields("PRECO_MKT")
        .DescontoPadrao = rst.Fields("DESCONTO_PADRAO")
        .PrecoTotalMKT = rst.Fields("PRECO_TOTAL")
        .ArredondamentoGross = rst.Fields("ARREDONDAMENTO_GROSS")
        .ArredondamentoMkt = rst.Fields("ARREDONDAMENTO_MKT")
        .Vendido = rst.Fields("VENDIDO")
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop

Set getProjetos = colObj

Set recObj = Nothing
Set colObj = Nothing

End Function

Public Function getProjetosOrcamentos(banco As clsBancos, obj As clsOrcamentos) As clsProjeto
Dim colObj As New clsProjeto
Dim recObj As New clsProjeto

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProjetos where controle = '" & obj.Controle & "' and vendedor = '" & obj.Vendedor & "' order by NumPROJETO,id"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsProjeto
    With recObj
        .ID = rst.Fields("ID")
        .NumControle = rst.Fields("CONTROLE")
        .Vendedor = rst.Fields("VENDEDOR")
        .NumProjeto = rst.Fields("NumPROJETO")
        .Linha = rst.Fields("LINHA")
        .Fasciculos = rst.Fields("FASCICULOS")
        .Venda = rst.Fields("VENDA")
        .Idioma = rst.Fields("IDIOMA")
        .Tiragem = rst.Fields("TIRAGEM")
        .Especificacao = rst.Fields("ESPECIFICACAO")
        .Moeda = rst.Fields("MOEDA")
        .RoyaltyPercentual = rst.Fields("ROYALTY_PERCENTUAL")
        .RoyaltyValor = rst.Fields("ROYALTY_VALOR")
        .ReImpressao = rst.Fields("RE_IMPRESSAO")
        .PrecoMKT = rst.Fields("PRECO_MKT")
        .DescontoPadrao = rst.Fields("DESCONTO_PADRAO")
        .PrecoTotalMKT = rst.Fields("PRECO_TOTAL")
        .ArredondamentoGross = rst.Fields("ARREDONDAMENTO_GROSS")
        .ArredondamentoMkt = rst.Fields("ARREDONDAMENTO_MKT")
        .Vendido = rst.Fields("VENDIDO")
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop

Set getProjetosOrcamentos = colObj

Set recObj = Nothing
Set colObj = Nothing

End Function

Public Function getProjetoIndex(banco As clsBancos, i As Long) As clsProjeto
Dim colObj As New clsProjeto
Dim recObj As New clsProjeto

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProjetos where id = " & i & " order by ID"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsProjeto
    With recObj
        .ID = rst.Fields("ID")
        .NumControle = rst.Fields("CONTROLE")
        .Vendedor = rst.Fields("VENDEDOR")
        .NumProjeto = rst.Fields("NumPROJETO")
        .Linha = rst.Fields("LINHA")
        .Fasciculos = rst.Fields("FASCICULOS")
        .Venda = rst.Fields("VENDA")
        .Idioma = rst.Fields("IDIOMA")
        .Tiragem = rst.Fields("TIRAGEM")
        .Especificacao = rst.Fields("ESPECIFICACAO")
        .Moeda = rst.Fields("MOEDA")
        .RoyaltyPercentual = rst.Fields("ROYALTY_PERCENTUAL")
        .RoyaltyValor = rst.Fields("ROYALTY_VALOR")
        .ReImpressao = rst.Fields("RE_IMPRESSAO")
        .PrecoMKT = rst.Fields("PRECO_MKT")
        .DescontoPadrao = rst.Fields("DESCONTO_PADRAO")
        .PrecoTotalMKT = rst.Fields("PRECO_TOTAL")
        .ArredondamentoGross = rst.Fields("ARREDONDAMENTO_GROSS")
        .ArredondamentoMkt = rst.Fields("ARREDONDAMENTO_MKT")
        .Vendido = rst.Fields("VENDIDO")
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop

Set getProjetoIndex = colObj

Set recObj = Nothing
Set colObj = Nothing

End Function


