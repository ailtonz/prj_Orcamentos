Attribute VB_Name = "basSynchronism"
Sub loadBancos()

    '' SERVER
    With banco(0)
        .strSource = Sheets("BANCOS").Range("C2")
        .strDriver = Sheets("BANCOS").Range("C3")
        .strLocation = Sheets("BANCOS").Range("C4")
        .strDatabase = Sheets("BANCOS").Range("C5")
        .strUser = Sheets("BANCOS").Range("C6")
        .strPassword = Sheets("BANCOS").Range("C7")
        .strPort = Sheets("BANCOS").Range("C8")
'        .strTabela = "qryUpdateSystem"
    End With
    
    '' LOCAL
    With banco(1)
        .strSource = Sheets("BANCOS").Range("F2")
        .strDriver = Sheets("BANCOS").Range("F3")
        .strLocation = Sheets("BANCOS").Range("F4")
        .strDatabase = Sheets("BANCOS").Range("F5")
        .strUser = Sheets("BANCOS").Range("F6")
        .strPassword = Sheets("BANCOS").Range("F7")
        .strPort = Sheets("BANCOS").Range("F8")
    End With

End Sub

Sub loadOrcamento(strVendedor As String, strControle As String, Optional strOperator As String, Optional strStatus As String)

    With orcamento
        .strVendedor = strVendedor
        .strControle = strControle
        .strOperator = strOperator
        .strStatus = strStatus
    End With

End Sub

Function Transferencia(strOperacao As String, strDepartamento As String, strOrcamento As infOrcamento, strLocal As infBanco, strServer As infBanco)
Dim connection As New ADODB.connection
Dim rstSincronismo As ADODB.Recordset
Set rstSincronismo = New ADODB.Recordset
Dim strSql As String

''Is Internet Connected
If IsInternetConnected() = True Then
    Set connection = OpenConnection(strLocal)
    '' Is Connected
    If connection.State = 1 Then
        strSql = "SELECT DISTINCT tabela FROM qrySincronismo where sincronismo = '" & strOperacao & "' and dpto = '" & strDepartamento & "'"
        Call rstSincronismo.Open(strSql, connection, adOpenStatic, adLockOptimistic)
        '' ENVIAR/RECEBER DADOS
        Do While Not rstSincronismo.EOF
            strSql = "SELECT * FROM " & rstSincronismo.Fields("tabela") & " WHERE controle = '" & strOrcamento.strControle & "' AND vendedor = '" & strOrcamento.strVendedor & "'"
            EnvioDeDados strLocal, strServer, strSql
            
            If strOperacao = "ENVIAR" Then
                '' server ( RECEBER )
                loadOrcamento strOrcamento.strVendedor, strOrcamento.strControle
                loadOrcamento strOrcamento.strVendedor, strOrcamento.strControle, strStatus:=ID_STATUS(banco(1), orcamento)
                Call admOrcamentoAtualizarEtapaADO(banco(0), orcamento)
            ElseIf strOperacao = "RECEBER" Then
                '' local ( ENVIAR )
                loadOrcamento strOrcamento.strVendedor, strOrcamento.strControle
                loadOrcamento strOrcamento.strVendedor, strOrcamento.strControle, strStatus:=ID_STATUS(banco(0), orcamento)
                Call admOrcamentoAtualizarEtapaADO(banco(1), orcamento)
            End If
            
            rstSincronismo.MoveNext
        Loop
    Else
        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "Falha na conexão com o banco. (" & strOperacao & ")"
    End If
    connection.Close
Else
    ' no connected
    MsgBox "SEM INTERNET.", vbOKOnly + vbExclamation
End If
    
End Function

Sub EnvioDeDados(dbOrigem As infBanco, dbDestino As infBanco, strSql As String)
Dim Origem As New ADODB.connection
Set Origem = OpenConnection(dbOrigem)
Dim rstOrigem As ADODB.Recordset
Set rstOrigem = New ADODB.Recordset
Dim Destino As New ADODB.connection
Set Destino = OpenConnection(dbDestino)
Dim rstDestino As ADODB.Recordset
Set rstDestino = New ADODB.Recordset
Dim fld As ADODB.Field
Dim NewFile As Boolean: NewFile = False
    
    Call rstOrigem.Open(strSql, Origem, , adLockOptimistic)
    
    If dbDestino.strDriver = "Access2003" Then
        Call rstDestino.Open(strSql, Destino, adOpenDynamic, adLockOptimistic, adCmdText)
    Else
        Call rstDestino.Open(strSql, Destino, adOpenDynamic, adLockOptimistic, adCmdText)
    End If
    
    '' SE Ñ EXISTE NO SERVER CADASTRAR
    If rstDestino.EOF Then
        NewFile = True
    End If

    Do While Not rstOrigem.EOF

        If NewFile Then
            rstDestino.AddNew
        End If

        For Each fld In rstDestino.Fields
            If fld.Name <> "Codigo" Then
                rstDestino(fld.Name).value = rstOrigem(fld.Name).value
            End If
        Next
        rstDestino.Update
        rstOrigem.MoveNext
    Loop
    
    rstDestino.Close
    rstOrigem.Close
    Destino.Close
    Origem.Close
End Sub

Function Departamento(strBanco As infBanco, strOrcamento As infOrcamento) As String
Dim connection As New ADODB.connection
Dim rst As New ADODB.Recordset
    Set connection = OpenConnection(strBanco)
    If connection.State = 1 Then
        Call rst.Open("SELECT * FROM qryUsuarios WHERE usuario = '" & strOrcamento.strOperator & "'", connection, adOpenStatic, adLockOptimistic)
        If Not rst.EOF Then
            Departamento = rst.Fields("DPTO").value
        Else
            Departamento = ""
        End If
    Else
        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "Falha na conexão com o banco."
    End If
    connection.Close
End Function

Function ID_STATUS(strBanco As infBanco, strOrcamento As infOrcamento) As String
Dim connection As New ADODB.connection
Dim rst As New ADODB.Recordset
    Set connection = OpenConnection(strBanco)
    If connection.State = 1 Then
        Call rst.Open("SELECT ID_ETAPA FROM Orcamentos WHERE controle = '" & strOrcamento.strControle & "' AND vendedor = '" & strOrcamento.strVendedor & "'", connection, adOpenStatic, adLockOptimistic)
        If Not rst.EOF Then
            ID_STATUS = rst.Fields("ID_ETAPA").value
        Else
            ID_STATUS = ""
        End If
    Else
        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "Falha na conexão com o banco."
    End If
    connection.Close
End Function
Sub updateStatus()
loadBancos

'' local ( ENVIAR )
'loadOrcamento "VANESSA VICTORELLO", "117-14"
'loadOrcamento "VANESSA VICTORELLO", "117-14", strStatus:=ID_STATUS(banco(0), orcamento)
'Call admOrcamentoAtualizarEtapaADO(banco(1), orcamento)

'' server ( RECEBER )
loadOrcamento "VANESSA VICTORELLO", "117-14"
loadOrcamento "VANESSA VICTORELLO", "117-14", strStatus:=ID_STATUS(banco(1), orcamento)
Call admOrcamentoAtualizarEtapaADO(banco(0), orcamento)


End Sub

Sub idStatus()

'loadBancos
'loadOrcamento "FABIANA", "134-14"
'loadOrcamento "FABIANA", "134-14", strStatus:=ID_STATUS(banco(1), orcamento)
'
''' local
''MsgBox ID_STATUS(banco(1),orcamento)
'
''' server
''MsgBox ID_STATUS(banco(0), orcamento)
'
'MsgBox orcamento.strStatus

End Sub

Sub admOrcamentoAtualizarEtapaADO(strBanco As infBanco, strOrcamento As infOrcamento)
Dim connection As New ADODB.connection
Set connection = OpenConnection(strBanco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = connection
    .CommandText = "admOrcamentoAtualizarEtapa"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_ETAPA", adVarChar, adParamInput, 50, strOrcamento.strStatus)
    .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, strOrcamento.strControle)
    .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, strOrcamento.strVendedor)
    Set rst = .Execute
End With
connection.Close

End Sub


Sub teste_admUpdateMoeda()
Dim sScript As String
Dim sValor As String: sValor = "13,56"
Dim sMoeda As String: sMoeda = "Dolar"
Dim sID As String: sID = "1"

sScript = "UPDATE admcategorias SET admcategorias.Descricao01 = '" & sValor & "' WHERE (((admcategorias.categoria)='" & sMoeda & "') AND ((admcategorias.codRelacao)=(SELECT admCategorias.codCategoria FROM admCategorias Where Categoria='MOEDA' and codRelacao = 0)))"

loadBancos
admUpdateMoeda banco(0), sID, sMoeda, sScript

End Sub

Function admUpdateMoeda(strBanco As infBanco, sID As String, sDescricao As String, sScript As String) As Boolean: admUpdateMoeda = True
On Error GoTo admUpdateMoeda_err
Dim cnn As New ADODB.connection
Set cnn = OpenConnection(strBanco)
Dim rst As ADODB.Recordset
Dim cmd As ADODB.Command

Set cmd = New ADODB.Command
With cmd
    .ActiveConnection = cnn
    .CommandText = "admUpdateMoeda"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_CATEGORIA", adVarChar, adParamInput, 100, "UPDATESYSTEM")
    .Parameters.Append .CreateParameter("@ATUALIZACAO_ID", adVarChar, adParamInput, 10, sID)
    .Parameters.Append .CreateParameter("@ATUALIZACAO_DESCRICAO", adVarChar, adParamInput, 100, sDescricao)
    .Parameters.Append .CreateParameter("@ATUALIZACAO_SCRIPT", adVarChar, adParamInput, 2000, sScript)
        
    Set rst = .Execute
End With
cnn.Close

admUpdateMoeda_Fim:
    Set cnn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    
    Exit Function
admUpdateMoeda_err:
    admUpdateMoeda = False
    MsgBox Err.Description
    Resume admUpdateMoeda_Fim

End Function

Function admNovoCliente_CADASTRAR(strBanco As infBanco, sID As String, sDescricao As String, sScript As String) As Boolean: admNovoCliente_CADASTRAR = True
On Error GoTo admNovoCliente_CADASTRAR_err
Dim cnn As New ADODB.connection
Set cnn = OpenConnection(strBanco)
Dim rst As ADODB.Recordset
Dim cmd As ADODB.Command

Set cmd = New ADODB.Command
With cmd
    .ActiveConnection = cnn
    .CommandText = "admNovoCliente_CADASTRAR"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_CATEGORIA", adVarChar, adParamInput, 100, "UPDATESYSTEM")
    .Parameters.Append .CreateParameter("@ATUALIZACAO_ID", adVarChar, adParamInput, 10, sID)
    .Parameters.Append .CreateParameter("@ATUALIZACAO_DESCRICAO", adVarChar, adParamInput, 100, sDescricao)
    .Parameters.Append .CreateParameter("@ATUALIZACAO_SCRIPT", adVarChar, adParamInput, 2000, sScript)
        
    Set rst = .Execute
End With
cnn.Close

admNovoCliente_CADASTRAR_Fim:
    Set cnn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    
    Exit Function
admNovoCliente_CADASTRAR_err:
    admNovoCliente_CADASTRAR = False
    MsgBox Err.Description
    Resume admNovoCliente_CADASTRAR_Fim

End Function


Function admNovoCliente_LIMPAR(strBanco As infBanco) As Boolean: admNovoCliente_LIMPAR = True
On Error GoTo admNovoCliente_LIMPAR_err
Dim cnn As New ADODB.connection
Set cnn = OpenConnection(strBanco)
Dim rst As ADODB.Recordset
Dim cmd As ADODB.Command

Set cmd = New ADODB.Command
With cmd
    .ActiveConnection = cnn
    .CommandText = "admNovoCliente_LIMPAR"
    .CommandType = adCmdStoredProc
    Set rst = .Execute
End With
cnn.Close

admNovoCliente_LIMPAR_Fim:
    Set cnn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    
    Exit Function
admNovoCliente_LIMPAR_err:
    admNovoCliente_LIMPAR = False
    MsgBox Err.Description
    Resume admNovoCliente_LIMPAR_Fim

End Function


Function admNovoCliente_ATUALIZAR(strBanco As infBanco) As Boolean: admNovoCliente_ATUALIZAR = True
On Error GoTo admNovoCliente_ATUALIZAR_err
Dim cnn As New ADODB.connection
Set cnn = OpenConnection(strBanco)
Dim rst As ADODB.Recordset
Dim cmd As ADODB.Command

Set cmd = New ADODB.Command
With cmd
    .ActiveConnection = cnn
    .CommandText = "admNovoCliente_ATUALIZAR"
    .CommandType = adCmdStoredProc
    Set rst = .Execute
End With
cnn.Close

admNovoCliente_ATUALIZAR_Fim:
    Set cnn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    
    Exit Function
admNovoCliente_ATUALIZAR_err:
    admNovoCliente_ATUALIZAR = False
    MsgBox Err.Description
    Resume admNovoCliente_ATUALIZAR_Fim

End Function

Sub UpdateSystem()

    loadBancos
    admUpdateSystem banco(0), banco(1), "1"
    admUpdateSystem banco(0), banco(1), "2"

End Sub

Function admUpdateSystem(strServidor As infBanco, strLocal As infBanco, idAtualizacao As Integer)
On Error GoTo admUpdateSystem_err
Dim cnnServidor As New ADODB.connection
Dim cnnLocal As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim rstLocal As New ADODB.Recordset
Dim cmdLocal As New ADODB.Command

    Set cnnServidor = OpenConnection(strBanco)
    If cnnServidor.State = 1 Then
        Call rst.Open("SELECT * FROM qryUpdateSystem WHERE ID = '" & idAtualizacao & "'", cnnServidor, adOpenStatic, adLockOptimistic)
        If Not rst.EOF Then
            Set cnnLocal = OpenConnection(strOrcamento)
            If cnnLocal.State = 1 Then
                cmdLocal.ActiveConnection = cnnLocal
                cmdLocal.CommandType = adCmdText
                cmdLocal.CommandText = rst.Fields("SCRIPT").value
                Set rstLocal = cmdLocal.Execute
'            Else
'                MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "ERROR DE FUNÇÃO: admUpdateSystem"
            End If
        End If
'    Else
'        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "ERROR DE FUNÇÃO: admUpdateSystem"
    End If
    
    cnnServidor.Close
    cnnLocal.Close
    
admUpdateSystem_Fim:
    Set cnnServidor = Nothing
    Set cnnLocal = Nothing
    Set rst = Nothing
    Set rstLocal = Nothing
    Set cmdLocal = Nothing
    
    Exit Function
admUpdateSystem_err:
    MsgBox Err.Description
    Resume admUpdateSystem_Fim
    
End Function



