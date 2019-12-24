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

    With Orcamento
        .strVendedor = strVendedor
        .strControle = strControle
        .strOperator = strOperator
        .strStatus = strStatus
    End With

End Sub

Function Transferencia(strOperacao As String, strDepartamento As String, strOrcamento As infOrcamento, strLocal As infBanco, strServer As infBanco)
Dim Connection As New ADODB.Connection
Dim rstSincronismo As ADODB.Recordset
Set rstSincronismo = New ADODB.Recordset
Dim strSql As String

''Is Internet Connected
If IsInternetConnected() = True Then
    Set Connection = OpenConnection(strLocal)
    '' Is Connected
    If Connection.State = 1 Then
        strSql = "SELECT DISTINCT tabela FROM qrySincronismo where sincronismo = '" & strOperacao & "' and dpto = '" & strDepartamento & "'"
        Call rstSincronismo.Open(strSql, Connection, adOpenStatic, adLockOptimistic)
        '' ENVIAR/RECEBER DADOS
        Do While Not rstSincronismo.EOF
            strSql = "SELECT * FROM " & rstSincronismo.Fields("tabela") & " WHERE controle = '" & strOrcamento.strControle & "' AND vendedor = '" & strOrcamento.strVendedor & "'"
            EnvioDeDados strLocal, strServer, strSql
            
            If strOperacao = "ENVIAR" Then
                '' server ( ENVIAR )
                loadOrcamento strOrcamento.strVendedor, strOrcamento.strControle
                loadOrcamento strOrcamento.strVendedor, strOrcamento.strControle, strStatus:=ID_STATUS(banco(1), Orcamento)
                Call admOrcamentoAtualizarEtapaADO(banco(0), Orcamento)
            ElseIf strOperacao = "RECEBER" Then
                '' local ( RECEBER )
                loadOrcamento strOrcamento.strVendedor, strOrcamento.strControle
                loadOrcamento strOrcamento.strVendedor, strOrcamento.strControle, strStatus:=ID_STATUS(banco(0), Orcamento)
                Call admOrcamentoAtualizarEtapaADO(banco(1), Orcamento)
            End If
            
            rstSincronismo.MoveNext
        Loop
    Else
        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "Falha na conexão com o banco. (" & strOperacao & ")"
    End If
    Connection.Close
Else
    ' no connected
    MsgBox "SEM INTERNET.", vbOKOnly + vbExclamation
End If
    
End Function

Sub EnvioDeDados(dbOrigem As infBanco, dbDestino As infBanco, strSql As String)
Dim Origem As New ADODB.Connection
Set Origem = OpenConnection(dbOrigem)
Dim rstOrigem As ADODB.Recordset
Set rstOrigem = New ADODB.Recordset
Dim Destino As New ADODB.Connection
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
Dim Connection As New ADODB.Connection
Dim rst As New ADODB.Recordset
    Set Connection = OpenConnection(strBanco)
    If Connection.State = 1 Then
        Call rst.Open("SELECT * FROM qryUsuarios WHERE usuario = '" & strOrcamento.strOperator & "'", Connection, adOpenStatic, adLockOptimistic)
        If Not rst.EOF Then
            Departamento = rst.Fields("DPTO").value
        Else
            Departamento = ""
        End If
    Else
        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "Falha na conexão com o banco."
    End If
    Connection.Close
End Function

Function ID_STATUS(strBanco As infBanco, strOrcamento As infOrcamento) As String
Dim Connection As New ADODB.Connection
Dim rst As New ADODB.Recordset
    Set Connection = OpenConnection(strBanco)
    If Connection.State = 1 Then
        Call rst.Open("SELECT ID_ETAPA FROM Orcamentos WHERE controle = '" & strOrcamento.strControle & "' AND vendedor = '" & strOrcamento.strVendedor & "'", Connection, adOpenStatic, adLockOptimistic)
        If Not rst.EOF Then
            ID_STATUS = rst.Fields("ID_ETAPA").value
        Else
            ID_STATUS = ""
        End If
    Else
        MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "Falha na conexão com o banco."
    End If
    Connection.Close
End Function
Sub updateStatus()
'loadBancos

'' local ( ENVIAR )
'loadOrcamento "VANESSA VICTORELLO", "117-14"
'loadOrcamento "VANESSA VICTORELLO", "117-14", strStatus:=ID_STATUS(banco(0), orcamento)
'Call admOrcamentoAtualizarEtapaADO(banco(1), orcamento)

''' server ( RECEBER )
'loadOrcamento "VANESSA VICTORELLO", "117-14"
'loadOrcamento "VANESSA VICTORELLO", "117-14", strStatus:=ID_STATUS(Banco(1), Orcamento)
'Call admOrcamentoAtualizarEtapaADO(Banco(0), Orcamento)


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
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(strBanco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admOrcamentoAtualizarEtapa"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_ETAPA", adVarChar, adParamInput, 50, strOrcamento.strStatus)
    .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, strOrcamento.strControle)
    .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, strOrcamento.strVendedor)
    Set rst = .Execute
End With
Connection.Close

End Sub
