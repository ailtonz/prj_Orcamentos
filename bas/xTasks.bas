Attribute VB_Name = "xTasks"
Sub addTasksShipping(strBanco As String, strVendedor As String, strControle As String)

Dim connection As New ADODB.connection
Set connection = OpenConnection(strBanco)

Dim rst As ADODB.Recordset

Dim cmd As ADODB.Command
Set cmd = New ADODB.Command

With cmd

    .ActiveConnection = connection
    
    .CommandText = "admOrcamentosAtualizacoesREMESSA"
    .CommandType = adCmdStoredProc
    
    .Parameters.Append cmd.CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, strVendedor)
    .Parameters.Append cmd.CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, strControle)
    
       
    Set rst = .Execute

End With

connection.Close

End Sub

Sub addTasksReturnADO(strBanco As String, strVendedor As String, strControle As String)

Dim connection As New ADODB.connection
Set connection = OpenConnection(strBanco)

Dim rst As ADODB.Recordset

Dim cmd As ADODB.Command
Set cmd = New ADODB.Command

With cmd

    .ActiveConnection = connection
    
    .CommandText = "admOrcamentosAtualizacoesRETORNO"
    .CommandType = adCmdStoredProc
    
    .Parameters.Append cmd.CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, strVendedor)
    .Parameters.Append cmd.CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, strControle)
    
       
    Set rst = .Execute

End With

connection.Close

End Sub

Sub addTasksReturn(strBanco As String, strVendedor As String, strControle As String)

Dim dbOrcamento As DAO.Database
Dim qdfadmUsuariosPermissoesExcluir As DAO.queryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(Range(BancoLocal), False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuariosPermissoesExcluir = dbOrcamento.QueryDefs("admRetorno")

With qdfadmUsuariosPermissoesExcluir

    .Parameters("NM_VENDEDOR") = strVendedor
    .Parameters("NM_CONTROLE") = strControle
    
    .Execute
    
End With

qdfadmUsuariosPermissoesExcluir.Close
dbOrcamento.Close

End Sub

Sub delTasks(strBanco As String, strUsuario As String)

Dim connection As New ADODB.connection
Set connection = OpenConnection(strBanco)

Dim rst As ADODB.Recordset

Dim cmd As ADODB.Command
Set cmd = New ADODB.Command

With cmd

    .ActiveConnection = connection
    
    .CommandText = "admOrcamentosAtualizacoesEXCLUSAO"
    .CommandType = adCmdStoredProc
    .Parameters.Append cmd.CreateParameter("@NM_USUARIO", adVarChar, adParamInput, 50, strUsuario)
    
    Set rst = .Execute

End With

connection.Close

End Sub

