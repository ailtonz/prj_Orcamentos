Attribute VB_Name = "mod_Sincronismo"
Sub TESTE()
        
        addTasksShipping "MySQL", "AZS", "001-14"

End Sub

Sub Sincronismo_dados()

    RECEBER_DADOS
    ENVIAR_DADOS

End Sub

Function RemessaNew(strOperacao As String, strLocal As String, strServer As String)
Dim connection As New ADODB.connection
Dim rstAtualizacoes As ADODB.Recordset
Dim strTabelas(3) As String, strSQL As String, strUsuario As String, strControle As String
Set rstAtualizacoes = New ADODB.Recordset
strTabelas(1) = "Orcamentos"
strTabelas(2) = "OrcamentosAnexos"
strTabelas(3) = "OrcamentosCustos"

    ''Is Internet Connected
    If IsInternetConnected() = True Then
        Set connection = OpenConnection(strLocal)
        '' Is Connected
        If connection.State = 1 Then
            '' Tasks
            If strLocal = "Access2003" Then
                Call rstAtualizacoes.Open("SELECT * FROM OrcamentosAtualizacoes ", connection, adOpenStatic, adLockOptimistic)
            Else
                Call rstAtualizacoes.Open("SELECT * FROM OrcamentosAtualizacoes where usuario = '" & Range(NomeUsuario) & "'", connection, adOpenStatic, adLockOptimistic)
            End If
            
            '' VERIFICAR ATUALIZAÇÕES
            Do While Not rstAtualizacoes.EOF
                
                '' CADASTRAR ATUALIZAÇÕES NO SERVIDOR
                strUsuario = rstAtualizacoes.Fields("VENDEDOR").Value
                strControle = rstAtualizacoes.Fields("CONTROLE").Value
                If strOperacao = "ENVIAR" And strUsuario <> "" Then
                    addTasksShipping strServer, strUsuario, strControle
                End If
                        
                '' ENVIAR/RECEBER DADOS
                For x = 1 To UBound(strTabelas)
                    strSQL = "SELECT * FROM " & strTabelas(x) & " WHERE controle = '" & rstAtualizacoes.Fields("CONTROLE").Value & "' AND vendedor = '" & rstAtualizacoes.Fields("VENDEDOR").Value & "'"
                    EnvioDeDados strLocal, strServer, strSQL
                Next x

                rstAtualizacoes.MoveNext
            Loop
                        
            '' EXCLUIR ATUALIZAÇÕES LOCAIS
            delTasks strLocal, Range(NomeUsuario)
            MsgBox strOperacao & " ok!", vbInformation + vbOKOnly, strOperacao
        Else
            MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "Falha na conexão com o banco. (" & strOperacao & ")"
        End If
        connection.Close
    Else
        ' no connected
        MsgBox "SEM INTERNET.", vbOKOnly + vbExclamation
    End If
    
End Function


'Function RemessaNew(strOperacao As String, strLocal As String, strServer As String)
'Dim connection As New ADODB.connection
'Dim rstAtualizacoes As ADODB.Recordset
'Dim strTabelas(3) As String, strSQL As String, strUsuario As String, strControle As String
'Set rstAtualizacoes = New ADODB.Recordset
'
'strTabelas(1) = "Orcamentos"
'strTabelas(2) = "OrcamentosAnexos"
'strTabelas(3) = "OrcamentosCustos"
'
'    ''Is Internet Connected
'    If IsInternetConnected() = True Then
'        Set connection = OpenConnection(strLocal)
'
'        '' Is Connected
'        If connection.State = 1 Then
'            '' Tasks
''            If strLocal = "Access2003" Then
'
'                Call rstAtualizacoes.Open("SELECT * FROM OrcamentosAtualizacoes ", connection, adOpenStatic, adLockOptimistic)
''            Else
''
''                Call rstAtualizacoes.Open("SELECT * FROM OrcamentosAtualizacoes where usuario = '" & Range(NomeUsuario) & "'", connection, adOpenStatic, adLockOptimistic)
''            End If
'
'                '' VERIFICAR ATUALIZAÇÕES
'                Do While Not rstAtualizacoes.EOF
'
'
'                    If rstAtualizacoes.Fields("VENDEDOR").Value <> "" Or rstAtualizacoes.Fields("CONTROLE").Value <> "" Then
'
'                        '' CADASTRAR ATUALIZAÇÕES NO SERVIDOR
'                        strUsuario = rstAtualizacoes.Fields("VENDEDOR").Value
'                        strControle = rstAtualizacoes.Fields("CONTROLE").Value
'                        If strOperacao = "REMESSA" And strUsuario <> "" Then
'                            addTasksShipping strServer, strUsuario, strControle
'                        End If
'
'                        '' ENVIAR/RECEBER DADOS
'                        For x = 1 To UBound(strTabelas)
'                            strSQL = "SELECT * FROM " & strTabelas(x) & " WHERE controle = '" & rstAtualizacoes.Fields("CONTROLE").Value & "' AND vendedor = '" & rstAtualizacoes.Fields("VENDEDOR").Value & "'"
'                            EnvioDeDados strLocal, strServer, strSQL
'                        Next x
'
'                    End If
'
'                    rstAtualizacoes.MoveNext
'                Loop
'
''                delTasks strLocal, Range(NomeUsuario)
'
'                MsgBox strOperacao & " ok!", vbInformation + vbOKOnly, strOperacao
'
'        Else
'            MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "Falha na conexão com o banco. (" & strOperacao & ")"
'        End If
'        connection.Close
'    Else
'        ' no connected
'        MsgBox "SEM INTERNET.", vbOKOnly + vbExclamation, strOperacao
'    End If
'
'End Function

Sub RECEBER_DADOS()

    '' Retorno
    RemessaNew "RECEBER", "MySQL", "Access2003"

End Sub

Sub ENVIAR_DADOS()

    '' Remessa
    RemessaNew "ENVIAR", "Access2003", "MySQL"

End Sub
