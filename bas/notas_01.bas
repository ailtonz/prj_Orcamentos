Attribute VB_Name = "notas_01"

Sub testeIsInternetConnected()

Dim connection As New ADODB.connection

    If IsInternetConnected() = True Then
        Set connection = OpenConnection("MySQL")
        
        ' connected
        If connection.State = 1 Then
            MsgBox "ok"
        Else
            MsgBox "ñ,ok"

        End If

        connection.Close
    Else
        ' no connected
        MsgBox "SEM INTERNET.", vbOKOnly + vbExclamation
    End If


End Sub


Sub TesteCadastro()

Dim connectionString  As String
connectionString = "Driver={" & mysql_driver & "};Server=" & location & ";Database=" & Database & ";PORT=" & port & ";UID=" & user & ";PWD=" & password

Dim OpenConnection As ADODB.connection
Set OpenConnection = New ADODB.connection
OpenConnection.Open connectionString

Dim records As ADODB.Recordset
Set records = New ADODB.Recordset

Call records.Open("SELECT * FROM admCategorias where Descricao01 = 'teste.ok' ", connection, , adLockOptimistic)

'    connection.BeginTrans

    'records.AddNew

    records("Categoria").Value = "teste.categoria" & Now()

    records.Update
    records.Close

'    connection.CommitTrans

    connection.Close

End Sub


Sub testeFields()

Dim connectionString  As String
connectionString = "Driver={" & mysql_driver & "};Server=" & location & ";Database=" & Database & ";PORT=" & port & ";UID=" & user & ";PWD=" & password

Dim connection As ADODB.connection
Set connection = New ADODB.connection
connection.Open connectionString

Dim records As ADODB.Recordset
Set records = New ADODB.Recordset

Dim intLoop As Integer
Dim strSQL As String: strSQL = "SELECT * FROM Orcamentos"

Dim fld As ADODB.Field


    records.Open strSQL, connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    
    For Each fld In records.Fields
        Debug.Print fld.Value & "|"
    Next
    
    records.Close

    connection.Close


End Sub


Sub testeConnect()

Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)

Dim connectionString  As String
connectionString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & BaseDeDados & ";Uid=admin;PWD=" & SenhaBanco & ""


Dim connection As ADODB.connection
Set connection = New ADODB.connection
connection.Open connectionString

Dim rstSERVER As ADODB.Recordset
Set rstSERVER = New ADODB.Recordset


Dim strServer As String: strServer = "Select * from Orcamentos"


rstSERVER.Open strServer, connection, adOpenDynamic, adLockOptimistic, adCmdText

If rstSERVER.State = adStateOpen Then

    MsgBox "conexão ativa!"

Else

    MsgBox "conexão inativa!"

End If

rstSERVER.Close



End Sub

Sub testeCommand_MySQL()

Dim connection As New ADODB.connection
Set connection = OpenConnection("MySQL")

Dim rstExclusao As ADODB.Recordset

Dim cmdExclusao As ADODB.Command
Set cmdExclusao = New ADODB.Command

With cmdExclusao

    .ActiveConnection = connection
    
    .CommandText = "admOrcamentosAtualizacoesEXCLUSAO"
    .CommandType = adCmdStoredProc
    .Parameters.Append cmdExclusao.CreateParameter("@NM_USUARIO", adVarChar, adParamInput, 50, "ANDREA LIMA")
    
    Set rstExclusao = .Execute

End With

connection.Close


End Sub

Sub testeCommand_Access2003()

Dim connection As New ADODB.connection
Set connection = OpenConnection("Access2003")

Dim rstExclusao As ADODB.Recordset

Dim cmdExclusao As ADODB.Command
Set cmdExclusao = New ADODB.Command

With cmdExclusao

    .ActiveConnection = connection
    
    .CommandText = "admOrcamentosAtualizacoesEXCLUSAO"
    .CommandType = adCmdStoredProc
    .Parameters.Append cmdExclusao.CreateParameter("@NM_USUARIO", adVarChar, adParamInput, 50, "ailton")
    
    Set rstExclusao = .Execute

End With

connection.Close


End Sub


Sub TesteCadastro02()

    Dim connection As ADODB.connection
    Set connection = OpenConnection("MySQL")
    
    Dim rstAtualizacoes As ADODB.Recordset
    Set rstAtualizacoes = New ADODB.Recordset

    Call rstAtualizacoes.Open("SELECT * FROM OrcamentosAtualizacoes where USUARIO = 'AILTON' ", connection, , adLockOptimistic)
    
    If Not rstAtualizacoes.EOF Then
    
        MsgBox "Ok!"
    
    End If
    
    
    connection.BeginTrans
    
    'records.AddNew
        
    rstAtualizacoes("Categoria").Value = "teste.categoria" & Now()
    
    rstAtualizacoes.Update
    rstAtualizacoes.Close
    
    connection.CommitTrans
    
    connection.Close

End Sub


'
'Sub tarefasLocais()
'Dim dbOrcamento As DAO.Database
'Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)
'Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
'
'Dim rstOrigem As DAO.Recordset
'Dim strOrigem As String: strOrigem = "SELECT DISTINCT vendedor,controle  FROM OrcamentosAtualizacoes"
'Set rstOrigem = dbOrcamento.OpenRecordset(strOrigem)
'
'Dim qdfadmOrcamentoAtualizacoes As DAO.queryDef
'Set qdfadmOrcamentoAtualizacoes = dbOrcamento.QueryDefs("admOrcamentosAtualizacoesEXCLUSAO")
'
'Dim strTabelas(3) As String, strSQL As String
'
'strTabelas(1) = "Orcamentos"
'strTabelas(2) = "OrcamentosAnexos"
'strTabelas(3) = "OrcamentosCustos"
'
'Application.ScreenUpdating = False
'
'
''' ATUALIZAÇÕES
'Do While Not rstOrigem.EOF
'
'    '' REMESSA
'    For x = 1 To UBound(strTabelas)
'        strSQL = "SELECT * FROM " & strTabelas(x) & " WHERE controle = '" & rstOrigem.Fields("CONTROLE").Value & "' AND vendedor = '" & rstOrigem.Fields("VENDEDOR").Value & "'"
'        RemessaDados strSQL, strSQL
'    Next x
'
'    '' CRIAR TAREFA NO SERVIDOR
'    criarTarefaServidor rstOrigem.Fields("BANCO").Value, rstOrigem.Fields("VENDEDOR").Value, rstOrigem.Fields("CONTROLE").Value
'
''    '' EXCLUIR TAREFAS LOCAIS
''    With qdfadmOrcamentoAtualizacoes
''        .Parameters("NM_CODIGO") = rstOrigem.Fields("CODIGO").Value
''        .Execute
''    End With
'
'    rstOrigem.MoveNext
'
'Loop
'
'qdfadmOrcamentoAtualizacoes.Close
'rstOrigem.Close
'dbOrcamento.Close
'
'Application.ScreenUpdating = True
'
'End Sub
'
'Sub RemessaDados(strLOCAL As String, strSERVER As String)
'
'    '' SERVER
'    Dim connectionString  As String
'    connectionString = "Driver={" & mysql_driver & "};Server=" & location & ";Database=" & Database & ";PORT=" & port & ";UID=" & user & ";PWD=" & password
'
'    Dim connection As ADODB.connection
'    Set connection = New ADODB.connection
'    connection.Open connectionString
'
'    Dim rstSERVER As ADODB.Recordset
'    Set rstSERVER = New ADODB.Recordset
'
'    Dim fld As ADODB.Field
'    Dim NewFile As Boolean: NewFile = False
'
'    rstSERVER.Open strSERVER, connection, adOpenDynamic, adLockOptimistic, adCmdText
'
'    '' LOCAL
'    Dim dbOrcamento As DAO.Database
'    Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)
'    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
'
'    Dim rstLOCAL As DAO.Recordset
'    Set rstLOCAL = dbOrcamento.OpenRecordset(strLOCAL)
'
'    connection.BeginTrans
'
'    '' SE Ñ EXISTE NO SERVER CADASTRAR
'    If rstSERVER.EOF Then
'        NewFile = True
'    End If
'
'    Do While Not rstLOCAL.EOF
'
'        If NewFile Then
'            rstSERVER.AddNew
'        End If
'
'        For Each fld In rstSERVER.Fields
'            If fld.Name <> "Codigo" Then
'                rstSERVER(fld.Name).Value = rstLOCAL(fld.Name).Value
'            End If
'        Next
'
'        rstSERVER.Update
'        rstLOCAL.MoveNext
'
'    Loop
'
'
'    connection.CommitTrans
'
'    rstSERVER.Close
'    rstLOCAL.Close
'
'    connection.Close
'
'End Sub
'
'
'Sub criarTarefaServidor(strBanco As String, strUsuario As String, strControle As String)
'    '' SERVER
'    Dim connectionString  As String
'    connectionString = "Driver={" & mysql_driver & "};Server=" & location & ";Database=" & Database & ";PORT=" & port & ";UID=" & user & ";PWD=" & password
'
'    Dim connection As ADODB.connection
'    Set connection = New ADODB.connection
'    connection.Open connectionString
'
'    connection.Execute "Insert into OrcamentosAtualizacoes (BANCO,USUARIO,CONTROLE) values ('" & strBanco & "','" & strUsuario & "','" & strControle & "')"
'
'    connection.Close
'
'End Sub


'Sub VerificarLinkWeb_RetornoDados()
'    If IsInternetConnected() = True Then
'        ' connected
'        MsgBox "RetornoDados"
'    Else
'        ' no connected
'        MsgBox "Ñ.RetornoDados"
'    End If
'End Sub
'
'Sub tarefasServidor(strUsuario As String)
'
''' SERVER
'Dim connectionString  As String
'connectionString = "Driver={" & mysql_driver & "};Server=" & location & ";Database=" & Database & ";PORT=" & port & ";UID=" & user & ";PWD=" & password
'
'Dim connection As ADODB.connection
'Set connection = New ADODB.connection
'connection.Open connectionString
'
'Dim rstSERVER As ADODB.Recordset
'Set rstSERVER = New ADODB.Recordset
'
'
'
'Dim dbOrcamento As DAO.Database
'Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)
'Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
'
'Dim rstUsuarios As DAO.Recordset
'Dim strUsuarios As String: strOrigem = "SELECT usuarios FROM qryUsuariosUsuarios where usuario = '" & strUsuario & "' ORDER BY usuarios"
'Set rstUsuarios = dbOrcamento.OpenRecordset(strUsuarios)
'
'Dim strATUALIZACOES As String
'
'
'Do While Not rstUsuarios.EOF
'
'    strATUALIZACOES = "Select * from OrcamentosAtualizacoes where usuario = '" & strUsuario & "'"
'    rstSERVER.Open strSERVER, connection, adOpenDynamic, adLockOptimistic, adCmdText
'
'
'rstUsuarios.MoveNext
'Loop
'
'
'Do While Not rstSERVER.EOF
'
'    MsgBox rstSERVER("CONTROLE").Value
'    rstSERVER.MoveNext
'
'Loop
'
'rstSERVER.Close
'connection.Close
'
'
'End Sub
'
'
'Sub RetornoDados(strLOCAL As String, strSERVER As String)
'
'    '' SERVER
'    Dim connectionString  As String
'    connectionString = "Driver={" & mysql_driver & "};Server=" & location & ";Database=" & Database & ";PORT=" & port & ";UID=" & user & ";PWD=" & password
'
'    Dim connection As ADODB.connection
'    Set connection = New ADODB.connection
'    connection.Open connectionString
'
'    Dim rstSERVER As ADODB.Recordset
'    Set rstSERVER = New ADODB.Recordset
'
'    Dim fld As ADODB.Field
'
'
'    rstSERVER.Open strSERVER, connection, adOpenDynamic, adLockOptimistic, adCmdText
'
'    '' LOCAL
'    Dim dbOrcamento As DAO.Database
'    Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)
'    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
'
'    Dim rstLOCAL As DAO.Recordset
'    Set rstLOCAL = dbOrcamento.OpenRecordset(strLOCAL)
'
'    BeginTrans
'
'    Do While Not rstLOCAL.EOF
'
'        '' SE Ñ EXISTE NO SERVER CADASTRAR
'        If rstLOCAL.EOF Then
'            rstLOCAL.AddNew
'        Else
'            rstLOCAL.Edit
'        End If
'
'        For Each fld In rstSERVER.Fields
'            If fld.Name <> "Codigo" Then
'                rstLOCAL(fld.Name).Value = rstSERVER(fld.Name).Value
'            End If
'        Next
'
'        rstSERVER.Update
'        rstLOCAL.MoveNext
'
'    Loop
'
'    CommitTrans
'
'    rstSERVER.Close
'    rstLOCAL.Close
'
'    connection.Close
'
'End Sub

'Sub LocalServer()
'
'
'
'    '' SERVIDOR
'    Dim connection As ADODB.connection
'    Dim rstServer As ADODB.Recordset
'
'
'    Dim fld As ADODB.Field
'    Dim connectionString  As String
'    connectionString = "Driver={" & mysql_driver & "};Server=" & location & ";Database=" & Database & ";PORT=" & port & ";UID=" & user & ";PWD=" & password
'
'    Set connection = New ADODB.connection
'    connection.Open connectionString
'
'    Dim strServer As String: strSQL = "SELECT * FROM Orcamentos"
'    Set rstServer = New ADODB.Recordset
'
'    rstServer.Open strServer, connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'
'
'    '' LOCAL
'    Dim dbOrcamento As DAO.Database
'    Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)
'    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
'
'    '' TABELA - LOCAL
'    Dim rstLocal As DAO.Recordset
'    Dim strLocal As String: strLocal = "SELECT * FROM Orcamentos WHERE controle = '076-11' AND vendedor = 'Shirlei'"
'    Set rstLocal = dbOrcamento.OpenRecordset(strLocal)
'
'
'    connection.BeginTrans
'
'    For Each fld In rstServer.Fields
'
'        rstServer(fld.Name).Value = rstLocal(fld.Name).Value
'
'        rstServer.Update
'        rstServer.Close
'
'    Next
'
'    connection.CommitTrans
'
'    dbOrcamento.Close
'    connection.Close
'
'
'End Sub
'
'
'
'Sub testeServer()
'
'    '' SERVER
'    Dim connectionString  As String
'    connectionString = "Driver={" & mysql_driver & "};Server=" & location & ";Database=" & Database & ";PORT=" & port & ";UID=" & user & ";PWD=" & password
'
'    Dim connection As ADODB.connection
'    Set connection = New ADODB.connection
'    connection.Open connectionString
'
'    Dim rstServer As ADODB.Recordset
'    Set rstServer = New ADODB.Recordset
'
'    Dim intLoop As Integer
'    Dim strSQL As String: strSQL = "SELECT * FROM Orcamentos WHERE controle = '076-11' AND vendedor = 'Shirlei'"
'
'    Dim fld As ADODB.Field
'
'    rstServer.Open strSQL, connection, adOpenDynamic, adLockOptimistic, adCmdText
'
'    '' LOCAL
'    Dim dbOrcamento As DAO.Database
'    Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)
'    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
'
'    Dim rstLocal As DAO.Recordset
'    Dim strLocal As String: strLocal = "SELECT * FROM Orcamentos WHERE controle = '076-11' AND vendedor = 'Shirlei'"
'    Set rstLocal = dbOrcamento.OpenRecordset(strLocal)
'
'
'    connection.BeginTrans
'
'    For Each fld In rstServer.Fields
'
'        If fld.Name <> "Codigo" Then
'            rstServer(fld.Name).Value = rstLocal(fld.Name).Value
'        End If
'
'    Next
'
'    rstServer.Update
'
'    connection.CommitTrans
'
'    rstServer.Close
'    rstLocal.Close
'
'
'    connection.Close
'
'
'End Sub
'

'Option Explicit
'
'Sub Compact(NameFile As String, NameZipFile As String)
'    Dim PathZipProgram As String
'    Dim ShellStr As String
'
'    'Path of the Zip program
'    PathZipProgram = "C:\program files\7-Zip\"
'    If Right(PathZipProgram, 1) <> "\" Then
'        PathZipProgram = PathZipProgram & "\"
'    End If
'
'    'Check if this is the path where 7z is installed.
'    If Dir(PathZipProgram & "7z.exe") = "" Then
'        MsgBox "Please find your copy of 7z.exe and try again"
'        Exit Sub
'    End If
'
'    ShellStr = PathZipProgram & "7z.exe a" _
'             & " " & Chr(34) & NameZipFile & Chr(34) _
'             & " " & NameFile
'
'    ShellAndWait ShellStr, vbHide
'
'
'
'End Sub
'
'Sub DesCompact(FileNameZip As Variant, NameUnZipFolder As String)
'    Dim PathZipProgram As String
'    Dim ShellStr As String
'
'    'Path of the Zip program
'    PathZipProgram = "C:\program files\7-Zip\"
'    If Right(PathZipProgram, 1) <> "\" Then
'        PathZipProgram = PathZipProgram & "\"
'    End If
'
'    'Check if this is the path where 7z is installed.
'    If Dir(PathZipProgram & "7z.exe") = "" Then
'        MsgBox "Please find your copy of 7z.exe and try again"
'        Exit Sub
'    End If
'
'    'There are a few commands/Switches that you can change in the ShellStr
'    'We use x command now to keep the folder stucture, replace it with e if you want only the files
'    '-aoa Overwrite All existing files without prompt.
'    '-aos Skip extracting of existing files.
'    '-aou aUto rename extracting file (for example, name.txt will be renamed to name_1.txt).
'    '-aot auto rename existing file (for example, name.txt will be renamed to name_1.txt).
'    'Use -r if you also want to unzip the subfolders from the zip file
'    'You can add -ppassword if you want to unzip a zip file with password (only 7zip files)
'    'Change "*.*" to for example "*.txt" if you only want to unzip the txt files
'    'Use "*.xl*" for all Excel files: xls, xlsx, xlsm, xlsb
'    ShellStr = PathZipProgram & "7z.exe x -aoa -r" _
'             & " " & Chr(34) & FileNameZip & Chr(34) _
'             & " -o" & Chr(34) & NameUnZipFolder & Chr(34) & " " & "*.*"
'
'    ShellAndWait ShellStr, vbHide
'
'End Sub
'


'Function Retorno()
''Dim connection As New ADODB.connection
''
''Dim rstAtualizacoes As ADODB.Recordset
''Set rstAtualizacoes = New ADODB.Recordset
''
''Dim strTabelas(3) As String, strSQL As String
''
''strTabelas(1) = "Orcamentos"
''strTabelas(2) = "OrcamentosAnexos"
''strTabelas(3) = "OrcamentosCustos"
''
''    ''Is Internet Connected
''    If IsInternetConnected() = True Then
''        Set connection = OpenConnection("MySQL")
''
''        ' connected
''        If connection.State = 1 Then
''            ' tasks
'''            Call rstAtualizacoes.Open("SELECT * FROM OrcamentosAtualizacoes where USUARIO = '" & strUsuario & "' ", connection, , adLockOptimistic)
''
''            Call rstAtualizacoes.Open("SELECT * FROM OrcamentosAtualizacoes", connection, , adLockOptimistic)
''
''            Do While Not rstAtualizacoes.EOF
''                '' REMESSA
''                For x = 1 To UBound(strTabelas)
''                    '' ATUALIZAÇÃO DE ORÇAMENTOS
''                    strSQL = "SELECT * FROM " & strTabelas(x) & " WHERE controle = '" & rstAtualizacoes.Fields("CONTROLE").Value & "' AND vendedor = '" & rstAtualizacoes.Fields("VENDEDOR").Value & "'"
''                    EnvioDeDados "MySQL", "Access2003", strSQL
''
''                Next x
''
''                '' CADASTRO NO HISTORICO DE ATUALIZAÇÕES
'''                addTasksReturn "Access2003", rstAtualizacoes.Fields("VENDEDOR").Value, rstAtualizacoes.Fields("CONTROLE").Value
''
''                rstAtualizacoes.MoveNext
''
''            Loop
''
''            delTasks "MySQL", Range(NomeUsuario)
''
''            MsgBox "Retorno ok!", vbInformation + vbOKOnly, "Retorno"
''
''        Else
''            MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "(BANCO LOCAL) - Falha na conexão!"
''        End If
''
''        connection.Close
''    Else
''        ' no connected
''        MsgBox "SEM INTERNET.", vbOKOnly + vbExclamation
''    End If
'
'End Function
'
'Sub VerificarLinkWeb_RemessaDados()
'    If IsInternetConnected() Then
'        ' connected
''        tarefasLocais
'    Else
'        ' no connected
'        MsgBox "Ñ.RemessaDados"
'    End If
'End Sub
'
'Function Remessa()
'Dim connection As New ADODB.connection
'
'Dim rstAtualizacoes As ADODB.Recordset
'Set rstAtualizacoes = New ADODB.Recordset
'
'Dim strTabelas(3) As String, strSQL As String, strUsuario As String
'
'strTabelas(1) = "Orcamentos"
'strTabelas(2) = "OrcamentosAnexos"
'strTabelas(3) = "OrcamentosCustos"
'
'    ''Is Internet Connected
'    If IsInternetConnected() = True Then
'        Set connection = OpenConnection("Access2003")
'
'        ' connected
'        If connection.State = 1 Then
'            ' tasks
'            Call rstAtualizacoes.Open("SELECT * FROM OrcamentosAtualizacoes", connection, , adLockOptimistic)
'
'            Do While Not rstAtualizacoes.EOF
'                '' REMESSA
'                For x = 1 To UBound(strTabelas)
'                    '' ATUALIZAÇÃO DE ORÇAMENTOS
'                    strSQL = "SELECT * FROM " & strTabelas(x) & " WHERE controle = '" & rstAtualizacoes.Fields("CONTROLE").Value & "' AND vendedor = '" & rstAtualizacoes.Fields("VENDEDOR").Value & "'"
'                    EnvioDeDados "Access2003", "MySQL", strSQL
'
'                Next x
'
'                '' CADASTRO NO HISTORICO DE ATUALIZAÇÕES
'                addTasksShipping "MySQL", rstAtualizacoes.Fields("VENDEDOR").Value, rstAtualizacoes.Fields("CONTROLE").Value
'
'                strUsuario = rstAtualizacoes.Fields("VENDEDOR").Value
'
'                rstAtualizacoes.MoveNext
'
'            Loop
'
'            delTasks "Access2003", strUsuario
'
'            MsgBox "Remessa ok!", vbInformation + vbOKOnly, "Remessa"
'
'        Else
'            MsgBox "Falha na conexão com o banco de dados!", vbCritical + vbOKOnly, "(BANCO LOCAL) - Falha na conexão!"
'        End If
'
'        connection.Close
'    Else
'        ' no connected
'        MsgBox "SEM INTERNET.", vbOKOnly + vbExclamation
'    End If
'
'End Function
'
'
'Sub testeRemessa()
'
'EnvioDeDados "Access2003", "MySQL", "SELECT * FROM OrcamentosAtualizacoes"
'
'End Sub
'
'Sub testeRemessa2()
'
'addTasksShipping "MySQL", "FABIANA", "098-14"
'
'End Sub
'
'Sub testeRetorno()
'
'addTasksReturn "MySQL", "FABIANA", "098-14"
'
'End Sub



