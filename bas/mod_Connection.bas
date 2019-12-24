Attribute VB_Name = "mod_Connection"
Public Const strSource As String = "MySQL"

'Public Const mysql_driver As String = "MySQL ODBC 3.51 Driver"
Public Const mysql_driver As String = "MySQL ODBC 5.2a Driver"

Public Const location As String = "186.202.152.40"

Public Const port As String = "3306"

Public Const Database As String = "Ailton_springer"

Public Const user As String = "Ailto_springer"

Public Const password As String = "41L70N@@"

'Sub testeBancoLocal()
'
'    MsgBox Range(BancoLocal)
'
'End Sub



Public Function OpenConnection(source As String) As ADODB.connection
''   Read type and location of the database, user login and password
'    Dim source As String, location As String, user As String, password As String, mysql_driver As String, port As String
'
'    source = Range("Source").Value
'    mysql_driver = Range("driver").Value
'    location = Range("location").Value
'    port = Range("port").Value
'    Database = Range("database").Value
'    user = Range("user").Value
'    password = Range("password").Value

'OpenConnection.Close


''    Handle relative path for the location of Access and SQLite database files
    If (source = "Access2003" Or source = "SQLite") And Not location Like "?:\*" Then
'        strlocation = ActiveWorkbook.Path & "\db\" & ActiveWorkbook.Name & ""
        
        strlocation = Range(BancoLocal)
        
        
    End If

''    Build the connection string depending on the source
    Dim connectionString As String
    Select Case source
        Case "Access"
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strlocation
        Case "Access2003"
            connectionString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & strlocation & ";Uid=admin;PWD=" & SenhaBanco & ""
        Case "MySQL"
            connectionString = "Driver={" & mysql_driver & "};Server=" & location & ";Database=" & Database & ";PORT=" & port & ";UID=" & user & ";PWD=" & password
        Case "PostgreSQL"
            connectionString = "Driver={PostgreSQL ANSI};Server=" & location & ";Database=test;UID=" & user & ";PWD=" & password
        Case "SQLite"
            connectionString = "Driver={SQLite3 ODBC Driver};Database=" & location
    End Select

''    Create and open a new connection to the selected source
    Set OpenConnection = New ADODB.connection
    Call OpenConnection.Open(connectionString)
    
End Function
