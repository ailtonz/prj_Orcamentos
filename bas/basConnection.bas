Attribute VB_Name = "basConnection"
Public Bnc As New clsBancos
Public tmpProposta As String

Public Function OpenConnection(strBanco As infBanco) As ADODB.Connection
'' Build the connection string depending on the source
Dim connectionString As String
    
Select Case strBanco.strSource
    Case "Access"
        connectionString = "Provider=" & strBanco.strDriver & ";Data Source=" & strBanco.strDatabase
    Case "Access2003"
        connectionString = "Driver={" & strBanco.strDriver & "};Dbq=" & strBanco.strLocation & strBanco.strDatabase & ";Uid=" & strBanco.strUser & ";PWD=" & strBanco.strPassword & ""
    Case "SQLite"
        connectionString = "Driver={" & strBanco.strDriver & "};Database=" & strBanco.strDatabase
    Case "MySQL"
        connectionString = "Driver={" & strBanco.strDriver & "};Server=" & strBanco.strLocation & ";Database=" & strBanco.strDatabase & ";PORT=" & strBanco.strPort & ";UID=" & strBanco.strUser & ";PWD=" & strBanco.strPassword
    Case "PostgreSQL"
        connectionString = "Driver={" & strBanco.strDriver & "};Server=" & strBanco.strLocation & ";Database=" & strBanco.strDatabase & ";UID=" & strBanco.strUser & ";PWD=" & strBanco.strPassword
End Select

'' Create and open a new connection to the selected source
Set OpenConnection = New ADODB.Connection
Call OpenConnection.Open(connectionString)
   
End Function


Public Function OpenConnectionNEW(banco As clsBancos) As ADODB.Connection
'' Build the connection string depending on the source
Dim connectionString As String
    
Select Case banco.Source
    Case "Access"
        connectionString = "Provider=" & banco.Driver & ";Data Source=" & banco.Database
    Case "Access2003"
        connectionString = "Driver={" & banco.Driver & "};Dbq=" & banco.Location & banco.Database & ";Uid=" & banco.User & ";PWD=" & banco.Password & ""
    Case "SQLite"
        connectionString = "Driver={" & banco.Driver & "};Database=" & banco.Database
    Case "MySQL"
        connectionString = "Driver={" & banco.Driver & "};Server=" & banco.Location & ";Database=" & banco.Database & ";PORT=" & banco.Port & ";UID=" & banco.User & ";PWD=" & banco.Password
    Case "PostgreSQL"
        connectionString = "Driver={" & banco.Driver & "};Server=" & banco.Location & ";Database=" & banco.Database & ";UID=" & banco.User & ";PWD=" & banco.Password
End Select

'' Create and open a new connection to the selected source
Set OpenConnectionNEW = New ADODB.Connection
Call OpenConnectionNEW.Open(connectionString)
   
End Function

Public Sub carregarBanco()
Dim wsBnc As Worksheet
Set wsBnc = Worksheets("BANCOS")

    With Bnc
        .Source = wsBnc.Range("F2").value
        .Driver = wsBnc.Range("F3").value
        .Location = wsBnc.Range("F4").value
        .Database = wsBnc.Range("F5").value
        .User = wsBnc.Range("F6").value
        .Password = wsBnc.Range("F7").value
        .Port = wsBnc.Range("F8").value
        .add Bnc
    End With

Set wsBnc = Nothing

End Sub
