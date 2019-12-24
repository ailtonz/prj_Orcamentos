Attribute VB_Name = "xEnviarDados"
Sub EnvioDeDados(dbOrigem As String, dbDestino As String, strSQL As String)

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

   
    Call rstOrigem.Open(strSQL, Origem, , adLockOptimistic)
           
    Call rstDestino.Open(strSQL, Destino, adOpenDynamic, adLockOptimistic, adCmdText)
           
               
'    Destino.BeginTrans
    
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
                rstDestino(fld.Name).Value = rstOrigem(fld.Name).Value
            End If
        Next
        
        rstDestino.Update
        rstOrigem.MoveNext
        
    Loop
    
'    Destino.CommitTrans
 
    rstDestino.Close
    rstOrigem.Close
    
    Destino.Close
    Origem.Close

End Sub
