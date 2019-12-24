Attribute VB_Name = "Módulo1"
Sub Obrigatorios()
Dim strObrigatorio() As Variant
Dim i As Integer

strObrigatorio = Array("J4", "C12", "D12", "E12", "F12", "G12", "H12", "I12", "J12")

    For i = 0 To UBound(strObrigatorio, 1)
        DesbloqueioDeGuia SenhaBloqueio
        MarcarObrigatorio strObrigatorio(i), False
        BloqueioDeGuia SenhaBloqueio
    Next i
    
    
    

End Sub

Public Function ListarCamposObrigatorios(BaseDeDados As String, strEtapa As String)
On Error GoTo ListarCamposObrigatorios_err
Dim dbOrcamento As dao.Database
Dim rstSelecao As dao.Recordset
Dim strSQL As String
Dim RetVal As Variant

RetVal = Dir(BaseDeDados)

If RetVal <> "" Then
   
    strSQL = "SELECT selecao FROM qryObrigatorios WHERE Etapa = '" & strEtapa & "' ORDER BY Ordem"
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstSelecao = dbOrcamento.OpenRecordset(strSQL)
    
    DesbloqueioDeGuia SenhaBloqueio
    While Not rstSelecao.EOF
        MarcarObrigatorio rstSelecao.Fields("Selecao").Value, False
        rstSelecao.MoveNext
    Wend
    BloqueioDeGuia SenhaBloqueio
    
    rstSelecao.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstSelecao = Nothing
    
End If

ListarCamposObrigatorios_Fim:
  
    Exit Function
ListarCamposObrigatorios_err:
    
    MsgBox Err.Description
    Resume ListarCamposObrigatorios_Fim
End Function


Function MarcarObrigatorio(ByVal strCelula As String, Marcar As Boolean)
'' Marcar celula obrigatoria quando estiver vasia
    If Range(strCelula) = "" Or Range(strCelula) = 0 Then
        Range(strCelula).Select
        If Marcar Then
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Else
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    End If
End Function
