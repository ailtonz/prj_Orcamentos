Attribute VB_Name = "frm_ComboBox"
Option Explicit

Function ComboBoxUpdate(wsGuia As String, lstListagem As String, cbo As ComboBox)
Dim cLoc As Range
Dim ws As Worksheet
Set ws = Worksheets(wsGuia)

cbo.Clear

For Each cLoc In ws.Range(lstListagem)
  With cbo
    .AddItem cLoc.Value
  End With
Next cLoc

End Function

Public Function ComboBoxCarregar(BaseDeDados As String, cbo As ComboBox, strCampo As String, strSQL As String) As Boolean: ComboBoxCarregar = True
On Error GoTo ComboBoxCarregar_err
Dim dbOrcamento As DAO.Database
Dim rstComboBoxCarregar As DAO.Recordset
Dim RetVal As Variant

RetVal = Dir(BaseDeDados)

If RetVal = "" Then

    ComboBoxCarregar = False
    
Else
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstComboBoxCarregar = dbOrcamento.OpenRecordset(strSQL)
    
    cbo.Clear
    
    While Not rstComboBoxCarregar.EOF
        cbo.AddItem rstComboBoxCarregar.Fields(strCampo)
        rstComboBoxCarregar.MoveNext
    Wend
        
    rstComboBoxCarregar.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstComboBoxCarregar = Nothing
    
End If

ComboBoxCarregar_Fim:
  
    Exit Function
ComboBoxCarregar_err:
    ComboBoxCarregar = False
    MsgBox Err.Description
    Resume ComboBoxCarregar_Fim
End Function
