Attribute VB_Name = "frm_UserForm"
Option Explicit

Public Function UserFormDesbloqueioDeFuncoes(BaseDeDados As String, frm As UserForm, strSQL As String, strCampo As String)
On Error GoTo UserFormDesbloqueioDeFuncoes_err

Dim dbOrcamento         As DAO.Database
Dim rstUserFormDesbloqueioDeFuncoes   As DAO.Recordset
Dim RetVal              As Variant
Dim Ctrl                As Control

RetVal = Dir(BaseDeDados)

If RetVal = "" Then

    UserFormDesbloqueioDeFuncoes = False
    
Else
        
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstUserFormDesbloqueioDeFuncoes = dbOrcamento.OpenRecordset(strSQL)
        
    While Not rstUserFormDesbloqueioDeFuncoes.EOF
        For Each Ctrl In frm.Controls
            If TypeName(Ctrl) = "CommandButton" Then
                If Right(Ctrl.Name, Len(Ctrl.Name) - 3) = rstUserFormDesbloqueioDeFuncoes.Fields(strCampo) Then
                    Ctrl.Enabled = True
                End If
            ElseIf TypeName(Ctrl) = "ListBox" Then
                If Right(Ctrl.Name, Len(Ctrl.Name) - 3) = rstUserFormDesbloqueioDeFuncoes.Fields(strCampo) Then
                    Ctrl.Enabled = True
                End If
            End If
        Next
        rstUserFormDesbloqueioDeFuncoes.MoveNext
    Wend
    
    rstUserFormDesbloqueioDeFuncoes.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstUserFormDesbloqueioDeFuncoes = Nothing
    
End If

UserFormDesbloqueioDeFuncoes_Fim:
  
    Exit Function
UserFormDesbloqueioDeFuncoes_err:
    UserFormDesbloqueioDeFuncoes = False
    MsgBox Err.Description
    Resume UserFormDesbloqueioDeFuncoes_Fim
End Function


Sub AtualizarProcesso(Percentual As Single, frm As UserForm) 'variável reservada para ser %

    With frm 'With usa o frmprocesso para as ações abaixo
    'sem ter que repetir o nome do objeto frmprocesso

        ' Atualiza o Título do Quadro que comporta a barra para %
'        .FrameProcesso.Caption = Format(Percentual, "0%")

        ' Atualza o tamanho da Barra (label)
        .lblProcesso.Width = Percentual * (100 - 10)
    End With 'final do uso de frmprocesso diretamente
    
    'Habilita o userform para ser atualizado
    DoEvents
End Sub

