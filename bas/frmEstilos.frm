VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstilos 
   Caption         =   "ESTILOS"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   OleObjectBlob   =   "frmEstilos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEstilos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    listarRegistros
    limparCampos
End Sub

Private Sub txtEstilo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtEstilo.Text = UCase(Me.txtEstilo.Text)
End Sub

Private Sub cmdSalvar_Click()
    salvarRegistro
End Sub

Private Sub cmdCancelar_Click()
    limparCampos
End Sub

Private Sub lstRegistros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '' CARREGAR REGISTRO
    If Not IsNull(Me.lstRegistros.value) Then
        Me.txtId.value = Me.lstRegistros.value
        Me.txtEstilo.value = Me.lstRegistros.Column(1)
        
        Me.cmdSalvar.Caption = "SALVAR"
    End If
    
End Sub

Private Sub lstRegistros_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '' EXCLUIR REGISTRO
    If KeyCode = vbKeyDelete Then
        If Not IsNull(Me.lstRegistros.value) Then
            Me.txtId.value = Me.lstRegistros.value
            Me.txtEstilo.value = Me.lstRegistros.Column(1)
            
            Me.cmdSalvar.Caption = "EXCLUIR"
        End If
    End If
End Sub

Private Sub salvarRegistro()

Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

Dim obj As clsEstilos

Set obj = New clsEstilos

carregarBanco

    With obj
        .ID = Me.txtId.value
        .Estilo = Me.txtEstilo.value
        .add obj
    End With

    If obj.ID = "" Then
        If (obj.Insert(Bnc, obj) = True) Then
            MsgBox "Cadastro realisado com sucesso!", vbInformation + vbOKOnly, "Cadastro"
        Else
            MsgBox "Não foi possivel realizar o cadastro!", vbCritical + vbOKOnly, "Cadastro - ERRO!"
        End If

    Else

        If Me.cmdSalvar.Caption = "SALVAR" Then
            If (obj.Update(Bnc, obj) = True) Then
                MsgBox "Alteração realizada com sucesso!", vbInformation + vbOKOnly, "Alteração"
            Else
                MsgBox "Não foi possivel realizar alteração!", vbCritical + vbOKOnly, "Alteração - ERRO!"
            End If
        Else
            If mostrarRegistro = vbYes Then
                If (obj.Delete(Bnc, obj) = True) Then
                    MsgBox "Exclusão realizada com sucesso!", vbInformation + vbOKOnly, "Exclusão"
                Else
                    MsgBox "Não foi possivel realizar Exclusão!", vbCritical + vbOKOnly, "Exclusão - ERRO!"
                End If
            End If

        End If

    End If

    listarRegistros
    limparCampos
    

Set obj = Nothing
Set Bnc = Nothing

End Sub

Private Function mostrarRegistro() As Variant
Dim retVal As Variant

    retVal = MsgBox("Você deseja realmente EXCLUIR o registro abaixo:" & vbNewLine & _
            vbNewLine & _
            "ESTILO: " & Me.lstRegistros.Column(1) & vbNewLine & _
            vbNewLine, vbCritical + vbYesNo, "EXCLUSÃO DE REGISTRO!")
            
    mostrarRegistro = retVal
            
Set retVal = Nothing

End Function

Private Sub listarRegistros()
Dim Prf As clsEstilos
Dim col As clsEstilos

carregarBanco

Set Prf = New clsEstilos

Set col = Prf.getEstilos(Bnc)

With Me.lstRegistros
    .Clear
    .ColumnCount = 2
    .ColumnWidths = "0;60"
    
    For Each Prf In col.Itens
        .AddItem Prf.ID
        .List(.ListCount - 1, 1) = Prf.Estilo
    Next Prf

End With


End Sub

Private Sub limparCampos()

    Me.txtId.value = ""
    Me.txtEstilo.value = ""
    
    Me.cmdSalvar.Caption = "NOVO"
    
End Sub
