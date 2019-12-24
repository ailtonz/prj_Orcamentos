VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjetosImpressao 
   Caption         =   "IMPRESSÕES"
   ClientHeight    =   8550.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7020
   OleObjectBlob   =   "frmProjetosImpressao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProjetosImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private registro As New clsProjetoImpressao

Private Sub UserForm_Initialize()
    listarRegistros
    listarTipos
    listarPapel
    listarNPaginas
    listarImpressao
    listarFormato
End Sub

Private Sub listarRegistros()
Dim Prf As clsProjetoImpressao
Dim col As clsProjetoImpressao
Dim orc As clsOrcamentos

carregarBanco

Set orc = New clsOrcamentos
Set Prf = New clsProjetoImpressao

Set col = Prf.getImpressaoProjeto(Bnc, objOrc)

With Me.lstRegistros
    .Clear
    .ColumnCount = 3
    .ColumnWidths = "0;200;60"
    
    For Each Prf In col.Itens
        .AddItem Prf.ID
        .List(.ListCount - 1, 1) = Prf.Tipo
        .List(.ListCount - 1, 2) = Prf.Papel
    Next Prf

End With

End Sub

Private Sub listarTipos()
Dim ws As Worksheet
Set ws = Worksheets("Apoio")

ComboBoxUpdate ws.Name, "TIPO", Me.cboTipo

Set ws = Nothing

End Sub

Private Sub listarPapel()
Dim ws As Worksheet
Set ws = Worksheets("Apoio")

ComboBoxUpdate ws.Name, "PAPEL", Me.cboPapel

Set ws = Nothing

End Sub

Private Sub listarNPaginas()
Dim ws As Worksheet
Set ws = Worksheets("Apoio")

ComboBoxUpdate ws.Name, "NPAGINAS", Me.cboNumPaginas

Set ws = Nothing

End Sub

Private Sub listarImpressao()
Dim ws As Worksheet
Set ws = Worksheets("Apoio")

ComboBoxUpdate ws.Name, "IMPRESSAO", Me.cboImpressao

Set ws = Nothing

End Sub

Private Sub listarFormato()
Dim ws As Worksheet
Set ws = Worksheets("Apoio")

ComboBoxUpdate ws.Name, "FORMATO", Me.cboFormato

Set ws = Nothing

End Sub

Private Sub limparCampos()

    Me.txtId.value = ""
               
    Me.cboTipo.value = ""
    Me.cboPapel.value = ""
    Me.cboNumPaginas.value = ""
    Me.cboImpressao.value = ""
    Me.cboFormato.value = ""
    
    Me.cboTipo.SetFocus
    
    Me.cmdSalvar.Caption = "NOVO"
    
End Sub

Private Function mostrarRegistro() As Variant
Dim retVal As Variant

carregarCampos
    
retVal = MsgBox("Você deseja realmente EXCLUIR o registro abaixo:" & vbNewLine & _
        vbNewLine & _
        "TIPO: " & vbTab & registro.Item(1).Tipo & vbNewLine & _
        "PAPEL: " & vbTab & registro.Item(1).Papel & vbNewLine & _
        "NUM.PAGINAS : " & vbTab & registro.Item(1).NumPaginas & vbNewLine & _
        "IMPRESSÃO : " & vbTab & registro.Item(1).impressao & vbNewLine & _
        "FORMATO : " & vbTab & registro.Item(1).Formato & vbNewLine, vbCritical + vbYesNo, "EXCLUSÃO DE REGISTRO!")
        
mostrarRegistro = retVal
            
Set retVal = Nothing

End Function

Private Sub carregarCampos()
Dim obj As clsProjetoImpressao
Dim col As clsProjetoImpressao
Dim i As Long: i = Me.lstRegistros.value

carregarBanco
limparCampos

Set obj = New clsProjetoImpressao

Set col = obj.getImpressaoIndex(Bnc, i)
Set registro = obj.getImpressaoIndex(Bnc, i)

For Each obj In col.Itens
   
    Me.txtId.value = obj.ID
            
    Me.cboTipo.SetFocus
    Me.cboTipo.SelText = obj.Tipo
    
    Me.cboPapel.SetFocus
    Me.cboPapel.SelText = obj.Papel
    
    Me.cboNumPaginas.SetFocus
    Me.cboNumPaginas.SelText = obj.NumPaginas
    
    Me.cboImpressao.SetFocus
    Me.cboImpressao.SelText = obj.impressao
        
    Me.cboFormato.SetFocus
    Me.cboFormato.SelText = obj.Formato
    
    Me.cboTipo.SetFocus

Next obj
                
End Sub

Private Sub salvarRegistro()

Dim obj As clsProjetoImpressao
Set obj = New clsProjetoImpressao

carregarBanco
            
    With obj
        .ID = Me.txtId.value
        
        .NumControle = objOrc.Controle
        .Vendedor = objOrc.Vendedor
        .NumProjeto = objOrc.NumProjeto
        
        .Tipo = Me.cboTipo.value
        .Papel = Me.cboPapel.value
        .NumPaginas = Me.cboNumPaginas.value
        .impressao = Me.cboImpressao.value
        .Formato = Me.cboFormato.value
        
        .add obj
    End With
                      
    If obj.ID = "" Then
        If (obj.Insert(Bnc, obj) = True) Then
            MsgBox "Cadastro realizado com sucesso!", vbInformation + vbOKOnly, "Cadastro"
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


Set orc = Nothing
Set obj = Nothing
Set Bnc = Nothing

End Sub

Private Sub lstRegistros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNull(Me.lstRegistros.value) Then
        carregarCampos
        Me.cmdSalvar.Caption = "SALVAR"
    End If
    
End Sub

Private Sub lstRegistros_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Not IsNull(Me.lstRegistros.value) Then
            carregarCampos
            Me.cmdSalvar.Caption = "EXCLUIR"
            Me.cmdSalvar.SetFocus
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    limparCampos
End Sub

Private Sub cmdSalvar_Click()
    salvarRegistro
End Sub


