VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustos 
   Caption         =   "CUSTOS DE PRODUÇÃO"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   OleObjectBlob   =   "frmCustos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCustos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    listarCustos
    listarTipos
    listarEstilos
    listarSubTipos
    limparCampos
End Sub

Private Sub cmdCancelar_Click()
    limparCampos
End Sub

Private Sub cmdSalvar_Click()
    salvarRegistro
End Sub

Private Sub txtValor_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtValor.value = FormatCurrency(Me.txtValor.value)
End Sub

Private Sub lstRegistros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '' CARREGAR REGISTRO
    If Not IsNull(Me.lstRegistros.value) Then
        Me.txtId.value = Me.lstRegistros.value
        Me.cboTipos.value = Me.lstRegistros.Column(2)
        Me.cboEstilos.value = Me.lstRegistros.Column(5)
        Me.cboSubTipos.value = Me.lstRegistros.Column(4)
        Me.txtPaginas.value = Me.lstRegistros.Column(1)
        Me.txtValor.value = FormatCurrency(Me.lstRegistros.Column(3))
        
        Me.cmdSalvar.Caption = "SALVAR"
    End If
    
End Sub

Private Sub lstRegistros_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '' EXCLUIR REGISTRO
    If KeyCode = vbKeyDelete Then
        If Not IsNull(Me.lstRegistros.value) Then
            Me.txtId.value = Me.lstRegistros.value
            Me.cboTipos.value = Me.lstRegistros.Column(2)
            Me.cboEstilos.value = Me.lstRegistros.Column(5)
            Me.cboSubTipos.value = Me.lstRegistros.Column(4)
            Me.txtPaginas.value = Me.lstRegistros.Column(1)
            Me.txtValor.value = FormatCurrency(Me.lstRegistros.Column(3))
            
            Me.cmdSalvar.Caption = "EXCLUIR"
        End If
    End If
End Sub

Private Sub salvarRegistro()

Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

Dim obj As clsCustoProducao

Set obj = New clsCustoProducao

carregarBanco

    With obj
        .ID = Me.txtId.value
        .Tipo = Me.cboTipos.value
        .Estilo = Me.cboEstilos.value
        .SubTipo = Me.cboSubTipos.value
        .Paginas = Me.txtPaginas.value
        .Valor = Me.txtValor.value
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

    listarCustos
    listarTipos
    listarEstilos
    listarSubTipos
    
    limparCampos
    

Set obj = Nothing
Set Bnc = Nothing

End Sub


Private Sub listarCustos()
Dim Prf As clsCustoProducao
Dim col As clsCustoProducao

carregarBanco

Set Prf = New clsCustoProducao

Set col = Prf.getCustosProducao(Bnc)

With Me.lstRegistros
    .Clear
    .ColumnCount = 6
    .ColumnWidths = "0;40;160;10;0;0"
    
    For Each Prf In col.Itens
        .AddItem Prf.ID
        .List(.ListCount - 1, 1) = Prf.Paginas
        .List(.ListCount - 1, 2) = Prf.Tipo
        .List(.ListCount - 1, 3) = FormatCurrency(Prf.Valor)
        .List(.ListCount - 1, 4) = Prf.SubTipo
        .List(.ListCount - 1, 5) = Prf.Estilo
    Next Prf

End With


End Sub

Private Sub listarEstilos()

carregarBanco

Dim Prf As clsEstilos
Set Prf = New clsEstilos

Dim col As clsEstilos
Set col = Prf.getEstilos(Bnc)

With Me.cboEstilos
    .Clear
    .Clear
    .ColumnCount = 1
    .ColumnWidths = "60"

    For Each Prf In col.Itens
        .AddItem Prf.Estilo
    Next Prf

End With

End Sub

Private Sub listarTipos()

carregarBanco

Dim Prf As clsTipos
Set Prf = New clsTipos

Dim col As clsTipos
Set col = Prf.getTipos(Bnc)

With Me.cboTipos
    .Clear
    .Clear
    .ColumnCount = 1
    .ColumnWidths = "60"

    For Each Prf In col.Itens
        .AddItem Prf.Tipo
    Next Prf

End With

End Sub

Private Sub listarSubTipos()

carregarBanco

Dim Prf As clsSubTipos
Set Prf = New clsSubTipos

Dim col As clsSubTipos
Set col = Prf.getSubTipos(Bnc)

With Me.cboSubTipos
    .Clear
    .Clear
    .ColumnCount = 1
    .ColumnWidths = "60"

    For Each Prf In col.Itens
        .AddItem Prf.SubTipo
    Next Prf

End With

End Sub

Private Sub limparCampos()

    Me.txtId.value = ""
    Me.cboTipos.value = ""
    Me.cboEstilos.value = ""
    Me.cboSubTipos.value = ""
    Me.txtPaginas.value = ""
    Me.txtValor.value = ""
    
    Me.cmdSalvar.Caption = "NOVO"
    
End Sub

Private Function mostrarRegistro() As Variant
Dim retVal As Variant

    retVal = MsgBox("Você deseja realmente EXCLUIR o registro abaixo:" & vbNewLine & _
            vbNewLine & _
            "TIPOS: " & Me.lstRegistros.Column(2) & vbNewLine & _
            "ESTILOS: " & Me.lstRegistros.Column(5) & vbNewLine & _
            "SUBTIPOS : " & (Me.lstRegistros.Column(4)) & vbNewLine & _
            "PAGINAS  : " & (Me.lstRegistros.Column(1)) & vbNewLine & _
            "VALOR  : " & FormatCurrency(Me.lstRegistros.Column(3)) & vbNewLine & _
            vbNewLine, vbCritical + vbYesNo, "EXCLUSÃO DE REGISTRO!")
            
    mostrarRegistro = retVal
            
Set retVal = Nothing

End Function

