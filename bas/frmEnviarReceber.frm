VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEnviarReceber 
   Caption         =   "Enviar / Receber"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13800
   OleObjectBlob   =   "frmEnviarReceber.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEnviarReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strPesquisar As String
Private strSql As String

Private Sub cboEtapas_Click()
    Call cboOperacao_Click
End Sub

Private Sub cboOperacao_Click()
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

    If Me.cboOperacao.value = "RECEBER" Then
        loadBancos
        strSql = "Select * from qryOrcamentosEnviar WHERE qryOrcamentosEnviar.Pesquisa Like '%" & strPesquisar & "%' and STATUS = '" & Me.cboEtapas.value & "' "
        strSql = strSql & " AND ((qryOrcamentosEnviar.VENDEDOR) In (SELECT  qryUsuariosUsuarios.Usuarios From qryUsuariosUsuarios WHERE (((qryUsuariosUsuarios.Usuario)='" & strUsuario & "')))) limit " & Me.txtLimiteRegistros.value & ""
        
        ListBoxCarregarADO banco(0), Me, Me.lstSelecao.Name, "Pesquisa", strSql
    ElseIf Me.cboOperacao.value = "ENVIAR" Then
        MontarPesquisa
        carregarLista
    End If
    
    Me.lblStatus.Caption = Me.lstSelecao.ListCount

End Sub

Private Sub cmdEnviar_Click()
Dim Matriz As Variant
Dim intCurrentRow As Integer

'' CARREGAR BANCOS
loadBancos

Me.lblStatus.Caption = Me.lstSelecao.ListCount

    ''' EXPORTA PARA O BANCO ITENS SELECIONADOS
    For intCurrentRow = 0 To Me.lstSelecao.ListCount - 1
        DoEvents

        If Me.lstSelecao.Selected(intCurrentRow) Then
            ''' CARREGA MATRIZ
            Matriz = Split(Me.lstSelecao.List(intCurrentRow), " - ")
           
            '' CARREGAR OR�AMENTO
            loadOrcamento CStr(Matriz(2)), CStr(Matriz(0)), Sheets(ActiveSheet.Name).Range(NomeUsuario)
            
            If Me.cboOperacao.value = "RECEBER" Then
                Transferencia Me.cboOperacao.value, Departamento(banco(1), Orcamento), Orcamento, banco(0), banco(1)
            ElseIf Me.cboOperacao.value = "ENVIAR" Then
                Transferencia Me.cboOperacao.value, Departamento(banco(1), Orcamento), Orcamento, banco(1), banco(0)
            End If
            
            Me.lblStatus.Caption = Me.lstSelecao.ListCount - 1 & " - " & intCurrentRow
            
            ''' DESMARCAR ITEM SELECIONADO
            Me.lstSelecao.Selected(intCurrentRow) = False
        End If

    Next intCurrentRow
    
    MsgBox "Conclu�do !", vbInformation + vbOKOnly, Me.cboOperacao.value
    

End Sub

Private Sub cmdPesquisar_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)

Dim retValor As Variant

    retValor = InputBox("Digite uma palavra para fazer o filtro:", "Filtro", strPesquisar, 0, 0)
    strPesquisar = retValor
    
    Call cboOperacao_Click

End Sub

Private Sub cmdTodos_Click()
Dim intCurrentRow As Integer
            
For intCurrentRow = 0 To Me.lstSelecao.ListCount - 1
    If Not IsNull(Me.lstSelecao.Column(0, intCurrentRow)) Then
        Me.lstSelecao.Selected(intCurrentRow) = True
    End If
Next intCurrentRow

End Sub

Private Sub cmdNenhum_Click()
Dim intCurrentRow As Integer
            
For intCurrentRow = 0 To Me.lstSelecao.ListCount - 1
    If Not IsNull(Me.lstSelecao.Column(0, intCurrentRow)) Then
        Me.lstSelecao.Selected(intCurrentRow) = False
    End If
Next intCurrentRow

End Sub

Private Sub UserForm_Initialize()

carregarOperacao
carregarEtapa

Call cboOperacao_Click

End Sub

Private Sub MontarPesquisa()
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

strSql = "Select top " & Me.txtLimiteRegistros.value & " * from qryOrcamentosEnviar WHERE ((qryOrcamentosEnviar.Pesquisa) Like '*" & strPesquisar & "*') and STATUS = '" & Me.cboEtapas.value & "'"
strSql = strSql & " AND ((qryOrcamentosEnviar.VENDEDOR) In (SELECT  qryUsuariosUsuarios.Usuarios From qryUsuariosUsuarios WHERE (((qryUsuariosUsuarios.Usuario)='" & strUsuario & "'))))"

End Sub

Private Sub carregarLista()
Dim strBanco As String: strBanco = Range(BancoLocal)

ListBoxCarregar strBanco, Me, Me.lstSelecao.Name, "Pesquisa", strSql
Me.Repaint

End Sub

Private Sub carregarOperacao()
Dim strBanco As String: strBanco = Range(BancoLocal)

ComboBoxCarregar strBanco, Me.cboOperacao, "Sincronismo", "Select Distinct Sincronismo from qrySincronismo order by Sincronismo"

Me.cboOperacao.value = "ENVIAR"

End Sub

Private Sub carregarEtapa()
Dim strBanco As String: strBanco = Range(BancoLocal)

ComboBoxCarregar strBanco, Me.cboEtapas, "Status", "SELECT DISTINCT qryEtapas.ATUAL, qryEtapas.Status From qryEtapas ORDER BY qryEtapas.ATUAL"

Me.cboEtapas.value = "Custo"

End Sub

