VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdministracao 
   Caption         =   "Administra��o Central do Sistema"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8340.001
   OleObjectBlob   =   "frmAdministracao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdministracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit
Dim sqlPermissoes As String
Dim sqlSelecao As String

Private Sub mensagemAtualizacao()
    MsgBox "Atualiza��o concluida", vbOKOnly + vbInformation, "Atualiza��o do sistema."
End Sub

Private Sub btnCustos_ATUALIZACAO_Click()
    mensagemAtualizacao
End Sub

Private Sub btnCustos_Click()
    frmCustos.Show
End Sub

Private Sub btnEstilos_ATUALIZACAO_Click()
    mensagemAtualizacao
End Sub

Private Sub btnEstilos_Click()
    frmEstilos.Show
End Sub

Private Sub btnIR_ATUALIZACAO_Click()
    mensagemAtualizacao
End Sub

Private Sub btnIR_Click()
    frmIR.Show
End Sub

Private Sub btnLinhas_ATUALIZACAO_Click()
    mensagemAtualizacao
End Sub

Private Sub btnLinhas_Click()
    frmLinhaProduto.Show
End Sub

Private Sub btnProfissoes_ATUALIZACAO_Click()
    mensagemAtualizacao
End Sub

Private Sub btnProfissoes_Click()
    frmProfissoes.Show
End Sub

Private Sub cboApoio_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strSql As String
Dim strParemetro As String: strParemetro = Me.cboApoio.Text

If Len(Me.cboApoio.Text) > 0 Then

    strSql = "SELECT admCategorias.Categoria AS DESCRICAO From admCategorias WHERE " & _
             " (((admCategorias.codRelacao)= " & _
             " (SELECT admCategorias.codCategoria FROM admCategorias Where Categoria = '" & strParemetro & "' and codRelacao = 0))) ORDER BY admCategorias.Categoria"
        
    Me.lstApoio.Clear
    ListBoxCarregar strBanco, Me, Me.lstApoio.Name, "DESCRICAO", strSql
        
End If

End Sub

Private Sub cboIndice_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strSql As String
Dim strParemetro As String: strParemetro = Me.cboIndice.Text

If Len(Me.cboIndice.Text) > 0 Then

'    strSQL = "SELECT admCategorias.Categoria AS DESCRICAO From admCategorias WHERE " & _
'             " (((admCategorias.codRelacao)= " & _
'             " (SELECT admCategorias.codCategoria FROM admCategorias Where Categoria = '" & strParemetro & "' and codRelacao = 0))) ORDER BY admCategorias.Categoria"


    strSql = "SELECT IIf(([DESCRICAO02])<>'',[CATEGORIA] & ' | ' & [DESCRICAO01] & ' | ' & [DESCRICAO02],[CATEGORIA] & ' | ' & [DESCRICAO01]) AS DESCRICAO " & _
                " From admCategorias " & _
                " WHERE (((admCategorias.codRelacao)=(SELECT admCategorias.codCategoria FROM admCategorias Where Categoria = '" & strParemetro & "' and codRelacao = 0))) " & _
                "ORDER BY IIf(([DESCRICAO02])<>'',[CATEGORIA] & ' | ' & [DESCRICAO01] & ' | ' & [DESCRICAO02],[CATEGORIA] & ' | ' & [DESCRICAO01])"


    Me.lstIndices.Clear
    ListBoxCarregar strBanco, Me, Me.lstIndices.Name, "DESCRICAO", strSql
End If

End Sub

Private Sub cboPermissoes_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
    
    sqlSelecao = "SELECT Selecionado FROM qryPermissoesUsuarios WHERE USUARIO = '" & Me.cboUsuario.Text & "' AND Categoria = '" & Me.cboPermissoes.Text & "'"
        
    sqlPermissoes = "Select * from qryPermissoesItens where Grupo = '" & Me.cboPermissoes.Text & "' and Item not in (" & sqlSelecao & ")"
       
    ListBoxCarregar strBanco, Me, Me.lstItensEmUso.Name, "Selecionado", sqlSelecao
    
    ListBoxCarregar strBanco, Me, Me.lstItensDisponiveis.Name, "ITEM", sqlPermissoes
        
End Sub

Private Sub cboPermissoes_Enter()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strSql As String

    strSql = "qryPermissoesGrupos"
    
    ComboBoxCarregar strBanco, Me.cboPermissoes, "Grupo", strSql

End Sub

Private Sub cboUsuario_Enter()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strSql As String

    strSql = "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No)) Order By Usuario"

    ComboBoxCarregar strBanco, Me.cboUsuario, "Usuario", strSql
    
    Me.cboPermissoes.Clear

    Me.lstItensDisponiveis.Clear
    
    Me.lstItensEmUso.Clear

End Sub

Private Sub cmdArquivoDeAtualizacao_Click()

''' CRIAR ARQUIVO DE ATUALIZA��O DO SISTEMA
Me.txtArquivoDeAtualizacao.Text = getFileNameAndExt(CriarArquivoDeAtualizacaoDoSistema)

End Sub

Private Sub cmdAtualizarApoio_Click()
    admAtualizarGuiaDeApoio
End Sub

Private Sub cmdAtualizarOperacional_Click()
'    AtualizarOperacional
End Sub

Private Sub cmdCopiar_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim Matriz As Variant
Dim strMsg As String
Dim strTitulo As String
Dim strSelecao As String


    If Me.lstUsuarios.value = "" Or IsNull(Me.lstUsuarios.value) Then
        strMsg = "ATEN��O: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "COPIAR!"
        
        MsgBox strMsg, vbInformation + vbOKOnly, strTitulo
    Else
      
        Matriz = Array()
        Matriz = Split(Me.lstUsuarios.value, " - ")
        
        admUsuarioCopiar Range(BancoLocal), CStr(Matriz(0)), CStr(Matriz(1))
        
        ListBoxCarregar strBanco, Me, Me.lstUsuarios.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No))"
        ListBoxCarregar strBanco, Me, Me.lstUsuariosExcluidos.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=yes))"
                
        limparCampos
        
    End If
End Sub

Private Sub cmdEnviarAtualizacoes_Click()

Dim strEmail As String
Dim strAssunto As String: strAssunto = Controle & "_" & UCase("ATUALIZACAO")
Dim strArquivo As String: strArquivo = pathWorkSheetAddress & Me.txtArquivoDeAtualizacao.Text
Dim strConteudo As String: strConteudo = ""
Dim intCurrentRow As Integer
Dim Matriz As Variant

If getFileStatus(strArquivo) Then

    Matriz = Array()
    ''' SELE��O DE USU�RIO PARA ENVIO DE ATUALIZA��ES
    For intCurrentRow = 0 To Me.lstAtulizacaoDeUsuarios.ListCount - 1
        DoEvents
           
            ''' ENVIO DE E-MAIL PARA O USU�RIO SELECIONADO
        If Me.lstAtulizacaoDeUsuarios.Selected(intCurrentRow) Then
            Matriz = Split(Me.lstAtulizacaoDeUsuarios.List(intCurrentRow), " - ")
            strEmail = CStr(Matriz(2))
            EnviarOrcamentos strEmail, strAssunto, strArquivo, strConteudo
            Me.lstAtulizacaoDeUsuarios.Selected(intCurrentRow) = False
        End If
    
    Next intCurrentRow
    
    ''' DELETA BASE DE DADOS TEMPORARIO COMPACTADO
    If Dir$(strArquivo) <> "" Then Kill strArquivo

End If

End Sub

Private Sub cmdExcluirApoio_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strMensagem As String: strMensagem = "ATEN��O: Voc� deseja realmente EXCLUIR este item ???"
Dim strTitulo As String: strTitulo = "EXCLUS�O DE ITEM !!!"
Dim retResposta As String

    retResposta = MsgBox(strMensagem, vbQuestion + vbYesNo, strTitulo)
    
    If (retResposta) = vbYes Then
        admGerenciarApoioExcluir strBanco, Me.cboApoio.Text, Me.lstApoio.value
        Me.txtApoio.Text = ""
    End If
    
    Call cboApoio_Click

End Sub

Private Sub cmdExcluirIndice_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strMensagem As String: strMensagem = "ATEN��O: Voc� deseja realmente EXCLUIR este item ???"
Dim strTitulo As String: strTitulo = "EXCLUS�O DE ITEM !!!"
Dim retResposta As String

    retResposta = MsgBox(strMensagem, vbQuestion + vbYesNo, strTitulo)

    If (retResposta) = vbYes Then
        admGerenciarIndiceExcluir strBanco, Me.cboIndice.Text, DivisorDeTexto(Me.lstIndices.value, "|", 0)
        Me.txtIndice.Text = ""
        Me.txtIndiceValor01.Text = ""
        Me.txtIndiceValor02.Text = ""
    End If

    Call cboIndice_Click
End Sub

Private Sub cmdNovoCaminhoDoBancoServer_Click()
Dim fd As Office.FileDialog
Dim strArq As String
Dim strCaminhoDoBancoServer As String: strCaminhoDoBancoServer = Me.txtCaminhoDoBancoServer


'If Not TestaExistenciaArquivo(strCaminhoDoBancoServer) Then

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.add "BDs do Access", "*.MDB"
    fd.Title = "Selecionar Banco Servidor"
    fd.AllowMultiSelect = False
    
    If fd.Show = -1 Then
        strArq = fd.SelectedItems(1)
    End If
    
    If strArq <> "" Then
'        DesbloqueioDeGuia SenhaBloqueio
        ''' ARMAZENAR BANCO SELECIONADO EM CONFIGIRA��ES DO SISTEMA (BANCO SERVIDOR)
        Sheets(cfgGuiaConfiguracao).Range(cfgBancoServidor) = strArq
        Me.txtCaminhoDoBancoServer.Text = strArq
'        BloqueioDeGuia SenhaBloqueio
    End If
    
'End If


End Sub

Private Sub cmdSalvarApoio_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strApoio As String: strApoio = Me.cboApoio.Text
Dim strAntigo As String: strAntigo = IIf(IsNull(Me.lstApoio.value), "", Me.lstApoio.value)
Dim strNovo As String: strNovo = Me.txtApoio

    If Len(Me.lstApoio.value) > 0 Then
        admGerenciarApoioAterar strBanco, strApoio, strAntigo, strNovo
    Else
        ''' N�O INCLUI APOIO SEM DESCRI��O
        If Len(Me.txtApoio.value) > 0 Then
            admGerenciarApoioIncluir strBanco, strApoio, strNovo
        End If
    End If
    
    Call cboApoio_Click
    Me.txtApoio.Text = ""
    
End Sub

Private Sub cmdSalvarIndice_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strIndice As String: strIndice = Me.cboIndice.Text
Dim strAntigo As String
Dim strNovo As String: strNovo = Me.txtIndice
Dim strValor_01 As String: strValor_01 = Me.txtIndiceValor01
Dim strValor_02 As String: strValor_02 = Me.txtIndiceValor02


    If IsNull(Me.lstIndices.value) Then
        strAntigo = ""
    Else
        strAntigo = DivisorDeTexto(Me.lstIndices.value, "|", 0)
    End If


    If Len(strAntigo) > 0 Then
        admGerenciarIndiceAterar strBanco, strIndice, strAntigo, strNovo, strValor_01, strValor_02
    Else
        ''' N�O INCLUI INDICES SEM DESCRI��O
        If Len(Me.txtIndice.value) > 0 Then
            admGerenciarIndiceIncluir strBanco, strIndice, strNovo, strValor_01, strValor_02
        End If
    End If

    Call cboIndice_Click
    Me.txtIndice.Text = ""
    Me.txtIndiceValor01.Text = ""
    Me.txtIndiceValor02.Text = ""
    
End Sub

Private Sub lstApoio_Click()
    Me.txtApoio = Me.lstApoio.value
End Sub

Private Sub lstIndices_Click()

    Me.txtIndice = Trim(DivisorDeTexto(Me.lstIndices.value, "|", 0))
    Me.txtIndiceValor01 = Trim(DivisorDeTexto(Me.lstIndices.value, "|", 1))
    Me.txtIndiceValor02 = Trim(DivisorDeTexto(Me.lstIndices.value, "|", 2))
    
End Sub

Private Sub lstItensDisponiveis_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strMsg As String
Dim strTitulo As String

    If IsNull(Me.lstItensDisponiveis.value) Then
        strMsg = "ATEN��O: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Sele��o de Item disponivel"
        
        MsgBox strMsg, vbInformation + vbOKOnly, strTitulo
    Else
    
        admUsuariosPermissoes strBanco, Me.cboUsuario, Me.lstItensDisponiveis, Me.cboPermissoes
        
        cboPermissoes_Click
    End If
    
End Sub

Private Sub lstItensEmUso_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strMsg As String
Dim strTitulo As String

    If IsNull(Me.lstItensEmUso.value) Then
        strMsg = "ATEN��O: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Remo��o de Item em uso"
        
        MsgBox strMsg, vbInformation + vbOKOnly, strTitulo
    Else

        admUsuariosPermissoesExcluir strBanco, Me.cboUsuario, Me.lstItensEmUso, Me.cboPermissoes
        
        cboPermissoes_Click
        
    End If
    
End Sub

''#########################################
''  FORMULARIO
''#########################################

Private Sub UserForm_Initialize()
Dim strBanco As String: strBanco = Range(BancoLocal)

    ''' ADICIONAR O "ADM" EM DEPARTAMENTOS
    Me.cboDepartamento.AddItem "ADM"
    
    ''' CARREGAR COMBO BOX DE DEPARTAMENTOS
    ComboBoxCarregar strBanco, Me.cboDepartamento, "Departamento", "qryDepartamentos"

    ''' CARREGAR LIST BOX DE USU�RIOS
    ListBoxCarregar strBanco, Me, Me.lstUsuarios.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No))"
    
    ''' CARREGAR LIST BOX DE USU�RIOS (EXCLU�DOS)
    ListBoxCarregar strBanco, Me, Me.lstUsuariosExcluidos.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=yes))"
    
    ''' CARREGAR COMBO BOX DE APOIO
    ComboBoxCarregar strBanco, Me.cboApoio, "Lista", "Select * from qryListas where TipoDeLista is null Order by Lista"

    ''' CARREGAR COMBO BOX DE INDICES
    ComboBoxCarregar strBanco, Me.cboIndice, "Lista", "Select * from qryListas where TipoDeLista = 'Indices' Order by Lista"

    ''' CARREGAR LIST BOX DE USU�RIOS PARA ATUALIZA��O
    ListBoxCarregar strBanco, Me, Me.lstAtulizacaoDeUsuarios.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No))"
    
    ''' CARREGAR CAMINHO DO BANCO SERVIDOR
'    Me.txtCaminhoDoBancoServer = Sheets("cfg").Range("B2")
    

End Sub

''#########################################
''  COMANDOS
''#########################################

Private Sub cmdSalvar_Enter()
    Me.txtEmail = LCase(Me.txtEmail)
End Sub

Private Sub cmdSalvar_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)

    If ExistenciaUsuario(Range(BancoLocal), Me.txtCodigo, Me.txtNome) Then
        admUsuarioSalvar Range(BancoLocal), Me.cboDepartamento, Me.txtCodigo, Me.txtNome, Me.txtEmail, Me.txtGerenteContas, Me.txtTelefone, Me.txtCelular01, Me.txtCelular02, Me.txtIdNextel
    Else
        admUsuarioNovo Range(BancoLocal), Me.cboDepartamento, Me.txtCodigo, Me.txtNome, Me.txtEmail, Me.txtGerenteContas, Me.txtTelefone, Me.txtCelular01, Me.txtCelular02, Me.txtIdNextel
    End If
    
    limparCampos
    
    ListBoxCarregar strBanco, Me, Me.lstUsuarios.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No))"
    
End Sub

Private Sub cmdCancelar_Click()
    limparCampos
End Sub

Private Sub cmdExcluir_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim Matriz As Variant
Dim strMsg As String
Dim strTitulo As String
Dim strSelecao As String


    If Me.lstUsuarios.value = "" Or IsNull(Me.lstUsuarios.value) Then
        strMsg = "ATEN��O: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "EXCLUIR!"
        
        MsgBox strMsg, vbInformation + vbOKOnly, strTitulo
    Else
      
        Matriz = Array()
        Matriz = Split(Me.lstUsuarios.value, " - ")
        
        admUsuarioExcluir Range(BancoLocal), CStr(Matriz(1)), True
        
        ListBoxCarregar strBanco, Me, Me.lstUsuarios.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No))"
        ListBoxCarregar strBanco, Me, Me.lstUsuariosExcluidos.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=yes))"
                
        limparCampos
        
    End If

End Sub

Private Sub cmdRestaurar_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim Matriz As Variant
Dim strMsg As String
Dim strTitulo As String
Dim strSelecao As String

    If Me.lstUsuariosExcluidos.value = "" Or IsNull(Me.lstUsuariosExcluidos.value) Then
        strMsg = "ATEN��O: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "RESTAURAR!"
        
        MsgBox strMsg, vbInformation + vbOKOnly, strTitulo
    Else

        Matriz = Array()
        Matriz = Split(Me.lstUsuariosExcluidos.value, " - ")
        
        admUsuarioExcluir strBanco, CStr(Matriz(1)), False
        
        ListBoxCarregar strBanco, Me, Me.lstUsuarios.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No))"
        ListBoxCarregar strBanco, Me, Me.lstUsuariosExcluidos.Name, "Pesquisa", "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=yes))"
        
        limparCampos

    End If

    
End Sub

''#########################################
''  CAMPOS
''#########################################

Private Sub txtNome_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtNome = UCase(Me.txtNome)
End Sub
Private Sub txtCodigo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtCodigo = UCase(Me.txtCodigo)
End Sub

Private Sub txtEmail_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtEmail = LCase(Me.txtEmail)
End Sub

''#########################################
''  LISTAS
''#########################################

Private Sub lstUsuarios_Click()
Dim Matriz As Variant

    Matriz = Array()
    Matriz = Split(Me.lstUsuarios.value, " - ")
    
    Me.cboDepartamento.Text = CStr(Matriz(0))
    Me.txtNome = CStr(Matriz(1))
    Me.txtEmail = CStr(Matriz(2))
    Me.txtCodigo = CStr(Matriz(3))
    
    Me.txtGerenteContas = CStr(Matriz(4))
    Me.txtTelefone = CStr(Matriz(5))
    Me.txtCelular01 = CStr(Matriz(6))
    Me.txtCelular02 = CStr(Matriz(7))
    Me.txtIdNextel = CStr(Matriz(8))
    
    Me.cmdSalvar.Enabled = True

End Sub

Private Sub lstUsuarios_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdExcluir_Click
End Sub

Private Sub lstUsuariosExcluidos_Click()
Dim Matriz As Variant

    Matriz = Array()
    Matriz = Split(Me.lstUsuariosExcluidos.value, " - ")
    
    Me.cboDepartamento.Text = CStr(Matriz(0))
    Me.txtNome = CStr(Matriz(1))
    Me.txtEmail = CStr(Matriz(2))
    Me.txtCodigo = CStr(Matriz(3))
    
    Me.txtGerenteContas = CStr(Matriz(4))
    Me.txtTelefone = CStr(Matriz(5))
    Me.txtCelular01 = CStr(Matriz(6))
    Me.txtCelular02 = CStr(Matriz(7))
    Me.txtIdNextel = CStr(Matriz(8))
    
    Me.cmdSalvar.Enabled = False
End Sub

Private Sub lstUsuariosExcluidos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdRestaurar_Click
End Sub

''#########################################
''  PROCEDIMENTOS
''#########################################

Private Sub limparCampos()

    Me.cboDepartamento.Text = "DPTO"
    Me.txtCodigo.Text = "CODIGO"
    Me.txtNome.Text = "NOME"
    Me.txtEmail.Text = "E-MAIL"
    Me.txtGerenteContas.Text = "GERENTE DE CONTAS"
    Me.txtTelefone.Text = "TELEFONE"
    Me.txtCelular01.Text = "CELULAR - 01"
    Me.txtCelular02.Text = "CELULAR - 02"
    Me.txtIdNextel.Text = "ID NEXTEL"
    
End Sub

