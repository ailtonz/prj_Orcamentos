VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPesquisar 
   Caption         =   "Pesquisa de Orçamentos"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11280
   OleObjectBlob   =   "frmPesquisar.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPesquisar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit

Public strPesquisar As String
Public strSQL As String
Public strUsuarios As String

Private Sub cboAmbienteDeTrabalho_Click()
Dim strBancoServidor As String: strBancoServidor = Sheets(cfgGuiaConfiguracao).Range(cfgBancoServidor)
Dim strBancoLocal As String: strBancoLocal = pathWorkSheetAddress & "bin\" & Controle & "_db" & "HOME" & ".mdb"
Dim strAmbiente As String: strAmbiente = Me.cboAmbienteDeTrabalho.Text


''' VERIFICAR EXISTENCIA (BANCO_SERVER)
If Not Dir$(strBancoServidor, vbArchive) <> "" Then

    ''' MENSAGEM DE ERRO DE PROCEDIMENTO
    MsgBox "Troca de Ambiente INTERROMPIDA!", vbCritical + vbOKOnly, "Troca de Ambiente"
    
Else

    ''' DESBLOQUEIO DE PLANILHA
    DesbloqueioDeGuia SenhaBloqueio
    Application.ScreenUpdating = False

    Select Case strAmbiente
    
        Case "CASA"
            ''' COPIAR BASE DE DADOS (SERVER) PARA PASTA LOCAL
            FileCopy strBancoServidor, strBancoLocal
            
            ''' EXCLUIR ORCAMENTOS SEM VINCULOS COM USUARIO
            admExcluirOrcamentosSemVinculosComUsuario strBancoLocal, Range(NomeUsuario)
            
            ''' ARMAZENAR BANCO SELECIONADO EM CONFIGIRAÇÕES DO SISTEMA (BANCO LOCAL)
            Sheets(cfgGuiaConfiguracao).Range(cfgBancoLocal) = strBancoLocal
            
            ''' SETA AMBIENTE DE TRABALHO COMO: CASA
            Range(AmbienteDeTrabalho) = strAmbiente
            
            ''' CADASTRA CAMINHO DO BANCO
            Range(BancoLocal) = strBancoLocal
        
        Case "ESCRITORIO"
        
            ''' SETA AMBIENTE DE TRABALHO COMO: ESCRITORIO
            Range(AmbienteDeTrabalho) = strAmbiente
        
            ''' CADASTRA CAMINHO DO BANCO
            Range(BancoLocal) = strBancoServidor
        
        
    End Select
    
    ''' BLOQUEIO DE PLANILHA
    BloqueioDeGuia SenhaBloqueio
    Application.ScreenUpdating = True
    
    ''' ATUALIZAR TITULO DA TELA
    Me.Caption = UCase(strAmbiente & " - " & " Pesquisa de Orçamentos ")
    
    ''' MENSAGEM DE CONCLUSÃO DE PROCEDIMENTO
    MsgBox "Troca de Ambiente Concluida!", vbInformation + vbOKOnly, "Troca de Ambiente"
    
End If
    
End Sub

Private Sub spbEtapas_Change()
Dim strBanco As String: strBanco = Range(BancoLocal)
    
    Me.txtEtapa = DescricaoEtapa(strBanco, Me.spbEtapas.Value)
    
    MontarPesquisa
    
    ListBoxCarregar strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
    
    Me.lstPesquisa.Enabled = EtapaUsuario(strBanco, Me.txtEtapa, Range(NomeUsuario))
    
    Me.lstPesquisa.Enabled = BloqueioEtapaUsuario(strBanco, Me.txtEtapa, Range(NomeUsuario))
    
    ''' DESBLOQUEIO DE FUNÇÕES
    UserFormDesbloqueioDeFuncoes strBanco, Me, "Select * from qryUsuariosFuncoes Where Usuario = '" & strUsuarios & "'", "Funcao"
    
    
    Me.Repaint
    
End Sub

Private Sub UserForm_Activate()
    
'    admVerificarBaseDeDados
'    AmbienteDeTrabalhoCarregar

    admOrcamentoFormulariosLimpar
    admOrcamentoFormulariosLiberar Range(NomeUsuario)
    ActiveSheet.Name = IIf(IsNull(Range(NomeUsuario)), "SEM_USUARIO", Range(NomeUsuario))
    Range(InicioCursor).Select
    spbEtapas_Change

       
    UserForm_Initialize
        
End Sub

''#########################################
''  FORMULARIO
''#########################################

Private Sub UserForm_Initialize()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim sqlUsuarios As String: strUsuarios = Range(NomeUsuario)
Dim strAmbiente As String: strAmbiente = Range(AmbienteDeTrabalho)
    
    ''' ATUALIZAR TITULO DA TELA
    Me.Caption = UCase(strAmbiente & " - " & " Pesquisa de Orçamentos ")
    
    ''' VERIFICAR EXISTENCIA DA BASE DE DADOS
    admVerificarBaseDeDados
    
    ''' MONTA PESQUISA
    MontarPesquisa
    
    ''' CARREGA VARIAVEL DE USUÁRIOS
    sqlUsuarios = "Select * from qryUsuarios WHERE (((qryUsuarios.ExclusaoVirtual)=No)) Order By Usuario"
    
'    Saida strSQL, "Pesquisa.log"
    
    ''' CARREGAR LIST BOX DE PESQUISA
    ListBoxCarregar strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
    
    ''' CARREGAR LIST BOX DE USUÁRIOS
    ComboBoxCarregar strBanco, Me.cboUsuario, "Usuario", sqlUsuarios
    
    ''' SELECIONAR USUÁRIO
    Me.cboUsuario.Text = strUsuarios
    
''    ''' DESBLOQUEIO DE FUNÇÕES
''    UserFormDesbloqueioDeFuncoes strBanco, Me, "Select * from qryUsuariosFuncoes Where Usuario = '" & strUsuarios & "'", "Funcao"
    
    ''' CARREGAR COMBO BOX DE AMBIENTE DE TRABALHO
    ComboBoxUpdate "cfg", "BANCOS", Me.cboAmbienteDeTrabalho
    
    ''' DESATIVA ATUALIZAÇÃO DA TELA
    Application.ScreenUpdating = False
    ''' DESBLOQUEIA GUIA
    DesbloqueioDeGuia SenhaBloqueio
    
    ''' LIMPAR LINHA DE PRODUTOS
    Range("L3:N3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ''' LIMPAR MOEDA
    Range("P3:Q3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ''' LIMPAR VENDA
    Range("S3:T3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ''' LIMPAR DESCONTOS
    Range("V3:W3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ''' BLOQUEIA GUIA
    BloqueioDeGuia SenhaBloqueio
    ''' ATIVA ATUALIZAÇÃO DA TELA
    Application.ScreenUpdating = True
    
    '''POSICIONA CURSOR
    Range(InicioCursor).Select
              
End Sub

''#########################################
''  COMANDOS
''#########################################

Private Sub cmdPesquisar_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)

Dim retValor As Variant

    retValor = InputBox("Digite uma palavra para fazer o filtro:", "Filtro", strPesquisar, 0, 0)
    strPesquisar = retValor
    
    MontarPesquisa
    
    ListBoxCarregar strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
    Me.Repaint

End Sub

Private Sub cmdNovo_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim sqlUsuario As String: strUsuarios = Range(NomeUsuario)

    admOrcamentoNovo strBanco, Me.cboUsuario
    ListBoxCarregar strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
    
End Sub

Private Sub cmdAlterar_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim Matriz As Variant
Dim strMSG As String
Dim strTitulo As String

    If IsNull(lstPesquisa.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "ALTERAR ORÇAMENTO!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
        Matriz = Array()
        Matriz = Split(lstPesquisa.Value, " - ")

        Application.ScreenUpdating = False

        admOrcamentoFormulariosLimpar
        
        CarregarOrcamento strBanco, CStr(Matriz(0)), CStr(Matriz(2))
        
        admIntervalosDeEdicaoControle strBanco, CStr(Matriz(0)), CStr(Matriz(2))
                        
        Range(InicioCursor).Select
        
        ActiveSheet.Name = CStr(Matriz(0))
            
        Application.ScreenUpdating = True
        
        frmPesquisar.Hide
        
    End If

End Sub

Private Sub cmdCopiar_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)
Dim Matriz As Variant
Dim strMSG As String
Dim strTitulo As String

    If IsNull(lstPesquisa.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "CÓPIA!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
        Matriz = Array()
        Matriz = Split(lstPesquisa.Value, " - ")
        
        admOrcamentoCopiar strBanco, CStr(Matriz(0)), CStr(Matriz(2)), strUsuario
        ListBoxCarregar strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
    End If
End Sub

Private Sub cmdExcluir_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim Matriz As Variant
Dim strMSG As String
Dim strTitulo As String
Dim varResposta As Variant


    If IsNull(Me.lstPesquisa.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "EXCLUIR!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else
        strMSG = "ATENÇÃO: Você deseja realmente EXCLUIR este registro?. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "EXCLUIR!"
        
        varResposta = MsgBox(strMSG, vbInformation + vbYesNo, strTitulo)
    
        If varResposta = vbYes Then
            Matriz = Array()
            Matriz = Split(lstPesquisa.Value, " - ")
    
    
            varResposta = InputBox("Informe o motivo pelo qual o Orçamento foi Excluido.", "Motivo da exclusão")
    
            If varResposta <> "" Then
            
                If admOrcamentoExcluirVirtual(strBanco, CStr(Matriz(0)), CStr(Matriz(2)), CStr(varResposta)) Then
                    strMSG = "Exclusão concluida com sucesso!" & Chr(10) & Chr(13) & Chr(13)
                    strTitulo = "EXCLUIR!"
                    
                    ListBoxCarregar strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
                Else
                    strMSG = "Operação cancelada! " & Chr(10) & Chr(13) & Chr(13)
                    strTitulo = "EXCLUIR!"
                End If
            
            Else
                strMSG = "Operação cancelada! " & Chr(10) & Chr(13) & Chr(13)
                strTitulo = "EXCLUIR!"
            End If
            
            MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
        Else
            strMSG = "Operação cancelada! " & Chr(10) & Chr(13) & Chr(13)
            strTitulo = "EXCLUIR!"
            
            MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
        End If
        
    End If
    
End Sub

Private Sub cmdBanco_Click()
' VINCULAR BANCO DE DADOS

Dim strMSG As String
Dim strTitulo As String
Dim strBanco As String

strBanco = SelecionarBanco

    If strBanco <> "" Then
    
        DesbloqueioDeGuia SenhaBloqueio
        Range(BancoLocal).Value = strBanco
        BloqueioDeGuia SenhaBloqueio
        
    Else
        strMSG = "Por favor Selecione a Base de dados para uso da planilha "
        strTitulo = "Seleção de Base de dados"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    End If
    

End Sub

Private Sub cmdControleDeUsuarios_Click()
    frmAdministracao.Show
End Sub

Private Sub cmdUsuarioPadrao_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)

    DesbloqueioDeGuia SenhaBloqueio
    Application.ScreenUpdating = False
    
    Range(NomeUsuario) = Me.cboUsuario
    
    ListBoxCarregar strBanco, Me, Me.lstPesquisa.Name, "Pesquisa", strSQL
    
    admOrcamentoFormulariosLiberar Range(NomeUsuario)
    
    ActiveSheet.Name = Me.cboUsuario
    
    Me.Repaint
    
    'POSICIONA CURSOR
    Range(InicioCursor).Select
    
    Application.ScreenUpdating = True
    BloqueioDeGuia SenhaBloqueio
    
    UserForm_Initialize

End Sub

Private Sub cmdVoltarEtapa_Click()
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strMSG As String
Dim strTitulo As String
Dim RetVal As Variant
Dim Matriz As Variant


    If IsNull(lstPesquisa.Value) Then
        strMSG = "ATENÇÃO: Por favor selecione um item da lista. " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Voltar Etapa"

        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    Else


        Matriz = Array()
        Matriz = Split(lstPesquisa.Value, " - ")

        Application.ScreenUpdating = False

        admOrcamentoFormulariosLimpar
        
        CarregarOrcamento strBanco, CStr(Matriz(0)), CStr(Matriz(2))

        admIntervalosDeEdicaoControle strBanco, CStr(Matriz(0)), CStr(Matriz(2))
        
        ActiveSheet.Name = CStr(Matriz(0))
            
        Application.ScreenUpdating = True

        frmEtapas.Show

    End If


End Sub

''#########################################
''  LISTAS
''#########################################

Private Sub lstPesquisa_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim strMSG As String
Dim strTitulo As String
    
    
    If Me.cmdAlterar.Enabled Then
        cmdAlterar_Click
    Else
        strMSG = "Função Bloqueada! " & Chr(10) & Chr(13) & Chr(13)
        strTitulo = "Alterar!"
        
        MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
    
    End If
    
End Sub

''#########################################
''  PROCEDIMENTOS
''#########################################


Private Sub MontarPesquisa()

''''
'''' ORIGINAL
''''

'strSQL = "SELECT qryOrcamentosListar.Pesquisa FROM qryOrcamentosListar WHERE ((qryOrcamentosListar.Pesquisa) Like '*" & strPesquisar & "*')"
'strSQL = strSQL + " AND ((qryOrcamentosListar.VENDEDOR) In (Select Descricao01 from admCategorias where codRelacao = (SELECT admCategorias.codCategoria FROM admCategorias WHERE ((admCategorias.Categoria)='" & strUsuarios & "')) and Categoria = 'Usuarios'))"
'strSQL = strSQL + " AND ((qryOrcamentosListar.DEPARTAMENTO) In (Select Descricao01 from admCategorias where codRelacao = (SELECT admCategorias.codCategoria FROM admCategorias WHERE ((admCategorias.Categoria)='" & strUsuarios & "')) and Categoria = 'Departamentos')) "
'strSQL = strSQL + " AND ((qryOrcamentosListar.STATUS) In (Select Descricao01 from admCategorias where codRelacao = (SELECT admCategorias.codCategoria FROM admCategorias WHERE admCategorias.Categoria = '" & strUsuarios & "') and Categoria = 'Status')) "
'strSQL = strSQL + "ORDER BY qryOrcamentosListar.CONTROLE DESC , qryOrcamentosListar.VENDEDOR"

strSQL = "SELECT qryOrcamentosListar.Pesquisa FROM qryOrcamentosListar WHERE ((qryOrcamentosListar.Pesquisa) Like '*" & strPesquisar & "*')"
strSQL = strSQL + " AND ((qryOrcamentosListar.VENDEDOR) In (Select Descricao01 from admCategorias where codRelacao = (SELECT admCategorias.codCategoria FROM admCategorias WHERE ((admCategorias.Categoria)='" & strUsuarios & "')) and Categoria = 'Usuarios'))"
strSQL = strSQL + " AND ((qryOrcamentosListar.DEPARTAMENTO) In (Select Descricao01 from admCategorias where codRelacao = (SELECT admCategorias.codCategoria FROM admCategorias WHERE ((admCategorias.Categoria)='" & strUsuarios & "')) and Categoria = 'Departamentos')) "
strSQL = strSQL + " AND ((qryOrcamentosListar.STATUS) In ('" & Me.txtEtapa & "')) "
strSQL = strSQL + "ORDER BY qryOrcamentosListar.Codigo DESC"






End Sub


























