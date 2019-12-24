Attribute VB_Name = "mod_00_Administracao"
Option Base 1
Option Explicit

'''   CONTROLES DO SISTEMA
Public Const VersaoDoSistema As String = "131104-1919"

'' SENHAS
Public Const SenhaBloqueio As String = "Ge456B!"
Public Const SenhaBanco As String = "abc"

'' BANCOS DE DADOS
Public Const BancoLocal As String = "B1"

'' GUIA DE CONFIGURAÇÃO
Public Const cfgGuiaConfiguracao As String = "CFG"
Public Const cfgBancoServidor As String = "B2"
Public Const cfgBancoLocal As String = "B3"

'' CONTROLE DE USUÁRIOS
Public Const NomeUsuario As String = "A1"
Public Const AmbienteDeTrabalho As String = "A2"
Public Const GerenteDeContas As String = "C3"

'' PLANILHA
Public Const InicioCursor As String = "C4"
Public Const ArquivoInicio As String = "A76"
Public Const ArquivoControle As String = "C75"

'' PROJETO ATUAL
Public ProjetoAtual As String


Sub admAtualizarCaminhoBaseDados(strCaminhoDaBase As String, strTipoDeBase As String)
    DesbloqueioDeGuia SenhaBloqueio
    Range(strTipoDeBase) = strCaminhoDaBase
    BloqueioDeGuia SenhaBloqueio
End Sub

Sub admVincularBaseDados(strBaseDeDados As String)
    DesbloqueioDeGuia SenhaBloqueio
    Range(BancoLocal) = strBaseDeDados
    BloqueioDeGuia SenhaBloqueio
End Sub

Function SelecionarAmbienteDeTrabalho(Ind As Integer)
    SelecionarAmbienteDeTrabalho = Choose(Ind, "CASA", "ESCRITORIO")
End Function

Sub Pesquisar(ByVal Control As IRibbonControl)
    frmPesquisar.Show
End Sub

Sub Cadastro(ByVal Control As IRibbonControl)
    frmCadastro.Show
End Sub

Sub AnexosArquivos(ByVal Control As IRibbonControl)
    frmAnexosArquivos.Show
End Sub

Sub Enviar(ByVal Control As IRibbonControl)
    frmEnviar.Show
End Sub

Sub Indices(ByVal Control As IRibbonControl)
Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)
Dim strMSG As String
Dim strTitulo As String
    
If LiberarIndice(strBanco, strUsuario) = False Then
    strMSG = "Ops!!! " & Chr(10) & Chr(13) & Chr(13)
    strMSG = strMSG & "Você não tem permissão para acessar este conteúdo. " & Chr(10) & Chr(13) & Chr(13)
    strTitulo = "Indices de calculos!"
    
    MsgBox strMSG, vbInformation + vbOKOnly, strTitulo
Else
    
    LiberarIndice strBanco, strUsuario
    frmIndices.Show
    
End If

End Sub


Sub Projeto01(ByVal Control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

If ActiveSheet.Name = strUsuario Then
    Unload frmIndices
    Exit Sub
Else
    ProjetoAtual = "C"
    frmPojeto.Show
End If


End Sub

Sub Projeto02(ByVal Control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

If ActiveSheet.Name = strUsuario Then
    Unload frmIndices
    Exit Sub
Else
    ProjetoAtual = "D"
    frmPojeto.Show
End If

End Sub

Sub Projeto03(ByVal Control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

If ActiveSheet.Name = strUsuario Then
    Unload frmIndices
    Exit Sub
Else
    ProjetoAtual = "E"
    frmPojeto.Show
End If

End Sub

Sub Projeto04(ByVal Control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

If ActiveSheet.Name = strUsuario Then
    Unload frmIndices
    Exit Sub
Else
    ProjetoAtual = "F"
    frmPojeto.Show
End If

End Sub

Sub Projeto05(ByVal Control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

If ActiveSheet.Name = strUsuario Then
    Unload frmIndices
    Exit Sub
Else
    ProjetoAtual = "G"
    frmPojeto.Show
End If

End Sub

Sub Projeto06(ByVal Control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

If ActiveSheet.Name = strUsuario Then
    Unload frmIndices
    Exit Sub
Else
    ProjetoAtual = "H"
    frmPojeto.Show
End If

End Sub

Sub Projeto07(ByVal Control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

If ActiveSheet.Name = strUsuario Then
    Unload frmIndices
    Exit Sub
Else
    ProjetoAtual = "I"
    frmPojeto.Show
End If

End Sub

Sub Projeto08(ByVal Control As IRibbonControl)
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

If ActiveSheet.Name = strUsuario Then
    Unload frmIndices
    Exit Sub
Else
    ProjetoAtual = "J"
    frmPojeto.Show
End If

End Sub



''#################################################
''      ADMINISTRAÇÃO CENTRAL DE PROCEDIMENTOS
''#################################################




Public Function admNotificacoes(BaseDeDados As String, strControle As String, strVendedor As String, strEtapa As String)
On Error GoTo admNotificacoes_err

'   BASE DE DADOS
Dim dbOrcamento As dao.Database

' NOTIFICAÇÕES (TODOS)
'Dim rstNotificacoesTodos As DAO.Recordset
Dim strNotificacoesTodos As String

' COM ANEXOS
Dim rstNotificacoesComANEXOS As dao.Recordset
Dim strNotificacoesComANEXOS As String


' SEM ANEXOS
Dim rstNotificacoesSemANEXOS As dao.Recordset
Dim strNotificacoesSemANEXOS As String


    '   BASE DE DADOS
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    
    
    '   NOTIFICAÇÕES (TODOS)
    strNotificacoesTodos = "SELECT DISTINCT qryPermissoesUsuarios.Selecionado AS Status, qryPermissoesUsuarios.eMail, " & _
                      " qryPermissoesUsuarios.DPTO, qryPermissoesUsuarios.Usuario From qryPermissoesUsuarios WHERE " & _
                      "(((qryPermissoesUsuarios.Selecionado)='" & strEtapa & "') AND ((qryPermissoesUsuarios.DPTO)<>'Vendas') AND " & _
                      "((qryPermissoesUsuarios.Categoria)='Notificações')) UNION SELECT qryPermissoesUsuarios.Selecionado AS Status, " & _
                      "qryPermissoesUsuarios.eMail, qryPermissoesUsuarios.DPTO, qryPermissoesUsuarios.Usuario From qryPermissoesUsuarios WHERE " & _
                      "(((qryPermissoesUsuarios.Selecionado)='" & strEtapa & "') AND ((qryPermissoesUsuarios.DPTO)='Vendas') " & _
                      "AND ((qryPermissoesUsuarios.Usuario)='" & strVendedor & "') AND ((qryPermissoesUsuarios.Categoria)='Notificações'))"
    
    
    '   NOTIFICAÇÕES (COM ANEXOS)
    strNotificacoesComANEXOS = "SELECT DISTINCT qryPermissoesUsuarios.Selecionado AS Status, qryPermissoesUsuarios.eMail, " & _
                      " qryPermissoesUsuarios.DPTO, qryPermissoesUsuarios.Usuario From qryPermissoesUsuarios WHERE " & _
                      "(((qryPermissoesUsuarios.Selecionado)='" & strEtapa & "') AND ((qryPermissoesUsuarios.DPTO)<>'Vendas') AND " & _
                      "((qryPermissoesUsuarios.Categoria)='Anexos')) UNION SELECT qryPermissoesUsuarios.Selecionado AS Status, " & _
                      "qryPermissoesUsuarios.eMail, qryPermissoesUsuarios.DPTO, qryPermissoesUsuarios.Usuario From qryPermissoesUsuarios WHERE " & _
                      "(((qryPermissoesUsuarios.Selecionado)='" & strEtapa & "') AND ((qryPermissoesUsuarios.DPTO)='Vendas') " & _
                      "AND ((qryPermissoesUsuarios.Usuario)='" & strVendedor & "') AND ((qryPermissoesUsuarios.Categoria)='Anexos'))"
                      
    Set rstNotificacoesComANEXOS = dbOrcamento.OpenRecordset(strNotificacoesComANEXOS)
    Dim strConsultaAnexos As String
    Dim strConsultaAnexosCampoCaminho As String
    
    strConsultaAnexos = "Select * from qryOrcamentosArquivosAnexos where Vendedor = '" & strVendedor & "' AND Controle = '" & strControle & "'"
    strConsultaAnexosCampoCaminho = "OBS_01"
    
    While Not rstNotificacoesComANEXOS.EOF

        EnviarEmail rstNotificacoesComANEXOS.Fields("eMail"), strEtapa & " : " & strControle & " - " & strVendedor, True, BaseDeDados, strConsultaAnexos, strConsultaAnexosCampoCaminho
        rstNotificacoesComANEXOS.MoveNext

    Wend
    
                      
    '   NOTIFICAÇÕES (SEM ANEXOS)
    strNotificacoesSemANEXOS = "SELECT tmpNotificacoes.Status, tmpNotificacoes.eMail, tmpNotificacoes.DPTO, tmpNotificacoes.Usuario " & _
                               "FROM (" & strNotificacoesTodos & ") as tmpNotificacoes  " & _
                               "WHERE (((tmpNotificacoes.Usuario) Not In (Select tmpAnexos.Usuario from (" & strNotificacoesComANEXOS & ") as tmpAnexos)))"
    
    Set rstNotificacoesSemANEXOS = dbOrcamento.OpenRecordset(strNotificacoesSemANEXOS)
    While Not rstNotificacoesSemANEXOS.EOF

        EnviarEmail rstNotificacoesSemANEXOS.Fields("eMail"), strEtapa & " : " & strControle & " - " & strVendedor, False
        rstNotificacoesSemANEXOS.MoveNext

    Wend
    

admNotificacoes_Fim:

    rstNotificacoesComANEXOS.Close
    rstNotificacoesSemANEXOS.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstNotificacoesComANEXOS = Nothing
    Set rstNotificacoesSemANEXOS = Nothing
    
    Exit Function
admNotificacoes_err:
    
    MsgBox Err.Description, , "admNotificacoes"
    Resume admNotificacoes_Fim


End Function

Public Function admOrcamentoNovo(BaseDeDados As String, strVendedor As String) As Boolean: admOrcamentoNovo = True
' CADASTRAR NOVO ORÇAMENTO

On Error GoTo admOrcamentoNovo_err
Dim dbOrcamento As dao.Database
Dim qdfOrcamentoNovo As dao.QueryDef
Dim qdfOrcamentoNovoCustos As dao.QueryDef
Dim qdfOrcamentoNovoLinha As dao.QueryDef
Dim qdfOrcamentoNovoMoeda As dao.QueryDef
Dim qdfOrcamentoNovoVenda As dao.QueryDef
Dim qdfOrcamentoNovoDescontos As dao.QueryDef

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)

'' ORÇAMENTO
Set qdfOrcamentoNovo = dbOrcamento.QueryDefs("admOrcamentoNovo")
With qdfOrcamentoNovo

    .Parameters("NM_VENDEDOR") = strVendedor
    .Execute
    
End With

'' PREVISÕES DE CUSTOS
Set qdfOrcamentoNovoCustos = dbOrcamento.QueryDefs("admOrcamentoNovoCustos")
With qdfOrcamentoNovoCustos

    .Parameters("NM_VENDEDOR") = strVendedor
    .Execute
    
End With

'' LINHA DE PRODUTOS
Set qdfOrcamentoNovoLinha = dbOrcamento.QueryDefs("admOrcamentoNovoLinha")
With qdfOrcamentoNovoLinha

    .Parameters("NM_VENDEDOR") = strVendedor
    .Execute
    
End With

'' MOEDAS
Set qdfOrcamentoNovoMoeda = dbOrcamento.QueryDefs("admOrcamentoNovoMoeda")
With qdfOrcamentoNovoMoeda

    .Parameters("NM_VENDEDOR") = strVendedor
    .Execute
    
End With

'' VENDAS
Set qdfOrcamentoNovoVenda = dbOrcamento.QueryDefs("admOrcamentoNovoVenda")
With qdfOrcamentoNovoVenda

    .Parameters("NM_VENDEDOR") = strVendedor
    .Execute
    
End With


'' DESCONTOS
Set qdfOrcamentoNovoDescontos = dbOrcamento.QueryDefs("admOrcamentoNovoDescontos")
With qdfOrcamentoNovoDescontos

    .Parameters("NM_VENDEDOR") = strVendedor
    .Execute
    
End With


admOrcamentoNovo_Fim:
    dbOrcamento.Close
    qdfOrcamentoNovo.Close
    qdfOrcamentoNovoCustos.Close
    qdfOrcamentoNovoLinha.Close
    qdfOrcamentoNovoMoeda.Close
    qdfOrcamentoNovoVenda.Close
    qdfOrcamentoNovoDescontos.Close
    
    Set dbOrcamento = Nothing
    Set qdfOrcamentoNovo = Nothing
    Set qdfOrcamentoNovoCustos = Nothing
    Set qdfOrcamentoNovoLinha = Nothing
    Set qdfOrcamentoNovoMoeda = Nothing
    Set qdfOrcamentoNovoVenda = Nothing
    Set qdfOrcamentoNovoDescontos = Nothing
    
    Exit Function
admOrcamentoNovo_err:
    admOrcamentoNovo = False
    MsgBox Err.Description
    Resume admOrcamentoNovo_Fim

End Function

Public Function admOrcamentoCopiar(BaseDeDados As String, strControle_SELECAO As String, strVendedor_SELECAO As String, strVendedor_ATUAL As String) As Boolean: admOrcamentoCopiar = True
' CRIAR CÓPIA DE ORÇAMENTO

On Error GoTo admOrcamentoCopiar_err
Dim dbOrcamento As dao.Database
Dim qdfadmOrcamentoCopiar As dao.QueryDef
Dim qdfOrcamentoNovoCustos As dao.QueryDef
Dim qdfOrcamentoNovoLinha As dao.QueryDef
Dim qdfOrcamentoNovoMoeda As dao.QueryDef
Dim qdfOrcamentoNovoVenda As dao.QueryDef
Dim qdfOrcamentoNovoDescontos As dao.QueryDef

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)

'' ORÇAMENTO
Set qdfadmOrcamentoCopiar = dbOrcamento.QueryDefs("admOrcamentoCopiar")
With qdfadmOrcamentoCopiar

    .Parameters("NM_VENDEDOR") = strVendedor_ATUAL
    .Parameters("SELECAO_CONTROLE") = strControle_SELECAO
    .Parameters("SELECAO_VENDEDOR") = strVendedor_SELECAO
    
    .Execute
    
End With

'' PREVISÕES DE CUSTOS
Set qdfOrcamentoNovoCustos = dbOrcamento.QueryDefs("admOrcamentoNovoCustos")
With qdfOrcamentoNovoCustos

    .Parameters("NM_VENDEDOR") = strVendedor_ATUAL
    .Execute
    
End With

'' LINHA DE PRODUTOS
Set qdfOrcamentoNovoLinha = dbOrcamento.QueryDefs("admOrcamentoNovoLinha")
With qdfOrcamentoNovoLinha

    .Parameters("NM_VENDEDOR") = strVendedor_ATUAL
    .Execute
    
End With

'' MOEDAS
Set qdfOrcamentoNovoMoeda = dbOrcamento.QueryDefs("admOrcamentoNovoMoeda")
With qdfOrcamentoNovoMoeda

    .Parameters("NM_VENDEDOR") = strVendedor_ATUAL
    .Execute
    
End With

'' VENDAS
Set qdfOrcamentoNovoVenda = dbOrcamento.QueryDefs("admOrcamentoNovoVenda")
With qdfOrcamentoNovoVenda

    .Parameters("NM_VENDEDOR") = strVendedor_ATUAL
    .Execute
    
End With


'' DESCONTOS
Set qdfOrcamentoNovoDescontos = dbOrcamento.QueryDefs("admOrcamentoNovoDescontos")
With qdfOrcamentoNovoDescontos

    .Parameters("NM_VENDEDOR") = strVendedor_ATUAL
    .Execute
    
End With


admOrcamentoCopiar_Fim:
    dbOrcamento.Close
    qdfadmOrcamentoCopiar.Close
    qdfOrcamentoNovoCustos.Close
    qdfOrcamentoNovoLinha.Close
    qdfOrcamentoNovoMoeda.Close
    qdfOrcamentoNovoVenda.Close
    qdfOrcamentoNovoDescontos.Close
    
    Set dbOrcamento = Nothing
    Set qdfadmOrcamentoCopiar = Nothing
    Set qdfOrcamentoNovoCustos = Nothing
    Set qdfOrcamentoNovoLinha = Nothing
    Set qdfOrcamentoNovoMoeda = Nothing
    Set qdfOrcamentoNovoVenda = Nothing
    Set qdfOrcamentoNovoDescontos = Nothing
    
    Exit Function
admOrcamentoCopiar_err:
    admOrcamentoCopiar = False
    MsgBox Err.Description
    Resume admOrcamentoCopiar_Fim
End Function

Public Sub admOrcamentoFormulariosLiberar(strUsuario As String)
On Error GoTo admOrcamentoFormulariosLiberar_err
Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)
Dim dbOrcamento As dao.Database
Dim rstLiberarFormularios As dao.Recordset
Dim rstBloquearFormularios As dao.Recordset

Dim Matriz As Variant
'Dim bloqueio As Variant

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstLiberarFormularios = dbOrcamento.OpenRecordset("Select * from qryUsuariosFormularios where Usuario = '" & strUsuario & "'")
Set rstBloquearFormularios = dbOrcamento.OpenRecordset("qryFormularios")

Matriz = Array()

DesbloqueioDeGuia SenhaBloqueio
Application.ScreenUpdating = False

While Not rstBloquearFormularios.EOF
    Matriz = Split(rstBloquearFormularios.Fields("VALOR_02"), "-")
    
    OcultarLinhas (Matriz(0)), (Matriz(1)), True
    rstBloquearFormularios.MoveNext

Wend


While Not rstLiberarFormularios.EOF
    Matriz = Split(rstLiberarFormularios.Fields("Formulario"), "-")

    OcultarLinhas CStr(Matriz(0)), CStr(Matriz(1)), False
    rstLiberarFormularios.MoveNext

Wend

BloqueioDeGuia SenhaBloqueio
Application.ScreenUpdating = True

admOrcamentoFormulariosLiberar_Fim:
    rstBloquearFormularios.Close
    rstLiberarFormularios.Close
    dbOrcamento.Close
    Set rstLiberarFormularios = Nothing
    Set dbOrcamento = Nothing
    
    Exit Sub
admOrcamentoFormulariosLiberar_err:
    MsgBox Err.Description
    Resume admOrcamentoFormulariosLiberar_Fim

End Sub

Public Function admOrcamentoExcluirVirtual(BaseDeDados As String, strControle As String, strNOME As String, strMotivo As String) As Boolean: admOrcamentoExcluirVirtual = True
On Error GoTo admOrcamentoExcluirVirtual_err
Dim dbOrcamento As dao.Database
Dim qdfadmOrcamentoExcluir As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmOrcamentoExcluir = dbOrcamento.QueryDefs("admOrcamentoExcluirVirtual")

With qdfadmOrcamentoExcluir

    .Parameters("NM_VENDEDOR") = strNOME
    .Parameters("NM_CONTROLE") = strControle
    .Parameters("NM_MOTIVO") = strMotivo
    
    .Execute
    
End With

qdfadmOrcamentoExcluir.Close
dbOrcamento.Close

admOrcamentoExcluirVirtual_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmOrcamentoExcluir = Nothing
    
    Exit Function
admOrcamentoExcluirVirtual_err:
    admOrcamentoExcluirVirtual = False
    MsgBox Err.Description
    Resume admOrcamentoExcluirVirtual_Fim
End Function

Public Function admOrcamentoExcluirAnexos(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean
On Error GoTo admOrcamentoExcluirAnexos_err
Dim dbOrcamento As dao.Database
Dim qdfadmOrcamentoEtapaAvancar As dao.QueryDef
Dim strSQL As String


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmOrcamentoEtapaAvancar = dbOrcamento.QueryDefs("admOrcamentoExcluirAnexos")

With qdfadmOrcamentoEtapaAvancar
    
    .Parameters("NM_VENDEDOR") = strVendedor
    .Parameters("NM_CONTROLE") = strControle
    
    .Execute
    
End With

qdfadmOrcamentoEtapaAvancar.Close
dbOrcamento.Close

admOrcamentoExcluirAnexos_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmOrcamentoEtapaAvancar = Nothing
    
    Exit Function
admOrcamentoExcluirAnexos_err:
    admOrcamentoExcluirAnexos = False
    MsgBox Err.Description
    Resume admOrcamentoExcluirAnexos_Fim
End Function

Public Function CodigoEtapa(BaseDeDados As String, strEtapa As String)
On Error GoTo CodigoEtapa_err
Dim dbOrcamento As dao.Database
Dim rstEtapas As dao.Recordset

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstEtapas = dbOrcamento.OpenRecordset("Select * from qryEtapas where Status = '" & strEtapa & "'")

If Not rstEtapas.EOF Then CodigoEtapa = rstEtapas.Fields("Atual")

rstEtapas.Close
dbOrcamento.Close


CodigoEtapa_Fim:

    Set dbOrcamento = Nothing
    Set rstEtapas = Nothing
    
    Exit Function
CodigoEtapa_err:
    MsgBox Err.Description, vbInformation + vbOKOnly, "Código da etapa!!!"
    Resume CodigoEtapa_Fim


End Function



Public Function DescricaoEtapa(BaseDeDados As String, intEtapa As Integer)
On Error GoTo DescricaoEtapa_err
Dim dbOrcamento As dao.Database
Dim rstEtapas As dao.Recordset

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstEtapas = dbOrcamento.OpenRecordset("Select * from qryEtapas where Atual = " & intEtapa)

If Not rstEtapas.EOF Then DescricaoEtapa = rstEtapas.Fields("Status")


rstEtapas.Close
dbOrcamento.Close


DescricaoEtapa_Fim:

    Set dbOrcamento = Nothing
    Set rstEtapas = Nothing
    
    Exit Function
DescricaoEtapa_err:
    MsgBox Err.Description, vbInformation + vbOKOnly, "Código da etapa!!!"
    Resume DescricaoEtapa_Fim


End Function

Public Function admOrcamentoAtualizarEtapa(BaseDeDados As String, strControle As String, strVendedor As String, strEtapa As String) As Boolean: admOrcamentoAtualizarEtapa = True
On Error GoTo admOrcamentoAtualizarEtapa_err
Dim dbOrcamento As dao.Database
Dim qdfadmOrcamentoEtapaAvancar As dao.QueryDef
Dim strSQL As String


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmOrcamentoEtapaAvancar = dbOrcamento.QueryDefs("admOrcamentoAtualizarEtapa")

With qdfadmOrcamentoEtapaAvancar
    
    .Parameters("NM_VENDEDOR") = strVendedor
    .Parameters("NM_CONTROLE") = strControle
    .Parameters("NM_ETAPA") = strEtapa
    
    .Execute
    
End With

qdfadmOrcamentoEtapaAvancar.Close
dbOrcamento.Close

admOrcamentoAtualizarEtapa_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmOrcamentoEtapaAvancar = Nothing
    
    Exit Function
admOrcamentoAtualizarEtapa_err:
    admOrcamentoAtualizarEtapa = False
    MsgBox Err.Description
    Resume admOrcamentoAtualizarEtapa_Fim
End Function

Public Function admIntervalosDeEdicaoControle(BaseDeDados As String, strControle As String, strVendedor As String) As Boolean: admIntervalosDeEdicaoControle = True
On Error GoTo admIntervalosDeEdicaoControle_err
Dim dbOrcamento As dao.Database
Dim rstOrcamento As dao.Recordset
Dim rstIntervalos As dao.Recordset
Dim strOrcamento As String
Dim strIntervalos As String

Dim strSelecao As String


strOrcamento = "SELECT Orcamentos.* " & _
                " FROM Orcamentos  " & _
                " WHERE (((CONTROLE)='" & strControle & "') AND ((VENDEDOR)= '" & strVendedor & "')) "


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstOrcamento = dbOrcamento.OpenRecordset(strOrcamento)

strIntervalos = "Select * from qryEtapasIntervalosEdicoes where Departamento = '" & rstOrcamento.Fields("Departamento") & "' and Status = '" & rstOrcamento.Fields("Status") & "'"

Set rstIntervalos = dbOrcamento.OpenRecordset(strIntervalos)

DesbloqueioDeGuia SenhaBloqueio

rstIntervalos.MoveFirst
While Not rstIntervalos.EOF

'    strSelecao = Replace(rstIntervalos.Fields("Selecao"), ",", ";")
    
    strSelecao = rstIntervalos.Fields("Selecao")
    
    IntervaloEditacaoCriar rstIntervalos.Fields("Intervalo"), strSelecao
'    Debug.Print rstIntervalos.Fields("Intervalo")
    Debug.Print strSelecao
    rstIntervalos.MoveNext

Wend


BloqueioDeGuia SenhaBloqueio

admIntervalosDeEdicaoControle_Fim:
    rstOrcamento.Close
    rstIntervalos.Close
    dbOrcamento.Close

    Set dbOrcamento = Nothing
    Set rstIntervalos = Nothing
    Set rstOrcamento = Nothing

    Exit Function
admIntervalosDeEdicaoControle_err:
    admIntervalosDeEdicaoControle = False
    MsgBox Err.Description
    Resume admIntervalosDeEdicaoControle_Fim


End Function

Public Function admIntervalosDeEdicaoMarcarSelecao(BaseDeDados As String)
On Error GoTo BloqueioDeSelecao_err
Dim dbOrcamento As dao.Database
Dim rstIntervalos As dao.Recordset
Dim strIntervalos As String

strIntervalos = "qryEtapasIntervalosEdicoes"

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstIntervalos = dbOrcamento.OpenRecordset(strIntervalos)


While Not rstIntervalos.EOF
    
    MarcaTexto rstIntervalos.Fields("Selecao")
    rstIntervalos.MoveNext

Wend


BloqueioDeSelecao_Fim:
    rstIntervalos.Close
    dbOrcamento.Close

    Set dbOrcamento = Nothing
    Set rstIntervalos = Nothing
    
    Exit Function
BloqueioDeSelecao_err:
    MsgBox Err.Description
    Resume BloqueioDeSelecao_Fim
End Function

Public Sub admIntervalosDeEdicaoLimparSelecao(BaseDeDados As String)
On Error GoTo admIntervalosDeEdicaoLimparSelecao_err
Dim dbOrcamento As dao.Database
Dim rstIntervalos As dao.Recordset
Dim strIntervalos As String

strIntervalos = "qryEtapasIntervalosEdicoes"

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstIntervalos = dbOrcamento.OpenRecordset(strIntervalos)

While Not rstIntervalos.EOF
    
    LimparTemplate rstIntervalos.Fields("Selecao"), rstIntervalos.Fields("ValorPadrao")
    rstIntervalos.MoveNext

Wend

admIntervalosDeEdicaoLimparSelecao_Fim:
    rstIntervalos.Close
    dbOrcamento.Close

    Set dbOrcamento = Nothing
    Set rstIntervalos = Nothing
    
    Exit Sub
admIntervalosDeEdicaoLimparSelecao_err:
    MsgBox Err.Description
    Resume admIntervalosDeEdicaoLimparSelecao_Fim
End Sub

Public Sub admOrcamentoFormulariosLimpar()
Dim strBanco As String: strBanco = Range(BancoLocal)
        
    ' DESATIVA A ATUALIZAÇÃO DA TELA
    Application.ScreenUpdating = False
    
    ' DESBLOQUEIO DA GUIA
    DesbloqueioDeGuia SenhaBloqueio
    
    ' REMOVER TODOS
    IntervaloEditacaoRemoverTodos
        
    ' LIMPAR
    admIntervalosDeEdicaoLimparSelecao strBanco
        
    ' POSICIONAR CURSOR
    Range(InicioCursor).Select
        
    ' MARCAR
    admIntervalosDeEdicaoMarcarSelecao strBanco
        
    '   LIMPAR GERENTE DE CONTAS
    LimparTemplate "C3,J3", ""
    MarcaTexto "C3,J3"
    
    '   LIMPAR VALOR
    LimparTemplate "J4", "0"
    
'    admArquivosAnexosExcluir ArquivoInicio, ArquivoControle, True
    
                
    ' BLOQUEIO DA GUIA
    BloqueioDeGuia SenhaBloqueio
        
    ' DESATIVA A ATUALIZAÇÃO DA TELA
    Application.ScreenUpdating = True

End Sub

Public Sub admArquivosAnexosCarregar(BaseDeDados As String, strControle As String, strVendedor As String)

    'Declare a variable as a FileDialog object.
    Dim fd As FileDialog

    'Create a FileDialog object as a File Picker dialog box.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    'Declare a variable to contain the path
    'of each selected item. Even though the path is aString,
    'the variable must be a Variant because For Each...Next
    'routines only work with Variants and Objects.
    Dim vrtSelectedItem As Variant
    Dim x As Long

    
    'Use a With...End With block to reference the FileDialog object.
    With fd

        'Allow the user to select multiple files.
        .Filters.Clear
        .Filters.Add "Todos os arquivos", "*.*", 1
        .Title = "Abrir"
        
        .AllowMultiSelect = True

        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                CadastroAnexoArquivo BaseDeDados, strControle, strVendedor, CStr(vrtSelectedItem)
            Next
        'If the user presses Cancel...
        Else
        End If
    End With
    
'    Range(ArquivoInicio).Select

    'Set the object variable to Nothing.
    Set fd = Nothing

End Sub

Public Function admOrcamentoExcluirAnexoArquivo(BaseDeDados As String, strUsuario As String, strControle As String, strArquivo As String) As Boolean: admOrcamentoExcluirAnexoArquivo = True
On Error GoTo admOrcamentoExcluirAnexoArquivo_err
Dim dbOrcamento As dao.Database
Dim qdfadmOrcamentoExcluirAnexoArquivo As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmOrcamentoExcluirAnexoArquivo = dbOrcamento.QueryDefs("admOrcamentoExcluirAnexoArquivo")

With qdfadmOrcamentoExcluirAnexoArquivo

    .Parameters("NM_VENDEDOR") = strUsuario
    .Parameters("NM_CONTROLE") = strControle
    .Parameters("NM_ARQUIVO") = strArquivo
    
    .Execute
    
End With

qdfadmOrcamentoExcluirAnexoArquivo.Close
dbOrcamento.Close

admOrcamentoExcluirAnexoArquivo_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmOrcamentoExcluirAnexoArquivo = Nothing
    
    Exit Function
admOrcamentoExcluirAnexoArquivo_err:
    admOrcamentoExcluirAnexoArquivo = False
    MsgBox Err.Description
    Resume admOrcamentoExcluirAnexoArquivo_Fim
End Function

Public Function admQuantidadeDeAnexos( _
                                BaseDeDados As String, _
                                strControle As String, _
                                strVendedor As String, _
                                strPropriedade As String) As Integer

On Error GoTo QuantidadeDeAnexos_err

Dim dbOrcamento As dao.Database
Dim qdfQuantidadeDeAnexos As dao.Recordset


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfQuantidadeDeAnexos = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & _
                                                            "' and Vendedor = '" & strVendedor & _
                                                            "' and PROPRIEDADE = '" & strPropriedade & "'")
qdfQuantidadeDeAnexos.MoveLast
qdfQuantidadeDeAnexos.MoveFirst
admQuantidadeDeAnexos = qdfQuantidadeDeAnexos.RecordCount


QuantidadeDeAnexos_Fim:
    qdfQuantidadeDeAnexos.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set qdfQuantidadeDeAnexos = Nothing
    
    Exit Function
QuantidadeDeAnexos_err:
    MsgBox Err.Description
    Resume QuantidadeDeAnexos_Fim
End Function

Public Function admProxEtapa( _
                        BaseDeDados As String, _
                        strDepartamento As String, _
                        strStatus As String) As String

On Error GoTo admProxEtapa_err

Dim dbOrcamento As dao.Database
Dim rstAdmProxEtapa As dao.Recordset

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstAdmProxEtapa = dbOrcamento.OpenRecordset("Select * from qryEtapas where Departamento = '" & strDepartamento & "' and Status = '" & strStatus & "'")

admProxEtapa = rstAdmProxEtapa.Fields("PROXIMO")


admProxEtapa_Fim:
    rstAdmProxEtapa.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstAdmProxEtapa = Nothing
    
    Exit Function
admProxEtapa_err:
    MsgBox Err.Description
    Resume admProxEtapa_Fim

End Function

Public Function admVerificarBaseDeDados() As Boolean: admVerificarBaseDeDados = False
Dim fd As Office.FileDialog
Dim strArq As String

If Not getFileStatus(Range(BancoLocal)) Then

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.Add "BDs do Access", "*.MDB"
    fd.Title = "Por favor Selecione a Base de dados para uso da planilha "
    fd.AllowMultiSelect = False
    
    If fd.Show = -1 Then
        strArq = fd.SelectedItems(1)
    End If
    
    If strArq <> "" Then
        DesbloqueioDeGuia SenhaBloqueio
        Range(BancoLocal) = strArq
        BloqueioDeGuia SenhaBloqueio
    End If
    
End If

End Function

Public Function admUsuarioNovoDepartamentos(BaseDeDados As String, strUsuario As String) As Boolean: admUsuarioNovoDepartamentos = True
On Error GoTo admUsuarioNovoDepartamentos_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuarioNovoDepartamentos As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuarioNovoDepartamentos = dbOrcamento.QueryDefs("admUsuarioNovoDepartamentos")

With qdfadmUsuarioNovoDepartamentos

    .Parameters("NM_USUARIO") = strUsuario
    
    .Execute
    
End With

qdfadmUsuarioNovoDepartamentos.Close
dbOrcamento.Close

admUsuarioNovoDepartamentos_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovoDepartamentos = Nothing
    
    Exit Function
admUsuarioNovoDepartamentos_err:
    admUsuarioNovoDepartamentos = False
    MsgBox Err.Description
    Resume admUsuarioNovoDepartamentos_Fim
End Function

Public Function admUsuarioNovoFuncoes(BaseDeDados As String, strUsuario As String) As Boolean: admUsuarioNovoFuncoes = True
On Error GoTo admUsuarioNovoFuncoes_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuarioNovoFuncoes As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuarioNovoFuncoes = dbOrcamento.QueryDefs("admUsuarioNovoFuncoes")

With qdfadmUsuarioNovoFuncoes

    .Parameters("NM_USUARIO") = strUsuario
    
    .Execute
    
End With

qdfadmUsuarioNovoFuncoes.Close
dbOrcamento.Close

admUsuarioNovoFuncoes_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovoFuncoes = Nothing
    
    Exit Function
admUsuarioNovoFuncoes_err:
    admUsuarioNovoFuncoes = False
    MsgBox Err.Description
    Resume admUsuarioNovoFuncoes_Fim
End Function

Public Function admUsuarioNovoNotificacoes(BaseDeDados As String, strUsuario As String) As Boolean: admUsuarioNovoNotificacoes = True
On Error GoTo admUsuarioNovoNotificacoes_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuarioNovoNotificacoes As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuarioNovoNotificacoes = dbOrcamento.QueryDefs("admUsuarioNovoNotificacoes")

With qdfadmUsuarioNovoNotificacoes

    .Parameters("NM_USUARIO") = strUsuario
    .Execute
    
End With

qdfadmUsuarioNovoNotificacoes.Close
dbOrcamento.Close

admUsuarioNovoNotificacoes_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovoNotificacoes = Nothing
    
    Exit Function
admUsuarioNovoNotificacoes_err:
    admUsuarioNovoNotificacoes = False
    MsgBox Err.Description
    Resume admUsuarioNovoNotificacoes_Fim
End Function

Public Function admUsuarioNovoStatus(BaseDeDados As String, strUsuario As String) As Boolean: admUsuarioNovoStatus = True
On Error GoTo admUsuarioNovoStatus_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuarioNovoStatus As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuarioNovoStatus = dbOrcamento.QueryDefs("admUsuarioNovoStatus")

With qdfadmUsuarioNovoStatus

    .Parameters("NM_USUARIO") = strUsuario
    
    .Execute
    
End With

qdfadmUsuarioNovoStatus.Close
dbOrcamento.Close

admUsuarioNovoStatus_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovoStatus = Nothing
    
    Exit Function
admUsuarioNovoStatus_err:
    admUsuarioNovoStatus = False
    MsgBox Err.Description
    Resume admUsuarioNovoStatus_Fim
End Function

Public Function admUsuarioNovoUsuarios(BaseDeDados As String, strUsuario As String) As Boolean: admUsuarioNovoUsuarios = True
On Error GoTo admUsuarioNovoUsuarios_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuarioNovoUsuarios As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuarioNovoUsuarios = dbOrcamento.QueryDefs("admUsuarioNovoUsuarios")

With qdfadmUsuarioNovoUsuarios

    .Parameters("NM_USUARIO") = strUsuario
    
    .Execute
    
End With

qdfadmUsuarioNovoUsuarios.Close
dbOrcamento.Close

admUsuarioNovoUsuarios_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovoUsuarios = Nothing
    
    Exit Function
admUsuarioNovoUsuarios_err:
    admUsuarioNovoUsuarios = False
    MsgBox Err.Description
    Resume admUsuarioNovoUsuarios_Fim
End Function

Public Function ExistenciaUsuario(BaseDeDados As String, strCODIGO As String, strNOME As String) As Boolean: ExistenciaUsuario = False
On Error GoTo ExistenciaUsuario_err
Dim dbOrcamento As dao.Database
Dim rstExistenciaUsuario As dao.Recordset
Dim strSQL As String
Dim RetVal As Variant

RetVal = Dir(BaseDeDados)

If RetVal = "" Then

    ExistenciaUsuario = True
    
Else
   
    strSQL = "SELECT * FROM qryUsuarios WHERE Usuario = '" & strNOME & "' AND  Codigo = '" & strCODIGO & "' "
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstExistenciaUsuario = dbOrcamento.OpenRecordset(strSQL)
      
    If rstExistenciaUsuario.EOF Then
        ExistenciaUsuario = False
    Else
        ExistenciaUsuario = True
    End If
        
    rstExistenciaUsuario.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstExistenciaUsuario = Nothing
    
End If

ExistenciaUsuario_Fim:
  
    Exit Function
ExistenciaUsuario_err:
    ExistenciaUsuario = True
    MsgBox Err.Description
    Resume ExistenciaUsuario_Fim
End Function

Public Function admUsuarioNovo(BaseDeDados As String, strDPTO As String, strCODIGO As String, strNOME As String, strEmail As String) As Boolean: admUsuarioNovo = True
On Error GoTo admUsuarioNovo_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuarioNovo As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuarioNovo = dbOrcamento.QueryDefs("admUsuarioNovo")

With qdfadmUsuarioNovo

    .Parameters("CODUSUARIO") = strCODIGO
    .Parameters("NOME_USUARIO") = strNOME
    .Parameters("EMAIL_USUARIO") = strEmail
    .Parameters("DPTO_USUARIO") = strDPTO
    
    .Execute
    
End With

admUsuarioNovoDepartamentos BaseDeDados, strNOME
admUsuarioNovoFuncoes BaseDeDados, strNOME
admUsuarioNovoNotificacoes BaseDeDados, strNOME
admUsuarioNovoStatus BaseDeDados, strNOME
admUsuarioNovoUsuarios BaseDeDados, strNOME

qdfadmUsuarioNovo.Close
dbOrcamento.Close

admUsuarioNovo_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioNovo = Nothing
    
    Exit Function
admUsuarioNovo_err:
    admUsuarioNovo = False
    MsgBox Err.Description
    Resume admUsuarioNovo_Fim
End Function

Public Function admUsuarioSalvar(BaseDeDados As String, strDPTO As String, strCODIGO As String, strNOME As String, strEmail As String) As Boolean: admUsuarioSalvar = True
On Error GoTo admUsuarioSalvar_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuarioSalvar As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuarioSalvar = dbOrcamento.QueryDefs("admUsuarioSalvar")

With qdfadmUsuarioSalvar

    .Parameters("CODUSUARIO") = strCODIGO
    .Parameters("NOME_USUARIO") = strNOME
    .Parameters("EMAIL_USUARIO") = strEmail
    .Parameters("DPTO_USUARIO") = strDPTO
    
    .Execute
    
End With

qdfadmUsuarioSalvar.Close
dbOrcamento.Close

admUsuarioSalvar_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioSalvar = Nothing
    
    Exit Function
admUsuarioSalvar_err:
    admUsuarioSalvar = False
    MsgBox Err.Description
    Resume admUsuarioSalvar_Fim
End Function

Public Function admUsuarioExcluir(BaseDeDados As String, strNOME As String, Excluir As Boolean) As Boolean: admUsuarioExcluir = True
On Error GoTo admUsuarioExcluir_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuarioExcluir As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuarioExcluir = dbOrcamento.QueryDefs("admUsuarioExcluir")

With qdfadmUsuarioExcluir

    .Parameters("NOME_USUARIO") = strNOME
    .Parameters("EXCLUSAO") = Excluir
    
    .Execute
    
End With

qdfadmUsuarioExcluir.Close
dbOrcamento.Close

admUsuarioExcluir_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioExcluir = Nothing
    
    Exit Function
admUsuarioExcluir_err:
    admUsuarioExcluir = False
    MsgBox Err.Description
    Resume admUsuarioExcluir_Fim
End Function

Public Function admUsuarioCopiar(BaseDeDados As String, strUSUARIO_DESTINO As String, strUSUARIO_SELECAO As String) As Boolean: admUsuarioCopiar = True
On Error GoTo admUsuarioCopiar_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuarioCopiar As dao.QueryDef
Dim qdfadmUsuarioCopiarConfiguracao As dao.QueryDef


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuarioCopiar = dbOrcamento.QueryDefs("admUsuarioCopiar")
Set qdfadmUsuarioCopiarConfiguracao = dbOrcamento.QueryDefs("admUsuarioCopiarConfiguracao")

With qdfadmUsuarioCopiar

    .Parameters("NM_USUARIO_SELECAO") = strUSUARIO_SELECAO
    
    .Execute
    
End With


With qdfadmUsuarioCopiarConfiguracao

    .Parameters("NM_USUARIO_SELECAO") = strUSUARIO_SELECAO
    .Parameters("NM_USUARIO_DESTINO") = strUSUARIO_DESTINO
    .Execute
    
End With





qdfadmUsuarioCopiarConfiguracao.Close
qdfadmUsuarioCopiar.Close
dbOrcamento.Close

admUsuarioCopiar_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuarioCopiar = Nothing
    Set qdfadmUsuarioCopiarConfiguracao = Nothing
    
    Exit Function
admUsuarioCopiar_err:
    admUsuarioCopiar = False
    MsgBox Err.Description
    Resume admUsuarioCopiar_Fim
End Function

Public Function admUsuariosPermissoes(BaseDeDados As String, strUsuario As String, strPERMISSAO As String, strCategoria As String) As Boolean: admUsuariosPermissoes = True
On Error GoTo admUsuariosPermissoes_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuariosPermissoes As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuariosPermissoes = dbOrcamento.QueryDefs("admUsuariosPermissoes")

With qdfadmUsuariosPermissoes

    .Parameters("NM_USUARIO") = strUsuario
    .Parameters("NM_PERMISSAO") = strPERMISSAO
    .Parameters("NM_CATEGORIA") = strCategoria
    
    .Execute
    
End With

qdfadmUsuariosPermissoes.Close
dbOrcamento.Close

admUsuariosPermissoes_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuariosPermissoes = Nothing
    
    Exit Function
admUsuariosPermissoes_err:
    admUsuariosPermissoes = False
    MsgBox Err.Description
    Resume admUsuariosPermissoes_Fim
End Function

Public Function admUsuariosPermissoesExcluir(BaseDeDados As String, strUsuario As String, strPERMISSAO As String, strCategoria As String) As Boolean: admUsuariosPermissoesExcluir = True
On Error GoTo admUsuariosPermissoesExcluir_err
Dim dbOrcamento As dao.Database
Dim qdfadmUsuariosPermissoesExcluir As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmUsuariosPermissoesExcluir = dbOrcamento.QueryDefs("admUsuariosPermissoesExcluir")

With qdfadmUsuariosPermissoesExcluir

    .Parameters("NM_USUARIO") = strUsuario
    .Parameters("NM_PERMISSAO") = strPERMISSAO
    .Parameters("NM_CATEGORIA") = strCategoria
    
    .Execute
    
End With

qdfadmUsuariosPermissoesExcluir.Close
dbOrcamento.Close

admUsuariosPermissoesExcluir_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmUsuariosPermissoesExcluir = Nothing
    
    Exit Function
admUsuariosPermissoesExcluir_err:
    admUsuariosPermissoesExcluir = False
    MsgBox Err.Description
    Resume admUsuariosPermissoesExcluir_Fim
End Function

Public Function EtapaUsuario(BaseDeDados As String, strCategoria As String, strNOME As String) As Boolean: EtapaUsuario = False
On Error GoTo EtapaUsuario_err
Dim dbOrcamento As dao.Database
Dim rstEtapaUsuario As dao.Recordset
Dim strSQL As String
Dim RetVal As Variant

RetVal = Dir(BaseDeDados)

If RetVal <> "" Then
   
    strSQL = "SELECT * FROM qryPermissoesUsuarios WHERE Usuario = '" & strNOME & "' AND  Categoria = 'Status' AND Selecionado = '" & strCategoria & "' "
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstEtapaUsuario = dbOrcamento.OpenRecordset(strSQL)
      
    If rstEtapaUsuario.EOF Then
        EtapaUsuario = False
    Else
        EtapaUsuario = True
    End If
        
    rstEtapaUsuario.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstEtapaUsuario = Nothing
    
End If

EtapaUsuario_Fim:
  
    Exit Function
EtapaUsuario_err:
    EtapaUsuario = True
    MsgBox Err.Description
    Resume EtapaUsuario_Fim
End Function


Public Function BloqueioEtapaUsuario(BaseDeDados As String, strCategoria As String, strNOME As String) As Boolean: BloqueioEtapaUsuario = False
On Error GoTo BloqueioEtapaUsuario_err
Dim dbOrcamento As dao.Database
Dim rstBloqueioEtapaUsuario As dao.Recordset
Dim strSQL As String
Dim RetVal As Variant

RetVal = Dir(BaseDeDados)

If RetVal <> "" Then
   
    strSQL = "SELECT * FROM qryPermissoesUsuarios WHERE Usuario = '" & strNOME & "' AND  Categoria = 'Bloqueio' AND Selecionado = '" & strCategoria & "' "
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstBloqueioEtapaUsuario = dbOrcamento.OpenRecordset(strSQL)
      
    If rstBloqueioEtapaUsuario.EOF Then
        BloqueioEtapaUsuario = True
    Else
        BloqueioEtapaUsuario = False
    End If
        
    rstBloqueioEtapaUsuario.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstBloqueioEtapaUsuario = Nothing
    
End If

BloqueioEtapaUsuario_Fim:
  
    Exit Function
BloqueioEtapaUsuario_err:
    BloqueioEtapaUsuario = True
    MsgBox Err.Description
    Resume BloqueioEtapaUsuario_Fim
End Function


Public Function UsuarioDepartamento(BaseDeDados As String, strUsuario As String) As Boolean
On Error GoTo UsuarioDepartamento_err
Dim dbOrcamento As dao.Database
Dim rstUsuarioDepartamento As dao.Recordset
Dim strSQL As String
Dim RetVal As Variant

RetVal = Dir(BaseDeDados)

If RetVal <> "" Then
   
    strSQL = "SELECT DPTO FROM qryUsuarios WHERE Usuario = '" & strUsuario & "' and DPTO in ('Financeiro','Produção')"
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstUsuarioDepartamento = dbOrcamento.OpenRecordset(strSQL)
      
    If Not rstUsuarioDepartamento.EOF Then
        UsuarioDepartamento = False
    Else
        UsuarioDepartamento = True
    End If
    
    rstUsuarioDepartamento.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstUsuarioDepartamento = Nothing
    
End If

UsuarioDepartamento_Fim:
  
    Exit Function
UsuarioDepartamento_err:
    UsuarioDepartamento = True
    MsgBox Err.Description
    Resume UsuarioDepartamento_Fim
End Function



Public Function admExecutarTarefa(BaseDeDados As String, strTarefa As String)
On Error GoTo admExecutarTarefa_err
Dim dbOrcamento As dao.Database
Dim qdfTarefa As dao.QueryDef

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfTarefa = CurrentDb.QueryDefs(strTarefa)

qdfTarefa.Execute

qdfTarefa.Close
dbOrcamento.Close

admExecutarTarefa_Fim:

    Set dbOrcamento = Nothing
    Set qdfTarefa = Nothing
    
    Exit Function
admExecutarTarefa_err:
    admExecutarTarefa = False
    MsgBox Err.Description
    Resume admExecutarTarefa_Fim
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''       NOVO
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub AmbienteDeTrabalhoDefinir(ByVal Ambiente As String)
'Dim strBancoServidor As String: strBancoServidor = Sheets(cfgGuiaConfiguracao).Range(cfgBancoServidor)
'Dim strBancoLocal As String: strBancoLocal = pathWorkSheetAddress & "bin\" & Controle & "_db" & "HOME" & ".mdb"
'
'''' VERIFICAR EXISTENCIA (BANCO_SERVER)
'If Dir$(strBancoServidor, vbArchive) <> "" Then
'
'    ''' DESBLOQUEIO DE PLANILHA
'    DesbloqueioDeGuia SenhaBloqueio
'    Application.ScreenUpdating = False
'
'    Select Case Ambiente
'
'        Case "CASA"
'            ''' COPIAR BASE DE DADOS (SERVER) PARA PASTA LOCAL
'            FileCopy strBancoServidor, strBancoLocal
'
'            ''' EXCLUIR ORCAMENTOS SEM VINCULOS COM USUARIO
'            admExcluirOrcamentosSemVinculosComUsuario strBancoLocal, Range(NomeUsuario)
'
'            ''' ARMAZENAR BANCO SELECIONADO EM CONFIGIRAÇÕES DO SISTEMA (BANCO LOCAL)
'            Sheets(cfgGuiaConfiguracao).Range(cfgBancoLocal) = strBancoLocal
'
'            ''' SETA AMBIENTE DE TRABALHO COMO: CASA
'            Range(AmbienteDeTrabalho) = Ambiente
'
'        Case "ESCRITORIO"
'
'            ''' SETA AMBIENTE DE TRABALHO COMO: CASA
'            Range(AmbienteDeTrabalho) = Ambiente
'
'    End Select
'
'    ''' BLOQUEIO DE PLANILHA
'    BloqueioDeGuia SenhaBloqueio
'    Application.ScreenUpdating = True
'
'Else
'
'
'
'End If


End Sub

Sub admAtualizarDaGuiaDeApoio()

Dim BaseDeDados As String: BaseDeDados = Range(BancoLocal)
Dim strConsultas(10) As String
Dim x As Integer

strConsultas(1) = "qryApoio_Journal"
strConsultas(2) = "qryApoio_Publisher"
strConsultas(3) = "qryApoio_Clientes"
strConsultas(4) = "qryApoio_Acabamento"
strConsultas(5) = "qryApoio_Idiomas"
strConsultas(6) = "qryApoio_Tipo"
strConsultas(7) = "qryApoio_Papel"
strConsultas(8) = "qryApoio_N_Paginas"
strConsultas(9) = "qryApoio_Impressao"
strConsultas(10) = "qryApoio_Formato"

Application.ScreenUpdating = False

For x = 1 To 10
    
    AtualizarListagens BaseDeDados, strConsultas(x), "APOIO", 2, x

Next x

Application.ScreenUpdating = True


End Sub

Sub AtualizarListagens(BaseDeDados As String, Consulta As String, Guia As String, Linha As Integer, Coluna As Integer)
Dim dbOrcamento As dao.Database
Dim rstListagem As dao.Recordset

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstListagem = dbOrcamento.OpenRecordset("Select * from " & Consulta & " as tmpListagem ")

'' LIMPAR CELULAS

Sheets(Guia).Visible = -1

With Sheets(Guia)
    .Select
    .Cells(Linha, Coluna).Select
    .Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

End With

Do While Not rstListagem.EOF

    With Sheets(Guia)
        .Cells(Linha, Coluna).Value = rstListagem.Fields("DESCRICAO")
        rstListagem.MoveNext
        Linha = Linha + 1
    End With
    
Loop

Sheets(Guia).Visible = 2
rstListagem.Close

End Sub

Public Function admCategoriaLimparTabela(BaseDeDados As String) As Boolean
On Error GoTo admCategoriaLimparTabela_err
Dim dbOrcamento As dao.Database
Dim qdfadmCategoriaLimparTabela As dao.QueryDef
Dim strSQL As String


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfadmCategoriaLimparTabela = dbOrcamento.QueryDefs("admCategoriaLimparTabela")

With qdfadmCategoriaLimparTabela
    
    .Execute
    .Close
    
End With

dbOrcamento.Close

admCategoriaLimparTabela_Fim:

    Set dbOrcamento = Nothing
    Set qdfadmCategoriaLimparTabela = Nothing
    
    Exit Function
admCategoriaLimparTabela_err:
    admCategoriaLimparTabela = False
    MsgBox Err.Description
    Resume admCategoriaLimparTabela_Fim
End Function

Public Function admExcluirOrcamentosSemVinculosComUsuario(BaseDeDados As String, strUsuario As String)
''' DEIXAR APENAS ORCAMENTOS VINCULADOS AO VENDEDOR
On Error GoTo admOrcamentosVinculadosVendedor_err
Dim dbOrcamento As dao.Database
Dim qdfExclusao As dao.QueryDef
Dim strSQL(3) As String

Dim l As Integer, c As Integer

Dim x As Integer ' contador de linhas
Dim y As Integer ' contador de colunas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)

strSQL(1) = "admOrcamentosEXCLUSAO"
strSQL(2) = "admOrcamentosCustosEXCLUSAO"
strSQL(3) = "admOrcamentosAnexosEXCLUSAO"


'Saida Now(), "admOrcamentosVinculadosVendedor.log"

For x = 1 To UBound(strSQL)
    
    Set qdfExclusao = dbOrcamento.QueryDefs(strSQL(x))
    
    With qdfExclusao
    
        .Parameters("NM_VENDEDOR") = strUsuario
        
        .Execute
        
    End With
    
    qdfExclusao.Close

Next
    
dbOrcamento.Close

'Saida Now(), "admOrcamentosVinculadosVendedor.log"

admOrcamentosVinculadosVendedor_Fim:

    Set dbOrcamento = Nothing
    Set qdfExclusao = Nothing
    
    Exit Function
admOrcamentosVinculadosVendedor_err:
    MsgBox Err.Description
    Resume admOrcamentosVinculadosVendedor_Fim
End Function

Public Function CaminhoDoBancoOffice(BaseDeDados As String, strTipo As String) As String
On Error GoTo CaminhoDoBancoOffice_err
Dim dbOrcamento As dao.Database
Dim rstBanco As dao.Recordset

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstBanco = dbOrcamento.OpenRecordset("Select * from qryBancoDeDados where Tipo = '" & strTipo & "'")

If Not rstBanco.EOF Then CaminhoDoBancoOffice = rstBanco.Fields("OrigemDoBancoOffice")


rstBanco.Close
dbOrcamento.Close


CaminhoDoBancoOffice_Fim:

    Set dbOrcamento = Nothing
    Set rstBanco = Nothing
    
    Exit Function
CaminhoDoBancoOffice_err:
    MsgBox Err.Description, vbInformation + vbOKOnly, "Caminho do Banco!!!"
    Resume CaminhoDoBancoOffice_Fim


End Function


Public Function admGerenciarApoioExcluir(BaseDeDados As String, strListagemDeApoio As String, strNomeApoio As String) As Boolean: admGerenciarApoioExcluir = True
On Error GoTo admGerenciarApoioExcluir_err
Dim dbOrcamento As dao.Database
Dim qdfQuery As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfQuery = dbOrcamento.QueryDefs("admGerenciarApoioExcluir")

With qdfQuery

    .Parameters("NM_APOIO") = strListagemDeApoio
    .Parameters("NM_EXCLUIR") = strNomeApoio
    
    .Execute
    
End With

qdfQuery.Close
dbOrcamento.Close

admGerenciarApoioExcluir_Fim:

    Set dbOrcamento = Nothing
    Set qdfQuery = Nothing
    
    Exit Function
admGerenciarApoioExcluir_err:
    admGerenciarApoioExcluir = False
    MsgBox Err.Description
    Resume admGerenciarApoioExcluir_Fim
End Function


Public Function admGerenciarApoioAterar(BaseDeDados As String, strListagemDeApoio As String, strNomeAntigo As String, strNomeNovo As String) As Boolean: admGerenciarApoioAterar = True
On Error GoTo admGerenciarApoioAterar_err
Dim dbOrcamento As dao.Database
Dim qdfQuery As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfQuery = dbOrcamento.QueryDefs("admGerenciarApoioAterar")

With qdfQuery

    .Parameters("NM_APOIO") = strListagemDeApoio
    .Parameters("NM_ANTIGO") = strNomeAntigo
    .Parameters("NM_NOVO") = strNomeNovo
    
    .Execute
    
End With

qdfQuery.Close
dbOrcamento.Close

admGerenciarApoioAterar_Fim:

    Set dbOrcamento = Nothing
    Set qdfQuery = Nothing
    
    Exit Function
admGerenciarApoioAterar_err:
    admGerenciarApoioAterar = False
    MsgBox Err.Description
    Resume admGerenciarApoioAterar_Fim
End Function

Public Function admGerenciarApoioIncluir(BaseDeDados As String, strListagemDeApoio As String, strNomeNovo As String) As Boolean: admGerenciarApoioIncluir = True
On Error GoTo admGerenciarApoioIncluir_err
Dim dbOrcamento As dao.Database
Dim qdfQuery As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfQuery = dbOrcamento.QueryDefs("admGerenciarApoioIncluir")

With qdfQuery

    .Parameters("NM_APOIO") = strListagemDeApoio
    .Parameters("NM_NOVO") = strNomeNovo
    
    .Execute
    
End With

qdfQuery.Close
dbOrcamento.Close

admGerenciarApoioIncluir_Fim:

    Set dbOrcamento = Nothing
    Set qdfQuery = Nothing
    
    Exit Function
admGerenciarApoioIncluir_err:
    admGerenciarApoioIncluir = False
    MsgBox Err.Description
    Resume admGerenciarApoioIncluir_Fim
End Function


Public Function admGerenciarIndiceExcluir(BaseDeDados As String, strIndice As String, strNomeIndice As String) As Boolean: admGerenciarIndiceExcluir = True
On Error GoTo admGerenciarIndiceExcluir_err
Dim dbOrcamento As dao.Database
Dim qdfQuery As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfQuery = dbOrcamento.QueryDefs("admGerenciarIndiceExcluir")

With qdfQuery

    .Parameters("NM_INDICE") = strIndice
    .Parameters("NM_EXCLUIR") = strNomeIndice

    .Execute

End With

qdfQuery.Close
dbOrcamento.Close

admGerenciarIndiceExcluir_Fim:

    Set dbOrcamento = Nothing
    Set qdfQuery = Nothing

    Exit Function
admGerenciarIndiceExcluir_err:
    admGerenciarIndiceExcluir = False
    MsgBox Err.Description
    Resume admGerenciarIndiceExcluir_Fim
End Function


Public Function admGerenciarIndiceAterar(BaseDeDados As String, strIndice As String, strNomeAntigo As String, strNomeNovo As String, strValor_01 As String, strValor_02 As String) As Boolean: admGerenciarIndiceAterar = True
On Error GoTo admGerenciarIndiceAterar_err
Dim dbOrcamento As dao.Database
Dim qdfQuery As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfQuery = dbOrcamento.QueryDefs("admGerenciarIndiceAterar")

With qdfQuery

    .Parameters("NM_INDICE") = strIndice
    .Parameters("NM_ANTIGO") = strNomeAntigo
    .Parameters("NM_NOVO") = strNomeNovo
    .Parameters("VALOR_01") = strValor_01
    .Parameters("VALOR_02") = strValor_02
    
    .Execute
    
End With

qdfQuery.Close
dbOrcamento.Close

admGerenciarIndiceAterar_Fim:

    Set dbOrcamento = Nothing
    Set qdfQuery = Nothing
    
    Exit Function
admGerenciarIndiceAterar_err:
    admGerenciarIndiceAterar = False
    MsgBox Err.Description
    Resume admGerenciarIndiceAterar_Fim
End Function

Public Function admGerenciarIndiceIncluir(BaseDeDados As String, strIndice As String, strNomeIndice As String, strValor_01 As String, strValor_02 As String) As Boolean: admGerenciarIndiceIncluir = True
On Error GoTo admGerenciarIndiceIncluir_err
Dim dbOrcamento As dao.Database
Dim qdfQuery As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfQuery = dbOrcamento.QueryDefs("admGerenciarIndiceIncluir")

With qdfQuery

    .Parameters("NM_INDICE") = strIndice
    .Parameters("NM_DESCRICAO") = strNomeIndice
    .Parameters("VALOR_01") = strValor_01
    .Parameters("VALOR_02") = strValor_02
    
    .Execute
    
End With

qdfQuery.Close
dbOrcamento.Close

admGerenciarIndiceIncluir_Fim:

    Set dbOrcamento = Nothing
    Set qdfQuery = Nothing
    
    Exit Function
admGerenciarIndiceIncluir_err:
    admGerenciarIndiceIncluir = False
    MsgBox Err.Description
    Resume admGerenciarIndiceIncluir_Fim
End Function


Public Function admGerenciarIndicesDeCalculos _
                    (BaseDeDados As String, _
                        strVendedor As String, _
                        strControle As String, _
                        strPropriedade As String, _
                        strIndice As String, _
                        strValor_01 As String, _
                        strValor_02 As String) _
                        As Boolean: admGerenciarIndicesDeCalculos = True
                        
On Error GoTo admGerenciarIndicesDeCalculos_err
Dim dbOrcamento As dao.Database
Dim qdfQuery As dao.QueryDef
Dim strSQL As String

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfQuery = dbOrcamento.QueryDefs("admGerenciarIndicesDeCalculos")

With qdfQuery

    .Parameters("NM_VENDEDOR") = strVendedor
    .Parameters("NM_CONTROLE") = strControle
    .Parameters("NM_PROPRIEDADE") = strPropriedade
    .Parameters("NM_INDICE") = strIndice
    .Parameters("NM_VALOR01") = strValor_01
    .Parameters("NM_VALOR02") = strValor_02
    
    .Execute
    
End With

qdfQuery.Close
dbOrcamento.Close

admGerenciarIndicesDeCalculos_Fim:

    Set dbOrcamento = Nothing
    Set qdfQuery = Nothing
    
    Exit Function
admGerenciarIndicesDeCalculos_err:
    admGerenciarIndicesDeCalculos = False
    MsgBox Err.Description
    Resume admGerenciarIndicesDeCalculos_Fim
End Function

Public Function LiberarIndice(BaseDeDados As String, strUsuario As String) As Boolean
On Error GoTo LiberarIndice_err
Dim dbOrcamento As dao.Database
Dim rstIndice As dao.Recordset

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstIndice = dbOrcamento.OpenRecordset("Select * from qryPermissoesUsuarios where Categoria = 'Indice' and Usuario = '" & strUsuario & "'")

If Not rstIndice.EOF Then
    LiberarIndice = True
Else
    LiberarIndice = False
End If


rstIndice.Close
dbOrcamento.Close


LiberarIndice_Fim:

    Set dbOrcamento = Nothing
    Set rstIndice = Nothing
    
    Exit Function
LiberarIndice_err:
    MsgBox Err.Description, vbInformation + vbOKOnly, "Liberar Indice."
    Resume LiberarIndice_Fim


End Function
