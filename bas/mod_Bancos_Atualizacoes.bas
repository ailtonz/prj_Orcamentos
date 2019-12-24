Attribute VB_Name = "mod_Bancos_Atualizacoes"
Option Explicit

Sub ReceberAtualizacoes(ByVal strPastaChegada As String)
''' RECEBER ATUALIZAÇÕES
Dim strBancoOrigem As String: strBancoOrigem = Range(BancoLocal)
Dim Matriz As Variant
Dim x As Long
Dim y As Long
Dim Lista As String
Dim CATEGORIA As String
    
    Matriz = Array()
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' LISTAR ARQUIVOS COMPACTADOS, DESCOMPACTA-LOS E EXCLUI-LOS
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''' LISTAR
    Lista = ListarDiretorio(strPastaChegada, "*.zip")
    Matriz = Split(Lista, ";")
    
    For x = 0 To UBound(Matriz)
        ''' DESCOMPACTAR
'        UnZip strPastaChegada & Matriz(x), strPastaChegada
        DesCompact strPastaChegada & Matriz(x), strPastaChegada
        
        ''' EXCLUIR
        If Dir$(strPastaChegada & Matriz(x)) <> "" Then Kill strPastaChegada & Matriz(x)
    Next x
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' LISTAR ARQUIVOS DESCOMPACTADOS, EXPORTA-LOS AO BANCO PRINCIPAL E EXCLUI-LOS
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''' LISTAR
    Lista = ListarDiretorio(strPastaChegada, "*.mdb")
    Matriz = Split(Lista, ";")

    For y = 0 To UBound(Matriz)

        CATEGORIA = getFileStep(CStr(Matriz(y)))

        Select Case CATEGORIA

            Case "TRANSITO"

                BancoEmTransito_Importar strPastaChegada & Matriz(y)

            Case "ATUALIZACAO"

                ''' LIMPAR CONFIGURAÇÕES ANTIGAS
                admCategoriaLimparTabela strBancoOrigem

                ''' CARREGAR NOVAS CONFIGURAÇÕES
                ExportarDadosTabelaOrigemParaTabelaDestino strPastaChegada & Matriz(y), strBancoOrigem, "admCategorias"

                ''' ATUALIZAR GUIAS DE APOIO
                admAtualizarDaGuiaDeApoio


        End Select

        ''' EXCLUIR
        If Dir$(strPastaChegada & Matriz(y)) <> "" Then Kill strPastaChegada & Matriz(y)
    Next y
    
    
End Sub

Sub BancoEmTransito_Importar(ByVal strBancoEmTransito As String)
On Error GoTo cmdImportar_err

Dim strBancoLocal As String: strBancoLocal = Range(BancoLocal)

Dim dbOrigem As DAO.Database
Dim rstOrcamentoORIGEM As DAO.Recordset

Set dbOrigem = DBEngine.OpenDatabase(strBancoEmTransito, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstOrcamentoORIGEM = dbOrigem.OpenRecordset("Select * from Orcamentos Order by Vendedor,Controle")

    While Not rstOrcamentoORIGEM.EOF
        ''' IMPORTAR ORÇAMENTOS EM TRANSITO
        ExportarOrcamento strBancoEmTransito, strBancoLocal, rstOrcamentoORIGEM.Fields("Controle"), rstOrcamentoORIGEM.Fields("Vendedor")
        rstOrcamentoORIGEM.MoveNext

    Wend


cmdImportar_Fim:

    MsgBox "Importação concluída", vbInformation + vbOKOnly, "Importar de Orçamento(s)"

    Exit Sub
cmdImportar_err:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Importar de Orçamento(s)"
    Resume cmdImportar_Fim

End Sub

Function CriarArquivoDeAtualizacaoDoSistema() As String
Dim strBancoOrigem As String: strBancoOrigem = Range(BancoLocal)
Dim strBancoDestino As String: strBancoDestino = pathWorkSheetAddress & Controle & "_db" & "ATUALIZACAO" & ".mdb"
Dim strArquivoCompactado As String: strArquivoCompactado = Left(strBancoDestino, Len(strBancoDestino) - 3) & "zip"

    ''' CRIA BASE DE DADOS PARA EXPORTAÇÃO DE DADOS
    CriarBancoParaExportacao strBancoDestino
    
    ''' CRIAR TABELA(S) EM BASE DE DADOS DE EXPORTAÇÃO
    CriarTabelaEmBancoParaExportacao strBancoOrigem, strBancoDestino, "admCategorias"
    
    ''' EXPORTAR DADOS DA TABELA ADMINISTRATIVA DO SISTEMA
    ExportarDadosTabelaOrigemParaTabelaDestino strBancoOrigem, strBancoDestino, "admCategorias"
    
    ''' COMPACTA BASE DE DADOS
    Zip strBancoDestino, strArquivoCompactado
    
    ''' DELETA BASE DE DADOS TEMPORARIA
    If Dir$(strBancoDestino) <> "" Then Kill strBancoDestino
    
    ''' RETORNO DE NOME DO ARQUIVO CRIADO
    CriarArquivoDeAtualizacaoDoSistema = strArquivoCompactado

End Function
