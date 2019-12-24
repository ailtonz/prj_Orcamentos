Attribute VB_Name = "mod_02_Cadastrar"
Option Base 1
Option Explicit

Public Function CadastroOrcamento( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String) As Boolean: CadastroOrcamento = True
On Error GoTo CadastroOrcamento_err

Dim dbOrcamento As DAO.Database
Dim qdfCadastroOrcamento As DAO.QueryDef
Dim strSql As String

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroOrcamento = dbOrcamento.QueryDefs("CadastroOrcamento")

    With qdfCadastroOrcamento
    
        .Parameters("NOME_VENDEDOR") = strVendedor
        .Parameters("NUMERO_CONTROLE") = strControle
        
        .Parameters("NM_CLIENTE") = Range("C4").value
        .Parameters("NM_RESPONSAVEL") = Range("C5").value
        
        .Parameters("DTPEDIDO") = Range("G3").value
        .Parameters("PREVENTREGA") = Range("G4").value
        
        .Parameters("VALORPROJETO") = Range("J4").value
        .Parameters("NM_PUBLISHER") = Range("C8").value
        .Parameters("NM_JOURNAL") = Range("C9").value
        .Parameters("NM_PAGS") = Range("C10").value
        
        CadastroOrcamentoProjeto BaseDeDados, strControle, strVendedor, Range("C6").value
    
        'FECHADO COM CLIENTE ( VENDIDO )
        l = 12
        c = 3
        For x = 1 To 8
            .Parameters(x & "FECHADO") = Cells(l, c).value
            c = c + 1
        Next x
        
        'LINHA DE PRODUTOS
        l = 13
        c = 3
        For x = 1 To 4
            .Parameters(x & "LINHA_PRODUTO") = Cells(l, c).value
            c = c + 1
        Next x
        
        'FASCICULOS
        l = 14
        c = 3
        For x = 1 To 4
            .Parameters(x & "FASCICULOS") = Cells(l, c).value
            c = c + 1
        Next x
                
        'VENDA
        l = 15
        c = 3
        For x = 1 To 8
            .Parameters(x & "VENDA") = Cells(l, c).value
            c = c + 1
        Next x
                
        'IMPOSTO
        l = 16
        c = 3
        For x = 1 To 8
            .Parameters(x & "IMPOSTO") = Cells(l, c).value
            c = c + 1
        Next x
        
        'IDIOMA
        l = 17
        c = 3
        For x = 1 To 8
            .Parameters(x & "IDIOMA") = Cells(l, c).value
            c = c + 1
        Next x
        
        'TIRAGEM
        l = 18
        c = 3
        For x = 1 To 8
            .Parameters(x & "TIRAGEM") = Cells(l, c).value
            c = c + 1
        Next x
                
        'ESPECIFICACAO
        l = 19
        c = 3
        For x = 1 To 8
            .Parameters(x & "ESPECIFICACAO") = Cells(l, c).value
            c = c + 1
        Next x
                
        'MOEDA
        l = 20
        c = 3
        For x = 1 To 8
            .Parameters(x & "MOEDA") = Cells(l, c).value
            c = c + 1
        Next x
        
        'ROYALTY PERCENTUAL
        l = 21
        c = 3
        For x = 1 To 8
            .Parameters(x & "ROYALTY_PERCENTUAL") = Cells(l, c).value
            c = c + 1
        Next x
            
        'ROYALTY ESPECIE
        l = 22
        c = 3
        For x = 1 To 8
            .Parameters(x & "ROYALTY_ESPECIE") = Cells(l, c).value
            c = c + 1
        Next x
                
        'RE IMPRESSAO
        l = 23
        c = 3
        For x = 1 To 8
            .Parameters(x & "RE_IMPRESSAO") = Cells(l, c).value
            c = c + 1
        Next x
            
'        'DESCONTO - ( PREÇOS )
'        L = 61
'        c = 3
'        For x = 1 To 8
'            .Parameters(x & "DESCONTO") = Cells(L, c).Value
'            c = c + 1
'        Next x
        
        'PREÇO MKT
        l = 65
        c = 3
        For x = 1 To 4
            .Parameters(x & "PrecoMKT") = Cells(l, c).value
            c = c + 1
        Next x

        'DESCONTO PADRÃO
        l = 71
        c = 3
        For x = 1 To 4
            .Parameters(x & "DescontoPadrao") = Cells(l, c).value
            c = c + 1
        Next x

        'PREÇO COMPRA TOTAL
        l = 73
        c = 3
        For x = 1 To 4
            .Parameters(x & "PrecoTotal") = Cells(l, c).value
            c = c + 1
        Next x
        
        'ARREDONDAMENTO
        l = 83
        c = 3
        For x = 1 To 4
            .Parameters(x & "Arredondamento") = Cells(l, c).value
            c = c + 1
        Next x
                
        .Execute
        
    End With

    CadastroAnexoDesconto BaseDeDados, strControle, strVendedor, 3, 22
    CadastroAnexoLinha BaseDeDados, strControle, strVendedor, 3, 12
    CadastroAnexoMoeda BaseDeDados, strControle, strVendedor, 3, 16
    CadastroAnexoVenda BaseDeDados, strControle, strVendedor, 3, 19
    
    

'    admOrcamentoExcluirAnexos BaseDeDados, strControle, strVendedor
    
'    'ARQUIVOS - ( ANEXOS )
'    Dim Terminio As Integer
'    Dim Inicio As Integer
'
'    Terminio = CInt(Range(ArquivoControle) - 1)
'    Inicio = CInt(Right(ArquivoInicio, Len(ArquivoInicio) - 1))
'
'    If Terminio > Inicio Then
'        l = Inicio
'        c = 2
'        For x = Inicio To Terminio
'            CadastroAnexoArquivo BaseDeDados, strControle, strVendedor, Cells(l, c).Value
'            l = l + 1
'        Next x
'    End If


CadastroOrcamento_Fim:
    qdfCadastroOrcamento.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set qdfCadastroOrcamento = Nothing
    
    Exit Function
CadastroOrcamento_err:
    CadastroOrcamento = False
    MsgBox Err.Description
    Resume CadastroOrcamento_Fim
End Function

Public Function CadastroOrcamentoImpressao( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String)

On Error GoTo CadastroOrcamentoImpressao_err

Dim dbOrcamento As DAO.Database
Dim qdfCadastroOrcamentoImpressao As DAO.QueryDef
Dim strSql As String

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroOrcamentoImpressao = dbOrcamento.QueryDefs("CadastroOrcamentoImpressao")

With qdfCadastroOrcamentoImpressao

    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle

    'TIPO
    l = 25
    c = 3
    For x = 1 To 4
        .Parameters(x & "TIPO") = Cells(l, c).value
        c = c + 1
    Next x

    'PAPEL
    l = 26
    c = 3
    For x = 1 To 4
        .Parameters(x & "PAPEL") = Cells(l, c).value
        c = c + 1
    Next x

    'PAGINAS
    l = 27
    c = 3
    For x = 1 To 4
        .Parameters(x & "PAGINAS") = Cells(l, c).value
        c = c + 1
    Next x

    'IMPRESSAO
    l = 28
    c = 3
    For x = 1 To 4
        .Parameters(x & "IMPRESSAO") = Cells(l, c).value
        c = c + 1
    Next x
    
    'FORMATO
    l = 29
    c = 3
    For x = 1 To 4
        .Parameters(x & "FORMATO") = Cells(l, c).value
        c = c + 1
    Next x

    'ACABAMENTO
    l = 31
    c = 2
    For x = 1 To 4
        CadastroOrcamentoAcabamento BaseDeDados, strControle, strVendedor, x & "_ACABAMENTO", Cells(l, c).value
        l = l + 1
    Next x
    
    .Execute
    
End With

qdfCadastroOrcamentoImpressao.Close
dbOrcamento.Close


CadastroOrcamentoImpressao_Fim:

    Set dbOrcamento = Nothing
    Set qdfCadastroOrcamentoImpressao = Nothing
    
    Exit Function
CadastroOrcamentoImpressao_err:
    CadastroOrcamentoImpressao = False
    MsgBox Err.Description
    Resume CadastroOrcamentoImpressao_Fim
End Function

Public Function CadastroOrcamentoCustos( _
                                 BaseDeDados As String, _
                                 strControle As String, _
                                 strVendedor As String) As Boolean: CadastroOrcamentoCustos = True

On Error GoTo CadastroOrcamentoCustos_err

Dim dbOrcamento As DAO.Database
Dim qdfCadastroCustos01 As DAO.QueryDef
Dim qdfCadastroCustos02 As DAO.QueryDef
Dim qdfCadastroCustos03 As DAO.QueryDef
Dim strSql As String

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroCustos01 = dbOrcamento.QueryDefs("CadastroOrcamentoCustos01")
Set qdfCadastroCustos02 = dbOrcamento.QueryDefs("CadastroOrcamentoCustos02")
'Set qdfCadastroCustos03 = dbOrcamento.QueryDefs("CadastroOrcamentoCustos03")

With qdfCadastroCustos01

    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle

    'INDEXACAO
    l = 37
    c = 3
    For x = 1 To 8
        .Parameters(x & "INDEXACAO") = Cells(l, c).value
        c = c + 1
    Next x
    
    'TRADUCAO
    l = 38
    c = 3
    For x = 1 To 8
        .Parameters(x & "TRADUCAO") = Cells(l, c).value
        c = c + 1
    Next x
    
    'REVISAO ORTOGRAFICA
    l = 39
    c = 3
    For x = 1 To 8
        .Parameters(x & "REVISAO_ORTOGRAFICA") = Cells(l, c).value
        c = c + 1
    Next x
    
    'REVISAO MEDICA
    l = 40
    c = 3
    For x = 1 To 8
        .Parameters(x & "REVISAO_MEDICA") = Cells(l, c).value
        c = c + 1
    Next x
    
    'CRIACAO
    l = 41
    c = 3
    For x = 1 To 8
        .Parameters(x & "CRIACAO") = Cells(l, c).value
        c = c + 1
    Next x
    
    'ILUSTRACAO
    l = 42
    c = 3
    For x = 1 To 8
        .Parameters(x & "ILUSTRACAO") = Cells(l, c).value
        c = c + 1
    Next x
    
    'REVISAO
    l = 43
    c = 3
    For x = 1 To 8
        .Parameters(x & "REVISAO") = Cells(l, c).value
        c = c + 1
    Next x
    
    'DIAGRAMACAO
    l = 44
    c = 3
    For x = 1 To 8
        .Parameters(x & "DIAGRAMACAO") = Cells(l, c).value
        c = c + 1
    Next x
    
    'MEDICO
    l = 45
    c = 3
    For x = 1 To 8
        .Parameters(x & "MEDICO") = Cells(l, c).value
        c = c + 1
    Next x
    
    'GRAFICA
    l = 46
    c = 3
    For x = 1 To 8
        .Parameters(x & "GRAFICA") = Cells(l, c).value
        c = c + 1
    Next x
    
    .Execute
    
End With


With qdfCadastroCustos02

    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle

    'MIDIA
    l = 47
    c = 3
    For x = 1 To 8
        .Parameters(x & "MIDIA") = Cells(l, c).value
        c = c + 1
    Next x

    'CORREIO
    l = 48
    c = 3
    For x = 1 To 8
        .Parameters(x & "CORREIO") = Cells(l, c).value
        c = c + 1
    Next x


    'ULTIMA CAPA
    l = 49
    c = 3
    For x = 1 To 8
        .Parameters(x & "ULTIMA_CAPA") = Cells(l, c).value
        c = c + 1
    Next x

    'IMPORT
    l = 50
    c = 3
    For x = 1 To 8
        .Parameters(x & "IMPORT") = Cells(l, c).value
        c = c + 1
    Next x

    'TRANSPORTE NACIONAL
    l = 51
    c = 3
    For x = 1 To 8
        .Parameters(x & "TRANSPORTE_NACIONAL") = Cells(l, c).value
        c = c + 1
    Next x

    'TRANSPORTE INTERNACIONAL
    l = 52
    c = 3
    For x = 1 To 8
        .Parameters(x & "TRANSPORTE_INTERNACIONAL") = Cells(l, c).value
        c = c + 1
    Next x

    'SEGUROS
    l = 53
    c = 3
    For x = 1 To 8
        .Parameters(x & "SEGUROS") = Cells(l, c).value
        c = c + 1
    Next x

    'EXTRAS
    l = 54
    c = 3
    For x = 1 To 8
        .Parameters(x & "EXTRAS") = Cells(l, c).value
        c = c + 1
    Next x

    'EDITOR FEE
    l = 55
    c = 3
    For x = 1 To 8
        .Parameters(x & "EDITOR_FEE") = Cells(l, c).value
        c = c + 1
    Next x

    'DESP VIAGEM
    l = 56
    c = 3
    For x = 1 To 8
        .Parameters(x & "DESP_VIAGEM") = Cells(l, c).value
        c = c + 1
    Next x

    'OUTROS
    l = 57
    c = 3
    For x = 1 To 8
        .Parameters(x & "OUTROS") = Cells(l, c).value
        c = c + 1
    Next x

    .Execute
    
End With

'With qdfCadastroCustos03
'
'    .Parameters("NOME_VENDEDOR") = strVendedor
'    .Parameters("NUMERO_CONTROLE") = strControle
'
'    'TRANSPORTE
'    L = 61
'    c = 3
'    For x = 1 To 8
'        .Parameters(x & "TRANSPORTE") = Cells(L, c).Value
'        c = c + 1
'    Next x
'
'    'IMPORT_DESEMB
'    L = 62
'    c = 3
'    For x = 1 To 8
'        .Parameters(x & "IMPORT_DESEMB") = Cells(L, c).Value
'        c = c + 1
'    Next x
'
'    .Execute
'
'End With

'qdfCadastroCustos03.Close
qdfCadastroCustos02.Close
qdfCadastroCustos01.Close
dbOrcamento.Close


CadastroOrcamentoCustos_Fim:

    Set dbOrcamento = Nothing
    Set qdfCadastroCustos01 = Nothing
    Set qdfCadastroCustos02 = Nothing
    Set qdfCadastroCustos03 = Nothing
    
    Exit Function
CadastroOrcamentoCustos_err:
    CadastroOrcamentoCustos = False
    MsgBox Err.Description
    Resume CadastroOrcamentoCustos_Fim
End Function

Public Function CadastroAnexoLinha( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)

On Error GoTo CadastroAnexoLinha_err

Dim dbOrcamento As DAO.Database
Dim qdfCadastroAnexoLinha As DAO.QueryDef

Dim x, y As Integer ' contador de linhas

y = admQuantidadeDeAnexos(BaseDeDados, strControle, strVendedor, "Linha")

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroAnexoLinha = dbOrcamento.QueryDefs("CadastroAnexoLinha")
    
    For x = 1 To y
        
        With qdfCadastroAnexoLinha
            
            .Parameters("NOME_VENDEDOR") = strVendedor
            .Parameters("NUMERO_CONTROLE") = strControle
            .Parameters("NM_LINHA") = Cells(intLinha, intColuna).value
            .Parameters("MAXIMO") = Cells(intLinha, intColuna + 1).value
            .Parameters("MINIMO") = Cells(intLinha, intColuna + 2).value
            
            .Execute
            
        
        End With
        
        intLinha = intLinha + 1
    Next x


CadastroAnexoLinha_Fim:
    qdfCadastroAnexoLinha.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set qdfCadastroAnexoLinha = Nothing
    
    Exit Function
CadastroAnexoLinha_err:
    MsgBox Err.Description
    Resume CadastroAnexoLinha_Fim


End Function

Public Function CadastroAnexoMoeda( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CadastroAnexoMoeda_err

Dim dbOrcamento As DAO.Database
Dim qdfCadastroAnexoMoeda As DAO.QueryDef

Dim x, y As Integer ' contador de linhas

y = admQuantidadeDeAnexos(BaseDeDados, strControle, strVendedor, "Moeda")

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroAnexoMoeda = dbOrcamento.QueryDefs("CadastroAnexoMoeda")

    For x = 1 To y
        
        With qdfCadastroAnexoMoeda
            
            .Parameters("NOME_VENDEDOR") = strVendedor
            .Parameters("NUMERO_CONTROLE") = strControle
            .Parameters("NM_MOEDA") = Cells(intLinha, intColuna).value
            .Parameters("INDICE") = Cells(intLinha, intColuna + 1).value
            
            .Execute
            
            
        End With
        
        intLinha = intLinha + 1
    Next x


CadastroAnexoMoeda_Fim:
    qdfCadastroAnexoMoeda.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set qdfCadastroAnexoMoeda = Nothing
    
    Exit Function
CadastroAnexoMoeda_err:
    MsgBox Err.Description
    Resume CadastroAnexoMoeda_Fim

End Function

Public Function CadastroAnexoVenda( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)

On Error GoTo CadastroAnexoVenda_err

Dim dbOrcamento As DAO.Database
Dim qdfCadastroAnexoVenda As DAO.QueryDef

Dim x, y As Integer ' contador de linhas

y = admQuantidadeDeAnexos(BaseDeDados, strControle, strVendedor, "Venda")

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroAnexoVenda = dbOrcamento.QueryDefs("CadastroAnexoVenda")

    For x = 1 To y
        
        With qdfCadastroAnexoVenda
                    
            .Parameters("NOME_VENDEDOR") = strVendedor
            .Parameters("NUMERO_CONTROLE") = strControle
            .Parameters("NM_VENDA") = Cells(intLinha, intColuna).value
            .Parameters("PORCENTAGEM") = Cells(intLinha, intColuna + 1).value
            
            .Execute
        
        End With
        
        intLinha = intLinha + 1
    Next x


CadastroAnexoVenda_Fim:
    qdfCadastroAnexoVenda.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set qdfCadastroAnexoVenda = Nothing
    
    Exit Function
CadastroAnexoVenda_err:
    MsgBox Err.Description
    Resume CadastroAnexoVenda_Fim

End Function

Public Function CadastroAnexoDesconto( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CadastroAnexoDescontos_err

Dim dbOrcamento As DAO.Database
Dim qdfCadastroAnexoDescontos As DAO.QueryDef

Dim x, y As Integer ' contador de linhas

y = admQuantidadeDeAnexos(BaseDeDados, strControle, strVendedor, "Desconto")

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroAnexoDescontos = dbOrcamento.QueryDefs("CadastroAnexoDescontos")
    
    For x = 1 To y
        
        With qdfCadastroAnexoDescontos
        
            .Parameters("NOME_VENDEDOR") = strVendedor
            .Parameters("NUMERO_CONTROLE") = strControle
            .Parameters("NM_MOTIVO") = Cells(intLinha, intColuna + 1).value
            .Parameters("VALOR01") = Val(Cells(intLinha, intColuna).value)
            
            .Execute
            
        End With
        
        intLinha = intLinha + 1
    Next x


CadastroAnexoDescontos_Fim:
    qdfCadastroAnexoDescontos.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set qdfCadastroAnexoDescontos = Nothing
    
    Exit Function
CadastroAnexoDescontos_err:
    MsgBox Err.Description
    Resume CadastroAnexoDescontos_Fim

End Function

Public Function CadastroAnexoArquivo( _
                                        BaseDeDados As String, _
                                        strControle As String, _
                                        strVendedor As String, _
                                        strArquivo As String)
                                        
On Error GoTo CadastroAnexoArquivo_err

Dim dbOrcamento As DAO.Database
Dim rstCadastroAnexoArquivo As DAO.Recordset

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCadastroAnexoArquivo = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos")

With rstCadastroAnexoArquivo

    .AddNew
    
    .Fields("CONTROLE") = strControle
    .Fields("VENDEDOR") = strVendedor
    .Fields("PROPRIEDADE") = "ARQUIVO"
    .Fields("OBS_01") = strArquivo

    .Update

End With


CadastroAnexoArquivo_Fim:
    rstCadastroAnexoArquivo.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCadastroAnexoArquivo = Nothing
    
    Exit Function
CadastroAnexoArquivo_err:
    MsgBox Err.Description
    Resume CadastroAnexoArquivo_Fim

End Function

Public Function CadastroOrcamentoProjeto( _
                                 BaseDeDados As String, _
                                 strControle As String, _
                                 strVendedor As String, _
                                 strProjeto As String)
                                 
On Error GoTo CadastroOrcamentoProjeto_err

Dim dbOrcamento As DAO.Database
Dim rstCadastroOrcamentoProjeto As DAO.Recordset


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCadastroOrcamentoProjeto = dbOrcamento.OpenRecordset("Select * from Orcamentos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")

With rstCadastroOrcamentoProjeto

    .Edit
    .Fields("PROJETO") = strProjeto
    .Update

End With


CadastroOrcamentoProjeto_Fim:
    rstCadastroOrcamentoProjeto.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCadastroOrcamentoProjeto = Nothing
    
    Exit Function
CadastroOrcamentoProjeto_err:
    MsgBox Err.Description
    Resume CadastroOrcamentoProjeto_Fim

End Function

Public Function CadastroOrcamentoAcabamento( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    strCampo As String, _
                                    strAcabamento As String)

On Error GoTo CadastroOrcamentoAcabamento_err

Dim dbOrcamento As DAO.Database
Dim rstCadastroOrcamentoAcabamento As DAO.Recordset


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCadastroOrcamentoAcabamento = dbOrcamento.OpenRecordset("Select * from Orcamentos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")

With rstCadastroOrcamentoAcabamento

    .Edit
    .Fields(strCampo) = strAcabamento
    .Update

End With



CadastroOrcamentoAcabamento_Fim:
    rstCadastroOrcamentoAcabamento.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCadastroOrcamentoAcabamento = Nothing
    
    Exit Function
CadastroOrcamentoAcabamento_err:
    MsgBox Err.Description
    Resume CadastroOrcamentoAcabamento_Fim

End Function

