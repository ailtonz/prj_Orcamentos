Attribute VB_Name = "mod_02_Cadastrar"
Option Base 1
Option Explicit

Public Function CadastroOrcamento( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String) As Boolean: CadastroOrcamento = True
On Error GoTo CadastroOrcamento_err

Dim dbOrcamento As dao.Database
Dim qdfCadastroOrcamento As dao.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroOrcamento = dbOrcamento.QueryDefs("CadastroOrcamento")

    With qdfCadastroOrcamento
    
        .Parameters("NOME_VENDEDOR") = strVendedor
        .Parameters("NUMERO_CONTROLE") = strControle
        
        .Parameters("NM_CLIENTE") = Range("C4").Value
        .Parameters("NM_RESPONSAVEL") = Range("C5").Value
        .Parameters("NM_LINHA_PRODUTO") = Range("G5").Value
        .Parameters("DTPEDIDO") = Range("G3").Value
        .Parameters("PREVENTREGA") = Range("G4").Value
        .Parameters("VALORPROJETO") = Range("J4").Value
        .Parameters("NM_PUBLISHER") = Range("C8").Value
        .Parameters("NM_JOURNAL") = Range("C9").Value
        .Parameters("NM_PAGS") = Range("C10").Value
        
        CadastroOrcamentoProjeto BaseDeDados, strControle, strVendedor, Range("C6").Value
    
        'FECHADO COM CLIENTE
        l = 12
        c = 3
        For x = 1 To 8
            .Parameters(x & "FECHADO") = Cells(l, c).Value
            c = c + 1
        Next x
    
        'VENDA
        l = 13
        c = 3
        For x = 1 To 8
            .Parameters(x & "VENDA") = Cells(l, c).Value
            c = c + 1
        Next x
    
        'IMPOSTO
        l = 14
        c = 3
        For x = 1 To 8
            .Parameters(x & "IMPOSTO") = Cells(l, c).Value
            c = c + 1
        Next x
    
        'IDIOMA
        l = 15
        c = 3
        For x = 1 To 8
            .Parameters(x & "IDIOMA") = Cells(l, c).Value
            c = c + 1
        Next x
    
        'TIRAGEM
        l = 16
        c = 3
        For x = 1 To 8
            .Parameters(x & "TIRAGEM") = Cells(l, c).Value
            c = c + 1
        Next x
    
        'ESPECIFICACAO
        l = 17
        c = 3
        For x = 1 To 8
            .Parameters(x & "ESPECIFICACAO") = Cells(l, c).Value
            c = c + 1
        Next x
        
        'MOEDA
        l = 18
        c = 3
        For x = 1 To 8
            .Parameters(x & "MOEDA") = Cells(l, c).Value
            c = c + 1
        Next x
        
        'ROYALTY PERCENTUAL
        l = 19
        c = 3
        For x = 1 To 8
            .Parameters(x & "ROYALTY_PERCENTUAL") = Cells(l, c).Value
            c = c + 1
        Next x
    
        'ROYALTY ESPECIE
        l = 20
        c = 3
        For x = 1 To 8
            .Parameters(x & "ROYALTY_ESPECIE") = Cells(l, c).Value
            c = c + 1
        Next x
        
        'RE IMPRESSAO
        l = 21
        c = 3
        For x = 1 To 8
            .Parameters(x & "RE_IMPRESSAO") = Cells(l, c).Value
            c = c + 1
        Next x
    
        'DESCONTO - ( PREÇOS )
        l = 60
        c = 3
        For x = 1 To 8
            .Parameters(x & "DESCONTO") = Cells(l, c).Value
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

Dim dbOrcamento As dao.Database
Dim qdfCadastroOrcamentoImpressao As dao.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroOrcamentoImpressao = dbOrcamento.QueryDefs("CadastroOrcamentoImpressao")

With qdfCadastroOrcamentoImpressao

    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle

    'TIPO
    l = 23
    c = 3
    For x = 1 To 4
        .Parameters(x & "TIPO") = Cells(l, c).Value
        c = c + 2
    Next x

    'PAPEL
    l = 24
    c = 3
    For x = 1 To 4
        .Parameters(x & "PAPEL") = Cells(l, c).Value
        c = c + 2
    Next x

    'PAGINAS
    l = 25
    c = 3
    For x = 1 To 4
        .Parameters(x & "PAGINAS") = Cells(l, c).Value
        c = c + 2
    Next x

    'IMPRESSAO
    l = 26
    c = 3
    For x = 1 To 4
        .Parameters(x & "IMPRESSAO") = Cells(l, c).Value
        c = c + 2
    Next x
    
    'FORMATO
    l = 27
    c = 3
    For x = 1 To 4
        .Parameters(x & "FORMATO") = Cells(l, c).Value
        c = c + 2
    Next x

    'ACABAMENTO
    l = 29
    c = 2
    For x = 1 To 4
        CadastroOrcamentoAcabamento BaseDeDados, strControle, strVendedor, x & "_ACABAMENTO", Cells(l, c).Value
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

Dim dbOrcamento As dao.Database
Dim qdfCadastroCustos01 As dao.QueryDef
Dim qdfCadastroCustos02 As dao.QueryDef
Dim strSQL As String

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroCustos01 = dbOrcamento.QueryDefs("CadastroOrcamentoCustos01")
Set qdfCadastroCustos02 = dbOrcamento.QueryDefs("CadastroOrcamentoCustos02")

With qdfCadastroCustos01

    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle

    'INDEXACAO
    l = 35
    c = 3
    For x = 1 To 8
        .Parameters(x & "INDEXACAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'TRADUCAO
    l = 36
    c = 3
    For x = 1 To 8
        .Parameters(x & "TRADUCAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'REVISAO ORTOGRAFICA
    l = 37
    c = 3
    For x = 1 To 8
        .Parameters(x & "REVISAO_ORTOGRAFICA") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'REVISAO MEDICA
    l = 38
    c = 3
    For x = 1 To 8
        .Parameters(x & "REVISAO_MEDICA") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'CRIACAO
    l = 39
    c = 3
    For x = 1 To 8
        .Parameters(x & "CRIACAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'ILUSTRACAO
    l = 40
    c = 3
    For x = 1 To 8
        .Parameters(x & "ILUSTRACAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'REVISAO
    l = 41
    c = 3
    For x = 1 To 8
        .Parameters(x & "REVISAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'DIAGRAMACAO
    l = 42
    c = 3
    For x = 1 To 8
        .Parameters(x & "DIAGRAMACAO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'MEDICO
    l = 43
    c = 3
    For x = 1 To 8
        .Parameters(x & "MEDICO") = Cells(l, c).Value
        c = c + 1
    Next x
    
    'GRAFICA
    l = 44
    c = 3
    For x = 1 To 8
        .Parameters(x & "GRAFICA") = Cells(l, c).Value
        c = c + 1
    Next x
    
    .Execute
    
End With


With qdfCadastroCustos02

    .Parameters("NOME_VENDEDOR") = strVendedor
    .Parameters("NUMERO_CONTROLE") = strControle

    'MIDIA
    l = 45
    c = 3
    For x = 1 To 8
        .Parameters(x & "MIDIA") = Cells(l, c).Value
        c = c + 1
    Next x

    'CORREIO
    l = 46
    c = 3
    For x = 1 To 8
        .Parameters(x & "CORREIO") = Cells(l, c).Value
        c = c + 1
    Next x


    'ULTIMA CAPA
    l = 47
    c = 3
    For x = 1 To 8
        .Parameters(x & "ULTIMA_CAPA") = Cells(l, c).Value
        c = c + 1
    Next x

    'IMPORT
    l = 48
    c = 3
    For x = 1 To 8
        .Parameters(x & "IMPORT") = Cells(l, c).Value
        c = c + 1
    Next x

    'TRANSPORTE NACIONAL
    l = 49
    c = 3
    For x = 1 To 8
        .Parameters(x & "TRANSPORTE_NACIONAL") = Cells(l, c).Value
        c = c + 1
    Next x

    'TRANSPORTE_INTERNACIONAL
    l = 50
    c = 3
    For x = 1 To 8
        .Parameters(x & "TRANSPORTE_INTERNACIONAL") = Cells(l, c).Value
        c = c + 1
    Next x

    'SEGUROS
    l = 51
    c = 3
    For x = 1 To 8
        .Parameters(x & "SEGUROS") = Cells(l, c).Value
        c = c + 1
    Next x

    'EXTRAS
    l = 52
    c = 3
    For x = 1 To 8
        .Parameters(x & "EXTRAS") = Cells(l, c).Value
        c = c + 1
    Next x

    'EDITOR_FEE
    l = 53
    c = 3
    For x = 1 To 8
        .Parameters(x & "EDITOR_FEE") = Cells(l, c).Value
        c = c + 1
    Next x

    'DESP_VIAGEM
    l = 54
    c = 3
    For x = 1 To 8
        .Parameters(x & "DESP_VIAGEM") = Cells(l, c).Value
        c = c + 1
    Next x

    'OUTROS
    l = 55
    c = 3
    For x = 1 To 8
        .Parameters(x & "OUTROS") = Cells(l, c).Value
        c = c + 1
    Next x

    .Execute
    
End With


qdfCadastroCustos02.Close
qdfCadastroCustos01.Close
dbOrcamento.Close


CadastroOrcamentoCustos_Fim:

    Set dbOrcamento = Nothing
    Set qdfCadastroCustos01 = Nothing
    Set qdfCadastroCustos02 = Nothing
    
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

Dim dbOrcamento As dao.Database
Dim qdfCadastroAnexoLinha As dao.QueryDef

Dim x, y As Integer ' contador de linhas

y = admQuantidadeDeAnexos(BaseDeDados, strControle, strVendedor, "Linha")

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroAnexoLinha = dbOrcamento.QueryDefs("CadastroAnexoLinha")
    
    For x = 1 To y
        
        With qdfCadastroAnexoLinha
            
            .Parameters("NOME_VENDEDOR") = strVendedor
            .Parameters("NUMERO_CONTROLE") = strControle
            .Parameters("NM_LINHA") = Cells(intLinha, intColuna).Value
            .Parameters("MAXIMO") = Cells(intLinha, intColuna + 1).Value
            .Parameters("MINIMO") = Cells(intLinha, intColuna + 2).Value
            
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

Dim dbOrcamento As dao.Database
Dim qdfCadastroAnexoMoeda As dao.QueryDef

Dim x, y As Integer ' contador de linhas

y = admQuantidadeDeAnexos(BaseDeDados, strControle, strVendedor, "Moeda")

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroAnexoMoeda = dbOrcamento.QueryDefs("CadastroAnexoMoeda")

    For x = 1 To y
        
        With qdfCadastroAnexoMoeda
            
            .Parameters("NOME_VENDEDOR") = strVendedor
            .Parameters("NUMERO_CONTROLE") = strControle
            .Parameters("NM_MOEDA") = Cells(intLinha, intColuna).Value
            .Parameters("INDICE") = Cells(intLinha, intColuna + 1).Value
            
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

Dim dbOrcamento As dao.Database
Dim qdfCadastroAnexoVenda As dao.QueryDef

Dim x, y As Integer ' contador de linhas

y = admQuantidadeDeAnexos(BaseDeDados, strControle, strVendedor, "Venda")

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroAnexoVenda = dbOrcamento.QueryDefs("CadastroAnexoVenda")

    For x = 1 To y
        
        With qdfCadastroAnexoVenda
                    
            .Parameters("NOME_VENDEDOR") = strVendedor
            .Parameters("NUMERO_CONTROLE") = strControle
            .Parameters("NM_VENDA") = Cells(intLinha, intColuna).Value
            .Parameters("PORCENTAGEM") = Cells(intLinha, intColuna + 1).Value
            
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

Dim dbOrcamento As dao.Database
Dim qdfCadastroAnexoDescontos As dao.QueryDef

Dim x, y As Integer ' contador de linhas

y = admQuantidadeDeAnexos(BaseDeDados, strControle, strVendedor, "Desconto")

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set qdfCadastroAnexoDescontos = dbOrcamento.QueryDefs("CadastroAnexoDescontos")
    
    For x = 1 To y
        
        With qdfCadastroAnexoDescontos
        
            .Parameters("NOME_VENDEDOR") = strVendedor
            .Parameters("NUMERO_CONTROLE") = strControle
            .Parameters("NM_MOTIVO") = Cells(intLinha, intColuna + 1).Value
            .Parameters("VALOR01") = Val(Cells(intLinha, intColuna).Value)
            
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

Dim dbOrcamento As dao.Database
Dim rstCadastroAnexoArquivo As dao.Recordset

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

Dim dbOrcamento As dao.Database
Dim rstCadastroOrcamentoProjeto As dao.Recordset


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

Dim dbOrcamento As dao.Database
Dim rstCadastroOrcamentoAcabamento As dao.Recordset


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

