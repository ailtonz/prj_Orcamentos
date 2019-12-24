Attribute VB_Name = "mod_01_Carregar"
Option Base 1
Option Explicit

Public Function CarregarOrcamento( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String)
On Error GoTo CarregarOrcamento_err

Dim dbOrcamento As dao.Database
Dim rstCarregarOrcamento As dao.Recordset
Dim rstCarregarCustos As dao.Recordset

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarOrcamento = dbOrcamento.OpenRecordset("Select * from Orcamentos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")
Set rstCarregarCustos = dbOrcamento.OpenRecordset("Select * from OrcamentosCustos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "'")

    Application.ScreenUpdating = False
    DesbloqueioDeGuia SenhaBloqueio
    
    '#########################
    '   ORÇAMENTO
    '#########################
    
    With rstCarregarOrcamento
    
        Range("C3").Value = .Fields("VENDEDOR")
        Range("C4").Value = .Fields("CLIENTE")
        Range("C5").Value = .Fields("RESPONSAVEL")
        Range("C6").Value = .Fields("PROJETO")
        Range("G5").Value = .Fields("LINHA_PRODUTO")
        Range("G3").Value = .Fields("DT_PEDIDO")
        Range("G4").Value = .Fields("PREV_ENTREGA")
        Range("J4").Value = .Fields("VALOR_PROJETO")
        Range("J3").Value = .Fields("STATUS")
        Range("C8").Value = .Fields("PUBLISHER")
        Range("C9").Value = .Fields("JOURNAL")
        Range("C10").Value = .Fields("PAGS")
    
        'FECHADO COM CLIENTE
        l = 12
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_FECHADO")
            c = c + 1
        Next x
    
        'VENDA
        l = 13
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_VENDA")
            c = c + 1
        Next x
    
        'IDIOMA
        l = 15
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_IDIOMA")
            c = c + 1
        Next x
    
        'TIRAGEM
        l = 16
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_TIRAGEM")
            c = c + 1
        Next x
    
        'ESPECIFICACAO
        l = 17
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_ESPECIFICACAO")
            c = c + 1
        Next x
    
        'MOEDA
        l = 18
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_MOEDA")
            c = c + 1
        Next x
    
        'ROYALTY PERCENTUAL
        l = 19
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_ROYALTY_PERCENTUAL")
            c = c + 1
        Next x
    
        'ROYALTY ESPECIE
        l = 20
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_ROYALTY_ESPECIE")
            c = c + 1
        Next x
    
        'RE IMPRESSAO
        l = 21
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_RE_IMPRESSAO")
            c = c + 1
        Next x
    
        'DESCONTO
        l = 60
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_DESCONTO")
            c = c + 1
        Next x
    
        'TIPO
        l = 23
        c = 3
        For x = 1 To 4
            Cells(l, c).Value = .Fields(x & "_TIPO")
            c = c + 2
        Next x
    
        'PAPEL
        l = 24
        c = 3
        For x = 1 To 4
            Cells(l, c).Value = .Fields(x & "_PAPEL")
            c = c + 2
        Next x
    
        'PAGINAS
        l = 25
        c = 3
        For x = 1 To 4
            Cells(l, c).Value = .Fields(x & "_PAGINAS")
            c = c + 2
        Next x
    
        'IMPRESSAO
        l = 26
        c = 3
        For x = 1 To 4
            Cells(l, c).Value = .Fields(x & "_IMPRESSAO")
            c = c + 2
        Next x
        
        'FORMATO
        l = 27
        c = 3
        For x = 1 To 4
            Cells(l, c).Value = .Fields(x & "_FORMATO")
            c = c + 2
        Next x
    
        'ACABAMENTO
        l = 29
        c = 2
        For x = 1 To 4
            Cells(l, c).Value = .Fields(x & "_ACABAMENTO")
            l = l + 1
        Next x
    
    End With
    
    '#########################
    '   CUSTOS
    '#########################
    
    With rstCarregarCustos
    
        'INDEXACAO
        l = 35
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_INDEXACAO")
            c = c + 1
        Next x
        
        'TRADUCAO
        l = 36
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_TRADUCAO")
            c = c + 1
        Next x
        
        'REVISAO ORTOGRAFICA
        l = 37
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_REVISAO_ORTOGRAFICA")
            c = c + 1
        Next x
        
        'REVISAO MEDICA
        l = 38
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_REVISAO_MEDICA")
            c = c + 1
        Next x
        
        'CRIACAO
        l = 39
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_CRIACAO")
            c = c + 1
        Next x
        
        'ILUSTRACAO
        l = 40
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_ILUSTRACAO")
            c = c + 1
        Next x
        
        'REVISAO
        l = 41
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_REVISAO")
            c = c + 1
        Next x
        
        'DIAGRAMACAO
        l = 42
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_DIAGRAMACAO")
            c = c + 1
        Next x
        
        'MEDICO
        l = 43
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_MEDICO")
            c = c + 1
        Next x
        
        'GRAFICA
        l = 44
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_GRAFICA")
            c = c + 1
        Next x
    
    
        'MIDIA
        l = 45
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_MIDIA")
            c = c + 1
        Next x
    
        'CORREIO
        l = 46
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_CORREIO")
            c = c + 1
        Next x
    
    
        'ULTIMA CAPA
        l = 47
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_ULTIMA_CAPA")
            c = c + 1
        Next x
    
        'IMPORT
        l = 48
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_IMPORT")
            c = c + 1
        Next x
    
        'TRANSPORTE NACIONAL
        l = 49
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_TRANSPORTE_NACIONAL")
            c = c + 1
        Next x
    
        'TRANSPORTE_INTERNACIONAL
        l = 50
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_TRANSPORTE_INTERNACIONAL")
            c = c + 1
        Next x
    
        'SEGUROS
        l = 51
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_SEGUROS")
            c = c + 1
        Next x
    
        'EXTRAS
        l = 52
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_EXTRAS")
            c = c + 1
        Next x
    
        'EDITOR_FEE
        l = 53
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_EDITOR_FEE")
            c = c + 1
        Next x
    
        'DESP_VIAGEM
        l = 54
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_DESP_VIAGEM")
            c = c + 1
        Next x
    
        'OUTROS
        l = 55
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_OUTROS")
            c = c + 1
        Next x
        
        'TRANSPORTE
        l = 61
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_TRANSPORTE")
            c = c + 1
        Next x


        'IMPORT
        l = 62
        c = 3
        For x = 1 To 8
            Cells(l, c).Value = .Fields(x & "_IMPORT_DESEMB")
            c = c + 1
        Next x
        
        
    
    End With
    

    CarregarAnexoDesconto BaseDeDados, strControle, strVendedor, 3, 22
    CarregarAnexoLinha BaseDeDados, strControle, strVendedor, 3, 12
    CarregarAnexoMoeda BaseDeDados, strControle, strVendedor, 3, 16
    CarregarAnexoVenda BaseDeDados, strControle, strVendedor, 3, 19
'    CarregarAnexoArquivo BaseDeDados, strControle, strVendedor, CInt(Right(ArquivoInicio, Len(ArquivoInicio) - 1)), 2
    
    BloqueioDeGuia SenhaBloqueio
    Application.ScreenUpdating = True


CarregarOrcamento_Fim:
    rstCarregarOrcamento.Close
    rstCarregarCustos.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarCustos = Nothing
    Set rstCarregarOrcamento = Nothing
    
    Exit Function
CarregarOrcamento_err:
    MsgBox Err.Description
    Resume CarregarOrcamento_Fim


End Function

Public Function CarregarAnexoDesconto( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoDesconto_err
                                    
Dim dbOrcamento As dao.Database
Dim rstCarregarAnexoDesconto As dao.Recordset

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x, y As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoDesconto = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'DESCONTO'")
    
rstCarregarAnexoDesconto.MoveLast
rstCarregarAnexoDesconto.MoveFirst
y = rstCarregarAnexoDesconto.RecordCount
    
'    l = 3
'    c = 22
    For x = 1 To y
        
        With rstCarregarAnexoDesconto
        
            Cells(intLinha, intColuna + 1).Value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna).Value = Val(.Fields("VALOR_01"))
            rstCarregarAnexoDesconto.MoveNext
            
        End With
        
        intLinha = intLinha + 1
    Next x


CarregarAnexoDesconto_Fim:
    rstCarregarAnexoDesconto.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoDesconto = Nothing
    
    Exit Function
CarregarAnexoDesconto_err:
    MsgBox Err.Description, , "Anexo Desconto"
    Resume CarregarAnexoDesconto_Fim

End Function


Public Function CarregarAnexoLinha( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoLinha_err
                                    
Dim dbOrcamento As dao.Database
Dim rstCarregarAnexoLinha As dao.Recordset

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x, y As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoLinha = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'LINHA'")

rstCarregarAnexoLinha.MoveLast
rstCarregarAnexoLinha.MoveFirst
y = rstCarregarAnexoLinha.RecordCount
    
'    l = 3
'    c = 12
    For x = 1 To y
        
        With rstCarregarAnexoLinha
            
            Cells(intLinha, intColuna).Value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna + 1).Value = .Fields("VALOR_01")
            Cells(intLinha, intColuna + 2).Value = .Fields("VALOR_02")
            rstCarregarAnexoLinha.MoveNext
            
        End With
        
        intLinha = intLinha + 1
    Next x


CarregarAnexoLinha_Fim:
    rstCarregarAnexoLinha.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoLinha = Nothing
    
    Exit Function
CarregarAnexoLinha_err:
    MsgBox Err.Description, , "Anexo Linha"
    Resume CarregarAnexoLinha_Fim

End Function


Public Function CarregarAnexoMoeda( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoMoeda_err
                                    
Dim dbOrcamento As dao.Database
Dim rstCarregarAnexoMoeda As dao.Recordset

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x, y As Integer ' contador de linhas


Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoMoeda = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'MOEDA'")
    
rstCarregarAnexoMoeda.MoveLast
rstCarregarAnexoMoeda.MoveFirst
y = rstCarregarAnexoMoeda.RecordCount
    
'    l = 3
'    c = 16
    For x = 1 To y
        
        With rstCarregarAnexoMoeda
        
            Cells(intLinha, intColuna).Value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna + 1).Value = .Fields("VALOR_01")
            rstCarregarAnexoMoeda.MoveNext
            
        End With
        
        intLinha = intLinha + 1
    Next x


CarregarAnexoMoeda_Fim:
    rstCarregarAnexoMoeda.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoMoeda = Nothing
    
    Exit Function
CarregarAnexoMoeda_err:
    MsgBox Err.Description, , "Anexo Moeda"
    Resume CarregarAnexoMoeda_Fim

End Function

Public Function CarregarAnexoVenda( _
                                    BaseDeDados As String, _
                                    strControle As String, _
                                    strVendedor As String, _
                                    intLinha As Integer, _
                                    intColuna As Integer)
                                    
On Error GoTo CarregarAnexoVenda_err
                                    
Dim dbOrcamento As dao.Database
Dim rstCarregarAnexoVenda As dao.Recordset

Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
Dim x, y As Integer ' contador de linhas

Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
Set rstCarregarAnexoVenda = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'VENDA'")

rstCarregarAnexoVenda.MoveLast
rstCarregarAnexoVenda.MoveFirst
y = rstCarregarAnexoVenda.RecordCount
    
'    l = 3
'    c = 19
    For x = 1 To y
        
        With rstCarregarAnexoVenda
        
            Cells(intLinha, intColuna).Value = .Fields("DESCRICAO")
            Cells(intLinha, intColuna + 1).Value = .Fields("VALOR_01")
            rstCarregarAnexoVenda.MoveNext
            
        End With
        
        intLinha = intLinha + 1
    Next x


CarregarAnexoVenda_Fim:
    rstCarregarAnexoVenda.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarAnexoVenda = Nothing
    
    Exit Function
CarregarAnexoVenda_err:
    MsgBox Err.Description, , "Anexo Venda"
    Resume CarregarAnexoVenda_Fim

End Function


'Public Function CarregarAnexoArquivo( _
'                                    BaseDeDados As String, _
'                                    strControle As String, _
'                                    strVendedor As String, _
'                                    intLinha As Integer, _
'                                    intColuna As Integer)
'
'On Error GoTo CarregarAnexoArquivo_err
'
'Dim dbOrcamento As dao.Database
'Dim rstCarregarAnexoArquivo As dao.Recordset
'
'Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
'Dim x, y As Integer ' contador de linhas
'
''ARQUIVOS - ( ANEXOS )
'Dim Terminio As Integer
'Dim Inicio As Integer
'
'Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
'Set rstCarregarAnexoArquivo = dbOrcamento.OpenRecordset("Select * from OrcamentosAnexos where controle = '" & strControle & "' and Vendedor = '" & strVendedor & "' and PROPRIEDADE = 'Arquivo'")
'
'If Not rstCarregarAnexoArquivo.EOF Then
'
'    Inicio = CInt(Right(ArquivoInicio, Len(ArquivoInicio) - 1))
'
'    rstCarregarAnexoArquivo.MoveLast
'    rstCarregarAnexoArquivo.MoveFirst
'    Terminio = (rstCarregarAnexoArquivo.RecordCount + Inicio) - 1
'
'    Range(ArquivoControle).Value = rstCarregarAnexoArquivo.RecordCount + Inicio
'
'    For x = Inicio To Terminio
'
'        With rstCarregarAnexoArquivo
'
'            Range(Chr(Asc(Left(ArquivoInicio, 1)) + 1) & x).Select
''            ActiveCell.FormulaR1C1 = vrtSelectedItem
'            Selection.Hyperlinks.Add Range(Chr(Asc(Left(ArquivoInicio, 1)) + 1) & x), "file://" & .Fields("OBS_01")
'            Selection.Font.Size = 12
'
''            Cells(intLinha, intColuna).Value = .Fields("OBS_01")
'            rstCarregarAnexoArquivo.MoveNext
'
'        End With
'
'        intLinha = intLinha + 1
'    Next x
'
'End If
'
'CarregarAnexoArquivo_Fim:
'    rstCarregarAnexoArquivo.Close
'    dbOrcamento.Close
'
'    Set dbOrcamento = Nothing
'    Set rstCarregarAnexoArquivo = Nothing
'
'    Exit Function
'CarregarAnexoArquivo_err:
'    MsgBox Err.Description, , "Anexo Arquivo"
'    Resume CarregarAnexoArquivo_Fim
'
'End Function
'
'
'
