Attribute VB_Name = "AZS_BancoDeDados"
Option Explicit

Sub CriarBancoParaExportacao(strBancoDestino As String)
Dim oAccess As Access.Application
Dim dbDestino As DAO.Database

Set oAccess = New Access.Application
Set dbDestino = DBEngine.CreateDatabase(strBancoDestino, dbLangGeneral & ";pwd=" & SenhaBanco, dbVersion40)

dbDestino.Close

Set dbDestino = Nothing
Set oAccess = Nothing

End Sub

Sub CriarTabelaEmBancoParaExportacao(strOrigem As String, strDestino As String, strTabela As String)
Dim dbOrigem As DAO.Database
Dim tbORIGEM As DAO.TableDef
Dim dbDestino As DAO.Database
Dim tdfNew As DAO.TableDef
    
    
Set dbOrigem = DBEngine.OpenDatabase(strOrigem, False, False, "MS Access;PWD=" & SenhaBanco)
Set tbORIGEM = dbOrigem.TableDefs(strTabela)
Set dbDestino = DBEngine.OpenDatabase(strDestino, False, False, "MS Access;PWD=" & SenhaBanco)
Set tdfNew = dbDestino.CreateTableDef(strTabela)

Dim x As Integer
    
    For x = 0 To dbOrigem.TableDefs(strTabela).Fields.count - 1
    
        With tdfNew
    
            .Fields.Append .CreateField(dbOrigem.TableDefs(strTabela).Fields(x).Properties("name"), dbOrigem.TableDefs(strTabela).Fields(x).Properties("type"), dbOrigem.TableDefs(strTabela).Fields(x).Properties("size"))
    
        End With
    
    Next x

   dbDestino.TableDefs.Append tdfNew

'''Delete new TableDef because this is a demonstration.
'''dbDESTINO.TableDefs.Delete tdfNew.Name
   
   dbDestino.Close
   dbOrigem.Close

End Sub

Sub ExportarDadosTabelaOrigemParaTabelaDestino(ByVal strOrigem As String, ByVal strDestino As String, ByVal strTabela As String)
''' EXPORTAR DADOS DA TABELA ORIGEM PARA A TABELA DESTINO (AMBAS COM A MESMA EXTRUTURA)
''==============================''
''           ORIGEM
''==============================''

'' POSICIONA O BANCO DE ORIGEM
Dim dbOrigem As DAO.Database
Set dbOrigem = DBEngine.OpenDatabase(strOrigem, False, False, "MS Access;PWD=" & SenhaBanco)

'' SELECIONA A TABELA DE ORIGEM
Dim tbORIGEM As DAO.TableDef
Set tbORIGEM = dbOrigem.TableDefs(strTabela)

'' SELECIONA OS REGISTROS DA ORIGEM
Dim rstOrigem As DAO.Recordset
Set rstOrigem = dbOrigem.OpenRecordset("Select * from " & strTabela & "")


''==============================''
''           DESTINO
''==============================''

'' POSICIONA O BANCO DE DESTINO
Dim dbDestino As DAO.Database
Set dbDestino = OpenDatabase(strDestino, False, False, "MS Access;PWD=" & SenhaBanco)

'' SELECIONA A TABELA DE DESTINO
Dim tdfNew As DAO.TableDef
Set tdfNew = dbDestino.CreateTableDef(strTabela)

'' SELECIONA OS REGISTROS DA ORIGEM
Dim rstDestino As DAO.Recordset
Set rstDestino = dbDestino.OpenRecordset("Select * from " & strTabela & "")

Dim x As Integer

'Saida Now(), "ExportarDadosTabelaOrigemParaTabelaDestino.log"

While Not rstOrigem.EOF

    rstDestino.AddNew

    For x = 0 To dbOrigem.TableDefs(strTabela).Fields.count - 1

        With tdfNew
             rstDestino.Fields(dbDestino.TableDefs(strTabela).Fields(x).Properties("name")) = rstOrigem.Fields(dbOrigem.TableDefs(strTabela).Fields(x).Properties("name"))
        End With

    Next x
    
    rstDestino.Update
    rstOrigem.MoveNext

Wend
   
'Saida Now(), "ExportarDadosTabelaOrigemParaTabelaDestino.log"
   
rstDestino.Close
rstOrigem.Close
dbDestino.Close
dbOrigem.Close

Set rstDestino = Nothing
Set rstOrigem = Nothing
Set dbDestino = Nothing
Set dbOrigem = Nothing

End Sub


