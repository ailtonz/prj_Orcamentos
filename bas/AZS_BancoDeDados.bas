Attribute VB_Name = "AZS_BancoDeDados"
Option Explicit

Sub CriarBancoParaExportacao(strBancoDestino As String)
Dim oAccess As Access.Application
Dim dbDestino As dao.Database

Set oAccess = New Access.Application
Set dbDestino = DBEngine.CreateDatabase(strBancoDestino, dbLangGeneral & ";pwd=" & SenhaBanco, dbVersion40)

dbDestino.Close

Set dbDestino = Nothing
Set oAccess = Nothing

End Sub

Sub CriarTabelaEmBancoParaExportacao(strOrigem As String, strDestino As String, strTabela As String)
Dim dbORIGEM As dao.Database
Dim tbORIGEM As dao.TableDef
Dim dbDestino As dao.Database
Dim tdfNew As dao.TableDef
    
    
Set dbORIGEM = DBEngine.OpenDatabase(strOrigem, False, False, "MS Access;PWD=" & SenhaBanco)
Set tbORIGEM = dbORIGEM.TableDefs(strTabela)
Set dbDestino = DBEngine.OpenDatabase(strDestino, False, False, "MS Access;PWD=" & SenhaBanco)
Set tdfNew = dbDestino.CreateTableDef(strTabela)

Dim x As Integer
    
    For x = 0 To dbORIGEM.TableDefs(strTabela).Fields.count - 1
    
        With tdfNew
    
            .Fields.Append .CreateField(dbORIGEM.TableDefs(strTabela).Fields(x).Properties("name"), dbORIGEM.TableDefs(strTabela).Fields(x).Properties("type"), dbORIGEM.TableDefs(strTabela).Fields(x).Properties("size"))
    
        End With
    
    Next x

   dbDestino.TableDefs.Append tdfNew

'''Delete new TableDef because this is a demonstration.
'''dbDESTINO.TableDefs.Delete tdfNew.Name
   
   dbDestino.Close
   dbORIGEM.Close

End Sub

Sub ExportarDadosTabelaOrigemParaTabelaDestino(ByVal strOrigem As String, ByVal strDestino As String, ByVal strTabela As String)
''' EXPORTAR DADOS DA TABELA ORIGEM PARA A TABELA DESTINO (AMBAS COM A MESMA EXTRUTURA)
''==============================''
''           ORIGEM
''==============================''

'' POSICIONA O BANCO DE ORIGEM
Dim dbORIGEM As dao.Database
Set dbORIGEM = DBEngine.OpenDatabase(strOrigem, False, False, "MS Access;PWD=" & SenhaBanco)

'' SELECIONA A TABELA DE ORIGEM
Dim tbORIGEM As dao.TableDef
Set tbORIGEM = dbORIGEM.TableDefs(strTabela)

'' SELECIONA OS REGISTROS DA ORIGEM
Dim rstORIGEM As dao.Recordset
Set rstORIGEM = dbORIGEM.OpenRecordset("Select * from " & strTabela & "")


''==============================''
''           DESTINO
''==============================''

'' POSICIONA O BANCO DE DESTINO
Dim dbDestino As dao.Database
Set dbDestino = OpenDatabase(strDestino, False, False, "MS Access;PWD=" & SenhaBanco)

'' SELECIONA A TABELA DE DESTINO
Dim tdfNew As dao.TableDef
Set tdfNew = dbDestino.CreateTableDef(strTabela)

'' SELECIONA OS REGISTROS DA ORIGEM
Dim rstDESTINO As dao.Recordset
Set rstDESTINO = dbDestino.OpenRecordset("Select * from " & strTabela & "")

Dim x As Integer

'Saida Now(), "ExportarDadosTabelaOrigemParaTabelaDestino.log"

While Not rstORIGEM.EOF

    rstDESTINO.AddNew

    For x = 0 To dbORIGEM.TableDefs(strTabela).Fields.count - 1

        With tdfNew
             rstDESTINO.Fields(dbDestino.TableDefs(strTabela).Fields(x).Properties("name")) = rstORIGEM.Fields(dbORIGEM.TableDefs(strTabela).Fields(x).Properties("name"))
        End With

    Next x
    
    rstDESTINO.Update
    rstORIGEM.MoveNext

Wend
   
'Saida Now(), "ExportarDadosTabelaOrigemParaTabelaDestino.log"
   
rstDESTINO.Close
rstORIGEM.Close
dbDestino.Close
dbORIGEM.Close

Set rstDESTINO = Nothing
Set rstORIGEM = Nothing
Set dbDestino = Nothing
Set dbORIGEM = Nothing

End Sub


