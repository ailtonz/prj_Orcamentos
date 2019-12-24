Attribute VB_Name = "basProposta"
Sub Gerar_Proposta()



Dim wdApp As Word.Application, wdDoc As Word.Document
'Dim tdate As Date
'Dim contract As String
Dim Arquivo As String: Arquivo = "Proposta - Springer Brasil - 2014.docx"

'contract = "Verizon FIOS"
'tdate = Date

On Error Resume Next

Set wdApp = GetObject(, "Word.Application")

If Err.Number <> 0 Then 'Word isn't already running
    Set wdApp = CreateObject("Word.Application")
End If

On Error GoTo 0

Set wdDoc = wdApp.Documents.Open(ActiveWorkbook.Path & "\db\" & Arquivo, ReadOnly:=True)

wdApp.Visible = True
'Writing Variables from Excel to the Checklist word doc.

'' PROPOSTA
wdDoc.Bookmarks("N_CONTROLE").Range.Text = ActiveSheet.Name
wdDoc.Bookmarks("CLIENTE").Range.Text = Range("C4").value
wdDoc.Bookmarks("RESPONSAVEL").Range.Text = Range("C5").value
wdDoc.Bookmarks("PROJETO").Range.Text = Range("C6").value
wdDoc.Bookmarks("JOURNAL").Range.Text = Range("C9").value
wdDoc.Bookmarks("AUTOR").Range.Text = Range("C10").value
wdDoc.Bookmarks("PUBLISHER").Range.Text = Range("C8").value

'' PROJETO
wdDoc.Bookmarks("FORMATO").Range.Text = Range("C29").value
wdDoc.Bookmarks("N_PAGINAS").Range.Text = Range("C27").value
wdDoc.Bookmarks("IDIOMA").Range.Text = Range("C17").value
wdDoc.Bookmarks("VOLUME").Range.Text = Range("").value
wdDoc.Bookmarks("PRC_VENDA").Range.Text = Range("").value
wdDoc.Bookmarks("PRC_TOTAL").Range.Text = Range("").value

'' GERENTE DE CONTAS
wdDoc.Bookmarks("G_CONTAS").Range.Text = Range("C3").value
wdDoc.Bookmarks("TELEFONE").Range.Text = Range("I2").value
wdDoc.Bookmarks("CELULAR_01").Range.Text = Range("I2").value
wdDoc.Bookmarks("CELULAR_02").Range.Text = Range("I2").value
wdDoc.Bookmarks("ID_NEXTEL").Range.Text = Range("I2").value

wdDoc.SaveAs pathDesktopAddress & "\" & Now() & "_" & Arquivo
wdDoc.Close

wdApp.Application.Quit
End Sub


Sub UpdateBookmark(BookmarkToUpdate As String, TextToUse As String)
    Dim BMRange As Range
    Set BMRange = ActiveDocument.Bookmarks(BookmarkToUpdate).Range
    BMRange.Text = TextToUse
    ActiveDocument.Bookmarks.add BookmarkToUpdate, BMRange
End Sub
