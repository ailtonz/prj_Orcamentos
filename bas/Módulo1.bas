Attribute VB_Name = "Módulo1"
Option Explicit

Sub TESTE_UP_CLIENTES()

Dim sScript As String
Dim sValor As String: sValor = "AILTON.OK.10-04-15_19-28"
Dim sDescricao As String: sDescricao = "CADASTRO DE CLIENTE"
Dim sID As String: sID = "2"

sScript = "INSERT INTO admCategorias (admCategorias.codRelacao, admCategorias.Categoria) SELECT (SELECT admCategorias.codCategoria FROM admCategorias Where Categoria='CLIENTES' and codRelacao = 0) AS idRelacao ,'" & sValor & "',"

loadBancos
If admNovoCliente(banco(0), sID, sDescricao, sScript) Then
    MsgBox "Valor do Dolar Atualizado com sucesso.", vbInformation + vbOKOnly, "Atualização de moeda"
Else
    MsgBox "ERROR AO: Valor do Dolar Atualizado com sucesso.", vbCritical + vbOKOnly, "Atualização de moeda"
End If



End Sub
