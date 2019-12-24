VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPojeto 
   Caption         =   "Projeto"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   OleObjectBlob   =   "frmPojeto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPojeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCadastrar_Click()
Dim strMSG As String: strMSG = "Favor Preencher campo "
Dim strTitulo As String: strTitulo = "CAMPO OBRIGATORIO!"

''' VENDAS
If Me.cboVendas <> "" Then
    Range(ProjetoAtual & "13") = Me.cboVendas
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboVendas.SetFocus
    Exit Sub
End If

''' IDIOMAS
If Me.cboIdiomas <> "" Then
    Range(ProjetoAtual & "15") = Me.cboIdiomas
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboIdiomas.SetFocus
    Exit Sub
End If

''' TIRAGEM
If Me.txtTiragem <> "" Then
    Range(ProjetoAtual & "16") = Me.txtTiragem
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtTiragem.SetFocus
    Exit Sub
End If

''' ESPECIFICAÇÃO
If Me.txtEspecificacao <> "" Then
    Range(ProjetoAtual & "17") = Me.txtEspecificacao
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtEspecificacao.SetFocus
    Exit Sub
End If

''' MOEDA
If Me.cboMoeda <> "" Then
    Range(ProjetoAtual & "18") = Me.cboMoeda
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboMoeda.SetFocus
    Exit Sub
End If

''' ROYALTY (%)
If Me.txtRoyalty_Percentual <> "" Then
    Range(ProjetoAtual & "19") = Me.txtRoyalty_Percentual
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtRoyalty_Percentual.SetFocus
    Exit Sub
End If

''' ROYALTY (VALOR)
If Me.txtRoyalty_Valor <> "" Then
    Range(ProjetoAtual & "20") = Me.txtRoyalty_Valor
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtRoyalty_Valor.SetFocus
    Exit Sub
End If

''' RE-IMPRESSÃO
If Me.txtReImpressao <> "" Then
    Range(ProjetoAtual & "21") = Me.txtReImpressao
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtReImpressao.SetFocus
    Exit Sub
End If

Call cmdFechar_Click

End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
Dim cPart As Range
Dim cLoc As Range
Dim strUsuario As String: strUsuario = Range(NomeUsuario)

Dim wsApoio As Worksheet
Set wsApoio = Worksheets("Apoio")

Dim wsPrincipal As Worksheet
Set wsPrincipal = Worksheets(ActiveSheet.Name)

    
    If ProjetoAtual = "" Then ProjetoAtual = "C"
    
    ''' CARREGAR COMBO BOX DE VENDAS
    ComboBoxUpdate wsPrincipal.Name, "VENDAS", Me.cboVendas
    
    Me.cboVendas = Range(ProjetoAtual & "13")
    
    ''' CARREGAR COMBO BOX DE IDIOMAS
    ComboBoxUpdate wsApoio.Name, "IDIOMAS", Me.cboIdiomas
    
    Me.cboIdiomas = Range(ProjetoAtual & "15")
    Me.txtTiragem = Range(ProjetoAtual & "16")
    Me.txtEspecificacao = Range(ProjetoAtual & "17")
    Me.cboVendas.SetFocus
    
    ''' CARREGAR COMBO BOX DE MOEDA
    ComboBoxUpdate wsPrincipal.Name, "MOEDA", Me.cboMoeda
    Me.cboMoeda = Range(ProjetoAtual & "18")
    
    Me.txtRoyalty_Percentual = Range(ProjetoAtual & "19")
    Me.txtRoyalty_Valor = Range(ProjetoAtual & "20")
    Me.txtReImpressao = Range(ProjetoAtual & "21")

End Sub

