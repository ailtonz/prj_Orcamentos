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
On Error GoTo cmdCadastrar_err
Dim strMSG As String: strMSG = "Favor Preencher campo "
Dim strTitulo As String: strTitulo = "CAMPO OBRIGATORIO!"

Dim strBanco As String: strBanco = Range(BancoLocal)
Dim strSheet As String: strSheet = ActiveSheet.Name
Dim strGerente As String: strGerente = Range(GerenteDeContas)
Dim Cadastro As Boolean: Cadastro = False


''' VENDAS
If Me.cboVendas <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboVendas.SetFocus
    Exit Sub
End If

''' IDIOMAS
If Me.cboIdiomas <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboIdiomas.SetFocus
    Exit Sub
End If

''' TIRAGEM
If Me.txtTiragem <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtTiragem.SetFocus
    Exit Sub
End If

''' ESPECIFICAÇÃO
If Me.txtEspecificacao <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtEspecificacao.SetFocus
    Exit Sub
End If

''' MOEDA
If Me.cboMoeda <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboMoeda.SetFocus
    Exit Sub
End If

''' ROYALTY (%)
If Me.txtRoyalty_Percentual <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtRoyalty_Percentual.SetFocus
    Exit Sub
End If

''' ROYALTY (VALOR)
If Me.txtRoyalty_Valor <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtRoyalty_Valor.SetFocus
    Exit Sub
End If

''' RE-IMPRESSÃO
If Me.txtReImpressao <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtReImpressao.SetFocus
    Exit Sub
End If

If Cadastro Then
    
    Application.ScreenUpdating = False
    
    admIntervalosDeEdicaoControle strBanco, strSheet, strGerente
    
    DesbloqueioDeGuia SenhaBloqueio
    
    '' CRIAR INTERVALO DE EDIÇÃO MAS Ñ DESMARCAR TEXTO
    IntervaloEditacaoCriar "ORÇAMENTO", "C4:E5,G3:H5,C6,C8:J10,C12:J13,C15:J21,C60:J60", True
    
    Range(ProjetoAtual & "13") = Me.cboVendas
    Range(ProjetoAtual & "15") = Me.cboIdiomas
    Range(ProjetoAtual & "16") = Me.txtTiragem
    Range(ProjetoAtual & "17") = Me.txtEspecificacao
    Range(ProjetoAtual & "18") = Me.cboMoeda
    Range(ProjetoAtual & "19") = Me.txtRoyalty_Percentual
    Range(ProjetoAtual & "20") = Me.txtRoyalty_Valor
    Range(ProjetoAtual & "21") = Me.txtReImpressao
    
    IntervaloEditacaoRemover "ORÇAMENTO", "C4:E5,G3:H5,C6,C8:J10,C12:J13,C15:J21,C60:J60"
    
    BloqueioDeGuia SenhaBloqueio
    
    Application.ScreenUpdating = True
    
End If


cmdCadastrar_Fim:
        
    Call cmdFechar_Click
        
    Exit Sub
cmdCadastrar_err:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Cadastro de projeto"
    Resume cmdCadastrar_Fim


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

    
    If ProjetoAtual = "" Then
    
        ProjetoAtual = "C"
        
    Else
    
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

    End If
    
End Sub

