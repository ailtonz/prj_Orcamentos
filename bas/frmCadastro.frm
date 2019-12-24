VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadastro 
   Caption         =   "Cadastro"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7755
   OleObjectBlob   =   "frmCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCadastrar_Click()
On Error GoTo cmdCadastrar_err

Dim strMSG As String: strMSG = "Favor Preencher campo"
Dim strTitulo As String: strTitulo = "CAMPO OBRIGATORIO!"

Dim strBanco As String: strBanco = Range(BancoLocal)
Dim Cadastro As Boolean: Cadastro = False

If Me.cboClientes <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboClientes.SetFocus
    Exit Sub
End If

If Me.txtResponsavel <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtResponsavel.SetFocus
    Exit Sub
End If

If Me.cboLinha <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboLinha.SetFocus
    Exit Sub
End If

If Me.txtTitulo <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtTitulo.SetFocus
    Exit Sub
End If

If Me.cboPublisher <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboPublisher.SetFocus
    Exit Sub
End If

If Me.cboJournal <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.cboJournal.SetFocus
    Exit Sub
End If

If Me.txtVolume <> "" Then
    
    Cadastro = True
Else
    MsgBox strMSG, vbCritical + vbOKOnly, strTitulo
    Me.txtVolume.SetFocus
    Exit Sub
End If

If Cadastro Then
    Application.ScreenUpdating = False
    
    admIntervalosDeEdicaoControle strBanco, ActiveSheet.Name, Range(GerenteDeContas)
    
    DesbloqueioDeGuia SenhaBloqueio
    
    '' CRIAR INTERVALO DE EDIÇÃO MAS Ñ DESMARCAR TEXTO
    IntervaloEditacaoCriar "ORÇAMENTO", "C4:E5,G3:H5,C6,C8:J10,C12:J13,C15:J21,C60:J60", True
    
    Range("C4") = Me.cboClientes
    Range("C5") = Me.txtResponsavel
    Range("G5") = Me.cboLinha
    Range("C6") = Me.txtTitulo
    Range("C8") = Me.cboPublisher
    Range("C9") = Me.cboJournal
    Range("C10") = Me.txtVolume
    
       
    IntervaloEditacaoRemover "ORÇAMENTO", "C4:E5,G3:H5,C6,C8:J10,C12:J13,C15:J21,C60:J60"
    
    BloqueioDeGuia SenhaBloqueio
    
    Application.ScreenUpdating = True
    
End If

cmdCadastrar_Fim:
        
    Call cmdFechar_Click
        
    Exit Sub
cmdCadastrar_err:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Envio de Orçamento(s)"
    Resume cmdCadastrar_Fim

End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub UserForm_Initialize()
Dim cPart As Range
Dim cLoc As Range

Dim ws As Worksheet
Set ws = Worksheets("Apoio")

Dim wsPrincipal As Worksheet
Set wsPrincipal = Worksheets(ActiveSheet.Name)

' LINHA
For Each cLoc In wsPrincipal.Range("LINHA")
  With Me.cboLinha
    .AddItem cLoc.Value
  End With
Next cLoc

Me.cboLinha = Range("G5")

' CLIENTES
For Each cLoc In ws.Range("CLIENTES")
  With Me.cboClientes
    .AddItem cLoc.Value
  End With
Next cLoc

Me.cboClientes = Range("C4")
Me.txtResponsavel = Range("C5")

Me.txtTitulo = Range("C6")

' PUBLISHER
For Each cLoc In ws.Range("PUBLISHER")
  With Me.cboPublisher
    .AddItem cLoc.Value
  End With
Next cLoc

Me.cboPublisher = Range("C8")

' JOURNAL
For Each cLoc In ws.Range("JOURNAL")
  With Me.cboJournal
    .AddItem cLoc.Value
  End With
Next cLoc

Me.cboJournal = Range("C9")
Me.txtVolume = Range("C10")
Me.cboClientes.SetFocus

End Sub


