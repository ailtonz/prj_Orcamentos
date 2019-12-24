VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGrand 
   Caption         =   "CADASTRO DE GRAND'S"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   OleObjectBlob   =   "frmGrand.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGrand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    listarGrands
    listarProfissoes
    limparCampos
End Sub

Private Sub cmdCancelar_Click()
    limparCampos
End Sub

Private Sub cmdSalvar_Click()
    salvarRegistro
End Sub

Private Sub lstRegistros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNull(Me.lstRegistros.value) Then
        carregarCampos
        Me.cmdSalvar.Caption = "SALVAR"
    End If
    
End Sub

Private Sub lstRegistros_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Not IsNull(Me.lstRegistros.value) Then
            carregarCampos
            Me.cmdSalvar.Caption = "EXCLUIR"
        End If
    End If
End Sub

Private Sub txtValorLiquido_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtValorLiquido.value <> "" Then
        Me.txtValorLiquido.value = FormatCurrency(Me.txtValorLiquido.value)
    End If
    
End Sub

Private Sub txtNome_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtNome.value = UCase(Me.txtNome.value)
End Sub

Private Sub salvarRegistro()

Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

Dim Grd As clsGrands
Set Grd = New clsGrands

carregarBanco
            
    With Grd
        .ID = Me.txtId.value
        
        .NumProjeto = objOrc.NumProjeto
        .NumControle = objOrc.Controle
        .Vendedor = objOrc.Vendedor
        
        .Profissao = Me.cboProfissao.value
        .Nome = Me.txtNome.value
        .ValorLiquido = Me.txtValorLiquido.value
        .CustoMedico = ws.Range("C45").value
        .CustoEditorFee = ws.Range("C55").value
        .add Grd
    End With
                      
    If Grd.ID = "" Then
        If (Grd.Insert(Bnc, Grd) = True) Then
            MsgBox "Cadastro realisado com sucesso!", vbInformation + vbOKOnly, "Cadastro"
        Else
            MsgBox "Não foi possivel realizar o cadastro!", vbCritical + vbOKOnly, "Cadastro - ERRO!"
        End If
        
    Else
    
        If Me.cmdSalvar.Caption = "SALVAR" Then
            If (Grd.Update(Bnc, Grd) = True) Then
                MsgBox "Alteração realizada com sucesso!", vbInformation + vbOKOnly, "Alteração"
            Else
                MsgBox "Não foi possivel realizar alteração!", vbCritical + vbOKOnly, "Alteração - ERRO!"
            End If
        Else
            If mostrarRegistro = vbYes Then
                If (Grd.Delete(Bnc, Grd) = True) Then
                    MsgBox "Exclusão realizada com sucesso!", vbInformation + vbOKOnly, "Exclusão"
                Else
                    MsgBox "Não foi possivel realizar Exclusão!", vbCritical + vbOKOnly, "Exclusão - ERRO!"
                End If
            End If
            
        End If
        
    End If
    

    
    listarGrands
    limparCampos
'    custosGrand

Set Grd = Nothing
Set Bnc = Nothing

End Sub

Private Function mostrarRegistro() As Variant
Dim retVal As Variant

    retVal = MsgBox("Você deseja realmente EXCLUIR o registro abaixo:" & vbNewLine & _
            vbNewLine & _
            "NOME: " & Me.lstRegistros.Column(1) & vbNewLine & _
            "PROFISSÃO: " & Me.lstRegistros.Column(2) & vbNewLine & _
            "VALOR : " & FormatCurrency(Me.lstRegistros.Column(3)) & vbNewLine, vbCritical + vbYesNo, "EXCLUSÃO DE REGISTRO!")
            
    mostrarRegistro = retVal
            
Set retVal = Nothing

End Function

Private Sub carregarCampos()
Dim obj As clsGrands
Dim col As clsGrands
Dim i As Long: i = Me.lstRegistros.value

carregarBanco
limparCampos

Set obj = New clsGrands

Set col = obj.getGrandsIndex(Bnc, i)
Set registro = obj.getGrandsIndex(Bnc, i)

For Each obj In col.Itens
   
    Me.txtId.value = obj.ID
            
    Me.cboProfissao.SetFocus
    Me.cboProfissao.SelText = obj.Profissao
    
    Me.txtNome.value = obj.Nome
    
    Me.txtValorLiquido.value = FormatCurrency(obj.ValorLiquido)
        
    Me.cboProfissao.SetFocus

Next obj
                
End Sub

Private Sub listarProfissoes()

carregarBanco

Dim Prf As clsProfissoes
Set Prf = New clsProfissoes

Dim col As clsProfissoes
Set col = Prf.getProfissoes(Bnc)
    
With Me.cboProfissao
    .Clear
    .Clear
    .ColumnCount = 1
    .ColumnWidths = "60"
    
    For Each Prf In col.Itens
        .AddItem Prf.Profissao
    Next Prf

End With

End Sub

Private Sub listarGrands()
Dim Prf As clsGrands
Dim col As clsGrands

carregarBanco

Set Prf = New clsGrands
Set col = Prf.getGrands(Bnc, objOrc)

With Me.lstRegistros
    .Clear
    .ColumnCount = 4
    .ColumnWidths = "0;200;60;0"
    
    For Each Prf In col.Itens
        .AddItem Prf.ID
        .List(.ListCount - 1, 1) = Prf.Nome
        .List(.ListCount - 1, 2) = Prf.Profissao
        .List(.ListCount - 1, 3) = Prf.ValorLiquido
    Next Prf

End With

'limparGrandsGuia
'listarGrandsGuia

End Sub

Private Sub limparGrandsGuia()
    DesbloqueioDeGuia SenhaBloqueio
    
    ''' LIMPAR GRAND
    Range("AC3:AE27").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("AC3").Select
    
    BloqueioDeGuia SenhaBloqueio
    
End Sub


Private Sub limparCampos()

    Me.txtId.value = ""
    Me.txtNome.value = ""
    Me.cboProfissao.value = ""
    Me.txtValorLiquido.value = ""
    
    Me.cmdSalvar.Caption = "NOVO"
    
End Sub


Private Sub listarGrandsGuia()
Dim Prf As clsGrands
Dim col As clsGrands
Dim orc As clsOrcamentos

carregarBanco

Set orc = New clsOrcamentos
Set Prf = New clsGrands

With orc
    .Controle = ActiveSheet.Name
    .Vendedor = Range(GerenteDeContas)
    .add orc
End With

Set col = Prf.getGrands(Bnc, orc)

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets(orc.Controle)

''find  first empty row in database
lRow = ws.Cells(Rows.count, 29).End(xlUp).Offset(1, 0).Row
    
DesbloqueioDeGuia SenhaBloqueio
    
For Each Prf In col.Itens
    ws.Range("AC" & lRow).value = Prf.Profissao
    ws.Range("AD" & lRow).value = Prf.Nome
    ws.Range("AE" & lRow).value = Prf.ValorLiquido
    lRow = lRow + 1
Next Prf

BloqueioDeGuia SenhaBloqueio

End Sub

Private Sub custosGrand()
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

DesbloqueioDeGuia SenhaBloqueio

'' MÉDICO
With ws
    .Range("AO10").Select
    Selection.Copy
    .Range("C45:J45").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End With


'' EDITOR FEE
With ws
    .Range("AO11").Select
    Selection.Copy
    .Range("C55:J55").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End With

BloqueioDeGuia SenhaBloqueio

End Sub

