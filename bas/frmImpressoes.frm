VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImpressoes 
   Caption         =   "IMPRESSÕES"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12090
   OleObjectBlob   =   "frmImpressoes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImpressoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private registro As New clsProjetoImpressao

Private Sub UserForm_Activate()
    carregarDados
    carregarListagens
    carregarProjetos
End Sub

Private Sub cmdSalvar_Click()
    carregarProjeto Me.lstRegistros.Column(1)
    Call cmdCancelar_Click
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub lstRegistros_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNull(Me.lstRegistros.value) Then
        carregarCampos Me.lstRegistros.Column(1)
        Me.cmdSalvar.Caption = "SALVAR"
        Me.cmdSalvar.Enabled = True
    End If
End Sub

Private Sub carregarProjetos()
Dim col As clsProjetoImpressao

With Me.lstRegistros
    .Clear
    .ColumnCount = 4
    .ColumnWidths = "20;0;70;90"

    For Each col In registro.Itens
        .AddItem col.ID
        .List(.ListCount - 1, 1) = col.ColunaExcel
        .List(.ListCount - 1, 2) = col.Tipo
        .List(.ListCount - 1, 3) = col.Papel
    Next col

End With

End Sub

Private Sub carregarProjeto(c As Integer)
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

With Me
    
    ws.Cells(25, c).value = .cboTipo.value
    ws.Cells(26, c).value = .cboPapel.value
    ws.Cells(27, c).value = .cboNumPaginas.value
    ws.Cells(28, c).value = .cboImpressao.value
    ws.Cells(29, c).value = .cboFormato.value

End With

End Sub

Private Sub carregarCampos(c As Integer)
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

With Me
    
    .cboTipo.value = ws.Cells(25, c).value
    .cboPapel.value = ws.Cells(26, c).value
    .cboNumPaginas.value = ws.Cells(27, c).value
    .cboImpressao.value = ws.Cells(28, c).value
    .cboFormato.value = ws.Cells(29, c).value

End With

End Sub

Private Sub carregarDados()
Dim col As New clsProjetoImpressao
Dim x As Integer

lCol = 8
c = 3
For x = 1 To lCol
    Set col = New clsProjetoImpressao
    With col
        .ID = x
        
        .Tipo = Cells(25, c).value
        .Papel = Cells(26, c).value
        .NumPaginas = Cells(27, c).value
        .impressao = Cells(28, c).value
        .Formato = Cells(29, c).value
                
        .ColunaExcel = c
        
        c = c + 1
        
        registro.add col
    End With
Next x

End Sub

Private Sub carregarListagens()
Dim ws As Worksheet
Set ws = Worksheets("Apoio")

ComboBoxUpdate ws.Name, "TIPO", Me.cboTipo
ComboBoxUpdate ws.Name, "PAPEL", Me.cboPapel
ComboBoxUpdate ws.Name, "NPAGINAS", Me.cboNumPaginas
ComboBoxUpdate ws.Name, "IMPRESSAO", Me.cboImpressao
ComboBoxUpdate ws.Name, "FORMATO", Me.cboFormato

End Sub
