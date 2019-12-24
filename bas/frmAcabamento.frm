VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAcabamento 
   Caption         =   "ACABAMENTOS"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9150.001
   OleObjectBlob   =   "frmAcabamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAcabamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private registro As New clsAcabamento

Private Sub UserForm_Activate()
    carregarDados
    carregarListagens
    carregarAcabamentos
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


Private Sub carregarProjeto(l As Integer)
Dim ws As Worksheet
Set ws = Worksheets(ActiveSheet.Name)

With Me
    
    ws.Cells(l, 2).value = .cboAcabamentos.value

End With

End Sub


Private Sub carregarAcabamentos()
Dim col As clsAcabamento

With Me.lstRegistros
    .Clear
    .ColumnCount = 3
    .ColumnWidths = "0;0;70"

    For Each col In registro.Itens
        .AddItem col.ID
        .List(.ListCount - 1, 1) = col.ColunaExcel
        .List(.ListCount - 1, 2) = col.Acabamento
    Next col

End With

End Sub


Private Sub carregarCampos(l As Integer)

With Me
    
    .cboAcabamentos.value = Cells(l, 2).value

End With

End Sub

Private Sub carregarDados()
Dim col As New clsAcabamento
Dim l  As Integer, c As Integer

lRow = 34
c = 2

For l = 31 To lRow
    Set col = New clsAcabamento
    With col
        .ID = l
        
        .Acabamento = Cells(l, c).value
        
        .ColunaExcel = l
            
        registro.add col
    End With
Next l

End Sub

Private Sub carregarListagens()

ComboBoxUpdate "apoio", "ACABAMENTO", Me.cboAcabamentos

End Sub
